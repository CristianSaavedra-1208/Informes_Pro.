import io
import os
import tempfile
import unittest
import warnings
from unittest.mock import patch

import openpyxl

from src.core.external_notes_manager import ExternalNotesManager


class TestExternalNotesAggregation(unittest.TestCase):
    def setUp(self):
        ExternalNotesManager._AGG_ANEXO_CACHE.clear()

    def test_real_consolidated_group_2025_12_code_aligned_aggregation(self):
        """
        Verifica que [GRUPO] Consolidado DB Terra Holdco en 2025-12 agregue correctamente
        el anexo de Pacifico Cable SpA pese al desfase de filas/columnas respecto a la plantilla máster:
        - #N07.1 Inventarios = 33.247.439 M$ (y no 0)
        - #N09.1 Activo Fijo Bruto = 372.803.055 / Deprec = -230.698.160 / Neto = 142.104.895 M$
        - #N04.3 Fondos Mutuos = 7.032.177 M$
        - #N16.1 Pasivos por derecho de uso incorporado dinámicamente (4.847.998 Cte / 5.820.163 No Cte)
        - #N20.1..3 Patrimonio conserva acciones de la Matriz (50.334.500) sin warnings de fecha serial.
        """
        grupo = "[GRUPO] Consolidado DB Terra Holdco"
        periodo = "2025-12"

        if not ExternalNotesManager.has_anexo(grupo, periodo):
            self.skipTest("No existe anexo de filial para 2025-12 en el entorno actual.")

        with warnings.catch_warnings(record=True) as caught:
            warnings.simplefilter("always")
            agg_io = ExternalNotesManager.build_aggregated_group_anexo(grupo, periodo)
            wb = openpyxl.load_workbook(agg_io, data_only=False)

        date_warnings = [w for w in caught if "outside the limits for dates" in str(w.message)]
        self.assertEqual(len(date_warnings), 0, f"No deben emitirse warnings de fecha serial: {date_warnings}")

        # 1. Inventarios (#N07.1): suma al menos Pacífico (33.247.439) y, si Holdco tiene anexo cargado, el consolidado completo
        ws_inv = wb["Inventarios"]
        inv_vals = [ws_inv.cell(r, 2).value or 0 for r in range(6, 13)]
        self.assertGreaterEqual(round(sum(inv_vals)), 33247439)

        # 2. Activo Fijo (#N09.1)
        ws_af = wb["Activo Fijo"]
        self.assertEqual(ws_af.cell(5, 2).value, "Activo fijo bruto")
        self.assertEqual(ws_af.cell(5, 3).value, "Depreciación acumulada")
        self.assertEqual(ws_af.cell(5, 4).value, "Activo fijo neto")
        bruto_sum = sum(ws_af.cell(r, 2).value or 0 for r in range(8, 24))
        deprec_sum = sum(ws_af.cell(r, 3).value or 0 for r in range(8, 24))
        self.assertGreaterEqual(round(bruto_sum), 372803055)
        self.assertLessEqual(round(deprec_sum), -230698160)
        self.assertGreaterEqual(round(bruto_sum + deprec_sum), 142104895)

        # 3. Efectivo — Detalle Fondos Mutuos (#N04.3)
        ws_ef = wb["Efectivo"]
        fm_vals = [ws_ef.cell(r, 4).value or 0 for r in range(7, ws_ef.max_row + 1) if isinstance(ws_ef.cell(r, 4).value, (int, float))]
        self.assertGreaterEqual(len(fm_vals), 5)
        self.assertGreaterEqual(round(sum(fm_vals)), 7032177)

        # 4. Pasivos derechos de uso (#N16.1 incorporado dinámicamente + #N16.2[COMPARATIVO])
        ws_pdu = wb["Pasivos derechos de  uso"]
        pdu_codes = ExternalNotesManager.get_code_locations(ws_pdu)
        self.assertIn("#N16.1", pdu_codes)
        r16_1 = pdu_codes["#N16.1"]["row"]
        cte_16_1 = sum(
            ws_pdu.cell(r, 2).value or 0
            for r in range(r16_1 + 4, r16_1 + 11)
            if isinstance(ws_pdu.cell(r, 2).value, (int, float))
        )
        ncte_16_1 = sum(
            ws_pdu.cell(r, 3).value or 0
            for r in range(r16_1 + 4, r16_1 + 11)
            if isinstance(ws_pdu.cell(r, 3).value, (int, float))
        )
        self.assertGreaterEqual(round(cte_16_1), 4847998)
        self.assertGreaterEqual(round(ncte_16_1), 5820163)

        # 5. Patrimonio (#N20.1 conserva 50.334.500 acciones de la Matriz)
        ws_pat = wb["Patrimonio"]
        self.assertEqual(ws_pat.cell(5, 2).value, 50334500)
        wb.close()

    def test_multi_subsidiary_shifted_columns_and_detail_expansion(self):
        """
        Verifica que cuando 2 empresas del grupo tienen sus cuadros #N07.1, #N04.3 y #N16.2[COMPARATIVO]
        en distintas filas y columnas, build_aggregated_group_anexo suma exactamente concepto por concepto
        y expande las filas de detalle multicolumna (#N04.3) cuando superan la capacidad base.
        """
        with tempfile.TemporaryDirectory() as tmpdir:
            sub_a_path = os.path.join(tmpdir, "sub_a.xlsx")
            sub_b_path = os.path.join(tmpdir, "sub_b.xlsx")

            # Construir Sub A: #N07.1 desplazado a Col B (col=2, row=1), 4 fondos mutuos en #N04.3
            wb_a = openpyxl.Workbook()
            ws_inv_a = wb_a.active
            ws_inv_a.title = "Inventarios"
            ws_inv_a.cell(1, 2, "#N07.1")
            ws_inv_a.cell(3, 2, "Detalle")
            ws_inv_a.cell(3, 3, "M$")
            ws_inv_a.cell(5, 2, "Bodega central (*)")
            ws_inv_a.cell(5, 3, 10000)
            ws_inv_a.cell(6, 2, "Otras bodegas (**)")
            ws_inv_a.cell(6, 3, 2500)
            ws_inv_a.cell(7, 2, "Deterioro por obsolescencia de materiales")
            ws_inv_a.cell(7, 3, -500)
            ws_inv_a.cell(8, 2, "Total")
            ws_inv_a.cell(8, 3, 12000)

            ws_ef_a = wb_a.create_sheet("Efectivo")
            ws_ef_a.cell(2, 1, "#N04.3")
            ws_ef_a.cell(5, 1, "Detalle")
            ws_ef_a.cell(5, 2, "Banco")
            ws_ef_a.cell(5, 3, "Moneda")
            ws_ef_a.cell(6, 4, "M$")
            for i in range(4):
                ws_ef_a.cell(7 + i, 1, "Fondos mutuos")
                ws_ef_a.cell(7 + i, 2, f"Banco A{i+1}")
                ws_ef_a.cell(7 + i, 3, "$")
                ws_ef_a.cell(7 + i, 4, 1000 * (i + 1))
            ws_ef_a.cell(11, 1, "Total")
            ws_ef_a.cell(11, 4, 10000)

            ws_pdu_a = wb_a.create_sheet("Pasivos derechos de  uso")
            ws_pdu_a.cell(5, 3, "#N16.2[COMPARATIVO]")
            ws_pdu_a.cell(7, 3, "Detalle al 31/12/2025")
            ws_pdu_a.cell(7, 4, "90 días")
            ws_pdu_a.cell(7, 5, "90 días a 1 años")
            ws_pdu_a.cell(8, 4, "M$")
            ws_pdu_a.cell(8, 5, "M$")
            ws_pdu_a.cell(9, 3, "Pasivos por derecho de uso")
            ws_pdu_a.cell(9, 4, 300)
            ws_pdu_a.cell(9, 5, 700)
            ws_pdu_a.cell(10, 3, "Total")
            wb_a.save(sub_a_path)
            wb_a.close()

            # Construir Sub B: #N07.1 desplazado a Col C (col=3, row=3), 4 fondos mutuos más en #N04.3 (total 8 > 6 slots)
            wb_b = openpyxl.Workbook()
            ws_inv_b = wb_b.active
            ws_inv_b.title = "Inventarios"
            ws_inv_b.cell(3, 3, "#N07.1")
            ws_inv_b.cell(5, 3, "Detalle")
            ws_inv_b.cell(5, 4, "M$")
            ws_inv_b.cell(7, 3, "Bodega central (*)")
            ws_inv_b.cell(7, 4, 5000)
            ws_inv_b.cell(8, 3, "Bodega tecnica")
            ws_inv_b.cell(8, 4, 3200)
            ws_inv_b.cell(9, 3, "Deterioro por obsolescencia de materiales")
            ws_inv_b.cell(9, 4, -200)
            ws_inv_b.cell(10, 3, "Total")
            ws_inv_b.cell(10, 4, 8000)

            ws_ef_b = wb_b.create_sheet("Efectivo")
            ws_ef_b.cell(4, 2, "#N04.3")
            ws_ef_b.cell(6, 2, "Detalle")
            ws_ef_b.cell(6, 3, "Banco")
            ws_ef_b.cell(6, 4, "Moneda")
            ws_ef_b.cell(7, 5, "M$")
            for i in range(4):
                ws_ef_b.cell(8 + i, 2, "Fondos mutuos")
                ws_ef_b.cell(8 + i, 3, f"Banco B{i+1}")
                ws_ef_b.cell(8 + i, 4, "US$")
                ws_ef_b.cell(8 + i, 5, 500 * (i + 1))
            ws_ef_b.cell(12, 2, "Total")
            ws_ef_b.cell(12, 5, 5000)

            ws_pdu_b = wb_b.create_sheet("Pasivos derechos de  uso")
            ws_pdu_b.cell(2, 1, "#N16.2[COMPARATIVO]")
            ws_pdu_b.cell(4, 1, "Detalle al 31/12/2025")
            ws_pdu_b.cell(4, 2, "90 días")
            ws_pdu_b.cell(4, 3, "90 días a 1 años")
            ws_pdu_b.cell(5, 2, "M$")
            ws_pdu_b.cell(5, 3, "M$")
            ws_pdu_b.cell(6, 1, "Pasivos por derecho de uso")
            ws_pdu_b.cell(6, 2, 150)
            ws_pdu_b.cell(6, 3, 250)
            ws_pdu_b.cell(7, 1, "Total")
            wb_b.save(sub_b_path)
            wb_b.close()

            path_map = {"Sub A": sub_a_path, "Sub B": sub_b_path}

            with patch.object(ExternalNotesManager, "get_group_companies", return_value=["Sub A", "Sub B"]), \
                 patch.object(ExternalNotesManager, "get_group_parent_company", return_value="Sub A"), \
                 patch.object(ExternalNotesManager, "has_explicit_anexo", side_effect=lambda emp, per: emp in path_map), \
                 patch.object(ExternalNotesManager, "get_anexo_path", side_effect=lambda emp, per: path_map[emp]):
                agg_io = ExternalNotesManager.build_aggregated_group_anexo("[GRUPO] Test Grupo", "2025-12")
                self.assertIsNotNone(agg_io)
                agg_wb = openpyxl.load_workbook(agg_io, data_only=False)

                # Inventarios: Bodega central = 15000, Otras bodegas = 2500, Bodega tecnica = 3200, Deterioro = -700 => Total = 20000
                ws_inv = agg_wb["Inventarios"]
                self.assertEqual(ws_inv.cell(6, 2).value, 15000)
                self.assertEqual(ws_inv.cell(7, 2).value, 2500)
                self.assertEqual(ws_inv.cell(8, 2).value, 3200)
                self.assertEqual(ws_inv.cell(12, 2).value, -700)
                inv_total = sum(ws_inv.cell(r, 2).value or 0 for r in range(6, 13))
                self.assertEqual(inv_total, 20000)

                # Efectivo #N04.3: 8 filas de detalle (expandidas desde 6), suma = 10000 + 5000 = 15000
                ws_ef = agg_wb["Efectivo"]
                ef_nums = [
                    ws_ef.cell(r, 4).value
                    for r in range(7, ws_ef.max_row + 1)
                    if isinstance(ws_ef.cell(r, 4).value, (int, float))
                ]
                self.assertEqual(len(ef_nums), 8)
                self.assertEqual(sum(ef_nums), 15000)
                self.assertEqual(ws_ef.cell(15, 4).value, "=SUM(D7:D14)")

                # Pasivos derechos de uso #N16.2[COMPARATIVO]: 90 días = 300+150=450, 90d-1a = 700+250=950
                ws_pdu = agg_wb["Pasivos derechos de  uso"]
                self.assertEqual(ws_pdu.cell(6, 2).value, 450)
                self.assertEqual(ws_pdu.cell(6, 3).value, 950)
                agg_wb.close()

    def test_inter_company_identical_descriptions_merge_and_normalize_accents_case(self):
        """
        Verifica que cuando distintas compañías informan la misma descripción completa
        (con diferencias de mayúsculas, tildes o espacios), se mantiene 1 sola línea y se suman los montos.
        """
        with tempfile.TemporaryDirectory() as tmpdir:
            sub_a_path = os.path.join(tmpdir, "sub_a.xlsx")
            sub_b_path = os.path.join(tmpdir, "sub_b.xlsx")

            wb_a = openpyxl.Workbook()
            ws_a = wb_a.active
            ws_a.title = "Efectivo"
            ws_a.cell(2, 1, "#N04.3")
            ws_a.cell(5, 1, "Detalle")
            ws_a.cell(5, 2, "Banco")
            ws_a.cell(5, 3, "Moneda")
            ws_a.cell(6, 4, "M$")
            ws_a.cell(7, 1, "Fondos mutuos")
            ws_a.cell(7, 2, "  SCOTIABANK ")
            ws_a.cell(7, 3, "$")
            ws_a.cell(7, 4, 346885)
            ws_a.cell(8, 1, "Fondos mutuos")
            ws_a.cell(8, 2, "Banco de Crédito")
            ws_a.cell(8, 3, "US$")
            ws_a.cell(8, 4, 1000)
            ws_a.cell(9, 1, "Total")
            wb_a.save(sub_a_path)
            wb_a.close()

            wb_b = openpyxl.Workbook()
            ws_b = wb_b.active
            ws_b.title = "Efectivo"
            ws_b.cell(2, 1, "#N04.3")
            ws_b.cell(5, 1, "Detalle")
            ws_b.cell(5, 2, "Banco")
            ws_b.cell(5, 3, "Moneda")
            ws_b.cell(6, 4, "M$")
            ws_b.cell(7, 1, "fondos mutuos")
            ws_b.cell(7, 2, "Scotiabank")
            ws_b.cell(7, 3, "$")
            ws_b.cell(7, 4, 2839260)
            ws_b.cell(8, 1, "Fondos mutuos")
            ws_b.cell(8, 2, "banco de credito")
            ws_b.cell(8, 3, "us$")
            ws_b.cell(8, 4, 2500)
            ws_b.cell(9, 1, "Total")
            wb_b.save(sub_b_path)
            wb_b.close()

            path_map = {"Sub A": sub_a_path, "Sub B": sub_b_path}
            with patch.object(ExternalNotesManager, "get_group_companies", return_value=["Sub A", "Sub B"]), \
                 patch.object(ExternalNotesManager, "get_group_parent_company", return_value="Sub A"), \
                 patch.object(ExternalNotesManager, "has_explicit_anexo", side_effect=lambda emp, per: emp in path_map), \
                 patch.object(ExternalNotesManager, "get_anexo_path", side_effect=lambda emp, per: path_map[emp]):
                agg_io = ExternalNotesManager.build_aggregated_group_anexo("[GRUPO] Test Merge", "2025-12")
                agg_wb = openpyxl.load_workbook(agg_io, data_only=False)
                ws_ef = agg_wb["Efectivo"]
                rows_found = [
                    (ws_ef.cell(r, 2).value, ws_ef.cell(r, 3).value, ws_ef.cell(r, 4).value)
                    for r in range(7, ws_ef.max_row + 1)
                    if isinstance(ws_ef.cell(r, 4).value, (int, float))
                ]
                self.assertEqual(len(rows_found), 2)
                self.assertEqual(rows_found[0][2], 3186145)
                self.assertEqual(rows_found[1][2], 3500)
                agg_wb.close()

    def test_intra_company_duplicate_lines_preserved_separately(self):
        """
        Verifica que si una misma compañía informa 2 o más líneas con idéntica descripción
        (ej. 2 series en Estado | $ o 2 tramos idénticos), se mantienen en líneas separadas.
        """
        with tempfile.TemporaryDirectory() as tmpdir:
            sub_a_path = os.path.join(tmpdir, "sub_a.xlsx")
            sub_b_path = os.path.join(tmpdir, "sub_b.xlsx")

            wb_a = openpyxl.Workbook()
            ws_a = wb_a.active
            ws_a.title = "Efectivo"
            ws_a.cell(2, 1, "#N04.3")
            ws_a.cell(5, 1, "Detalle")
            ws_a.cell(5, 2, "Banco")
            ws_a.cell(5, 3, "Moneda")
            ws_a.cell(6, 4, "M$")
            ws_a.cell(7, 1, "Fondos mutuos")
            ws_a.cell(7, 2, "Santander")
            ws_a.cell(7, 3, "$")
            ws_a.cell(7, 4, 100000)
            ws_a.cell(8, 1, "Total")
            wb_a.save(sub_a_path)
            wb_a.close()

            wb_b = openpyxl.Workbook()
            ws_b = wb_b.active
            ws_b.title = "Efectivo"
            ws_b.cell(2, 1, "#N04.3")
            ws_b.cell(5, 1, "Detalle")
            ws_b.cell(5, 2, "Banco")
            ws_b.cell(5, 3, "Moneda")
            ws_b.cell(6, 4, "M$")
            ws_b.cell(7, 1, "Fondos mutuos")
            ws_b.cell(7, 2, "Estado")
            ws_b.cell(7, 3, "$")
            ws_b.cell(7, 4, 866211)
            ws_b.cell(8, 1, "Fondos mutuos")
            ws_b.cell(8, 2, "Estado")
            ws_b.cell(8, 3, "$")
            ws_b.cell(8, 4, 61178)
            ws_b.cell(9, 1, "Total")
            wb_b.save(sub_b_path)
            wb_b.close()

            path_map = {"Sub A": sub_a_path, "Sub B": sub_b_path}
            with patch.object(ExternalNotesManager, "get_group_companies", return_value=["Sub A", "Sub B"]), \
                 patch.object(ExternalNotesManager, "get_group_parent_company", return_value="Sub A"), \
                 patch.object(ExternalNotesManager, "has_explicit_anexo", side_effect=lambda emp, per: emp in path_map), \
                 patch.object(ExternalNotesManager, "get_anexo_path", side_effect=lambda emp, per: path_map[emp]):
                agg_io = ExternalNotesManager.build_aggregated_group_anexo("[GRUPO] Test IntraDup", "2025-12")
                agg_wb = openpyxl.load_workbook(agg_io, data_only=False)
                ws_ef = agg_wb["Efectivo"]
                rows_found = [
                    (ws_ef.cell(r, 2).value, ws_ef.cell(r, 4).value)
                    for r in range(7, ws_ef.max_row + 1)
                    if isinstance(ws_ef.cell(r, 4).value, (int, float))
                ]
                self.assertEqual(rows_found, [("Santander", 100000), ("Estado", 866211), ("Estado", 61178)])
                agg_wb.close()

    def test_different_bank_names_kept_in_separate_lines(self):
        """
        Verifica que nombres distintos (ej. 'BCI' vs 'Banco de Crédito e Inversiones')
        no se unifiquen mágicamente, sino que permanezcan en 2 líneas separadas.
        """
        with tempfile.TemporaryDirectory() as tmpdir:
            sub_a_path = os.path.join(tmpdir, "sub_a.xlsx")
            sub_b_path = os.path.join(tmpdir, "sub_b.xlsx")

            wb_a = openpyxl.Workbook()
            ws_a = wb_a.active
            ws_a.title = "Efectivo"
            ws_a.cell(2, 1, "#N04.3")
            ws_a.cell(5, 1, "Detalle")
            ws_a.cell(5, 2, "Banco")
            ws_a.cell(5, 3, "Moneda")
            ws_a.cell(6, 4, "M$")
            ws_a.cell(7, 1, "Fondos mutuos")
            ws_a.cell(7, 2, "BCI")
            ws_a.cell(7, 3, "$")
            ws_a.cell(7, 4, 8504606)
            ws_a.cell(8, 1, "Total")
            wb_a.save(sub_a_path)
            wb_a.close()

            wb_b = openpyxl.Workbook()
            ws_b = wb_b.active
            ws_b.title = "Efectivo"
            ws_b.cell(2, 1, "#N04.3")
            ws_b.cell(5, 1, "Detalle")
            ws_b.cell(5, 2, "Banco")
            ws_b.cell(5, 3, "Moneda")
            ws_b.cell(6, 4, "M$")
            ws_b.cell(7, 1, "Fondos mutuos")
            ws_b.cell(7, 2, "Banco de Crédito e Inversiones")
            ws_b.cell(7, 3, "$")
            ws_b.cell(7, 4, 145029)
            ws_b.cell(8, 1, "Total")
            wb_b.save(sub_b_path)
            wb_b.close()

            path_map = {"Sub A": sub_a_path, "Sub B": sub_b_path}
            with patch.object(ExternalNotesManager, "get_group_companies", return_value=["Sub A", "Sub B"]), \
                 patch.object(ExternalNotesManager, "get_group_parent_company", return_value="Sub A"), \
                 patch.object(ExternalNotesManager, "has_explicit_anexo", side_effect=lambda emp, per: emp in path_map), \
                 patch.object(ExternalNotesManager, "get_anexo_path", side_effect=lambda emp, per: path_map[emp]):
                agg_io = ExternalNotesManager.build_aggregated_group_anexo("[GRUPO] Test DiffNames", "2025-12")
                agg_wb = openpyxl.load_workbook(agg_io, data_only=False)
                ws_ef = agg_wb["Efectivo"]
                rows_found = [
                    (ws_ef.cell(r, 2).value, ws_ef.cell(r, 4).value)
                    for r in range(7, ws_ef.max_row + 1)
                    if isinstance(ws_ef.cell(r, 4).value, (int, float))
                ]
                self.assertEqual(rows_found, [("BCI", 8504606), ("Banco de Crédito e Inversiones", 145029)])
                agg_wb.close()

    def test_single_column_new_concept_appended_before_total(self):
        """
        Verifica que en una tabla de 1 columna descriptiva (#N07.1 Inventarios), si una sociedad
        agrega un concepto nuevo con saldo != 0 ('Bodega zona norte'), se inserta antes del Total,
        se suma si otra sociedad también lo trae, y se expande el rango =SUM(...) en la nota final.
        """
        with tempfile.TemporaryDirectory() as tmpdir:
            sub_a_path = os.path.join(tmpdir, "sub_a.xlsx")
            sub_b_path = os.path.join(tmpdir, "sub_b.xlsx")

            wb_a = openpyxl.Workbook()
            ws_a = wb_a.active
            ws_a.title = "Inventarios"
            ws_a.cell(2, 1, "#N07.1")
            ws_a.cell(4, 1, "Detalle")
            ws_a.cell(4, 2, "M$")
            ws_a.cell(6, 1, "Bodega central (*)")
            ws_a.cell(6, 2, 10000)
            ws_a.cell(7, 1, "Bodega zona norte")
            ws_a.cell(7, 2, 3000)
            ws_a.cell(8, 1, "Total")
            wb_a.save(sub_a_path)
            wb_a.close()

            wb_b = openpyxl.Workbook()
            ws_b = wb_b.active
            ws_b.title = "Inventarios"
            ws_b.cell(2, 1, "#N07.1")
            ws_b.cell(4, 1, "Detalle")
            ws_b.cell(4, 2, "M$")
            ws_b.cell(6, 1, "Bodega central (*)")
            ws_b.cell(6, 2, 5000)
            ws_b.cell(7, 1, "Bodega zona norte")
            ws_b.cell(7, 2, 2000)
            ws_b.cell(8, 1, "Total")
            wb_b.save(sub_b_path)
            wb_b.close()

            path_map = {"Sub A": sub_a_path, "Sub B": sub_b_path}
            with patch.object(ExternalNotesManager, "get_group_companies", return_value=["Sub A", "Sub B"]), \
                 patch.object(ExternalNotesManager, "get_group_parent_company", return_value="Sub A"), \
                 patch.object(ExternalNotesManager, "has_explicit_anexo", side_effect=lambda emp, per: emp in path_map if emp != "[GRUPO] Test NewConcept" else False), \
                 patch.object(ExternalNotesManager, "get_anexo_path", side_effect=lambda emp, per: path_map[emp]):
                with open("Plantilla de notas_v1.xlsx", "rb") as f_tpl:
                    not_bytes = io.BytesIO(f_tpl.read())
                out_io = ExternalNotesManager.inject_external_data(
                    not_bytes, "[GRUPO] Test NewConcept", "2026-10", None
                )
                out_wb = openpyxl.load_workbook(out_io, data_only=False)
                ws_out = out_wb["Inventarios"]
                concepts = {
                    str(ws_out.cell(r, 1).value).strip(): (r, ws_out.cell(r, 2).value)
                    for r in range(6, 20)
                    if ws_out.cell(r, 1).value is not None
                }
                self.assertIn("Bodega zona norte", concepts)
                self.assertEqual(concepts["Bodega central (*)"][1], 15000.0)
                self.assertEqual(concepts["Bodega zona norte"][1], 5000.0)
                r_norte = concepts["Bodega zona norte"][0]
                r_total, f_total = concepts["Total"]
                self.assertGreater(r_total, r_norte)
                self.assertEqual(f_total, f"=SUM(B6:B{r_norte})")
                out_wb.close()

    def test_multicol_actual_vs_comparative_key_alignment(self):
        """
        Verifica que al inyectar simultáneamente Período Actual (2026-10) y Comparativo (2025-12)
        en un cuadro multicolumna (#N04.3), las filas con la misma llave descriptiva compartan fila
        y las filas exclusivas de cada período ocupen su propia línea sin pisar bancos distintos.
        """
        with tempfile.TemporaryDirectory() as tmpdir:
            act_path = os.path.join(tmpdir, "act_2026_10.xlsx")
            cmp_path = os.path.join(tmpdir, "cmp_2025_12.xlsx")

            wb_act = openpyxl.Workbook()
            ws_act = wb_act.active
            ws_act.title = "Efectivo"
            ws_act.cell(2, 1, "#N04.3")
            ws_act.cell(5, 1, "Detalle")
            ws_act.cell(5, 2, "Banco")
            ws_act.cell(5, 3, "Moneda")
            ws_act.cell(6, 4, "M$")
            ws_act.cell(7, 1, "Fondos mutuos")
            ws_act.cell(7, 2, "Scotiabank")
            ws_act.cell(7, 3, "$")
            ws_act.cell(7, 4, 5000000)
            ws_act.cell(8, 1, "Fondos mutuos")
            ws_act.cell(8, 2, "Itaú")
            ws_act.cell(8, 3, "$")
            ws_act.cell(8, 4, 1200000)
            ws_act.cell(9, 1, "Total")
            wb_act.save(act_path)
            wb_act.close()

            wb_cmp = openpyxl.Workbook()
            ws_cmp = wb_cmp.active
            ws_cmp.title = "Efectivo"
            ws_cmp.cell(2, 1, "#N04.3")
            ws_cmp.cell(5, 1, "Detalle")
            ws_cmp.cell(5, 2, "Banco")
            ws_cmp.cell(5, 3, "Moneda")
            ws_cmp.cell(6, 4, "M$")
            ws_cmp.cell(7, 1, "Fondos mutuos")
            ws_cmp.cell(7, 2, "Santander")
            ws_cmp.cell(7, 3, "US$")
            ws_cmp.cell(7, 4, 6664968)
            ws_cmp.cell(8, 1, "Fondos mutuos")
            ws_cmp.cell(8, 2, "scotiabank ")
            ws_cmp.cell(8, 3, "$")
            ws_cmp.cell(8, 4, 3186145)
            ws_cmp.cell(9, 1, "Total")
            wb_cmp.save(cmp_path)
            wb_cmp.close()

            per_paths = {"2026-10": act_path, "2025-12": cmp_path}
            with patch.object(ExternalNotesManager, "has_explicit_anexo", side_effect=lambda emp, per: per in per_paths), \
                 patch.object(ExternalNotesManager, "get_anexo_path", side_effect=lambda emp, per: per_paths[per]):
                with open("Plantilla de notas_v1.xlsx", "rb") as f_tpl:
                    not_bytes = io.BytesIO(f_tpl.read())
                out_io = ExternalNotesManager.inject_external_data(
                    not_bytes, "Empresa Test SpA", "2026-10", "2025-12"
                )
                out_wb = openpyxl.load_workbook(out_io, data_only=False)
                ws_out = out_wb["Efectivo"]
                locs = ExternalNotesManager.get_code_locations(ws_out)
                r0 = locs["#N04.3"]["row"]
                rows_out = []
                for r in range(r0 + 1, ws_out.max_row + 1):
                    lbl = str(ws_out.cell(r, 1).value or "").strip()
                    if lbl.lower() == "total":
                        break
                    if lbl.lower() != "fondos mutuos":
                        continue
                    banco = ws_out.cell(r, 2).value
                    mon = ws_out.cell(r, 3).value
                    v_act = ws_out.cell(r, 4).value
                    v_cmp = ws_out.cell(r, 5).value
                    if banco is not None or v_act is not None or v_cmp is not None:
                        rows_out.append((banco, mon, v_act, v_cmp))

                self.assertEqual(
                    rows_out,
                    [
                        ("Scotiabank", "$", 5000000.0, 3186145.0),
                        ("Itaú", "$", 1200000.0, None),
                        ("Santander", "US$", None, 6664968.0),
                    ],
                )
                out_wb.close()

    def test_n15_4_adiciones_header_alignment_individual_and_consolidated(self):
        """
        Verifica que #N15.4 preserve la columna 'Adiciones' y alinee por encabezado cuando
        una filial tiene 10 columnas (con Adiciones, ej. Holdco 2025-12) y otra 9 columnas
        (sin Adiciones, ej. Pacífico 2025-12), cuadrando Saldo final exacto (377.501.685 y 380.541.644).
        """
        from src.core.rounding_reconciler import RoundingReconciler

        for emp, expected_adic, expected_cap, expected_int, expected_end in [
            ("Db Terra Chile Holdco SpA", 314725.0, -9711710.0, -34293536.0, 377501685.0),
            ("Pacifico Cable SpA", 0.0, -1308112.0, -393540.0, 3039959.0),
            ("[GRUPO] Consolidado DB Terra Holdco", 314725.0, -11019822.0, -34687076.0, 380541644.0),
        ]:
            with open("Plantilla de notas_v1.xlsx", "rb") as f_tpl:
                not_bytes = io.BytesIO(f_tpl.read())
            injected_io = ExternalNotesManager.inject_external_data(
                not_bytes, emp, "2026-09", "2025-12"
            )
            _, eval_io = RoundingReconciler.reconcile_and_evaluate_note_workbook(
                injected_io,
                empresa=emp,
                periodo_actual="2026-09",
                scale_factor=1000.0,
                target_sheets=["Pasivos financieros "],
                return_raw=False,
            )
            wb_ev = openpyxl.load_workbook(io.BytesIO(eval_io.getvalue()), data_only=True)
            ws_pf = wb_ev["Pasivos financieros "]

            comp_hdr_row = None
            for r in range(1, ws_pf.max_row + 1):
                if str(ws_pf.cell(r, 1).value or "").strip() == "Tipo":
                    last_hdr = str(ws_pf.cell(r, 11).value or "")
                    if "2025" in last_hdr:
                        comp_hdr_row = r
                        break
            self.assertIsNotNone(comp_hdr_row, f"No se encontró bloque comparativo 2025 de #N15.4 para {emp}")
            self.assertEqual(ws_pf.cell(comp_hdr_row, 8).value, "Adiciones")
            self.assertEqual(ws_pf.cell(comp_hdr_row, 9).value, "Capital")
            self.assertEqual(ws_pf.cell(comp_hdr_row, 10).value, "Intereses (1)")

            tot_row = None
            for r in range(comp_hdr_row + 1, min(comp_hdr_row + 12, ws_pf.max_row + 1)):
                if "total otros pasivos financieros" in str(ws_pf.cell(r, 1).value or "").strip().lower():
                    tot_row = r
                    break
            self.assertIsNotNone(tot_row)
            self.assertAlmostEqual(float(ws_pf.cell(tot_row, 8).value or 0.0), expected_adic, places=0)
            self.assertAlmostEqual(float(ws_pf.cell(tot_row, 9).value or 0.0), expected_cap, places=0)
            self.assertAlmostEqual(float(ws_pf.cell(tot_row, 10).value or 0.0), expected_int, places=0)
            self.assertAlmostEqual(float(ws_pf.cell(tot_row, 11).value or 0.0), expected_end, places=0)
            wb_ev.close()


if __name__ == "__main__":
    unittest.main()

