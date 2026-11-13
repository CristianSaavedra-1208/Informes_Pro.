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

        # 1. Inventarios (#N07.1)
        ws_inv = wb["Inventarios"]
        inv_vals = [ws_inv.cell(r, 2).value or 0 for r in range(6, 13)]
        self.assertEqual(round(sum(inv_vals)), 33247439)

        # 2. Activo Fijo (#N09.1)
        ws_af = wb["Activo Fijo"]
        self.assertEqual(ws_af.cell(5, 2).value, "Activo fijo bruto")
        self.assertEqual(ws_af.cell(5, 3).value, "Depreciación acumulada")
        self.assertEqual(ws_af.cell(5, 4).value, "Activo fijo neto")
        bruto_sum = sum(ws_af.cell(r, 2).value or 0 for r in range(8, 24))
        deprec_sum = sum(ws_af.cell(r, 3).value or 0 for r in range(8, 24))
        self.assertEqual(round(bruto_sum), 372803055)
        self.assertEqual(round(deprec_sum), -230698160)
        self.assertEqual(round(bruto_sum + deprec_sum), 142104895)

        # 3. Efectivo — Detalle Fondos Mutuos (#N04.3)
        ws_ef = wb["Efectivo"]
        fm_vals = [ws_ef.cell(r, 4).value or 0 for r in range(7, 13) if isinstance(ws_ef.cell(r, 4).value, (int, float))]
        self.assertEqual(len(fm_vals), 5)
        self.assertEqual(round(sum(fm_vals)), 7032177)

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
        self.assertEqual(round(cte_16_1), 4847998)
        self.assertEqual(round(ncte_16_1), 5820163)

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


if __name__ == "__main__":
    unittest.main()

