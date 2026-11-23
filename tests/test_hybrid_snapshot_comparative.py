import os
import pickle
import tempfile
import unittest
from io import BytesIO
from unittest.mock import patch

import numpy as np
import openpyxl
import pandas as pd

from src.core.snapshot_vault_manager import SnapshotVaultManager, _HYBRID_COMP_CACHE


class TestHybridSnapshotComparative(unittest.TestCase):
    """
    Verifica los escenarios solicitados para comparativos con períodos cerrados y abiertos:
    1) Actual Cerrado (Snapshot) vs Comparativo Abierto (calculado al vuelo) -> EEFF, Excel y Notas.
    2) Actual Cerrado (Snapshot) vs Comparativo Cerrado (Snapshot propio) -> extracción inmutable desde ambos Snapshots.
    3) Actual Abierto (al vuelo) vs Comparativo Cerrado (Snapshot) -> columna comparativa extraída del Snapshot.
    4) Comparativo 'Ninguno' -> columna comparativa limpia manteniendo período actual congelado.
    """

    EMP_GRUPO = "[GRUPO] Consolidado DB Terra Holdco"

    def setUp(self):
        _HYBRID_COMP_CACHE.clear()

    def test_closed_actual_vs_open_comp_er_and_excel_and_note(self):
        if not SnapshotVaultManager.has_active_snapshot(self.EMP_GRUPO, "2026-10"):
            self.skipTest("No existe snapshot activo 2026-10 para Consolidado DB Terra Holdco en este entorno.")

        # 1. Cargar Snapshot puro sin override (nativo de ER 2026-10 era 2025-10)
        frozen_only, _ = SnapshotVaultManager.get_statement(
            self.EMP_GRUPO, "2026-10", "ER", comp_period=None, scale_factor=1000.0
        )
        self.assertIsNotNone(frozen_only)
        self.assertEqual(frozen_only.columns[1], "2026-10")

        # 2. Solicitar comparativo abierto 2025-12 -> debe calcular 2025-12 al vuelo sin alterar 2026-10
        hybrid_er, _ = SnapshotVaultManager.get_statement(
            self.EMP_GRUPO, "2026-10", "ER", comp_period="2025-12", scale_factor=1000.0
        )
        self.assertIsNotNone(hybrid_er)
        self.assertEqual(list(hybrid_er.columns), ["Clasificación", "2026-10", "2025-12"])

        # Inmutabilidad estricta del período cerrado 2026-10
        pd.testing.assert_series_equal(
            hybrid_er["2026-10"],
            frozen_only["2026-10"],
            check_names=True,
        )

        # La columna 2025-12 calculada al vuelo debe tener datos reales (no ceros/NaN) y cuadrar verticalmente
        def _row_val(df: pd.DataFrame, label_sub: str, col: str) -> float:
            m = df[df.iloc[:, 0].astype(str).str.lower().str.strip() == label_sub.lower().strip()]
            self.assertFalse(m.empty, f"No se encontró fila '{label_sub}'")
            return float(m[col].iloc[0])

        ing_2025 = _row_val(hybrid_er, "Ingresos de actividades ordinarias", "2025-12")
        costo_2025 = _row_val(hybrid_er, "Costo de ventas", "2025-12")
        dep_op_2025 = _row_val(hybrid_er, "Depreciación y amortización operacional", "2025-12")
        gan_bruta_2025 = _row_val(hybrid_er, "Ganancia bruta", "2025-12")
        perdida_2025 = _row_val(hybrid_er, "Pérdida", "2025-12")

        self.assertGreater(abs(ing_2025), 100_000_000)
        self.assertNotEqual(perdida_2025, 0.0)
        self.assertEqual(
            int(round(ing_2025 + costo_2025 + dep_op_2025)),
            int(round(gan_bruta_2025)),
        )

        # 3. Excel exportado debe reflejar encabezado '31 de diciembre 2025' y los montos de 2025-12
        ex_bytes = SnapshotVaultManager.get_statement_excel(
            self.EMP_GRUPO,
            "2026-10",
            "ER",
            comp_period="2025-12",
            scale_factor=1000.0,
            df_override=hybrid_er,
        )
        self.assertIsNotNone(ex_bytes)
        wb = openpyxl.load_workbook(BytesIO(ex_bytes))
        ws = wb.active
        self.assertEqual(ws.cell(row=1, column=3).value, "31 de octubre 2026")
        self.assertEqual(ws.cell(row=1, column=4).value, "31 de diciembre 2025")
        self.assertEqual(int(ws.cell(row=4, column=4).value), int(round(ing_2025)))
        self.assertEqual(int(ws.cell(row=4, column=3).value), int(round(_row_val(hybrid_er, "Ingresos de actividades ordinarias", "2026-10"))))

        # 4. Nota P&L #N21 con 2026-10 cerrado vs 2025-12 abierto debe mostrar 31.12.2025 y cuadrar con ER
        els_n21_frozen = SnapshotVaultManager.get_note_elements(
            self.EMP_GRUPO, "2026-10", "#N21", scale_factor=1000.0, comp_period=None
        )
        els_n21_hybrid = SnapshotVaultManager.get_note_elements(
            self.EMP_GRUPO, "2026-10", "#N21", scale_factor=1000.0, comp_period="2025-12"
        )
        self.assertIsNotNone(els_n21_hybrid)
        tbl_frozen = next(v for t, v, *_ in els_n21_frozen if t == "table")
        tbl_hybrid = next(v for t, v, *_ in els_n21_hybrid if t == "table")

        # Columna 2026-10 inmutable
        pd.testing.assert_series_equal(tbl_hybrid.iloc[:, 1], tbl_frozen.iloc[:, 1], check_names=False)

        # Encabezado comparativo actualizado a 31.12.2025 y suma vertical exacta de la columna 2025-12
        hdr_texts_comp = " ".join(str(x) for x in tbl_hybrid.iloc[:3, 2].tolist())
        self.assertIn("31.12.2025", hdr_texts_comp)
        self.assertNotIn("31.10.2025", hdr_texts_comp)
        total_row_n21 = tbl_hybrid[tbl_hybrid.iloc[:, 0].astype(str).str.lower().str.contains("total")]
        self.assertFalse(total_row_n21.empty)
        total_ing_n21_2025 = int(round(float(total_row_n21.iloc[-1, 2])))
        detail_vals_2025 = [
            int(round(float(v)))
            for v in tbl_hybrid.iloc[3:-1, 2].tolist()
            if pd.notna(v) and str(v).strip() != ""
        ]
        self.assertGreater(total_ing_n21_2025, 200_000_000)
        self.assertEqual(sum(detail_vals_2025), total_ing_n21_2025)

        # 5. Excel de la Nota #N21 también refleja 31.12.2025 y el total calculado al vuelo
        from src.reporting.notes import get_full_note_registry
        sh_n21 = get_full_note_registry()["#N21"]["sheets"][0]
        ex_n21 = SnapshotVaultManager.get_note_excel(
            self.EMP_GRUPO, "2026-10", "#N21", comp_period="2025-12", scale_factor=1000.0
        )
        self.assertIsNotNone(ex_n21)
        wb_n21 = openpyxl.load_workbook(BytesIO(ex_n21), data_only=True)
        ws_n21 = wb_n21[sh_n21]
        flat_vals = [str(c.value) for row in ws_n21.iter_rows(max_row=20, max_col=6) for c in row if c.value is not None]
        self.assertTrue(any("31.12.2025" in v or "2025-12" in v for v in flat_vals))

    def test_closed_actual_vs_closed_comp_uses_both_snapshots(self):
        """
        Simula en un directorio temporal dos períodos cerrados (2026-10 y 2025-12) con saldos certificados
        y verifica que ambos se extraen estrictamente desde sus respectivos Snapshots (sin recalcular al vuelo).
        """
        with tempfile.TemporaryDirectory() as tmpdir:
            emp_test = "Empresa_Test_Unit"
            snap_2026 = os.path.join(tmpdir, "2026-10")
            snap_2025 = os.path.join(tmpdir, "2025-12")
            os.makedirs(os.path.join(snap_2026, "statements"), exist_ok=True)
            os.makedirs(os.path.join(snap_2025, "statements"), exist_ok=True)

            df_2026_er = pd.DataFrame({
                "Clasificación": ["Ingresos de actividades ordinarias", "Costo de ventas", "Ganancia bruta"],
                "2026-10": [500000, -200000, 300000],
                "2025-10": [0, 0, 0],
            })
            df_2025_er = pd.DataFrame({
                "Clasificación": ["Ingresos de actividades ordinarias", "Costo de ventas", "Ganancia bruta"],
                "2025-12": [777777, -333333, 444444],
                "2024-12": [100000, -50000, 50000],
            })

            with open(os.path.join(snap_2026, "statements", "er.pkl"), "wb") as f:
                pickle.dump((df_2026_er, "ESTADO DE RESULTADOS"), f)
            with open(os.path.join(snap_2025, "statements", "er.pkl"), "wb") as f:
                pickle.dump((df_2025_er, "ESTADO DE RESULTADOS"), f)

            meta_2026 = {"base_scale": 1000.0, "comp_period_bal": "2025-12", "comp_period_pl": "2025-10"}
            meta_2025 = {"base_scale": 1000.0, "comp_period_bal": "2024-12", "comp_period_pl": "2024-12"}

            def fake_has_snap(emp, per):
                return emp == emp_test and per in ("2026-10", "2025-12")

            def fake_get_dir(emp, per):
                return snap_2026 if per == "2026-10" else snap_2025

            def fake_get_meta(emp, per):
                return meta_2026 if per == "2026-10" else meta_2025

            with patch.object(SnapshotVaultManager, "has_active_snapshot", side_effect=fake_has_snap), \
                 patch.object(SnapshotVaultManager, "get_snapshot_dir", side_effect=fake_get_dir), \
                 patch.object(SnapshotVaultManager, "_get_snapshot_meta", side_effect=fake_get_meta), \
                 patch.object(SnapshotVaultManager, "_compute_live_statement_for_period") as mock_live:
                res_df, _ = SnapshotVaultManager.get_statement(
                    emp_test, "2026-10", "ER", comp_period="2025-12", scale_factor=1000.0
                )
                mock_live.assert_not_called()
                self.assertEqual(list(res_df.columns), ["Clasificación", "2026-10", "2025-12"])
                self.assertEqual(list(res_df["2026-10"]), [500000, -200000, 300000])
                self.assertEqual(list(res_df["2025-12"]), [777777, -333333, 444444])

                # Escenario 3: Actual abierto (live_df) + Comparativo cerrado (2025-12)
                df_open_2026_11 = pd.DataFrame({
                    "Clasificación": ["Ingresos de actividades ordinarias", "Costo de ventas", "Ganancia bruta"],
                    "2026-11": [610000, -250000, 360000],
                    "2025-12": [999999, -999999, 0],  # Valor vivo sin certificar que debe ser reemplazado por Snapshot
                })
                spliced = SnapshotVaultManager.apply_closed_comp_to_statement(
                    df_open_2026_11, emp_test, "ER", "2026-11", "2025-12", scale_factor=1000.0
                )
                self.assertEqual(list(spliced["2026-11"]), [610000, -250000, 360000])
                self.assertEqual(list(spliced["2025-12"]), [777777, -333333, 444444])

                # Escenario 4: Comparativo 'Ninguno'
                res_none, _ = SnapshotVaultManager.get_statement(
                    emp_test, "2026-10", "ER", comp_period="Ninguno", scale_factor=1000.0
                )
                self.assertEqual(list(res_none.columns), ["Clasificación", "2026-10", "Comp"])
                self.assertEqual(list(res_none["2026-10"]), [500000, -200000, 300000])
                self.assertTrue(res_none["Comp"].isna().all())

    def test_open_parent_or_subsidiary_consolidation_and_patrimonio_shield(self):
        """
        Verifica que si una sociedad del grupo tiene su período abierto en memoria activa
        (sin registros en HistoricalDataRecord), obtener_saldos_base / generar_hoja_trabajo
        calcule sus saldos al vuelo y el Estado de Cambios en el Patrimonio cuadre 2025-12 completo.
        """
        from src.core.consolidacion_engine import generar_hoja_trabajo
        from src.reporting.patrimonio_generator import get_patrimonio_balances_from_db

        df_h, _ = generar_hoja_trabajo(1, "2025-12")
        self.assertIsNotNone(df_h)
        row_cap = df_h[df_h["Balance clasificado"] == "Capital emitido"]
        row_ej = df_h[df_h["Balance clasificado"] == "Ganancias (Pérdida) del Ejercicio"]
        self.assertFalse(row_cap.empty)
        self.assertAlmostEqual(round(float(row_cap["CONSOLIDADO"].iloc[0]) / 1000.0), -503344995, delta=1)
        self.assertAlmostEqual(round(float(row_ej["CONSOLIDADO"].iloc[0]) / 1000.0), 33412509, delta=1)

        row_gan_h = df_h[df_h["Balance clasificado"] == "Resultados acumulados"]
        expected_gan_m = round(-float(row_gan_h["CONSOLIDADO"].iloc[0]) / 1000.0) if not row_gan_h.empty else -100817186

        cap, gan, res = get_patrimonio_balances_from_db("[GRUPO] Consolidado DB Terra Holdco", "2025-12")
        self.assertAlmostEqual(round(cap), 503344995, delta=1)
        if not SnapshotVaultManager.has_active_snapshot("[GRUPO] Consolidado DB Terra Holdco", "2025-12"):
            self.assertAlmostEqual(round(gan), expected_gan_m, delta=2)
        self.assertAlmostEqual(round(res), 178355741, delta=1)

    def test_closed_actual_note_recomputes_open_native_comp_with_new_anexo(self):
        """
        Verifica que cuando el período actual está cerrado (ej. 2026-10 en Holdco) pero su
        período comparativo nativo (2025-12) está ABIERTO y recibe un nuevo Anexo Externo,
        get_note_elements recalcule la columna 2025-12 al vuelo preservando intacto 2026-10.
        """
        emp = "Db Terra Chile Holdco SpA"
        if not SnapshotVaultManager.has_active_snapshot(emp, "2026-10"):
            self.skipTest("Db Terra Chile Holdco SpA no tiene snapshot activo 2026-10.")
        if SnapshotVaultManager.has_active_snapshot(emp, "2025-12"):
            self.skipTest("Db Terra Chile Holdco SpA 2025-12 está cerrado; este test valida 2025-12 abierto.")

        els = SnapshotVaultManager.get_note_elements(emp, "2026-10", "#N07", scale_factor=1000.0, comp_period="2025-12")
        self.assertIsNotNone(els)
        tables = [el[1] for el in els if el[0] == "table"]
        self.assertGreaterEqual(len(tables), 2)
        # En #N07.1, Otras bodegas (fila 4 col 2) y #N07.2 Saldo inicial (tabla 1 fila 3 col 2) deben traer el anexo 2025-12
        t_inv = tables[0]
        self.assertEqual(round(float(t_inv.iloc[3, 1])), 32391252)
        self.assertEqual(round(float(t_inv.iloc[4, 2])), 1730997)

    def test_intangibles_stacked_comparative_blocks_and_movements(self):
        """
        Verifica que #N08 (Intangibles) en bloques verticales apilados separe correctamente
        el período actual y el comparativo (tanto entre años distintos 2026-10 vs 2025-12
        como dentro del mismo año 2026-10 vs 2026-06), preserve la cabecera 'Licencias' en #N08.2
        y cuadre #N08.2 contra #N08.1.
        """
        emp = "Db Terra Chile Holdco SpA"
        if not SnapshotVaultManager.has_active_snapshot(emp, "2026-10"):
            self.skipTest("Db Terra Chile Holdco SpA no tiene snapshot activo 2026-10.")

        # 1. Comparación inter-anual: 2026-10 vs 2025-12
        els_25 = SnapshotVaultManager.get_note_elements(emp, "2026-10", "#N08", scale_factor=1000.0, comp_period="2025-12")
        self.assertIsNotNone(els_25)
        tables_25 = [el[1] for el in els_25 if el[0] == "table"]
        self.assertEqual(len(tables_25), 2)

        n08_1 = tables_25[0]
        self.assertEqual(str(n08_1.iloc[0, 0]).strip(), "31.10.2026")
        self.assertEqual(str(n08_1.iloc[9, 0]).strip(), "31.12.2025")
        self.assertEqual(round(float(n08_1.iloc[7, 3])), 3472288)   # Total Neto 31.10.2026
        self.assertEqual(round(float(n08_1.iloc[16, 3])), 1530575)  # Total Neto 31.12.2025

        n08_2 = tables_25[1]
        self.assertEqual(str(n08_2.iloc[1, 2]).strip(), "Licencias")
        self.assertEqual(str(n08_2.iloc[9, 2]).strip(), "Licencias")
        self.assertEqual(round(float(n08_2.iloc[3, 5])), 1530575)   # Inicio 2026 = Cierre 2025
        self.assertEqual(round(float(n08_2.iloc[4, 5])), 3015368)   # Adiciones 2026
        self.assertEqual(round(float(n08_2.iloc[5, 5])), -1073655)  # Amortización 2026
        self.assertEqual(round(float(n08_2.iloc[6, 5])), 3472288)   # Cierre 2026 == #N08.1

        # 2. Comparación intra-anual (mismo año): 2026-10 vs 2026-06
        if SnapshotVaultManager.has_active_snapshot(emp, "2026-06"):
            els_26 = SnapshotVaultManager.get_note_elements(emp, "2026-10", "#N08", scale_factor=1000.0, comp_period="2026-06")
            self.assertIsNotNone(els_26)
            tables_26 = [el[1] for el in els_26 if el[0] == "table"]
            self.assertEqual(str(tables_26[0].iloc[0, 0]).strip(), "31.10.2026")
            self.assertEqual(str(tables_26[0].iloc[9, 0]).strip(), "30.06.2026")
            self.assertEqual(round(float(tables_26[0].iloc[7, 3])), 3472288)
            self.assertEqual(round(float(tables_26[0].iloc[16, 3])), 3013718)
            # Además, como 2025-12 existe como año previo de 2026-06, el bloque comparativo de #N08.2 también cuadra
            self.assertEqual(round(float(tables_26[1].iloc[14, 5])), 3013718)


if __name__ == "__main__":
    unittest.main()

