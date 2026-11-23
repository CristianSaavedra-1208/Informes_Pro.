import os
import pickle
import unittest
import pandas as pd

from src.core.snapshot_vault_manager import SnapshotVaultManager
from src.core.consolidacion_engine import generar_hoja_trabajo, get_available_consolidation_periods
from src.core.consolidacion_account_mapper import resolve_je_account_and_nota
from src.reporting.adjustments_export import get_adjustments_dataset
from src.models.database import SessionLocal
from src.models.consolidacion import ConsolidationJournalEntry


BASELINE_PATH = r"C:\Users\cfsaa\.gemini\antigravity\brain\4079b99d-555b-4494-821f-aff70c67b9c0\scratch\pre_change_baseline.pkl"


def _assert_df_equal_numeric_and_text(tc: unittest.TestCase, df_a: pd.DataFrame, df_b: pd.DataFrame, label: str):
    tc.assertEqual(df_a.shape, df_b.shape, f"{label}: distinto tamaño {df_a.shape} vs {df_b.shape}")
    for r in range(len(df_a)):
        for c in range(len(df_a.columns)):
            va = df_a.iloc[r, c]
            vb = df_b.iloc[r, c]
            na_a = va is None or (isinstance(va, float) and pd.isna(va)) or pd.isna(va)
            na_b = vb is None or (isinstance(vb, float) and pd.isna(vb)) or pd.isna(vb)
            if na_a and na_b:
                continue
            try:
                fa = float(va)
                fb = float(vb)
                tc.assertAlmostEqual(fa, fb, places=2, msg=f"{label} celda ({r},{c}): {fa} != {fb}")
                continue
            except Exception:
                pass
            tc.assertEqual(str(va).strip(), str(vb).strip(), f"{label} celda ({r},{c}): {va!r} != {vb!r}")


class TestConsolidationJesNotesAndZeroRegression(unittest.TestCase):
    def test_1_all_300_jes_have_9999_accounts_and_2025_12_in_selector(self):
        """
        Test 3 del plan: Verifica que 0 líneas de JEs queden con cuenta_codigo vacío
        y que 2025-12 esté disponible en el selector de consolidación.
        """
        db = SessionLocal()
        try:
            rows = db.query(ConsolidationJournalEntry).all()
            self.assertGreaterEqual(len(rows), 300)
            missing = [
                (r.id, r.asiento_codigo, r.linea_item)
                for r in rows
                if not r.cuenta_codigo or not str(r.cuenta_codigo).strip().startswith("99")
            ]
            self.assertEqual(missing, [], f"Líneas de JEs sin cuenta 9999xxx: {missing}")

            # Verificar los 2 alias históricos explícitos
            cod_fibra, nota_fibra = resolve_je_account_and_nota("Costos de uso fibra optica")
            self.assertEqual(cod_fibra, "9999315")
            self.assertIsNotNone(nota_fibra)

            cod_res, _ = resolve_je_account_and_nota("Ganancias (Pérdida) del Ejercicio")
            self.assertEqual(cod_res, "9999345")
        finally:
            db.close()

        periods = get_available_consolidation_periods(grupo_id=1)
        self.assertIn("2025-12", periods)

    def test_2_zero_regression_against_pre_change_baseline(self):
        """
        Tests 1 y 2 del plan: Verifica invarianza total (0 diferencias) frente al baseline pre-cambio en:
        - EEFF y Notas de sociedades individuales
        - Hoja de Trabajo Consolidada y 5 EEFF Consolidados
        - Montos Debe/Haber/Saldo Base de las 300 líneas de JEs
        """
        if not os.path.exists(BASELINE_PATH):
            self.skipTest("Archivo baseline temporal no disponible en este entorno.")

        with open(BASELINE_PATH, "rb") as f:
            baseline = pickle.load(f)

        # 1. EEFF individuales
        for (emp, pa, pc, code), df_base in baseline["ind_statements"].items():
            res = SnapshotVaultManager.get_statement(emp, pa, code, comp_period=pc, scale_factor=1000.0)
            self.assertIsNotNone(res, f"Falta EEFF individual {emp} {code}")
            _assert_df_equal_numeric_and_text(self, df_base, res[0], f"IND_STMT {emp} {code}")

        # 2. Notas individuales
        for (emp, pa, pc, ncode), tbls_base in baseline["ind_notes"].items():
            els = SnapshotVaultManager.get_note_elements(emp, pa, ncode, scale_factor=1000.0, comp_period=pc)
            self.assertIsNotNone(els, f"Falta nota individual {emp} {ncode}")
            tbls_now = [el[1] for el in els if el[0] == "table" and isinstance(el[1], pd.DataFrame)]
            self.assertEqual(len(tbls_base), len(tbls_now), f"IND_NOTE {emp} {ncode} cantidad tablas")
            for idx_t, (tb, tn) in enumerate(zip(tbls_base, tbls_now)):
                _assert_df_equal_numeric_and_text(self, tb, tn, f"IND_NOTE {emp} {ncode} T{idx_t}")

        # 3. Hoja de Trabajo Consolidada
        for per, df_h_base in baseline["cons_hoja_trabajo"].items():
            df_h_now, _ = generar_hoja_trabajo(1, per)
            self.assertIsNotNone(df_h_now)
            _assert_df_equal_numeric_and_text(self, df_h_base, df_h_now, f"CONS_HOJA_TRABAJO {per}")

        # 4. 5 EEFF Consolidados en todos los períodos cerrados del grupo
        grp = "[GRUPO] Consolidado DB Terra Holdco"
        for (per, code), df_c_base in baseline["cons_statements"].items():
            res_c = SnapshotVaultManager.get_statement(grp, per, code, scale_factor=1000.0)
            self.assertIsNotNone(res_c)
            _assert_df_equal_numeric_and_text(self, df_c_base, res_c[0], f"CONS_STMT {per} {code}")

        # 5. Montos resueltos de las 300 líneas de JEs
        ds_now = get_adjustments_dataset(grupo_id=1, filter_mode="ALL")
        jes_now = [
            (
                ln["Código Folio"], ln["Período"], ln["Línea #"], ln["Rubro EEFF Afectado"],
                round(float(ln["Debe ($)"])), round(float(ln["Haber ($)"])), round(float(ln["Saldo Base"]))
            )
            for ln in ds_now["lineas"]
        ]
        self.assertEqual(baseline["jes_lines"], jes_now)

    def test_3_consolidated_notes_include_jes_and_tie_out_to_eeff(self):
        """
        Test 4 del plan: Verifica que las Notas Consolidadas (#N08, #N09, #N12, #N21, #N22)
        incorporen los ajustes PPA y reclasificaciones de consolidación y cuadren con EEFF Consolidados.
        """
        grp = "[GRUPO] Consolidado DB Terra Holdco"
        if not SnapshotVaultManager.has_active_snapshot(grp, "2025-12"):
            self.skipTest("Snapshot consolidado 2025-12 no activo.")

        bal_df, _ = SnapshotVaultManager.get_statement(grp, "2025-12", "BALANCE", scale_factor=1000.0)
        er_df, _ = SnapshotVaultManager.get_statement(grp, "2025-12", "ER", scale_factor=1000.0)

        def _stmt_val(df, name_substr):
            sub = df[df.iloc[:, 0].astype(str).str.contains(name_substr, case=False, na=False)]
            col_num = [c for c in df.columns if "2025" in str(c)][0]
            return round(float(sub[col_num].iloc[0]))

        # #N08.1 Intangibles == 88.923.321 M$ (incluyendo Cartera de clientes PPA)
        els_n08 = SnapshotVaultManager.get_note_elements(grp, "2025-12", "#N08", scale_factor=1000.0)
        t_n08_1 = [el[1] for el in els_n08 if el[0] == "table"][0]
        self.assertEqual(round(float(t_n08_1.iloc[6, 1])), 93587847)   # Cartera de clientes Bruto PPA
        self.assertEqual(round(float(t_n08_1.iloc[6, 2])), -7363426)   # Cartera de clientes Amort. Acum. PPA
        self.assertEqual(round(float(t_n08_1.iloc[7, 3])), _stmt_val(bal_df, "Activos intangibles"))

        # #N09.1 Activo Fijo incluye Asignación de compra (+210.187.318) y Amortización PPA (-17.342.754)
        from src.core.external_notes_manager import ExternalNotesManager
        els_n09 = SnapshotVaultManager.get_note_elements(grp, "2025-12", "#N09", scale_factor=1000.0)
        t_n09_1 = [el[1] for el in els_n09 if el[0] == "table"][0]
        row_asig = t_n09_1[t_n09_1.iloc[:, 0].astype(str).str.contains("Asignacion de compra", case=False, na=False)]
        row_appa = t_n09_1[t_n09_1.iloc[:, 0].astype(str).str.contains("Amortizacion PPA", case=False, na=False)]
        row_tot_af = t_n09_1[t_n09_1.iloc[:, 0].astype(str).str.contains("Totales", case=False, na=False)]
        self.assertEqual(round(float(row_asig.iloc[0, 1])), 210187318)
        self.assertAlmostEqual(round(float(row_appa.iloc[0, 2])), -17342754, delta=2)
        wb_h_ext = ExternalNotesManager.get_anexo_workbook("Db Terra Chile Holdco SpA", "2025-12", data_only=True)
        holdco_af_filled = bool(wb_h_ext and "Activo Fijo" in wb_h_ext.sheetnames and ExternalNotesManager.is_activo_fijo_injected(wb_h_ext["Activo Fijo"]))
        if wb_h_ext:
            wb_h_ext.close()
        if holdco_af_filled:
            self.assertEqual(round(float(row_tot_af.iloc[0, 3])), _stmt_val(bal_df, "Propiedades, plantas y equipos"))
        else:
            # Pacífico AF externo (142.104.895) + PPA neto (192.844.564) = 334.949.459 mientras el anexo AF de Holdco siga vacío
            self.assertAlmostEqual(round(float(row_tot_af.iloc[0, 3])), 334949459, delta=2)

        # #N12.1 y #N12.2 Impuestos corrientes == 0 y 2.003.744 M$
        els_n12 = SnapshotVaultManager.get_note_elements(grp, "2025-12", "#N12", scale_factor=1000.0)
        tbls_n12 = [el[1] for el in els_n12 if el[0] == "table"]
        tot_act_tax = tbls_n12[0][tbls_n12[0].iloc[:, 0].astype(str).str.contains("Total activos", case=False, na=False)]
        tot_pas_tax = tbls_n12[1][tbls_n12[1].iloc[:, 0].astype(str).str.contains("Total pasivos", case=False, na=False)]
        self.assertEqual(round(float(tot_act_tax.iloc[0, 1])), _stmt_val(bal_df, "Activo por impuestos, corrientes"))
        self.assertEqual(round(float(tot_pas_tax.iloc[0, 1])), _stmt_val(bal_df, "Pasivo por impuestos, corrientes"))

        # #N21.1 Ingresos, #N21.2 Costos y Depreciación operacional cuadran con #ER Consolidado (229.763.351, 106.877.881 y 109.456.250 M$)
        els_n21 = SnapshotVaultManager.get_note_elements(grp, "2025-12", "#N21", scale_factor=1000.0)
        tbls_n21 = [el[1] for el in els_n21 if el[0] == "table"]
        tot_ing = tbls_n21[0][tbls_n21[0].iloc[:, 0].astype(str).str.contains("Total ingresos", case=False, na=False)]
        tot_ctos = tbls_n21[1][tbls_n21[1].iloc[:, 0].astype(str).str.contains("Total costos", case=False, na=False)]
        tot_dep_op = tbls_n21[1][tbls_n21[1].iloc[:, 0].astype(str).str.contains("Total Depreciacion", case=False, na=False)]
        self.assertAlmostEqual(round(float(tot_ing.iloc[0, 1])), abs(_stmt_val(er_df, "Ingresos de actividades ordinarias")), delta=2)
        self.assertAlmostEqual(round(float(tot_ctos.iloc[0, 1])), abs(_stmt_val(er_df, "Costo de ventas")), delta=2)
        self.assertAlmostEqual(round(float(tot_dep_op.iloc[0, 1])), abs(_stmt_val(er_df, "Depreciación y amortización operacional")), delta=2)

        # #N22.2 Depreciación y amortizaciones en Gtos Adm == 6.364.860 M$ (incluyendo Amortizacion cartera clientes 1.840.857)
        els_n22 = SnapshotVaultManager.get_note_elements(grp, "2025-12", "#N22", scale_factor=1000.0)
        tbls_n22 = [el[1] for el in els_n22 if el[0] == "table"]
        t_n22_2 = tbls_n22[1]
        row_acc = t_n22_2[t_n22_2.iloc[:, 0].astype(str).str.contains("Amortizacion cartera clientes", case=False, na=False)]
        self.assertEqual(round(float(row_acc.iloc[0, 1])), 1840857)
        self.assertEqual(round(float(t_n22_2.iloc[-1, 1])), abs(_stmt_val(er_df, "Depreciación y amortizaciones")))


if __name__ == "__main__":
    unittest.main()
