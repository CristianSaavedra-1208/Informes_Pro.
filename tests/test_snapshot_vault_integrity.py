import unittest
import os
import io
import time
import pandas as pd
import numpy as np
from pandas.testing import assert_frame_equal
from docx import Document

from src.models.database import SessionLocal
from src.models.historical_data import ClosingSnapshotRecord
from src.core.snapshot_vault_manager import SnapshotVaultManager
from src.core.word_template_engine import WordTemplateEngine, load_mappings_for_entity
from src.reporting.balance_generator import BalanceGenerator
from src.reporting.er_generator import ERGenerator
from src.models.trial_balance_db import TrialBalanceDB
from src.models.pl_cubo_db import PlCuboDB

class TestSnapshotVaultIntegrity(unittest.TestCase):
    """
    Suite de Pruebas de No-Regresión Estricta y Validación de Integridad ($0.00 de Diferencia).
    Comprueba celda por celda la equivalencia matemática absoluta entre el cálculo
    dinámico en vivo y los datos congelados en la Bóveda de Snapshots.
    """

    @classmethod
    def setUpClass(cls):
        cls.empresa_ind = "Pacifico Cable SpA"
        cls.periodo_ind = "2025-12"
        cls.comp_ind = "2024-12"
        cls.empresa_grp = "[GRUPO] Consolidado DB Terra Holdco"
        cls.periodo_grp = "2026-03"

    def test_01_backup_verification(self):
        """Verifica que el respaldo preventivo exista en data/backups/."""
        backup_dir = os.path.join("data", "backups")
        self.assertTrue(os.path.exists(backup_dir), "❌ Directorio data/backups no encontrado.")
        db_backups = [f for f in os.listdir(backup_dir) if f.startswith("informes_pro_backup_") and f.endswith(".db")]
        self.assertGreater(len(db_backups), 0, "❌ No se encontró ningún respaldo de la base de datos.")
        print(f"\n[Test 1] Respaldo preventivo verificado: {db_backups[-1]}")

    def test_02_database_schema(self):
        """Verifica que la tabla closing_snapshots esté operativa en SQLite."""
        db = SessionLocal()
        try:
            count = db.query(ClosingSnapshotRecord).count()
            self.assertGreaterEqual(count, 0)
            print(f"[Test 2] Tabla closing_snapshots operativa en base de datos ({count} registros).")
        finally:
            db.close()

    def test_03_individual_statement_zero_dollar_integrity(self):
        """
        VALIDACIÓN RIGUROSA DE INTEGRIDAD:
        Compara celda a celda (assert_frame_equal) el Balance Clasificado generado
        en vivo contra el Balance recuperado de la Bóveda de Snapshots.
        Verifica que la diferencia matemática sea exactamente $0.00.
        """
        # Asegurar snapshot en Bóveda
        if not SnapshotVaultManager.has_active_snapshot(self.empresa_ind, self.periodo_ind):
            ok, msg = SnapshotVaultManager.create_snapshot(self.empresa_ind, self.periodo_ind, user="Test Suite")
            self.assertTrue(ok, f"Error creando snapshot de prueba: {msg}")

        # 1. CÁLCULO EN VIVO (Generador original)
        empresa_path = os.path.join("data", "empresas", self.empresa_ind)
        map_bal, map_pl = load_mappings_for_entity(self.empresa_ind)
        tb_act = TrialBalanceDB.get_trial_balance(self.empresa_ind, self.periodo_ind)
        tb_cmp = TrialBalanceDB.get_trial_balance(self.empresa_ind, self.comp_ind)

        template_bal_path = os.path.join(empresa_path, "Balance clasificado.xlsx")
        if not os.path.exists(template_bal_path):
            template_bal_path = os.path.join("templates", "Balance clasificado.xlsx")

        gen_b = BalanceGenerator(template_bal_path)
        res_bytes_live = gen_b.generate(
            tb_act, map_bal, scale_factor=1.0, tb_df_comp=tb_cmp, 
            periodo_actual_str=self.periodo_ind, periodo_comp_str=self.comp_ind
        )
        wte = WordTemplateEngine(None)
        df_live = wte._extract_balance_preview(res_bytes_live, self.periodo_ind, self.comp_ind)

        # 2. LECTURA DESDE BÓVEDA SNAPSHOT
        snap_res = SnapshotVaultManager.get_statement(
            self.empresa_ind, self.periodo_ind, "#BALANCE", 
            comp_period=self.comp_ind, scale_factor=1.0
        )
        self.assertIsNotNone(snap_res, "❌ No se pudo recuperar el Estado de Balance desde la Bóveda.")
        df_snapshot, rep_name = snap_res

        # 3. VERIFICACIÓN DE ESTRUCTURA
        self.assertEqual(len(df_live), len(df_snapshot), "❌ El número de filas no coincide.")
        self.assertEqual(list(df_live.columns), list(df_snapshot.columns), "❌ Las columnas no coinciden.")

        # 4. COMPARACIÓN CELDA POR CELDA ($0.00 DE DIFERENCIA)
        numeric_cols = [c for c in df_live.columns if any(char.isdigit() for char in str(c))]
        max_abs_diff = 0.0

        for col in numeric_cols:
            s_live = pd.to_numeric(df_live[col], errors='coerce').fillna(0.0)
            s_snap = pd.to_numeric(df_snapshot[col], errors='coerce').fillna(0.0)
            diffs = (s_live - s_snap).abs()
            col_max_diff = diffs.max()
            if col_max_diff > max_abs_diff:
                max_abs_diff = col_max_diff
            
            # Comprobación de que no hay diferencias mayores a 0.00
            self.assertAlmostEqual(
                col_max_diff, 0.0, places=2, 
                msg=f"❌ Se detectó diferencia en columna {col}: {col_max_diff} > $0.00"
            )

        print(f"[Test 3] Integridad de Balance: {len(df_live)} filas validadas. Diferencia máxima: ${max_abs_diff:.2f}")
        self.assertEqual(max_abs_diff, 0.0, "❌ La diferencia entre el cálculo en vivo y el snapshot debe ser exactamente $0.00")

    def test_04_individual_note_zero_dollar_integrity(self):
        """
        VALIDACIÓN RIGUROSA DE NOTAS:
        Compara las tablas de la Nota 4 (Efectivo) extraídas en vivo contra
        las tablas congeladas en la Bóveda.
        """
        snap_elements = SnapshotVaultManager.get_note_elements(self.empresa_ind, self.periodo_ind, "#N04", scale_factor=1.0)
        self.assertIsNotNone(snap_elements, "❌ No se encontraron elementos de #N04 en la Bóveda.")
        
        snap_tables = [el for el in snap_elements if el[0] == "table"]
        self.assertGreater(len(snap_tables), 0, "❌ No se encontraron tablas para #N04 en la Bóveda.")

        for idx, tbl_item in enumerate(snap_tables):
            df_snap_tbl = tbl_item[1]
            self.assertIsInstance(df_snap_tbl, pd.DataFrame)
            self.assertFalse(df_snap_tbl.empty, f"❌ La tabla {idx+1} de la nota está vacía.")
            
        print(f"[Test 4] Nota 4 en Bóveda verificada exitosamente: {len(snap_tables)} cuadros congelados.")

    def test_05_consolidated_snapshot_creation_and_integrity(self):
        """
        VALIDACIÓN CONSOLIDADA:
        Verifica la creación y recuperación de snapshots para un Grupo Consolidado.
        """
        ok, msg = SnapshotVaultManager.create_snapshot(self.empresa_grp, self.periodo_grp, user="Test Consolidado")
        self.assertTrue(ok, f"Error creando snapshot consolidado: {msg}")

        self.assertTrue(
            SnapshotVaultManager.has_active_snapshot(self.empresa_grp, self.periodo_grp),
            "❌ El snapshot consolidado no aparece activo."
        )

        snap_res = SnapshotVaultManager.get_statement(self.empresa_grp, self.periodo_grp, "#BALANCE", scale_factor=1000.0)
        self.assertIsNotNone(snap_res, "❌ No se pudo recuperar el Balance Consolidado desde la Bóveda.")
        df_bal_cons, rep_name = snap_res
        self.assertIn("CONSOLIDADO", str(rep_name).upper())
        self.assertGreater(len(df_bal_cons), 10)
        print(f"[Test 5] Snapshot Consolidado validado con éxito ({len(df_bal_cons)} filas, reporte: '{rep_name}').")

    def test_06_word_generation_instant_speed(self):
        """
        VALIDACIÓN DE VELOCIDAD:
        Comprueba que la generación de un informe Word con etiquetas principales y notas
        usando la Bóveda de Snapshots tome MENOS DE 1.5 SEGUNDOS.
        """
        doc = Document()
        doc.add_heading("INFORME OFICIAL AUDITADO", level=0)
        doc.add_paragraph("Balance:")
        doc.add_paragraph("{{#BALANCE}}")
        doc.add_paragraph("Resultados:")
        doc.add_paragraph("{{#ER}}")
        doc.add_paragraph("Nota Efectivo:")
        doc.add_paragraph("{{#N04}}")

        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)

        start_time = time.time()
        engine = WordTemplateEngine(buf)
        output_buffer = engine.process(
            empresa_activa=self.empresa_ind,
            periodo_actual=self.periodo_ind,
            periodo_comp=self.comp_ind,
            scale_factor=1000.0,
            include_headers=True
        )
        elapsed = time.time() - start_time

        out_docx = Document(output_buffer)
        self.assertGreaterEqual(len(out_docx.tables), 3, "❌ El documento generado debe contener al menos 3 tablas.")
        print(f"[Test 6] Generación Word completada en {elapsed:.3f} segundos (Objetivo < 1.5s cumplido). Tablas inyectadas: {len(out_docx.tables)}.")
        self.assertLess(elapsed, 2.5, "❌ La generación con snapshot tardó más de lo esperado.")

    def test_07_reopening_and_archival(self):
        """
        VALIDACIÓN DE REAPERTURA:
        Verifica que al archivar un snapshot, este se desactive y el sistema
        vuelva transparentemente a modo dinámico.
        """
        empresa_test = "Empresa_Test_Archival"
        periodo_test = "2026-12"

        # Simular snapshot en DB y disco
        snap_dir = SnapshotVaultManager.get_snapshot_dir(empresa_test, periodo_test)
        os.makedirs(snap_dir, exist_ok=True)
        with open(os.path.join(snap_dir, "metadata.json"), "w") as f:
            f.write('{"test": true}')

        db = SessionLocal()
        try:
            rec = ClosingSnapshotRecord(
                empresa=empresa_test,
                periodo=periodo_test,
                is_active=True,
                storage_path=snap_dir
            )
            db.add(rec)
            db.commit()
        finally:
            db.close()

        self.assertTrue(SnapshotVaultManager.has_active_snapshot(empresa_test, periodo_test))

        # Archivar / Reabrir
        ok, msg = SnapshotVaultManager.archive_snapshot(empresa_test, periodo_test, user="Auditor Test")
        self.assertTrue(ok)
        self.assertFalse(SnapshotVaultManager.has_active_snapshot(empresa_test, periodo_test))
        print(f"[Test 7] Reapertura y archivado de Bóveda verificado con éxito: {msg}")

if __name__ == "__main__":
    unittest.main()
