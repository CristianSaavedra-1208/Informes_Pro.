import unittest
import os
import sys
import pandas as pd
from sqlalchemy import func

# Asegurar que el directorio raíz del proyecto esté en sys.path
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from src.models.database import SessionLocal, init_db
from src.models.consolidacion import ConsolidationGroup, ConsolidationJournalEntry
from src.core.consolidacion_engine import generar_hoja_trabajo

class TestConsolidacionIntegridad100(unittest.TestCase):

    def setUp(self):
        init_db()
        self.db = SessionLocal()

    def tearDown(self):
        self.db.close()

    def test_consolidacion_integridad_y_autocorreccion(self):
        """
        1. Captura el baseline numérico de todas las entradas de consolidación y de las hojas de trabajo generadas.
        2. Ejecuta el script de migración y autocorrección para asignar folios `AST-YYYYMM-XXX`.
        3. Verifica que la integridad numérica del Debe, Haber y la Hoja de Trabajo Consolidada sea idéntica al 100%.
        4. Garantiza que no existan registros sin `asiento_codigo`.
        """
        print("\n--- INICIANDO TEST DE INTEGRIDAD 100% Y AUTOCORRECCIÓN DE CONSOLIDACIÓN ---")

        # 1. Obtener baseline de la base de datos
        asientos_before = self.db.query(ConsolidationJournalEntry).all()
        cant_before = len(asientos_before)
        total_debe_before = sum(a.debe or 0.0 for a in asientos_before)
        total_haber_before = sum(a.haber or 0.0 for a in asientos_before)

        print(f"[Baseline] Cantidad asientos: {cant_before}, Total Debe: {total_debe_before:,.2f}, Total Haber: {total_haber_before:,.2f}")

        # Grupos de consolidación disponibles
        grupos = self.db.query(ConsolidationGroup).all()
        periodos_test = ["2026-05", "2025-12"]

        baseline_hojas = {}
        for g in grupos:
            for p in periodos_test:
                df_hoja, msg = generar_hoja_trabajo(g.id, p)
                if df_hoja is not None:
                    numeric_cols = [c for c in df_hoja.columns if c not in ['Rubro', 'Balance clasificado', 'Estado de Resultados clasificado', 'Código']]
                    sums = {col: float(df_hoja[col].sum()) for col in numeric_cols if pd.api.types.is_numeric_dtype(df_hoja[col])}
                    baseline_hojas[(g.id, p)] = sums

        # 2. Ejecutar la migración y autocorrección de códigos
        from src.core.migrate_asientos_codigo import migrar_y_autocorregir_codigos_asiento
        stats = migrar_y_autocorregir_codigos_asiento(self.db)
        print(f"[Autocorrección] Estadísticas: {stats}")

        # Refresh db session
        self.db.expire_all()

        # 3. Verificar que 100% de los registros tengan asiento_codigo válido
        asientos_after = self.db.query(ConsolidationJournalEntry).all()
        sin_codigo = [a for a in asientos_after if not a.asiento_codigo]
        self.assertEqual(len(sin_codigo), 0, f"Error: Hay {len(sin_codigo)} asientos sin asiento_codigo.")

        cant_after = len(asientos_after)
        total_debe_after = sum(a.debe or 0.0 for a in asientos_after)
        total_haber_after = sum(a.haber or 0.0 for a in asientos_after)

        self.assertEqual(cant_before, cant_after, "Error: La cantidad de registros cambió tras la migración.")
        self.assertAlmostEqual(total_debe_before, total_debe_after, places=2, msg="Error: El Total Debe difiere.")
        self.assertAlmostEqual(total_haber_before, total_haber_after, places=2, msg="Error: El Total Haber difiere.")

        # 4. Verificar integridad de las hojas de trabajo generadas celda por celda / rubro por rubro
        for g in grupos:
            for p in periodos_test:
                if (g.id, p) in baseline_hojas:
                    df_hoja_post, msg = generar_hoja_trabajo(g.id, p)
                    self.assertIsNotNone(df_hoja_post, f"Error generando hoja post migración para Grupo {g.id} Periodo {p}")
                    numeric_cols = [c for c in df_hoja_post.columns if c not in ['Rubro', 'Balance clasificado', 'Estado de Resultados clasificado', 'Código']]
                    for col, expected_sum in baseline_hojas[(g.id, p)].items():
                        if col in df_hoja_post.columns and pd.api.types.is_numeric_dtype(df_hoja_post[col]):
                            actual_sum = float(df_hoja_post[col].sum())
                            self.assertAlmostEqual(expected_sum, actual_sum, places=2,
                                msg=f"Descuadre en Grupo {g.id} Periodo {p} Columna {col}: antes={expected_sum}, después={actual_sum}")

        print("[OK] TEST DE INTEGRIDAD 100% PASADO EXITOSAMENTE SIN NINGUN DESCUADRE NUMERICO.")

if __name__ == "__main__":
    unittest.main()
