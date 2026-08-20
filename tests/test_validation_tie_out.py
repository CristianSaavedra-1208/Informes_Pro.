import unittest
import os
import sys
import pandas as pd

# Asegurar que el directorio raíz del proyecto esté en sys.path
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from src.models.database import SessionLocal, init_db
from src.core.validation_tie_out import ValidationTieOutEngine
from src.models.consolidacion import ConsolidationGroup

class TestValidationTieOut(unittest.TestCase):

    def setUp(self):
        init_db()
        self.db = SessionLocal()

    def tearDown(self):
        self.db.close()

    def test_validation_tie_out_grupo_consolidado(self):
        """
        Verifica que la matriz de validación Tie-Out responda correctamente para grupos consolidados.
        """
        grupos = self.db.query(ConsolidationGroup).all()
        if not grupos:
            self.skipTest("No se encontraron grupos de consolidación en la base de datos.")

        grp_name = grupos[0].nombre_grupo
        periodo = "2026-05"

        df_matrix, health = ValidationTieOutEngine.obtener_matriz_tie_out(
            empresa_o_grupo=grp_name,
            periodo=periodo,
            is_consolidated=True
        )

        self.assertIsInstance(df_matrix, pd.DataFrame, "La matriz Tie-Out debe ser un DataFrame.")
        self.assertIn("is_valid", health, "Debe incluir el estado global is_valid.")
        self.assertIn("total_descuadres", health, "Debe incluir el conteo de descuadres.")

        if not df_matrix.empty:
            self.assertIn("N° Nota", df_matrix.columns)
            self.assertEqual(df_matrix.columns[0], "N° Nota", "La columna N° Nota debe ser la primera a la izquierda.")
            self.assertIn("Reporte", df_matrix.columns)
            self.assertIn("Rubro Estado Financiero", df_matrix.columns)
            self.assertIn("Saldo EEFF ($)", df_matrix.columns)
            self.assertIn("Suma Nota ($)", df_matrix.columns)
            self.assertIn("Estado", df_matrix.columns)

            excel_bytes = ValidationTieOutEngine.generar_excel_tie_out(df_matrix, health, grp_name, periodo)
            self.assertGreater(len(excel_bytes), 0, "Debe generar los bytes del archivo Excel de Tie-Out.")

    def test_validation_tie_out_empresa_individual(self):
        """
        Verifica que la matriz de validación Tie-Out funcione para empresas individuales.
        """
        empresa_test = "Pacifico SpA"
        periodo = "2026-05"

        df_matrix, health = ValidationTieOutEngine.obtener_matriz_tie_out(
            empresa_o_grupo=empresa_test,
            periodo=periodo,
            is_consolidated=False
        )

        self.assertIsInstance(df_matrix, pd.DataFrame, "La matriz Tie-Out individual debe ser un DataFrame.")
        self.assertIn("is_valid", health, "Debe incluir el estado global is_valid.")
        if not df_matrix.empty:
            self.assertIn("N° Nota", df_matrix.columns)
            self.assertEqual(df_matrix.columns[0], "N° Nota", "La columna N° Nota debe ser la primera a la izquierda.")

print("[OK] Test de validación Tie-Out cargado correctamente.")

if __name__ == "__main__":
    unittest.main()
