import unittest
import pandas as pd
from src.core.sabana_builder import build_consolidated_balance_sabana, build_consolidated_pl_sabana

class TestConsolidatedSabana(unittest.TestCase):
    """
    Pruebas unitarias para verificar la generación de sábanas de auditoría consolidadas (Balance y P&L).
    """

    @classmethod
    def setUpClass(cls):
        cls.grupo_test = "[GRUPO] Consolidado DB Terra Holdco"
        cls.periodo_test = "2026-03"

    def test_01_build_consolidated_balance_sabana(self):
        """Verifica que la sábana consolidada de balance se genere con columnas de empresas y total."""
        df_sab = build_consolidated_balance_sabana(self.grupo_test, self.periodo_test)
        self.assertIsNotNone(df_sab, "❌ La sábana consolidada de balance no debe ser None")
        self.assertFalse(df_sab.empty, "❌ La sábana consolidada de balance no debe estar vacía")
        
        self.assertIn("N° de Cuenta", df_sab.columns)
        self.assertIn("TOTAL CONSOLIDADO", df_sab.columns)
        
        # Verificar que Pacifico SpA esté presente como columna
        self.assertIn("Pacifico SpA", df_sab.columns, "❌ Debe existir la columna 'Pacifico SpA' en la sábana consolidada")

    def test_02_build_consolidated_pl_sabana(self):
        """Verifica que la sábana consolidada de P&L se genere con columnas de empresas y total."""
        df_sab = build_consolidated_pl_sabana(self.grupo_test, self.periodo_test)
        self.assertIsNotNone(df_sab, "❌ La sábana consolidada de P&L no debe ser None")
        self.assertFalse(df_sab.empty, "❌ La sábana consolidada de P&L no debe estar vacía")
        
        self.assertIn("N° de Cuenta", df_sab.columns)
        self.assertIn("TOTAL CONSOLIDADO", df_sab.columns)
        self.assertIn("Pacifico SpA", df_sab.columns, "❌ Debe existir la columna 'Pacifico SpA' en la sábana consolidada")

if __name__ == '__main__':
    unittest.main()
