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
        """Verifica que la sábana consolidada de balance se genere con columnas de empresas, ajustes mensuales YTD, total ajustes y total consolidado."""
        df_sab = build_consolidated_balance_sabana(self.grupo_test, self.periodo_test)
        self.assertIsNotNone(df_sab, "❌ La sábana consolidada de balance no debe ser None")
        self.assertFalse(df_sab.empty, "❌ La sábana consolidada de balance no debe estar vacía")
        
        self.assertIn("N° de Cuenta", df_sab.columns)
        self.assertIn("Nombre de la Cuenta", df_sab.columns)
        self.assertIn("Db Terra Chile Holdco SpA", df_sab.columns)
        self.assertIn("Pacifico Cable SpA", df_sab.columns)
        
        # Verificar columnas de ajustes mensuales YTD
        self.assertIn("Ajustes 2026-Ene", df_sab.columns)
        self.assertIn("Ajustes 2026-Feb", df_sab.columns)
        self.assertIn("Ajustes 2026-Mar", df_sab.columns)
        self.assertIn("TOTAL AJUSTES", df_sab.columns)
        self.assertIn("TOTAL CONSOLIDADO", df_sab.columns)
        
        # Verificar presencia de cuentas de consolidación (serie 9999xxx)
        acc_ids = df_sab["N° de Cuenta"].astype(str).tolist()
        has_99 = any(a.startswith("99") for a in acc_ids)
        self.assertTrue(has_99, "❌ Deben existir cuentas de consolidación (serie 9999xxx) en la sábana")

    def test_02_build_consolidated_pl_sabana(self):
        """Verifica que la sábana consolidada de P&L se genere con columnas de empresas, ajustes mensuales YTD, total ajustes y total consolidado."""
        df_sab = build_consolidated_pl_sabana(self.grupo_test, self.periodo_test)
        self.assertIsNotNone(df_sab, "❌ La sábana consolidada de P&L no debe ser None")
        self.assertFalse(df_sab.empty, "❌ La sábana consolidada de P&L no debe estar vacía")
        
        self.assertIn("N° de Cuenta", df_sab.columns)
        self.assertIn("Nombre de la Cuenta", df_sab.columns)
        self.assertIn("Db Terra Chile Holdco SpA", df_sab.columns)
        self.assertIn("Pacifico Cable SpA", df_sab.columns)
        
        # Verificar columnas de ajustes mensuales YTD
        self.assertIn("Ajustes 2026-Ene", df_sab.columns)
        self.assertIn("Ajustes 2026-Feb", df_sab.columns)
        self.assertIn("Ajustes 2026-Mar", df_sab.columns)
        self.assertIn("TOTAL AJUSTES", df_sab.columns)
        self.assertIn("TOTAL CONSOLIDADO", df_sab.columns)
        
        # Verificar presencia de cuentas de consolidación (serie 9999xxx)
        acc_ids = df_sab["N° de Cuenta"].astype(str).tolist()
        has_99 = any(a.startswith("99") for a in acc_ids)
        self.assertTrue(has_99, "❌ Deben existir cuentas de consolidación (serie 9999xxx) en la sábana de P&L")

if __name__ == '__main__':
    unittest.main()
