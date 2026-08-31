import unittest
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.core.consolidacion_lock_manager import ConsolidationLockManager
from src.core.consolidacion_engine import generar_hoja_trabajo
from src.models.database import SessionLocal
from src.models.consolidacion import ConsolidationGroup, ConsolidationJournalEntry, ConsolidationPeriodLock

class TestConsolidationLocks(unittest.TestCase):
    def setUp(self):
        self.grupo_id = 1
        self.periodo_test = "2025-12"
        self.periodo_open = "2026-07"
        
    def test_01_lock_and_unlock_cycle(self):
        """Valida que el ciclo de bloqueo y reapertura funcione perfectamente."""
        # 1. Bloquear
        res, msg = ConsolidationLockManager.lock_period(self.grupo_id, self.periodo_test, user="TestAuditor")
        self.assertTrue(res)
        self.assertTrue(ConsolidationLockManager.is_period_locked(self.grupo_id, self.periodo_test))
        
        # 2. Verificar lista de bloqueados
        locked_list = ConsolidationLockManager.get_locked_periods(self.grupo_id)
        self.assertTrue(any(l["periodo"] == self.periodo_test for l in locked_list))
        
        # 3. Desbloquear
        res_un, msg_un = ConsolidationLockManager.unlock_period(self.grupo_id, self.periodo_test, user="TestAuditor")
        self.assertTrue(res_un)
        self.assertFalse(ConsolidationLockManager.is_period_locked(self.grupo_id, self.periodo_test))

    def test_02_numerical_integrity_locked_vs_unlocked(self):
        """Valida que la hoja de consolidación produzca exactamente los mismos números con el período abierto o cerrado."""
        # Hoja con período desbloqueado
        ConsolidationLockManager.unlock_period(self.grupo_id, self.periodo_test)
        df_unlocked, _ = generar_hoja_trabajo(self.grupo_id, self.periodo_test)
        self.assertIsNotNone(df_unlocked)
        
        # Hoja con período bloqueado
        ConsolidationLockManager.lock_period(self.grupo_id, self.periodo_test)
        df_locked, _ = generar_hoja_trabajo(self.grupo_id, self.periodo_test)
        self.assertIsNotNone(df_locked)
        
        import pandas as pd
        # Comparar celda por celda la columna CONSOLIDADO
        for i in range(len(df_unlocked)):
            val_un = pd.to_numeric(df_unlocked.iloc[i]["CONSOLIDADO"], errors='coerce')
            val_lk = pd.to_numeric(df_locked.iloc[i]["CONSOLIDADO"], errors='coerce')
            if pd.notna(val_un) and pd.notna(val_lk):
                self.assertAlmostEqual(float(val_un), float(val_lk), places=2)
                
        # Limpieza (dejar desbloqueado para pruebas generales)
        ConsolidationLockManager.unlock_period(self.grupo_id, self.periodo_test)

    def test_03_available_periods_detection(self):
        """Valida que se detecten los períodos disponibles para el grupo."""
        available = ConsolidationLockManager.get_available_periods_for_group(self.grupo_id)
        self.assertIn("2026-07", available)
        self.assertIn("2025-12", available)

if __name__ == '__main__':
    unittest.main()
