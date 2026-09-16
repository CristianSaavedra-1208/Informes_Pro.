import unittest
import os
import pandas as pd
from src.core.consolidacion_account_mapper import (
    normalize_text,
    load_consolidation_account_map,
    get_consolidation_accounts_list,
    get_account_info,
    find_account_by_rubro_nota
)

class TestConsolidationAccountMapper(unittest.TestCase):
    
    def test_normalize_text(self):
        self.assertEqual(normalize_text("Pérdida (Ganancia) del Ejercicio"), "perdida (ganancia) del ejercicio")
        self.assertEqual(normalize_text("  Activos   Intangibles  "), "activos intangibles")
        self.assertEqual(normalize_text(None), "")
        self.assertEqual(normalize_text(float('nan')), "")

    def test_load_consolidation_account_map(self):
        acc_map = load_consolidation_account_map("Pacifico Cable SpA")
        self.assertIsInstance(acc_map, dict)
        self.assertGreater(len(acc_map), 20, "Deberían cargarse al menos 20 cuentas de consolidación")

        # Verificar que todas las cuentas comiencen con 99
        for cod, info in acc_map.items():
            self.assertTrue(cod.startswith('99'), f"Cuenta {cod} debe comenzar con '99'")
            self.assertIn("linea_item", info)
            self.assertIn("tipo", info)
            self.assertIn("display_label", info)
            self.assertTrue(len(info["linea_item"]) > 0)

    def test_known_accounts_balance_and_pl(self):
        acc_map = load_consolidation_account_map("Pacifico Cable SpA")
        
        # 9999010 - Efectivo
        if "9999010" in acc_map:
            efectivo = acc_map["9999010"]
            self.assertEqual(efectivo["tipo"], "Balance")
            self.assertIn("Efectivo", efectivo["linea_item"])

        # 9999050 - Intangibles con nota
        if "9999050" in acc_map:
            intangibles = acc_map["9999050"]
            self.assertEqual(intangibles["tipo"], "Balance")
            self.assertIn("intangibles", intangibles["linea_item"].lower())
            self.assertEqual(intangibles.get("linea_nota"), "Cartera de clientes")

        # P&L account (ej. 9999300 - Costo de ventas o 9999330)
        pl_accounts = [a for a in acc_map.values() if a["tipo"] == "P&L"]
        self.assertGreater(len(pl_accounts), 5, "Debe haber cuentas de resultado (P&L)")

    def test_get_consolidation_accounts_list(self):
        acc_list = get_consolidation_accounts_list("Pacifico Cable SpA")
        self.assertIsInstance(acc_list, list)
        self.assertGreater(len(acc_list), 20)
        
        # Verificar que esté ordenada por código
        codigos = [a["codigo"] for a in acc_list]
        self.assertEqual(codigos, sorted(codigos))

    def test_get_account_info(self):
        info = get_account_info("9999010", "Pacifico Cable SpA")
        self.assertIsNotNone(info)
        self.assertEqual(info["codigo"], "9999010")

        # Código inexistente
        info_none = get_account_info("0000000", "Pacifico Cable SpA")
        self.assertIsNone(info_none)
        
        # Código None / vacío
        self.assertIsNone(get_account_info(None))
        self.assertIsNone(get_account_info(""))

    def test_find_account_by_rubro_nota(self):
        # Búsqueda de intangibles con nota Cartera de clientes
        cod = find_account_by_rubro_nota(
            linea_item="Activos intangibles distintos de la plusvalía",
            linea_nota="Cartera de clientes",
            empresa="Pacifico Cable SpA"
        )
        self.assertEqual(cod, "9999050")

    def test_db_persistence_with_cuenta_codigo(self):
        from src.models.database import SessionLocal, init_db
        from src.models.consolidacion import ConsolidationJournalEntry
        init_db()
        db = SessionLocal()
        try:
            # Crear un registro de prueba
            entry = ConsolidationJournalEntry(
                grupo_id=1,
                periodo="9999-99",
                glosa="Asiento de Prueba Unitario Cuentas Consolidacion",
                columna_ajuste="PPA",
                linea_item="Activos intangibles distinto a la plusvalía",
                linea_nota="Cartera de clientes",
                cuenta_codigo="9999050",
                debe=50000.0,
                haber=0.0,
                asiento_codigo="TEST-999999-001",
                num_linea=1
            )
            db.add(entry)
            db.commit()

            # Consultar y verificar
            saved = db.query(ConsolidationJournalEntry).filter_by(asiento_codigo="TEST-999999-001").first()
            self.assertIsNotNone(saved)
            self.assertEqual(saved.cuenta_codigo, "9999050")
            self.assertEqual(saved.linea_item, "Activos intangibles distinto a la plusvalía")
            self.assertEqual(saved.linea_nota, "Cartera de clientes")
            self.assertEqual(saved.debe, 50000.0)

            # Limpiar
            db.delete(saved)
            db.commit()
        finally:
            db.close()

if __name__ == '__main__':
    unittest.main()
