import unittest
import pandas as pd
from io import BytesIO
from src.ui_pages.informes_y_notas import split_sheet_into_elements
from src.reporting.word_export import WordExportEngine

class TestExplicitNoteCodes(unittest.TestCase):

    def test_explicit_codes_and_comparative_grouping(self):
        # Simulated sheet with explicit codes and [COMPARATIVO]
        test_rows = [
            ["#N15.1 Detalle de Pasivos Financieros", None, None],
            ["Concepto", "2026", "2025"],
            ["Prestamos", 100, 90],
            ["Total", 100, 90],
            [None, None, None],
            ["#N15.2 [COMPARATIVO] Vencimiento de Pasivos (NIC 32)", None, None],
            ["Detalle al 30.06.2026", "90 dias", "Total"],
            ["Prestamos", 50, 50],
            ["Total", 50, 50],
            [None, None, None],
            ["#N15.2 Vencimiento de Pasivos (NIC 32)", None, None], # Duplicate comparative header
            ["Detalle al 31.12.2025", "90 dias", "Total"],
            ["Prestamos", 40, 40],
            ["Total", 40, 40],
            [None, None, None],
            ["#N15.3 [COMPARATIVO] Cambios en Pasivo Financiero", None, None],
            ["Saldo 2026", "Aumento", "Total"],
            ["Prestamos", 10, 60],
            ["Total", 10, 60],
            [None, None, None],
            ["#N15.3 Cambios en Pasivo Financiero", None, None], # Duplicate comparative header
            ["Saldo 2025", "Aumento", "Total"],
            ["Prestamos", 8, 48],
            ["Total", 8, 48]
        ]
        df = pd.DataFrame(test_rows)
        elements = split_sheet_into_elements(df, sheet_name="Pasivos financieros ")

        # Should produce 3 tables with exact codes #N15.1, #N15.2, #N15.3
        tables = [e for e in elements if e[0] == "table"]
        self.assertEqual(len(tables), 3)

        self.assertEqual(tables[0][3], "#N15.1")
        self.assertEqual(tables[1][3], "#N15.2")
        self.assertEqual(tables[2][3], "#N15.3")

        # Table 2 must contain both 2026 and 2025 rows (total 6 rows of data)
        self.assertEqual(len(tables[1][1]), 6)
        # Table 3 must contain both 2026 and 2025 rows (total 6 rows of data)
        self.assertEqual(len(tables[2][1]), 6)

    def test_legacy_sheet_fallback(self):
        # Sheet without explicit # codes
        test_rows = [
            ["Detalle de Efectivo", None, None],
            ["Concepto", "2026", "2025"],
            ["Banco", 500, 400],
            ["Total", 500, 400]
        ]
        df = pd.DataFrame(test_rows)
        elements = split_sheet_into_elements(df, sheet_name="Efectivo")

        tables = [e for e in elements if e[0] == "table"]
        self.assertEqual(len(tables), 1)
        # Legacy mode returns 2-tuple or None for explicit code
        exp_code = tables[0][3] if len(tables[0]) > 3 else None
        self.assertIsNone(exp_code)

if __name__ == "__main__":
    unittest.main()
