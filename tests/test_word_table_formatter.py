import unittest
import pandas as pd
import numpy as np
from docx import Document
from src.reporting.word_table_formatter import (
    clean_table_df,
    is_valid_note_table,
    classify_columns,
    compute_optimal_column_widths,
    apply_table_formatting_and_widths,
    get_dynamic_font_and_spacing,
    should_use_landscape,
)

class TestWordTableFormatter(unittest.TestCase):
    def test_clean_table_df_removes_empty_rows_and_cols(self):
        df = pd.DataFrame({
            0: ["Concepto", "Terrenos", np.nan, "Total"],
            1: [np.nan, np.nan, np.nan, np.nan], # Empty col
            2: ["31.12.2025", 1000, np.nan, 1000],
            3: [np.nan, np.nan, np.nan, np.nan], # Empty col
        })
        cleaned = clean_table_df(df)
        self.assertEqual(cleaned.shape[1], 2)
        self.assertEqual(len(cleaned), 3)

    def test_is_valid_note_table(self):
        # Header-only single row -> Invalid
        header_only = pd.DataFrame([["31.12.2025", "31.12.2024", np.nan]])
        self.assertFalse(is_valid_note_table(header_only))
        
        # Real table with values -> Valid
        real_table = pd.DataFrame({
            0: ["Efectivo", "Bancos", "Total"],
            1: [5000, 15000, 20000]
        })
        self.assertTrue(is_valid_note_table(real_table))

    def test_classify_columns(self):
        # Table with Concepto, Nota, Val1, Val2
        df = pd.DataFrame({
            'Concepto': ["Caja", "Clientes"],
            'Nota': ["4", "5.1"],
            '2025': [150000, -25000],
            '2024': [120000, 10000]
        })
        classes = classify_columns(df)
        self.assertEqual(classes, ['text', 'note', 'numeric', 'numeric'])

        # Table with multiple text columns (Currency note)
        df_curr = pd.DataFrame({
            0: ["Efectivo", "Cuentas por pagar"],
            1: ["Dólares estadounidenses", "Dólares estadounidenses"],
            2: [-14171193, 3128229125],
            3: [0, 0]
        })
        classes_curr = classify_columns(df_curr)
        self.assertEqual(classes_curr, ['text', 'text', 'numeric', 'numeric'])

    def test_compute_optimal_column_widths(self):
        df_curr = pd.DataFrame({
            0: ["Efectivo y efectivo equivalente", "Cuentas por cobrar con entidades relacionadas"],
            1: ["Dólares estadounidenses", "Dólares estadounidenses"],
            2: [-14171193, 3128229125],
            3: [0, 0]
        })
        classes = ['text', 'text', 'numeric', 'numeric']
        widths = compute_optimal_column_widths(df_curr, classes, total_width_inches=6.85)
        self.assertEqual(len(widths), 4)
        # Sum must strictly equal 6.85
        self.assertAlmostEqual(sum(widths), 6.85, places=2)
        # Text column 1 must have at least 1.8 inches so it doesn't wrap into 3 vertical lines!
        self.assertGreaterEqual(widths[1], 1.8)
        # Numeric columns must have appropriate widths (col 2 with 10 digits >= 1.15, col 3 with 1 digit >= 0.90)
        self.assertGreaterEqual(widths[2], 1.15)
        self.assertGreaterEqual(widths[3], 0.90)

    def test_apply_table_formatting_and_widths(self):
        doc = Document()
        tbl = doc.add_table(rows=3, cols=4)
        df = pd.DataFrame({
            'Item': ['A', 'B'],
            'Nota': ['1', '2'],
            'Val1': [100, 200],
            'Val2': [300, 400]
        })
        apply_table_formatting_and_widths(tbl, df, total_width_inches=6.85)
        
        # Verify cantSplit on rows
        for row in tbl.rows:
            self.assertTrue(any(c.tag.endswith('cantSplit') for c in row._tr.trPr))
        # Verify tblHeader on row 0
        self.assertTrue(any(c.tag.endswith('tblHeader') for c in tbl.rows[0]._tr.trPr))
        # Verify noWrap on numeric and note cells
        for row in tbl.rows:
            self.assertTrue(any(c.tag.endswith('noWrap') for c in row.cells[1]._tc.tcPr))
            self.assertTrue(any(c.tag.endswith('noWrap') for c in row.cells[2]._tc.tcPr))
            self.assertTrue(any(c.tag.endswith('noWrap') for c in row.cells[3]._tc.tcPr))
            # Text column 0 should NOT have noWrap
            self.assertFalse(any(c.tag.endswith('noWrap') for c in row.cells[0]._tc.tcPr))

    def test_should_use_landscape(self):
        self.assertTrue(should_use_landscape(8))
        self.assertTrue(should_use_landscape(4, tag_code="{{#PATRIMONIO:HORIZONTAL}}"))
        self.assertFalse(should_use_landscape(4, tag_code="{{#BALANCE}}"))

if __name__ == '__main__':
    unittest.main()
