import unittest
import pandas as pd
from io import BytesIO
import openpyxl
from docx import Document

from src.reporting.word_export import generate_word_report, WordExportEngine

class TestWordExportFidelity(unittest.TestCase):
    def test_flujo_word_export_matches_excel_template(self):
        # Create a mock openpyxl workbook representing an Excel template
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Flujo"
        
        ws.cell(row=8, column=1, value="Descripción")
        ws.cell(row=8, column=2, value="Nota")
        ws.cell(row=8, column=3, value="2026-07")
        ws.cell(row=8, column=4, value="2025-12")

        # Row 1: Section Header (Bold in Excel)
        ws.cell(row=9, column=1, value="Flujos de efectivo procedentes de (utilizados en) actividades de operación")
        ws.cell(row=9, column=1).font = openpyxl.styles.Font(bold=True)

        # Row 2: Data item (Regular)
        ws.cell(row=10, column=1, value="Cobros procedentes de las ventas de bienes y prestación de servicios")
        ws.cell(row=10, column=3, value=1966337)
        ws.cell(row=10, column=4, value=-98847)

        # Row 3: Subtotal (Bold, with top and bottom thin borders)
        ws.cell(row=11, column=1, value="Flujos de efectivo procedentes de actividades de operación")
        ws.cell(row=11, column=1).font = openpyxl.styles.Font(bold=True)
        ws.cell(row=11, column=3, value=16164335)
        ws.cell(row=11, column=3).font = openpyxl.styles.Font(bold=True)
        ws.cell(row=11, column=3).border = openpyxl.styles.Border(
            top=openpyxl.styles.Side(style='thin'),
            bottom=openpyxl.styles.Side(style='thin')
        )

        # Row 4: Incremento row (Regular in Excel, no borders)
        ws.cell(row=12, column=1, value="Incremento (decremento) neto en efectivo y equivalentes al efectivo")
        ws.cell(row=12, column=3, value=2720738)
        ws.cell(row=12, column=4, value=-1088400)

        excel_io = BytesIO()
        wb.save(excel_io)
        excel_bytes = excel_io.getvalue()

        df = pd.DataFrame({
            "Descripción": [
                "Flujos de efectivo procedentes de (utilizados en) actividades de operación",
                "Cobros procedentes de las ventas de bienes y prestación de servicios",
                "Flujos de efectivo procedentes de actividades de operación",
                "Incremento (decremento) neto en efectivo y equivalentes al efectivo"
            ],
            "Nota": ["", "", "", ""],
            "2026-07": [None, 1966337, 16164335, 2720738],
            "2025-12": [None, -98847, None, -1088400]
        })

        # Generate Word with excel_bytes
        docx_bytes = generate_word_report(
            df,
            title="Estado de Flujos de Efectivo",
            subtitle="Expresado en M$",
            excel_bytes=excel_bytes
        )

        self.assertIsNotNone(docx_bytes)
        doc = Document(BytesIO(docx_bytes))
        self.assertEqual(len(doc.tables), 1)
        table = doc.tables[0]

        # Check Header: Row 0 must have white bg, bold black text
        hdr_c0 = table.rows[0].cells[0]
        self.assertTrue(hdr_c0.paragraphs[0].runs[0].font.bold)

        # Check Section Header: Row 1 must be bold (inherited from Excel)
        r1_c0 = table.rows[1].cells[0]
        self.assertTrue(r1_c0.paragraphs[0].runs[0].font.bold)

        # Check Subtotal: Row 3 must be bold (inherited from Excel)
        r3_c0 = table.rows[3].cells[0]
        self.assertTrue(r3_c0.paragraphs[0].runs[0].font.bold)

        # Check Incremento: Row 4 must NOT be bold (matches screen & Excel!)
        r4_c0 = table.rows[4].cells[0]
        self.assertFalse(r4_c0.paragraphs[0].runs[0].font.bold)

    def test_balance_word_export_compact(self):
        # Create a 40-row dataframe representing a full classified balance
        rows = []
        for i in range(40):
            rows.append({
                "Clasificación": f"Cuenta de Balance {i+1}",
                "Nota": str(i % 10 + 1) if i % 4 == 0 else "",
                "2026-07": 1000000 + i * 50000,
                "2025-12": 950000 + i * 45000
            })
        df_bal = pd.DataFrame(rows)

        docx_bal = WordExportEngine.generate_classified_balance_word(
            df=df_bal,
            title="Estado de Situación Financiera Clasificado",
            unit="M$",
            entity_name="PACIFICO CABLE SPA"
        )
        self.assertIsNotNone(docx_bal)
        doc = Document(docx_bal)
        self.assertEqual(len(doc.tables), 1)
        # Verify 41 rows (1 header + 40 data)
        self.assertEqual(len(doc.tables[0].rows), 41)

    def test_big4_audited_table_styling(self):
        """
        Verifica que las tablas generadas sigan el estándar exacto de Formato base de programa_v1.docx:
        1. Cabecera con fondo #E8EEF6, sin borde superior, borde inferior #1F3864 (sz="8"), y texto #1F3864.
        2. Filas de datos con fondo blanco limpio #FFFFFF y línea tenue #E1E1E1 (sz="2") de separación.
        3. Totales mayores con línea superior #404040 y doble inferior #404040 solo en columnas numéricas/notas, mientras que col 0 no tiene línea superior.
        """
        df_test = pd.DataFrame({
            "Concepto": ["Efectivo y equivalentes", "Subtotal Activo Corriente", "Total Activos"],
            "Nota": ["1", "", ""],
            "2026": [150000, 150000, 150000],
            "2025": [120000, 120000, 120000]
        })
        docx_bytes = generate_word_report(df_test, title="Prueba Estilo Big 4", subtitle="M$")
        doc = Document(BytesIO(docx_bytes))
        table = doc.tables[0]

        # 1. Cabecera: Celda 0
        hdr_c0 = table.rows[0].cells[0]
        # Shading #E8EEF6
        shd = hdr_c0._tc.xpath('.//w:tcPr/w:shd/@w:fill')
        self.assertEqual(shd, ["E8EEF6"])
        # Bordes de cabecera: top none, bottom 1F3864
        top_val = hdr_c0._tc.xpath('.//w:tcPr/w:tcBorders/w:top/@w:val')
        self.assertEqual(top_val, ['none'])
        bottom_border_color = hdr_c0._tc.xpath('.//w:tcPr/w:tcBorders/w:bottom/@w:color')
        self.assertEqual(bottom_border_color, ["1F3864"])
        # Color de texto #1F3864
        hdr_run = hdr_c0.paragraphs[0].runs[0]
        self.assertEqual(str(hdr_run.font.color.rgb), "1F3864")

        # 2. Filas de datos sin cebrado (todas con fondo blanco #FFFFFF)
        for r_idx in range(1, len(table.rows)):
            for cell in table.rows[r_idx].cells:
                r_shd = cell._tc.xpath('.//w:tcPr/w:shd/@w:fill')
                self.assertEqual(r_shd, ["FFFFFF"])

        # 3. Fila de datos normal (fila 1): debe tener separador tenue E1E1E1
        r1_c0_bottom_color = table.rows[1].cells[0]._tc.xpath('.//w:tcPr/w:tcBorders/w:bottom/@w:color')
        self.assertEqual(r1_c0_bottom_color, ['E1E1E1'])

        # 4. Totales: Column 0 no debe tener línea superior
        total_row = table.rows[3] # Total Activos
        col0_top = total_row.cells[0]._tc.xpath('.//w:tcPr/w:tcBorders/w:top/@w:val')
        self.assertEqual(col0_top, ['none'])
        # Columna 2 (numérica) tiene línea superior 404040 y subrayado doble 404040
        col2_top_val = total_row.cells[2]._tc.xpath('.//w:tcPr/w:tcBorders/w:top/@w:val')
        col2_top_color = total_row.cells[2]._tc.xpath('.//w:tcPr/w:tcBorders/w:top/@w:color')
        col2_bot_val = total_row.cells[2]._tc.xpath('.//w:tcPr/w:tcBorders/w:bottom/@w:val')
        col2_bot_color = total_row.cells[2]._tc.xpath('.//w:tcPr/w:tcBorders/w:bottom/@w:color')
        self.assertEqual(col2_top_val, ['single'])
        self.assertEqual(col2_top_color, ['404040'])
        self.assertEqual(col2_bot_val, ['double'])
        self.assertEqual(col2_bot_color, ['404040'])

if __name__ == '__main__':
    unittest.main()

