import os
import io
import time
import unittest
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

import pandas as pd
import numpy as np

from src.core.word_template_engine import (
    WordTemplateEngine,
    set_table_metadata,
    get_table_metadata,
    update_table_cells_inplace
)
from src.models.trial_balance_db import TrialBalanceDB
from src.reporting.balance_generator import BalanceGenerator
from src.core.word_template_engine import load_mappings_for_entity

class TestWordInplaceUpdate(unittest.TestCase):
    
    def test_inplace_preserves_custom_widths_and_styles(self):
        """
        Test 1: Verifica que update_table_cells_inplace actualice los valores numéricos
        respetando al 100% los anchos de columna (tcW), márgenes, bordes y fuentes de Word.
        """
        doc = docx.Document()
        table = doc.add_table(rows=3, cols=3)
        
        # Asignar anchos personalizados deliberados (ej: col0=3.2 in, col1=1.8 in, col2=1.8 in)
        custom_widths = [3.2, 1.8, 1.8]
        for row in table.rows:
            for c_idx, cell in enumerate(row.cells):
                cell.width = Inches(custom_widths[c_idx])
                tcPr = cell._tc.get_or_add_tcPr()
                tcPr.append(parse_xml(f'<w:tcW {nsdecls("w")} w:w="{int(round(custom_widths[c_idx] * 1440))}" w:type="dxa"/>'))
                
        # Estilos iniciales en celdas
        hdr_row = table.rows[0]
        hdr_row.cells[0].paragraphs[0].add_run("Concepto Contable")
        hdr_row.cells[1].paragraphs[0].add_run("31-12-2026 M$")
        hdr_row.cells[2].paragraphs[0].add_run("31-12-2025 M$")
        
        data_row = table.rows[1]
        r0 = data_row.cells[0].paragraphs[0].add_run("Efectivo y efectivo equivalente")
        r0.font.name = "Georgia"
        r0.font.size = Pt(11)
        r0.font.bold = True
        
        r1 = data_row.cells[1].paragraphs[0].add_run("0")
        r1.font.name = "Georgia"
        r1.font.size = Pt(11)
        
        r2 = data_row.cells[2].paragraphs[0].add_run("0")
        r2.font.name = "Georgia"
        r2.font.size = Pt(11)
        
        # Capturar anchos en dxa antes de la inyección
        widths_before = [cell._tc.xpath('.//w:tcW/@w:w') for row in table.rows for cell in row.cells]
        
        # DataFrame con nuevas cifras
        df_update = pd.DataFrame({
            "Concepto": ["Efectivo y efectivo equivalente", "Total Activos"],
            "2026": [15181513, 316263153],
            "2025": [8170902, 305969866]
        })
        
        # Ejecutar actualización in-place con período comparativo explícito
        update_table_cells_inplace(table, df_update, periodo_actual="2026", periodo_comp="2025")
        
        # 1. Validar que los anchos de celda en dxa son IDÉNTICOS
        widths_after = [cell._tc.xpath('.//w:tcW/@w:w') for row in table.rows for cell in row.cells]
        self.assertEqual(widths_before, widths_after, "Los anchos de columna fueron modificados por error.")
        
        # 2. Validar que el valor numérico se actualizó con formato contable chileno
        txt_cell_2026 = table.rows[1].cells[1].text.strip()
        txt_cell_2025 = table.rows[1].cells[2].text.strip()
        self.assertEqual(txt_cell_2026, "15.181.513", f"Valor 2026 incorrecto: {txt_cell_2026}")
        self.assertEqual(txt_cell_2025, "8.170.902", f"Valor 2025 incorrecto: {txt_cell_2025}")
        
        # 3. Validar preservación de tipografía
        run_updated = table.rows[1].cells[1].paragraphs[0].runs[0]
        self.assertEqual(run_updated.font.name, "Georgia", "La tipografía personalizada no se preservó.")
        self.assertEqual(run_updated.font.size, Pt(11), "El tamaño de fuente no se preservó.")

    def test_numeric_accuracy_and_tie_out(self):
        """
        Test 2: Integridad numérica y exactitud al centavo con datos reales de Pacifico Cable SpA.
        """
        empresa = "Pacifico Cable SpA"
        p_act = "2026-03"
        p_cmp = "2025-12"
        
        mb, mp = load_mappings_for_entity(empresa)
        tb_act = TrialBalanceDB.get_trial_balance(empresa, p_act)
        tb_cmp = TrialBalanceDB.get_trial_balance(empresa, p_cmp)
        
        gen = BalanceGenerator(os.path.join("data", "empresas", empresa, "Balance clasificado.xlsx"))
        res_bytes = gen.generate(tb_act, mb, scale_factor=1000.0, tb_df_comp=tb_cmp, periodo_actual_str=p_act, periodo_comp_str=p_cmp)
        
        wte = WordTemplateEngine(None)
        bal_df = wte._extract_balance_preview(res_bytes, p_act, p_cmp)
        
        # Crear tabla de Word simulada con los mismos conceptos
        doc = docx.Document()
        tbl = doc.add_table(rows=len(bal_df), cols=len(bal_df.columns))
        for i, row in bal_df.iterrows():
            tbl.rows[i].cells[0].text = str(row.iloc[0])
            tbl.rows[i].cells[1].text = "0"
            tbl.rows[i].cells[2].text = "0"
            
        update_table_cells_inplace(tbl, bal_df)
        
        # Validar filas clave
        for i, row in bal_df.iterrows():
            c_name = str(row.iloc[0]).strip()
            val_act = row.iloc[1]
            if pd.notna(val_act) and isinstance(val_act, (int, float)) and val_act != 0:
                expected_str = f"{val_act:,.0f}".replace(",", ".") if val_act > 0 else f"({abs(val_act):,.0f})".replace(",", ".")
                actual_word_str = tbl.rows[i].cells[1].text.strip()
                self.assertEqual(actual_word_str, expected_str, f"Diferencia numérica en {c_name}: esperado {expected_str}, obtenido {actual_word_str}")

    def test_coexistence_mode_a_and_mode_b(self):
        """
        Test 3: Convivencia y fallback transparente.
        Verifica que una tabla prediseñada se actualice in-place y que una etiqueta suelta
        genere una tabla nueva automáticamente.
        """
        doc = docx.Document()
        
        # 1. Tabla existente marcada con metadata #BALANCE
        tbl_existente = doc.add_table(rows=2, cols=3)
        tbl_existente.rows[0].cells[0].text = "Concepto"
        tbl_existente.rows[0].cells[1].text = "2026"
        tbl_existente.rows[0].cells[2].text = "2025"
        tbl_existente.rows[1].cells[0].text = "Efectivo y efectivo equivalente"
        tbl_existente.rows[1].cells[1].text = "0"
        tbl_existente.rows[1].cells[2].text = "0"
        set_table_metadata(tbl_existente, "#BALANCE")
        
        # 2. Párrafo suelto con etiqueta Modo A
        p_nuevo = doc.add_paragraph("{{#ER}}")
        
        doc_io = io.BytesIO()
        doc.save(doc_io)
        doc_io.seek(0)
        
        engine = WordTemplateEngine(doc_io)
        out_buf = engine.process(
            empresa_activa="Pacifico Cable SpA",
            periodo_actual="2026-03",
            periodo_comp="2025-12",
            scale_factor=1000.0,
            use_ifrs_auto=False,
            include_headers=False
        )
        
        doc_out = docx.Document(out_buf)
        # Debe haber al menos 2 tablas: la primera actualizada in-place y la segunda creada para #ER
        self.assertGreaterEqual(len(doc_out.tables), 2)
        
        # Verificar que la primera tabla contiene el saldo in-place
        first_tbl = doc_out.tables[0]
        self.assertEqual(first_tbl.rows[1].cells[1].text.strip(), "15.181.513")

    def test_end_to_end_user_template_performance(self):
        """
        Test 4: Ejecución completa en el archivo real del usuario verificando tiempo < 10 segundos.
        """
        doc_path = r'C:\Mundo\Programa de informes\Informes\TEST Pacifico auditado\Cierre anual\TEST 2_Pacifico Cable SpA 31-12-2026_DRAFT_v1.docx'
        if not os.path.exists(doc_path):
            self.skipTest(f"Archivo de prueba {doc_path} no disponible en el entorno.")
            
        with open(doc_path, 'rb') as f:
            template_bytes = io.BytesIO(f.read())
            
        engine = WordTemplateEngine(template_bytes)
        t0 = time.time()
        res_buffer = engine.process(
            empresa_activa="Pacifico Cable SpA",
            periodo_actual="2026-03",
            periodo_comp="2025-12",
            scale_factor=1000.0,
            use_ifrs_auto=False,
            include_headers=False
        )
        elapsed = time.time() - t0
        
        self.assertLess(elapsed, 10.0, f"El proceso tomó {elapsed:.2f}s, excediendo el objetivo de < 10s.")
        self.assertGreater(len(res_buffer.getvalue()), 100000, "El documento generado está incompleto.")

    def test_comparative_preservation_when_periodo_comp_is_ninguno(self):
        """
        Test 5: Verifica que cuando periodo_comp='Ninguno', las columnas y valores de años
        anteriores (ej. 2025) en Word queden 100% intactos y protegidos, modificando ÚNICAMENTE el año actual.
        """
        doc = docx.Document()
        table = doc.add_table(rows=3, cols=3)
        
        # Cabeceras
        table.rows[0].cells[0].text = "Concepto"
        table.rows[0].cells[1].text = "30-06-2026"
        table.rows[0].cells[2].text = "31-12-2025"
        
        # Fila 1: Efectivo
        table.rows[1].cells[0].text = "Efectivo y efectivo equivalente"
        table.rows[1].cells[1].text = "0"
        table.rows[1].cells[2].text = "9.999.999" # Valor previo auditado
        
        # Fila 2: Total Activos
        table.rows[2].cells[0].text = "Total Activos"
        table.rows[2].cells[1].text = "0"
        table.rows[2].cells[2].text = "88.888.888" # Valor previo auditado
        
        # DataFrame que trae datos tanto para 2026 como para 2025
        df_update = pd.DataFrame({
            "Concepto": ["Efectivo y efectivo equivalente", "Total Activos"],
            "2026": [15181513, 316263153],
            "2025": [11111111, 22222222] # Cifras que NO deben inyectarse
        })
        
        # Ejecutar con periodo_comp="Ninguno"
        update_table_cells_inplace(table, df_update, periodo_actual="2026-06", periodo_comp="Ninguno")
        
        # 1. 2026 DEBE estar actualizado
        self.assertEqual(table.rows[1].cells[1].text.strip(), "15.181.513")
        self.assertEqual(table.rows[2].cells[1].text.strip(), "316.263.153")
        
        # 2. 2025 DEBE permanecer 100% INTACTO con su valor previo auditado
        self.assertEqual(table.rows[1].cells[2].text.strip(), "9.999.999", "El valor 2025 preexistente fue alterado indebidamente.")
        self.assertEqual(table.rows[2].cells[2].text.strip(), "88.888.888", "El total 2025 preexistente fue alterado indebidamente.")

    def test_note_14_1_structural_rebuild_with_ninguno(self):
        """
        Test 6: Verifica que si una tabla existente en Word (ej. #N14.1 con 5 columnas antiguas)
        cambia de formato en la Plantilla Maestra Global (ahora 7 columnas con Corriente y No corriente),
        al emitir con Período Actual='2026-09' y Período Comparativo='Ninguno' se reconstruya a 7 columnas
        con las cifras de 2026 y 2025 completas.
        """
        doc = docx.Document()
        old_tbl = doc.add_table(rows=5, cols=5)
        old_tbl.rows[0].cells[0].text = "Nombre entidad"
        old_tbl.rows[0].cells[1].text = "relación"
        old_tbl.rows[0].cells[2].text = "Moneda"
        old_tbl.rows[0].cells[3].text = "M$"
        old_tbl.rows[0].cells[4].text = "30.06.2026"
        set_table_metadata(old_tbl, "#N14.1")

        doc_io = io.BytesIO()
        doc.save(doc_io)
        doc_io.seek(0)

        engine = WordTemplateEngine(doc_io)
        out_buf = engine.process(
            empresa_activa="Pacifico Cable SpA",
            periodo_actual="2026-09",
            periodo_comp="Ninguno",
            scale_factor=1000.0,
            use_ifrs_auto=True,
            include_headers=False
        )

        doc_out = docx.Document(out_buf)
        self.assertEqual(len(doc_out.tables), 1)
        rebuilt_tbl = doc_out.tables[0]
        self.assertEqual(len(rebuilt_tbl.columns), 7, "La tabla #N14.1 debe haberse reconstruido con las 7 columnas nuevas.")
        self.assertEqual(rebuilt_tbl.rows[3].cells[3].text.strip(), "28.797.507")
        self.assertEqual(rebuilt_tbl.rows[3].cells[4].text.strip(), "22.150.376")
        self.assertEqual(rebuilt_tbl.rows[3].cells[5].text.strip(), "131.548.096")
        self.assertEqual(rebuilt_tbl.rows[3].cells[6].text.strip(), "121.965.389")

if __name__ == '__main__':
    unittest.main()

