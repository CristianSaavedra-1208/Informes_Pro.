import os
import sys
import re
import pandas as pd
from io import BytesIO
import docx
from docx import Document
from docxtpl import DocxTemplate
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor

# Asegurar importabilidad de src
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.append(ROOT_DIR)

from src.models.database import SessionLocal
from src.models.consolidacion import ConsolidationGroup
from src.models.trial_balance_db import TrialBalanceDB
from src.models.pl_cubo_db import PlCuboDB

from docx.shared import Inches, Pt, RGBColor
from src.reporting.word_table_formatter import (
    clean_table_df,
    is_valid_note_table,
    classify_columns,
    compute_optimal_column_widths,
    apply_table_formatting_and_widths,
    set_cell_nowrap,
    set_table_width,
    get_dynamic_font_and_spacing,
    should_use_landscape,
    create_section_break_xml,
)

def set_cell_bg(cell, fill_color):
    """Inyecta XML shading para cambiar el color de fondo de una celda de Word."""
    tcPr = cell._tc.get_or_add_tcPr()
    for c in list(tcPr):
        if c.tag.endswith('shd'):
            tcPr.remove(c)
    shading_elm = parse_xml(r'<w:shd {} w:val="clear" w:color="auto" w:fill="{}"/>'.format(nsdecls('w'), fill_color))
    tcPr.append(shading_elm)

def set_cell_margins(cell, top=50, bottom=50, left=100, right=100):
    """Ajusta los márgenes internos (padding) de la celda de Word en dxa (1 pt = 20 dxa)."""
    tcPr = cell._tc.get_or_add_tcPr()
    for c in list(tcPr):
        if c.tag.endswith('tcMar'):
            tcPr.remove(c)
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def set_cell_borders(cell, top=None, bottom=None):
    """Inyecta bordes OXML a una celda de Word (útil para filas de totales/subtotales contables)."""
    tcPr = cell._tc.get_or_add_tcPr()
    for c in list(tcPr):
        if c.tag.endswith('tcBorders'):
            tcPr.remove(c)
    borders_xml = f'<w:tcBorders {nsdecls("w")}>'
    if top:
        borders_xml += f'<w:top w:val="{top.get("val", "single")}" w:sz="{top.get("sz", "6")}" w:space="0" w:color="{top.get("color", "000000")}"/>'
    else:
        borders_xml += f'<w:top w:val="none"/>'
    if bottom:
        borders_xml += f'<w:bottom w:val="{bottom.get("val", "single")}" w:sz="{bottom.get("sz", "6")}" w:space="0" w:color="{bottom.get("color", "000000")}"/>'
    else:
        borders_xml += f'<w:bottom w:val="none"/>'
    borders_xml += f'<w:left w:val="none"/><w:right w:val="none"/>'
    borders_xml += '</w:tcBorders>'
    tcPr.append(parse_xml(borders_xml))

def set_table_metadata(tbl, tag_code):
    """Asigna metadatos invisibles w:tblCaption y w:tblDescription a la tabla para actualización in-place futura."""
    try:
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        tblPr = tbl._tbl.tblPr
        for child in list(tblPr):
            if child.tag.endswith(('tblCaption', 'tblDescription')):
                tblPr.remove(child)
        caption = OxmlElement('w:tblCaption')
        caption.set(qn('w:val'), str(tag_code))
        tblPr.append(caption)
        desc = OxmlElement('w:tblDescription')
        desc.set(qn('w:val'), str(tag_code))
        tblPr.append(desc)
    except Exception:
        pass

def get_table_metadata(tbl):
    """Lee metadatos invisibles w:tblCaption o w:tblDescription de la tabla de Word."""
    try:
        caps = tbl._tbl.xpath('.//w:tblPr/w:tblCaption/@w:val')
        if caps and str(caps[0]).strip():
            return str(caps[0]).strip()
        descs = tbl._tbl.xpath('.//w:tblPr/w:tblDescription/@w:val')
        if descs and str(descs[0]).strip():
            return str(descs[0]).strip()
    except Exception:
        pass
    return None

def apply_table_column_widths(table, num_cols, total_width_inches=6.85, df=None):
    """Calcula y aplica anchos de columna óptimos para informes financieros (100% de margen a margen)."""
    table.autofit = False
    set_table_width(table, total_width_inches)
    
    if df is not None:
        apply_table_formatting_and_widths(table, df, total_width_inches=total_width_inches)
        return

    if num_cols <= 1:
        col_widths = [total_width_inches]
    elif num_cols == 2:
        col_widths = [total_width_inches * 0.70, total_width_inches * 0.30]
    elif num_cols == 3:
        col_widths = [total_width_inches * 0.64, total_width_inches * 0.18, total_width_inches * 0.18]
    elif num_cols == 4:
        col_widths = [total_width_inches * 0.56, total_width_inches * 0.10, total_width_inches * 0.17, total_width_inches * 0.17]
    else:
        col0 = max(2.8, total_width_inches - (0.95 * (num_cols - 1)))
        rem = max(0.6, (total_width_inches - col0) / (num_cols - 1))
        col_widths = [col0] + [rem] * (num_cols - 1)
        factor = total_width_inches / sum(col_widths)
        col_widths = [w * factor for w in col_widths]

    for row_idx, row in enumerate(table.rows):
        trPr = row._tr.get_or_add_trPr()
        if not any(c.tag.endswith('cantSplit') for c in trPr):
            trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        if row_idx == 0 and not any(c.tag.endswith('tblHeader') for c in trPr):
            trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
        for j, cell in enumerate(row.cells):
            if j < len(col_widths):
                cell.width = Inches(col_widths[j])
                tcPr = cell._tc.get_or_add_tcPr()
                for c in list(tcPr):
                    if c.tag.endswith('tcW'):
                        tcPr.remove(c)
                tcPr.append(parse_xml(f'<w:tcW {nsdecls("w")} w:w="{int(round(col_widths[j] * 1440))}" w:type="dxa"/>'))

def move_element_before(target_p, new_element):
    """Mueve un elemento XML de Word (párrafo o tabla) justo antes del párrafo objetivo."""
    parent = target_p._p.getparent()
    idx = parent.index(target_p._p)
    parent.insert(idx, new_element)

def populate_docx_table(table, df, numeric_cols, unit="M$", is_balance=False):
    """Rellena y da estilo corporativo a una tabla de Word idéntico a la pantalla."""
    table.style = 'Normal Table'
    df = clean_table_df(df)
    num_cols = len(df.columns)
    col_classes = classify_columns(df)
    is_landscape = should_use_landscape(num_cols)
    total_width_inches = 9.60 if is_landscape else (7.50 if is_balance else 6.85)
    
    # Configuración de espaciados dinámicos según columnas y orientación
    base_font_size, hdr_font_size, p_spacing, pad_top, pad_lr = get_dynamic_font_and_spacing(num_cols, is_compact=is_balance, is_landscape=is_landscape)
    pad_bottom = pad_top
    hdr_pad_v = pad_top
    hdr_line_spacing = Pt(p_spacing.pt - 0.5) if hasattr(p_spacing, 'pt') else Pt(8.5)

    # Fila de cabecera
    is_miles = any(m in str(unit).lower() for m in ["miles", "m$", "mch$"])
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(df.columns):
        header_text = str(col_name)
        if is_miles and (any(char.isdigit() for char in header_text) or (i > 0 and header_text.lower() not in ["nota", "notas"])):
            if "M$" not in header_text:
                header_text = f"{header_text}\nM$"
        hdr_cells[i].text = header_text
        set_cell_bg(hdr_cells[i], "E8EEF6")
        set_cell_margins(hdr_cells[i], top=hdr_pad_v, bottom=hdr_pad_v, left=pad_lr, right=pad_lr)
        set_cell_borders(hdr_cells[i], top=None, bottom={'val': 'single', 'sz': '8', 'color': '1F3864'})
        for p in hdr_cells[i].paragraphs:
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = hdr_line_spacing
            if i < len(col_classes) and col_classes[i] == 'numeric':
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                set_cell_nowrap(hdr_cells[i])
            elif i < len(col_classes) and col_classes[i] == 'note':
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                set_cell_nowrap(hdr_cells[i])
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            for run in p.runs:
                run.font.bold = True
                run.font.color.rgb = RGBColor(31, 56, 100) # #1F3864
                run.font.name = 'Arial'
                run.font.size = Pt(hdr_font_size)
                
    # Filas de datos
    for row_i, (_, row_vals) in enumerate(df.iterrows()):
        row_cells = table.rows[row_i + 1].cells
        first_cell_str = str(row_vals.iloc[0]).strip()
        first_cell_clean = first_cell_str.lower()
        
        # Verificar si la fila contiene datos numéricos
        has_data_values = any(
            isinstance(row_vals.iloc[c], (int, float)) or 
            (pd.notna(row_vals.iloc[c]) and str(row_vals.iloc[c]).strip() not in ["", "-", "nan", "None"] and any(ch.isdigit() for ch in str(row_vals.iloc[c])))
            for c in range(1, num_cols)
        )
        
        is_category_header = (not has_data_values) or (
            any(k in first_cell_clean for k in ["(utilizados en)", "actividades de operación", "actividades de inversión", "actividades de financiación"]) and 
            not any(k in first_cell_clean for k in ["total", "subtotal"]) and not has_data_values
        ) or first_cell_str in ["Activos", "Activos corrientes", "Activos no corrientes", "Patrimonio y pasivos", "Pasivos corrientes", "Pasivos no corrientes", "Patrimonio"]
        
        from src.reporting.word_export import _is_total_or_subtotal_label
        is_grand_total = not is_category_header and _is_total_or_subtotal_label(first_cell_clean)
        is_subtotal = False

        # Estilo auditado Big 4: fondo blanco sin cebrado
        bg_color = "FFFFFF"

        for col_idx, col_name in enumerate(df.columns):
            val = row_vals[col_name]
            
            if pd.isna(val) or val == "None" or str(val) == "nan":
                text_val = ""
            else:
                if isinstance(val, (int, float)):
                    try:
                        if val == 0:
                            text_val = "0" if is_category_header else "-"
                        elif val < 0:
                            text_val = f"({abs(val):,.0f})".replace(",", ".")
                        else:
                            text_val = f"{val:,.0f}".replace(",", ".")
                    except:
                        text_val = str(val)
                else:
                    text_val = str(val)
                    
            cell = row_cells[col_idx]
            cell.text = text_val
            set_cell_bg(cell, bg_color)
            set_cell_margins(cell, top=pad_top, bottom=pad_bottom, left=pad_lr, right=pad_lr)
            
            if is_grand_total or is_subtotal:
                top_b = {'val': 'single', 'sz': '6', 'color': '404040'}
                bottom_b = {'val': 'single', 'sz': '6', 'color': '404040'}
            else:
                top_b = None
                bottom_b = {'val': 'single', 'sz': '2', 'color': 'E1E1E1'}
            
            set_cell_borders(cell, top=top_b, bottom=bottom_b)

            for p in cell.paragraphs:
                p.paragraph_format.space_before = Pt(0) if is_balance else Pt(0.5)
                p.paragraph_format.space_after = Pt(0) if is_balance else Pt(0.5)
                p.paragraph_format.line_spacing = p_spacing
                
                if col_idx < len(col_classes) and col_classes[col_idx] == 'numeric':
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    set_cell_nowrap(cell)
                elif col_idx < len(col_classes) and col_classes[col_idx] == 'note':
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    set_cell_nowrap(cell)
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    
                for run in p.runs:
                    run.font.name = 'Arial'
                    run.font.size = Pt(base_font_size)
                    if is_grand_total or is_subtotal or is_category_header:
                        run.font.bold = True
                    run.font.color.rgb = RGBColor(0, 0, 0)

    apply_table_formatting_and_widths(table, df, classifications=col_classes, total_width_inches=total_width_inches, is_compact=is_balance, is_landscape=is_landscape)

def populate_notes_docx_table(table, chunk_df):
    """Rellena y da estilo a una tabla de notas en base a un DataFrame de openpyxl/slicing."""
    table.style = 'Normal Table'
    chunk_df = clean_table_df(chunk_df)
    num_cols = len(chunk_df.columns)
    col_classes = classify_columns(chunk_df)
    is_landscape = should_use_landscape(num_cols)
    total_width_inches = 9.60 if is_landscape else 6.85
    
    base_font_size, hdr_font_size, p_spacing, pad_top, pad_lr = get_dynamic_font_and_spacing(num_cols, is_compact=False, is_landscape=is_landscape)
    pad_bottom = pad_top
    
    from src.reporting.word_export import _is_total_or_subtotal_label
    from src.reporting.formatting import detect_internal_header_rows, is_separator_row_series, ensure_subtable_separator_rows
    chunk_df = ensure_subtable_separator_rows(chunk_df)
    internal_hdr_rows, bottom_hdr_rows = detect_internal_header_rows(chunk_df)

    for row_i, (_, row) in enumerate(chunk_df.iterrows()):
        row_cells = table.rows[row_i].cells if row_i < len(table.rows) else table.add_row().cells
        first_cell_raw = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) and str(row.iloc[0]).strip() not in ["None", "nan"] else ""
        first_cell_val = first_cell_raw.lower()
        is_internal_hdr = (row_i in internal_hdr_rows)
        is_bottom_internal_hdr = (row_i in bottom_hdr_rows)
        is_sep_row = (not is_internal_hdr) and is_separator_row_series(row)
        is_grand_total = (not is_internal_hdr) and (not is_sep_row) and _is_total_or_subtotal_label(first_cell_val)
        is_subtotal = False
        bg_color = "E8EEF6" if is_internal_hdr else "FFFFFF"
            
        for col_idx, col_name in enumerate(chunk_df.columns):
            val = row[col_name]
            if is_sep_row or pd.isna(val) or val == "None" or str(val) == "nan" or str(val).strip() == "":
                text_val = ""
            else:
                if isinstance(val, (int, float)):
                    try:
                        if val == 0:
                            text_val = "0" if is_internal_hdr else "-"
                        elif val < 0:
                            text_val = f"({abs(val):,.0f})".replace(",", ".")
                        else:
                            text_val = f"{val:,.0f}".replace(",", ".")
                    except:
                        text_val = str(val)
                else:
                    text_val = str(val).strip()
                    
            cell = row_cells[col_idx]
            cell.text = text_val
            set_cell_bg(cell, bg_color)
            set_cell_margins(cell, top=pad_top, bottom=pad_bottom, left=pad_lr, right=pad_lr)
            
            if is_internal_hdr:
                top_b = None
                bottom_b = {'val': 'single', 'sz': '8', 'color': '1F3864'} if is_bottom_internal_hdr else None
                set_cell_borders(cell, top=top_b, bottom=bottom_b)
            elif is_sep_row:
                prev_is_tot = (row_i > 0) and _is_total_or_subtotal_label(str(chunk_df.iloc[row_i - 1, 0]).strip().lower())
                top_b = {'val': 'single', 'sz': '6', 'color': '404040'} if prev_is_tot else None
                set_cell_borders(cell, top=top_b, bottom=None)
            elif is_grand_total or is_subtotal:
                set_cell_borders(cell, top={'val': 'single', 'sz': '6', 'color': '404040'}, bottom={'val': 'single', 'sz': '6', 'color': '404040'})
            else:
                set_cell_borders(cell, top=None, bottom={'val': 'single', 'sz': '2', 'color': 'E1E1E1'})
                
            for p in cell.paragraphs:
                p.paragraph_format.space_before = Pt(0.5)
                p.paragraph_format.space_after = Pt(0.5)
                p.paragraph_format.line_spacing = p_spacing
                
                is_num_hdr = (col_idx >= 1) and (
                    text_val.lower() in ["m$", "$", "%", "activos", "pasivos", "corriente", "no corriente", "us$", "uf", "th$"] or
                    any(k in text_val for k in ["31.12", "30.06", "31.03", "30.09", "202", "31-", "30-"])
                )
                if (col_idx < len(col_classes) and col_classes[col_idx] == 'numeric') or is_num_hdr:
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    set_cell_nowrap(cell)
                elif col_idx < len(col_classes) and col_classes[col_idx] == 'note':
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    set_cell_nowrap(cell)
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    
                for run in p.runs:
                    run.font.name = 'Arial'
                    run.font.size = Pt(hdr_font_size if is_internal_hdr else base_font_size)
                    if is_internal_hdr or is_grand_total or is_subtotal:
                        run.font.bold = True
                    run.font.color.rgb = RGBColor(31, 56, 100) if is_internal_hdr else RGBColor(0, 0, 0)
                        
    apply_table_formatting_and_widths(table, chunk_df, classifications=col_classes, total_width_inches=total_width_inches, is_compact=False, is_landscape=is_landscape)

def _find_vertical_comp_block_start_texts(row_texts_matrix):
    """
    Detecta estructuralmente la fila donde comienza el 2° bloque vertical (Período Comparativo)
    en tablas apiladas verticalmente (ej. #PATRIMONIO, #N08.1, #N09.2, #N14.2, #N15.4, #N16.2, #N16.3).
    Retorna el índice de fila de inicio del 2° bloque (o de su fila separadora), o None si es de 1 solo bloque.
    """
    num_r = len(row_texts_matrix)
    if num_r < 6:
        return None

    import re as _re
    seen_data_or_open = False
    seen_block1_total_or_close = False
    first_block_year = None

    for r_idx, cells_txt in enumerate(row_texts_matrix):
        c0 = (cells_txt[0] if cells_txt else "").strip()
        c0_low = c0.lower()
        all_empty = all(not (t or "").strip() for t in cells_txt)

        # Detectar año solitario en col 0 al inicio del bloque 1 (ej. '2026' o '31-06-2026')
        if r_idx <= 2 and c0 and _re.match(r'^(\d{2}[\.\-\/]\d{2}[\.\-\/])?20\d{2}([\.\-\/]\d{2})?$', c0):
            yrs0 = _re.findall(r'20\d{2}', c0)
            if yrs0:
                first_block_year = yrs0[0]

        if r_idx >= 2 and not all_empty:
            if any(k in c0_low for k in ("saldo inicial", "saldo al 1", "saldo al 01", "opening balance", "detalle al")):
                if seen_block1_total_or_close or (seen_data_or_open and r_idx >= 5):
                    return r_idx
                seen_data_or_open = True
            elif any(
                _re.search(r'\d', (t or "")) and not _re.search(r'20\d{2}', (t or ""))
                for t in cells_txt[1:]
            ):
                seen_data_or_open = True

        # Caso A: Fila separadora completamente vacía después de filas de datos del Bloque 1
        if r_idx >= 4 and all_empty and seen_data_or_open:
            has_non_empty_after = any(
                any((t or "").strip() for t in row_texts_matrix[r_nxt])
                for r_nxt in range(r_idx + 1, num_r)
            )
            if has_non_empty_after:
                return r_idx

        # Detectar cierre del Bloque 1
        if r_idx >= 3 and any(k in c0_low for k in ("saldo final", "saldo al 31", "saldo al 30", "patrimonio total", "total", "totales", "ending balance")):
            if "1 de enero" not in c0_low and "01/01" not in c0_low and "01.01" not in c0_low and "inicial" not in c0_low:
                seen_block1_total_or_close = True
                continue

        # Caso B: Después del cierre del Bloque 1, comienza un nuevo encabezado o un nuevo año/bloque
        if seen_block1_total_or_close and r_idx >= 4 and not all_empty:
            return r_idx

        # Caso C: Segundo encabezado de año solitario en col 0 (ej. #N08.1, #N09.2, #N14.2)
        if r_idx >= 5 and first_block_year and c0 and _re.match(r'^(\d{2}[\.\-\/]\d{2}[\.\-\/])?20\d{2}([\.\-\/]\d{2})?$', c0):
            return r_idx

    return None


def update_table_cells_inplace(table, df, preserve_styles=True, periodo_actual=None, periodo_comp="Ninguno"):
    """
    Actualiza los valores numéricos y fechas de una tabla de Word existente sin alterar
    sus anchos de columnas (w:tcW), márgenes (w:tcMar), bordes (w:tcBorders) ni tipografía.
    Soporta mapeo semántico por concepto (columna 0) y mapeo posicional 1 a 1.
    Si periodo_comp == 'Ninguno':
      - Propuesta 1: Identifica columnas y bloques comparativos por POSICIÓN ESTRUCTURAL FIJA
        (nunca por coincidencia de año), evitando colisiones cuando el año del período actual
        coincide con el año escrito en la columna comparativa de la plantilla Word.
      - Propuesta 2: Toma un Snapshot pre-inyección de todas las celdas comparativas (columnas y
        bloques protegidos) y garantiza su restauración 100% íntegra al finalizar.
    """
    if table is None or df is None:
        return
        
    import unicodedata
    import re
    import numpy as np
    
    clean_df = clean_table_df(df)
    if clean_df.empty:
        return

    is_no_comp = (not periodo_comp) or str(periodo_comp).strip().lower() in ["ninguno", "none", ""]
    act_year = str(periodo_actual)[:4] if periodo_actual and len(str(periodo_actual)) >= 4 else None

    def normalize_text(text):
        if not text:
            return ""
        s = str(text).lower().strip()
        s = unicodedata.normalize('NFKD', s).encode('ASCII', 'ignore').decode('utf-8')
        s = re.sub(r'[^a-z0-9]', '', s)
        return s

    def format_num_val(val):
        if pd.isna(val) or val is None or str(val).strip().lower() in ["", "nan", "none"]:
            return ""
        if isinstance(val, (int, float, np.number)):
            if val == 0:
                return "-"
            elif val < 0:
                return f"({abs(val):,.0f})".replace(",", ".")
            else:
                return f"{val:,.0f}".replace(",", ".")
        return str(val).strip()

    def set_cell_value(cell, text_val, align_right=False):
        first_p = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
        orig_alignment = first_p.alignment
        
        # Eliminar párrafos adicionales en la celda si los hubiera
        for extra_p in list(cell.paragraphs[1:]):
            p_elm = extra_p._p
            if p_elm.getparent() is not None:
                p_elm.getparent().remove(p_elm)
                
        if first_p.runs:
            r0 = first_p.runs[0]
            r0.text = str(text_val)
            for extra_r in list(first_p.runs[1:]):
                r_elm = extra_r._r
                if r_elm.getparent() is not None:
                    r_elm.getparent().remove(r_elm)
        else:
            first_p.add_run(str(text_val))
            
        if orig_alignment is not None:
            first_p.alignment = orig_alignment
        elif align_right:
            first_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    def format_period_into_matched_date(matched_str, target_period_iso, target_dot_date):
        """Reemplaza una fecha detectada conservando el formato exacto de la celda original."""
        if not target_period_iso:
            return matched_str
        t_yr = str(target_period_iso)[:4]
        if len(matched_str) == 4 and re.match(r'^20\d{2}$', matched_str):
            return t_yr
        if re.match(r'^20\d{2}[\.\-\/]\d{2}$', matched_str):
            sep = matched_str[4]
            return str(target_period_iso).replace("-", sep)
        if re.match(r'^20\d{2}-\d{2}-\d{2}', matched_str):
            parts_d = target_dot_date.split(".")
            if len(parts_d) == 3:
                return f"{parts_d[2]}-{parts_d[1]}-{parts_d[0]}"
            return str(target_period_iso)
        if "-" in matched_str and len(matched_str) >= 8:
            return target_dot_date.replace(".", "-")
        if "." in matched_str and len(matched_str) >= 8:
            return target_dot_date
        if "/" in matched_str and len(matched_str) >= 8:
            return target_dot_date.replace(".", "/")
        return str(target_period_iso)

    num_rows_tbl = len(table.rows)
    num_cols_tbl = len(table.columns)
    df_cols = list(clean_df.columns)

    # Identificar columnas cuyos encabezados son atributos de contrato o saldos inicial/final de un mismo período
    contract_col_keywords = ("inicio", "termino", "término", "vencimiento", "emisión", "emision", "issue date", "maturity date")
    movement_col_keywords = ("saldo inicial", "saldo final", "opening balance", "closing balance")
    non_period_cols = set()
    movement_date_cols = set()

    for r_chk in range(min(4, num_rows_tbl)):
        row_c = table.rows[r_chk].cells
        for c_idx in range(1, len(row_c)):
            t_low = row_c[c_idx].text.strip().lower()
            if any(k in t_low for k in contract_col_keywords):
                non_period_cols.add(c_idx)
            elif any(k in t_low for k in movement_col_keywords):
                non_period_cols.add(c_idx)
                movement_date_cols.add(c_idx)

    # 1. Detectar columnas de fecha de período en las cabeceras de Word (primeras filas, excluyendo col 0 y non_period_cols)
    tbl_col_years = {}
    for r_idx in range(min(3, num_rows_tbl)):
        row = table.rows[r_idx]
        c0_txt_low = row.cells[0].text.strip().lower() if len(row.cells) > 0 else ""
        # Evitar tomar como cabecera una fila de datos que tenga montos formateados en otras celdas
        for c_idx in range(1, len(row.cells)):
            if c_idx in non_period_cols:
                continue
            txt = row.cells[c_idx].text.strip()
            years = re.findall(r'20\d{2}', txt)
            if years and c_idx not in tbl_col_years:
                # Validar que el texto sea un encabezado de fecha/período y no un monto largo
                txt_clean_chk = re.sub(r'[\d\.\-\/\s\:\(\)mchth\$us%a-zA-ZáéíóúñÁÉÍÓÚÑ]', '', txt)
                if len(txt) <= 35:
                    tbl_col_years[c_idx] = years[0]

    hdr_date_cols = sorted(tbl_col_years.keys())

    # Detectar si es una tabla apilada verticalmente por bloques de período (ej. #PATRIMONIO, #N08.1, #N09.2, #N15.4, #N16.2, #N16.3)
    protected_tbl_rows = set()
    comp_block_start_row = None
    df_comp_block_start = None
    if not hdr_date_cols:
        tbl_texts_matrix = [[c.text for c in r.cells] for r in table.rows]
        comp_block_start_row = _find_vertical_comp_block_start_texts(tbl_texts_matrix)
        df_texts_matrix = [["" if pd.isna(v) else str(v) for v in r_vals] for _, r_vals in clean_df.iterrows()]
        df_comp_block_start = _find_vertical_comp_block_start_texts(df_texts_matrix)
        if is_no_comp and comp_block_start_row is not None:
            protected_tbl_rows = set(range(comp_block_start_row, num_rows_tbl))

    # 2. Determinar columnas Actuales (act_tbl_cols) y Comparativas Protegidas (protected_tbl_cols) por POSICIÓN ESTRUCTURAL
    protected_tbl_cols = set()
    act_tbl_cols = []

    if is_no_comp and hdr_date_cols:
        r0_joined_low = " ".join(c.text.strip().lower() for c in table.rows[0].cells[1:]) if num_rows_tbl > 0 else ""
        is_interleaved_corr_nocorr = (
            len(hdr_date_cols) == 4 and
            "corriente" in r0_joined_low and
            "no corriente" in r0_joined_low and
            not any(re.search(r'20\d{2}', c.text) for c in table.rows[0].cells[1:])
        )
        if is_interleaved_corr_nocorr:
            # Caso #N14.1: Pares intercalados (Corriente: [Actual, Comp], No corriente: [Actual, Comp])
            act_tbl_cols = [hdr_date_cols[0], hdr_date_cols[2]]
            protected_tbl_cols = {hdr_date_cols[1], hdr_date_cols[3]}
        elif len(hdr_date_cols) == 2:
            group_width = hdr_date_cols[1] - hdr_date_cols[0]
            if group_width <= 1:
                # Caso estándar de 2 columnas de período: 1ª = Actual, 2ª = Comparativo Protegido
                act_tbl_cols = [hdr_date_cols[0]]
                protected_tbl_cols = set(range(hdr_date_cols[1], num_cols_tbl))
            else:
                # Caso de bloques multicolumna bajo cada fecha (#N09.1, #N10.1, #N13.1, #N15.1, #N15.2, #N15.5, #N15.6, #N15.8, #N20.1, #N26.1)
                start_p_col = max(1, min(hdr_date_cols[0], num_cols_tbl - 2 * group_width))
                act_tbl_cols = list(range(start_p_col, start_p_col + group_width))
                protected_tbl_cols = set(range(start_p_col + group_width, num_cols_tbl))
        elif len(hdr_date_cols) > 2:
            # Celdas combinadas que repiten la fecha en cada subcolumna (ej. #N16.1 con 4 celdas de fecha):
            # Primera mitad = Actual, Segunda mitad = Comparativo Protegido
            half_len = len(hdr_date_cols) // 2
            act_tbl_cols = hdr_date_cols[:half_len]
            protected_tbl_cols = set(hdr_date_cols[half_len:])
        elif len(hdr_date_cols) == 1:
            # Solo 1 columna tiene fecha explícita; la primera columna de datos es Actual y las posteriores tras ella se protegen si hay >=3 cols
            c_only = hdr_date_cols[0]
            if c_only == 1 and num_cols_tbl >= 3:
                act_tbl_cols = [1]
                protected_tbl_cols = set(range(2, num_cols_tbl))
            elif c_only == 2 and num_cols_tbl == 3 and "nota" not in (table.rows[0].cells[1].text.lower() if num_rows_tbl > 0 else ""):
                # El template tenía año solo en col 2 o col 1 sin año: por posición col 1 es Actual y col 2 es Comparativo
                act_tbl_cols = [1]
                protected_tbl_cols = {2}
            else:
                act_tbl_cols = [c_only]
                if c_only + 1 < num_cols_tbl:
                    protected_tbl_cols = set(range(c_only + 1, num_cols_tbl))
    elif (
        is_no_comp
        and not hdr_date_cols
        and comp_block_start_row is None
        and not non_period_cols
        and num_cols_tbl == 3
        and len(df_cols) == 3
        and num_rows_tbl >= 2
        and table.rows[0].cells[1].text.strip() == ""
        and table.rows[0].cells[2].text.strip() == ""
    ):
        # Sub-tabla comparativa de 3 columnas sin fila de fechas repetida (ej. #N21.3 'ARRIENDO DE FIBRA IC')
        act_tbl_cols = [1]
        protected_tbl_cols = {2}

    # PROPUESTA 2: Capturar Snapshot inmutable de todas las celdas comparativas (columnas y filas protegidas) antes de modificar la tabla
    protected_cells_snapshot = {}
    if is_no_comp:
        for r_i in range(num_rows_tbl):
            row_cells_snap = table.rows[r_i].cells
            for c_i in range(len(row_cells_snap)):
                if (r_i in protected_tbl_rows) or (c_i in protected_tbl_cols):
                    protected_cells_snapshot[(r_i, c_i)] = row_cells_snap[c_i].text

    def _enforce_protected_snapshot():
        """Verifica y restaura cualquier celda comparativa protegida a su valor exacto original."""
        if is_no_comp and protected_cells_snapshot:
            for (r_i, c_i), orig_txt in protected_cells_snapshot.items():
                if r_i < len(table.rows) and c_i < len(table.rows[r_i].cells):
                    cur_cell = table.rows[r_i].cells[c_i]
                    if cur_cell.text != orig_txt:
                        set_cell_value(cur_cell, orig_txt, align_right=(c_i > 0))

    # 3. Mapeo de columnas de clean_df a columnas de table (df_col_idx -> tbl_col_idx)
    col_mapping = {}
    
    if not hdr_date_cols and not protected_tbl_cols:
        # La tabla no tiene columnas por fecha (ej. Patrimonio, cuadros de movimientos), sino componentes en columnas
        for j in range(min(num_cols_tbl, len(df_cols))):
            col_mapping[j] = j
        act_tbl_cols = [j for j in range(1, num_cols_tbl) if j not in non_period_cols or j in movement_date_cols]
    elif is_no_comp:
        # Modo estricto sin comparativo: SOLO mapear columna de concepto, columnas descriptivas y columnas del período actual
        col_mapping[0] = 0
        if num_cols_tbl == len(df_cols):
            for c_idx in range(1, num_cols_tbl):
                if c_idx not in protected_tbl_cols:
                    col_mapping[c_idx] = c_idx
        else:
            # Mapear columna 'Nota' si existe en ambos
            for j, col_name in enumerate(df_cols):
                if j == 0:
                    continue
                if str(col_name).strip().lower() in ['nota', 'notas']:
                    for tj in range(1, num_cols_tbl):
                        if tj not in protected_tbl_cols and "nota" in table.rows[0].cells[tj].text.lower():
                            col_mapping[j] = tj
                            break
            # Identificar las columnas de datos del período actual en clean_df (por posición: primeras columnas no-Nota)
            df_data_cols = [
                j for j, col_name in enumerate(df_cols)
                if j > 0 and str(col_name).strip().lower() not in ['nota', 'notas', 'comp']
            ]
            for d_c, t_c in zip(df_data_cols, act_tbl_cols):
                if t_c not in protected_tbl_cols:
                    col_mapping[d_c] = t_c
    else:
        # Modo con comparativo normal
        if num_cols_tbl == len(df_cols):
            for j in range(num_cols_tbl):
                col_mapping[j] = j
        else:
            used_tbl_cols = set()
            col_mapping[0] = 0
            used_tbl_cols.add(0)
            for df_idx, col_name in enumerate(df_cols):
                if df_idx == 0:
                    continue
                col_str = str(col_name)
                df_years = re.findall(r'20\d{2}', col_str)
                if df_years:
                    df_yr = df_years[0]
                    for tbl_c, tbl_yr in tbl_col_years.items():
                        if tbl_yr == df_yr and tbl_c not in used_tbl_cols:
                            col_mapping[df_idx] = tbl_c
                            used_tbl_cols.add(tbl_c)
                            break
            if len(col_mapping) <= 1:
                for df_idx in range(1, len(df_cols)):
                    if df_idx < num_cols_tbl:
                        col_mapping[df_idx] = df_idx

    has_header_row = (num_rows_tbl == len(clean_df) + 1)

    def _update_actual_header_dates():
        """Actualiza las fechas en las cabeceras de las columnas del período actual y de movimientos no protegidos."""
        if not periodo_actual:
            return
        from src.reporting.note_generator import format_date_dots
        dot_date = format_date_dots(periodo_actual)
        open_dot = f"01.01.{act_year}" if act_year else dot_date
        date_regex = re.compile(r'(\d{2}[\.\-\/]\d{2}[\.\-\/])?20\d{2}([\.\-\/]\d{2}([\.\-\/]\d{2})?)?')

        if is_no_comp and act_tbl_cols:
            for act_c in act_tbl_cols:
                if act_c in protected_tbl_cols:
                    continue
                for hr_idx in range(min(4, num_rows_tbl)):
                    if (has_header_row and hr_idx == 0) or hr_idx in protected_tbl_rows or act_c >= len(table.rows[hr_idx].cells):
                        continue
                    hdr_cell = table.rows[hr_idx].cells[act_c]
                    hdr_txt = hdr_cell.text
                    m_date = date_regex.search(hdr_txt)
                    if m_date:
                        matched_str = m_date.group(0)
                        new_date_str = format_period_into_matched_date(matched_str, periodo_actual, dot_date)
                        updated_hdr = hdr_txt.replace(matched_str, new_date_str)
                        set_cell_value(hdr_cell, updated_hdr, align_right=True)

        if movement_date_cols:
            for mov_c in movement_date_cols:
                if mov_c in protected_tbl_cols:
                    continue
                for hr_idx in range(min(4, num_rows_tbl)):
                    if hr_idx in protected_tbl_rows or mov_c >= len(table.rows[hr_idx].cells):
                        continue
                    hdr_cell = table.rows[hr_idx].cells[mov_c]
                    hdr_txt = hdr_cell.text
                    hdr_low = hdr_txt.lower()
                    m_date = date_regex.search(hdr_txt)
                    if m_date:
                        matched_str = m_date.group(0)
                        is_open = ("inicial" in hdr_low or "01/01" in hdr_low or "01.01" in hdr_low or "01-01" in hdr_low)
                        target_dot_use = open_dot if is_open else dot_date
                        target_iso_use = f"{act_year}-01" if (is_open and act_year) else periodo_actual
                        new_date_str = format_period_into_matched_date(matched_str, target_iso_use, target_dot_use)
                        set_cell_value(hdr_cell, hdr_txt.replace(matched_str, new_date_str), align_right=False)

    from src.reporting.word_export import _is_total_or_subtotal_label
    from src.reporting.formatting import detect_internal_header_rows, is_separator_row_series

    # Si la tabla existente en Word estaba colapsada/incompleta (menos filas que clean_df, ej. #N24.1 y #N24.2 sin líneas de detalle)
    # y tiene la misma estructura de columnas horizontales de fecha, expandir todas las filas preservando 100% la columna comparativa.
    if (
        num_rows_tbl < len(clean_df)
        and num_cols_tbl == len(df_cols)
        and bool(hdr_date_cols)
        and comp_block_start_row is None
        and df_comp_block_start is None
    ):
        exp_int_hdr_rows, _ = detect_internal_header_rows(clean_df)
        hdr_cutoff = (max(exp_int_hdr_rows) + 1) if exp_int_hdr_rows else min(3, num_rows_tbl)
        df_expanded = clean_df.copy().astype(object)

        if is_no_comp and protected_tbl_cols:
            orig_hdr_comp = {}
            for r_h in range(min(hdr_cutoff, num_rows_tbl)):
                for c_i in protected_tbl_cols:
                    if c_i < len(table.rows[r_h].cells):
                        orig_hdr_comp[(r_h, c_i)] = table.rows[r_h].cells[c_i].text

            orig_data_comp = {}
            for r_d in range(min(hdr_cutoff, num_rows_tbl), num_rows_tbl):
                c0_raw = table.rows[r_d].cells[0].text.strip() if len(table.rows[r_d].cells) > 0 else ""
                c0_n = normalize_text(c0_raw)
                row_comp_vals = {
                    c_i: table.rows[r_d].cells[c_i].text
                    for c_i in protected_tbl_cols
                    if c_i < len(table.rows[r_d].cells)
                }
                if c0_n:
                    orig_data_comp[c0_n] = row_comp_vals
                elif r_d == num_rows_tbl - 1 and len(clean_df) > 0:
                    last_df_norm = normalize_text(str(clean_df.iloc[-1, 0]))
                    if last_df_norm and last_df_norm not in orig_data_comp:
                        orig_data_comp[last_df_norm] = row_comp_vals

            for r_idx in range(len(df_expanded)):
                if r_idx in exp_int_hdr_rows:
                    for c_i in protected_tbl_cols:
                        if (r_idx, c_i) in orig_hdr_comp:
                            df_expanded.iloc[r_idx, c_i] = orig_hdr_comp[(r_idx, c_i)]
                else:
                    c_df_norm = normalize_text(str(df_expanded.iloc[r_idx, 0]))
                    matched_comp_dict = orig_data_comp.get(c_df_norm)
                    if matched_comp_dict is None and len(c_df_norm) >= 7:
                        for k_norm, v_dict in orig_data_comp.items():
                            if k_norm in c_df_norm or c_df_norm in k_norm:
                                matched_comp_dict = v_dict
                                break
                    if matched_comp_dict is not None:
                        for c_i in protected_tbl_cols:
                            df_expanded.iloc[r_idx, c_i] = matched_comp_dict.get(c_i, "-")
                    else:
                        has_act_val = any(
                            pd.notna(df_expanded.iloc[r_idx, ac]) and str(df_expanded.iloc[r_idx, ac]).strip() not in ["", "nan", "None"]
                            for ac in act_tbl_cols if ac < len(df_cols)
                        )
                        for c_i in protected_tbl_cols:
                            df_expanded.iloc[r_idx, c_i] = "-" if has_act_val else ""

        orig_hdr_alignments = {}
        for r_h in range(min(hdr_cutoff, num_rows_tbl)):
            for c_i in range(len(table.rows[r_h].cells)):
                if table.rows[r_h].cells[c_i].paragraphs:
                    p_al = table.rows[r_h].cells[c_i].paragraphs[0].alignment
                    if p_al is not None:
                        orig_hdr_alignments[(r_h, c_i)] = p_al

        populate_notes_docx_table(table, df_expanded)
        num_rows_tbl = len(table.rows)
        _update_actual_header_dates()
        for (r_h, c_i), p_al in orig_hdr_alignments.items():
            if r_h < len(table.rows) and c_i < len(table.rows[r_h].cells) and table.rows[r_h].cells[c_i].paragraphs:
                table.rows[r_h].cells[c_i].paragraphs[0].alignment = p_al
        return

    if num_rows_tbl == len(clean_df):
        int_hdr_rows_all, bot_hdr_rows_all = detect_internal_header_rows(clean_df)
    else:
        int_hdr_rows_all, bot_hdr_rows_all = set(), set()

    def _normalize_total_row_borders():
        for r_i, t_row in enumerate(table.rows):
            if r_i in protected_tbl_rows or not t_row.cells:
                continue
            if r_i in int_hdr_rows_all:
                is_bot_h = (r_i in bot_hdr_rows_all)
                for cell_c in t_row.cells:
                    if not cell_c._tc.xpath('.//w:shd'):
                        set_cell_bg(cell_c, "E8EEF6")
                    if not cell_c._tc.xpath('.//w:tcBorders'):
                        set_cell_borders(
                            cell_c,
                            top=None,
                            bottom={'val': 'single', 'sz': '8', 'color': '1F3864'} if is_bot_h else None
                        )
                    for p_c in cell_c.paragraphs:
                        for r_c in p_c.runs:
                            r_c.font.bold = True
                            r_c.font.color.rgb = RGBColor(31, 56, 100)
                if is_bot_h and (r_i + 1) < len(table.rows) and (r_i + 1) not in protected_tbl_rows:
                    for cell_nxt in table.rows[r_i + 1].cells:
                        if not cell_nxt._tc.xpath('.//w:tcBorders'):
                            set_cell_borders(
                                cell_nxt,
                                top={'val': 'single', 'sz': '8', 'color': '1F3864'},
                                bottom={'val': 'single', 'sz': '2', 'color': 'E1E1E1'}
                            )
                continue
            if r_i > 0 and all(not (c.text or "").strip() for c in t_row.cells):
                prev_is_tot = _is_total_or_subtotal_label((table.rows[r_i - 1].cells[0].text or "").strip().lower()) if len(table.rows[r_i - 1].cells) > 0 else False
                for cell_c in t_row.cells:
                    set_cell_bg(cell_c, "FFFFFF")
                    set_cell_borders(
                        cell_c,
                        top={'val': 'single', 'sz': '6', 'color': '404040'} if prev_is_tot else None,
                        bottom=None
                    )
                continue
            if r_i == 0:
                continue
            # Respetar los bordes (w:tcBorders) que el usuario ya haya configurado manualmente en su plantilla Word
            if any(c._tc.xpath('.//w:tcBorders') for c in t_row.cells):
                continue
            c0_txt = t_row.cells[0].text.strip()
            if _is_total_or_subtotal_label(c0_txt):
                for cell_c in t_row.cells:
                    set_cell_borders(
                        cell_c,
                        top={'val': 'single', 'sz': '6', 'color': '404040'},
                        bottom={'val': 'single', 'sz': '6', 'color': '404040'}
                    )

    has_header_row = (num_rows_tbl == len(clean_df) + 1)

    # Actualizar encabezados de fecha en la fila 0 cuando la tabla tiene fila de cabecera de columnas (ej. #BALANCE, #ER, #ORI, #EFE)
    if has_header_row and num_rows_tbl > 0:
        hdr_row_cells = table.rows[0].cells
        for df_c_idx, tbl_c_idx in col_mapping.items():
            if df_c_idx == 0 or tbl_c_idx >= len(hdr_row_cells) or tbl_c_idx in protected_tbl_cols:
                continue
            col_label = str(df_cols[df_c_idx]).strip()
            if col_label.lower() in ["comp", "nota", "notas", "clasificación", "clasificacion", "descripción", "descripcion", "detalle", ""]:
                continue
            old_hdr_txt = hdr_row_cells[tbl_c_idx].text
            unit_suffix = ""
            for u_tok in ["\nM$", "\nCh$", "\nTh$", " M$", " Ch$", " Th$"]:
                if u_tok in old_hdr_txt and u_tok.strip() not in col_label:
                    unit_suffix = u_tok
                    break
            new_hdr_txt = f"{col_label}{unit_suffix}"
            set_cell_value(hdr_row_cells[tbl_c_idx], new_hdr_txt, align_right=True)

    # En modo comparativo (not is_no_comp), actualizar también fechas en filas de encabezado internas (filas 0..3) de tablas de Notas
    if not is_no_comp and periodo_actual:
        from src.reporting.note_generator import format_date_dots
        dot_act = format_date_dots(periodo_actual)
        dot_comp = format_date_dots(periodo_comp) if periodo_comp and str(periodo_comp).strip().lower() not in ["ninguno", "none", ""] else None
        date_tbl_cols = [c_idx for c_idx in range(1, num_cols_tbl) if c_idx not in protected_tbl_cols and c_idx not in non_period_cols and any(re.search(r'20\d{2}', table.rows[hr].cells[c_idx].text) for hr in range(min(4, num_rows_tbl)) if c_idx < len(table.rows[hr].cells))]
        for pair_pos, t_col in enumerate(date_tbl_cols):
            target_dot = dot_act if (pair_pos % 2 == 0 or not dot_comp) else dot_comp
            target_iso = str(periodo_actual) if (pair_pos % 2 == 0 or not dot_comp) else str(periodo_comp)
            for hr_idx in range(min(4, num_rows_tbl)):
                if has_header_row and hr_idx == 0:
                    continue
                if t_col >= len(table.rows[hr_idx].cells):
                    continue
                hdr_cell = table.rows[hr_idx].cells[t_col]
                hdr_txt = hdr_cell.text
                m_date = re.search(r'(\d{2}[\.\-\/]\d{2}[\.\-\/])?20\d{2}([\.\-\/]\d{2}([\.\-\/]\d{2})?)?', hdr_txt)
                if m_date and len(m_date.group(0)) >= 7:
                    matched_str = m_date.group(0)
                    new_date_str = format_period_into_matched_date(matched_str, target_iso, target_dot)
                    set_cell_value(hdr_cell, hdr_txt.replace(matched_str, new_date_str), align_right=True)

    # Caso 1: Coincidencia posicional exacta de filas (o coincidencia posicional del Bloque 1 en tablas multi-bloque)
    can_use_positional = (num_rows_tbl == len(clean_df) or has_header_row)
    pos_df_rows_limit = len(clean_df)
    row_offset = 1 if has_header_row else 0

    if not can_use_positional and is_no_comp and comp_block_start_row is not None and df_comp_block_start is not None:
        # Verificar si la fila 0 de Word es cabecera de columnas mientras que la fila 0 de clean_df ya es fila de datos
        tbl_r0_c0 = table.rows[0].cells[0].text.strip() if (num_rows_tbl > 0 and len(table.rows[0].cells) > 0) else ""
        df_r0_c0 = str(clean_df.iloc[0, 0]).strip() if (len(clean_df) > 0 and pd.notna(clean_df.iloc[0, 0])) else ""
        tbl_has_extra_top_hdr = (not tbl_r0_c0 and bool(df_r0_c0))
        if not tbl_has_extra_top_hdr and comp_block_start_row == df_comp_block_start:
            can_use_positional = True
            row_offset = 0
            pos_df_rows_limit = df_comp_block_start
        elif tbl_has_extra_top_hdr and comp_block_start_row == df_comp_block_start + 1:
            can_use_positional = True
            row_offset = 1
            pos_df_rows_limit = df_comp_block_start

    if can_use_positional:
        for r_idx in range(pos_df_rows_limit):
            tbl_row_idx = r_idx + row_offset
            if tbl_row_idx >= num_rows_tbl or tbl_row_idx in protected_tbl_rows:
                continue
            df_row = clean_df.iloc[r_idx]
            tbl_row = table.rows[tbl_row_idx]
            # Actualizar etiquetas dinámicas de fecha en columna 0 para filas no protegidas (ej. Bloque 1 de Patrimonio o #N16.2)
            if len(tbl_row.cells) > 0:
                c0_df = str(df_row.iloc[0]).strip() if pd.notna(df_row.iloc[0]) else ""
                c0_tbl = tbl_row.cells[0].text.strip()
                c0_tbl_yrs = re.findall(r'20\d{2}', c0_tbl)
                if c0_df and (re.search(r'20\d{2}', c0_df) or c0_tbl_yrs):
                    if not is_no_comp or len(c0_tbl_yrs) == 1:
                        set_cell_value(tbl_row.cells[0], c0_df, align_right=False)
                elif not c0_tbl and c0_df and tbl_row_idx not in int_hdr_rows_all:
                    set_cell_value(tbl_row.cells[0], c0_df, align_right=False)
            for df_c_idx, tbl_c_idx in col_mapping.items():
                if df_c_idx == 0:
                    continue
                if tbl_c_idx >= len(tbl_row.cells) or tbl_c_idx in protected_tbl_cols:
                    continue
                val = df_row.iloc[df_c_idx]
                fmt_val = format_num_val(val)
                is_num = isinstance(val, (int, float, np.number))
                if fmt_val != "":
                    set_cell_value(tbl_row.cells[tbl_c_idx], fmt_val, align_right=is_num)
                elif tbl_row_idx not in int_hdr_rows_all and (not has_header_row or tbl_row_idx > 0):
                    # Si la celda del período actual tenía un saldo numérico de otro período en el template y ahora es NaN, limpiar a '-'
                    old_cell_txt = tbl_row.cells[tbl_c_idx].text.strip()
                    if old_cell_txt and re.match(r'^[\(\-\d\.,\)]+$', old_cell_txt) and any(ch.isdigit() for ch in old_cell_txt):
                        set_cell_value(tbl_row.cells[tbl_c_idx], "-", align_right=True)
        _update_actual_header_dates()
        _normalize_total_row_borders()
        _enforce_protected_snapshot()
        return

    # Caso 2: Coincidencia semántica por concepto en Columna 0 (respetando bloques verticales protegidos)
    df_for_semantic = clean_df.iloc[:df_comp_block_start] if (is_no_comp and df_comp_block_start is not None) else clean_df
    df_concepts = {}
    df_concepts_by_year = {}
    df_sec_year = act_year
    for df_r_idx, (_, df_row) in enumerate(df_for_semantic.iterrows()):
        row_joined = " ".join(str(v) for v in df_row if pd.notna(v))
        yrs_found = re.findall(r'20\d{2}', row_joined)
        if yrs_found:
            df_sec_year = yrs_found[0]
        concept_raw = str(df_row.iloc[0])
        c_norm = normalize_text(concept_raw)
        if c_norm and c_norm not in ['nan', 'none', 'm', 'detalle', 'concepto']:
            if c_norm not in df_concepts:
                df_concepts[c_norm] = df_row
            if df_sec_year and (df_sec_year, c_norm) not in df_concepts_by_year:
                df_concepts_by_year[(df_sec_year, c_norm)] = df_row

    current_section_year = act_year

    for tbl_r_idx, tbl_row in enumerate(table.rows):
        if tbl_r_idx in protected_tbl_rows or not tbl_row.cells:
            continue
        c_cell = tbl_row.cells[0]
        c_text = c_cell.text.strip()
        c_norm = normalize_text(c_text)
        
        row_full_text = ' '.join(c.text for c in tbl_row.cells)
        found_years = re.findall(r'20\d{2}', row_full_text)
        if found_years:
            current_section_year = found_years[0]
            
        # Si es modo sin comparativo y NO se pudo detectar comp_block_start_row estructuralmente,
        # usar detección de bloque por segundo año solo después de haber pasado las primeras filas del Bloque 1
        if is_no_comp and comp_block_start_row is None and not hdr_date_cols:
            if tbl_r_idx >= 5 and act_year and current_section_year and current_section_year != act_year:
                continue
        
        is_closing_patrimonio = (
            "saldoal" in c_norm and 
            "1deenero" not in c_norm and 
            "0101" not in c_norm and
            "inicial" not in c_norm
        )
        is_opening_patrimonio = (
            ("saldoal" in c_norm or "saldoinicial" in c_norm) and
            ("1deenero" in c_norm or "0101" in c_norm or "inicial" in c_norm)
        )
        
        matched_df_row = None
        if is_closing_patrimonio:
            for _, df_r_cand in df_for_semantic.iterrows():
                cand_norm = normalize_text(str(df_r_cand.iloc[0]))
                if ("saldoal" in cand_norm or "saldofinal" in cand_norm) and "1deenero" not in cand_norm and "0101" not in cand_norm and "inicial" not in cand_norm:
                    matched_df_row = df_r_cand
                    if df_r_cand.iloc[0] and str(df_r_cand.iloc[0]).strip():
                        set_cell_value(c_cell, str(df_r_cand.iloc[0]).strip(), align_right=False)
                    break
        elif is_opening_patrimonio and is_no_comp:
            for _, df_r_cand in df_for_semantic.iterrows():
                cand_norm = normalize_text(str(df_r_cand.iloc[0]))
                if ("saldoal" in cand_norm or "saldoinicial" in cand_norm) and ("1deenero" in cand_norm or "0101" in cand_norm or "inicial" in cand_norm):
                    matched_df_row = df_r_cand
                    if df_r_cand.iloc[0] and str(df_r_cand.iloc[0]).strip():
                        set_cell_value(c_cell, str(df_r_cand.iloc[0]).strip(), align_right=False)
                    break
        elif current_section_year and (current_section_year, c_norm) in df_concepts_by_year:
            matched_df_row = df_concepts_by_year[(current_section_year, c_norm)]
        elif c_norm in df_concepts:
            matched_df_row = df_concepts[c_norm]
        elif not c_norm and tbl_r_idx == num_rows_tbl - 1 and len(df_for_semantic) > 0:
            last_df_r = df_for_semantic.iloc[-1]
            last_df_c0 = str(last_df_r.iloc[0]).strip() if pd.notna(last_df_r.iloc[0]) else ""
            if _is_total_or_subtotal_label(last_df_c0.lower()):
                set_cell_value(c_cell, last_df_c0, align_right=False)
                matched_df_row = last_df_r
        else:
            for (cand_sec_yr, df_c_norm), df_r_candidate in df_concepts_by_year.items():
                if not is_no_comp and current_section_year and cand_sec_yr and cand_sec_yr != current_section_year:
                    continue
                if len(c_norm) >= 7 and (c_norm in df_c_norm or df_c_norm in c_norm):
                    matched_df_row = df_r_candidate
                    break
            if matched_df_row is None:
                for df_c_norm, df_r_candidate in df_concepts.items():
                    if len(c_norm) >= 7 and (c_norm in df_c_norm or df_c_norm in c_norm):
                        if not is_no_comp:
                            cand_str = str(df_r_candidate.iloc[0])
                            cand_yrs = re.findall(r'20\d{2}', cand_str)
                            row_yrs = re.findall(r'20\d{2}', c_text)
                            if cand_yrs and row_yrs and cand_yrs[0] != row_yrs[0]:
                                continue
                        matched_df_row = df_r_candidate
                        break
                    
        if matched_df_row is not None:
            for df_c_idx, tbl_c_idx in col_mapping.items():
                if df_c_idx == 0:
                    continue
                if tbl_c_idx in protected_tbl_cols:
                    continue
                if tbl_c_idx < len(tbl_row.cells):
                    val = matched_df_row.iloc[df_c_idx]
                    fmt_val = format_num_val(val)
                    is_num = isinstance(val, (int, float, np.number))
                    if fmt_val != "":
                        set_cell_value(tbl_row.cells[tbl_c_idx], fmt_val, align_right=is_num)
                    elif tbl_r_idx not in int_hdr_rows_all and tbl_r_idx > 0:
                        old_cell_txt = tbl_row.cells[tbl_c_idx].text.strip()
                        if old_cell_txt and re.match(r'^[\(\-\d\.,\)]+$', old_cell_txt) and any(ch.isdigit() for ch in old_cell_txt):
                            set_cell_value(tbl_row.cells[tbl_c_idx], "-", align_right=True)

    _update_actual_header_dates()
    _normalize_total_row_borders()
    _enforce_protected_snapshot()

def load_mappings_for_entity(empresa):
    """Carga los mapeos de balance y PL para una entidad."""
    empresa_clean = str(empresa).replace("[GRUPO] ", "").strip()
    empresa_path = os.path.join("data", "empresas", empresa)
    if not os.path.exists(empresa_path):
        empresa_path = os.path.join("data", "empresas", empresa_clean)
        
    if empresa.startswith("[GRUPO]"):
        from src.models.database import SessionLocal
        from src.models.consolidacion import ConsolidationGroup
        db = SessionLocal()
        try:
            grupo = db.query(ConsolidationGroup).filter_by(nombre_grupo=empresa_clean).first()
            if grupo and grupo.empresa_matriz:
                empresa_path = os.path.join("data", "empresas", grupo.empresa_matriz)
        finally:
            try:
                db.close()
            except:
                pass
    
    map_balance_df = None
    map_pl_df = None
    
    # 1. Intentar cargar desde los archivos individuales específicos (Source of Truth)
    p_bal = os.path.join(empresa_path, "map_balance.xlsx")
    if os.path.exists(p_bal):
        try:
            map_balance_df = pd.read_excel(p_bal)
        except:
            pass
            
    p_pl = os.path.join(empresa_path, "map_pl.xlsx")
    if os.path.exists(p_pl):
        try:
            map_pl_df = pd.read_excel(p_pl)
        except:
            pass
            
    # 2. Si no existen, intentar cargar desde la plantilla de notas global
    if map_balance_df is None or map_pl_df is None:
        template_nota = "Plantilla de notas_v1.xlsx"
        if os.path.exists(template_nota):
            if map_balance_df is None:
                try:
                    map_balance_df = pd.read_excel(template_nota, sheet_name="Mapeo Balance")
                except:
                    pass
            if map_pl_df is None:
                try:
                    map_pl_df = pd.read_excel(template_nota, sheet_name="Mapeo Ctas P&L Cubo")
                except:
                    pass
                    
    return map_balance_df, map_pl_df

class WordTemplateEngine:
    def __init__(self, template_bytes_io):
        self.template_bytes_io = template_bytes_io

    def _extract_cf_preview(self, result_bytes, col_actual, col_comp):
        result_bytes.seek(0)
        import openpyxl
        import datetime
        import re
        wb_check = openpyxl.load_workbook(result_bytes, data_only=True)
        ws_check = wb_check.active
        
        name_col_idx = 1
        for col in range(1, 10):
            for row in range(1, 15):
                val = ws_check.cell(row=row, column=col).value
                if val and str(val).strip().lower() in ["concepto", "descripcion", "detalle", "flujos", "origen/aplicacion"]:
                    name_col_idx = col
                    break
                    
        date_cols = []
        for col in range(1, ws_check.max_column + 1):
            if col == name_col_idx:
                continue
            for row in range(1, 10):
                val = ws_check.cell(row=row, column=col).value
                if val is not None:
                    is_date = (
                        isinstance(val, (datetime.datetime, datetime.date)) or
                        (isinstance(val, str) and re.search(r'20\d{2}', val))
                    )
                    if is_date:
                        date_cols.append(col)
                        break
        date_cols = sorted(list(set(date_cols)))
        
        val25_col_idx = 3
        val24_col_idx = 4
        if len(date_cols) >= 2:
            val25_col_idx = date_cols[0]
            val24_col_idx = date_cols[1]
            
        header_row_idx = 8
        for row in range(1, 15):
            v25 = ws_check.cell(row=row, column=val25_col_idx).value
            v24 = ws_check.cell(row=row, column=val24_col_idx).value
            v_name = ws_check.cell(row=row, column=name_col_idx).value
            v_str = f"{v25} {v24} {v_name}".lower()
            if any(k in v_str for k in ["flujos de efectivo procedentes", "cobros procedentes", "utilidad (pérdida)", "utilidad (perdida)", "ganancia (pérdida)", "resultado del ejercicio"]):
                header_row_idx = row - 1
                break
            elif any(k in v_str for k in ["actual", "anterior", "m$", "31-12", "31 de"]):
                header_row_idx = row
                
        result_bytes.seek(0)
        df_raw = pd.read_excel(result_bytes, header=None)
        df_data = df_raw.iloc[header_row_idx:].reset_index(drop=True)
        
        def is_cf_title_or_header_row(row_val):
            if pd.isna(row_val):
                return True
            r_str = str(row_val).strip().lower()
            if r_str == "" or r_str == "nan":
                return True
            title_patterns = [
                "estados separados de flujos",
                "estado de flujos de efectivo",
                "por los ejercicios terminados",
                "pacifico cable",
                "db terra chile holdco",
                "nota",
                "m$"
            ]
            return any(pat in r_str for pat in title_patterns)

        cols_to_keep = [name_col_idx - 1, val25_col_idx - 1, val24_col_idx - 1]
        preview_df = df_data.iloc[:, cols_to_keep].copy()
        preview_df.columns = ["Descripción", col_actual, col_comp]
        preview_df = preview_df[~preview_df["Descripción"].apply(is_cf_title_or_header_row)].reset_index(drop=True)
        if str(col_comp).strip().lower() in ["comp", "ninguno", "none", ""]:
            import numpy as np
            preview_df[col_comp] = np.nan

        return preview_df

    def _extract_balance_preview(self, result_bytes, col_actual, col_comp):
        result_bytes.seek(0)
        import openpyxl
        wb_check = openpyxl.load_workbook(result_bytes, data_only=True)
        ws_check = wb_check.active
        
        from src.core.excel_utils import detect_balance_columns, read_template_config, read_excel_preview
        cfg = read_template_config(wb_check)
        if cfg is None:
            name_col_idx, nota_col_idx, val25_col_idx, val24_col_idx = detect_balance_columns(ws_check, wb_check)
            cfg = {
                "name_col": name_col_idx,
                "nota_col": nota_col_idx if nota_col_idx else 0,
                "val_actual_col": val25_col_idx,
                "val_comp_col": val24_col_idx,
                "data_start_row": 5,
            }

        col_act_use = str(col_actual)
        col_cmp_use = str(col_comp)
        if col_act_use == col_cmp_use:
            col_cmp_use = f"{col_cmp_use} (Comp)"

        result_bytes.seek(0)
        preview_df = read_excel_preview(result_bytes, cfg, col_act_use, col_cmp_use)

        if "Nota" in preview_df.columns:
            def clean_nota(val):
                if pd.isna(val) or str(val).strip() == "" or str(val).lower() == "nan" or val == 0:
                    return ""
                try:
                    float_val = float(val)
                    if float_val.is_integer():
                        return str(int(float_val))
                    return str(float_val)
                except Exception:
                    return str(val).strip()
            preview_df["Nota"] = preview_df["Nota"].apply(clean_nota)

        if col_act_use in preview_df.columns:
            preview_df[col_act_use] = pd.to_numeric(preview_df[col_act_use], errors='coerce')
        if col_cmp_use in preview_df.columns:
            if str(col_comp).strip().lower() in ["comp", "ninguno", "none", ""]:
                import numpy as np
                preview_df[col_cmp_use] = np.nan
            else:
                preview_df[col_cmp_use] = pd.to_numeric(preview_df[col_cmp_use], errors='coerce')

        return preview_df

    def _extract_er_preview(self, result_bytes, col_actual, col_comp):
        result_bytes.seek(0)
        import openpyxl
        import datetime
        import re
        wb_check = openpyxl.load_workbook(result_bytes, data_only=True)
        ws_check = wb_check.active
        
        name_col_idx = 1
        for col in range(1, 10):
            for row in range(1, 15):
                val = ws_check.cell(row=row, column=col).value
                if val and str(val).strip().lower() in ["clasificacion", "descripcion", "detalle", "concepto", "ingresos"]:
                    name_col_idx = col
                    break
                    
        date_cols = []
        for col in range(1, ws_check.max_column + 1):
            if col == name_col_idx:
                continue
            for row in range(1, 10):
                val = ws_check.cell(row=row, column=col).value
                if val is not None:
                    is_date = (
                        isinstance(val, (datetime.datetime, datetime.date)) or
                        (isinstance(val, str) and re.search(r'20\d{2}', val))
                    )
                    if is_date:
                        date_cols.append(col)
                        break
        date_cols = sorted(list(set(date_cols)))
        
        val25_col_idx = 3
        val24_col_idx = 4
        if len(date_cols) >= 2:
            val25_col_idx = date_cols[0]
            val24_col_idx = date_cols[1]
            
        nota_col_idx = None
        for col in range(1, ws_check.max_column + 1):
            for row in range(1, 10):
                val = ws_check.cell(row=row, column=col).value
                if val and str(val).strip().lower() == "nota":
                    nota_col_idx = col
                    break
                    
        start_row = 1
        for r in range(1, 15):
            val = ws_check.cell(row=r, column=name_col_idx).value
            if val and any(k in str(val).lower() for k in ["ingresos", "costo", "ganancia", "resultado", "actividades"]):
                start_row = r
                break

        rows_data = []
        has_any_nota = False
        if nota_col_idx is not None:
            for r in range(start_row, ws_check.max_row + 1):
                nv = ws_check.cell(row=r, column=nota_col_idx).value
                if nv is not None and str(nv).strip() not in ["", "nan", "None", "0"]:
                    has_any_nota = True
                    break

        is_no_comp_er = str(col_comp).strip().lower() in ["comp", "ninguno", "none", ""]
        import numpy as np

        for r in range(start_row, ws_check.max_row + 1):
            name_val = ws_check.cell(row=r, column=name_col_idx).value
            v25 = ws_check.cell(row=r, column=val25_col_idx).value
            v24 = np.nan if is_no_comp_er else ws_check.cell(row=r, column=val24_col_idx).value
            
            # Si la fila es completamente vacía, verificar si hay más contenido abajo
            if name_val is None and v25 is None and (v24 is None or pd.isna(v24)):
                has_more = any(ws_check.cell(row=r_next, column=name_col_idx).value is not None for r_next in range(r + 1, ws_check.max_row + 1))
                if not has_more:
                    break
                name_val = ""
                
            row_dict = {
                "Clasificación": str(name_val).strip() if name_val is not None else "",
            }
            if has_any_nota:
                row_dict["Nota"] = ws_check.cell(row=r, column=nota_col_idx).value or ""
            row_dict[col_actual] = v25
            row_dict[col_comp] = v24
                
            rows_data.append(row_dict)
            
        preview_df = pd.DataFrame(rows_data)
        return preview_df

    def _extract_ori_preview(self, result_bytes, col_actual, col_comp):
        result_bytes.seek(0)
        import openpyxl
        import datetime
        import re
        wb_check = openpyxl.load_workbook(result_bytes, data_only=True)
        ws_check = wb_check.active
        
        name_col_idx = 1
        for col in range(1, 10):
            for row in range(1, 15):
                val = ws_check.cell(row=row, column=col).value
                if val and str(val).strip().lower() in ["detalle", "concepto", "clasificacion", "descripcion"]:
                    name_col_idx = col
                    break
                    
        date_cols = []
        for col in range(1, ws_check.max_column + 1):
            if col == name_col_idx:
                continue
            for row in range(1, 10):
                val = ws_check.cell(row=row, column=col).value
                if val is not None:
                    is_date = (
                        isinstance(val, (datetime.datetime, datetime.date)) or
                        (isinstance(val, str) and re.search(r'20\d{2}', val))
                    )
                    if is_date:
                        date_cols.append(col)
                        break
        date_cols = sorted(list(set(date_cols)))
        
        val25_col_idx = 2
        val24_col_idx = 3
        if len(date_cols) >= 2:
            val25_col_idx = date_cols[0]
            val24_col_idx = date_cols[1]
            
        result_bytes.seek(0)
        from src.core.excel_utils import detect_general_skiprows
        ori_skip = detect_general_skiprows(result_bytes)
        result_bytes.seek(0)
        df_raw = pd.read_excel(result_bytes, skiprows=ori_skip)
        cols_to_keep = [name_col_idx - 1, val25_col_idx - 1, val24_col_idx - 1]
        preview_df = df_raw.iloc[:, cols_to_keep].copy()
        preview_df.columns = ["Detalle", col_actual, col_comp]
        preview_df = preview_df.dropna(how='all', subset=["Detalle"]).reset_index(drop=True)
        if str(col_comp).strip().lower() in ["comp", "ninguno", "none", ""]:
            import numpy as np
            preview_df[col_comp] = np.nan
        return preview_df

    def _get_consolidated_bal_and_pl(self, grupo_id, periodo_actual, comp_period, scale_factor, empresa_path, grupo_name):
        from src.core.consolidacion_engine import generar_hoja_trabajo
        import openpyxl
        df_hoja_act, _ = generar_hoja_trabajo(grupo_id, periodo_actual)
        df_hoja_comp = None
        if comp_period != "Ninguno":
            df_hoja_comp, _ = generar_hoja_trabajo(grupo_id, comp_period)
            
        def clean_str(s):
            if pd.isna(s): return ""
            return str(s).strip().lower().replace('á','a').replace('é','e').replace('í','i').replace('ó','o').replace('ú','u')

        bs_subtotals = {
            "activos corrientes", "activos corrientes totales",
            "activos no corrientes", "activos no corrientes totales",
            "total activos", "patrimonio y pasivos",
            "pasivos corrientes", "pasivo corrientes totales",
            "pasivos no corrientes", "pasivo no corrientes totales",
            "total pasivos", "patrimonio", "patrimonio total",
            "total patrimonio y pasivos", "estado de resultados"
        }
        bs_subtotals_normalized = {clean_str(name) for name in bs_subtotals}
        
        idx_er_b = df_hoja_act[df_hoja_act['Balance clasificado'] == "Estado de Resultados"].index
        df_hoja_act_sec_b = df_hoja_act.loc[:idx_er_b[0]-1] if not idx_er_b.empty else df_hoja_act
        df_hoja_act_clean_b = df_hoja_act_sec_b[df_hoja_act_sec_b['Balance clasificado'].notna() & (df_hoja_act_sec_b['Balance clasificado'].str.strip() != "")]
        df_hoja_act_clean_b = df_hoja_act_clean_b[~df_hoja_act_clean_b['Balance clasificado'].apply(clean_str).isin(bs_subtotals_normalized)]
        
        tb_df_bal = pd.DataFrame({
            'cuenta_id': df_hoja_act_clean_b['Balance clasificado'],
            'saldo_final': df_hoja_act_clean_b['CONSOLIDADO']
        })
        tb_df_comp_bal = None
        if df_hoja_comp is not None:
            idx_er_comp_b = df_hoja_comp[df_hoja_comp['Balance clasificado'] == "Estado de Resultados"].index
            df_hoja_comp_sec_b = df_hoja_comp.loc[:idx_er_comp_b[0]-1] if not idx_er_comp_b.empty else df_hoja_comp
            df_hoja_comp_clean_b = df_hoja_comp_sec_b[df_hoja_comp_sec_b['Balance clasificado'].notna() & (df_hoja_comp_sec_b['Balance clasificado'].str.strip() != "")]
            df_hoja_comp_clean_b = df_hoja_comp_clean_b[~df_hoja_comp_clean_b['Balance clasificado'].apply(clean_str).isin(bs_subtotals_normalized)]
            tb_df_comp_bal = pd.DataFrame({
                'cuenta_id': df_hoja_comp_clean_b['Balance clasificado'],
                'saldo_final': df_hoja_comp_clean_b['CONSOLIDADO']
            })
            
        def is_pl_line_local(li):
            norm = str(li).lower().strip().replace('á','a').replace('é','e').replace('í','i').replace('ó','o').replace('ú','u')
            if any(x in norm for x in ["ingreso", "gasto", "costo", "egreso", "depreciacion", "amortizacion", "diferencia de cambio", "unidad de reajuste", "ganancia", "perdida"]):
                if not any(x in norm for x in ["diferido", "diferidos", "anticipado", "anticipados", "por pagar", "por cobrar", "acumulado", "acumulados", "acumulada", "acumuladas", "capital", "reserva", "ejercicio", "periodo"]):
                    return True
            return False
            
        tb_df_bal = tb_df_bal[~tb_df_bal['cuenta_id'].apply(is_pl_line_local)]
        if tb_df_comp_bal is not None:
            tb_df_comp_bal = tb_df_comp_bal[~tb_df_comp_bal['cuenta_id'].apply(is_pl_line_local)]
            
        template_bal_path = os.path.join(empresa_path, "Balance clasificado.xlsx")
        if not os.path.exists(template_bal_path):
            template_bal_path = os.path.join("templates", "Balance clasificado.xlsx")
            
        wb_bal = openpyxl.load_workbook(template_bal_path)
        ws_bal = wb_bal.active
        template_accounts = [str(ws_bal.cell(row=r, column=2).value).strip() for r in range(1, ws_bal.max_row+1) if ws_bal.cell(row=r, column=2).value]
        
        dummy_map = pd.DataFrame({
            'N° de Cuenta': tb_df_bal['cuenta_id'],
            'Clasificación balance': tb_df_bal['cuenta_id']
        })
        dummy_map.loc[~dummy_map['N° de Cuenta'].isin(template_accounts), 'Clasificación balance'] = pd.NA
        
        from src.reporting.balance_generator import BalanceGenerator
        gen_bal = BalanceGenerator(template_bal_path)
        result_bytes_bal = gen_bal.generate(
            tb_df=tb_df_bal, 
            map_balance_df=dummy_map, 
            scale_factor=scale_factor,
            tb_df_comp=tb_df_comp_bal,
            periodo_actual_str=periodo_actual,
            periodo_comp_str=comp_period if comp_period != "Ninguno" else None
        )
        
        col_actual_tmp = str(periodo_actual)
        col_comp_tmp = str(comp_period) if comp_period != "Ninguno" else "Comp"
        if col_actual_tmp == col_comp_tmp:
            col_comp_tmp = f"{col_comp_tmp} (Comp)"
            
        from src.core.excel_utils import detect_balance_columns, read_template_config, read_excel_preview
        wb_check_bal = openpyxl.load_workbook(result_bytes_bal, data_only=True)
        ws_check_bal = wb_check_bal.active
        cfg_bal = read_template_config(wb_check_bal)
        if cfg_bal is None:
            name_col_idx_b, nota_col_idx_b, val25_col_idx_b, val24_col_idx_b = detect_balance_columns(ws_check_bal, wb_check_bal)
            cfg_bal = {
                "name_col": name_col_idx_b,
                "nota_col": nota_col_idx_b if nota_col_idx_b else 0,
                "val_actual_col": val25_col_idx_b,
                "val_comp_col": val24_col_idx_b,
                "data_start_row": 5,
            }
        result_bytes_bal.seek(0)
        bal_df = read_excel_preview(result_bytes_bal, cfg_bal, col_actual_tmp, col_comp_tmp)
        if "Nota" in bal_df.columns:
            bal_df = bal_df.drop(columns=["Nota"])
            
        pl_subtotals = {
            "ganancia bruta", "resultado antes de impuestos",
            "ganancias (perdida) del ejercicio", "ganancia (perdida) del ejercicio",
            "estado de resultados", "otros rubros no clasificados"
        }
        pl_subtotals_normalized = {clean_str(name) for name in pl_subtotals}
        
        idx_er_p = df_hoja_act[df_hoja_act['Balance clasificado'] == "Estado de Resultados"].index
        df_hoja_act_sec_pl = df_hoja_act.loc[idx_er_p[0]+1:] if not idx_er_p.empty else df_hoja_act
        df_hoja_act_clean_pl = df_hoja_act_sec_pl[df_hoja_act_sec_pl['Balance clasificado'].notna() & (df_hoja_act_sec_pl['Balance clasificado'].str.strip() != "")]
        df_hoja_act_clean_pl = df_hoja_act_clean_pl[~df_hoja_act_clean_pl['Balance clasificado'].apply(clean_str).isin(pl_subtotals_normalized)]
        
        pl_dict_act = {row['Balance clasificado']: [row['CONSOLIDADO']] for _, row in df_hoja_act_clean_pl.iterrows()}
        pl_df_wide = pd.DataFrame(pl_dict_act)
        
        pl_df_comp_wide = None
        if df_hoja_comp is not None:
            idx_er_comp_p = df_hoja_comp[df_hoja_comp['Balance clasificado'] == "Estado de Resultados"].index
            df_hoja_comp_sec_pl = df_hoja_comp.loc[idx_er_comp_p[0]+1:] if not idx_er_comp_p.empty else df_hoja_comp
            df_hoja_comp_clean_pl = df_hoja_comp_sec_pl[df_hoja_comp_sec_pl['Balance clasificado'].notna() & (df_hoja_comp_sec_pl['Balance clasificado'].str.strip() != "")]
            df_hoja_comp_clean_pl = df_hoja_comp_clean_pl[~df_hoja_comp_clean_pl['Balance clasificado'].apply(clean_str).isin(pl_subtotals_normalized)]
            pl_dict_comp = {row['Balance clasificado']: [row['CONSOLIDADO']] for _, row in df_hoja_comp_clean_pl.iterrows()}
            pl_df_comp_wide = pd.DataFrame(pl_dict_comp)
            
        template_er_path = os.path.join(empresa_path, "Estado de Resultados Clasificados.xlsx")
        if not os.path.exists(template_er_path):
            template_er_path = os.path.join("templates", "Estado de Resultados Clasificados.xlsx")
            
        from src.reporting.er_generator import ERGenerator
        gen_er = ERGenerator(template_er_path)
        result_bytes_er, pl_df = gen_er.generate(
            pl_df=pl_df_wide, 
            scale_factor=scale_factor,
            pl_df_comp=pl_df_comp_wide,
            periodo_actual_str=periodo_actual,
            periodo_comp_str=comp_period if comp_period != "Ninguno" else None
        )
        return bal_df, pl_df, result_bytes_bal, result_bytes_er

    def process(self, empresa_activa, periodo_actual, periodo_comp="Ninguno", scale_factor=1.0, use_ifrs_auto=True, include_headers=True, progress_callback=None):
        is_consolidated = empresa_activa.startswith("[GRUPO]")
        empresa_path = os.path.join("data", "empresas", empresa_activa)
        
        # 1. Cargar mapeos de balance/PL
        map_balance_df, map_pl_df = load_mappings_for_entity(empresa_activa)
        
        # 2. Calcular variables de texto
        periodo_actual_str = str(periodo_actual)
        periodo_comp_str = str(periodo_comp)
        
        # Calcular periodos comparativos específicos según reglas IFRS o manuales
        is_no_comp_param = (not periodo_comp) or str(periodo_comp).strip().lower() in ["ninguno", "none", ""]
        if is_no_comp_param:
            periodo_comp = "Ninguno"
            periodo_comp_balance = "Ninguno"
            periodo_comp_resultados = "Ninguno"
        else:
            periodo_comp_balance = periodo_comp
            periodo_comp_resultados = periodo_comp
            if use_ifrs_auto:
                try:
                    parts = str(periodo_actual).split('-')
                    if len(parts) >= 2:
                        year = int(parts[0])
                        month = parts[1]
                        prev_year = str(year - 1)
                        periodo_comp_balance = f"{prev_year}-12"
                        periodo_comp_resultados = f"{prev_year}-{month}"
                except Exception:
                    pass
        
        from src.core.excel_utils import format_periodo
        txt_periodo_act = format_periodo(periodo_actual)
        txt_periodo_comp = format_periodo(periodo_comp) if periodo_comp != "Ninguno" else ""
        
        # Obtener utilidad neta
        from src.core.snapshot_vault_manager import SnapshotVaultManager
        val_utilidad_pesos = 0.0
        if is_consolidated:
            grupo_name = empresa_activa.replace("[GRUPO] ", "").strip()
            db = SessionLocal()
            try:
                grupo_obj = db.query(ConsolidationGroup).filter_by(nombre_grupo=grupo_name).first()
                grupo_id = grupo_obj.id if grupo_obj else None
            finally:
                db.close()
                
            if grupo_id:
                from src.core.consolidacion_engine import generar_hoja_trabajo
                df_hoja_act, msg_act = generar_hoja_trabajo(grupo_id, periodo_actual)
                if df_hoja_act is not None:
                    idx_er_act = df_hoja_act[df_hoja_act['Balance clasificado'] == "Estado de Resultados"].index
                    df_sec_act = df_hoja_act.loc[idx_er_act[0]+1:] if not idx_er_act.empty else df_hoja_act
                    row_hoja_act = df_sec_act[df_sec_act['Balance clasificado'].astype(str).str.strip() == "Ganancias (Pérdida) del Ejercicio"]
                    val_utilidad_pesos = -1 * float(row_hoja_act['CONSOLIDADO'].values[0]) if not row_hoja_act.empty else 0.0
                else:
                    raise ValueError(f"No hay información financiera cargada para consolidar el grupo '{grupo_name}' en el periodo {periodo_actual}. Por favor verifica que la empresa matriz tenga cargado dicho periodo.")
        else:
            from src.core.snapshot_vault_manager import SnapshotVaultManager
            has_snap = SnapshotVaultManager.has_active_snapshot(empresa_activa, periodo_actual)
            tb_check = TrialBalanceDB.get_trial_balance(empresa_activa, periodo_actual)
            if (tb_check is None or tb_check.empty) and not has_snap:
                raise ValueError(f"No se encontraron saldos de Balance para la empresa '{empresa_activa}' en el periodo {periodo_actual}. Por favor sube el archivo en '1️⃣ Cargas de datos'.")
            val_utilidad_pesos = -1 * PlCuboDB.get_pl_cubo_total_sum(empresa_activa, periodo_actual)
            # Si no hay saldo en Cubo pero hay snapshot, obtener utilidad de snapshot ER
            if val_utilidad_pesos == 0.0 and has_snap:
                er_snap = SnapshotVaultManager.get_statement(empresa_activa, periodo_actual, "#ER", scale_factor=1.0)
                if er_snap and er_snap[0] is not None:
                    er_df = er_snap[0]
                    # Buscar fila de ganancia/resultado del ejercicio
                    for _, r_er in er_df.iterrows():
                        r_name = str(r_er.iloc[0]).lower().strip()
                        if any(k in r_name for k in ["ganancia (pérdida) del ejercicio", "ganancia (perdida) del ejercicio", "resultado del ejercicio"]):
                            val_utilidad_pesos = float(r_er.iloc[1]) if pd.notna(r_er.iloc[1]) else 0.0
                            break
            
        txt_utilidad_neta = ""
        if val_utilidad_pesos < 0:
            txt_utilidad_neta = f"({abs(val_utilidad_pesos / scale_factor):,.0f})".replace(",", ".")
        else:
            txt_utilidad_neta = f"{val_utilidad_pesos / scale_factor:,.0f}".replace(",", ".")
            
        # Contexto de variables Jinja2 para docxtpl
        context = {
            "EMPRESA_ACTIVA": empresa_activa.replace("[GRUPO] ", "").strip(),
            "PERIODO_ACTUAL": txt_periodo_act,
            "PERIODO_COMPARATIVO": txt_periodo_comp,
            "UTILIDAD_NETA": txt_utilidad_neta,
            "empresa_activa": empresa_activa.replace("[GRUPO] ", "").strip(),
            "periodo_actual": txt_periodo_act,
            "periodo_comparativo": txt_periodo_comp,
            "utilidad_neta": txt_utilidad_neta
        }

        # Cargar documento docx y procesar variables de plantilla de forma ultrarrápida
        self.template_bytes_io.seek(0)
        doc = Document(self.template_bytes_io)

        # Detectar si hay variables de texto tipo {{VARIABLE}} (no {{#...}})
        import zipfile
        has_jinja_vars = False
        try:
            self.template_bytes_io.seek(0)
            with zipfile.ZipFile(self.template_bytes_io) as z:
                for name in z.namelist():
                    if name.startswith('word/') and name.endswith('.xml'):
                        c_xml = z.read(name).decode('utf-8', errors='ignore')
                        if re.search(r'\{\{\s*[^#\s]', c_xml):
                            has_jinja_vars = True
                            break
        except Exception:
            has_jinja_vars = False

        if has_jinja_vars:
            def replace_vars(p_obj):
                for k, v in context.items():
                    if k in p_obj.text:
                        p_obj.text = re.sub(r'\{\{\s*' + re.escape(k) + r'\s*\}\}', str(v), p_obj.text)

            for p in doc.paragraphs:
                if '{{' in p.text:
                    replace_vars(p)
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            if '{{' in p.text:
                                replace_vars(p)

        target_lang = context.get('target_lang')
        if not target_lang:
            try:
                import streamlit as st
                target_lang = st.session_state.get('idioma_reporte', 'es')
            except Exception:
                target_lang = 'es'

        # 3. Procesar e inyectar reportes principales y notas
        # Recolectar todos los párrafos (tanto del cuerpo principal como de tablas)
        all_paragraphs = list(doc.paragraphs)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        all_paragraphs.append(p)

        matched_items = []
        seen_tags = set()

        # 1. Párrafos con etiquetas {{#...}}
        for p in all_paragraphs:
            text_clean = p.text.replace('\xa0', ' ').replace('\u200b', '')
            m = re.search(r'\{\{\s*(#[\w\.\:\-]+)\s*\}\}', text_clean)
            if m:
                tag = m.group(1).strip()
                matched_items.append(('paragraph', p, tag))
                seen_tags.add(tag.split(':', 1)[0].strip().upper())

        # 2. Tablas existentes con metadatos w:tblCaption / w:tblDescription
        note_subtable_counts = {}
        duplicate_tables_to_remove = []
        for tbl in doc.tables:
            meta = get_table_metadata(tbl)
            if meta:
                meta_clean = meta.strip()
                parts_m = meta_clean.split(':', 1)
                base_meta = parts_m[0].strip().upper()
                suffix_m = f":{parts_m[1]}" if len(parts_m) > 1 else ""
                m_note_dot1 = re.match(r'^(#N\d+)\.1$', base_meta)
                if m_note_dot1:
                    n_prefix = m_note_dot1.group(1)
                    note_subtable_counts[n_prefix] = note_subtable_counts.get(n_prefix, 0) + 1
                    if base_meta in seen_tags and note_subtable_counts[n_prefix] > 1:
                        healed_base = f"{n_prefix}.{note_subtable_counts[n_prefix]}"
                        while healed_base in seen_tags:
                            note_subtable_counts[n_prefix] += 1
                            healed_base = f"{n_prefix}.{note_subtable_counts[n_prefix]}"
                        base_meta = healed_base
                        meta_clean = f"{healed_base}{suffix_m}"
                if base_meta not in seen_tags:
                    matched_items.append(('table', tbl, meta_clean))
                    seen_tags.add(base_meta)
                else:
                    duplicate_tables_to_remove.append((tbl, base_meta))

        # 3. Detección inteligente por títulos y contexto para tablas NO marcadas aún
        # (SOLO si no se detectaron etiquetas explícitas ni metadatos previos)
        if not matched_items:
            seen_table_elements = {target._tbl for item_type, target, _ in matched_items if item_type == 'table'}
            
            def is_statement_heading(txt):
                txt_clean = txt.strip().upper()
                if len(txt_clean) > 80:
                    return None
                if any(w in txt_clean for w in ['REGISTRAD', 'ADJUNTO', 'PRESENTACIÓN EN', 'ADOPCIÓN', 'DESCRIT', 'RECLASIFICACI', 'DURANTE EL AÑO', 'RESPECTIVAMENTE']):
                    return None
                if any(p in txt_clean for p in ['ESTADO DE SITUACIÓN FINANCIERA', 'ESTADOS SEPARADOS DE SITUACIÓN FINANCIERA', 'BALANCE GENERAL', 'BALANCE CLASIFICADO']):
                    return '#BALANCE'
                if any(p in txt_clean for p in ['ESTADO DE RESULTADO INTEGRAL', 'ESTADOS SEPARADOS DE RESULTADOS INTEGRALES', 'ESTADO DE OTROS RESULTADOS INTEGRALES']):
                    return '#ORI'
                if any(p in txt_clean for p in ['ESTADO DE RESULTADOS', 'ESTADOS SEPARADOS DE RESULTADOS', 'ESTADO DE RESULTADO POR FUNCIÓN']):
                    return '#ER'
                if any(p in txt_clean for p in ['ESTADO DE CAMBIOS EN EL PATRIMONIO', 'ESTADOS SEPARADOS DE CAMBIOS EN EL PATRIMONIO']):
                    return '#PATRIMONIO'
                if any(p in txt_clean for p in ['ESTADO DE FLUJOS DE EFECTIVO', 'ESTADOS SEPARADOS DE FLUJOS DE EFECTIVO']):
                    return '#EFE'
                return None

            current_sec = None
            sec_tbl_idx = 0
            for el in doc.element.body:
                if el.tag.endswith('p'):
                    txt = el.text.strip().upper() if el.text else ''
                    m_n = re.match(r'^NOTA\s+(\d+)', txt)
                    if m_n:
                        current_sec = f"#N{int(m_n.group(1)):02d}"
                        sec_tbl_idx = 0
                    elif not (current_sec and current_sec.startswith('#N')):
                        stmt = is_statement_heading(txt)
                        if stmt:
                            current_sec = stmt
                            sec_tbl_idx = 0
                elif el.tag.endswith('tbl'):
                    if el in seen_table_elements:
                        continue
                    if current_sec:
                        sec_tbl_idx += 1
                        if current_sec.startswith('#N'):
                            code = f"{current_sec}.{sec_tbl_idx}"
                            tbl_obj = docx.table.Table(el, doc)
                            matched_items.append(('table', tbl_obj, code))
                            seen_table_elements.add(el)
                        elif current_sec not in seen_tags and sec_tbl_idx == 1:
                            code = current_sec
                            tbl_obj = docx.table.Table(el, doc)
                            matched_items.append(('table', tbl_obj, code))
                            seen_tags.add(current_sec)
                            seen_table_elements.add(el)

        total_tags = len(matched_items)
        self.total_tags_found = total_tags

        # Caches para optimizar drásticamente la velocidad
        contexts_cache = {}
        def get_entity_contexts(comp_period_str):
            if comp_period_str not in contexts_cache:
                from src.ui_pages.informes_y_notas import load_all_entity_contexts
                contexts_cache[comp_period_str] = load_all_entity_contexts(
                    active_entity=empresa_activa,
                    periodo_actual=periodo_actual,
                    periodo_comp=comp_period_str,
                    map_balance_df=map_balance_df,
                    map_pl_df=map_pl_df
                )
            return contexts_cache[comp_period_str]

        hoja_trabajo_cache = {}
        def get_consolidated_hoja(g_id, p_str):
            if p_str not in hoja_trabajo_cache:
                from src.core.consolidacion_engine import generar_hoja_trabajo
                df_h, _ = generar_hoja_trabajo(g_id, p_str)
                hoja_trabajo_cache[p_str] = df_h
            return hoja_trabajo_cache[p_str]

        statements_cache = {}  # Key: (base_code_clean, comp_period) -> (preview_df, report_name)
        note_elements_cache = {}  # Key: (base_code, note_comp) -> list of elements
        from src.core.rounding_reconciler import RoundingReconciler
        RoundingReconciler.clear_cache()

        template_nota = "Plantilla de notas_v1.xlsx"
        engine_notes = None
        if os.path.exists(template_nota):
            from src.reporting.note_generator import NoteGenerator
            engine_notes = NoteGenerator(template_nota)

        paragraphs_to_delete = []
        
        for idx, (item_type, target_item, tag_content) in enumerate(matched_items):
            p = target_item if item_type == 'paragraph' else None
            old_tbl = target_item if item_type == 'table' else None
            if progress_callback:
                try:
                    progress_callback(idx + 1, total_tags, f"{{{{{tag_content}}}}}")
                except Exception:
                    pass

            try:
                if ":" in tag_content:
                    parts = tag_content.split(":", 1)
                    code = parts[0].strip() # ej. #BALANCE
                    override_period = parts[1].strip() # ej. 2025-12
                else:
                    code = tag_content
                    override_period = None
                
                # Normalizar código para detectar si es solo tabla o con encabezado completo
                code_upper = code.upper()
                is_only_table = (not include_headers) or code_upper.startswith("#TABLA_") or code_upper.endswith("_TABLA")
                base_code_clean = code_upper
                if base_code_clean.startswith("#TABLA_"):
                    base_code_clean = "#" + base_code_clean.replace("#TABLA_", "")
                elif base_code_clean.endswith("_TABLA"):
                    base_code_clean = base_code_clean.replace("_TABLA", "")
                
                # --- CASO REPORTES PRINCIPALES ---
                preview_df = None
                report_name = ""
                result_bytes = None
                
                # Seleccionar comparativo dinámico o manual (IFRS: Balance, EFE y Patrimonio usan cierre anual anterior; ER y ORI usan mismo periodo anterior)
                comp_period = override_period if override_period else (periodo_comp_balance if base_code_clean in ["#BALANCE", "#ESF", "#EFE", "#PATRIMONIO"] else periodo_comp_resultados)
                
                stmt_cache_key = (base_code_clean, comp_period)
                was_loaded_from_snapshot = False
                if base_code_clean in ["#BALANCE", "#ESF", "#ER", "#ORI", "#PATRIMONIO", "#EFE"] and stmt_cache_key in statements_cache:
                    preview_df, report_name = statements_cache[stmt_cache_key]
                    was_loaded_from_snapshot = True
                elif base_code_clean in ["#BALANCE", "#ESF", "#ER", "#ORI", "#PATRIMONIO", "#EFE"] and SnapshotVaultManager.has_active_snapshot(empresa_activa, periodo_actual):
                    snap_stmt = SnapshotVaultManager.get_statement(
                        empresa=empresa_activa,
                        periodo_actual=periodo_actual,
                        code=base_code_clean,
                        comp_period=comp_period,
                        scale_factor=scale_factor
                    )
                    if snap_stmt and snap_stmt[0] is not None:
                        preview_df, report_name = snap_stmt
                        statements_cache[stmt_cache_key] = (preview_df, report_name)
                        was_loaded_from_snapshot = True
                        result_bytes = SnapshotVaultManager.get_statement_excel(
                            empresa=empresa_activa,
                            periodo_actual=periodo_actual,
                            code=base_code_clean,
                            comp_period=comp_period,
                            scale_factor=scale_factor,
                            df_override=preview_df,
                            target_lang=target_lang,
                        )
                
                if preview_df is None and base_code_clean in ["#BALANCE", "#ESF"]:
                    report_name = "ESTADO DE SITUACIÓN FINANCIERA CONSOLIDADO" if is_consolidated else "ESTADO DE SITUACIÓN FINANCIERA CLASIFICADO"
                    if is_consolidated:
                        grupo_name = empresa_activa.replace("[GRUPO] ", "").strip()
                        db = SessionLocal()
                        try:
                            grupo_obj = db.query(ConsolidationGroup).filter_by(nombre_grupo=grupo_name).first()
                            grupo_id = grupo_obj.id
                        finally:
                            db.close()
                        df_hoja_act = get_consolidated_hoja(grupo_id, periodo_actual)
                        if df_hoja_act is None:
                            raise ValueError(f"No hay datos para consolidar el grupo '{grupo_name}' en el periodo {periodo_actual}. Verifica que la empresa matriz tenga cargado dicho periodo.")
                            
                        df_hoja_comp = None
                        if comp_period != "Ninguno":
                            df_hoja_comp = get_consolidated_hoja(grupo_id, comp_period)
                        
                        def clean_str(s):
                            if pd.isna(s): return ""
                            return str(s).strip().lower().replace('á','a').replace('é','e').replace('í','i').replace('ó','o').replace('ú','u')

                        bs_subtotals = {
                            "activos corrientes", "activos corrientes totales",
                            "activos no corrientes", "activos no corrientes totales",
                            "total activos", "patrimonio y pasivos",
                            "pasivos corrientes", "pasivo corrientes totales",
                            "pasivos no corrientes", "pasivo no corrientes totales",
                            "total pasivos", "patrimonio", "patrimonio total",
                            "total patrimonio y pasivos", "estado de resultados"
                        }
                        bs_subtotals_normalized = {clean_str(name) for name in bs_subtotals}

                        # Slice assets
                        idx_er = df_hoja_act[df_hoja_act['Balance clasificado'] == "Estado de Resultados"].index
                        df_hoja_act_sec = df_hoja_act.loc[:idx_er[0]-1] if not idx_er.empty else df_hoja_act
                        df_hoja_act_clean = df_hoja_act_sec[df_hoja_act_sec['Balance clasificado'].notna() & (df_hoja_act_sec['Balance clasificado'].str.strip() != "")]
                        df_hoja_act_clean = df_hoja_act_clean[~df_hoja_act_clean['Balance clasificado'].apply(clean_str).isin(bs_subtotals_normalized)]
                        
                        tb_df = pd.DataFrame({'cuenta_id': df_hoja_act_clean['Balance clasificado'], 'saldo_final': df_hoja_act_clean['CONSOLIDADO']})
                        tb_df_comp = None
                        if df_hoja_comp is not None:
                            idx_er_comp = df_hoja_comp[df_hoja_comp['Balance clasificado'] == "Estado de Resultados"].index
                            df_hoja_comp_sec = df_hoja_comp.loc[:idx_er_comp[0]-1] if not idx_er_comp.empty else df_hoja_comp
                            df_hoja_comp_clean = df_hoja_comp_sec[df_hoja_comp_sec['Balance clasificado'].notna() & (df_hoja_comp_sec['Balance clasificado'].str.strip() != "")]
                            df_hoja_comp_clean = df_hoja_comp_clean[~df_hoja_comp_clean['Balance clasificado'].apply(clean_str).isin(bs_subtotals_normalized)]
                            tb_df_comp = pd.DataFrame({'cuenta_id': df_hoja_comp_clean['Balance clasificado'], 'saldo_final': df_hoja_comp_clean['CONSOLIDADO']})
                            
                        # Excluir cuentas de resultado para evitar distorsiones en el Balance
                        def is_pl_line_local(li):
                            norm = str(li).lower().strip().replace('á','a').replace('é','e').replace('í','i').replace('ó','o').replace('ú','u')
                            if any(x in norm for x in ["ingreso", "gasto", "costo", "egreso", "depreciacion", "amortizacion", "diferencia de cambio", "unidad de reajuste", "ganancia", "perdida"]):
                                if not any(x in norm for x in ["diferido", "diferidos", "anticipado", "anticipados", "por pagar", "por cobrar", "acumulado", "acumulados", "acumulada", "acumuladas", "capital", "reserva", "ejercicio", "periodo"]):
                                    return True
                            return False
                            
                        tb_df = tb_df[~tb_df['cuenta_id'].apply(is_pl_line_local)]
                        if tb_df_comp is not None:
                            tb_df_comp = tb_df_comp[~tb_df_comp['cuenta_id'].apply(is_pl_line_local)]
                            
                        template_bal_path = os.path.join(empresa_path, "Balance clasificado.xlsx")
                        if not os.path.exists(template_bal_path):
                            template_bal_path = os.path.join("templates", "Balance clasificado.xlsx")
                            
                        dummy_map = pd.DataFrame({'N° de Cuenta': tb_df['cuenta_id'], 'Clasificación balance': tb_df['cuenta_id']})
                        from src.reporting.balance_generator import BalanceGenerator
                        gen = BalanceGenerator(template_bal_path)
                        result_bytes = gen.generate(tb_df, dummy_map, scale_factor=scale_factor, tb_df_comp=tb_df_comp, periodo_actual_str=periodo_actual, periodo_comp_str=comp_period if comp_period != "Ninguno" else None, target_lang=target_lang)
                        col_actual = str(periodo_actual)
                        col_comp = str(comp_period) if comp_period != "Ninguno" else "Comp"
                        preview_df = self._extract_balance_preview(result_bytes, col_actual, col_comp)
                    else:
                        tb_df = TrialBalanceDB.get_trial_balance(empresa_activa, periodo_actual)
                        tb_comp = TrialBalanceDB.get_trial_balance(empresa_activa, comp_period) if comp_period != "Ninguno" else None
                        from src.reporting.balance_generator import BalanceGenerator
                        gen = BalanceGenerator(os.path.join(empresa_path, "Balance clasificado.xlsx"))
                        result_bytes = gen.generate(tb_df, map_balance_df, scale_factor=scale_factor, tb_df_comp=tb_comp, periodo_actual_str=periodo_actual, periodo_comp_str=comp_period if comp_period != "Ninguno" else None, target_lang=target_lang)
                        col_actual = str(periodo_actual)
                        col_comp = str(comp_period) if comp_period != "Ninguno" else "Comp"
                        preview_df = self._extract_balance_preview(result_bytes, col_actual, col_comp)
                        
                elif preview_df is None and base_code_clean == "#ER":
                    report_name = "ESTADO DE RESULTADOS CONSOLIDADO" if is_consolidated else "ESTADO DE RESULTADOS CLASIFICADOS"
                    if is_consolidated:
                        grupo_name = empresa_activa.replace("[GRUPO] ", "").strip()
                        db = SessionLocal()
                        try:
                            grupo_obj = db.query(ConsolidationGroup).filter_by(nombre_grupo=grupo_name).first()
                            grupo_id = grupo_obj.id
                        finally:
                            db.close()
                        df_hoja_act = get_consolidated_hoja(grupo_id, periodo_actual)
                        if df_hoja_act is None:
                            raise ValueError(f"No hay datos para consolidar el grupo '{grupo_name}' en el periodo {periodo_actual}. Verifica que la empresa matriz tenga cargado dicho periodo.")
                            
                        df_hoja_comp = None
                        if comp_period != "Ninguno":
                            df_hoja_comp = get_consolidated_hoja(grupo_id, comp_period)
                            
                        def clean_str(s):
                            if pd.isna(s): return ""
                            return str(s).strip().lower().replace('á','a').replace('é','e').replace('í','i').replace('ó','o').replace('ú','u')

                        pl_subtotals = {
                            "ganancia bruta", "resultado antes de impuestos",
                            "ganancias (perdida) del ejercicio", "ganancia (perdida) del ejercicio",
                            "estado de resultados", "otros rubros no clasificados"
                        }
                        pl_subtotals_normalized = {clean_str(name) for name in pl_subtotals}

                        # Slice PL
                        idx_er = df_hoja_act[df_hoja_act['Balance clasificado'] == "Estado de Resultados"].index
                        df_hoja_act_sec = df_hoja_act.loc[idx_er[0]+1:] if not idx_er.empty else df_hoja_act
                        df_hoja_act_clean = df_hoja_act_sec[df_hoja_act_sec['Balance clasificado'].notna() & (df_hoja_act_sec['Balance clasificado'].str.strip() != "")]
                        df_hoja_act_clean = df_hoja_act_clean[~df_hoja_act_clean['Balance clasificado'].apply(clean_str).isin(pl_subtotals_normalized)]
                        pl_dict_act = {row['Balance clasificado']: [row['CONSOLIDADO']] for _, row in df_hoja_act_clean.iterrows()}
                        pl_df_wide = pd.DataFrame(pl_dict_act)
                        
                        pl_df_comp_wide = None
                        if df_hoja_comp is not None:
                            idx_er_comp = df_hoja_comp[df_hoja_comp['Balance clasificado'] == "Estado de Resultados"].index
                            df_hoja_comp_sec = df_hoja_comp.loc[idx_er_comp[0]+1:] if not idx_er_comp.empty else df_hoja_comp
                            df_hoja_comp_clean = df_hoja_comp_sec[df_hoja_comp_sec['Balance clasificado'].notna() & (df_hoja_comp_sec['Balance clasificado'].str.strip() != "")]
                            df_hoja_comp_clean = df_hoja_comp_clean[~df_hoja_comp_clean['Balance clasificado'].apply(clean_str).isin(pl_subtotals_normalized)]
                            pl_dict_comp = {row['Balance clasificado']: [row['CONSOLIDADO']] for _, row in df_hoja_comp_clean.iterrows()}
                            pl_df_comp_wide = pd.DataFrame(pl_dict_comp)
                            
                        template_er_path = os.path.join(empresa_path, "Estado de Resultados Clasificados.xlsx")
                        if not os.path.exists(template_er_path):
                            template_er_path = os.path.join("templates", "Estado de Resultados Clasificados.xlsx")
                            
                        from src.reporting.er_generator import ERGenerator
                        gen = ERGenerator(template_er_path)
                        result_bytes, _ = gen.generate(pl_df_wide, scale_factor=scale_factor, pl_df_comp=pl_df_comp_wide, periodo_actual_str=periodo_actual, periodo_comp_str=comp_period if comp_period != "Ninguno" else None)
                        col_actual = str(periodo_actual)
                        col_comp = str(comp_period) if comp_period != "Ninguno" else "Comp"
                        preview_df = self._extract_er_preview(result_bytes, col_actual, col_comp)
                    else:
                        pl_df_wide = PlCuboDB.get_pl_cubo(empresa_activa, periodo_actual)
                        pl_df_comp_wide = PlCuboDB.get_pl_cubo(empresa_activa, comp_period) if comp_period != "Ninguno" else None
                        
                        from src.reporting.er_generator import ERGenerator
                        gen = ERGenerator(os.path.join(empresa_path, "Estado de Resultados Clasificados.xlsx"))
                        result_bytes, _ = gen.generate(pl_df_wide, scale_factor=scale_factor, pl_df_comp=pl_df_comp_wide, periodo_actual_str=periodo_actual, periodo_comp_str=comp_period if comp_period != "Ninguno" else None)
                        col_actual = str(periodo_actual)
                        col_comp = str(comp_period) if comp_period != "Ninguno" else "Comp"
                        preview_df = self._extract_er_preview(result_bytes, col_actual, col_comp)
                        
                elif preview_df is None and base_code_clean == "#ORI":
                    report_name = "ESTADO DE RESULTADOS INTEGRALES CONSOLIDADO" if is_consolidated else "ESTADO DE RESULTADOS INTEGRALES"
                    if is_consolidated:
                        grupo_name = empresa_activa.replace("[GRUPO] ", "").strip()
                        db = SessionLocal()
                        try:
                            grupo_obj = db.query(ConsolidationGroup).filter_by(nombre_grupo=grupo_name).first()
                            grupo_id = grupo_obj.id
                        finally:
                            db.close()
                        from src.core.consolidacion_engine import generar_hoja_trabajo
                        df_hoja_act, _ = generar_hoja_trabajo(grupo_id, periodo_actual)
                        df_hoja_comp = None
                        if comp_period != "Ninguno":
                            df_hoja_comp, _ = generar_hoja_trabajo(grupo_id, comp_period)
                            
                        # Slice PL
                        idx_er = df_hoja_act[df_hoja_act['Balance clasificado'] == "Estado de Resultados"].index
                        df_hoja_act_sec = df_hoja_act.loc[idx_er[0]+1:] if not idx_er.empty else df_hoja_act
                        df_hoja_act_clean = df_hoja_act_sec[df_hoja_act_sec['Balance clasificado'].notna() & (df_hoja_act_sec['Balance clasificado'].str.strip() != "")]
                        pl_dict_act = {row['Balance clasificado']: [row['CONSOLIDADO']] for _, row in df_hoja_act_clean.iterrows()}
                        pl_df_wide = pd.DataFrame(pl_dict_act)
                        
                        pl_df_comp_wide = None
                        if df_hoja_comp is not None:
                            idx_er_comp = df_hoja_comp[df_hoja_comp['Balance clasificado'] == "Estado de Resultados"].index
                            df_hoja_comp_sec = df_hoja_comp.loc[idx_er_comp[0]+1:] if not idx_er_comp.empty else df_hoja_comp
                            df_hoja_comp_clean = df_hoja_comp_sec[df_hoja_comp_sec['Balance clasificado'].notna() & (df_hoja_comp_sec['Balance clasificado'].str.strip() != "")]
                            pl_dict_comp = {row['Balance clasificado']: [row['CONSOLIDADO']] for _, row in df_hoja_comp_clean.iterrows()}
                            pl_df_comp_wide = pd.DataFrame(pl_dict_comp)
                            
                        template_er_path = os.path.join(empresa_path, "Estado de Resultados Clasificados.xlsx")
                        if not os.path.exists(template_er_path):
                            template_er_path = os.path.join("templates", "Estado de Resultados Clasificados.xlsx")
                            
                        from src.reporting.er_generator import ERGenerator
                        gen_er = ERGenerator(template_er_path)
                        _, preview_er = gen_er.generate(pl_df_wide, scale_factor=scale_factor, pl_df_comp=pl_df_comp_wide, periodo_actual_str=periodo_actual, periodo_comp_str=comp_period if comp_period != "Ninguno" else None, target_lang=target_lang)
                    else:
                        pl_df_wide = PlCuboDB.get_pl_cubo(empresa_activa, periodo_actual)
                        pl_df_comp_wide = PlCuboDB.get_pl_cubo(empresa_activa, comp_period) if comp_period != "Ninguno" else None
                        
                        from src.reporting.er_generator import ERGenerator
                        gen_er = ERGenerator(os.path.join(empresa_path, "Estado de Resultados Clasificados.xlsx"))
                        _, preview_er = gen_er.generate(pl_df_wide, scale_factor=scale_factor, pl_df_comp=pl_df_comp_wide, periodo_actual_str=periodo_actual, periodo_comp_str=comp_period if comp_period != "Ninguno" else None, target_lang=target_lang)
                    
                    from src.reporting.ori_generator import OriGenerator
                    template_ori_path = os.path.join(empresa_path, "Estado de Resultados Integrales.xlsx")
                    if not os.path.exists(template_ori_path):
                        template_ori_path = os.path.join("templates", "Estado de Resultados Integrales.xlsx")
                    if os.path.exists(template_ori_path):
                        gen_ori = OriGenerator(template_ori_path)
                        result_bytes = gen_ori.generate(
                            preview_er, 
                            periodo_actual_str=str(periodo_actual), 
                            periodo_comp_str=str(comp_period) if comp_period != "Ninguno" else None,
                            empresa=empresa_activa,
                            target_lang=target_lang
                        )
                        col_actual = str(periodo_actual)
                        col_comp = str(comp_period) if comp_period != "Ninguno" else "Comp"
                        preview_df = self._extract_ori_preview(result_bytes, col_actual, col_comp)
                    else:
                        preview_df = None
                        
                elif preview_df is None and base_code_clean == "#PATRIMONIO":
                    report_name = "ESTADO DE CAMBIOS EN EL PATRIMONIO CONSOLIDADO" if is_consolidated else "ESTADO DE CAMBIOS EN EL PATRIMONIO"
                    template_pat_path = os.path.join(empresa_path, "Estado de Cambios en el Patrimonio.xlsx")
                    if not os.path.exists(template_pat_path):
                        template_pat_path = os.path.join("templates", "Estado de Cambios en el Patrimonio.xlsx")
                    if os.path.exists(template_pat_path):
                        from src.reporting.patrimonio_generator import PatrimonioGenerator
                        pat_engine = PatrimonioGenerator(template_pat_path)
                        
                        bal_df_pat = None
                        pl_df_pat = None
                        pat_comp_for_open = comp_period
                        if str(comp_period).strip().lower() in ["ninguno", "none", ""]:
                            try:
                                pat_comp_for_open = f"{int(str(periodo_actual)[:4]) - 1}-12"
                            except Exception:
                                pat_comp_for_open = "Ninguno"

                        if is_consolidated:
                            grupo_name = empresa_activa.replace("[GRUPO] ", "").strip()
                            db = SessionLocal()
                            try:
                                grupo_obj = db.query(ConsolidationGroup).filter_by(nombre_grupo=grupo_name).first()
                                grupo_id = grupo_obj.id if grupo_obj else None
                            finally:
                                db.close()
                                
                            if grupo_id:
                                bal_df_pat, pl_df_pat, _, _ = self._get_consolidated_bal_and_pl(
                                    grupo_id=grupo_id,
                                    periodo_actual=periodo_actual,
                                    comp_period=pat_comp_for_open,
                                    scale_factor=scale_factor,
                                    empresa_path=empresa_path,
                                    grupo_name=grupo_name
                                )
                        else:
                            # Cargar Balance para empresa individual (incluyendo cierre año anterior para el Saldo Inicial al 01.01 del Bloque 1)
                            tb_df_p = TrialBalanceDB.get_trial_balance(empresa_activa, periodo_actual)
                            tb_comp_p = TrialBalanceDB.get_trial_balance(empresa_activa, pat_comp_for_open) if pat_comp_for_open != "Ninguno" else None
                            template_bal_p = os.path.join(empresa_path, "Balance clasificado.xlsx")
                            if not os.path.exists(template_bal_p):
                                template_bal_p = os.path.join("templates", "Balance clasificado.xlsx")
                            from src.reporting.balance_generator import BalanceGenerator
                            gen_b = BalanceGenerator(template_bal_p)
                            b_out = gen_b.generate(tb_df_p, map_balance_df, scale_factor=scale_factor, tb_df_comp=tb_comp_p, periodo_actual_str=periodo_actual, periodo_comp_str=pat_comp_for_open if pat_comp_for_open != "Ninguno" else None)
                            col_act_str = str(periodo_actual)
                            col_cmp_str = str(pat_comp_for_open) if pat_comp_for_open != "Ninguno" else "Comp"
                            bal_df_pat = self._extract_balance_preview(b_out, col_act_str, col_cmp_str)

                            # Cargar Estado de Resultados para empresa individual
                            pl_df_wide_p = PlCuboDB.get_pl_cubo(empresa_activa, periodo_actual)
                            pl_df_comp_wide_p = PlCuboDB.get_pl_cubo(empresa_activa, comp_period) if comp_period != "Ninguno" else None
                            template_er_p = os.path.join(empresa_path, "Estado de Resultados Clasificados.xlsx")
                            if not os.path.exists(template_er_p):
                                template_er_p = os.path.join("templates", "Estado de Resultados Clasificados.xlsx")
                            from src.reporting.er_generator import ERGenerator
                            gen_er = ERGenerator(template_er_p)
                            _, pl_df_pat = gen_er.generate(pl_df_wide_p, scale_factor=scale_factor, pl_df_comp=pl_df_comp_wide_p, periodo_actual_str=periodo_actual, periodo_comp_str=comp_period if comp_period != "Ninguno" else None)

                        if pat_comp_for_open != "Ninguno" and SnapshotVaultManager.has_active_snapshot(empresa_activa, pat_comp_for_open):
                            if bal_df_pat is not None:
                                bal_df_pat = SnapshotVaultManager.apply_closed_comp_to_statement(
                                    bal_df_pat, empresa_activa, "#BALANCE", periodo_actual, pat_comp_for_open, scale_factor=scale_factor
                                )
                        if comp_period != "Ninguno" and SnapshotVaultManager.has_active_snapshot(empresa_activa, comp_period):
                            if pl_df_pat is not None:
                                pl_df_pat = SnapshotVaultManager.apply_closed_comp_to_statement(
                                    pl_df_pat, empresa_activa, "#ER", periodo_actual, comp_period, scale_factor=scale_factor
                                )
                                
                        result_bytes = pat_engine.generate(
                            bal_preview_df=bal_df_pat,
                            pl_preview_df=pl_df_pat,
                            periodo_actual_str=str(periodo_actual),
                            periodo_comp_str=str(comp_period) if comp_period != "Ninguno" else None,
                            empresa=empresa_activa
                        )
                        result_bytes.seek(0)
                        from src.core.excel_utils import detect_patrimonio_skiprows, clean_patrimonio_preview_df
                        pat_skip = detect_patrimonio_skiprows(result_bytes)
                        result_bytes.seek(0)
                        preview_df = pd.read_excel(result_bytes, skiprows=pat_skip)
                        preview_df = clean_patrimonio_preview_df(preview_df)
                        result_bytes.seek(0)
                    else:
                        preview_df = None
                        
                elif preview_df is None and base_code_clean == "#EFE":
                    report_name = "ESTADO DE FLUJOS DE EFECTIVO CONSOLIDADO" if is_consolidated else "ESTADO DE FLUJOS DE EFECTIVO"
                    cf_filename = "Estado de Flujos de Efectivo.xlsx"
                    template_cf_path = os.path.join(empresa_path, cf_filename)
                    if not os.path.exists(template_cf_path):
                        import shutil
                        shutil.copy2(os.path.join("templates", cf_filename), template_cf_path)
                        
                    from src.reporting.cash_flow_generator import CashFlowGenerator
                    gen = CashFlowGenerator(template_cf_path)
                    
                    if is_consolidated:
                        from src.core.consolidacion_engine import generar_hoja_trabajo
                        db = SessionLocal()
                        try:
                            grupo_obj = db.query(ConsolidationGroup).filter_by(nombre_grupo=empresa_activa.replace("[GRUPO] ", "").strip()).first()
                            grupo_id = grupo_obj.id if grupo_obj else None
                        finally:
                            db.close()
                        df_hoja_act = get_consolidated_hoja(grupo_id, periodo_actual)
                        df_hoja_comp = None
                        if comp_period != "Ninguno":
                            df_hoja_comp = get_consolidated_hoja(grupo_id, comp_period)
                            
                        result_bytes, _ = gen.generate(
                            empresa=empresa_activa,
                            periodo_actual_str=periodo_actual,
                            periodo_comp_str=comp_period if comp_period != "Ninguno" else None,
                            map_balance_df=map_balance_df,
                            map_pl_df=map_pl_df,
                            method="Directo",
                            is_consolidado=True,
                            consolidated_hoja_trabajo_df=df_hoja_act,
                            consolidated_hoja_trabajo_comp_df=df_hoja_comp,
                            scale_factor=scale_factor,
                            target_lang=target_lang
                        )
                    else:
                        result_bytes, _ = gen.generate(
                            empresa=empresa_activa,
                            periodo_actual_str=periodo_actual,
                            periodo_comp_str=comp_period if comp_period != "Ninguno" else None,
                            map_balance_df=map_balance_df,
                            map_pl_df=map_pl_df,
                            method="Directo",
                            is_consolidado=False,
                            scale_factor=scale_factor,
                            target_lang=target_lang
                        )
                        
                    col_actual = str(periodo_actual)
                    col_comp = str(comp_period) if comp_period != "Ninguno" else "Comp"
                    preview_df = self._extract_cf_preview(result_bytes, col_actual, col_comp)
                    
                if preview_df is not None:
                    if was_loaded_from_snapshot and str(comp_period).strip().lower() in ["ninguno", "none", ""] and base_code_clean in ["#BALANCE", "#ESF", "#ER", "#ORI", "#EFE"]:
                        import numpy as np
                        preview_df = preview_df.copy()
                        if len(preview_df.columns) >= 3:
                            last_col = preview_df.columns[-1]
                            preview_df[last_col] = np.nan
                            if last_col != "Comp":
                                preview_df = preview_df.rename(columns={last_col: "Comp"})
                    elif not was_loaded_from_snapshot and comp_period != "Ninguno" and base_code_clean in ["#BALANCE", "#ESF", "#ER", "#ORI", "#EFE"] and SnapshotVaultManager.has_active_snapshot(empresa_activa, comp_period):
                        preview_df = SnapshotVaultManager.apply_closed_comp_to_statement(
                            preview_df, empresa_activa, base_code_clean, periodo_actual, comp_period, scale_factor=scale_factor
                        )
                    if base_code_clean in ["#BALANCE", "#ESF", "#ER", "#ORI", "#PATRIMONIO", "#EFE"]:
                        statements_cache[stmt_cache_key] = (preview_df, report_name)
                    is_bal = (base_code_clean == "#BALANCE" or "balance" in str(report_name).lower() or len(preview_df) >= 30)
                    
                    # Inyectar encabezado institucional superior si no es solo tabla y es un párrafo nuevo
                    if (not is_only_table) and item_type == 'paragraph':
                        # 1. Nombre de la Entidad
                        p_ent = doc.add_paragraph()
                        p_ent.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p_ent.paragraph_format.space_before = Pt(0)
                        p_ent.paragraph_format.space_after = Pt(0)
                        p_ent.paragraph_format.line_spacing = Pt(9.5) if is_bal else Pt(10.0)
                        r_ent = p_ent.add_run(empresa_activa.replace("[GRUPO] ", "").upper())
                        r_ent.font.name = 'Arial'
                        r_ent.font.size = Pt(9.5) if is_bal else Pt(10)
                        r_ent.font.bold = True
                        r_ent.font.color.rgb = RGBColor(31, 78, 120)
                        move_element_before(p, p_ent._p)

                        # 2. Título formal del Reporte
                        p_tit = doc.add_paragraph()
                        p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p_tit.paragraph_format.space_before = Pt(0)
                        p_tit.paragraph_format.space_after = Pt(0)
                        p_tit.paragraph_format.line_spacing = Pt(9.5) if is_bal else Pt(10.0)
                        r_tit = p_tit.add_run(report_name.upper())
                        r_tit.font.name = 'Arial'
                        r_tit.font.size = Pt(9.5) if is_bal else Pt(10)
                        r_tit.font.bold = True
                        move_element_before(p, p_tit._p)

                        # 3. Período y Unidad
                        col_dates = [str(c) for c in preview_df.columns if any(char.isdigit() for char in str(c))]
                        period_str = f"Al {col_dates[0]} y {col_dates[1]}" if len(col_dates) >= 2 else (f"Al {col_dates[0]}" if col_dates else "")
                        unit_label = "M$ (Miles de pesos)" if scale_factor == 1000.0 else "Ch$ (Pesos)"
                        
                        p_sub = doc.add_paragraph()
                        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p_sub.paragraph_format.space_before = Pt(0)
                        p_sub.paragraph_format.space_after = Pt(1) if is_bal else Pt(2)
                        p_sub.paragraph_format.line_spacing = Pt(8.5) if is_bal else Pt(9.5)
                        if period_str:
                            r_per = p_sub.add_run(f"{period_str}\n")
                            r_per.font.name = 'Arial'
                            r_per.font.size = Pt(7.5) if is_bal else Pt(8.5)
                            r_per.font.italic = True
                        r_unit = p_sub.add_run(f"(Expresado en {unit_label})")
                        r_unit.font.name = 'Arial'
                        r_unit.font.size = Pt(7.5) if is_bal else Pt(8)
                        r_unit.font.italic = True
                        r_unit.font.color.rgb = RGBColor(90, 90, 90)
                        move_element_before(p, p_sub._p)

                    # Inyectar o actualizar tabla
                    is_landscape = should_use_landscape(len(preview_df.columns), tag_code=tag_content)
                    if item_type == 'paragraph':
                        unit_str = "M$" if scale_factor == 1000.0 else "Ch$"
                        tbl = doc.add_table(rows=1, cols=len(preview_df.columns))
                        from src.reporting.word_export import populate_word_table_styled
                        ex_b = result_bytes.getvalue() if (result_bytes is not None and hasattr(result_bytes, 'getvalue')) else (result_bytes if isinstance(result_bytes, bytes) else None)
                        populate_word_table_styled(tbl, preview_df, excel_bytes=ex_b, unit=unit_str, compact=is_bal)
                        set_table_metadata(tbl, base_code_clean)
                        
                        if is_landscape:
                            sec_before = create_section_break_xml(is_landscape=False)
                            sec_after = create_section_break_xml(is_landscape=True)
                            move_element_before(p, sec_before)
                            move_element_before(p, tbl._tbl)
                            move_element_before(p, sec_after)
                        else:
                            move_element_before(p, tbl._tbl)
                        paragraphs_to_delete.append(p)

                        # Si justo después de este párrafo hay una tabla heredada sin metadata, removerla para evitar duplicación
                        nxt = p._p.getnext()
                        if nxt is not None and nxt.tag.endswith('p') and not (nxt.text and nxt.text.strip()):
                            nxt = nxt.getnext()
                        if nxt is not None and nxt.tag.endswith('tbl'):
                            try:
                                t_obj = docx.table.Table(nxt, doc)
                                if not get_table_metadata(t_obj):
                                    nxt.getparent().remove(nxt)
                            except Exception:
                                pass
                    else:
                        # item_type == 'table': Actualización in-place respetando anchos y estilos nativos de Word
                        # Si es #PATRIMONIO y la tabla Word previa no tenía el segundo bloque de encabezado, reconstruir estructura (SOLO en modo con comparativo)
                        pat_needs_rebuild = False
                        is_pat_no_comp = (not comp_period) or str(comp_period).strip().lower() in ["ninguno", "none", ""]
                        old_pat_matrix_chk = [[c.text.strip() for c in r_o.cells] for r_o in old_tbl.rows]
                        old_has_two_blocks = (_find_vertical_comp_block_start_texts(old_pat_matrix_chk) is not None)
                        if base_code_clean == "#PATRIMONIO" and not is_pat_no_comp and not old_has_two_blocks:
                            pat_needs_rebuild = True
                        if pat_needs_rebuild:
                            # Si se solicitó con comp_period == 'Ninguno', preservar los valores del bloque comparativo
                            # que ya existían en la tabla Word anterior por POSICIÓN ESTRUCTURAL de bloque (2° bloque)
                            if str(comp_period).strip().lower() in ["ninguno", "none", ""]:
                                old_tbl_matrix = [[c.text.strip() for c in r_old.cells] for r_old in old_tbl.rows]
                                old_comp_start = _find_vertical_comp_block_start_texts(old_tbl_matrix)
                                df_pat_matrix = [["" if pd.isna(v) else str(v).strip() for v in r_p] for _, r_p in preview_df.iterrows()]
                                comp_start_df = _find_vertical_comp_block_start_texts(df_pat_matrix)
                                if old_comp_start is not None and comp_start_df is not None:
                                    # Si old_tbl no tenía fila separadora vacía al inicio del bloque 2 pero preview_df sí la tiene, alinear tras la fila separadora
                                    if all(not t for t in df_pat_matrix[comp_start_df]) and any(t for t in old_tbl_matrix[old_comp_start]):
                                        comp_start_df += 1
                                    # Si preview_df tiene fila de segundo encabezado y old_tbl no la tenía, alinear en la primera fila de saldo inicial
                                    if comp_start_df < len(df_pat_matrix) and not df_pat_matrix[comp_start_df][0] and old_tbl_matrix[old_comp_start][0]:
                                        comp_start_df += 1
                                    old_comp_rows = old_tbl_matrix[old_comp_start:]
                                    for offset_c, old_vals in enumerate(old_comp_rows):
                                        target_df_idx = comp_start_df + offset_c
                                        if target_df_idx < len(preview_df):
                                            for col_c in range(min(len(old_vals), len(preview_df.columns))):
                                                if old_vals[col_c] != "":
                                                    preview_df.iloc[target_df_idx, col_c] = old_vals[col_c]
                            unit_str = "M$" if scale_factor == 1000.0 else "Ch$"
                            while len(old_tbl.rows) > 1:
                                old_tbl._tbl.remove(old_tbl.rows[-1]._tr)
                            from src.reporting.word_export import populate_word_table_styled
                            ex_b = result_bytes.getvalue() if (result_bytes is not None and hasattr(result_bytes, 'getvalue')) else (result_bytes if isinstance(result_bytes, bytes) else None)
                            populate_word_table_styled(old_tbl, preview_df, excel_bytes=ex_b, unit=unit_str, compact=is_bal)
                        else:
                            update_table_cells_inplace(old_tbl, preview_df, periodo_actual=periodo_actual, periodo_comp=comp_period)
                        set_table_metadata(old_tbl, base_code_clean)
                        
                        # Actualizar texto del subtítulo de período en el párrafo previo si existe
                        prev_p = old_tbl._tbl.getprevious()
                        if prev_p is not None and prev_p.tag.endswith('p'):
                            col_dates = [str(c) for c in preview_df.columns if any(char.isdigit() for char in str(c))]
                            is_no_comp = (not comp_period) or str(comp_period).strip().lower() in ["ninguno", "none", ""]
                            p_obj = docx.text.paragraph.Paragraph(prev_p, doc)
                            if "al 202" in p_obj.text.lower() or "expresado en" in p_obj.text.lower():
                                if is_no_comp and any(yr in p_obj.text for yr in ["2024", "2025", "2023"]):
                                    # Preservar el subtítulo del template que ya incluye el período anterior auditado
                                    pass
                                else:
                                    period_str = f"Al {col_dates[0]} y {col_dates[1]}" if len(col_dates) >= 2 else (f"Al {col_dates[0]}" if col_dates else "")
                                    unit_label = "M$ (Miles de pesos)" if scale_factor == 1000.0 else "Ch$ (Pesos)"
                                    p_obj.text = ""
                                    if period_str:
                                        r_per = p_obj.add_run(f"{period_str}\n")
                                        r_per.font.name = 'Arial'
                                        r_per.font.size = Pt(7.5) if is_bal else Pt(8.5)
                                        r_per.font.italic = True
                                    r_unit = p_obj.add_run(f"(Expresado en {unit_label})")
                                    r_unit.font.name = 'Arial'
                                    r_unit.font.size = Pt(7.5) if is_bal else Pt(8)
                                    r_unit.font.italic = True
                                    r_unit.font.color.rgb = RGBColor(90, 90, 90)
                                    p_obj.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                # --- CASO NOTAS ---
                else:
                    base_code = code
                    table_idx = None
                    if "." in code:
                        parts_dot = code.split(".", 1)
                        base_code = parts_dot[0].strip()
                        try:
                            table_idx = int(parts_dot[1].strip()) - 1
                        except ValueError:
                            pass

                    from src.reporting.notes import NOTE_REGISTRY, get_full_note_registry
                    full_registry = get_full_note_registry()
                    if base_code in full_registry:
                        info = full_registry[base_code]
                        target_sheets = info['sheets']
                        
                        if target_sheets:
                            if engine_notes is not None:
                                category = info.get('category', 'resultados')
                                is_balance_note = category in ["activos_corrientes", "activos_no_corrientes", "pasivos_corrientes", "pasivos_no_corrientes", "patrimonio"]
                                note_comp = override_period if override_period else (periodo_comp_balance if is_balance_note else periodo_comp_resultados)
                                
                                note_cache_key = (base_code, note_comp)
                                if note_cache_key in note_elements_cache:
                                    elements = note_elements_cache[note_cache_key]
                                else:
                                    from src.core.snapshot_vault_manager import SnapshotVaultManager
                                    snap_note_els = None
                                    if SnapshotVaultManager.has_active_snapshot(empresa_activa, periodo_actual):
                                        snap_note_els = SnapshotVaultManager.get_note_elements(
                                            empresa=empresa_activa,
                                            periodo_actual=periodo_actual,
                                            note_code=base_code,
                                            scale_factor=scale_factor,
                                            comp_period=note_comp
                                        )
                                    if snap_note_els:
                                        raw_els = snap_note_els
                                        note_elements_cache[note_cache_key] = raw_els
                                        elements = raw_els
                                    else:
                                        entity_contexts = get_entity_contexts(note_comp)
                                        excel_nota_out = engine_notes.generate(
                                            sheet_names=target_sheets,
                                            entity_contexts=entity_contexts,
                                            active_entity_name=empresa_activa,
                                            is_consolidated=is_consolidated,
                                            scale_factor=scale_factor,
                                            periodo_actual_str=periodo_actual,
                                            periodo_comp_str=note_comp,
                                            map_balance_df=map_balance_df,
                                            map_pl_df=map_pl_df,
                                            target_lang=target_lang,
                                            evaluate_formulas=False
                                        )
                                        
                                        from src.core.external_notes_manager import ExternalNotesManager
                                        excel_nota_out = ExternalNotesManager.inject_external_data(
                                            excel_nota_out,
                                            empresa=empresa_activa,
                                            periodo_actual=periodo_actual,
                                            periodo_comp=note_comp if note_comp != "Ninguno" else None,
                                            target_lang=target_lang
                                        )
                                        
                                        # Evaluar las fórmulas y reconciliar redondeos (<= 4 M$) contra EEFF antes de procesar
                                        _, excel_nota_out = RoundingReconciler.reconcile_and_evaluate_note_workbook(
                                            excel_nota_out,
                                            empresa=empresa_activa,
                                            periodo_actual=periodo_actual,
                                            scale_factor=scale_factor,
                                            target_sheets=target_sheets,
                                        )
                                        
                                        # Intentar extraer usando rangos nombrados
                                        from src.core.excel_utils import extract_named_ranges_from_excel
                                        named_ranges = extract_named_ranges_from_excel(excel_nota_out, base_code)
                                        
                                        if named_ranges:
                                            raw_els = []
                                            for r_name, df_chunk in named_ranges:
                                                raw_els.append(("table", df_chunk))
                                        else:
                                            raw_els = []
                                            from src.ui_pages.informes_y_notas import split_sheet_into_elements
                                            excel_nota_out.seek(0)
                                            xl_nota_f = pd.ExcelFile(excel_nota_out)
                                            for s_name in target_sheets:
                                                if s_name in xl_nota_f.sheet_names:
                                                    preview_nota_df = xl_nota_f.parse(sheet_name=s_name, header=None)
                                                    raw_els.extend(split_sheet_into_elements(preview_nota_df, sheet_name=s_name))
                                        note_elements_cache[note_cache_key] = raw_els
                                        elements = raw_els
                                
                                # Si se solicitó un código o índice de tabla específico (ej: #N04.1, #N14.1)
                                if table_idx is not None:
                                    matched_explicit = [
                                        el for el in elements
                                        if el[0] == "table" and len(el) > 3 and el[3] and str(el[3]).strip().upper() == code.strip().upper()
                                    ]
                                    if matched_explicit:
                                        target_elements = [matched_explicit[0]]
                                    else:
                                        tables = [el for el in elements if el[0] == "table" and is_valid_note_table(clean_table_df(el[1]))]
                                        if 0 <= table_idx < len(tables):
                                            target_elements = [tables[table_idx]]
                                        else:
                                            target_elements = [] # Fuera de rango o sin tablas
                                else:
                                    target_elements = elements
                                
                                if item_type == 'paragraph':
                                    # Inyectar elementos uno a uno antes del párrafo p
                                    table_counter = 0
                                    for item in target_elements:
                                        el_type = item[0]
                                        el_val = item[1]
                                        if el_type == "text":
                                            new_p = doc.add_paragraph()
                                            run = new_p.add_run(str(el_val))
                                            run.font.bold = True
                                            run.font.name = 'Arial'
                                            run.font.size = Pt(9)
                                            move_element_before(p, new_p._p)
                                        else:
                                            chunk_df = clean_table_df(el_val)
                                            if not is_valid_note_table(chunk_df):
                                                continue
                                            
                                            table_counter += 1
                                            note_tag = item[3] if (len(item) > 3 and item[3]) else (code if table_idx is not None else f"{base_code}.{table_counter}")
                                            is_landscape = should_use_landscape(len(chunk_df.columns), tag_code=tag_content)
                                            
                                            tbl = doc.add_table(rows=len(chunk_df), cols=len(chunk_df.columns))
                                            populate_notes_docx_table(tbl, chunk_df)
                                            set_table_metadata(tbl, note_tag)
                                            
                                            if is_landscape:
                                                sec_before = create_section_break_xml(is_landscape=False)
                                                sec_after = create_section_break_xml(is_landscape=True)
                                                move_element_before(p, sec_before)
                                                move_element_before(p, tbl._tbl)
                                                move_element_before(p, sec_after)
                                            else:
                                                move_element_before(p, tbl._tbl)
                                            
                                    paragraphs_to_delete.append(p)
                                else:
                                    # item_type == 'table': Actualización in-place o reconstrucción estructural si cambió el formato de la tabla
                                    tbl_elements = [el for el in target_elements if el[0] == "table"]
                                    t_df = tbl_elements[0][1] if tbl_elements else None
                                        
                                    if t_df is not None:
                                        chunk_df = clean_table_df(t_df)
                                        if is_valid_note_table(chunk_df):
                                            r0_cells_txt = [c.text.strip().replace("\n", " ") for c in old_tbl.rows[0].cells] if old_tbl.rows else []
                                            has_bogus_index_hdr = bool(r0_cells_txt) and all(
                                                re.match(r"^\d+(\s*(M\$|Th\$|Ch\$))?$", t, re.IGNORECASE) or t == ""
                                                for t in r0_cells_txt
                                            ) and any(
                                                re.match(r"^\d+(\s*(M\$|Th\$|Ch\$))?$", t, re.IGNORECASE)
                                                for t in r0_cells_txt
                                            )
                                            has_corrupted_rows = has_bogus_index_hdr or any(
                                                "transacciones significativas" in r.cells[0].text.lower()
                                                for r in old_tbl.rows if len(r.cells) > 0
                                            )
                                            is_note_no_comp = (not note_comp) or str(note_comp).strip().lower() in ["ninguno", "none", ""]
                                            needs_rebuild = has_corrupted_rows or (not is_note_no_comp and len(old_tbl.columns) != len(chunk_df.columns))
                                            if needs_rebuild:
                                                # Si se emitió con 'Ninguno' para preservar el comparativo pero la tabla cambió de formato/columnas,
                                                # cargar el comparativo IFRS por defecto para que las nuevas columnas comparativas no queden vacías
                                                if note_comp == "Ninguno":
                                                    try:
                                                        parts_p = str(periodo_actual).split('-')
                                                        if len(parts_p) >= 2:
                                                            prev_yr_str = str(int(parts_p[0]) - 1)
                                                            fallback_comp = f"{prev_yr_str}-12" if is_balance_note else f"{prev_yr_str}-{parts_p[1]}"
                                                            fb_cache_key = (base_code, fallback_comp)
                                                            if fb_cache_key not in note_elements_cache:
                                                                fb_contexts = get_entity_contexts(fallback_comp)
                                                                fb_excel = engine_notes.generate(
                                                                    sheet_names=target_sheets,
                                                                    entity_contexts=fb_contexts,
                                                                    active_entity_name=empresa_activa,
                                                                    is_consolidated=is_consolidated,
                                                                    scale_factor=scale_factor,
                                                                    periodo_actual_str=periodo_actual,
                                                                    periodo_comp_str=fallback_comp,
                                                                    map_balance_df=map_balance_df,
                                                                    map_pl_df=map_pl_df,
                                                                    target_lang=target_lang,
                                                                    evaluate_formulas=False
                                                                )
                                                                from src.core.external_notes_manager import ExternalNotesManager
                                                                fb_excel = ExternalNotesManager.inject_external_data(
                                                                    fb_excel,
                                                                    empresa=empresa_activa,
                                                                    periodo_actual=periodo_actual,
                                                                    periodo_comp=fallback_comp,
                                                                    target_lang=target_lang
                                                                )
                                                                from src.ui_pages.informes_y_notas import evaluate_formulas_in_workbook, split_sheet_into_elements
                                                                fb_excel = evaluate_formulas_in_workbook(fb_excel)
                                                                fb_els = []
                                                                for s_name in target_sheets:
                                                                    fb_excel.seek(0)
                                                                    fb_df_s = pd.read_excel(fb_excel, sheet_name=s_name, header=None)
                                                                    fb_els.extend(split_sheet_into_elements(fb_df_s, sheet_name=s_name))
                                                                note_elements_cache[fb_cache_key] = fb_els
                                                            fb_elements = note_elements_cache[fb_cache_key]
                                                            fb_matched = [
                                                                el for el in fb_elements
                                                                if el[0] == "table" and len(el) > 3 and el[3] and str(el[3]).strip().upper() == code.strip().upper()
                                                            ]
                                                            if not fb_matched and table_idx is not None:
                                                                fb_tbls = [el for el in fb_elements if el[0] == "table" and is_valid_note_table(clean_table_df(el[1]))]
                                                                if 0 <= table_idx < len(fb_tbls):
                                                                    fb_matched = [fb_tbls[table_idx]]
                                                            if fb_matched:
                                                                fb_chunk_df = clean_table_df(fb_matched[0][1])
                                                                if is_valid_note_table(fb_chunk_df):
                                                                    chunk_df = fb_chunk_df
                                                    except Exception:
                                                        pass

                                                new_tbl = doc.add_table(rows=len(chunk_df), cols=len(chunk_df.columns))
                                                populate_notes_docx_table(new_tbl, chunk_df)
                                                set_table_metadata(new_tbl, tag_content)
                                                old_tbl._tbl.addprevious(new_tbl._tbl)
                                                parent_elm = old_tbl._tbl.getparent()
                                                if parent_elm is not None:
                                                    parent_elm.remove(old_tbl._tbl)
                                            else:
                                                update_table_cells_inplace(old_tbl, chunk_df, periodo_actual=periodo_actual, periodo_comp=note_comp)
                                                set_table_metadata(old_tbl, tag_content)
                                    elif old_tbl is not None and old_tbl.rows:
                                        r0_first = (old_tbl.rows[0].cells[0].text or "").strip().lower() if old_tbl.rows[0].cells else ""
                                        if any(k in r0_first for k in ("cambios en pasivo", "vencimiento de pasivos")):
                                            duplicate_tables_to_remove.append((old_tbl, code))
            except Exception as tag_err:
                print(f"Error procesando {tag_content}: {tag_err}")
                if item_type == 'paragraph' and p is not None:
                    new_p = doc.add_paragraph()
                    r_err = new_p.add_run(f"[{tag_content}: No se pudo inyectar automáticamente - {str(tag_err)}]")
                    r_err.font.italic = True
                    r_err.font.size = Pt(8.5)
                    move_element_before(p, new_p._p)
                    paragraphs_to_delete.append(p)
                            
        # Eliminar tablas duplicadas/huérfanas (ej. segundo cuadro #N16.3 cuando #N16.3 ya incluye actual y comparativo)
        for dup_tbl, dup_code in duplicate_tables_to_remove:
            try:
                tbl_el = dup_tbl._tbl
                parent = tbl_el.getparent()
                if parent is not None:
                    prev_el = tbl_el.getprevious()
                    steps_checked = 0
                    while prev_el is not None and steps_checked < 6:
                        steps_checked += 1
                        prev_prev = prev_el.getprevious()
                        if prev_el.tag.endswith('p'):
                            p_txt = (prev_el.text or "").strip()
                            p_low = p_txt.lower()
                            if not p_txt or (
                                "cambios en pasivos por derechos de uso al" in p_low and
                                not p_low.startswith("b)") and
                                not p_low.startswith("a)")
                            ):
                                parent.remove(prev_el)
                                if p_txt:
                                    break
                            else:
                                break
                        else:
                            break
                        prev_el = prev_prev
                    parent.remove(tbl_el)
            except Exception:
                pass

        # Eliminar párrafos marcadores procesados
        for p in paragraphs_to_delete:
            try:
                parent = p._p.getparent()
                if parent is not None:
                    parent.remove(p._p)
            except Exception:
                pass
                
        # Guardar en buffer
        output_buffer = BytesIO()
        doc.save(output_buffer)
        output_buffer.seek(0)
        return output_buffer
