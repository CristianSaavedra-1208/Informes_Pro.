import re
import datetime
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from io import BytesIO
import pandas as pd
import openpyxl


def set_cell_bg(cell, fill_color):
    """Inyecta XML OXML para dar color de fondo a una celda de Word."""
    if not fill_color or fill_color.lower() in ['none', 'transparent']:
        return
    clean_color = str(fill_color).replace('#', '').strip()
    if len(clean_color) == 8:
        clean_color = clean_color[2:]
    shading_elm = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), clean_color))
    cell._tc.get_or_add_tcPr().append(shading_elm)


def set_cell_margins(cell, top=35, bottom=35, left=50, right=50):
    """Ajusta los márgenes internos (padding) de la celda de Word en dxa (1 pt = 20 dxa)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)


def set_cell_borders(cell, top=None, bottom=None):
    """Inyecta bordes OXML a una celda de Word (estilo contable idéntico a pantalla)."""
    tcPr = cell._tc.get_or_add_tcPr()
    borders_xml = f'<w:tcBorders {nsdecls("w")}>'
    if top and isinstance(top, dict) and top.get('val') not in ('none', None):
        borders_xml += f'<w:top w:val="{top.get("val", "single")}" w:sz="{top.get("sz", "6")}" w:space="0" w:color="{top.get("color", "000000")}"/>'
    else:
        borders_xml += '<w:top w:val="none"/>'
        
    if bottom and isinstance(bottom, dict) and bottom.get('val') not in ('none', None):
        borders_xml += f'<w:bottom w:val="{bottom.get("val", "single")}" w:sz="{bottom.get("sz", "6")}" w:space="0" w:color="{bottom.get("color", "000000")}"/>'
    else:
        borders_xml += '<w:bottom w:val="none"/>'
        
    borders_xml += '<w:left w:val="none"/><w:right w:val="none"/>'
    borders_xml += '</w:tcBorders>'
    tcPr.append(parse_xml(borders_xml))


def convert_border_to_docx(border_side):
    """Convierte un borde de openpyxl al formato de diccionario docx OXML."""
    if not border_side or border_side.style is None or str(border_side.style).lower() in ['none', '']:
        return None
    style = border_side.style
    color_hex = '000000'
    if border_side.color and border_side.color.rgb:
        c_rgb = str(border_side.color.rgb)
        if len(c_rgb) == 8 and c_rgb != '00000000' and not c_rgb.startswith('000000'):
            color_hex = c_rgb[2:]
        elif len(c_rgb) == 6:
            color_hex = c_rgb

    if style == 'double':
        return {'val': 'double', 'sz': '12', 'color': color_hex}
    elif style in ('medium', 'mediumDashed', 'mediumDashDot', 'mediumDashDotDot', 'thick'):
        return {'val': 'single', 'sz': '12', 'color': color_hex}
    elif style in ('dashed', 'dashDot', 'dashDotDot'):
        return {'val': 'dashed', 'sz': '6', 'color': color_hex}
    elif style == 'dotted':
        return {'val': 'dotted', 'sz': '6', 'color': color_hex}
    else:
        return {'val': 'single', 'sz': '6', 'color': color_hex}


def apply_table_column_widths(table, num_cols, total_width_inches=7.1):
    """Calcula y aplica anchos de columna óptimos para informes financieros."""
    table.autofit = False
    
    if num_cols <= 1:
        col_widths = [total_width_inches]
    elif num_cols == 2:
        col_widths = [total_width_inches * 0.70, total_width_inches * 0.30]
    elif num_cols == 3:
        col_widths = [total_width_inches * 0.64, total_width_inches * 0.18, total_width_inches * 0.18]
    elif num_cols == 4:
        col_widths = [total_width_inches * 0.56, total_width_inches * 0.10, total_width_inches * 0.17, total_width_inches * 0.17]
    elif num_cols == 5:
        col_widths = [total_width_inches * 0.46, total_width_inches * 0.10, total_width_inches * 0.14, total_width_inches * 0.15, total_width_inches * 0.15]
    else:
        col0 = max(3.0, total_width_inches - (0.9 * (num_cols - 1)))
        rem = max(0.7, (total_width_inches - col0) / (num_cols - 1))
        col_widths = [col0] + [rem] * (num_cols - 1)

    for row in table.rows:
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        for j, cell in enumerate(row.cells):
            if j < len(col_widths):
                cell.width = Inches(col_widths[j])

    if len(table.rows) > 0:
        hdr_trPr = table.rows[0]._tr.get_or_add_trPr()
        hdr_trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))


def format_accounting_val(val, is_category_header=False, row_name=""):
    """Formatea valores numéricos contables idéntico a la plantilla Excel y pantalla."""
    if pd.isna(val) or val is None or str(val).strip().lower() in ["", "nan", "none"]:
        return ""
        
    is_per_share = "por acción" in str(row_name).lower() or "por accion" in str(row_name).lower()
    
    # Si es número de punto flotante o entero
    if isinstance(val, (int, float)):
        if is_per_share or (isinstance(val, float) and abs(val) < 10.0 and val != 0 and not val.is_integer()):
            if val < 0:
                return f"({abs(val):,.2f})".replace(",", "X").replace(".", ",").replace("X", ".")
            else:
                return f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        if val == 0:
            return "0,00" if is_per_share else ("0" if is_category_header else "-")
        elif val < 0:
            return f"({abs(val):,.0f})".replace(",", ".")
        else:
            return f"{val:,.0f}".replace(",", ".")
            
    try:
        num = float(str(val).replace(",", ""))
        if pd.isna(num):
            return ""
        if is_per_share or (abs(num) < 10.0 and num != 0 and not num.is_integer()):
            if num < 0:
                return f"({abs(num):,.2f})".replace(",", "X").replace(".", ",").replace("X", ".")
            else:
                return f"{num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        if num == 0:
            return "0,00" if is_per_share else ("0" if is_category_header else "-")
        elif num < 0:
            return f"({abs(num):,.0f})".replace(",", ".")
        else:
            return f"{num:,.0f}".replace(",", ".")
    except:
        if "00:00:00" in str(val):
            return str(val).split(" ")[0]
        return str(val).strip()


def clean_str(s):
    if s is None:
        return ""
    return str(s).replace("\xa0", " ").strip().lower()


def populate_word_table_styled(table, df, excel_bytes=None, sheet_name=None, unit=None, is_miles=None, col_indices=None, compact=False):
    """
    Rellena una tabla de Word aplicando exactamente el mismo formato que la pantalla:
    - Cabecera blanca con borde superior e inferior negro (2px).
    - Extracción fiel de estilos de la plantilla Excel (openpyxl) si excel_bytes está presente:
      negrita, bordes, alineación y colores celda a celda.
    - Fallback contable limpio idéntico al estilo corporativo de pantalla.
    - Si compact=True (Balance Clasificado), reduce padding e interlíneas para calzar en 1 sola página.
    """
    table.style = 'Normal Table'
    df = df.copy()
    df_cols = list(df.columns)
    num_cols = len(df_cols)

    # Configuración de espaciados y tamaños según modo
    pad_top = 15 if compact else 30
    pad_bottom = 15 if compact else 30
    pad_lr = 40 if compact else 50
    p_spacing = Pt(8.0) if compact else Pt(10.0)
    base_font_size = 7.0 if compact else 8.0
    hdr_font_size = 7.5 if compact else 8.5
    hdr_pad_v = 12 if compact else 20
    hdr_line_spacing = Pt(8.0) if compact else Pt(9.5)

    # Detectar si la unidad es en miles
    if is_miles is None:
        if unit is not None:
            is_miles = any(m in str(unit).lower() for m in ["miles", "m$", "mch$"])
        else:
            is_miles = False

    # 1. CABECERA DE TABLA (Idéntica a pantalla: fondo blanco, texto negro negrita, bordes negros arriba y abajo)
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(df_cols):
        header_text = str(col_name)
        if is_miles and (any(char.isdigit() for char in header_text) or (i > 0 and header_text.lower().strip() not in ["nota", "notas", "concepto", "cuenta", "item", "rubro"])):
            if "M$" not in header_text:
                header_text = f"{header_text}\nM$"
        hdr_cells[i].text = header_text
        set_cell_bg(hdr_cells[i], "FFFFFF")
        set_cell_margins(hdr_cells[i], top=hdr_pad_v, bottom=hdr_pad_v, left=pad_lr, right=pad_lr)
        # Bordes contables de cabecera: 2px (sz="12") superior e inferior
        set_cell_borders(hdr_cells[i], top={'val': 'single', 'sz': '12', 'color': '000000'}, bottom={'val': 'single', 'sz': '12', 'color': '000000'})
        
        for paragraph in hdr_cells[i].paragraphs:
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
            paragraph.paragraph_format.line_spacing = hdr_line_spacing
            col_lower = str(col_name).lower().strip()
            if i == 0:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            elif col_lower in ["nota", "notas"]:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                
            for run in paragraph.runs:
                run.font.bold = True
                run.font.color.rgb = RGBColor(0, 0, 0)
                run.font.name = 'Arial'
                run.font.size = Pt(hdr_font_size)

    # Identificar columnas numéricas vs texto
    non_numeric_names = {'balance clasificado', 'nota', 'notas', 'descripcion', 'concepto', 'cuenta_id', 'glosa'}
    numeric_df_cols = [
        j for j in range(1, num_cols)
        if str(df_cols[j]).lower().strip() not in non_numeric_names
    ]

    # 2. PROCESAR ESTILOS DINÁMICOS DESDE EXCEL (Si se suministran excel_bytes)
    dynamic_styled = False
    row_styles_data = {} # idx -> list of cell style dicts

    if excel_bytes is not None:
        try:
            if isinstance(excel_bytes, BytesIO):
                excel_bytes.seek(0)
                wb = openpyxl.load_workbook(excel_bytes, data_only=True)
            else:
                wb = openpyxl.load_workbook(BytesIO(excel_bytes), data_only=True)
                
            if sheet_name and sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
            else:
                ws = wb.active

            ws_cols_mapping = {}
            if col_indices:
                ws_cols_mapping = {j: col_indices[j] for j in range(num_cols)}
            else:
                excel_name_col = 1
                for col in range(1, 11):
                    for row in range(1, 20):
                        val = ws.cell(row=row, column=col).value
                        if val and str(val).strip().lower() in ["concepto", "descripcion", "detalle", "flujos", "origen/aplicacion", "balance clasificado", "clasificacion", "activos"]:
                            excel_name_col = col
                            break
                    if excel_name_col > 1:
                        break
                        
                if excel_name_col == 1:
                    found = False
                    for col in range(1, 11):
                        for row in range(1, 25):
                            val = ws.cell(row=row, column=col).value
                            if val and isinstance(val, str) and len(val.strip()) > 3:
                                if not re.match(r'^\d+$', val.strip()):
                                    excel_name_col = col
                                    found = True
                                    break
                        if found:
                            break

                excel_date_cols = []
                for col in range(1, ws.max_column + 1):
                    for row in range(1, 15):
                        val = ws.cell(row=row, column=col).value
                        if val is not None:
                            is_date = (
                                isinstance(val, (datetime.datetime, datetime.date)) or
                                (isinstance(val, str) and re.search(r'20\d{2}', val))
                            )
                            if is_date:
                                excel_date_cols.append(col)
                                break
                excel_date_cols = sorted(list(set(excel_date_cols)))

                excel_nota_col = None
                for col in range(1, ws.max_column + 1):
                    for row in range(1, 15):
                        val = ws.cell(row=row, column=col).value
                        if val and str(val).strip().lower() == "nota":
                            excel_nota_col = col
                            break
                if excel_nota_col is None and excel_date_cols:
                    first_date_col = excel_date_cols[0]
                    if first_date_col > excel_name_col + 1:
                        excel_nota_col = excel_name_col + 1

                non_numeric_df_cols = [0]
                for j in range(1, num_cols):
                    col_name_lower = str(df_cols[j]).strip().lower()
                    if col_name_lower in ['nota', 'notas']:
                        non_numeric_df_cols.append(j)

                excel_non_numeric_cols = [excel_name_col]
                if excel_nota_col is not None:
                    excel_non_numeric_cols.append(excel_nota_col)
                
                curr = excel_name_col + 1
                while len(excel_non_numeric_cols) < len(non_numeric_df_cols):
                    if curr not in excel_non_numeric_cols:
                        excel_non_numeric_cols.append(curr)
                    curr += 1
                
                for idx_m, j in enumerate(non_numeric_df_cols):
                    ws_cols_mapping[j] = excel_non_numeric_cols[idx_m]

                excel_numeric_cols = list(excel_date_cols)
                curr = (excel_date_cols[-1] + 1) if excel_date_cols else (excel_name_col + 1)
                while len(excel_numeric_cols) < len(numeric_df_cols):
                    excel_numeric_cols.append(curr)
                    curr += 1
                
                for idx_m, j in enumerate(numeric_df_cols):
                    ws_cols_mapping[j] = excel_numeric_cols[idx_m]

            # Construir mapa de filas de Excel
            row_map = {}
            for r in range(1, ws.max_row + 1):
                for col_idx in range(1, min(ws.max_column + 1, 6)):
                    val = ws.cell(row=r, column=col_idx).value
                    if val is not None:
                        s_val = clean_str(val)
                        if s_val and s_val not in row_map:
                            row_map[s_val] = r

            row_top_borders = {}
            row_bottom_borders = {}

            for idx, df_row in df.iterrows():
                search_val = clean_str(df_row.iloc[0])
                found_row = row_map.get(search_val)
                cell_styles = []

                if found_row is not None:
                    r_top_b = None
                    r_bottom_b = None

                    for j in range(num_cols):
                        excel_col = ws_cols_mapping.get(j, j + 1)
                        cell = ws.cell(row=found_row, column=excel_col)
                        
                        bold = False
                        italic = False
                        font_size = base_font_size
                        text_color = RGBColor(0, 0, 0)
                        
                        if cell.font:
                            bold = bool(cell.font.bold)
                            italic = bool(cell.font.italic)
                            if cell.font.size:
                                font_size = min(max(float(cell.font.size) - (2.5 if compact else 2.0), 6.5 if compact else 7.5), 8.5 if compact else 9.5)
                            if cell.font.color and cell.font.color.rgb:
                                c_rgb = str(cell.font.color.rgb)
                                if len(c_rgb) == 8 and c_rgb != '00000000' and not c_rgb.startswith('000000'):
                                    try:
                                        r_c = int(c_rgb[2:4], 16)
                                        g_c = int(c_rgb[4:6], 16)
                                        b_c = int(c_rgb[6:8], 16)
                                        text_color = RGBColor(r_c, g_c, b_c)
                                    except:
                                        pass

                        align = WD_ALIGN_PARAGRAPH.RIGHT if (j in numeric_df_cols or j > 0) else WD_ALIGN_PARAGRAPH.LEFT
                        if str(df_cols[j]).lower().strip() in ['nota', 'notas']:
                            align = WD_ALIGN_PARAGRAPH.CENTER
                        if cell.alignment and cell.alignment.horizontal:
                            if cell.alignment.horizontal == 'center':
                                align = WD_ALIGN_PARAGRAPH.CENTER
                            elif cell.alignment.horizontal == 'left':
                                align = WD_ALIGN_PARAGRAPH.LEFT
                            elif cell.alignment.horizontal == 'right':
                                align = WD_ALIGN_PARAGRAPH.RIGHT

                        fill_bg = None
                        if cell.fill and cell.fill.fill_type == 'solid' and cell.fill.fgColor:
                            c_rgb = str(cell.fill.fgColor.rgb)
                            if len(c_rgb) == 8 and c_rgb != '00000000' and not c_rgb.startswith('000000'):
                                fill_bg = c_rgb[2:]
                            elif len(c_rgb) == 6:
                                fill_bg = c_rgb

                        top_b = convert_border_to_docx(cell.border.top) if cell.border else None
                        bottom_b = convert_border_to_docx(cell.border.bottom) if cell.border else None

                        if top_b and not r_top_b:
                            r_top_b = top_b
                        if bottom_b and not r_bottom_b:
                            r_bottom_b = bottom_b

                        cell_styles.append({
                            'bold': bold,
                            'italic': italic,
                            'font_size': font_size,
                            'text_color': text_color,
                            'align': align,
                            'fill_bg': fill_bg,
                            'top_b': top_b,
                            'bottom_b': bottom_b
                        })

                    if r_top_b:
                        row_top_borders[idx] = r_top_b
                    if r_bottom_b:
                        row_bottom_borders[idx] = r_bottom_b
                else:
                    for j in range(num_cols):
                        align = WD_ALIGN_PARAGRAPH.RIGHT if (j in numeric_df_cols or j > 0) else WD_ALIGN_PARAGRAPH.LEFT
                        if str(df_cols[j]).lower().strip() in ['nota', 'notas']:
                            align = WD_ALIGN_PARAGRAPH.CENTER
                        cell_styles.append({
                            'bold': False,
                            'italic': False,
                            'font_size': base_font_size,
                            'text_color': RGBColor(0, 0, 0),
                            'align': align,
                            'fill_bg': None,
                            'top_b': None,
                            'bottom_b': None
                        })

                row_styles_data[idx] = cell_styles

            # Propagación horizontal de bordes de totales/subtotales
            for idx in row_styles_data:
                r_top_b = row_top_borders.get(idx)
                r_bottom_b = row_bottom_borders.get(idx)
                if r_top_b or r_bottom_b:
                    for j in range(num_cols):
                        if r_top_b and not row_styles_data[idx][j]['top_b']:
                            row_styles_data[idx][j]['top_b'] = r_top_b
                        if r_bottom_b and not row_styles_data[idx][j]['bottom_b']:
                            row_styles_data[idx][j]['bottom_b'] = r_bottom_b

            dynamic_styled = True
        except Exception as ex:
            import sys
            print(f"Error cargando estilos de Excel en Word: {ex}", file=sys.stderr)
            dynamic_styled = False

    # 3. CONSTRUCCIÓN DE FILAS EN WORD
    for row_idx, (index, row) in enumerate(df.iterrows()):
        if row_idx + 1 < len(table.rows):
            row_cells = table.rows[row_idx + 1].cells
        else:
            row_cells = table.add_row().cells
        first_cell_str = str(row.iloc[0]).strip()
        first_cell_clean = first_cell_str.lower()

        # Determinar si la fila no tiene datos (categoría)
        has_data_values = any(
            isinstance(row.iloc[c], (int, float)) or 
            (pd.notna(row.iloc[c]) and str(row.iloc[c]).strip() not in ["", "-", "nan", "None"] and any(ch.isdigit() for ch in str(row.iloc[c])))
            for c in range(1, num_cols)
        )
        is_category_header = not has_data_values

        for j, col_name in enumerate(df_cols):
            val = row[col_name]
            text_val = format_accounting_val(val, is_category_header=is_category_header, row_name=first_cell_str)
            
            cell = row_cells[j]
            cell.text = text_val
            set_cell_margins(cell, top=pad_top, bottom=pad_bottom, left=pad_lr, right=pad_lr)

            # Extraer formato si es dinámico desde Excel
            if dynamic_styled and index in row_styles_data and j < len(row_styles_data[index]):
                st_info = row_styles_data[index][j]
                bold = st_info['bold']
                italic = st_info['italic']
                font_size = st_info['font_size']
                text_color = st_info['text_color']
                align = st_info['align']
                fill_bg = st_info['fill_bg']
                top_b = st_info['top_b']
                bottom_b = st_info['bottom_b']
                
                # Cebrado suave idéntico a pantalla si no hay color específico
                if not fill_bg:
                    fill_bg = "E3F0FE" if (index % 2 == 1) else "FFFFFF"
                    
                set_cell_bg(cell, fill_bg)
                set_cell_borders(cell, top=top_b, bottom=bottom_b)
            else:
                # Fallback idéntico a las reglas de pantalla
                is_even = (index % 2 == 1)
                fill_bg = "E3F0FE" if is_even else "FFFFFF"
                set_cell_bg(cell, fill_bg)

                is_grand_total = not is_category_header and any(k in first_cell_clean for k in [
                    "total activos", "total patrimonio y pasivos", "patrimonio total", 
                    "ganancia (pérdida) del ejercicio", "resultado del ejercicio", "saldo final de efectivo"
                ])
                is_subtotal = not is_category_header and not is_grand_total and any(k in first_cell_clean for k in [
                    "total", "sub total", "subtotal", "ganancia bruta"
                ])

                top_b = None
                bottom_b = None
                if is_grand_total:
                    top_b = {'val': 'single', 'sz': '6', 'color': '000000'}
                    bottom_b = {'val': 'double', 'sz': '12', 'color': '000000'}
                elif is_subtotal:
                    top_b = {'val': 'single', 'sz': '6', 'color': '000000'}
                    bottom_b = {'val': 'single', 'sz': '6', 'color': '000000'}
                
                set_cell_borders(cell, top=top_b, bottom=bottom_b)
                
                bold = is_grand_total or is_subtotal or is_category_header
                italic = False
                font_size = base_font_size
                text_color = RGBColor(0, 0, 0)
                if str(col_name).lower().strip() in ['nota', 'notas']:
                    align = WD_ALIGN_PARAGRAPH.CENTER
                elif j in numeric_df_cols or j > 0:
                    align = WD_ALIGN_PARAGRAPH.RIGHT
                else:
                    align = WD_ALIGN_PARAGRAPH.LEFT

            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.space_before = Pt(0) if compact else Pt(0.5)
                paragraph.paragraph_format.space_after = Pt(0) if compact else Pt(0.5)
                paragraph.paragraph_format.line_spacing = p_spacing
                paragraph.alignment = align
                for run in paragraph.runs:
                    run.font.name = 'Arial'
                    run.font.size = Pt(font_size)
                    run.font.bold = bold
                    run.font.italic = italic
                    run.font.color.rgb = text_color

    apply_table_column_widths(table, num_cols, total_width_inches=7.1)


class WordExportEngine:
    @staticmethod
    def generate_classified_balance_word(df, title="Estado de Situación Financiera Clasificado", unit="Ch$", entity_name="DB TERRA CHILE HOLDCO SPA AND SUBSIDIARIES", excel_bytes=None, sheet_name=None, *args, **kwargs):
        """
        Genera un informe en Word de Balance Clasificado con formato financiero idéntico a pantalla
        calibrado para entrar en 1 sola página.
        """
        doc = Document()
        
        # Márgenes compactos de 0.35 pulgadas para garantizar ajuste a 1 sola página
        sections = doc.sections
        for section in sections:
            section.top_margin = Inches(0.35)
            section.bottom_margin = Inches(0.35)
            section.left_margin = Inches(0.5)
            section.right_margin = Inches(0.5)

        # Encabezado superior compacto
        p_ent = doc.add_paragraph()
        p_ent.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_ent.paragraph_format.space_before = Pt(0)
        p_ent.paragraph_format.space_after = Pt(0)
        p_ent.paragraph_format.line_spacing = Pt(9.5)
        r_ent = p_ent.add_run(entity_name.upper())
        r_ent.font.name = 'Arial'
        r_ent.font.size = Pt(9.5)
        r_ent.font.bold = True
        r_ent.font.color.rgb = RGBColor(31, 78, 120)

        p_tit = doc.add_paragraph()
        p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_tit.paragraph_format.space_before = Pt(0)
        p_tit.paragraph_format.space_after = Pt(0)
        p_tit.paragraph_format.line_spacing = Pt(9.5)
        r_tit = p_tit.add_run(title.upper())
        r_tit.font.name = 'Arial'
        r_tit.font.size = Pt(9.5)
        r_tit.font.bold = True

        # Períodos y Unidad
        col_dates = [str(c) for c in df.columns if any(char.isdigit() for char in str(c))]
        period_str = f"Al {col_dates[0]} y {col_dates[1]}" if len(col_dates) >= 2 else (f"Al {col_dates[0]}" if col_dates else "")
        
        p_sub = doc.add_paragraph()
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_sub.paragraph_format.space_before = Pt(0)
        p_sub.paragraph_format.space_after = Pt(1)
        p_sub.paragraph_format.line_spacing = Pt(8.5)
        if period_str:
            r_per = p_sub.add_run(f"{period_str}\n")
            r_per.font.name = 'Arial'
            r_per.font.size = Pt(7.5)
            r_per.font.italic = True
        r_unit = p_sub.add_run(f"(Expresado en {unit})")
        r_unit.font.name = 'Arial'
        r_unit.font.size = Pt(7.5)
        r_unit.font.italic = True
        r_unit.font.color.rgb = RGBColor(90, 90, 90)

        num_cols = len(df.columns)
        table = doc.add_table(rows=1, cols=num_cols)
        # Modo compacto activo para que las 38-42 líneas calcen en 1 sola página
        populate_word_table_styled(table, df, excel_bytes=excel_bytes, sheet_name=sheet_name, unit=unit, compact=True)

        # Pie inferior
        is_cons = kwargs.get('is_consolidado', False) or ("[grupo]" in str(entity_name).lower()) or ("consolidado" in str(title).lower())
        if is_cons:
            note_text = "Las notas adjuntas N°s 1 a XX forman parte integral de estos estados financieros consolidados."
        else:
            note_text = "Las notas adjuntas N°s 1 a XX forman parte integral de estos estados financieros."

        p_footer = doc.add_paragraph()
        p_footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_footer.paragraph_format.space_before = Pt(3)
        p_footer.paragraph_format.space_after = Pt(0)
        p_footer.paragraph_format.line_spacing = Pt(8.5)
        r_foot = p_footer.add_run(note_text)
        r_foot.font.name = 'Arial'
        r_foot.font.size = Pt(7.5)
        r_foot.font.italic = True
        r_foot.font.color.rgb = RGBColor(80, 80, 80)

        output = BytesIO()
        doc.save(output)
        output.seek(0)
        return output

    @staticmethod
    def generate_notes_word(elements, title="Nota", unit="M$", note_code=None, excel_bytes=None, *args, **kwargs):
        """
        Genera documento Word de Notas con formato contable idéntico a pantalla.
        """
        doc = Document()
        
        for section in doc.sections:
            section.top_margin = Inches(0.8)
            section.bottom_margin = Inches(0.8)
            section.left_margin = Inches(0.85)
            section.right_margin = Inches(0.85)
        
        heading = doc.add_heading(title, level=1)
        heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in heading.runs:
            r.font.name = 'Arial'
            r.font.size = Pt(15)
            r.font.color.rgb = RGBColor(31, 78, 120)
            r.font.bold = True
            
        subtitle = doc.add_paragraph(f"Expresado en {unit}")
        subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in subtitle.runs:
            r.font.name = 'Arial'
            r.font.size = Pt(9)
            r.font.italic = True
            
        doc.add_paragraph()
        
        table_counter = 0
        for item in elements:
            if len(item) == 3:
                el_type, el_val, sh_name = item
            else:
                el_type, el_val = item[0], item[1]
                sh_name = None

            if el_type == "text":
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(4)
                p.paragraph_format.space_after = Pt(2)
                run = p.add_run(str(el_val))
                run.font.bold = True
                run.font.name = 'Arial'
                run.font.size = Pt(9.5)
            else:
                chunk_df = el_val.dropna(how='all', axis=0).reset_index(drop=True)
                if chunk_df.empty:
                    continue
                    
                table_counter += 1
                sub_code = f"{note_code}.{table_counter}" if note_code else f"Cuadro {table_counter}"
                
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(6)
                p.paragraph_format.space_after = Pt(2)
                run = p.add_run(f"[{sub_code}]")
                run.font.bold = True
                run.font.name = 'Arial'
                run.font.size = Pt(9.5)
                run.font.color.rgb = RGBColor(31, 78, 120)
                
                num_cols = len(chunk_df.columns)
                table = doc.add_table(rows=1, cols=num_cols)
                populate_word_table_styled(table, chunk_df, excel_bytes=excel_bytes, sheet_name=sh_name, unit=unit)
                doc.add_paragraph()
                
        output = BytesIO()
        doc.save(output)
        output.seek(0)
        return output


def generate_word_report(df, title="Reporte Financiero", subtitle="Expresado en pesos", entity_name="DB TERRA CHILE HOLDCO SPA AND SUBSIDIARIES", excel_bytes=None, sheet_name=None, unit=None, is_miles=None, *args, **kwargs):
    """
    Exporta cualquier estado financiero (ER, Flujo, Patrimonio, ORI, etc.) a Word con formato idéntico a pantalla.
    """
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.6)
        section.right_margin = Inches(0.6)

    # Encabezado superior
    p_ent = doc.add_paragraph()
    p_ent.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ent.paragraph_format.space_before = Pt(0)
    p_ent.paragraph_format.space_after = Pt(1)
    r_ent = p_ent.add_run(entity_name.upper())
    r_ent.font.name = 'Arial'
    r_ent.font.size = Pt(10)
    r_ent.font.bold = True
    r_ent.font.color.rgb = RGBColor(31, 78, 120)

    p_tit = doc.add_paragraph()
    p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_tit.paragraph_format.space_before = Pt(1)
    p_tit.paragraph_format.space_after = Pt(1)
    r_tit = p_tit.add_run(title.upper())
    r_tit.font.name = 'Arial'
    r_tit.font.size = Pt(10)
    r_tit.font.bold = True

    sub = doc.add_paragraph(f"({subtitle})")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub.paragraph_format.space_before = Pt(1)
    sub.paragraph_format.space_after = Pt(2)
    for r in sub.runs:
        r.font.name = 'Arial'
        r.font.size = Pt(8)
        r.font.italic = True
        r.font.color.rgb = RGBColor(80, 80, 80)
        
    num_cols = len(df.columns)
    table = doc.add_table(rows=1, cols=num_cols)
    populate_word_table_styled(table, df, excel_bytes=excel_bytes, sheet_name=sheet_name, unit=unit or subtitle, is_miles=is_miles)

    # Pie inferior
    p_footer_space = doc.add_paragraph()
    p_footer_space.paragraph_format.space_before = Pt(10)
    p_footer_space.paragraph_format.space_after = Pt(0)
    p_footer_space.paragraph_format.line_spacing = 1.0

    is_cons = kwargs.get('is_consolidado', False) or ("[grupo]" in str(entity_name).lower()) or ("consolidado" in str(title).lower()) or ("consolidada" in str(title).lower())
    if is_cons:
        note_text = "Las notas adjuntas N°s 1 a XX forman parte integral de estos estados financieros consolidados."
    else:
        note_text = "Las notas adjuntas N°s 1 a XX forman parte integral de estos estados financieros."

    p_footer = doc.add_paragraph()
    p_footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_footer.paragraph_format.space_before = Pt(0)
    p_footer.paragraph_format.space_after = Pt(0)
    r_foot = p_footer.add_run(note_text)
    r_foot.font.name = 'Arial'
    r_foot.font.size = Pt(8)
    r_foot.font.italic = True
    r_foot.font.color.rgb = RGBColor(80, 80, 80)

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output.getvalue()
