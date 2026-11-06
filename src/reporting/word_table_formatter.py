import re
import numpy as np
import pandas as pd
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

def clean_table_df(df):
    """
    Limpia un DataFrame antes de convertirlo a tabla de Word:
    - Elimina filas completamente vacías (todas NaN o vacías).
    - Elimina columnas completamente vacías (todas NaN o vacías), evitando columnas fantasma.
    """
    if df is None or df.empty:
        return pd.DataFrame()
        
    # Copia para no mutar el original
    cleaned = df.copy()
    
    # Reemplazar cadenas vacías o espacios en blanco por NaN para el filtrado
    cleaned = cleaned.replace(r'^\s*$', np.nan, regex=True)
    
    # Eliminar filas completamente NaN
    cleaned = cleaned.dropna(how='all', axis=0)
    
    # Eliminar columnas completamente NaN
    cleaned = cleaned.dropna(how='all', axis=1)

    # Función auxiliar para formatear fechas tipo YYYY-MM-DD o YYYY-MM-DD 00:00:00 a DD.MM.YYYY
    def format_date_str(val):
        if pd.isna(val):
            return val
        s = str(val).strip()
        m = re.match(r'^(\d{4})-(\d{2})-(\d{2})(?:\s+\d{2}:\d{2}:\d{2})?$', s)
        if m:
            return f"{m.group(3)}.{m.group(2)}.{m.group(1)}"
        return val

    # Formatear nombres de columnas si son fechas con horas
    cleaned.columns = [format_date_str(col) for col in cleaned.columns]
    
    # Formatear celdas que sean fechas con horas
    if hasattr(cleaned, 'map'):
        cleaned = cleaned.map(format_date_str)
    else:
        cleaned = cleaned.applymap(format_date_str)

    cleaned = cleaned.reset_index(drop=True)
    from src.reporting.formatting import ensure_subtable_separator_rows
    return ensure_subtable_separator_rows(cleaned)

def is_valid_note_table(chunk_df):
    """
    Valida si un chunk de nota es una tabla financiera real con datos.
    Filtra tablas huérfanas de 1 sola fila que solo contienen encabezados o fechas vacías.
    """
    if chunk_df is None or chunk_df.empty:
        return False
        
    cleaned = clean_table_df(chunk_df)
    if len(cleaned) <= 1:
        # Una tabla de 1 fila no tiene datos (es un encabezado huérfano)
        return False
        
    # Verificar si existe al menos un valor numérico contable en la tabla
    has_numeric = False
    for col in cleaned.columns:
        for val in cleaned[col]:
            if pd.notna(val):
                if isinstance(val, (int, float)):
                    # Descartar años solitarios (ej: 2024, 2025, 2026) que son parte de encabezados
                    if not (2000 <= val <= 2100):
                        has_numeric = True
                        break
                elif isinstance(val, str):
                    s = val.strip()
                    # Limpiar signos contables
                    s_clean = re.sub(r'[\$\%\(\)\s\.\,]|(M\$)|(MM\$)|(CLP)|(UF)|(USD)', '', s)
                    if s_clean.lstrip('-+').isdigit() and len(s_clean) > 0:
                        try:
                            n = float(s_clean)
                            if not (2000 <= n <= 2100):
                                has_numeric = True
                                break
                        except ValueError:
                            pass
        if has_numeric:
            break
            
    return has_numeric

def classify_columns(df):
    """
    Clasifica dinámicamente cada columna de un DataFrame financiero en:
    - 'note': columna de número de nota de referencia (corta, números 1-35 o vacía).
    - 'numeric': cifras contables, montos, porcentajes, fechas numéricas.
    - 'text': descripciones, conceptos, monedas, bancos, instituciones, glosas.
    """
    classifications = []
    num_cols = len(df.columns)
    
    for c_idx in range(num_cols):
        col_name = str(df.columns[c_idx]).strip().lower()
        
        # 1. Chequeo por nombre de cabecera obvio
        if col_name in ['nota', 'notas', 'n°', 'n° nota', 'ref', 'ref.']:
            classifications.append('note')
            continue
            
        # 2. Recolectar valores no nulos
        vals = []
        for val in df.iloc[:, c_idx]:
            if pd.notna(val):
                s = str(val).strip()
                if s not in ["", "None", "nan"]:
                    vals.append(val)
                    
        if not vals:
            # Columna vacía -> tratar como texto
            classifications.append('text')
            continue
            
        # Comprobar si los valores son números de nota (ej: enteros pequeños 1-35 o '4.1')
        is_all_note_like = False
        if c_idx == 1 and num_cols >= 3:
            note_matches = 0
            for v in vals:
                sv = str(v).strip()
                if re.match(r'^\d{1,2}(\.\d{1,2})?$', sv) or sv in ["-", "—"]:
                    note_matches += 1
            if len(vals) > 0 and (note_matches / len(vals)) >= 0.7:
                try:
                    nums = [float(v) for v in vals if str(v).replace('.','').isdigit()]
                    if nums and max(nums) <= 50:
                        is_all_note_like = True
                except:
                    pass
        if is_all_note_like:
            classifications.append('note')
            continue

        # Comprobar si es columna numérica contable
        numeric_count = 0
        for v in vals:
            if isinstance(v, (int, float)):
                numeric_count += 1
            elif isinstance(v, str):
                sv = v.strip()
                if sv in ["-", "—", "–", "0"]:
                    numeric_count += 1
                else:
                    cleaned = re.sub(r'[\$\%\(\)\s\.\,]|(M\$)|(MM\$)|(CLP)|(UF)|(USD)', '', sv)
                    if cleaned.lstrip('-+').isdigit() and len(cleaned) > 0:
                        numeric_count += 1
                        
        ratio = numeric_count / len(vals) if vals else 0.0
        is_date_col = any(k in col_name for k in ["31.12", "30.06", "31.03", "30.09", "202", "201", "actual", "anterior", "comparativo"])
        
        if ratio >= 0.40 or (is_date_col and ratio >= 0.20):
            classifications.append('numeric')
        else:
            classifications.append('text')
            
    return classifications

def compute_optimal_column_widths(df, classifications, total_width_inches=6.85):
    """
    Calcula y distribuye el ancho de cada columna para llenar el 100% del ancho imprimible
    de extremo a extremo sin dejar espacios vacíos a la derecha.
    - Asegura que las columnas de texto tengan suficiente espacio para no apretar el texto.
    - Asegura que las columnas numéricas tengan suficiente espacio para cifras largas y paréntesis.
    """
    num_cols = len(df.columns)
    if num_cols <= 0:
        return []
    if num_cols == 1:
        return [total_width_inches]
        
    widths = [0.0] * num_cols
    
    # 1. Ancho para columnas tipo 'note'
    note_indices = [i for i, c in enumerate(classifications) if c == 'note']
    for idx in note_indices:
        widths[idx] = 0.60
        
    # 2. Ancho para columnas tipo 'numeric'
    numeric_indices = [i for i, c in enumerate(classifications) if c == 'numeric']
    for idx in numeric_indices:
        max_len = 0
        for val in df.iloc[:, idx]:
            if pd.notna(val):
                s = str(val).strip()
                if s not in ["", "nan", "None"]:
                    max_len = max(max_len, len(s))
        
        col_hdr = str(df.columns[idx]).strip()
        max_len = max(max_len, len(col_hdr))
        
        # Ajustar ancho base por largo de caracteres
        if num_cols >= 10:
            w = 0.65 if max_len < 10 else 0.80
        elif num_cols >= 7:
            w = 0.85 if max_len < 10 else 1.05
        else:
            if max_len >= 13:
                w = 1.35
            elif max_len >= 9:
                w = 1.20
            elif max_len >= 6:
                w = 1.05
            else:
                w = 0.95
            
        widths[idx] = w

    # 3. Ancho para columnas tipo 'text'
    text_indices = [i for i, c in enumerate(classifications) if c == 'text']
    used_width = sum(widths)
    remaining_width = total_width_inches - used_width
    
    if text_indices:
        if len(text_indices) == 1:
            # Una sola columna de texto: toma TODO el espacio restante hasta el margen derecho
            widths[text_indices[0]] = max(1.5, remaining_width)
        else:
            # Múltiples columnas de texto: repartir asegurando que cada columna
            # tenga suficiente espacio para sus palabras más largas sin apretarlas
            text_lens = []
            for idx in text_indices:
                lens = []
                for val in df.iloc[:, idx]:
                    if pd.notna(val) and str(val).strip() not in ["", "nan", "None"]:
                        lens.append(len(str(val).strip()))
                avg_len = np.percentile(lens, 75) if lens else len(str(df.columns[idx]))
                text_lens.append(max(avg_len, 8))
                
            total_lens = sum(text_lens)
            for i, idx in enumerate(text_indices):
                prop = text_lens[i] / total_lens
                # Ponderación suave para no castigar columnas secundarias de texto
                soft_prop = 0.5 * prop + 0.5 * (1.0 / len(text_indices))
                widths[idx] = max(1.2, remaining_width * soft_prop)
    else:
        # Si no hay columnas de texto, repartir el exceso entre las numéricas
        if numeric_indices:
            diff = total_width_inches - sum(widths)
            extra = diff / len(numeric_indices)
            for idx in numeric_indices:
                widths[idx] += extra

    # 4. Normalización estricta: La suma debe ser EXACTAMENTE total_width_inches
    # para usar de extremo a extremo la hoja sin dejar espacios vacíos
    curr_sum = sum(widths)
    if curr_sum > 0:
        factor = total_width_inches / curr_sum
        widths = [round(w * factor, 3) for w in widths]
        # Ajustar redondeo mínimo en la columna más ancha
        residual = round(total_width_inches - sum(widths), 3)
        max_idx = widths.index(max(widths))
        widths[max_idx] = round(widths[max_idx] + residual, 3)
        
    return widths

def set_cell_nowrap(cell):
    """Evita que Word quiebre líneas dentro de la celda (cifras, fechas, signos negativos)."""
    tcPr = cell._tc.get_or_add_tcPr()
    if not any(child.tag.endswith('noWrap') for child in tcPr):
        tcPr.append(parse_xml(f'<w:noWrap {nsdecls("w")}/>'))

def set_table_width(table, width_inches):
    """Fija el ancho total exacto de la tabla de Word al 100% del área imprimible."""
    tblPr = table._tbl.tblPr
    for child in list(tblPr):
        if child.tag.endswith('tblW'):
            tblPr.remove(child)
    width_dxa = int(round(width_inches * 1440))
    tblW = parse_xml(f'<w:tblW {nsdecls("w")} w:w="{width_dxa}" w:type="dxa"/>')
    tblPr.append(tblW)

def apply_table_formatting_and_widths(table, df, classifications=None, total_width_inches=6.85, is_compact=False, is_landscape=False):
    """
    Aplica anchos óptimos, cantSplit, tblHeader, noWrap y tblW al 100% de la página.
    """
    num_cols = len(df.columns)
    if classifications is None:
        classifications = classify_columns(df)
        
    col_widths = compute_optimal_column_widths(df, classifications, total_width_inches=total_width_inches)
    
    table.autofit = False
    set_table_width(table, total_width_inches)
    
    for row_idx, row in enumerate(table.rows):
        trPr = row._tr.get_or_add_trPr()
        # Evitar que la fila se parta entre dos páginas
        if not any(child.tag.endswith('cantSplit') for child in trPr):
            trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
            
        # Repetir cabecera en cada página si la tabla salta de página
        if row_idx == 0:
            if not any(child.tag.endswith('tblHeader') for child in trPr):
                trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
                
        for j, cell in enumerate(row.cells):
            if j < len(col_widths):
                cell.width = Inches(col_widths[j])
                tcPr = cell._tc.get_or_add_tcPr()
                for child in list(tcPr):
                    if child.tag.endswith('tcW'):
                        tcPr.remove(child)
                cell_dxa = int(round(col_widths[j] * 1440))
                tcPr.append(parse_xml(f'<w:tcW {nsdecls("w")} w:w="{cell_dxa}" w:type="dxa"/>'))
                
                # Cifras contables y notas nunca deben quebrar líneas
                if j < len(classifications) and classifications[j] in ['numeric', 'note']:
                    set_cell_nowrap(cell)

def get_dynamic_font_and_spacing(num_cols, is_compact=False, is_landscape=False):
    """
    Calcula dinámicamente el tamaño de letra, interlínea y espaciado de celdas
    según el número de columnas y la orientación para un ajuste perfecto sin interlíneas gruesas.
    """
    if is_compact:
        # Balance Clasificado compacto para entrar en 1 página
        return 7.0, 7.5, Pt(8.0), 12, 35
        
    if is_landscape:
        if num_cols >= 12:
            return 6.5, 7.0, Pt(8.0), 12, 18
        elif num_cols >= 8:
            return 7.0, 7.5, Pt(8.5), 16, 25
        else:
            return 7.5, 8.0, Pt(9.0), 20, 30

    # Portrait estándar
    if num_cols <= 4:
        return 8.0, 8.5, Pt(9.5), 25, 40
    elif num_cols <= 6:
        return 7.5, 8.0, Pt(9.0), 20, 30
    elif num_cols <= 8:
        return 7.0, 7.5, Pt(8.5), 16, 22
    else:
        return 6.5, 7.0, Pt(8.0), 12, 18

def should_use_landscape(num_cols, tag_code=""):
    """
    Determina si una tabla debe generarse o inyectarse en orientación horizontal (Landscape).
    Se activa si:
    - La etiqueta contiene explícitamente ':HORIZONTAL' o ':H'
    - La tabla tiene 7 o más columnas (ej: cuadro de movimientos de activo fijo #N09.2, pasivos #N15, patrimonio)
    """
    tag_upper = str(tag_code).upper()
    if ":HORIZONTAL" in tag_upper or ":H" in tag_upper:
        return True
    if num_cols >= 7:
        return True
    return False

def create_section_break_xml(is_landscape=False):
    """
    Crea un elemento XML de salto de sección de Word con los márgenes y dimensiones exactos.
    """
    if is_landscape:
        # 11 x 8.5 pulgadas, márgenes 0.7"
        return parse_xml(r'''
<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:pPr>
    <w:sectPr>
      <w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>
      <w:pgMar w:top="1008" w:right="1008" w:bottom="1008" w:left="1008" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:pPr>
</w:p>
''')
    else:
        # 8.5 x 11 pulgadas, márgenes 0.8"
        return parse_xml(r'''
<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:pPr>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1152" w:right="1152" w:bottom="1152" w:left="1152" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:pPr>
</w:p>
''')
