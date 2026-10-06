import pandas as pd

def apply_corporate_style(df, excel_bytes=None, sheet_name=None, col_indices=None, unit=None, is_miles=None, target_lang=None):
    import openpyxl
    from io import BytesIO

    if target_lang is None:
        try:
            import streamlit as st
            target_lang = st.session_state.get('idioma_reporte', 'es')
        except Exception:
            target_lang = 'es'

    df = df.copy()

    if str(target_lang).lower() == 'en':
        from src.core.ifrs_glossary import translate_ifrs_term
        new_cols = []
        for col in df.columns:
            new_cols.append(translate_ifrs_term(str(col), target_lang='en'))
        df.columns = new_cols

        if len(df) > 0:
            df.iloc[:, 0] = df.iloc[:, 0].apply(lambda x: translate_ifrs_term(str(x), target_lang='en') if pd.notna(x) else x)

    # Detectar si la unidad es en miles
    if is_miles is None:
        if unit is not None:
            is_miles = any(m in str(unit).lower() for m in ["miles", "m$", "mch$"])
        else:
            try:
                import streamlit as st
                um_candidates = [
                    st.session_state.get('unidad_medida'),
                    st.session_state.get('unidad'),
                    st.session_state.get('um_er'),
                    st.session_state.get('um_cf'),
                    st.session_state.get('um_bal'),
                    st.session_state.get('um_pat'),
                    st.session_state.get('um_ori'),
                ]
                for um in um_candidates:
                    if um and any(m in str(um).lower() for m in ["miles", "m$", "mch$"]):
                        is_miles = True
                        break
                if is_miles is None:
                    scale = st.session_state.get('scale_factor', 1.0)
                    if scale == 1000.0 or scale == 1000:
                        is_miles = True
            except Exception:
                is_miles = False

    if is_miles:
        new_cols = []
        for i, col in enumerate(df.columns):
            col_str = str(col)
            # Agregar \nM$ a columnas de fecha/período o numéricas (no primera col y no Nota)
            if any(char.isdigit() for char in col_str) or (i > 0 and col_str.lower().strip() not in ["nota", "notas", "concepto", "cuenta", "item", "rubro"]):
                if "M$" not in col_str:
                    new_cols.append(f"{col_str}\nM$")
                else:
                    new_cols.append(col_str)
            else:
                new_cols.append(col_str)
        df.columns = new_cols

    if len(df.columns) > 0:
        df.iloc[:, 0] = df.iloc[:, 0].fillna("").astype(str).replace({"nan": "", "None": ""})

    for i in range(1, len(df.columns)):
        col_series = df.iloc[:, i]
        if getattr(col_series, 'dtype', None) == 'object':
            try:
                converted = pd.to_numeric(col_series, errors='ignore')
                df.iloc[:, i] = converted
            except Exception:
                pass

    # Definir columnas de valores (numéricas) de forma robusta
    non_numeric_names = {'balance clasificado', 'nota', 'notas', 'descripcion', 'concepto', 'cuenta_id', 'glosa'}
    numeric_cols = [
        c for c in df.columns[1:]
        if str(c).lower().strip() not in non_numeric_names
    ]
    text_cols = [c for c in df.columns if c not in numeric_cols]
    
    is_en = (str(target_lang).lower() == 'en')

    def format_accounting(x):
        if pd.isna(x) or str(x).lower() == 'nan':
            return ""
        
        # If it's already a number
        if isinstance(x, (int, float)):
            if x == 0: return "-"
            if is_en:
                if x < 0: return f"({abs(x):,.0f})"
                return f"{x:,.0f}"
            else:
                if x < 0: return f"({abs(x):,.0f})".replace(',', '.')
                return f"{x:,.0f}".replace(',', '.')
            
        # Try to cast string numbers
        try:
            num = float(str(x).replace(',', ''))
            if pd.isna(num): return ""
            if num == 0: return "-"
            if is_en:
                if num < 0: return f"({abs(num):,.0f})"
                return f"{num:,.0f}"
            else:
                if num < 0: return f"({abs(num):,.0f})".replace(',', '.')
                return f"{num:,.0f}".replace(',', '.')
        except:
            # If it's text, date, or "M$"
            if "00:00:00" in str(x):
                return str(x).split(" ")[0] # Clean up dates
            return str(x)
        
    format_dict = {
        col: format_accounting
        for col in df.columns[1:]
    }
    
    styler = df.style.format(format_dict)
    
    table_styles = [
        {
            'selector': '',
            'props': [
                ('min-width', '100% !important'),
                ('width', 'auto !important'),
                ('table-layout', 'fixed !important'),
                ('border-collapse', 'collapse !important'),
                ('margin-left', '0px !important'),
                ('margin-right', 'auto !important')
            ]
        },
        {
            'selector': 'th',
            'props': [
                ('background-color', '#FFFFFF'),
                ('color', '#000000'),
                ('font-weight', 'bold'),
                ('font-family', 'Inter, sans-serif'),
                ('border', 'none'),
                ('padding', '8px'),
                ('border-top', '2px solid #000000'),
                ('border-bottom', '2px solid #000000'),
                ('text-align', 'left')
            ]
        },
        {
            'selector': 'td',
            'props': [
                ('font-family', 'Inter, sans-serif'),
                ('font-size', '0.85em'),
                ('border-left', 'none'),
                ('border-right', 'none'),
                ('border-top', 'none'),
                ('border-bottom', '1px solid #f0f0f0'),
                ('padding', '6px 10px'),
                ('background-clip', 'padding-box')
            ]
        },
        {
            'selector': 'tr:nth-child(even) td',
            'props': [('background-color', '#e3f0fe')]
        },
        {
            'selector': 'tr:nth-child(odd) td',
            'props': [('background-color', '#FFFFFF')]
        }
    ]
    
    styler = styler.set_table_styles(table_styles)
    styler = styler.hide(axis="index")
    
    def _check_is_num(v):
        if pd.isna(v) or v is None or v == '': return False
        if isinstance(v, (int, float)): return True
        sv = str(v).strip()
        if sv in ['-', '—', '–']: return True
        c_v = sv.replace('.', '').replace(',', '').replace('(', '').replace(')', '').replace('$', '').replace('%', '').replace('M$', '').strip()
        return c_v.lstrip('-+').isdigit()

    for idx, col in enumerate(df.columns):
        col_series = df[col]
        non_empty_vals = [v for v in col_series if pd.notna(v) and str(v).strip() != '']
        num_count = sum(1 for v in col_series if _check_is_num(v))
        is_num_col = (num_count >= max(1, len(col_series) * 0.4))
        
        # Detectar columnas de viñetas / numerales cortos (ej: 1), 2), a), b))
        max_str_len = max((len(str(v).strip()) for v in non_empty_vals), default=0)
        is_bullet_col = (not is_num_col) and (max_str_len <= 5) and (idx == 0 or idx == 1)

        if is_num_col:
            styler = styler.set_properties(
                subset=[col],
                **{
                    'text-align': 'right',
                    'width': '130px',
                    'min-width': '110px',
                    'white-space': 'nowrap'
                }
            )
        elif is_bullet_col:
            styler = styler.set_properties(
                subset=[col],
                **{
                    'text-align': 'left',
                    'width': '35px',
                    'min-width': '30px',
                    'max-width': '50px',
                    'white-space': 'nowrap'
                }
            )
        else:
            styler = styler.set_properties(
                subset=[col],
                **{
                    'text-align': 'left',
                    'min-width': '280px',
                    'white-space': 'normal'
                }
            )

    dynamic_styled = False
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
                
            def convert_border(border_side):
                if not border_side or border_side.style is None:
                    return None
                style = border_side.style
                if style == 'double':
                    css_style = 'double'
                    width = '3px'
                elif style in ('medium', 'mediumDashed', 'mediumDashDot', 'mediumDashDotDot'):
                    css_style = 'solid'
                    width = '2px'
                elif style == 'thick':
                    css_style = 'solid'
                    width = '3px'
                elif style in ('dashed', 'dashDot', 'dashDotDot'):
                    css_style = 'dashed'
                    width = '1px'
                elif style == 'dotted':
                    css_style = 'dotted'
                    width = '1px'
                else:
                    css_style = 'solid'
                    width = '1px'
                
                color_hex = '000000'
                if border_side.color and border_side.color.rgb:
                    c_rgb = str(border_side.color.rgb)
                    if len(c_rgb) == 8:
                        color_hex = c_rgb[2:]
                    elif len(c_rgb) == 6:
                        color_hex = c_rgb
                return f"{width} {css_style} #{color_hex}"

            import datetime
            import re

            df_cols = list(df.columns)
            ws_cols_mapping = {}
            excel_cols_from_attrs = getattr(df, 'attrs', {}).get('excel_cols')
            if excel_cols_from_attrs and len(excel_cols_from_attrs) == len(df_cols):
                ws_cols_mapping = {j: excel_cols_from_attrs[j] for j in range(len(df_cols))}
            elif col_indices:
                ws_cols_mapping = {j: col_indices[j] for j in range(len(df_cols))}
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
                    for row in range(1, 7):
                        val = ws.cell(row=row, column=col).value
                        if val is not None:
                            val_str = str(val).strip()
                            if any(k in val_str.lower() for k in ['saldo inicial', 'saldo final', 'adiciones', 'informe', 'validacion']):
                                continue
                            is_date = (
                                isinstance(val, (datetime.datetime, datetime.date)) or
                                bool(re.search(r'\b20\d{2}\b', val_str))
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
                
                # Separar columnas del DataFrame en numéricas y no numéricas para alineación correcta
                non_numeric_df_cols = [0]
                for j in range(1, len(df_cols)):
                    col_name_lower = str(df_cols[j]).strip().lower()
                    if col_name_lower in ['nota', 'notas']:
                        non_numeric_df_cols.append(j)
                
                numeric_df_cols = [j for j in range(len(df_cols)) if j not in non_numeric_df_cols]

                # Mapear columnas no numéricas (Concepto, Moneda, Nota, etc.)
                excel_non_numeric_cols = [excel_name_col]
                if excel_nota_col is not None:
                    excel_non_numeric_cols.append(excel_nota_col)
                
                curr = excel_name_col + 1
                while len(excel_non_numeric_cols) < len(non_numeric_df_cols):
                    if curr not in excel_non_numeric_cols:
                        excel_non_numeric_cols.append(curr)
                    curr += 1
                
                for idx, j in enumerate(non_numeric_df_cols):
                    ws_cols_mapping[j] = excel_non_numeric_cols[idx]

                # Mapear columnas numéricas (Actual, Comparativa)
                excel_numeric_cols = list(excel_date_cols)
                curr = (excel_date_cols[-1] + 1) if excel_date_cols else (excel_name_col + 1)
                while len(excel_numeric_cols) < len(numeric_df_cols):
                    excel_numeric_cols.append(curr)
                    curr += 1
                
                for idx, j in enumerate(numeric_df_cols):
                    ws_cols_mapping[j] = excel_numeric_cols[idx]
            
            def clean_str(s):
                if s is None: return ""
                return str(s).replace("\xa0", " ").strip().lower()

            # Pre-construir mapa de búsqueda de filas en openpyxl
            row_map = {}
            name_search_col = ws_cols_mapping.get(0, 1)
            search_cols = [name_search_col] + [c for c in range(1, min(ws.max_column + 1, 25)) if c != name_search_col]
            for r in range(1, ws.max_row + 1):
                for col_idx in search_cols:
                    val = ws.cell(row=r, column=col_idx).value
                    if val is not None:
                        s_val = clean_str(val)
                        if s_val and s_val not in row_map:
                            row_map[s_val] = r

            row_styles_matrix = pd.DataFrame('', index=df.index, columns=df.columns)
            row_top_borders = {}
            row_bottom_borders = {}
            
            for idx, df_row in df.iterrows():
                search_val = clean_str(df_row.iloc[0])
                if not search_val and len(df_row) > 1:
                    search_val = clean_str(df_row.iloc[1])
                found_row = row_map.get(search_val)
                
                row_top_b = None
                row_bottom_b = None
                
                for j, col_name in enumerate(df_cols):
                    css_styles = []
                    if found_row is not None:
                        excel_col = ws_cols_mapping.get(j, j + 1)
                        cell = ws.cell(row=found_row, column=excel_col)
                        
                        if cell.font:
                            if cell.font.bold:
                                css_styles.append("font-weight: bold")
                            else:
                                css_styles.append("font-weight: normal")
                            if cell.font.italic:
                                css_styles.append("font-style: italic")
                            if cell.font.size:
                                css_styles.append(f"font-size: {cell.font.size}pt")
                            if cell.font.name:
                                css_styles.append(f"font-family: '{cell.font.name}', sans-serif")
                        
                        val_curr = df_row.iloc[j] if j < len(df_row) else None
                        def _is_num_cell(v):
                            if pd.isna(v) or v is None or v == "": return False
                            if isinstance(v, (int, float)): return True
                            sv = str(v).strip()
                            if sv in ["-", "—", "–"]: return True
                            c_v = sv.replace(".", "").replace(",", "").replace("(", "").replace(")", "").replace("$", "").replace("%", "").replace("M$", "").strip()
                            return c_v.lstrip("-+").isdigit()

                        if _is_num_cell(val_curr):
                            css_styles = [s for s in css_styles if not s.startswith("text-align")]
                            css_styles.append("text-align: right !important")
                        else:
                            val_str = str(val_curr).strip() if val_curr is not None else ""
                            is_num_col_header = (j >= 2) and (val_str.lower() in ["m$", "$", "%", "activos", "pasivos"] or any(k in val_str for k in ["31.12", "30.06", "31.03", "30.09", "202"]))
                            if not is_num_col_header:
                                css_styles = [s for s in css_styles if not s.startswith("text-align")]
                                css_styles.append("text-align: left !important")
                                
                        if cell.fill and cell.fill.fill_type == 'solid' and cell.fill.fgColor:
                            c_rgb = str(cell.fill.fgColor.rgb)
                            if len(c_rgb) == 8 and c_rgb != '00000000' and not c_rgb.startswith('000000'):
                                css_styles.append(f"background-color: #{c_rgb[2:]}")
                            elif len(c_rgb) == 6:
                                css_styles.append(f"background-color: #{c_rgb}")
                        
                        top_b = None
                        bottom_b = None
                        if cell.border:
                            top_b = convert_border(cell.border.top)
                            bottom_b = convert_border(cell.border.bottom)
                        
                        if top_b:
                            css_styles.append(f"border-top: {top_b} !important")
                            if not row_top_b and top_b != 'none':
                                row_top_b = top_b
                        else:
                            css_styles.append("border-top: none")
                            
                        if bottom_b:
                            css_styles.append(f"border-bottom: {bottom_b} !important")
                            if not row_bottom_b and bottom_b not in ('none', '1px solid #f0f0f0'):
                                row_bottom_b = bottom_b
                        else:
                            css_styles.append("border-bottom: 1px solid #f0f0f0")
                        
                    if css_styles:
                        row_styles_matrix.loc[idx, col_name] = "; ".join(css_styles)
                            
                if row_top_b:
                    row_top_borders[idx] = row_top_b
                if row_bottom_b:
                    row_bottom_borders[idx] = row_bottom_b
            
            # Aplicar propagación de bordes horizontalmente en filas de totales
            for idx in row_styles_matrix.index:
                row_top_b = row_top_borders.get(idx)
                row_bottom_b = row_bottom_borders.get(idx)
                if row_top_b or row_bottom_b:
                    for col_name in df_cols:
                        style_str = row_styles_matrix.loc[idx, col_name]
                        if row_top_b:
                            style_str = re.sub(r'border-top:[^;]+', f'border-top: {row_top_b} !important', style_str)
                        if row_bottom_b:
                            style_str = re.sub(r'border-bottom:[^;]+', f'border-bottom: {row_bottom_b} !important', style_str)
                        row_styles_matrix.loc[idx, col_name] = style_str

            # Pasada de resolución de conflictos de colapso de bordes
            for i in range(len(df)):
                idx = df.index[i]
                row_top_b = row_top_borders.get(idx)
                row_bottom_b = row_bottom_borders.get(idx)
                if row_top_b and i > 0:
                    prev_idx = df.index[i - 1]
                    for col_name in df_cols:
                        row_styles_matrix.loc[prev_idx, col_name] += f"; border-bottom: {row_top_b} !important"
                if row_bottom_b and i < len(df) - 1:
                    next_idx = df.index[i + 1]
                    for col_name in df_cols:
                        row_styles_matrix.loc[next_idx, col_name] += f"; border-top: {row_bottom_b} !important"
            
            styler = styler.apply(lambda _: row_styles_matrix, axis=None)
            dynamic_styled = True
        except Exception as ex:
            import sys
            print(f"Error parsing Excel styles dynamically: {ex}", file=sys.stderr)
            dynamic_styled = False
            
    if not dynamic_styled:
        def highlight_totals(row):
            classif = str(row.iloc[0]).lower().strip()
            
            styles = [''] * len(row)
            if classif == "estado de resultados":
                for i in range(len(row)):
                    styles[i] = 'border-top: 2px solid #000000 !important; border-bottom: 2px solid #000000 !important; font-weight: bold; color: #000000; background-color: #FFFFFF;'
                return styles

            is_bottom_line = (
                "ganancias (pérdida) del ejercicio" in classif or
                "ganancia (pérdida) del ejercicio" in classif or
                "resultado final" in classif or
                classif == "ganancia (pérdida)"
            )
            
            if is_bottom_line:
                for i, col_name in enumerate(row.index):
                    if col_name in numeric_cols:
                        styles[i] = 'border-top: 1px solid #000000 !important; border-bottom: 3px double #000000 !important; font-weight: bold; color: #000000;'
                    elif i == 0:
                        styles[i] = 'font-weight: bold; color: #000000; border-top: 1px solid #000000 !important; border-bottom: 3px double #000000 !important;'
                return styles
                
            is_total = (
                "total" in classif or 
                "bruta" in classif or 
                "antes de" in classif or 
                "antes del" in classif or 
                "incremento" in classif or
                ("ganancia" in classif and "acumulada" not in classif) or
                ("pérdida" in classif and "acumulada" not in classif and "operacionales" not in classif and "operaciones continuadas" in classif) or
                "saldo final" in classif
            )
            
            if is_total:
                for i, col_name in enumerate(row.index):
                    if col_name in numeric_cols:
                        styles[i] = 'border-top: 1px solid #000000 !important; border-bottom: 1px solid #000000 !important; font-weight: bold; color: #000000; background-color: #f8f9fa;'
                    elif i == 0:
                        styles[i] = 'font-weight: bold; color: #000000; border-top: 1px solid #000000 !important; border-bottom: 1px solid #000000 !important; background-color: #f8f9fa;'
            return styles

        styler = styler.apply(highlight_totals, axis=1)
        
    return styler

