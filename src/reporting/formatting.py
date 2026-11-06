import pandas as pd
import re


def detect_internal_header_rows(df: pd.DataFrame):
    """
    Identifica los índices posicionales de fila (0-based en df) que conforman el encabezado
    interno de una tabla cuando el DataFrame no tiene nombres de columna reales (ej. cuadros
    de Notas donde los encabezados vienen en las primeras filas) o cuando existe un segundo
    bloque de encabezado intermedio (ej. segundo año en #PATRIMONIO o en cuadros de notas).
    Retorna: (hdr_rows_set, bottom_hdr_rows_set)
    """
    hdr_rows = set()
    bottom_hdr_rows = set()
    if df is None or df.empty:
        return hdr_rows, bottom_hdr_rows

    num_rows = len(df)
    num_cols = len(df.columns)
    df_cols = list(df.columns)

    def _is_dummy_col(c):
        if isinstance(c, (int, float)):
            return True
        s = str(c).replace("\nM$", "").replace("\nCh$", "").replace("\nTh$", "").strip()
        return s == "" or s.isdigit()

    has_real_header = not all(_is_dummy_col(c) for c in df_cols)

    def _clean(v):
        if pd.isna(v) or v is None:
            return ""
        return str(v).replace("\xa0", " ").strip()

    def _is_unit_token(s):
        sl = s.lower()
        return sl in ["m$", "ch$", "th$", "us$", "$", "%", "uf", "mus$", "mch$", "miles de $"]

    def _row_is_unit_line(r_idx):
        if r_idx < 0 or r_idx >= num_rows:
            return False
        row_vals = [_clean(df.iloc[r_idx, c]) for c in range(num_cols)]
        right_vals = [v for v in row_vals[1:] if v != ""]
        if not right_vals:
            return False
        has_unit = any(_is_unit_token(v) for v in right_vals)
        if not has_unit:
            return False
        non_unit_cells = [v for v in row_vals if v != "" and not _is_unit_token(v)]
        allowed_col_labels = {
            "detalle", "concepto", "movimiento", "movimientos", "tipo", "tipo de contrato",
            "rubro", "item", "ítem", "cuenta", "descripción", "descripcion", "clasificación",
            "clasificacion", "rut", "r.u.t.", "entidad", "institución", "institucion",
            "nombre", "nombre entidad", "participación", "participacion", "moneda", "nota", "notas",
            "cuentas comerciales según plazo", "cuentas comerciales segun plazo"
        }
        return all(v.lower() in allowed_col_labels for v in non_unit_cells)

    def _cell_looks_like_date_or_hdr(s):
        if not s:
            return False
        sl = s.lower()
        if _is_unit_token(sl):
            return True
        if re.search(r'\b20\d{2}\b', sl) or re.search(r'\d{2}[./-]\d{2}[./-]\d{2,4}', sl):
            return True
        hdr_keywords = [
            "corriente", "no corriente", "activos", "pasivos", "bruto", "neto", "amortizaci",
            "depreciaci", "moneda", "banco", "vencimiento", "tasa", "interés", "interes",
            "concesiones", "licencias", "marcas", "cartera", "terrenos", "edificios", "redes",
            "plantas", "equipos", "vehículos", "vehiculos", "muebles", "instalación", "instalacion",
            "obras", "leasing", "total", "residencial", "mayorista", "eliminaciones", "aumento",
            "baja", "valor al cierre", "10%", "90 días", "90 dias", "1 año", "1 ano", "5 años",
            "5 anos", "3 años", "3 anos", "7 años", "7 anos", "saldo inicial", "saldo final",
            "nuevos", "adiciones", "capital", "reajuste", "otros", "mark to market", "instrument",
            "description", "issuer", "currency", "issue date", "maturity", "transaction",
            "annual", "cobertura", "partida cubierta", "riesgo", "liquidación", "liquidacion",
            "naturaleza", "relación", "relacion", "transacción", "transaccion", "efecto",
            "participación", "participacion", "acciones", "monto", "inicio", "termino", "término",
            "ganancia", "acumulad", "reserva", "patrimonio", "issued", "retained", "equity",
            "activo m$", "pasivo m$"
        ]
        return any(k in sl for k in hdr_keywords)

    if not has_real_header:
        last_unit_row = None
        for r in range(min(5, num_rows)):
            if _row_is_unit_line(r):
                last_unit_row = r

        if last_unit_row is not None:
            for r in range(0, last_unit_row + 1):
                hdr_rows.add(r)
            bottom_hdr_rows.add(last_unit_row)
        else:
            max_hdr_r = None
            for r in range(min(3, num_rows)):
                row_vals = [_clean(df.iloc[r, c]) for c in range(num_cols)]
                right_vals = [v for v in row_vals[1:] if v != ""]
                c0 = row_vals[0]
                if r == 0 and num_rows > 1:
                    next_vals = [_clean(df.iloc[1, c]) for c in range(1, num_cols) if _clean(df.iloc[1, c]) != ""]
                    if next_vals and any(_cell_looks_like_date_or_hdr(v) for v in next_vals):
                        if (c0 != "" and not right_vals) or (c0 == "" and right_vals and all(_cell_looks_like_date_or_hdr(v) for v in right_vals)):
                            hdr_rows.add(0)
                            max_hdr_r = 0
                            continue
                if right_vals and all(_cell_looks_like_date_or_hdr(v) for v in right_vals):
                    all_pure_nums = all(
                        v.replace(".", "").replace(",", "").replace("-", "").replace("(", "").replace(")", "").isdigit()
                        and not re.match(r'^20\d{2}$', v.strip())
                        for v in right_vals
                    )
                    if not all_pure_nums:
                        for prev_r in range(0, r + 1):
                            hdr_rows.add(prev_r)
                        max_hdr_r = r
            if max_hdr_r is not None:
                bottom_hdr_rows.add(max_hdr_r)
            elif num_rows > 0:
                hdr_rows.add(0)
                bottom_hdr_rows.add(0)

    allowed_col_labels = {
        "detalle", "concepto", "movimiento", "movimientos", "tipo", "tipo de contrato",
        "rubro", "item", "ítem", "cuenta", "descripción", "descripcion", "clasificación",
        "clasificacion", "rut", "r.u.t.", "entidad", "institución", "institucion",
        "nombre", "nombre entidad", "participación", "participacion", "moneda", "nota", "notas",
        "cuentas comerciales según plazo", "cuentas comerciales segun plazo"
    }

    # Detectar bloques de encabezados intermedios (ej. segundo año en #PATRIMONIO o en cuadros multi-año de Notas)
    for r in range(1, num_rows - 1):
        if r in hdr_rows:
            continue
        c0 = _clean(df.iloc[r, 0])
        c0_low = c0.lower()
        is_allowed_c0 = (
            c0 == "" or
            c0_low in allowed_col_labels or
            c0_low.startswith("detalle al") or
            c0_low.startswith("concepto al")
        )
        if is_allowed_c0 and _row_is_unit_line(r + 1):
            right_vals = [_clean(df.iloc[r, c]).lower() for c in range(1, num_cols) if _clean(df.iloc[r, c]) != ""]
            if right_vals and any(_cell_looks_like_date_or_hdr(v) for v in right_vals):
                hdr_rows.add(r)
                hdr_rows.add(r + 1)
                bottom_hdr_rows.add(r + 1)
                for back_step in [1, 2]:
                    prev_r = r - back_step
                    if prev_r <= 0 or prev_r in hdr_rows:
                        break
                    prev_all = [_clean(df.iloc[prev_r, c]) for c in range(num_cols)]
                    prev_non_empty = [v for v in prev_all if v != ""]
                    if not prev_non_empty:
                        break
                    prev_c0_low = prev_all[0].lower()
                    if any(prev_c0_low.startswith(k) for k in ["total", "saldo final", "saldo al", "valores netos", "sub total", "subtotal"]):
                        break
                    # Caso A: fila superior con 1 o 2 celdas de fecha/año o título de cuadro (ej. '2025', '31.12.2024', 'VENCIMIENTO DE PASIVOS...')
                    if len(prev_non_empty) <= 2 and all(
                        re.search(r'\b20\d{2}\b', pv) or
                        re.search(r'\d{2}[./-]\d{2}[./-]\d{2,4}', pv) or
                        any(k in pv.lower() for k in ["vencimiento", "cambios en pasivo", "movimiento"])
                        for pv in prev_non_empty
                    ):
                        hdr_rows.add(prev_r)
                        continue
                    # Caso B: fila superior de títulos de columna multilínea (c0 vacío y celdas derechas son etiquetas de cabecera)
                    if prev_all[0] == "" and all(_cell_looks_like_date_or_hdr(pv) or pv.lower() in ["pagos", "equipo", "descripción de la", "descripcion de la", "edificios y", "redes de", "plantas", "equipos"] for pv in prev_non_empty):
                        hdr_rows.add(prev_r)
                        continue
                    break

    return hdr_rows, bottom_hdr_rows


def is_separator_row_series(row_series) -> bool:
    """Retorna True si todas las celdas de la fila están vacías o son marcadores nulos."""
    for v in row_series:
        if pd.notna(v) and v is not None:
            s = str(v).replace("\xa0", "").strip()
            if s not in ["", "nan", "None", "||"]:
                return False
    return True


def ensure_subtable_separator_rows(df: pd.DataFrame) -> pd.DataFrame:
    """
    Garantiza que cuando un DataFrame contiene dos o más cuadros apilados verticalmente
    (detectados por bloques de encabezado intermedios en detect_internal_header_rows),
    exista exactamente una fila vacía de separación antes de cada cuadro intermedio,
    respetando la línea de separación de las plantillas Excel.
    Es 100% idempotente.
    """
    if df is None or df.empty:
        return df

    orig_attrs = dict(getattr(df, "attrs", {}))
    df_non_empty = df[~df.apply(is_separator_row_series, axis=1)].reset_index(drop=True)
    if df_non_empty.empty:
        df_non_empty.attrs.update(orig_attrs)
        return df_non_empty

    hdr_rows, _ = detect_internal_header_rows(df_non_empty)
    mid_starts = sorted([r for r in hdr_rows if r > 0 and (r - 1) not in hdr_rows])
    if not mid_starts:
        df_non_empty.attrs.update(orig_attrs)
        return df_non_empty

    chunks = []
    prev_idx = 0
    sep_row_df = pd.DataFrame([[""] + [None] * (len(df_non_empty.columns) - 1)], columns=df_non_empty.columns)
    for r_start in mid_starts:
        chunks.append(df_non_empty.iloc[prev_idx:r_start])
        chunks.append(sep_row_df)
        prev_idx = r_start
    chunks.append(df_non_empty.iloc[prev_idx:])

    result = pd.concat(chunks, ignore_index=True)
    result.attrs.update(orig_attrs)
    return result


def apply_corporate_style(df, excel_bytes=None, sheet_name=None, col_indices=None, unit=None, is_miles=None, target_lang=None):
    import openpyxl
    from io import BytesIO

    if target_lang is None:
        try:
            import streamlit as st
            target_lang = st.session_state.get('idioma_reporte', 'es')
        except Exception:
            target_lang = 'es'

    df = ensure_subtable_separator_rows(df.copy())

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
        has_real_hdr_cols = not all(
            isinstance(c, (int, float)) or str(c).strip() == "" or str(c).strip().isdigit()
            for c in df.columns
        )
        if has_real_hdr_cols:
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

    internal_hdr_rows, bottom_hdr_rows = detect_internal_header_rows(df)
    sep_rows = {
        r_pos for r_pos in range(len(df))
        if r_pos not in internal_hdr_rows and is_separator_row_series(df.iloc[r_pos])
    }
    if sep_rows and len(df.columns) > 0:
        for r_pos in sep_rows:
            df.iloc[r_pos, 0] = "\u00a0"

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
                ('background-color', '#E8EEF6 !important'),
                ('color', '#1F3864 !important'),
                ('font-weight', 'bold !important'),
                ('font-family', 'Inter, sans-serif'),
                ('border', 'none !important'),
                ('padding', '8px'),
                ('border-top', 'none !important'),
                ('border-left', 'none !important'),
                ('border-right', 'none !important'),
                ('border-bottom', '2px solid #1F3864 !important'),
                ('text-align', 'left')
            ]
        },
        {
            'selector': 'tr',
            'props': [
                ('border', 'none !important')
            ]
        },
        {
            'selector': 'td',
            'props': [
                ('font-family', 'Inter, sans-serif'),
                ('font-size', '0.85em'),
                ('border-left', 'none !important'),
                ('border-right', 'none !important'),
                ('border-top', 'none !important'),
                ('border-bottom', '1px solid #E1E1E1'),
                ('padding', '6px 10px'),
                ('background-clip', 'padding-box')
            ]
        },
        {
            'selector': 'tr:nth-child(even) td',
            'props': [('background-color', '#FFFFFF')]
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

    # Detectar de antemano las filas de cabecera internas (Notas y bloques multi-año como #PATRIMONIO)
    internal_hdr_rows, bottom_hdr_rows = detect_internal_header_rows(df)

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
                    css_style = 'solid'
                    width = '1px'
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
                    if col <= excel_name_col:
                        continue
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
            curr_excel_row = 1

            def _is_unit_only_row(r_series):
                c0 = clean_str(r_series.iloc[0]) if len(r_series) > 0 else ""
                if c0 != "":
                    return False
                rest = [clean_str(r_series.iloc[c]) for c in range(1, len(r_series)) if pd.notna(r_series.iloc[c]) and clean_str(r_series.iloc[c]) != ""]
                return bool(rest) and all(v in ["m$", "ch$", "th$", "us$", "$", "%", "uf"] for v in rest)
            
            for row_pos, (idx, df_row) in enumerate(df.iterrows()):
                first_col_str = clean_str(df_row.iloc[0])
                first_col_empty = (first_col_str == "")
                search_val = first_col_str
                search_excel_col = ws_cols_mapping.get(0, 1)
                if not search_val and len(df_row) > 1:
                    search_val = clean_str(df_row.iloc[1])
                    search_excel_col = ws_cols_mapping.get(1, search_excel_col + 1)

                found_row = None
                if search_val:
                    for r in range(curr_excel_row, ws.max_row + 1):
                        val = ws.cell(row=r, column=search_excel_col).value
                        if val is not None and clean_str(val) == search_val:
                            found_row = r
                            curr_excel_row = r + 1
                            break
                    if found_row is None:
                        found_row = row_map.get(search_val)
                
                next_is_unit_row = (row_pos + 1 < len(df)) and _is_unit_only_row(df.iloc[row_pos + 1])
                curr_is_mid_hdr_title = first_col_empty and next_is_unit_row
                curr_is_unit_row = _is_unit_only_row(df_row)
                is_internal_hdr = (row_pos in internal_hdr_rows)
                is_bot_hdr = (row_pos in bottom_hdr_rows)
                next_is_internal_hdr = ((row_pos + 1) in internal_hdr_rows)

                row_top_b = None
                row_bottom_b = None
                
                for j, col_name in enumerate(df_cols):
                    css_styles = []
                    if found_row is not None:
                        excel_col = ws_cols_mapping.get(j, j + 1)
                        cell = ws.cell(row=found_row, column=excel_col)
                        
                        if cell.font:
                            if cell.font.bold or curr_is_mid_hdr_title or is_internal_hdr:
                                css_styles.append("font-weight: bold")
                            else:
                                css_styles.append("font-weight: normal")
                            if cell.font.italic:
                                css_styles.append("font-style: italic")
                            if cell.font.size:
                                css_styles.append(f"font-size: {cell.font.size}pt")
                            if cell.font.name:
                                css_styles.append(f"font-family: '{cell.font.name}', sans-serif")
                        elif curr_is_mid_hdr_title or is_internal_hdr:
                            css_styles.append("font-weight: bold")
                        
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
                            is_num_col_header = (j >= 1) and (
                                first_col_empty or
                                val_str.lower() in ["m$", "$", "%", "activos", "pasivos", "corriente", "no corriente", "us$", "uf", "th$"] or
                                any(k in val_str.lower() for k in ["capital", "ganancia", "acumulad", "reserva", "patrimonio", "issued", "retained", "equity"]) or
                                any(k in val_str for k in ["31.12", "30.06", "31.03", "30.09", "202", "31-", "30-"])
                            )
                            css_styles = [s for s in css_styles if not s.startswith("text-align")]
                            if is_num_col_header:
                                css_styles.append("text-align: right !important")
                                css_styles.append("white-space: normal !important")
                            else:
                                css_styles.append("text-align: left !important")
                                
                        if is_internal_hdr:
                            css_styles.append("background-color: #E8EEF6 !important")
                            css_styles.append("color: #1F3864 !important")
                            css_styles.append("font-weight: bold !important")
                            css_styles.append("border-top: none !important")
                            css_styles.append("border-left: none !important")
                            css_styles.append("border-right: none !important")
                            if is_bot_hdr:
                                css_styles.append("border-bottom: 2px solid #1F3864 !important")
                            else:
                                css_styles.append("border-bottom: none !important")
                        else:
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
                            
                            if curr_is_mid_hdr_title:
                                css_styles.append("padding-top: 14px !important")
                                css_styles.append("padding-bottom: 2px !important")
                                bottom_b = None

                            if curr_is_unit_row and not bottom_b:
                                bottom_b = "1px solid #000000"

                            if top_b:
                                css_styles.append(f"border-top: {top_b} !important")
                                if not row_top_b and top_b != 'none':
                                    row_top_b = top_b
                            else:
                                css_styles.append("border-top: none")
                                
                            if curr_is_mid_hdr_title or next_is_internal_hdr:
                                css_styles.append("border-bottom: none !important")
                            elif bottom_b:
                                css_styles.append(f"border-bottom: {bottom_b} !important")
                                if not row_bottom_b and bottom_b not in ('none', '1px solid #f0f0f0'):
                                    row_bottom_b = bottom_b
                            else:
                                css_styles.append("border-bottom: 1px solid #f0f0f0")
                    elif is_internal_hdr:
                        css_styles.append("background-color: #E8EEF6 !important")
                        css_styles.append("color: #1F3864 !important")
                        css_styles.append("font-weight: bold !important")
                        css_styles.append("border-top: none !important")
                        css_styles.append("border-left: none !important")
                        css_styles.append("border-right: none !important")
                        if is_bot_hdr:
                            css_styles.append("border-bottom: 2px solid #1F3864 !important")
                        else:
                            css_styles.append("border-bottom: none !important")
                        
                    if css_styles:
                        row_styles_matrix.loc[idx, col_name] = "; ".join(css_styles)
                            
                if not is_internal_hdr:
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
                if row_top_b and i > 0 and (i - 1) not in internal_hdr_rows:
                    prev_idx = df.index[i - 1]
                    for col_name in df_cols:
                        row_styles_matrix.loc[prev_idx, col_name] += f"; border-bottom: {row_top_b} !important"
                if row_bottom_b and i < len(df) - 1 and (i + 1) not in internal_hdr_rows:
                    next_idx = df.index[i + 1]
                    for col_name in df_cols:
                        row_styles_matrix.loc[next_idx, col_name] += f"; border-top: {row_bottom_b} !important"

            # Reforzar que el bloque de cabecera interna sea 100% sólido sin líneas internas
            for row_pos in internal_hdr_rows:
                idx_h = df.index[row_pos]
                is_bot_h = (row_pos in bottom_hdr_rows)
                b_bot_h = "2px solid #1F3864" if is_bot_h else "none"
                for col_name in df_cols:
                    style_str = row_styles_matrix.loc[idx_h, col_name]
                    style_str = re.sub(r'border-top:[^;]+', 'border-top: none !important', style_str)
                    style_str = re.sub(r'border-bottom:[^;]+', f'border-bottom: {b_bot_h} !important', style_str)
                    row_styles_matrix.loc[idx_h, col_name] = style_str
                if (row_pos + 1) < len(df) and (row_pos + 1) not in internal_hdr_rows and is_bot_h:
                    nxt_idx = df.index[row_pos + 1]
                    for col_name in df_cols:
                        style_str = row_styles_matrix.loc[nxt_idx, col_name]
                        style_str = re.sub(r'border-top:[^;]+', 'border-top: 2px solid #1F3864 !important', style_str)
                        row_styles_matrix.loc[nxt_idx, col_name] = style_str
            
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
            
            from src.reporting.word_export import _is_total_or_subtotal_label
            if is_bottom_line or _is_total_or_subtotal_label(classif):
                for i in range(len(row)):
                    styles[i] = 'border-top: 1px solid #404040 !important; border-bottom: 1px solid #404040 !important; font-weight: bold; color: #000000; background-color: #FFFFFF;'
            return styles

        styler = styler.apply(highlight_totals, axis=1)

    # Aplicar estilo corporativo celeste (#E8EEF6) y texto azul marino (#1F3864) en negrita
    # a las filas de encabezado internas de cuadros de Notas y bloques intermedios (ej. #PATRIMONIO),
    # así como la fila en blanco de separación entre cuadros.
    if internal_hdr_rows or sep_rows:
        df_cols_list = list(df.columns)
        hdr_styles_matrix = pd.DataFrame('', index=df.index, columns=df.columns)
        for row_pos, idx_val in enumerate(df.index):
            if row_pos in sep_rows:
                prev_bot = "1px solid #404040"
                if row_pos > 0:
                    prev_idx_val = df.index[row_pos - 1]
                    if dynamic_styled and 'row_bottom_borders' in locals() and prev_idx_val in row_bottom_borders:
                        prev_bot = row_bottom_borders[prev_idx_val]
                    for j, c_name in enumerate(df_cols_list):
                        hdr_styles_matrix.loc[prev_idx_val, c_name] = f"border-bottom: {prev_bot} !important;"
                for j, c_name in enumerate(df_cols_list):
                    hdr_styles_matrix.loc[idx_val, c_name] = (
                        f"background-color: #FFFFFF !important; "
                        f"border-top: {prev_bot} !important; "
                        f"border-bottom: none !important; "
                        f"border-left: none !important; "
                        f"border-right: none !important; "
                        f"height: 22px !important;"
                    )
            elif row_pos in internal_hdr_rows:
                is_bot_hdr = (row_pos in bottom_hdr_rows)
                b_bottom = "2px solid #1F3864" if is_bot_hdr else "none"
                for j, c_name in enumerate(df_cols_list):
                    hdr_styles_matrix.loc[idx_val, c_name] = (
                        f"background-color: #E8EEF6 !important; "
                        f"color: #1F3864 !important; "
                        f"font-weight: bold !important; "
                        f"border-top: none !important; "
                        f"border-left: none !important; "
                        f"border-right: none !important; "
                        f"border-bottom: {b_bottom} !important;"
                    )
        styler = styler.apply(lambda _: hdr_styles_matrix, axis=None)
        
    return styler

