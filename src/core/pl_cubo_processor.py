import pandas as pd
import numpy as np
import unicodedata

def normalize_text(text):
    """
    Normaliza el texto en minúsculas, sin espacios extras ni acentos (diacríticos)
    para comparaciones estables e independientes de codificación.
    """
    if pd.isna(text):
        return ""
    s = str(text).strip().lower()
    # Descomponer caracteres y remover marcas de acento
    s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
    return s

def parse_month_to_num(val):
    """
    Convierte cualquier representación de mes (número, texto en español/inglés,
    abreviación de 3 o 4 letras como 'sept', formato YYYY-MM, etc.) a un entero entre 1 y 12.
    """
    if pd.isna(val):
        return None
    s = str(val).strip().lower()
    if not s or s == 'nan' or s == 'none':
        return None
    
    # 1. Si contiene separadores de fecha o periodo ('YYYY-MM', 'YYYY/MM', 'DD-MM-YYYY')
    if '-' in s or '/' in s:
        parts = s.replace('/', '-').split('-')
        if len(parts) >= 2:
            # Formato ISO: YYYY-MM o YYYY-MM-DD
            if len(parts[0]) == 4 and parts[0].isdigit():
                m_part = parts[1].strip()
                if m_part.isdigit() and 1 <= int(m_part) <= 12:
                    return int(m_part)
            # Formato tradicional: DD-MM-YYYY
            elif len(parts[-1]) == 4 and parts[-1].isdigit():
                m_part = parts[1].strip()
                if m_part.isdigit() and 1 <= int(m_part) <= 12:
                    return int(m_part)

    # 2. Intento directo como número entero o float ('09', '9', 9, 9.0)
    try:
        n = int(float(s))
        if 1 <= n <= 12:
            return n
    except (ValueError, TypeError):
        pass

    # 3. Normalización de texto (sin acentos, minúsculas)
    s_norm = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')

    month_map = {
        # Enero
        'ene': 1, 'ener': 1, 'enero': 1, 'jan': 1, 'january': 1, '01': 1, '1': 1, 'i': 1,
        # Febrero
        'feb': 2, 'febr': 2, 'febrero': 2, 'february': 2, '02': 2, '2': 2, 'ii': 2,
        # Marzo
        'mar': 3, 'marz': 3, 'marzo': 3, 'march': 3, '03': 3, '3': 3, 'iii': 3,
        # Abril
        'abr': 4, 'abri': 4, 'abril': 4, 'apr': 4, 'april': 4, '04': 4, '4': 4, 'iv': 4,
        # Mayo
        'may': 5, 'mayo': 5, '05': 5, '5': 5, 'v': 5,
        # Junio
        'jun': 6, 'juni': 6, 'junio': 6, 'june': 6, '06': 6, '6': 6, 'vi': 6,
        # Julio
        'jul': 7, 'juli': 7, 'julio': 7, 'july': 7, '07': 7, '7': 7, 'vii': 7,
        # Agosto
        'ago': 8, 'agos': 8, 'agosto': 8, 'aug': 8, 'august': 8, '08': 8, '8': 8, 'viii': 8,
        # Septiembre (incluye 'sept', 'sep', 'set', 'seti', 'septiembre', 'setiembre', etc.)
        'sep': 9, 'sept': 9, 'set': 9, 'seti': 9, 'septiembre': 9, 'setiembre': 9, 'september': 9, '09': 9, '9': 9, 'ix': 9,
        # Octubre
        'oct': 10, 'octu': 10, 'octubre': 10, 'october': 10, '10': 10, 'x': 10,
        # Noviembre
        'nov': 11, 'novi': 11, 'noviembre': 11, 'november': 11, '11': 11, 'xi': 11,
        # Diciembre
        'dic': 12, 'dici': 12, 'diciembre': 12, 'dec': 12, 'december': 12, '12': 12, 'xii': 12,
    }

    if s_norm in month_map:
        return month_map[s_norm]

    # 4. Prefijos comunes
    for prefix, m_num in [
        ('ener', 1), ('ene', 1), ('jan', 1),
        ('febr', 2), ('feb', 2),
        ('marz', 3), ('mar', 3),
        ('abri', 4), ('abr', 4), ('apr', 4),
        ('mayo', 5), ('may', 5),
        ('juni', 6), ('jun', 6),
        ('juli', 7), ('jul', 7),
        ('agos', 8), ('ago', 8), ('aug', 8),
        ('sept', 9), ('sep', 9), ('seti', 9), ('set', 9),
        ('octu', 10), ('oct', 10),
        ('novi', 11), ('nov', 11),
        ('dici', 12), ('dic', 12), ('dec', 12)
    ]:
        if s_norm.startswith(prefix):
            return m_num

    return None

def parse_year_to_num(val):
    """
    Extrae el año como número entero de 4 dígitos.
    """
    if pd.isna(val):
        return None
    s = str(val).strip()
    if not s or s == 'nan' or s == 'none':
        return None
    import re
    m = re.search(r'\b(20\d\d|19\d\d)\b', s)
    if m:
        return int(m.group(1))
    try:
        n = int(float(s))
        if 1900 <= n <= 2100:
            return n
    except (ValueError, TypeError):
        pass
    return None

def process_odoo_cubo(df_cubo, year, month, map_pl_df, standard_categories=None, return_audit_log=False):
    """
    Procesa un DataFrame transaccional del Cubo de Odoo, filtra por periodo YTD,
    realiza la homologación de cuentas contra map_pl.xlsx y pivota los datos
    al formato P&L configurado.
    
    Si return_audit_log=True, retorna (df_pivot, df_audit) donde df_audit contiene
    las filas del Excel que no tenían clasificación o requerían asignación por descarte.
    """
    # 1. Detectar nombres de columnas de forma dinámica
    col_nivel1 = next((c for c in df_cubo.columns if normalize_text(c) in ['nivel 1', 'nivel_1', 'nivel1', 'n1']), None)

    col_year = next((c for c in df_cubo.columns if any(k in normalize_text(c) for k in ['ano', 'año', 'year']) or ('fec_doc' in normalize_text(c) and 'mes' not in normalize_text(c) and normalize_text(c) != 'fec_doc')), None)
    if not col_year and 'fec_doc' in df_cubo.columns:
        col_year = 'fec_doc'

    col_month = next((c for c in df_cubo.columns if 'mes' in normalize_text(c) or 'month' in normalize_text(c)), None)
    if not col_month and 'fec_doc' in df_cubo.columns:
        col_month = 'fec_doc'

    # Priorizar nombres exactos para cuenta y nombre
    col_cuenta = next((c for c in df_cubo.columns if normalize_text(c) in ["cuenta", "n° de cuenta", "n de cuenta", "cuenta_id", "cod_cuenta"]), None)
    if not col_cuenta:
        col_cuenta = next((c for c in df_cubo.columns if "cuenta" in normalize_text(c) and "nombre" not in normalize_text(c) and "agrup" not in normalize_text(c) and "fcst" not in normalize_text(c) and "flujo" not in normalize_text(c)), None)

    col_nombre = next((c for c in df_cubo.columns if normalize_text(c) in ["nombre_cuenta", "nombre de la cuenta", "nombre cuenta", "descripcion", "desc"]), None)
    if not col_nombre:
        col_nombre = next((c for c in df_cubo.columns if "nombre" in normalize_text(c) and "empresa" not in normalize_text(c) and "depto" not in normalize_text(c) and "departamento" not in normalize_text(c)), None)

    # Priorizar importe_mn / importe
    col_importe = next((c for c in df_cubo.columns if normalize_text(c) in ["importe_mn", "importe mn", "importe", "monto", "saldo", "monto_mn"]), None)
    if not col_importe:
        col_importe = next((c for c in df_cubo.columns if any(k in normalize_text(c) for k in ["importe", "monto", "saldo"])), None)

    # Priorizar informe_ee_rr (columna 37 / con guiones o espacios) sobre informe_eerr
    col_category = next((c for c in df_cubo.columns if normalize_text(c) in ["informe_ee_rr", "informeee_rr", "informe ee rr", "ee_rr", "eerr"]), None)
    if not col_category:
        col_category = next((c for c in df_cubo.columns if "informe_ee_rr" in normalize_text(c) or "ee_rr" in normalize_text(c)), None)
    if not col_category:
        col_category = next((c for c in df_cubo.columns if "informe_eerr" in normalize_text(c) or "eerr" in normalize_text(c)), None)

    # Fallbacks de detección por posición por si fallan los nombres
    if not col_cuenta:
        col_cuenta = df_cubo.columns[2] if len(df_cubo.columns) > 2 else df_cubo.columns[0]
    if not col_nombre:
        col_nombre = df_cubo.columns[3] if len(df_cubo.columns) > 3 else df_cubo.columns[1]
    if not col_importe:
        col_importe = next((c for c in df_cubo.columns if df_cubo[c].dtype in [np.float64, np.int64]), df_cubo.columns[4])
    if not col_category:
        col_category = df_cubo.columns[6] if len(df_cubo.columns) > 6 else df_cubo.columns[0]

    df_filtered = df_cubo.copy()
    # Guardar número de fila original de Excel (índice + 2 asumiendo fila 1 de encabezados)
    df_filtered['_excel_row'] = df_cubo.index + 2

    # 1.1 Filtrar por Nivel 1 (EBITDA / NO EBITDA) si existe la columna
    if col_nivel1 and col_nivel1 in df_filtered.columns:
        series_n1 = df_filtered[col_nivel1].apply(normalize_text)
        mask_n1 = series_n1.isin(["ebitda", "no ebitda", "no_ebitda", "no-ebitda"])
        if mask_n1.any():
            df_filtered = df_filtered[mask_n1]

    # 2. Filtrar por Año y Rango YTD de Meses
    target_year = parse_year_to_num(year)
    target_month = parse_month_to_num(month)
    if target_month is None:
        target_month = 12
    
    if col_year and col_year in df_filtered.columns and target_year is not None:
        series_year = df_filtered[col_year].apply(parse_year_to_num)
        if series_year.notna().any():
            df_filtered = df_filtered[series_year == target_year]
        else:
            df_filtered = df_filtered[df_filtered[col_year].astype(str).str.contains(str(target_year))]

    if col_month and col_month in df_filtered.columns and target_month is not None:
        series_month = df_filtered[col_month].apply(parse_month_to_num)
        if series_month.notna().any():
            df_filtered = df_filtered[series_month.notna() & (series_month >= 1) & (series_month <= target_month)]

    # Asegurar montos numéricos
    if col_importe and col_importe in df_filtered.columns:
        df_filtered[col_importe] = pd.to_numeric(df_filtered[col_importe], errors='coerce').fillna(0.0)
    else:
        df_filtered[col_importe] = 0.0

    df_filtered[col_cuenta] = df_filtered[col_cuenta].astype(str).str.strip()
    df_filtered[col_nombre] = df_filtered[col_nombre].astype(str).str.strip()

    # 3. Definir categorías estándar del P&L
    if standard_categories is not None:
        STANDARD_CATEGORIES = list(standard_categories)
    else:
        STANDARD_CATEGORIES = [
            "Ingresos de arriendo fibra optica",
            "Ingresos de actividades ordinarias", 
            "Costo de ventas", 
            "Acceso a infraestructura fibra óptica",
            "Costos de uso fibra optica",
            "Depreciación operacional", 
            "Otros ingresos por función", 
            "Gastos de administración", 
            "Depreciación y amortizaciones", 
            "Otros egresos por función", 
            "Ingresos financieros", 
            "Ingresos financieros IC",
            "Costos financieros", 
            "Diferencias de cambio", 
            "Resultados por unidades de reajuste", 
            "Resultado por impuestos a las ganancias"
        ]

    # 4. Procesar el mapeo del maestro map_pl.xlsx
    map_database = {}
    if map_pl_df is not None and not map_pl_df.empty:
        col_map_cuenta = map_pl_df.columns[0]
        map_pl_df_copy = map_pl_df.copy()
        map_pl_df_copy[col_map_cuenta] = map_pl_df_copy[col_map_cuenta].astype(str).str.strip()
        
        # Pre-calcular el mapeo de columnas una sola vez
        normalized_cols = {normalize_text(c): c for c in map_pl_df_copy.columns}
        valid_cat_cols = [
            (cat, normalized_cols[normalize_text(cat)])
            for cat in STANDARD_CATEGORIES
            if normalize_text(cat) in normalized_cols
        ]
        
        cols_to_use = [col_map_cuenta] + [col for _, col in valid_cat_cols]
        for row in map_pl_df_copy[cols_to_use].itertuples(index=False):
            acc_id = str(row[0]).strip()
            row_map = {}
            for i, (cat, _) in enumerate(valid_cat_cols, start=1):
                val = row[i]
                if pd.notna(val) and str(val).strip() != "" and str(val).strip().lower() != "nan":
                    row_map[cat] = str(val).strip()
            map_database[acc_id] = row_map

    # 5. Normalizar mapeo de Odoo
    category_mapping = {
        "ingresos de actividades ordinarias": "Ingresos de actividades ordinarias",
        "costo de ventas": "Costo de ventas",
        "acceso a infraestructura fibra optica": "Acceso a infraestructura fibra óptica",
        "costos de uso fibra optica": "Costos de uso fibra optica",
        "depreciacion operacional": "Depreciación operacional",
        "otros ingresos por funcion": "Otros ingresos por función",
        "gastos de administracion": "Gastos de administración",
        "depreciacion y amortizaciones": "Depreciación y amortizaciones",
        "otros egresos por funcion": "Otros egresos por función",
        "ingresos financieros": "Ingresos financieros",
        "ingresos financiero": "Ingresos financieros",
        "ingresos financieros ic": "Ingresos financieros IC",
        "ingresos financieros con empresas relacionadas": "Ingresos financieros IC",
        "intereses con empresas relacionadas": "Ingresos financieros IC",
        "ingresos de arriendo fibra optica": "Ingresos de arriendo fibra optica",
        "costos financieros": "Costos financieros",
        "diferencia de cambio": "Diferencias de cambio",
        "diferencias de cambio": "Diferencias de cambio",
        "resultado por unidad de reajuste": "Resultados por unidades de reajuste",
        "resultados por unidades de reajuste": "Resultados por unidades de reajuste",
        "gastos por impuesto a las ganancias": "Resultado por impuestos a las ganancias",
        "resultado por impuestos a las ganancias": "Resultado por impuestos a las ganancias"
    }

    def map_odoo_category(raw_cat):
        norm = normalize_text(raw_cat)
        if not norm or norm in ['nan', 'none']:
            return None
        if norm in category_mapping:
            return category_mapping[norm]
        for k, v in category_mapping.items():
            if k in norm or norm in k:
                return v
        return None

    # Categorías específicas con prioridad de override
    SPECIFIC_CATEGORIES = [
        "Ingresos de arriendo fibra optica",
        "Acceso a infraestructura fibra óptica",
        "Costos de uso fibra optica",
        "Depreciación operacional",
        "Depreciación y amortizaciones",
        "Ingresos financieros",
        "Ingresos financieros IC",
        "Costos financieros",
        "Diferencias de cambio",
        "Resultados por unidades de reajuste",
        "Resultado por impuestos a las ganancias"
    ]

    # Mapeos forzados específicos requeridos por el usuario
    overrides = {
        "3105301": "Gastos de administración",
        "3105302": "Costo de ventas",
        "3105312": "Acceso a infraestructura fibra óptica",
        "3105702": "Depreciación y amortizaciones",
        "3105703": "Depreciación operacional",
        "3105711": "Depreciación y amortizaciones",
        "3105834": "Depreciación operacional",
        "3105835": "Depreciación operacional",
        "3108112": "Ingresos financieros IC",
        "3103111": "Ingresos de arriendo fibra optica",
        "3103113": "Ingresos de actividades ordinarias",
        "3103112": "Costo de ventas",
        "3103122": "Ingresos de actividades ordinarias",
        "3105704": "Depreciación operacional"
    }

    def resolve_category_and_motivo(cuenta_id, raw_odoo_cat, odoo_cat_mapped):
        if cuenta_id in overrides:
            return overrides[cuenta_id], ""
            
        final_cat = None
        motivo = ""

        # Si no está mapeado en el maestro, usar la clasificación directa de Odoo
        if cuenta_id not in map_database:
            if odoo_cat_mapped:
                final_cat = odoo_cat_mapped
            else:
                final_cat = "Otros egresos por función"
                motivo = "Sin clasificación en ERP y sin cuenta en map_pl (asignado a 'Otros egresos por función')"
        else:
            mappings = map_database[cuenta_id]
            
            # Desempate operacional/no operacional para depreciaciones mapeadas a ambos
            if "Depreciación operacional" in mappings and "Depreciación y amortizaciones" in mappings:
                if odoo_cat_mapped == "Costo de ventas":
                    final_cat = "Depreciación operacional"
                elif odoo_cat_mapped == "Gastos de administración":
                    final_cat = "Depreciación y amortizaciones"
                else:
                    final_cat = "Depreciación operacional"
                    motivo = "Sin clasificación en ERP (asignado por descarte a 'Depreciación operacional')"
            
            if final_cat is None:
                # Verificar prioridades de categorías específicas
                for spec_cat in SPECIFIC_CATEGORIES:
                    if spec_cat in mappings:
                        final_cat = spec_cat
                        break
                        
            if final_cat is None and odoo_cat_mapped and odoo_cat_mapped in mappings:
                # Usar la columna que coincida con la categoría mapeada de Odoo
                final_cat = odoo_cat_mapped
                
            if final_cat is None and mappings:
                # Tomar la primera clasificación disponible por orden estándar
                if len(mappings) == 1:
                    final_cat = list(mappings.keys())[0]
                    motivo = f"Sin clasificación en ERP (único rubro mapeado: '{final_cat}')"
                else:
                    for cat in STANDARD_CATEGORIES:
                        if cat in mappings:
                            final_cat = cat
                            motivo = f"Sin clasificación en ERP (asignado por orden prioritario a '{final_cat}')"
                            break
                            
            if final_cat is None:
                final_cat = odoo_cat_mapped if odoo_cat_mapped else "Otros egresos por función"
                motivo = f"Sin clasificación en ERP ni coincidencia (asignado a '{final_cat}')"

        return final_cat, motivo

    # Clasificar filas de forma vectorizada/agrupada
    df_audit = None
    if col_category and col_category in df_filtered.columns:
        unique_odoo = df_filtered[col_category].unique()
        odoo_map_cache = {raw: map_odoo_category(raw) for raw in unique_odoo}

        unique_combos = df_filtered[[col_cuenta, col_category]].drop_duplicates()
        combo_cat = {}
        combo_mot = {}
        for cuenta_val, cat_val in unique_combos.itertuples(index=False):
            acc_id = str(cuenta_val).strip()
            mapped_odoo = odoo_map_cache.get(cat_val)
            f_cat, f_mot = resolve_category_and_motivo(acc_id, cat_val, mapped_odoo)
            combo_cat[(cuenta_val, cat_val)] = f_cat
            combo_mot[(cuenta_val, cat_val)] = f_mot

        combo_keys = list(zip(df_filtered[col_cuenta], df_filtered[col_category]))
        df_filtered['mapped_category'] = [combo_cat.get(k) for k in combo_keys]

        if return_audit_log:
            raw_series = df_filtered[col_category].fillna("").astype(str).str.strip()
            odoo_mapped_series = pd.Series([odoo_map_cache.get(c) for c in df_filtered[col_category]], index=df_filtered.index)
            
            is_empty_or_unmapped = (
                (raw_series == "") | 
                (raw_series.str.lower().isin(['nan', 'none', '0', '0.0'])) | 
                (odoo_mapped_series.isna())
            )
            
            audit_df_rows = df_filtered[is_empty_or_unmapped]
            if not audit_df_rows.empty:
                audit_keys = [combo_keys[i] for i in range(len(combo_keys)) if is_empty_or_unmapped.iloc[i]]
                motivos = [combo_mot.get(k, "") for k in audit_keys]
                
                df_audit = pd.DataFrame({
                    'Fila Excel': audit_df_rows['_excel_row'].astype(int) if '_excel_row' in audit_df_rows.columns else 0,
                    'N° de Cuenta': audit_df_rows[col_cuenta].astype(str).str.strip(),
                    'Nombre de la cuenta': audit_df_rows[col_nombre].astype(str).str.strip(),
                    'Importe MN': audit_df_rows[col_importe].astype(float),
                    'Clasificación ERP': raw_series.loc[audit_df_rows.index].apply(lambda s: s if s else "(Vacío)"),
                    'Rubro Asignado': audit_df_rows['mapped_category'],
                    'Motivo / Observación': [m if m else f"Clasificación ERP vacía/no reconocida (asignado a '{c}')" for m, c in zip(motivos, audit_df_rows['mapped_category'])]
                })
            else:
                df_audit = pd.DataFrame(columns=[
                    'Fila Excel', 'N° de Cuenta', 'Nombre de la cuenta', 
                    'Importe MN', 'Clasificación ERP', 'Rubro Asignado', 'Motivo / Observación'
                ])
    else:
        unique_cuentas = df_filtered[col_cuenta].unique()
        cuenta_cat = {}
        for c in unique_cuentas:
            acc_id = str(c).strip()
            if acc_id in overrides:
                cuenta_cat[c] = overrides[acc_id]
            else:
                mappings = map_database.get(acc_id, {})
                cuenta_cat[c] = list(mappings.keys())[0] if mappings else "Otros egresos por función"
        df_filtered['mapped_category'] = df_filtered[col_cuenta].map(cuenta_cat)
        
        if return_audit_log:
            df_audit = pd.DataFrame({
                'Fila Excel': df_filtered['_excel_row'].astype(int) if '_excel_row' in df_filtered.columns else 0,
                'N° de Cuenta': df_filtered[col_cuenta].astype(str).str.strip(),
                'Nombre de la cuenta': df_filtered[col_nombre].astype(str).str.strip(),
                'Importe MN': df_filtered[col_importe].astype(float),
                'Clasificación ERP': "(Columna ausente)",
                'Rubro Asignado': df_filtered['mapped_category'],
                'Motivo / Observación': df_filtered['mapped_category'].apply(lambda cat: f"Columna informe_ee_rr no encontrada (asignado a '{cat}')")
            })

    # 6. Agrupar y pivotar a formato ancho
    df_grouped = df_filtered.groupby([col_cuenta, col_nombre, 'mapped_category'])[col_importe].sum().reset_index()

    df_pivot = df_grouped.pivot_table(
        index=[col_cuenta, col_nombre],
        columns='mapped_category',
        values=col_importe,
        aggfunc='sum',
        fill_value=0.0
    ).reset_index()

    # Formatear columnas finales
    df_pivot.rename(columns={col_cuenta: 'N° de cuenta', col_nombre: 'Nombre de la cuenta'}, inplace=True)
    df_pivot.columns.name = None

    # Inyectar columnas faltantes del estándar
    for cat in STANDARD_CATEGORIES:
        if cat not in df_pivot.columns:
            df_pivot[cat] = 0.0

    # Reordenar columnas exactamente al formato esperado
    df_pivot = df_pivot[['N° de cuenta', 'Nombre de la cuenta'] + STANDARD_CATEGORIES]
    
    if return_audit_log:
        if df_audit is None or df_audit.empty:
            df_audit = pd.DataFrame(columns=[
                'Fila Excel', 'N° de Cuenta', 'Nombre de la cuenta', 
                'Importe MN', 'Clasificación ERP', 'Rubro Asignado', 'Motivo / Observación'
            ])
        return df_pivot, df_audit
        
    return df_pivot
