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

def process_odoo_cubo(df_cubo, year, month, map_pl_df, standard_categories=None):
    """
    Procesa un DataFrame transaccional del Cubo de Odoo, filtra por periodo YTD,
    realiza la homologación de cuentas contra map_pl.xlsx y pivota los datos
    al formato P&L configurado.
    """
    # 1. Detectar nombres de columnas de forma dinámica
    col_year = next((c for c in df_cubo.columns if 'fec_doc' in c and 'mes' not in c.lower() and c != 'fec_doc'), None)
    col_month = next((c for c in df_cubo.columns if 'mes' in c.lower()), None)
    col_cuenta = next((c for c in df_cubo.columns if "cuenta" in c.lower() and "nombre" not in c.lower()), None)
    col_nombre = next((c for c in df_cubo.columns if "nombre_cuenta" in c.lower() or "nombre de la cuenta" in c.lower()), None)
    col_importe = next((c for c in df_cubo.columns if "importe" in c.lower() or "monto" in c.lower()), None)
    col_category = next((c for c in df_cubo.columns if "informe_eerr" in c.lower() or "eerr" in c.lower()), None)

    # Fallbacks de detección por posición por si fallan los nombres
    if not col_cuenta:
        col_cuenta = df_cubo.columns[2] if len(df_cubo.columns) > 2 else df_cubo.columns[0]
    if not col_nombre:
        col_nombre = df_cubo.columns[3] if len(df_cubo.columns) > 3 else df_cubo.columns[1]
    if not col_importe:
        col_importe = next((c for c in df_cubo.columns if df_cubo[c].dtype in [np.float64, np.int64]), df_cubo.columns[4])
    if not col_category:
        col_category = df_cubo.columns[6] if len(df_cubo.columns) > 6 else df_cubo.columns[0]

    # 2. Filtrar por Año y Rango YTD de Meses
    MONTH_LIST = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
    try:
        month_idx = int(month)
    except Exception:
        month_idx = 12
    allowed_months = MONTH_LIST[:month_idx]

    df_filtered = df_cubo.copy()
    
    if col_year and col_year in df_filtered.columns:
        df_filtered = df_filtered[df_filtered[col_year].astype(str).str.contains(str(year))]
    if col_month and col_month in df_filtered.columns:
        df_filtered = df_filtered[df_filtered[col_month].astype(str).str.strip().str.lower().isin(allowed_months)]

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
        
        for idx, row in map_pl_df_copy.iterrows():
            acc_id = row[col_map_cuenta]
            map_database[acc_id] = {}
            for cat in STANDARD_CATEGORIES:
                map_col = next((c for c in map_pl_df_copy.columns if normalize_text(c) == normalize_text(cat)), None)
                if map_col:
                    val = row[map_col]
                    if pd.notna(val) and str(val).strip() != "" and str(val).strip().lower() != "nan":
                        map_database[acc_id][cat] = str(val).strip()

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

    def determine_final_category(cuenta_id, raw_odoo_cat):
        cuenta_id = str(cuenta_id).strip()
        
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
        if cuenta_id in overrides:
            return overrides[cuenta_id]
            
        odoo_cat_mapped = map_odoo_category(raw_odoo_cat)
        
        # Si no está mapeado en el maestro, usar la clasificación directa de Odoo
        if cuenta_id not in map_database:
            return odoo_cat_mapped if odoo_cat_mapped else "Otros egresos por función"
            
        mappings = map_database[cuenta_id]
        
        # Desempate operacional/no operacional para depreciaciones mapeadas a ambos
        if "Depreciación operacional" in mappings and "Depreciación y amortizaciones" in mappings:
            if odoo_cat_mapped == "Costo de ventas":
                return "Depreciación operacional"
            elif odoo_cat_mapped == "Gastos de administración":
                return "Depreciación y amortizaciones"
        
        # Verificar prioridades de categorías específicas
        for spec_cat in SPECIFIC_CATEGORIES:
            if spec_cat in mappings:
                return spec_cat
                
        # Usar la columna que coincida con la categoría mapeada de Odoo
        if odoo_cat_mapped and odoo_cat_mapped in mappings:
            return odoo_cat_mapped
            
        # Tomar la primera clasificación disponible
        if mappings:
            if len(mappings) == 1:
                return list(mappings.keys())[0]
            for cat in STANDARD_CATEGORIES:
                if cat in mappings:
                    return cat
                    
        return odoo_cat_mapped if odoo_cat_mapped else "Otros egresos por función"

    # Aplicar clasificación a cada fila
    if col_category and col_category in df_filtered.columns:
        df_filtered['mapped_category'] = df_filtered.apply(
            lambda r: determine_final_category(r[col_cuenta], r[col_category]), axis=1
        )
    else:
        def fallback_map(acc_id):
            mappings = map_database.get(acc_id, {})
            if mappings:
                return list(mappings.keys())[0]
            return "Otros egresos por función"
        df_filtered['mapped_category'] = df_filtered[col_cuenta].apply(fallback_map)

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
    return df_pivot
