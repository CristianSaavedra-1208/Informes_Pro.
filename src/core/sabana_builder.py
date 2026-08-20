import pandas as pd
from src.core.excel_utils import sort_accounts

def build_balance_sabana(tb_df: pd.DataFrame, map_balance_df: pd.DataFrame) -> pd.DataFrame:
    """
    Combina los saldos del Trial Balance con las clasificaciones del Mapeo de Balance.
    Realiza un Left Join para asegurar que todas las cuentas del Trial Balance queden
    representadas, incluso si no tienen clasificación asignada (huérfanas).
    """
    if tb_df is None or tb_df.empty:
        return pd.DataFrame()
        
    tb_df = tb_df.copy()
    
    # Identificar columna de cuenta en TB (coincidencia exacta con original)
    tb_acc_cols = ['N° de Cuenta \n', 'N de Cuenta', 'N° de Cuenta', 'Cuenta', 'cuenta_id']
    tb_cuenta_col = next((c for c in tb_df.columns if str(c).strip() in [ac.strip() for ac in tb_acc_cols]), None)
    if not tb_cuenta_col:
        tb_cuenta_col = tb_df.columns[0]
        
    tb_df[tb_cuenta_col] = tb_df[tb_cuenta_col].astype(str).str.strip()
    
    if map_balance_df is None or map_balance_df.empty:
        tb_df = sort_accounts(tb_df, tb_cuenta_col)
        return tb_df
        
    map_balance_df = map_balance_df.copy()
    
    # Identificar columna de cuenta en Mapeo de Balance (coincidencia exacta con original)
    map_acc_cols = ['N° de Cuenta \n', 'N de Cuenta', 'N° de Cuenta', 'Cuenta', 'cuenta_id']
    map_cuenta_col = next((c for c in map_balance_df.columns if str(c).strip() in [ac.strip() for ac in map_acc_cols]), None)
    if not map_cuenta_col:
        map_cuenta_col = map_balance_df.columns[0]
        
    map_balance_df[map_cuenta_col] = map_balance_df[map_cuenta_col].astype(str).str.strip()
    
    # Realizar cruce de datos
    merged = pd.merge(tb_df, map_balance_df, left_on=tb_cuenta_col, right_on=map_cuenta_col, how='left')
    
    # Ordenar numéricamente por código de cuenta
    merged = sort_accounts(merged, tb_cuenta_col)
    
    return merged

def build_pl_sabana(pl_df: pd.DataFrame, map_pl_df: pd.DataFrame, tb_df: pd.DataFrame = None) -> pd.DataFrame:
    """
    Combina los saldos del P&L Cubo (o del Trial Balance) con el Mapeo de P&L.
    Su base es el Mapeo de P&L filtrado por cuentas de P&L (excluyendo cuentas de balance
    que tengan mapeos vacíos), cruzando la data del Cubo (pl_df) y, si existe,
    el Trial Balance (tb_df) solo para los saldos contables de estas cuentas.
    """
    if map_pl_df is None or map_pl_df.empty:
        return pl_df.copy() if pl_df is not None else pd.DataFrame()
        
    map_pl_df = map_pl_df.copy()
    
    # Identificar columna de cuenta en Mapeo de P&L
    map_acc_cols = ['N° de Cuenta \n', 'N de Cuenta', 'N° de Cuenta', 'Cuenta', 'cuenta_id']
    map_cuenta_col = next((c for c in map_pl_df.columns if str(c).strip() in [ac.strip() for ac in map_acc_cols]), None)
    if not map_cuenta_col:
        map_cuenta_col = map_pl_df.columns[0]
    map_pl_df[map_cuenta_col] = map_pl_df[map_cuenta_col].astype(str).str.strip()
    
    # Filtrar map_pl_df para eliminar cuentas de balance (que tienen todas las columnas de P&L vacías/NaN)
    exclude_from_mapping = ['cuenta', 'detalle', 'nombre', 'descripcion', 'flujo', 'unnamed']
    mapping_cols_to_check = [
        c for c in map_pl_df.columns 
        if not any(x in c.lower() for x in exclude_from_mapping)
    ]
    if mapping_cols_to_check:
        has_valid_mapping = map_pl_df[mapping_cols_to_check].apply(
            lambda row: any(str(val).strip().lower() not in ('nan', 'none', '') for val in row), axis=1
        )
        map_pl_df = map_pl_df[has_valid_mapping].copy()
        
    # 1. Base es el Mapeo de P&L. Si hay P&L Cubo mensual (pl_df), los cruzamos.
    if pl_df is not None and not pl_df.empty:
        pl_df = pl_df.copy()
        pl_cuenta_col = next((c for c in pl_df.columns if "cuenta" in str(c).lower() and "nombre" not in str(c).lower()), None)
        if not pl_cuenta_col:
            pl_cuenta_col = pl_df.columns[0]
        pl_df[pl_cuenta_col] = pl_df[pl_cuenta_col].astype(str).str.strip()
        
        # Dropear columna de nombre del cubo
        pl_name_col = next((c for c in pl_df.columns if "nombre" in str(c).lower()), None)
        if pl_name_col:
            pl_df.drop(columns=[pl_name_col], inplace=True, errors='ignore')
            
        # Identificar y renombrar conflictos
        conflict_cols = [c for c in pl_df.columns if c in map_pl_df.columns and c != pl_cuenta_col and c != map_cuenta_col]
        if conflict_cols:
            rename_dict = {c: f"{c} (Saldo)" for c in conflict_cols}
            pl_df.rename(columns=rename_dict, inplace=True)
            
        base_df = pd.merge(map_pl_df, pl_df, left_on=map_cuenta_col, right_on=pl_cuenta_col, how='outer')
        
        # Rellenar la llave del mapa si el cubo trajo cuentas nuevas no mapeadas
        if map_cuenta_col != pl_cuenta_col and pl_cuenta_col in base_df.columns:
            base_df[map_cuenta_col] = base_df[map_cuenta_col].fillna(base_df[pl_cuenta_col])
    else:
        base_df = map_pl_df
        
    # 2. Si se proporciona tb_df, le cruzamos las columnas de saldos del ERP (saldo_inicial, debitos, creditos, saldo_final)
    # pero usando left join sobre la base de P&L, para no meter cuentas de balance.
    if tb_df is not None and not tb_df.empty:
        tb_df = tb_df.copy()
        tb_acc_cols = ['N° de Cuenta \n', 'N de Cuenta', 'N° de Cuenta', 'Cuenta', 'cuenta_id']
        tb_cuenta_col = next((c for c in tb_df.columns if str(c).strip() in [ac.strip() for ac in tb_acc_cols]), None)
        if not tb_cuenta_col:
            tb_cuenta_col = tb_df.columns[0]
        tb_df[tb_cuenta_col] = tb_df[tb_cuenta_col].astype(str).str.strip()
        
        # Identificar columna de nombre/descripción de la cuenta en el TB
        tb_nombre_cols = ['Nombre de la cuenta', 'Nombre cuenta', 'Nombre', 'descripcion', 'Detalle']
        tb_nombre_col = next((c for c in tb_df.columns if str(c).strip() in [nc.strip() for nc in tb_nombre_cols]), None)
        
        # Seleccionamos las columnas del TB para traer
        tb_saldo_cols = [tb_cuenta_col, 'saldo_inicial', 'debitos', 'creditos', 'saldo_final']
        if tb_nombre_col:
            tb_saldo_cols.append(tb_nombre_col)
            
        tb_cols_to_merge = [c for c in tb_saldo_cols if c in tb_df.columns]
        tb_subset = tb_df[tb_cols_to_merge].copy()
        
        base_df = pd.merge(base_df, tb_subset, left_on=map_cuenta_col, right_on=tb_cuenta_col, how='left')
        
    base_df = sort_accounts(base_df, map_cuenta_col)
    
    # 3. Calcular Totales para cada columna numérica
    totales = {}
    totales[map_cuenta_col] = "TOTAL"
    
    tb_nombre_cols = ['Nombre de la cuenta', 'Nombre cuenta', 'Nombre', 'descripcion', 'Detalle']
    desc_col = next((c for c in base_df.columns if str(c).strip() in [nc.strip() for nc in tb_nombre_cols] or "detalle" in str(c).lower()), None)
    if desc_col:
        totales[desc_col] = "Total General"
        
    for col in base_df.columns:
        if col in totales:
            continue
        # Evitar sumar llaves de cuentas contables o columnas de clasificación
        col_clean = str(col).lower()
        if any(x in col_clean for x in ['cuenta', 'flujo', 'unnamed', 'class', 'clasificac']):
            totales[col] = ""
            continue
            
        numeric_vals = pd.to_numeric(base_df[col], errors='coerce')
        if not numeric_vals.isna().all():
            totales[col] = float(numeric_vals.sum())
        else:
            totales[col] = ""
            
    totales_df = pd.DataFrame([totales])
    base_df = pd.concat([base_df, totales_df], ignore_index=True)
    
    return base_df


def get_group_companies(grupo_name: str):
    """
    Obtiene la lista de nombres de empresas filiales y matriz pertenecientes a un grupo.
    """
    from src.models.database import SessionLocal
    from src.models.consolidacion import ConsolidationGroup
    
    clean_group = grupo_name.replace("[GRUPO] ", "").strip()
    db = SessionLocal()
    companies = []
    try:
        grupo_obj = db.query(ConsolidationGroup).filter_by(nombre_grupo=clean_group).first()
        if grupo_obj:
            companies.append(grupo_obj.empresa_matriz)
            if grupo_obj.filial_is_group:
                def get_sub_companies(sub_g_id):
                    sub_g = db.query(ConsolidationGroup).filter_by(id=sub_g_id).first()
                    if sub_g:
                        c = [sub_g.empresa_matriz]
                        if sub_g.filial_is_group:
                            c.extend(get_sub_companies(int(sub_g.empresa_filial)))
                        else:
                            c.append(sub_g.empresa_filial)
                        return c
                    return []
                companies.extend(get_sub_companies(int(grupo_obj.empresa_filial)))
            else:
                companies.append(grupo_obj.empresa_filial)
    finally:
        db.close()
    return sorted(list(dict.fromkeys(companies)))


def build_consolidated_balance_sabana(grupo_name: str, periodo: str, map_balance_df: pd.DataFrame = None) -> pd.DataFrame:
    """
    Construye la sábana de auditoría consolidada para el Balance General.
    Muestra cada filial en una columna independiente junto al total consolidado.
    """
    from src.models.trial_balance_db import TrialBalanceDB
    
    companies = get_group_companies(grupo_name)
    if not companies:
        return pd.DataFrame()
        
    all_accounts = {}
    
    for co in companies:
        tb_df = TrialBalanceDB.get_trial_balance(co, periodo)
        if tb_df is not None and not tb_df.empty:
            tb_acc_cols = ['N° de Cuenta \n', 'N de Cuenta', 'N° de Cuenta', 'Cuenta', 'cuenta_id']
            cuenta_col = next((c for c in tb_df.columns if str(c).strip() in [ac.strip() for ac in tb_acc_cols]), tb_df.columns[0])
            
            tb_nombre_cols = ['Nombre de la cuenta', 'Nombre cuenta', 'Nombre', 'descripcion', 'Detalle']
            nombre_col = next((c for c in tb_df.columns if str(c).strip() in [nc.strip() for nc in tb_nombre_cols]), None)
            
            saldo_col = 'saldo_final' if 'saldo_final' in tb_df.columns else tb_df.columns[-1]
            
            for _, row in tb_df.iterrows():
                acct_id = str(row[cuenta_col]).strip()
                if not acct_id or acct_id.lower() in ('nan', 'none', ''):
                    continue
                acct_name = str(row[nombre_col]).strip() if nombre_col and pd.notna(row[nombre_col]) else ""
                val = pd.to_numeric(row[saldo_col], errors='coerce')
                val = float(val) if pd.notna(val) else 0.0
                
                if acct_id not in all_accounts:
                    all_accounts[acct_id] = {'N° de Cuenta': acct_id, 'Nombre de la Cuenta': acct_name}
                elif not all_accounts[acct_id]['Nombre de la Cuenta'] and acct_name:
                    all_accounts[acct_id]['Nombre de la Cuenta'] = acct_name
                    
                all_accounts[acct_id][co] = all_accounts[acct_id].get(co, 0.0) + val
                
    if not all_accounts:
        return pd.DataFrame()
        
    df_res = pd.DataFrame(list(all_accounts.values()))
    
    for co in companies:
        if co not in df_res.columns:
            df_res[co] = 0.0
        else:
            df_res[co] = df_res[co].fillna(0.0)
            
    df_res['TOTAL CONSOLIDADO'] = df_res[companies].sum(axis=1)
    
    if map_balance_df is not None and not map_balance_df.empty:
        map_acc_cols = ['N° de Cuenta \n', 'N de Cuenta', 'N° de Cuenta', 'Cuenta', 'cuenta_id']
        map_col = next((c for c in map_balance_df.columns if str(c).strip() in [ac.strip() for ac in map_acc_cols]), map_balance_df.columns[0])
        map_balance_df = map_balance_df.copy()
        map_balance_df[map_col] = map_balance_df[map_col].astype(str).str.strip()
        
        map_cols_to_use = [c for c in map_balance_df.columns if c not in ('Nombre de la cuenta', 'Nombre cuenta', 'Nombre', 'descripcion') or c == map_col]
        map_sub = map_balance_df[map_cols_to_use]
        
        df_res = pd.merge(df_res, map_sub, left_on='N° de Cuenta', right_on=map_col, how='left')
        if map_col != 'N° de Cuenta' and map_col in df_res.columns:
            df_res.drop(columns=[map_col], inplace=True)
            
    df_res = sort_accounts(df_res, 'N° de Cuenta')
    
    totales = {'N° de Cuenta': 'TOTAL', 'Nombre de la Cuenta': 'Total General'}
    for col in df_res.columns:
        if col in ('N° de Cuenta', 'Nombre de la Cuenta'):
            continue
        numeric_vals = pd.to_numeric(df_res[col], errors='coerce')
        if not numeric_vals.isna().all():
            totales[col] = float(numeric_vals.sum())
        else:
            totales[col] = ""
            
    df_res = pd.concat([df_res, pd.DataFrame([totales])], ignore_index=True)
    return df_res


def build_consolidated_pl_sabana(grupo_name: str, periodo: str, map_pl_df: pd.DataFrame = None) -> pd.DataFrame:
    """
    Construye la sábana de auditoría consolidada para el Estado de Resultados / P&L.
    Muestra cada filial en una columna independiente junto al total consolidado.
    """
    from src.models.pl_cubo_db import PlCuboDB
    from src.models.trial_balance_db import TrialBalanceDB
    
    companies = get_group_companies(grupo_name)
    if not companies:
        return pd.DataFrame()
        
    all_accounts = {}
    
    for co in companies:
        pl_df = PlCuboDB.get_pl_cubo(co, periodo)
        if pl_df is not None and not pl_df.empty:
            cuenta_col = next((c for c in pl_df.columns if "cuenta" in str(c).lower() and "nombre" not in str(c).lower()), pl_df.columns[0])
            nombre_col = next((c for c in pl_df.columns if "nombre" in str(c).lower() or "desc" in str(c).lower()), None)
            numeric_cols = [c for c in pl_df.columns if c not in (cuenta_col, nombre_col) and not any(x in str(c).lower() for x in ['cuenta', 'nombre', 'unnamed', 'detalle'])]
            
            for _, row in pl_df.iterrows():
                acct_id = str(row[cuenta_col]).strip()
                if not acct_id or acct_id.lower() in ('nan', 'none', ''):
                    continue
                acct_name = str(row[nombre_col]).strip() if nombre_col and pd.notna(row[nombre_col]) else ""
                
                row_sum = 0.0
                for nc in numeric_cols:
                    v = pd.to_numeric(row[nc], errors='coerce')
                    if pd.notna(v):
                        row_sum += float(v)
                        
                if acct_id not in all_accounts:
                    all_accounts[acct_id] = {'N° de Cuenta': acct_id, 'Nombre de la Cuenta': acct_name}
                elif not all_accounts[acct_id]['Nombre de la Cuenta'] and acct_name:
                    all_accounts[acct_id]['Nombre de la Cuenta'] = acct_name
                    
                all_accounts[acct_id][co] = all_accounts[acct_id].get(co, 0.0) + row_sum
        else:
            tb_df = TrialBalanceDB.get_trial_balance(co, periodo)
            if tb_df is not None and not tb_df.empty:
                tb_acc_cols = ['N° de Cuenta \n', 'N de Cuenta', 'N° de Cuenta', 'Cuenta', 'cuenta_id']
                cuenta_col = next((c for c in tb_df.columns if str(c).strip() in [ac.strip() for ac in tb_acc_cols]), tb_df.columns[0])
                tb_nombre_cols = ['Nombre de la cuenta', 'Nombre cuenta', 'Nombre', 'descripcion', 'Detalle']
                nombre_col = next((c for c in tb_df.columns if str(c).strip() in [nc.strip() for nc in tb_nombre_cols]), None)
                saldo_col = 'saldo_final' if 'saldo_final' in tb_df.columns else tb_df.columns[-1]
                
                for _, row in tb_df.iterrows():
                    acct_id = str(row[cuenta_col]).strip()
                    if not acct_id or acct_id.lower() in ('nan', 'none', ''):
                        continue
                    acct_name = str(row[nombre_col]).strip() if nombre_col and pd.notna(row[nombre_col]) else ""
                    val = pd.to_numeric(row[saldo_col], errors='coerce')
                    val = float(val) if pd.notna(val) else 0.0
                    
                    if acct_id not in all_accounts:
                        all_accounts[acct_id] = {'N° de Cuenta': acct_id, 'Nombre de la Cuenta': acct_name}
                    elif not all_accounts[acct_id]['Nombre de la Cuenta'] and acct_name:
                        all_accounts[acct_id]['Nombre de la Cuenta'] = acct_name
                        
                    all_accounts[acct_id][co] = all_accounts[acct_id].get(co, 0.0) + val

    if not all_accounts:
        return pd.DataFrame()
        
    df_res = pd.DataFrame(list(all_accounts.values()))
    
    for co in companies:
        if co not in df_res.columns:
            df_res[co] = 0.0
        else:
            df_res[co] = df_res[co].fillna(0.0)
            
    df_res['TOTAL CONSOLIDADO'] = df_res[companies].sum(axis=1)
    
    if map_pl_df is not None and not map_pl_df.empty:
        map_acc_cols = ['N° de Cuenta \n', 'N de Cuenta', 'N° de Cuenta', 'Cuenta', 'cuenta_id']
        map_col = next((c for c in map_pl_df.columns if str(c).strip() in [ac.strip() for ac in map_acc_cols]), map_pl_df.columns[0])
        map_pl_df = map_pl_df.copy()
        map_pl_df[map_col] = map_pl_df[map_col].astype(str).str.strip()
        
        map_cols_to_use = [c for c in map_pl_df.columns if c not in ('Nombre de la cuenta', 'Nombre cuenta', 'Nombre', 'descripcion') or c == map_col]
        map_sub = map_pl_df[map_cols_to_use]
        
        df_res = pd.merge(df_res, map_sub, left_on='N° de Cuenta', right_on=map_col, how='left')
        if map_col != 'N° de Cuenta' and map_col in df_res.columns:
            df_res.drop(columns=[map_col], inplace=True)
            
    df_res = sort_accounts(df_res, 'N° de Cuenta')
    
    totales = {'N° de Cuenta': 'TOTAL', 'Nombre de la Cuenta': 'Total General'}
    for col in df_res.columns:
        if col in ('N° de Cuenta', 'Nombre de la Cuenta'):
            continue
        numeric_vals = pd.to_numeric(df_res[col], errors='coerce')
        if not numeric_vals.isna().all():
            totales[col] = float(numeric_vals.sum())
        else:
            totales[col] = ""
            
    df_res = pd.concat([df_res, pd.DataFrame([totales])], ignore_index=True)
    return df_res
