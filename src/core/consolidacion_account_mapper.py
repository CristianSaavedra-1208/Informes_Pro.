import os
import re
import unicodedata
import pandas as pd
from typing import Dict, List, Optional, Any

# Base path resolution
BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DATA_EMPRESAS_DIR = os.path.join(BASE_DIR, 'data', 'empresas')

def normalize_text(t: Any) -> str:
    """
    Normaliza texto eliminando acentos, espacios duplicados y pasando a minúsculas.
    """
    if t is None or pd.isna(t):
        return ''
    text_str = str(t).strip().lower()
    return re.sub(
        r'\s+', ' ',
        ''.join(c for c in unicodedata.normalize('NFD', text_str) if unicodedata.category(c) != 'Mn')
    )

def _resolve_empresa_dir(empresa: Optional[str] = None) -> Optional[str]:
    """
    Encuentra el directorio de empresa que contiene los archivos de mapeo.
    """
    if empresa:
        direct_path = os.path.join(DATA_EMPRESAS_DIR, empresa)
        if os.path.exists(direct_path):
            return direct_path
            
    # Intentar con empresas por defecto si no se encuentra
    fallbacks = ['Pacifico Cable SpA', 'Db Terra Chile Holdco SpA', 'Db Terra Chile Parent SpA']
    for fb in fallbacks:
        fb_path = os.path.join(DATA_EMPRESAS_DIR, fb)
        if os.path.exists(fb_path):
            return fb_path
            
    # Si ninguna de las anteriores, buscar el primer directorio en data/empresas que tenga map_balance.xlsx
    if os.path.exists(DATA_EMPRESAS_DIR):
        for item in os.listdir(DATA_EMPRESAS_DIR):
            item_path = os.path.join(DATA_EMPRESAS_DIR, item)
            if os.path.isdir(item_path) and os.path.exists(os.path.join(item_path, 'map_balance.xlsx')):
                return item_path
                
    return None

def load_consolidation_account_map(empresa: Optional[str] = None) -> Dict[str, Dict[str, Any]]:
    """
    Carga todas las cuentas de consolidación (serie 99...) desde map_balance.xlsx y map_pl.xlsx.
    Retorna un diccionario indexado por código de cuenta (ej. '9999050').
    """
    empresa_dir = _resolve_empresa_dir(empresa)
    if not empresa_dir:
        return {}

    account_map: Dict[str, Dict[str, Any]] = {}

    # 1. Cargar map_balance.xlsx
    path_balance = os.path.join(empresa_dir, 'map_balance.xlsx')
    if os.path.exists(path_balance):
        try:
            df_b = pd.read_excel(path_balance)
            col_cuenta = df_b.columns[0]
            col_nombre = df_b.columns[1] if len(df_b.columns) > 1 else col_cuenta
            col_clasif = df_b.columns[2] if len(df_b.columns) > 2 else None
            col_nota = df_b.columns[3] if len(df_b.columns) > 3 else None

            for _, row in df_b.iterrows():
                val_cuenta = row[col_cuenta]
                if pd.isna(val_cuenta):
                    continue
                codigo = str(val_cuenta).strip()
                if not codigo.startswith('99'):
                    continue

                nombre_raw = str(row[col_nombre]).strip() if pd.notna(row[col_nombre]) else ''
                clasif_raw = str(row[col_clasif]).strip() if col_clasif and pd.notna(row[col_clasif]) else ''
                nota_raw = str(row[col_nota]).strip() if col_nota and pd.notna(row[col_nota]) else None

                if nota_raw and (nota_raw.lower() == 'nan' or nota_raw == ''):
                    nota_raw = None

                if clasif_raw and clasif_raw.lower() != 'nan':
                    lbl_nota = f" [Nota: {nota_raw}]" if nota_raw else ""
                    display_label = f"{codigo} — {clasif_raw}{lbl_nota}"
                    account_map[codigo] = {
                        "codigo": codigo,
                        "nombre_cuenta": nombre_raw,
                        "linea_item": clasif_raw,
                        "linea_nota": nota_raw,
                        "tipo": "Balance",
                        "display_label": display_label
                    }
        except Exception as e:
            print(f"Error cargando map_balance.xlsx en {empresa_dir}: {e}")

    # 2. Cargar map_pl.xlsx
    path_pl = os.path.join(empresa_dir, 'map_pl.xlsx')
    if os.path.exists(path_pl):
        try:
            df_pl = pd.read_excel(path_pl)
            col_cuenta = df_pl.columns[0]
            col_nombre = df_pl.columns[1] if len(df_pl.columns) > 1 else col_cuenta
            rubro_cols = list(df_pl.columns[2:])

            for _, row in df_pl.iterrows():
                val_cuenta = row[col_cuenta]
                if pd.isna(val_cuenta):
                    continue
                codigo = str(val_cuenta).strip()
                if not codigo.startswith('99'):
                    continue

                nombre_raw = str(row[col_nombre]).strip() if pd.notna(row[col_nombre]) else ''

                for rc in rubro_cols:
                    val_nota = row[rc]
                    if pd.notna(val_nota) and str(val_nota).strip() not in ['', 'nan']:
                        linea_item = str(rc).strip()
                        linea_nota = str(val_nota).strip()
                        if linea_nota.lower() == 'nan' or linea_nota == '':
                            linea_nota = None

                        lbl_nota = f" [Nota: {linea_nota}]" if linea_nota else ""
                        display_label = f"{codigo} — {linea_item}{lbl_nota}"
                        account_map[codigo] = {
                            "codigo": codigo,
                            "nombre_cuenta": nombre_raw,
                            "linea_item": linea_item,
                            "linea_nota": linea_nota,
                            "tipo": "P&L",
                            "display_label": display_label
                        }
                        break
        except Exception as e:
            print(f"Error cargando map_pl.xlsx en {empresa_dir}: {e}")

    return account_map

def get_consolidation_accounts_list(empresa: Optional[str] = None) -> List[Dict[str, Any]]:
    """
    Retorna la lista ordenada de cuentas de consolidación para usar en selectores de UI.
    """
    account_map = load_consolidation_account_map(empresa)
    accounts = list(account_map.values())
    accounts.sort(key=lambda a: a["codigo"])
    return accounts

def get_account_info(codigo: str, empresa: Optional[str] = None) -> Optional[Dict[str, Any]]:
    """
    Obtiene la información mapeada para un código de cuenta de consolidación.
    """
    if not codigo:
        return None
    account_map = load_consolidation_account_map(empresa)
    return account_map.get(str(codigo).strip())

def _stem_tokens(text: str) -> set:
    norm = normalize_text(text)
    stop_words = {'de', 'a', 'la', 'el', 'los', 'las', 'y', 'o', 'en', 'por', 'del', 'al'}
    tokens = set()
    for w in norm.split():
        if w not in stop_words and len(w) > 1:
            tokens.add(re.sub(r's$', '', w) if len(w) > 3 else w)
    return tokens

def find_account_by_rubro_nota(linea_item: str, linea_nota: Optional[str] = None, empresa: Optional[str] = None) -> Optional[str]:
    """
    Búsqueda inversa: dado un linea_item y opcionalmente linea_nota, busca el código 9999xxx correspondiente.
    """
    if not linea_item:
        return None
        
    norm_item = normalize_text(linea_item)
    norm_nota = normalize_text(linea_nota) if linea_nota else None

    account_map = load_consolidation_account_map(empresa)
    
    # 1. Búsqueda exacta normalizada
    candidates = []
    for cod, info in account_map.items():
        cand_item = normalize_text(info.get('linea_item'))
        if cand_item == norm_item:
            cand_nota = normalize_text(info.get('linea_nota')) if info.get('linea_nota') else None
            if norm_nota and cand_nota == norm_nota:
                return cod
            candidates.append((cod, cand_nota))

    if candidates:
        candidates.sort(key=lambda x: (0 if (norm_nota and x[1] == norm_nota) else (1 if x[1] is None else 2), x[0]))
        return candidates[0][0]

    # 2. Búsqueda por similitud de tokens (manejando singular/plural y stopwords)
    target_tokens = _stem_tokens(linea_item)
    if not target_tokens:
        return None

    scored_cands = []
    for cod, info in account_map.items():
        cand_tokens = _stem_tokens(info.get('linea_item', ''))
        if cand_tokens and target_tokens:
            intersection = target_tokens.intersection(cand_tokens)
            union = target_tokens.union(cand_tokens)
            jaccard = len(intersection) / len(union) if union else 0.0
            
            # Si coinciden casi todos los tokens principales
            if jaccard >= 0.7 or (intersection == target_tokens and intersection == cand_tokens):
                cand_nota = normalize_text(info.get('linea_nota')) if info.get('linea_nota') else None
                nota_match = (norm_nota and cand_nota == norm_nota)
                score = jaccard + (1.0 if nota_match else 0.0)
                scored_cands.append((score, cod, cand_nota))

    if scored_cands:
        scored_cands.sort(key=lambda x: (-x[0], x[1]))
        return scored_cands[0][1]

    return None

