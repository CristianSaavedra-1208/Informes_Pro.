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

import copy

_ACCOUNT_MAP_CACHE: Dict[tuple, Dict[str, Dict[str, Any]]] = {}
_RUBRO_LOOKUP_CACHE: Dict[tuple, Optional[str]] = {}

def _get_mapping_cache_key(empresa_dir: str) -> tuple:
    path_balance = os.path.join(empresa_dir, 'map_balance.xlsx')
    path_pl = os.path.join(empresa_dir, 'map_pl.xlsx')
    mtime_b = os.path.getmtime(path_balance) if os.path.exists(path_balance) else 0.0
    mtime_p = os.path.getmtime(path_pl) if os.path.exists(path_pl) else 0.0
    return (empresa_dir, mtime_b, mtime_p)

def load_consolidation_account_map(empresa: Optional[str] = None) -> Dict[str, Dict[str, Any]]:
    """
    Carga todas las cuentas de consolidación (serie 99...) desde map_balance.xlsx y map_pl.xlsx.
    Retorna un diccionario indexado por código de cuenta (ej. '9999050').
    Usa caché en memoria invalidado automáticamente por mtime de los archivos Excel.
    """
    empresa_dir = _resolve_empresa_dir(empresa)
    if not empresa_dir:
        return {}

    cache_key = _get_mapping_cache_key(empresa_dir)
    if cache_key in _ACCOUNT_MAP_CACHE:
        return copy.deepcopy(_ACCOUNT_MAP_CACHE[cache_key])

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

    _ACCOUNT_MAP_CACHE[cache_key] = account_map
    return copy.deepcopy(account_map)

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

def find_account_by_rubro_nota(
    linea_item: str,
    linea_nota: Optional[str] = None,
    empresa: Optional[str] = None,
    account_map: Optional[Dict[str, Dict[str, Any]]] = None,
) -> Optional[str]:
    """
    Búsqueda inversa: dado un linea_item y opcionalmente linea_nota, busca el código 9999xxx correspondiente.
    """
    if not linea_item:
        return None
        
    norm_item = normalize_text(linea_item)
    norm_nota = normalize_text(linea_nota) if linea_nota else None

    cache_lookup_key = None
    if account_map is None:
        empresa_dir = _resolve_empresa_dir(empresa)
        if not empresa_dir:
            return None
        map_key = _get_mapping_cache_key(empresa_dir)
        cache_lookup_key = (map_key, norm_item, norm_nota)
        if cache_lookup_key in _RUBRO_LOOKUP_CACHE:
            return _RUBRO_LOOKUP_CACHE[cache_lookup_key]
        if map_key not in _ACCOUNT_MAP_CACHE:
            load_consolidation_account_map(empresa)
        account_map = _ACCOUNT_MAP_CACHE.get(map_key, {})
    
    # 0. Alias históricos explícitos (asientos previos al catálogo 9999xxx)
    HISTORICAL_ITEM_ALIASES = {
        "costos de uso fibra optica": "9999315",
        "costo de uso fibra optica": "9999315",
        "ganancias (perdida) del ejercicio": "9999345",
        "ganancia (perdida) del ejercicio": "9999345",
    }
    if norm_item in HISTORICAL_ITEM_ALIASES and HISTORICAL_ITEM_ALIASES[norm_item] in account_map:
        res_alias = HISTORICAL_ITEM_ALIASES[norm_item]
        if cache_lookup_key is not None:
            _RUBRO_LOOKUP_CACHE[cache_lookup_key] = res_alias
        return res_alias

    # 1. Búsqueda exacta normalizada (por rubro o por desglose de nota cuando el histórico guardó la nota en linea_item)
    candidates = []
    nota_as_item_candidates = []
    for cod, info in account_map.items():
        cand_item = normalize_text(info.get('linea_item'))
        cand_nota = normalize_text(info.get('linea_nota')) if info.get('linea_nota') else None
        if cand_item == norm_item:
            if norm_nota and cand_nota == norm_nota:
                if cache_lookup_key is not None:
                    _RUBRO_LOOKUP_CACHE[cache_lookup_key] = cod
                return cod
            candidates.append((cod, cand_nota))
        elif cand_nota and cand_nota == norm_item:
            nota_as_item_candidates.append(cod)

    if candidates:
        candidates.sort(key=lambda x: (0 if (norm_nota and x[1] == norm_nota) else (1 if x[1] is None else 2), x[0]))
        res_cod = candidates[0][0]
        if cache_lookup_key is not None:
            _RUBRO_LOOKUP_CACHE[cache_lookup_key] = res_cod
        return res_cod

    if nota_as_item_candidates:
        nota_as_item_candidates.sort()
        res_cod = nota_as_item_candidates[0]
        if cache_lookup_key is not None:
            _RUBRO_LOOKUP_CACHE[cache_lookup_key] = res_cod
        return res_cod

    # 2. Búsqueda por similitud de tokens (manejando singular/plural y stopwords)
    target_tokens = _stem_tokens(linea_item)
    if not target_tokens:
        if cache_lookup_key is not None:
            _RUBRO_LOOKUP_CACHE[cache_lookup_key] = None
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
        res_cod = scored_cands[0][1]
        if cache_lookup_key is not None:
            _RUBRO_LOOKUP_CACHE[cache_lookup_key] = res_cod
        return res_cod

    if cache_lookup_key is not None:
        _RUBRO_LOOKUP_CACHE[cache_lookup_key] = None
    return None


def resolve_je_account_and_nota(
    linea_item: str,
    cuenta_codigo: Optional[str] = None,
    linea_nota: Optional[str] = None,
    empresa: Optional[str] = None,
    account_map: Optional[Dict[str, Dict[str, Any]]] = None,
) -> tuple[Optional[str], Optional[str]]:
    """
    Resuelve (cuenta_codigo, linea_nota) para una línea de asiento de consolidación,
    infiriendo desde el catálogo 9999xxx cuando alguno de los campos viene vacío.
    """
    if account_map is None:
        account_map = load_consolidation_account_map(empresa)

    cod_clean = str(cuenta_codigo).strip() if cuenta_codigo and str(cuenta_codigo).strip().lower() not in ("none", "nan", "") else None
    nota_clean = str(linea_nota).strip() if linea_nota and str(linea_nota).strip().lower() not in ("none", "nan", "", "sin detalle") else None

    if not cod_clean:
        cod_clean = find_account_by_rubro_nota(
            linea_item,
            nota_clean,
            empresa=empresa,
            account_map=account_map,
        )

    if cod_clean and not nota_clean:
        acc_info = account_map.get(cod_clean)
        if acc_info and acc_info.get("linea_nota"):
            nota_clean = acc_info.get("linea_nota")

    return cod_clean, nota_clean


def backfill_missing_je_account_codes(db=None) -> Dict[str, int]:
    """
    Homologa en base de datos `cuenta_codigo` (y `linea_nota` cuando viene vacía)
    en todas las líneas de `ConsolidationJournalEntry` sin modificar ningún monto ni período.
    """
    from src.models.database import SessionLocal
    from src.models.consolidacion import ConsolidationJournalEntry

    local_db = False
    if db is None:
        db = SessionLocal()
        local_db = True

    updated_codes = 0
    updated_notas = 0
    total = 0
    try:
        acc_map = load_consolidation_account_map()
        rows = db.query(ConsolidationJournalEntry).all()
        total = len(rows)
        for r in rows:
            cur_cod = getattr(r, "cuenta_codigo", None)
            cur_nota = getattr(r, "linea_nota", None)
            new_cod, new_nota = resolve_je_account_and_nota(
                r.linea_item,
                cuenta_codigo=cur_cod,
                linea_nota=cur_nota,
                account_map=acc_map,
            )
            if (not cur_cod or str(cur_cod).strip() == "") and new_cod:
                r.cuenta_codigo = new_cod
                updated_codes += 1
            if (not cur_nota or str(cur_nota).strip() == "") and new_nota:
                r.linea_nota = new_nota
                updated_notas += 1
        if updated_codes > 0 or updated_notas > 0:
            db.commit()
    finally:
        if local_db:
            db.close()

    return {"total": total, "updated_codes": updated_codes, "updated_notas": updated_notas}


