"""
Motor de Glosario Canónico IFRS/NIIF en RAM con Búsqueda Normalizada y Flexible.
Ofrece búsquedas O(1) e insensibles a acentos/puntuaciones con respuesta instantánea (<0.01s).
"""

import os
import json
import re
import unicodedata

_GLOSSARY_CACHE = None

def normalize_key(text):
    """
    Normaliza un texto removiendo acentos, signos de puntuación, paréntesis y espacios extras.
    Permite encontrar coincidencias aunque la plantilla contenga diferencias de comas, paréntesis o tildes.
    """
    if not text or not isinstance(text, str):
        return ""
    clean = unicodedata.normalize('NFD', str(text)).encode('ascii', 'ignore').decode('utf-8').lower()
    clean = re.sub(r'[\,\.\-\_\;\:\/\\\(\)\°\º]', ' ', clean)
    clean = re.sub(r'\s+', ' ', clean).strip()
    return clean

def load_ifrs_glossary():
    """
    Carga el diccionario IFRS/NIIF desde JSON a RAM y construye la tabla normalizada.
    """
    global _GLOSSARY_CACHE
    if _GLOSSARY_CACHE is not None:
        return _GLOSSARY_CACHE

    json_path = os.path.join(os.path.dirname(__file__), "diccionario_ifrs_en.json")
    try:
        if os.path.exists(json_path):
            with open(json_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                exact = data.get("exact_matches", {})
                norm_map = {}
                for k, v in exact.items():
                    nk = normalize_key(k)
                    if nk and nk not in norm_map:
                        norm_map[nk] = v

                _GLOSSARY_CACHE = {
                    "exact": exact,
                    "normalized": norm_map,
                    "patterns": data.get("contains_patterns", [])
                }
        else:
            _GLOSSARY_CACHE = {"exact": {}, "normalized": {}, "patterns": []}
    except Exception as e:
        print(f"[IFRS Glossary] Error al cargar glosario IFRS: {e}")
        _GLOSSARY_CACHE = {"exact": {}, "normalized": {}, "patterns": []}

    return _GLOSSARY_CACHE

def extract_mapping_en_overrides(map_df):
    """
    Extrae un diccionario de overrides en inglés a partir de columnas con sufijo _EN
    o de nombre 'Cuenta_EN', 'Rubro_EN', 'Clasificacion_EN', etc.
    """
    overrides = {}
    if map_df is None:
        return overrides

    try:
        import pandas as pd
        if isinstance(map_df, pd.DataFrame) and not map_df.empty:
            en_cols = [c for c in map_df.columns if str(c).strip().lower().endswith('_en')]
            for en_col in en_cols:
                raw_col_name = str(en_col).strip()
                base_col_name = raw_col_name[:-3].strip()
                matching_base = next((c for c in map_df.columns if str(c).strip().lower() == base_col_name.lower()), None)
                if not matching_base:
                    matching_base = next((c for c in map_df.columns if any(x in str(c).lower() for x in ['clasifica', 'rubro', 'cuenta', 'nombre'])), map_df.columns[0])
                
                for _, row in map_df.iterrows():
                    base_val = row.get(matching_base)
                    en_val = row.get(en_col)
                    if pd.notna(base_val) and pd.notna(en_val):
                        b_str = str(base_val).strip()
                        e_str = str(en_val).strip()
                        if b_str and e_str:
                            overrides[b_str] = e_str
                            overrides[normalize_key(b_str)] = e_str
    except Exception as e:
        print(f"[IFRS Glossary] Error extrayendo overrides _EN: {e}")

    return overrides

def translate_ifrs_term(term, target_lang="es", user_override=None, overrides_dict=None):
    """
    Traduce un término o rubro financiero utilizando búsqueda exact y normalizada.
    """
    if not term or not isinstance(term, str):
        return term if term is not None else ""

    target_lang = str(target_lang).strip().lower()
    if target_lang != "en":
        return term

    # 1. Override puntual explícito
    if user_override and isinstance(user_override, str) and user_override.strip():
        return user_override.strip()

    cleaned_term = term.strip()
    if not cleaned_term:
        return term
    norm_term = normalize_key(cleaned_term)

    # 2. Diccionario de overrides de la plantilla/mapeo (_EN)
    if overrides_dict and isinstance(overrides_dict, dict):
        if cleaned_term in overrides_dict:
            return overrides_dict[cleaned_term]
        if norm_term in overrides_dict:
            return overrides_dict[norm_term]

    glossary = load_ifrs_glossary()
    exact_map = glossary.get("exact", {})
    norm_map = glossary.get("normalized", {})

    # 3. Búsqueda exacta directa O(1)
    if cleaned_term in exact_map:
        return exact_map[cleaned_term]

    # 4. Búsqueda normalizada (sin acentos, comas ni espacios extra)
    if norm_term in norm_map:
        return norm_map[norm_term]

    # 5. Sustitución de patrones si aplica
    result = cleaned_term
    for pat in glossary.get("patterns", []):
        pattern_str = pat.get("pattern")
        replacement_str = pat.get("replacement")
        if pattern_str and pattern_str in result:
            result = result.replace(pattern_str, replacement_str)

    return result

def translate_dataframe_columns(df, target_lang="es", overrides_dict=None):
    """
    Traduce los nombres de las columnas de un DataFrame de pandas de acuerdo al glosario IFRS.
    """
    if df is None or df.empty or str(target_lang).lower() != "en":
        return df

    new_cols = []
    for col in df.columns:
        col_str = str(col)
        new_cols.append(translate_ifrs_term(col_str, target_lang="en", overrides_dict=overrides_dict))

    df_copy = df.copy()
    df_copy.columns = new_cols
    return df_copy

def get_glossary_dataframe():
    """
    Retorna un DataFrame de pandas con los términos del glosario IFRS para su edición interactiva.
    """
    import pandas as pd
    glossary = load_ifrs_glossary()
    exact_map = glossary.get("exact", {})
    
    data = []
    for es_term, en_term in exact_map.items():
        data.append({
            "Término en Español": es_term,
            "Traducción en Inglés (IASB)": en_term
        })
    return pd.DataFrame(data)

def save_glossary_from_dataframe(df):
    """
    Guarda el DataFrame modificado en el archivo JSON diccionario_ifrs_en.json
    y recarga la memoria RAM (_GLOSSARY_CACHE) al instante.
    """
    global _GLOSSARY_CACHE
    if df is None or df.empty:
        return False, "El DataFrame está vacío."

    try:
        json_path = os.path.join(os.path.dirname(__file__), "diccionario_ifrs_en.json")
        patterns = []
        if os.path.exists(json_path):
            try:
                with open(json_path, "r", encoding="utf-8") as f:
                    old_data = json.load(f)
                    patterns = old_data.get("contains_patterns", [])
            except Exception:
                pass

        new_exact_matches = {}
        for _, row in df.iterrows():
            es_term = str(row.get("Término en Español", "")).strip()
            en_term = str(row.get("Traducción en Inglés (IASB)", "")).strip()
            if es_term and en_term and es_term.lower() != "nan" and en_term.lower() != "nan":
                new_exact_matches[es_term] = en_term

        payload = {
            "exact_matches": new_exact_matches,
            "contains_patterns": patterns
        }

        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)

        # Invalida caché en RAM para recargar inmediatamente
        _GLOSSARY_CACHE = None
        load_ifrs_glossary()

        return True, f"Glosario guardado exitosamente. ({len(new_exact_matches)} términos activos)."
    except Exception as e:
        return False, f"Error al guardar glosario: {e}"

