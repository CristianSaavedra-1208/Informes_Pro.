"""
Motor de Traducción de Textos Narrativos y Cualitativos de Notas.
Utiliza traducción automática con almacenamiento en Caché Local (disco/RAM) para cero latencia
y funcionamiento continuo.
"""

import os
import json
import hashlib

_CACHE_FILE = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), "data", "translation_cache.json")
_NARRATIVE_CACHE = None

def _load_cache():
    global _NARRATIVE_CACHE
    if _NARRATIVE_CACHE is not None:
        return _NARRATIVE_CACHE

    os.makedirs(os.path.dirname(_CACHE_FILE), exist_ok=True)
    if os.path.exists(_CACHE_FILE):
        try:
            with open(_CACHE_FILE, "r", encoding="utf-8") as f:
                _NARRATIVE_CACHE = json.load(f)
        except Exception:
            _NARRATIVE_CACHE = {}
    else:
        _NARRATIVE_CACHE = {}

    return _NARRATIVE_CACHE

def _save_cache():
    global _NARRATIVE_CACHE
    if _NARRATIVE_CACHE is None:
        return
    try:
        os.makedirs(os.path.dirname(_CACHE_FILE), exist_ok=True)
        with open(_CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(_NARRATIVE_CACHE, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"[Narrative Translator] Error al guardar caché de traducción: {e}")

def translate_narrative_text(text, target_lang="es"):
    """
    Traduce párrafos narrativos o textos explicativos de notas.
    """
    if not text or not isinstance(text, str) or not text.strip():
        return text

    target_lang = str(target_lang).strip().lower()
    if target_lang != "en":
        return text

    cache = _load_cache()
    text_hash = hashlib.md5(text.encode("utf-8")).hexdigest()

    if text_hash in cache:
        return cache[text_hash]

    translated = None

    # Intentar usar deep_translator si está disponible
    try:
        from deep_translator import GoogleTranslator
        translator = GoogleTranslator(source="es", target="en")
        translated = translator.translate(text)
    except Exception:
        # Fallback si no hay conexión a internet o la librería no está instalada
        pass

    if not translated:
        # Fallback usando el glosario de términos IFRS
        from src.core.ifrs_glossary import translate_ifrs_term
        translated = translate_ifrs_term(text, target_lang="en")

    cache[text_hash] = translated
    _save_cache()

    return translated
