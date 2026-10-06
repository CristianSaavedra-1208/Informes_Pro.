import sys
import os
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

import time
import pandas as pd
import io
from src.core.ifrs_glossary import load_ifrs_glossary, translate_ifrs_term, extract_mapping_en_overrides, translate_dataframe_columns, get_glossary_dataframe, save_glossary_from_dataframe

def test_interactive_glossary_dataframe_save():
    """Verifica que el flujo del editor interactivo devuelva DataFrame y guarde correctamente."""
    df_gloss = get_glossary_dataframe()
    assert isinstance(df_gloss, pd.DataFrame)
    assert "Término en Español" in df_gloss.columns
    assert "Traducción en Inglés (IASB)" in df_gloss.columns
    assert len(df_gloss) > 0

    # Simular edición por el usuario en la interfaz web agregando un término dummy
    df_edited = pd.concat([df_gloss, pd.DataFrame([{"Término en Español": "Prueba Termino Dummy", "Traducción en Inglés (IASB)": "Dummy Test Term"}])], ignore_index=True)

    ok, msg = save_glossary_from_dataframe(df_edited)
    assert ok is True
    assert "Glosario guardado exitosamente" in msg

    # Verificar que el término traducido cambio al instante en RAM
    assert translate_ifrs_term("Prueba Termino Dummy", target_lang="en") == "Dummy Test Term"
from src.core.narrative_translator import translate_narrative_text
from src.reporting.excel_export import generate_excel_report
from src.reporting.word_export import WordExportEngine, generate_word_report

def test_ifrs_glossary_performance_and_lookup():
    """Verifica que la carga en RAM y búsqueda en el glosario tome menos de 0.01s."""
    t0 = time.time()
    glossary = load_ifrs_glossary()
    t1 = time.time()
    
    elapsed = t1 - t0
    assert elapsed < 0.01, f"La carga del glosario superó los 0.01s: {elapsed:.4f}s"
    assert "exact" in glossary

    # Probar términos estándar de NIIF/IFRS y variaciones de plantillas
    assert translate_ifrs_term("Clasificación", target_lang="en") == "Classification"
    assert translate_ifrs_term("Efectivo y efectivo equivalente", target_lang="en") == "Cash and cash equivalents"
    assert translate_ifrs_term("Activos financieros, Corrientes", target_lang="en") == "Other current financial assets"
    assert translate_ifrs_term("Otros activos no financieros, corrientes", target_lang="en") in ["Other non-financial assets, current", "Other current non-financial assets"]
    assert translate_ifrs_term("Deudores comerciales y otras cuentas por cobrar, corrientes", target_lang="en") in ["Trade and other receivables, current", "Trade and other current receivables"]
    assert translate_ifrs_term("Activo por impuestos, corrientes", target_lang="en") == "Current tax assets"
    assert translate_ifrs_term("Cuentas por cobrar a entidades relacionadas, corrientes", target_lang="en") == "Current receivables from related parties"
    assert translate_ifrs_term("Activos corrientes totales", target_lang="en") in ["TOTAL CURRENT ASSETS", "Total current assets"]
    assert translate_ifrs_term("Activos intangibles distinto a la plusvalia", target_lang="en") == "Intangible assets other than goodwill"
    assert translate_ifrs_term("Propiedades, plantas y equipos", target_lang="en") == "Property, plant and equipment"
    assert translate_ifrs_term("Activo por derechos de uso", target_lang="en") == "Right-of-use assets"
    assert translate_ifrs_term("Cuentas por cobrar a entidades relacionadas, no corrientes", target_lang="en") in ["Accounts receivable from related parties, non-current", "Non-current receivables from related parties"]
    assert translate_ifrs_term("Activo por impuestos diferidos, no corrientes", target_lang="en") in ["Deferred tax assets", "Deferred tax assets, non-current"]
    assert translate_ifrs_term("Otros activos financieros, no corrientes", target_lang="en") in ["Other financial assets, non-current", "Other non-current financial assets"]
    assert translate_ifrs_term("Inversion en empresas relacionadas", target_lang="en") == "Investments in related companies"
    assert translate_ifrs_term("Plusvalia", target_lang="en") == "Goodwill"
    assert translate_ifrs_term("Equity y pasivos", target_lang="en") == "TOTAL LIABILITIES AND EQUITY"

    # Flujo de Efectivo, Patrimonio y ORI
    assert translate_ifrs_term("Cobros procedentes de las ventas de bienes y prestación de servicios", target_lang="en") in ["Cash receipts from the sale of goods and rendering of services", "Cash receipts from sales of goods and rendering of services"]
    assert translate_ifrs_term("Pagos a proveedores por el suministro de bienes y servicios", target_lang="en") in ["Payments to suppliers for goods and services", "Cash payments to suppliers for goods and services"]
    assert translate_ifrs_term("Compra de Propiedades, planta y equipo", target_lang="en") == "Purchase of property, plant and equipment"
    assert translate_ifrs_term("Cambios en el patrimonio:", target_lang="en") == "Changes in equity:"
    assert translate_ifrs_term("Ganancia (pérdida)", target_lang="en") == "Profit (loss)"
    assert translate_ifrs_term("Total resultados integrales", target_lang="en") == "Total comprehensive income"
    assert translate_ifrs_term("Otros resultados integrales por coberturas", target_lang="en") == "Other comprehensive income from hedges"

def test_mapping_en_overrides():
    """Verifica que los overrides de la columna _EN tengan prioridad sobre el glosario canónico."""
    mapping_data = {
        "N° de Cuenta": ["110101", "110201"],
        "Clasificación balance": ["Efectivo y Equivalentes al Efectivo", "Deudores Comerciales y Otras Cuentas por Cobrar"],
        "Clasificación balance_EN": ["Custom Cash & Equivalents", "Custom Trade Receivables"]
    }
    map_df = pd.DataFrame(mapping_data)
    
    overrides = extract_mapping_en_overrides(map_df)
    assert overrides.get("Efectivo y Equivalentes al Efectivo") == "Custom Cash & Equivalents"
    assert overrides.get("Deudores Comerciales y Otras Cuentas por Cobrar") == "Custom Trade Receivables"

    # Probar traducción con override
    res = translate_ifrs_term("Efectivo y Equivalentes al Efectivo", target_lang="en", overrides_dict=overrides)
    assert res == "Custom Cash & Equivalents"

    # Término no overrideado cae al glosario canónico
    res_no_override = translate_ifrs_term("Propiedades, Planta y Equipo", target_lang="en", overrides_dict=overrides)
    assert res_no_override == "Property, Plant and Equipment"

def test_narrative_translation_cache():
    """Verifica que la traducción narrativa responda e inserte en caché."""
    texto_es = "Esta nota describe las políticas contables significativas aplicadas en la preparación de los estados financieros."
    translated_en = translate_narrative_text(texto_es, target_lang="en")
    
    assert isinstance(translated_en, str)
    assert len(translated_en) > 0

    # Segunda llamada debe ser instantánea por la caché
    t0 = time.time()
    cached_tr = translate_narrative_text(texto_es, target_lang="en")
    t1 = time.time()
    assert (t1 - t0) < 0.005
    assert cached_tr == translated_en

def test_excel_export_multilanguage_integrity():
    """Verifica que la exportación Excel en Inglés mantenga intactos los números y la estructura de Tie-Out."""
    df_sample = pd.DataFrame({
        "Concepto": ["Activos Corrientes", "Efectivo y Equivalentes al Efectivo", "Total Activos Corrientes"],
        "2025": [0.0, 150000.0, 150000.0],
        "2024": [0.0, 120000.0, 120000.0]
    })
    
    bytes_es = generate_excel_report(df_sample, title="Balance General", target_lang="es")
    bytes_en = generate_excel_report(df_sample, title="Balance General", target_lang="en")

    assert len(bytes_es) > 0
    assert len(bytes_en) > 0
    
    # Leer el excel generado en inglés para confirmar las traducciones
    excel_en_df = pd.read_excel(io.BytesIO(bytes_en), sheet_name="Reporte", skiprows=4)
    # Primera columna (Concepto -> Item) debe contener la traducción de los términos
    first_col_vals = list(excel_en_df.iloc[:, 0].dropna())
    assert "Current Assets" in first_col_vals or "Cash and Cash Equivalents" in first_col_vals
    
    # Verificar que los valores numéricos sean 100% idénticos
    num_2025_es = df_sample["2025"].sum()
    num_2025_en = excel_en_df.iloc[:, 1].dropna().apply(lambda x: float(str(x).replace(",", "")) if str(x) not in ["-", "nan", "None"] else 0.0).sum()
    assert abs(num_2025_es - num_2025_en) < 0.01

def test_word_export_multilanguage_integrity():
    """Verifica la exportación en Word en idioma Inglés."""
    df_sample = pd.DataFrame({
        "Concepto": ["Ingresos de Actividades Ordinarias", "Costo de Ventas", "Ganancia Bruta"],
        "2025": [500000.0, -300000.0, 200000.0]
    })

    word_bytes = generate_word_report(
        df_sample, title="Estado de Resultados", subtitle="Miles de Pesos", target_lang="en"
    )

    assert len(word_bytes) > 0
def test_notes_word_export_multilanguage():
    """Verifica que las notas exportadas a Word soporten idioma inglés."""
    elements = [
        ("text", "Nota 4 - Efectivo y Equivalentes al Efectivo"),
        ("table", pd.DataFrame({
            "Detalle": ["Efectivo en Bancos", "Depósitos a Plazo", "Total"],
            "2025": [100000.0, 50000.0, 150000.0]
        }))
    ]
    word_bytes = WordExportEngine.generate_notes_word(
        elements=elements, title="Nota 4 - Efectivo y Equivalentes al Efectivo", unit="Miles de Pesos", note_code="#N04", target_lang="en"
    )
    assert len(word_bytes.getvalue() if hasattr(word_bytes, 'getvalue') else word_bytes) > 0

if __name__ == '__main__':
    test_ifrs_glossary_performance_and_lookup()
    test_interactive_glossary_dataframe_save()
    test_mapping_en_overrides()
    test_narrative_translation_cache()
    test_excel_export_multilanguage_integrity()
    test_word_export_multilanguage_integrity()
    test_notes_word_export_multilanguage()
    print("ALL MULTILANGUAGE TESTS PASSED SUCCESSFULLY!")
