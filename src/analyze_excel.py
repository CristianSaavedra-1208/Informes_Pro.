import pandas as pd
import warnings
warnings.simplefilter(action='ignore', category=UserWarning)

file_path = r"c:\Users\cfsaa\OneDrive\Desktop\Informes_Pro\BASE INFORME_Balance Diciembre 2025_PACIFICO_v7_DICIEMBRE_Je Inv+IFRS16_JEs PwC_v2_final 19.02.26.xlsx"

try:
    xl = pd.ExcelFile(file_path, engine="openpyxl")
    print("=== Hojas del Archivo Excel ===")
    for sheet in xl.sheet_names:
        print(f"- {sheet}")
        
    print("\n\n=== Buscando hojas clave (EEFF, Notas, Balance) ===")
    target_keywords = ["EEFF", "BALANCE", "ESTADO", "NOTA", "SFP", "ERI", "EFE"]
    
    for sheet in xl.sheet_names:
        if any(kw in sheet.upper() for kw in target_keywords):
            print(f"\n--- Analizando hoja: {sheet} ---")
            # Parseamos un poco de datos, manejando correctamente headers vacíos
            df = xl.parse(sheet, nrows=15, header=None)
            print("Primeras 10 filas sin asumir header:")
            print(df.head(10).to_string())
            print("="*60)
            
except Exception as e:
    print("Error leyendo excel:", e)
