import pandas as pd
from openpyxl import load_workbook
import traceback

def test_er():
    import pickle
    try:
        df = pd.read_excel('data/empresas/Pacifico SpA/Estado de Resultados Clasificados.xlsx')
        for i in reversed(range(len(df))):
            clasif = str(df.iloc[i, 0]).lower()
            if 'continuadas' in clasif or ('ganancia' in clasif and 'bruta' not in clasif and 'antes' not in clasif and 'otras' not in clasif and 'impuesto' not in clasif):
                v25 = df.iloc[i, 1]
                print(f"FOUND: {clasif} -> {v25}")
                break
    except Exception as e:
        print(e)
        
if __name__ == '__main__':
    test_er()
