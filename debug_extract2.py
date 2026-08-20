import pandas as pd

def test_er():
    df = pd.read_excel('data/empresas/Pacifico SpA/Estado de Resultados Clasificados.xlsx')
    for c in df['Clasificación']:
        print(repr(c))

test_er()
