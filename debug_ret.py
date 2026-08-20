import sys
import os
import pandas as pd
from src.reporting.ori_generator import OriGenerator

df = pd.read_excel('data/empresas/Pacifico SpA/Estado de Resultados Clasificados.xlsx')
gen = OriGenerator('data/empresas/Pacifico SpA/Estado de Resultados Integrales.xlsx')

res = gen.generate(df)
print(type(res))
try:
    print("Length:", len(res))
except:
    pass
