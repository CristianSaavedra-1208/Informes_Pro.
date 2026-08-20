import streamlit as st
import pandas as pd
import pickle

# Actually, I'll just load the template to see what's in it, because session_state is only in Streamlit's memory.
template_df = pd.read_excel("data/empresas/Pacifico SpA/Estado de Resultados Clasificados.xlsx")

print("COLUMNS:")
print(template_df.columns)

print("\nCLASES:")
for idx, row in template_df.iterrows():
    clasif = str(row.iloc[0]).lower()
    if 'continuadas' in clasif or ('ganancia' in clasif and 'bruta' not in clasif and 'antes' not in clasif and 'otras' not in clasif and 'impuesto' not in clasif):
        print("FOUND:", repr(row.iloc[0]))
    elif 'pérdida' in clasif or 'ganancia' in clasif:
        print("PARTIAL MATCH:", repr(row.iloc[0]))
