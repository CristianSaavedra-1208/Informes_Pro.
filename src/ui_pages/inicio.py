import streamlit as st
import pandas as pd
import os
from src.core.excel_utils import df_to_excel_bytes

def render(empresa_seleccionada, empresa_path):
    st.title("Bienvenido a Informes Pro")
    st.write("Sistema automatizado de Emisión de Estados Financieros bajo IFRS.")
    st.info("👈 Navega por el menú lateral secuencialmente para comenzar tu proceso de cierre.")


