import streamlit as st
import pandas as pd
import os
from src.core.excel_utils import df_to_excel_bytes

def render(empresa_seleccionada, empresa_path):
    st.title("⚙️ Parámetros Globales del Reporte")
    st.write("Ingresa los valores macroeconómicos y fechas clave que afectarán los cálculos, consolidación y notas de todo el ciclo contable.")
    
    with st.form("form_parametros_globales"):
        st.subheader("1. Indicadores Económicos")
        col1, col2, col3, col4 = st.columns(4)
        uf = col1.number_input("Valor UF", value=st.session_state.get('param_uf', 38000.0), step=100.0)
        usd = col2.number_input("Valor USD", value=st.session_state.get('param_usd', 950.0), step=10.0)
        euro = col3.number_input("Valor Euro", value=st.session_state.get('param_euro', 1050.0), step=10.0)
        yuan = col4.number_input("Valor Yuan", value=st.session_state.get('param_yuan', 130.0), step=5.0)
        
        st.subheader("2. Marco Temporal")
        fecha_reporte = st.date_input("Fecha de Reportes a Generar", value=st.session_state.get('param_fecha', pd.Timestamp.today().date()))
        
        st.subheader("3. Data Adicional para Notas")
        data_prestamos = st.text_area("Información de Préstamos Bancarios (Condiciones, covenants, tasas)", value=st.session_state.get('param_prestamos', ''))
        
        submitted = st.form_submit_button("Guardar Parámetros", type="primary")
        
        if submitted:
            st.session_state['param_uf'] = uf
            st.session_state['param_usd'] = usd
            st.session_state['param_euro'] = euro
            st.session_state['param_yuan'] = yuan
            st.session_state['param_fecha'] = fecha_reporte
            st.session_state['param_prestamos'] = data_prestamos
            st.success("✅ Parámetros Globales guardados exitosamente. Las tasas y fechas aplicarán transversalmente.")

    st.divider()
    st.subheader("4. Administración de Datos")
    st.write("Selecciona los módulos de los cuales deseas purgar la data activa y los archivos temporales de esta sesión.")
    
    col_del1, col_del2, col_del3 = st.columns(3)
    del_plan = col_del1.checkbox("Plan de Cuentas")
    del_tb = col_del2.checkbox("Trial Balance")
    del_pl = col_del3.checkbox("P&L")
    
    if st.button("🗑️ Eliminar Datos Seleccionados", type="secondary"):
        deleted_items = []
        
        if del_plan:
            if 'plan_cuentas_df' in st.session_state:
                del st.session_state['plan_cuentas_df']
            plan_path = os.path.join(empresa_path, "plan_cuentas.xlsx")
            if os.path.exists(plan_path):
                os.remove(plan_path)
            deleted_items.append("Plan de Cuentas")
            
        if del_tb:
            if 'tb_df' in st.session_state:
                del st.session_state['tb_df']
            tb_path = os.path.join(empresa_path, "temp_uploaded.xlsx")
            if os.path.exists(tb_path):
                os.remove(tb_path)
            deleted_items.append("Trial Balance")
            
        if del_pl:
            if 'pl_df' in st.session_state:
                del st.session_state['pl_df']
            pl_path = os.path.join(empresa_path, "pl_cubo.xlsx")
            if os.path.exists(pl_path):
                os.remove(pl_path)
            deleted_items.append("P&L")
            
        if deleted_items:
            st.session_state['success_msg'] = f"✅ Se han eliminado los datos de: {', '.join(deleted_items)}."
            st.rerun()
        else:
            st.warning("⚠️ No seleccionaste ninguna casilla para eliminar.")


