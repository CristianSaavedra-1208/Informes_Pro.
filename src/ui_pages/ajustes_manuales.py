import streamlit as st
import pandas as pd
import os
from src.core.excel_utils import df_to_excel_bytes, format_periodo

def render(empresa_seleccionada, empresa_path):
    if empresa_seleccionada == "[GLOBAL] Configuración General" or "GLOBAL" in empresa_seleccionada:
        st.warning("Módulo de Sociedad Activa: Por favor, selecciona una empresa de trabajo específica (ej. Pacifico SpA) en la barra lateral izquierda para acceder a esta sección.")
        st.stop()
        
    st.title("Ajustes de Auditoría")
    st.write("Registra ajustes post-cierre o extracontables generados por auditores (provisiones, reclasificaciones, etc.). Estos ajustes se sumarán al Trial Balance activo.")
    
    from src.models.trial_balance_db import TrialBalanceDB
    TrialBalanceDB.initialize()
    available_periods = TrialBalanceDB.get_available_periods(empresa_seleccionada)
    
    if not available_periods:
        st.warning("No hay periodos activos para aplicar ajustes. Carga un Trial Balance primero desde el módulo de Cargas.")
    else:
        with st.form("form_ajuste"):
            st.subheader("Nuevo Asiento de Ajuste")
            col_p, col_c = st.columns(2)
            periodo_asiento = col_p.selectbox("Selecciona Periodo Afectado", available_periods, format_func=format_periodo)
            cuenta_asiento = col_c.text_input("N° de Cuenta")
            
            from src.core.cierre_engine import es_periodo_cerrado
            cerrado = es_periodo_cerrado(empresa_seleccionada, periodo_asiento)
            if cerrado:
                st.warning("**Periodo Cerrado:** Este periodo está cerrado en el histórico. No se pueden registrar nuevos ajustes manuales. Reversa el cierre en la sección de **Históricos** si necesitas realizar cambios.")
            
            glosa = st.text_input("Glosa / Descripción")
            
            col_debe, col_haber = st.columns(2)
            debe = col_debe.number_input("Monto al DEBE (+)", min_value=0.0)
            haber = col_haber.number_input("Monto al HABER (-)", min_value=0.0)
            
            submit = st.form_submit_button("Registrar Ajuste en Base de Datos", type="primary", disabled=cerrado)
            if submit:
                if cerrado:
                    st.error("No se puede registrar el ajuste: el periodo seleccionado está cerrado.")
                elif not cuenta_asiento.strip():
                    st.error("Debes ingresar el N° de cuenta.")
                elif debe == 0.0 and haber == 0.0:
                    st.error("El ajuste debe tener un monto distinto de cero.")
                else:
                    monto_neto = debe - haber
                    try:
                        from src.models.database import SessionLocal
                        from src.models.audit_adjustment import AuditAdjustmentRecord
                        db = SessionLocal()
                        nuevo_ajuste = AuditAdjustmentRecord(
                            empresa=empresa_seleccionada,
                            periodo=periodo_asiento,
                            cuenta_id=cuenta_asiento.strip(),
                            monto=monto_neto,
                            descripcion=glosa.strip()
                        )
                        db.add(nuevo_ajuste)
                        db.commit()
                        db.close()
                        st.success("Ajuste guardado exitosamente.")
                    except Exception as e:
                        st.error(f"Error al guardar ajuste: {e}")
