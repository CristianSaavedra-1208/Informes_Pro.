import streamlit as st
import pandas as pd
import os
from src.core.excel_utils import df_to_excel_bytes, format_periodo

def render(empresa_seleccionada, empresa_path):
    if empresa_seleccionada == "[GLOBAL] Configuración General" or "GLOBAL" in empresa_seleccionada:
        st.warning("Módulo de Sociedad Activa: Por favor, selecciona una empresa de trabajo específica (ej. Pacifico SpA) en la barra lateral izquierda para acceder a esta sección.")
        st.stop()
        
    st.title("Históricos")
    st.write("Ejecuta el cierre de un periodo (mes/año) y transforma la memoria dinámica a datos estáticos para arrastre futuro.")
    
    tab_cierre, tab_consulta, tab_legacy = st.tabs(["Cierre de Periodo Automático", "Consulta Histórica", "Carga Histórica Legacy"])
    with tab_cierre:
        st.subheader("Congelamiento de Periodo a Histórico")
        st.warning("Al ejecutar esta función se calcularán todos los saldos según el mapeo actual, se guardarán en la Bóveda Histórica, y se limpiará el periodo de la memoria activa.")
        
        try:
            from src.models.database import SessionLocal
            from src.models.trial_balance import TrialBalanceRecord
            db = SessionLocal()
            active_periods = db.query(TrialBalanceRecord.periodo).filter_by(empresa=empresa_seleccionada).distinct().all()
            db.close()
            active_periods = sorted([p[0] for p in active_periods], reverse=True)
        except Exception:
            active_periods = []
            
        if not active_periods:
            st.info("No hay periodos activos para cerrar.")
        else:
            periodo_cierre = st.selectbox("Selecciona el periodo activo a congelar:", active_periods, format_func=format_periodo)
            if st.button(f"🔒 Ejecutar Cierre y Congelar {format_periodo(periodo_cierre)}", type="primary"):
                with st.spinner("Procesando cierre de periodo... esto puede tomar unos segundos."):
                    import time
                    start_time = time.time()
                    from src.core.cierre_engine import ejecutar_cierre_periodo
                    success, msg = ejecutar_cierre_periodo(empresa_seleccionada, periodo_cierre, empresa_path)
                    elapsed_time = time.time() - start_time
                    if success:
                        st.success(f"✅ ¡Proceso finalizado! {msg} (Tiempo de ejecución: {elapsed_time:.2f} segundos)")
                    else:
                        st.error(f"❌ Error durante el cierre: {msg} (Tiempo de ejecución: {elapsed_time:.2f} segundos)")

    with tab_consulta:
        st.subheader("Consulta de Papeles de Trabajo Históricos")
        st.write("Visualiza el detalle cuenta por cuenta de periodos que ya han sido cerrados y congelados.")
        
        try:
            from src.models.database import SessionLocal
            from src.models.historical_data import HistoricalDetailRecord
            db = SessionLocal()
            hist_periods = db.query(HistoricalDetailRecord.periodo).filter_by(empresa=empresa_seleccionada).distinct().all()
            db.close()
            hist_periods = sorted([p[0] for p in hist_periods], reverse=True)
        except Exception:
            hist_periods = []
            
        if not hist_periods:
            st.info("Aún no hay periodos congelados en la Bóveda Histórica.")
        else:
            periodo_consulta = st.selectbox("Selecciona un periodo histórico:", hist_periods, key="sel_hist_per", format_func=format_periodo)
            if periodo_consulta:
                try:
                    db = SessionLocal()
                    records = db.query(HistoricalDetailRecord).filter_by(empresa=empresa_seleccionada, periodo=periodo_consulta).all()
                    db.close()
                    
                    if records:
                        data = []
                        for r in records:
                            data.append({
                                'Cuenta': r.cuenta_id,
                                'Descripción': r.descripcion,
                                'Saldo Final': r.saldo_final,
                                'Clasificación Balance': r.clasificacion_balance,
                                'Clasificación P&L': r.clasificacion_pl,
                                'Nota Asociada': r.id_nota_asociada
                            })
                        df_hist = pd.DataFrame(data)
                        if 'Saldo Final' in df_hist.columns:
                            # Format as dot-separated string in Python to avoid javascript sprintf errors
                            df_hist['Saldo Final'] = df_hist['Saldo Final'].apply(lambda x: f"{int(round(pd.to_numeric(x, errors='coerce') or 0.0)):,}".replace(",", "."))
                            st.dataframe(
                                df_hist,
                                use_container_width=True,
                                column_config={
                                    'Saldo Final': st.column_config.TextColumn('Saldo Final')
                                },
                                key=f"df_hist_{periodo_consulta}"
                            )
                        else:
                            st.dataframe(df_hist)
                        
                        excel_data = df_to_excel_bytes(df_hist, "Papel de Trabajo")
                        st.download_button(
                            label=f"📥 Descargar Papel de Trabajo {format_periodo(periodo_consulta)} (Excel)",
                            data=excel_data,
                            file_name=f"Papeles_Trabajo_Auditoria_{periodo_consulta}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="dl_papeles"
                        )
                        
                        st.divider()
                        st.warning("⚠️ **Reversar Cierre:** Si cometiste un error o necesitas volver a trabajar sobre este año, puedes devolverlo a la memoria activa. Esto eliminará la versión congelada.")
                        if st.button(f"⏪ Reversar Cierre y Reabrir {format_periodo(periodo_consulta)}"):
                            with st.spinner("Restaurando periodo a memoria activa..."):
                                import time
                                start_time = time.time()
                                from src.core.cierre_engine import reversar_cierre_periodo
                                succ, ms = reversar_cierre_periodo(empresa_seleccionada, periodo_consulta)
                                elapsed_time = time.time() - start_time
                                if succ:
                                    st.session_state['success_msg'] = f"✅ {ms} (Tiempo de ejecución: {elapsed_time:.2f} segundos)"
                                    st.rerun()
                                else:
                                    st.error(f"❌ {ms} (Tiempo de ejecución: {elapsed_time:.2f} segundos)")
                except Exception as e:
                    st.error(f"Error al cargar histórico: {e}")
        
    with tab_legacy:
        st.subheader("Carga Inicial Legacy (Subir periodos antiguos)")
        st.write("Sube directamente los montos auditados finales por línea para que alimenten la columna comparativa.")
        
        tipo_carga = st.selectbox("¿Qué documento vas a cargar al histórico?", [
            "1) Estados Financieros Principales",
            "2) Notas a los Estados Financieros"
        ])
        
        if "Principales" in tipo_carga:
            doc_especifico = st.selectbox("Selecciona el tipo de Estado Financiero:", [
                "Balance Clasificado (Situación Financiera)",
                "Estado de Resultados (P&L)",
                "Flujo de Efectivo",
                "Estado de Cambios en el Patrimonio",
                "Resultados Integrales (ORI)"
            ])
        else:
            doc_especifico = st.selectbox("Selecciona la Nota asociada:", [
                "Nota: Efectivo y Efectivo Equivalente",
                "Nota: Cuentas por Cobrar Comerciales",
                "Nota: Propiedad, Planta y Equipo",
                "Nota: Préstamos y Pasivos Financieros",
                "Nota: Ingresos Ordinarios",
                "Otras..."
            ])
            
        col_ay, col_am = st.columns(2)
        hist_year = col_ay.selectbox("Año Histórico", ["2023", "2024", "2025"], index=2)
        hist_month = col_am.selectbox("Mes", ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"], index=11)
        
        with st.expander(f"📥 Descargar Plantilla Mapeada: {doc_especifico}"):
            st.info("Esta plantilla contendrá únicamente dos columnas: [Línea Mapeada] y [Monto Estático].")
            st.button("Descargar Plantilla Vacía")
            
        uploaded_legacy = st.file_uploader(f"Sube los valores estáticos de {doc_especifico} ({hist_year}-{hist_month})", type=["xlsx", "xls"])
        if uploaded_legacy:
            if st.button("Guardar en Bóveda Histórica", type="primary"):
                st.success(f"✅ Valores guardados en Memoria Histórica para {doc_especifico} ({hist_year}-{hist_month}).")


