import streamlit as st
import pandas as pd
import os
from src.core.excel_utils import df_to_excel_bytes, format_periodo
from src.models.database import SessionLocal
from src.models.consolidacion import ConsolidationGroup
from src.core.company_resolver import CompanyResolver
from src.core.consolidacion_lock_manager import ConsolidationLockManager

def render(empresa_seleccionada, empresa_path):
    if empresa_seleccionada == "[GLOBAL] Configuración General" or "GLOBAL" in empresa_seleccionada:
        st.warning("Módulo de Sociedad Activa: Por favor, selecciona una empresa de trabajo específica (ej. Pacifico SpA) en la barra lateral izquierda para acceder a esta sección.")
        st.stop()
        
    st.title("Históricos")
    
    # -------------------------------------------------------------
    # DETECTAR SI LA EMPRESA ACTIVA ES UN GRUPO CONSOLIDADO
    # -------------------------------------------------------------
    is_group = False
    grupo_obj = None
    clean_name = empresa_seleccionada.replace("[GRUPO]", "").strip()
    
    db = SessionLocal()
    try:
        grupo_obj = db.query(ConsolidationGroup).filter(
            (ConsolidationGroup.nombre_grupo == clean_name) | 
            (ConsolidationGroup.nombre_grupo == empresa_seleccionada)
        ).first()
        if grupo_obj or empresa_seleccionada.startswith("[GRUPO]"):
            is_group = True
    finally:
        db.close()

    # =============================================================
    # MODO A: GRUPO CONSOLIDADO (CANDADO DE CIERRE Y SNAPSHOT)
    # =============================================================
    if is_group and grupo_obj:
        st.write(f"Gestión de Cierre, Candados de Seguridad y Protección contra escritura para **{empresa_seleccionada}**.")
        
        tab_cierre, tab_consulta = st.tabs(["🔒 Candado de Cierre Consolidado", "📊 Consulta Histórica Consolidada"])
        
        with tab_cierre:
            st.subheader("Cierre y Bloqueo de Períodos de Consolidación")
            st.info("💡 **Función de Control de Auditoría**: Al cerrar un período consolidado, este queda sellado contra modificaciones o eliminaciones directas accidentales de sus asientos de ajuste. Podrás consultarlo y **copiar sus asientos hacia períodos abiertos** en cualquier momento.")
            
            grupo_id = grupo_obj.id
            available_pers = ConsolidationLockManager.get_available_periods_for_group(grupo_id)
            
            if not available_pers:
                st.warning(f"No se detectaron períodos con información cargada en la matriz ni filiales del grupo.")
            else:
                col_sel1, col_sel2 = st.columns([2, 1])
                with col_sel1:
                    sel_per = st.selectbox("Selecciona el período a gestionar:", available_pers, format_func=format_periodo, key="sel_per_lock_grp")
                
                is_locked = ConsolidationLockManager.is_period_locked(grupo_id, sel_per)
                
                with col_sel2:
                    st.write("")
                    st.write("")
                    if is_locked:
                        st.success(f"🔒 **Estado**: CERRADO Y BLOQUEADO")
                    else:
                        st.info(f"🔓 **Estado**: ABIERTO (En Elaboración)")
                
                st.divider()
                
                user_curr = st.session_state.get('user_name', 'Administrador')
                
                if not is_locked:
                    st.write(f"#### 🔒 Bloquear Período {format_periodo(sel_per)}")
                    st.write("Al aplicar el candado, los asientos y hojas de trabajo de este período quedarán protegidos como **versión oficial auditada**.")
                    if st.button(f"🔒 Cerrar y Bloquear Período {format_periodo(sel_per)}", type="primary", key="btn_lock_grp_per"):
                        success, msg = ConsolidationLockManager.lock_period(grupo_id, sel_per, user=user_curr)
                        if success:
                            st.success(f"✅ {msg}")
                            st.rerun()
                        else:
                            st.error(f"❌ {msg}")
                else:
                    st.write(f"#### 🔓 Reapertura de Período {format_periodo(sel_per)}")
                    st.warning("Este período se encuentra actualmente cerrado. Si necesitas agregar o corregir asientos directamente en este período, confirma y ejecuta la reapertura.")
                    chk_reabrir = st.checkbox("Confirmo que deseo reabrir este período para realizar ajustes.", key="chk_confirm_unlock")
                    if st.button(f"🔓 Reabrir Período {format_periodo(sel_per)}", disabled=not chk_reabrir, key="btn_unlock_grp_per"):
                        success, msg = ConsolidationLockManager.unlock_period(grupo_id, sel_per, user=user_curr)
                        if success:
                            st.success(f"✅ {msg}")
                            st.rerun()
                        else:
                            st.error(f"❌ {msg}")
                            
            st.divider()
            st.write("### 📋 Registro Histórico de Períodos Bloqueados")
            locked_list = ConsolidationLockManager.get_locked_periods(grupo_obj.id)
            if locked_list:
                df_locks = pd.DataFrame([{
                    "Período": format_periodo(l["periodo"]),
                    "Código Período": l["periodo"],
                    "Fecha de Cierre (UTC)": l["locked_at"].strftime("%Y-%m-%d %H:%M:%S") if l["locked_at"] else "N/A",
                    "Cerrado Por": l["locked_by"],
                    "Estado": "🔒 Bloqueado contra escritura"
                } for l in locked_list])
                st.dataframe(df_locks, use_container_width=True)
            else:
                st.info("Aún no hay períodos cerrados con candado en este grupo.")

        with tab_consulta:
            st.subheader("Consulta de Hojas de Trabajo Consolidadas")
            st.write("Genera y descarga la Hoja de Trabajo y Reportes Oficiales de cualquier período.")
            
            if available_pers:
                sel_per_cons = st.selectbox("Seleccionar período a consultar:", available_pers, format_func=format_periodo, key="sel_per_hist_grp")
                if sel_per_cons:
                    from src.core.consolidacion_engine import generar_hoja_trabajo
                    df_hoja, msg_h = generar_hoja_trabajo(grupo_obj.id, sel_per_cons)
                    if df_hoja is not None:
                        is_l = ConsolidationLockManager.is_period_locked(grupo_obj.id, sel_per_cons)
                        badge_l = "🔒 Versión Cerrada y Auditada" if is_l else "🔓 Versión Dinámica Abierta"
                        st.caption(f"**Estado del Período**: {badge_l}")
                        
                        st.dataframe(df_hoja, use_container_width=True)
                        
                        excel_hoja = df_to_excel_bytes(df_hoja, f"Consolidado {sel_per_cons}")
                        st.download_button(
                            label=f"📥 Descargar Hoja de Consolidación {format_periodo(sel_per_cons)} (Excel)",
                            data=excel_hoja,
                            file_name=f"Hoja_Consolidada_{empresa_seleccionada}_{sel_per_cons}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"dl_hoja_hist_{sel_per_cons}"
                        )
                    else:
                        st.error(f"Error generando hoja: {msg_h}")
        return

    # =============================================================
    # MODO B: EMPRESA INDIVIDUAL (CIERRE DE SALDOS TRADICIONAL)
    # =============================================================
    st.write("Ejecuta el cierre de un periodo (mes/año) y transforma la memoria dinámica a datos estáticos para arrastre futuro.")
    
    tab_cierre, tab_consulta, tab_legacy = st.tabs(["Cierre de Periodo Automático", "Consulta Histórica", "Carga Histórica Legacy"])
    with tab_cierre:
        st.subheader("Congelamiento de Periodo a Histórico")
        st.warning("Al ejecutar esta función se calcularán todos los saldos según el mapeo actual, se guardarán en la Bóveda Histórica, y se limpiará el periodo de la memoria activa.")
        
        try:
            from src.models.trial_balance import TrialBalanceRecord
            db = SessionLocal()
            active_periods = db.query(TrialBalanceRecord.periodo).filter_by(empresa=empresa_seleccionada).distinct().all()
            db.close()
            active_periods = sorted([p[0] for p in active_periods], reverse=True)
        except Exception:
            active_periods = []
            
        if not active_periods:
            st.info("No hay períodos activos pendientes de cerrar en la memoria viva.")
        else:
            periodo_cierre = st.selectbox("Selecciona el periodo activo a congelar:", active_periods, format_func=format_periodo, key="sel_act_per_close")
            if st.button(f"🔒 Ejecutar Cierre y Congelar {format_periodo(periodo_cierre)}", type="primary", key="btn_close_act_per"):
                with st.spinner("Procesando cierre de periodo... esto puede tomar unos segundos."):
                    import time
                    start_time = time.time()
                    from src.core.cierre_engine import ejecutar_cierre_periodo
                    success, msg = ejecutar_cierre_periodo(empresa_seleccionada, periodo_cierre, empresa_path)
                    elapsed_time = time.time() - start_time
                    if success:
                        st.success(f"✅ ¡Proceso finalizado! {msg} (Tiempo de ejecución: {elapsed_time:.2f} segundos)")
                        st.rerun()
                    else:
                        st.error(f"❌ Error durante el cierre: {msg} (Tiempo de ejecución: {elapsed_time:.2f} segundos)")

        st.divider()
        st.subheader("🔓 Reapertura de Períodos Históricos a Memoria Activa")
        st.write("Restaura un período congelado desde la Bóveda Histórica de vuelta a la memoria viva para permitir ajustes, recargas o modificaciones de mapeo.")
        
        try:
            from src.models.historical_data import HistoricalDetailRecord
            db_h = SessionLocal()
            hist_periods_cierre = db_h.query(HistoricalDetailRecord.periodo).filter_by(empresa=empresa_seleccionada).distinct().all()
            db_h.close()
            hist_periods_cierre = sorted([p[0] for p in hist_periods_cierre], reverse=True)
        except Exception:
            hist_periods_cierre = []
            
        if not hist_periods_cierre:
            st.info("No hay períodos congelados en la Bóveda Histórica para reabrir.")
        else:
            col_reab1, col_reab2 = st.columns([2, 1])
            with col_reab1:
                periodo_reabrir = st.selectbox("Selecciona el período histórico a reabrir:", hist_periods_cierre, format_func=format_periodo, key="sel_hist_per_reopen")
            with col_reab2:
                st.write("")
                st.write("")
                chk_reopen_ind = st.checkbox("Confirmo que deseo reabrir este período.", key="chk_reopen_ind")
                
            if st.button(f"🔓 Reabrir {format_periodo(periodo_reabrir)} a Memoria Activa", type="primary", disabled=not chk_reopen_ind, key="btn_reopen_ind_per"):
                with st.spinner("Restaurando período a la memoria activa..."):
                    import time
                    start_time = time.time()
                    from src.core.cierre_engine import reversar_cierre_periodo
                    success, msg = reversar_cierre_periodo(empresa_seleccionada, periodo_reabrir)
                    elapsed_time = time.time() - start_time
                    if success:
                        st.success(f"✅ ¡Período restaurado! {msg} (Tiempo: {elapsed_time:.2f}s)")
                        st.rerun()
                    else:
                        st.error(f"❌ Error al reabrir período: {msg}")

    with tab_consulta:
        st.subheader("Consulta de Papeles de Trabajo Históricos")
        st.write("Visualiza el detalle cuenta por cuenta de periodos que ya han sido cerrados y congelados.")
        
        try:
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
                            key=f"dl_hist_{periodo_consulta}"
                        )
                except Exception as e:
                    st.error(f"Error cargando detalle histórico: {e}")

    with tab_legacy:
        st.subheader("Carga de Saldos Históricos Legacy")
        st.write("Importa balances congelados de sistemas anteriores para análisis comparativo.")
        uploaded_legacy = st.file_uploader("Subir archivo de Saldos Legacy (Excel)", type=["xlsx", "xls"], key="up_legacy")
        if uploaded_legacy:
            st.info("Funcionalidad de importación legacy activa.")
