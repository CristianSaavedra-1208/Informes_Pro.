import streamlit as st
import pandas as pd
import os
import time
from datetime import datetime
from src.core.excel_utils import df_to_excel_bytes, format_periodo
from src.models.database import SessionLocal
from src.models.consolidacion import ConsolidationGroup
from src.core.company_resolver import CompanyResolver
from src.core.consolidacion_lock_manager import ConsolidationLockManager


def _format_execution_time(elapsed_sec: float) -> str:
    if elapsed_sec < 60.0:
        return f"{elapsed_sec:.2f} s"
    mins = int(elapsed_sec // 60)
    rem = elapsed_sec % 60.0
    return f"{mins} min {rem:.2f} s ({elapsed_sec:.2f} s)"


def _store_close_notice(empresa: str, level: str, title: str, detail: str, elapsed_sec: float, periodo_lbl: str = ""):
    st.session_state["hist_flash_notice"] = {
        "empresa": empresa,
        "level": level,
        "title": title,
        "detail": detail,
        "elapsed_str": _format_execution_time(elapsed_sec),
        "periodo_lbl": periodo_lbl,
        "finished_at": datetime.now().strftime("%H:%M:%S"),
        "toast_shown": False,
    }


def _render_close_notice(empresa_actual: str):
    notice = st.session_state.get("hist_flash_notice")
    if not notice:
        return
    if notice.get("empresa") not in (empresa_actual, "__ALL__"):
        return

    if not notice.get("toast_shown"):
        icon_t = "✅" if notice.get("level") == "success" else "❌"
        try:
            st.toast(f"{notice['title']} — ⏱️ {notice['elapsed_str']}", icon=icon_t)
        except Exception:
            pass
        notice["toast_shown"] = True

    with st.container(border=True):
        c_msg, c_time, c_btn = st.columns([4.2, 1.5, 0.5])
        with c_msg:
            if notice.get("level") == "success":
                st.success(f"**{notice['title']}**\n\n{notice['detail']}")
            else:
                st.error(f"**{notice['title']}**\n\n{notice['detail']}")
        with c_time:
            st.metric(
                label=f"⏱️ Tiempo de Ejecución ({notice.get('finished_at', '')})",
                value=notice["elapsed_str"],
                delta=notice.get("periodo_lbl") or None,
            )
        with c_btn:
            st.write("")
            if st.button("✖", key="btn_dismiss_hist_notice", help="Cerrar aviso"):
                st.session_state.pop("hist_flash_notice", None)
                st.rerun()


def render(empresa_seleccionada, empresa_path):
    if empresa_seleccionada == "[GLOBAL] Configuración General" or "GLOBAL" in empresa_seleccionada:
        st.warning("Módulo de Sociedad Activa: Por favor, selecciona una empresa de trabajo específica (ej. Pacifico SpA) en la barra lateral izquierda para acceder a esta sección.")
        st.stop()
        
    st.title("Históricos")

    # Mostrar aviso persistente de cierre / reapertura / sincronización junto al marcador de tiempo
    _render_close_notice(empresa_seleccionada)
    
    # -------------------------------------------------------------
    # BÓVEDA DE SNAPSHOTS: SINCRONIZACIÓN AUTOMÁTICA
    # -------------------------------------------------------------
    with st.expander("⚡ Bóveda de Snapshots: Sincronización Automática", expanded=False):
        st.write("Genera y actualiza los Snapshots inmutables de Estados Financieros y Notas explicativas (#N01 a #N30) para todos los períodos históricos cerrados en la base de datos.")
        force_rebuild = st.checkbox(
            "🔄 Forzar actualización de Snapshots ya existentes (regenerar con últimas mejoras del motor)",
            value=False,
            key="chk_force_rebuild_snaps"
        )
        if st.button("⚡ Sincronizar Bóveda para Todos los Períodos Cerrados", key="btn_sync_all_snaps", type="secondary"):
            from src.core.snapshot_vault_manager import SnapshotVaultManager
            t0_sync = time.perf_counter()
            prog_bar = st.progress(0.0)
            status = st.empty()
            def on_s_prog(cur, tot, msg):
                prog_bar.progress(min(1.0, cur / max(1, tot)))
                el = time.perf_counter() - t0_sync
                status.info(f"⏳ {msg} ({cur}/{tot}) — ⏱️ Tiempo transcurrido: {_format_execution_time(el)}")
            succ, errs, logs = SnapshotVaultManager.sync_all_existing_historical_periods(
                progress_callback=on_s_prog,
                force_rebuild=force_rebuild
            )
            elapsed_sync = time.perf_counter() - t0_sync
            prog_bar.progress(1.0)
            status.empty()
            _store_close_notice(
                empresa="__ALL__",
                level="success" if errs == 0 else "error",
                title="✅ ¡Sincronización de Bóveda finalizada con éxito!" if errs == 0 else "⚠️ Sincronización de Bóveda finalizada con observaciones",
                detail=f"{succ} períodos listos en Bóveda ({errs} errores).",
                elapsed_sec=elapsed_sync,
                periodo_lbl=f"{succ} períodos",
            )
            st.rerun()

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
                        per_fmt_grp = format_periodo(sel_per)
                        t0_grp = time.perf_counter()
                        prog_grp = st.progress(0.0)
                        status_grp = st.empty()
                        status_grp.info(f"⏳ Iniciando cierre consolidado de {per_fmt_grp}... — ⏱️ Tiempo transcurrido: 0.00 s")

                        def _on_grp_prog(cur, tot, msg_step):
                            prog_grp.progress(min(1.0, cur / max(1, tot)))
                            el = time.perf_counter() - t0_grp
                            status_grp.info(f"⏳ {msg_step} ({cur}/{tot}) — ⏱️ Tiempo transcurrido: {_format_execution_time(el)}")

                        success, msg = ConsolidationLockManager.lock_period(
                            grupo_id,
                            sel_per,
                            user=user_curr,
                            progress_callback=_on_grp_prog
                        )
                        elapsed_grp = time.perf_counter() - t0_grp
                        prog_grp.progress(1.0)
                        status_grp.empty()

                        if success:
                            _store_close_notice(
                                empresa=empresa_seleccionada,
                                level="success",
                                title=f"✅ ¡Cierre Histórico Consolidado exitoso ({per_fmt_grp})!",
                                detail=msg,
                                elapsed_sec=elapsed_grp,
                                periodo_lbl=per_fmt_grp,
                            )
                            st.rerun()
                        else:
                            st.error(f"❌ {msg} (⏱️ Tiempo de ejecución: {_format_execution_time(elapsed_grp)})")
                else:
                    st.write(f"#### 🔓 Reapertura de Período {format_periodo(sel_per)}")
                    st.warning("Este período se encuentra actualmente cerrado. Si necesitas agregar o corregir asientos directamente en este período, confirma y ejecuta la reapertura.")
                    chk_reabrir = st.checkbox("Confirmo que deseo reabrir este período para realizar ajustes.", key="chk_confirm_unlock")
                    if st.button(f"🔓 Reabrir Período {format_periodo(sel_per)}", disabled=not chk_reabrir, key="btn_unlock_grp_per"):
                        per_fmt_un = format_periodo(sel_per)
                        t0_un = time.perf_counter()
                        with st.spinner(f"Reabriendo período consolidado {per_fmt_un}..."):
                            success, msg = ConsolidationLockManager.unlock_period(grupo_id, sel_per, user=user_curr)
                        elapsed_un = time.perf_counter() - t0_un
                        if success:
                            _store_close_notice(
                                empresa=empresa_seleccionada,
                                level="success",
                                title=f"🔓 ¡Período Consolidado {per_fmt_un} reabierto con éxito!",
                                detail=msg,
                                elapsed_sec=elapsed_un,
                                periodo_lbl=per_fmt_un,
                            )
                            st.rerun()
                        else:
                            st.error(f"❌ {msg} (⏱️ Tiempo de ejecución: {_format_execution_time(elapsed_un)})")
                            
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
                per_fmt_ind = format_periodo(periodo_cierre)
                user_ind = st.session_state.get('user_name', 'Usuario Cierre')
                t0_ind = time.perf_counter()
                prog_ind = st.progress(0.0)
                status_ind = st.empty()
                status_ind.info(f"⏳ Iniciando cierre histórico de {per_fmt_ind}... — ⏱️ Tiempo transcurrido: 0.00 s")

                def _on_ind_prog(cur, tot, msg_step):
                    prog_ind.progress(min(1.0, cur / max(1, tot)))
                    el = time.perf_counter() - t0_ind
                    status_ind.info(f"⏳ {msg_step} ({cur}/{tot}) — ⏱️ Tiempo transcurrido: {_format_execution_time(el)}")

                from src.core.cierre_engine import ejecutar_cierre_periodo
                success, msg = ejecutar_cierre_periodo(
                    empresa_seleccionada,
                    periodo_cierre,
                    empresa_path,
                    progress_callback=_on_ind_prog,
                    user=user_ind,
                )
                elapsed_time = time.perf_counter() - t0_ind
                prog_ind.progress(1.0)
                status_ind.empty()

                if success:
                    _store_close_notice(
                        empresa=empresa_seleccionada,
                        level="success",
                        title=f"✅ ¡Cierre Histórico ejecutado con éxito ({per_fmt_ind})!",
                        detail=msg,
                        elapsed_sec=elapsed_time,
                        periodo_lbl=per_fmt_ind,
                    )
                    st.rerun()
                else:
                    st.error(f"❌ Error durante el cierre: {msg} (⏱️ Tiempo de ejecución: {_format_execution_time(elapsed_time)})")

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
                per_fmt_reab = format_periodo(periodo_reabrir)
                t0_reab = time.perf_counter()
                with st.spinner(f"Restaurando período {per_fmt_reab} a la memoria activa..."):
                    from src.core.cierre_engine import reversar_cierre_periodo
                    success, msg = reversar_cierre_periodo(empresa_seleccionada, periodo_reabrir)
                elapsed_time = time.perf_counter() - t0_reab
                if success:
                    _store_close_notice(
                        empresa=empresa_seleccionada,
                        level="success",
                        title=f"🔓 ¡Período {per_fmt_reab} restaurado a Memoria Activa!",
                        detail=msg,
                        elapsed_sec=elapsed_time,
                        periodo_lbl=per_fmt_reab,
                    )
                    st.rerun()
                else:
                    st.error(f"❌ Error al reabrir período: {msg} (⏱️ Tiempo de ejecución: {_format_execution_time(elapsed_time)})")

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
                from src.core.snapshot_vault_manager import SnapshotVaultManager
                has_s = SnapshotVaultManager.has_active_snapshot(empresa_seleccionada, periodo_consulta)
                if has_s:
                    st.success("🔒 **Snapshot Oficial en Bóveda**: Los Estados Financieros y Notas explicativas están congelados e inmutables para reportes instantáneos.")
                else:
                    st.info("ℹ️ **Memoria Histórica Básica**: Los saldos están guardados. Puedes sincronizar la Bóveda completa en el panel superior.")
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
