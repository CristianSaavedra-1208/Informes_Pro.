import streamlit as st
import pandas as pd
import os
from src.core.excel_utils import df_to_excel_bytes, format_periodo

def render(empresa_seleccionada, empresa_path):
    # Vista simplificada exclusiva para el rol Analista de Reportes y Supervisor
    if st.session_state.get('auth_role') in ["Analista de Reportes", "Supervisor"]:
        st.title("🔑 Mi Perfil y Contraseña")
        st.write(f"Usuario: **{st.session_state.get('auth_name')}** (`{st.session_state.get('auth_user')}`)")
        st.write(f"Rol: `{st.session_state.get('auth_role')}`")
        st.divider()
        
        with st.form("form_own_password"):
            st.subheader("Modificar Mi Contraseña")
            new_p1 = st.text_input("Nueva Contraseña:", type="password", key="rep_new_p1")
            new_p2 = st.text_input("Confirmar Nueva Contraseña:", type="password", key="rep_new_p2")
            sub = st.form_submit_button("💾 Guardar Nueva Contraseña", type="primary")
            if sub:
                if not new_p1 or not new_p2:
                    st.error("⚠️ La contraseña no puede estar vacía.")
                elif new_p1 != new_p2:
                    st.error("❌ Las contraseñas no coinciden.")
                else:
                    from src.core.security_engine import change_user_password
                    ok, msg = change_user_password(st.session_state.get('auth_user'), new_p1, actor_username=st.session_state.get('auth_user'))
                    if ok:
                        st.success(f"✅ {msg}")
                    else:
                        st.error(f"❌ {msg}")
        return

    st.title("Configuraciones del Sistema")
    st.write("Administración general de empresas, integraciones y seguridad.")
    
    empresas_dir = os.path.join("data", "empresas")
    empresas = sorted([d for d in os.listdir(empresas_dir) if os.path.isdir(os.path.join(empresas_dir, d))])
    
    is_global = "GLOBAL" in str(empresa_seleccionada).upper()
    if not is_global:
        st.warning("🔒 **Acceso Restringido:** El módulo de Roles, Usuarios y Configuración General del Sistema solo puede ser administrado desde el entorno **🌐 [GLOBAL] Configuración General**.")
        st.info("Para acceder a estas funciones de administración, selecciona **🌐 [GLOBAL] Configuración General** en el selector de empresa de la barra lateral.")
        return

    tabs = st.tabs(["🏢 Empresas y Entornos", "🌐 Glosario IFRS (ES ↔ EN)", "🔌 Conexiones ERP (API)", "👥 Roles & Settings", "🗑️ Eliminación de Data"])
    tab_empresas, tab_glossary, tab_erp, tab_roles, tab_danger = tabs
    
    with tab_glossary:
        st.subheader("🌐 Editor de Glosario IFRS / NIIF (ES ↔ EN)")
        st.caption("Modifica o agrega traducciones de rubros contables y títulos de Notas. Los cambios se aplicarán instantáneamente a todos los reportes Excel, Word y vistas en pantalla.")

        from src.core.ifrs_glossary import get_glossary_dataframe, save_glossary_from_dataframe

        df_glossary = get_glossary_dataframe()

        # Buscador en vivo
        search_query = st.text_input("🔍 Buscar término en el glosario:", key="search_ifrs_glossary")
        if search_query.strip():
            sq = search_query.strip().lower()
            mask = (
                df_glossary["Término en Español"].astype(str).str.lower().str.contains(sq) |
                df_glossary["Traducción en Inglés (IASB)"].astype(str).str.lower().str.contains(sq)
            )
            filtered_df = df_glossary[mask].reset_index(drop=True)
        else:
            filtered_df = df_glossary

        st.markdown(f"**Términos activos:** `{len(df_glossary)}` (mostrando `{len(filtered_df)}`).")

        edited_df = st.data_editor(
            filtered_df,
            num_rows="dynamic",
            use_container_width=True,
            key="data_editor_ifrs_glossary",
            column_config={
                "Término en Español": st.column_config.TextColumn("Término en Español", width="large", required=True),
                "Traducción en Inglés (IASB)": st.column_config.TextColumn("Traducción en Inglés (IASB)", width="large", required=True),
            }
        )

        col_save, col_info = st.columns([1, 1])
        with col_save:
            if st.button("💾 Guardar Cambios en Glosario IFRS", type="primary", use_container_width=True):
                # Si se usó filtro, combinar la edición con los términos no filtrados
                if search_query.strip():
                    merged_dict = dict(zip(df_glossary["Término en Español"], df_glossary["Traducción en Inglés (IASB)"]))
                    for _, r in edited_df.iterrows():
                        k_es = str(r.get("Término en Español", "")).strip()
                        v_en = str(r.get("Traducción en Inglés (IASB)", "")).strip()
                        if k_es and v_en:
                            merged_dict[k_es] = v_en
                    final_df = pd.DataFrame([{"Término en Español": k, "Traducción en Inglés (IASB)": v} for k, v in merged_dict.items()])
                else:
                    final_df = edited_df

                ok, msg = save_glossary_from_dataframe(final_df)
                if ok:
                    st.success(f"✅ {msg}")
                    st.rerun()
                else:
                    st.error(f"❌ {msg}")
        with col_info:
            st.info("💡 **Consejo:** Puedes pegar directamente múltiples filas desde Excel o presionar `+` en la tabla para agregar nuevos términos.")
    
    with tab_empresas:
        col1, col2, col3 = st.columns(3)
        db_path = os.path.join(os.path.dirname(empresas_dir), "informes_pro.db")
        from src.core.company_resolver import CompanyResolver
        CompanyResolver.ensure_catalog_table()
        all_catalog = CompanyResolver.get_all_companies(include_groups=True, only_active=True)
        
        with col1:
            with st.expander("➕ Crear Nueva Empresa", expanded=True):
                nueva_empresa = st.text_input("Nombre / Razón Social", key="create_empresa_input")
                es_grupo_check = st.checkbox("Es Grupo Consolidado", key="create_is_group_check")
                if st.button("Crear Empresa", key="btn_create_empresa", type="primary"):
                    if nueva_empresa.strip():
                        success, msg, code = CompanyResolver.register_company(nueva_empresa.strip(), es_consolidado=es_grupo_check)
                        if success:
                            st.success(f"✅ {msg}")
                            st.rerun()
                        else:
                            st.error(f"❌ {msg}")
                    else:
                        st.warning("El nombre de la empresa no puede estar vacío.")

        with col2:
            with st.expander("✏️ Modificar Carátula / Nombre", expanded=True):
                if all_catalog:
                    sel_cat = st.selectbox(
                        "Selecciona sociedad a modificar", 
                        all_catalog, 
                        format_func=lambda c: f"[{c['codigo_id']}] {c['nombre_caratula']}",
                        key="rename_select_cat"
                    )
                    st.caption(f"🔑 **Código Inmutable:** `{sel_cat['codigo_id']}` {'(Grupo Consolidado)' if sel_cat['es_consolidado'] else '(Sociedad Individual)'}")
                    nuevo_nombre = st.text_input("Nuevo Nombre / Razón Social", value=sel_cat['nombre_caratula'], key="rename_input_cat")
                    
                    if st.button("Guardar Nuevo Nombre", key="btn_rename_empresa_cat"):
                        nuevo_clean = nuevo_nombre.rstrip('.').strip()
                        if nuevo_clean and nuevo_clean != sel_cat['nombre_caratula']:
                            from src.core.company_manager import rename_company_cascade
                            # Actualizar nombre en cascada y catálogo
                            success_cas, msg_cas = rename_company_cascade(empresas_dir, db_path, sel_cat['nombre_caratula'], nuevo_clean)
                            if success_cas:
                                CompanyResolver.update_display_name(sel_cat['codigo_id'], nuevo_clean)
                                
                                # Limpiar cache local forzosamente
                                for key in ['plan_cuentas_df', 'tb_df', 'tb_df_comp', 'map_balance_df', 'map_pl_df', 'pl_df', 'pl_df_comp', 'er_preview_df', 'preview_df']:
                                    if key in st.session_state:
                                        del st.session_state[key]
                                
                                if st.session_state.get('empresa_activa') == sel_cat['nombre_caratula']:
                                    st.session_state['empresa_activa'] = nuevo_clean
                                    st.session_state['empresa_activa_prev'] = nuevo_clean
                                    
                                st.success(f"✅ Carátula actualizada a '{nuevo_clean}' para el código [{sel_cat['codigo_id']}].")
                                st.rerun()
                            else:
                                st.error(msg_cas)
                        elif nuevo_clean == sel_cat['nombre_caratula']:
                            st.info("El nombre ingresado es idéntico al actual.")
                        else:
                            st.warning("El nuevo nombre no puede estar vacío.")
                else:
                    st.warning("No hay empresas registradas.")

        with col3:
            with st.expander("🗑️ Eliminar Empresa", expanded=True):
                if all_catalog:
                    from src.core.company_manager import check_company_has_data, delete_company
                    sel_cat_del = st.selectbox(
                        "Selecciona empresa a eliminar", 
                        all_catalog, 
                        format_func=lambda c: f"[{c['codigo_id']}] {c['nombre_caratula']}",
                        key="delete_empresa_select_cat"
                    )
                    empresa_a_eliminar = sel_cat_del['nombre_caratula']
                    
                    has_data, reasons = check_company_has_data(empresas_dir, db_path, empresa_a_eliminar)
                    if has_data:
                        st.warning(f"⚠️ **[{sel_cat_del['codigo_id']}] {empresa_a_eliminar}** tiene datos registrados:\n" + "\n".join([f"- {r}" for r in reasons]) + "\n\n*Antes de eliminar, debe borrar la data desde la pestaña **Eliminación de Data**.*")
                    else:
                        st.info(f"ℹ️ **[{sel_cat_del['codigo_id']}] {empresa_a_eliminar}** no contiene datos transaccionales. Puede ser eliminada de forma segura.")
                        
                    confirm_del = st.checkbox("Confirmo que deseo eliminar esta empresa", key="check_confirm_del_empresa")
                    if st.button("Eliminar Empresa", type="primary", disabled=not confirm_del or has_data, key="btn_delete_empresa"):
                        success, msg = delete_company(empresas_dir, db_path, empresa_a_eliminar)
                        if success:
                            # Limpiar cache local
                            for key in ['plan_cuentas_df', 'tb_df', 'tb_df_comp', 'map_balance_df', 'map_pl_df', 'pl_df', 'pl_df_comp', 'er_preview_df', 'preview_df']:
                                if key in st.session_state:
                                    del st.session_state[key]
                                    
                            remaining_cos = CompanyResolver.get_all_companies(include_groups=True, only_active=True)
                            if st.session_state.get('empresa_activa') == empresa_a_eliminar:
                                next_co = remaining_cos[0]['nombre_caratula'] if remaining_cos else ""
                                st.session_state['empresa_activa'] = next_co
                                st.session_state['empresa_activa_prev'] = next_co
                                
                            st.success(msg)
                            st.rerun()
                        else:
                            st.error(msg)
                else:
                    st.warning("No hay empresas registradas.")
                
    with tab_erp:
        st.write(f"Configura la extracción automática de saldos directamente desde tu ERP para **{empresa_seleccionada}**.")
        
        with st.expander("⚙️ Configuración de Conexión (API ERP)", expanded=True):
            erp_col1, erp_col2 = st.columns(2)
            with erp_col1:
                tipo_erp = st.selectbox("Proveedor de Software ERP", ["Netsuite", "Odoo", "SAP Business One", "SAP S/4HANA", "Microsoft Dynamics", "Oracle NetSuite", "Xero", "Otro"])
                api_url = st.text_input("Endpoint URL (Base API)")
                
            with erp_col2:
                api_key = st.text_input("API Key / Client ID")
                api_secret = st.text_input("Client Secret / Token", type="password")
                
            st.markdown("*(Estas credenciales se utilizarán por el futuro adaptador ETL para extraer el balance automáticamente, obviando los archivos Excel).*")
            
            test_col1, test_col2 = st.columns([1,3])
            with test_col1:
                if st.button("Guardar Credenciales ERP", type="primary"):
                    import json
                    settings_path = os.path.join(empresa_path, "erp_settings.json")
                    with open(settings_path, 'w') as f:
                        json.dump({
                            "erp": tipo_erp,
                            "url": api_url,
                            "key": api_key,
                            "configured": True
                        }, f)
                    st.session_state['success_msg'] = "✅ Credenciales de conexión al ERP guardadas con éxito."
                    st.rerun()
            with test_col2:
                if st.button("Probar Conexión al ERP"):
                    st.info(f"Haciendo ping a {tipo_erp}... (Módulo ETL Backend en construcción. El enchufe UI está instalado correctamente).")

    with tab_roles:
        st.subheader("🛡️ Módulo de Seguridad, Usuarios & Permisos")
        st.markdown("Administra las cuentas de acceso, roles de usuario, permisos y la bitácora de auditoría global del sistema.")
        
        from src.core.security_engine import (
            get_all_users, create_user, update_user_role, 
            update_user_status, change_user_password, delete_user, get_audit_logs
        )
        
        subtab_users, subtab_audit = st.tabs(["Usuarios y Roles", "Bitácora de Auditoría"])
        
        with subtab_users:
            u_col1, u_col2 = st.columns([1, 1])
            
            with u_col1:
                with st.expander("➕ Crear Nuevo Usuario", expanded=True):
                    with st.form("form_new_user", clear_on_submit=True):
                        nu_user = st.text_input("Nombre de Usuario (Login):", placeholder="ej: jgonzalez").strip()
                        nu_name = st.text_input("Nombre Completo:", placeholder="ej: Juan González")
                        nu_email = st.text_input("Correo Electrónico:", placeholder="ej: jgonzalez@empresa.cl")
                        nu_pass = st.text_input("Contraseña:", type="password", placeholder="••••••••")
                        nu_role = st.selectbox("Rol Asignado:", ["Administrador", "Supervisor", "Analista Contable", "Analista de Reportes", "Auditor Lector"])
                        
                        btn_nu = st.form_submit_button("👤 Guardar Usuario", type="primary", use_container_width=True)
                        if btn_nu:
                            if not nu_user or not nu_pass:
                                st.error("⚠️ El usuario y la contraseña son obligatorios.")
                            else:
                                ok, msg = create_user(nu_user, nu_pass, nu_name, nu_email, nu_role, created_by=st.session_state.get('auth_user', 'admin'))
                                if ok:
                                    st.success(f"✅ {msg}")
                                    st.rerun()
                                else:
                                    st.error(f"❌ {msg}")

            with u_col2:
                users_list = get_all_users()
                st.subheader(f"👥 Usuarios Registrados ({len(users_list)})")
                if users_list:
                    df_users = pd.DataFrame(users_list)
                    df_users = df_users[['usuario', 'nombre_completo', 'rol', 'activo', 'email', 'created_at', 'last_login']]
                    df_users.columns = ['Usuario', 'Nombre Completo', 'Rol', 'Activo', 'Email', 'Fecha Creación', 'Último Acceso']
                    st.dataframe(df_users, use_container_width=True, hide_index=True)
                else:
                    st.info("No hay usuarios registrados.")
                    
            st.divider()
            
            # --- ACCIONES SOBRE USUARIO EXISTENTE ---
            st.subheader("⚙️ Modificar o Administrar Usuario Registrado")
            if users_list:
                usernames_available = [u['usuario'] for u in users_list]
                sel_user = st.selectbox("Selecciona Usuario a Modificar:", usernames_available, key="sel_user_mod")
                
                user_info = next((u for u in users_list if u['usuario'] == sel_user), None)
                if user_info:
                    act_col1, act_col2, act_col3 = st.columns(3)
                    
                    with act_col1:
                        with st.expander("🎭 Cambiar Rol", expanded=True):
                            roles_all = ["Administrador", "Supervisor", "Analista Contable", "Analista de Reportes", "Auditor Lector"]
                            curr_idx = roles_all.index(user_info['rol']) if user_info['rol'] in roles_all else 1
                            is_admin_user = (sel_user == "admin")
                            new_role = st.selectbox(
                                "Nuevo Rol:", 
                                roles_all, 
                                index=curr_idx, 
                                key=f"sel_new_role_{sel_user}",
                                disabled=is_admin_user
                            )
                            if is_admin_user:
                                st.caption("🔒 El rol de la cuenta principal 'admin' es fijo (Administrador).")
                            else:
                                if st.button("Actualizar Rol", key=f"btn_update_role_{sel_user}", use_container_width=True):
                                    ok, msg = update_user_role(sel_user, new_role, admin_username=st.session_state.get('auth_user', 'admin'))
                                    if ok:
                                        st.success(f"✅ {msg}")
                                        st.rerun()
                                    else:
                                        st.error(f"❌ {msg}")
                                        
                    with act_col2:
                        with st.expander("🔑 Resetear Contraseña", expanded=True):
                            new_pass_val = st.text_input("Nueva Contraseña:", type="password", key=f"input_reset_pass_{sel_user}")
                            if st.button("Guardar Nueva Clave", key=f"btn_reset_pass_{sel_user}", use_container_width=True):
                                if new_pass_val:
                                    ok, msg = change_user_password(sel_user, new_pass_val, actor_username=st.session_state.get('auth_user', 'admin'))
                                    if ok:
                                        st.success(f"✅ {msg}")
                                        st.rerun()
                                    else:
                                        st.error(f"❌ {msg}")
                                else:
                                    st.warning("Escribe una contraseña válida.")
                                        
                    with act_col3:
                        with st.expander("🚫 Estado / Eliminar", expanded=True):
                            curr_status = user_info['activo']
                            toggle_label = "Deshabilitar Usuario" if curr_status else "Habilitar Usuario"
                            if st.button(toggle_label, key=f"btn_toggle_status_{sel_user}", disabled=(sel_user == "admin"), use_container_width=True):
                                ok, msg = update_user_status(sel_user, not curr_status, admin_username=st.session_state.get('auth_user', 'admin'))
                                if ok:
                                    st.success(f"✅ {msg}")
                                    st.rerun()
                                else:
                                    st.error(f"❌ {msg}")
                                    
                            if sel_user != "admin":
                                if st.button("🗑️ Eliminar Usuario", type="primary", key=f"btn_del_user_{sel_user}", use_container_width=True):
                                    ok, msg = delete_user(sel_user, admin_username=st.session_state.get('auth_user', 'admin'))
                                    if ok:
                                        st.success(f"✅ {msg}")
                                        st.rerun()
                                    else:
                                        st.error(f"❌ {msg}")

        with subtab_audit:
            st.subheader("📜 Regístro y Bitácora de Auditoría Global")
            st.markdown("Consulta en tiempo real todas las acciones de inicio de sesión, cambios de roles y modificaciones realizadas por los usuarios.")
            
            f_col1, f_col2 = st.columns(2)
            with f_col1:
                filter_user_input = st.text_input("Filtrar por Usuario:", placeholder="ej: admin").strip()
            with f_col2:
                filter_action_input = st.text_input("Filtrar por Acción:", placeholder="ej: LOGIN, CAMBIO_ROL").strip()
                
            logs = get_audit_logs(filter_user=filter_user_input or None, filter_action=filter_action_input or None, limit=300)
            if logs:
                df_logs = pd.DataFrame(logs)
                df_logs = df_logs[['id', 'fecha_hora', 'usuario', 'accion', 'entidad_id', 'detalles']]
                df_logs.columns = ['ID', 'Fecha y Hora', 'Usuario', 'Acción', 'Entidad / Objeto', 'Detalles del Evento']
                st.dataframe(df_logs, use_container_width=True, hide_index=True)
                
                # Exportar bitácora a Excel
                excel_bytes_audit = df_to_excel_bytes(df_logs, sheet_name="Bitacora_Auditoria")
                st.download_button(
                    label="📥 Exportar Bitácora a Excel (.xlsx)",
                    data=excel_bytes_audit,
                    file_name="Bitacora_Auditoria_InformesPro.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="dl_audit_logs"
                )
            else:
                st.info("No hay registros en la bitácora con los filtros aplicados.")

    with tab_danger:
        st.write(f"Atención: Las acciones aquí realizadas aplicarán a la empresa actualmente activa: **{empresa_seleccionada}**.")
        
        danger_col1, danger_col2 = st.columns(2)
        
        with danger_col1:
            with st.expander("🗑️ Borrar Plan de Cuentas"):
                st.warning("Esto eliminará el archivo del servidor y la memoria caché.")
                check_plan = st.checkbox("Entiendo que esto es irreversible", key="check_del_plan")
                if st.button("Borrar Plan de Cuentas", type="primary", disabled=not check_plan):
                    plan_path = os.path.join(empresa_path, "plan_cuentas.xlsx")
                    if os.path.exists(plan_path): os.remove(plan_path)
                    st.session_state.pop('plan_cuentas_df', None)
                    st.session_state['success_msg'] = "✅ Plan de Cuentas Maestro eliminado y formateado exitosamente de la base de datos."
                    st.rerun()
                
        with st.expander("🗑️ Borrar Mapeos F/S (Balance y P&L)"):
            st.warning("Elimina Mapeos, Diccionarios y borra la Bóveda Maestra Taxonómica de esta empresa.")
            check_map = st.checkbox("Entiendo que esto es irreversible", key="check_del_map")
            if st.button("Resetear Mapeos F/S", type="primary", disabled=not check_map):
                # Archivos Fisicos
                for map_f in ["map_balance.xlsx", "map_pl.xlsx"]:
                    mp = os.path.join(empresa_path, map_f)
                    if os.path.exists(mp): os.remove(mp)
                # DB Taxonomia
                try:
                    from src.models.database import SessionLocal
                    from src.models.taxonomy_master import TaxonomyMasterRecord
                    db = SessionLocal()
                    db.query(TaxonomyMasterRecord).filter(TaxonomyMasterRecord.empresa == empresa_seleccionada).delete()
                    db.commit()
                    db.close()
                except Exception as e:
                    pass
                # Cache Amnesia
                st.session_state.pop('map_balance_df', None)
                st.session_state.pop('map_pl_df', None)
                st.session_state['success_msg'] = "✅ Mapeos y bóveda de Taxonomía formateados exitosamente. El cerebro del programa ha olvidado esos cruces."
                st.rerun()

        # Si la empresa activa es Global, permitir elegir a qué empresa aplicar la eliminación
        is_global = "GLOBAL" in empresa_seleccionada
        if is_global:
            empresas_dir = os.path.join("data", "empresas")
            real_cos = sorted([d for d in os.listdir(empresas_dir) if os.path.isdir(os.path.join(empresas_dir, d))])
            target_empresa = st.selectbox("🏢 Selecciona la empresa objetivo para eliminar datos:", real_cos, key="del_target_empresa")
            target_empresa_path = os.path.join(empresas_dir, target_empresa)
        else:
            target_empresa = empresa_seleccionada
            target_empresa_path = empresa_path

        with danger_col2:
            with st.expander("🗑️ Borrar Mes (Transaccional)"):
                st.warning(f"Extirpa los saldos del Balance y del Cubo P&L de un periodo específico para **{target_empresa}**.")
                from src.models.trial_balance_db import TrialBalanceDB
                from src.models.pl_cubo_db import PlCuboDB
                from src.models.database import SessionLocal
                from src.models.historical_data import HistoricalDataRecord
                
                per_avail = set()
                try:
                    per_tb = TrialBalanceDB.get_available_periods(target_empresa)
                    per_pl = PlCuboDB.get_available_periods(target_empresa)
                    per_avail.update(per_tb)
                    per_avail.update(per_pl)
                    
                    db = SessionLocal()
                    hist_recs = db.query(HistoricalDataRecord.periodo).filter(
                        (HistoricalDataRecord.empresa == target_empresa) | 
                        (HistoricalDataRecord.empresa == target_empresa.replace("[GRUPO] ", ""))
                    ).distinct().all()
                    if not hist_recs and target_empresa.startswith("[GRUPO]"):
                        hist_recs = db.query(HistoricalDataRecord.periodo).distinct().all()
                    for r in hist_recs:
                        if r[0]: per_avail.add(r[0])
                    db.close()
                except Exception:
                    pass
                
                # Buscar también periodos en archivos físicos
                if os.path.exists(target_empresa_path):
                    for f in os.listdir(target_empresa_path):
                        if f.startswith("pl_cubo_") and f.endswith(".xlsx"):
                            p_str = f.replace("pl_cubo_", "").replace(".xlsx", "")
                            if len(p_str) == 7 and "-" in p_str:
                                per_avail.add(p_str)
                
                per_list = sorted(list(per_avail), reverse=True)
                
                if not per_list:
                    st.info(f"No hay meses guardados para {target_empresa}.")
                else:
                    per_to_del = st.selectbox("Periodo a Eliminar", per_list, key=f"sel_del_per_{target_empresa}", format_func=format_periodo)
                    check_per = st.checkbox(f"Entiendo que borraré todos los datos de {format_periodo(per_to_del)} en {target_empresa}", key=f"check_del_per_{target_empresa}")
                    if st.button(f"Aniquilar transacciones de {format_periodo(per_to_del)}", type="primary", disabled=not check_per):
                        try:
                            db = SessionLocal()
                            # 1. Borrar Trial Balance
                            from src.models.trial_balance import TrialBalanceRecord
                            db.query(TrialBalanceRecord).filter(
                                (TrialBalanceRecord.empresa == target_empresa) & 
                                (TrialBalanceRecord.periodo == per_to_del)
                            ).delete()
                            
                            # 2. Borrar P&L Cubo
                            from src.models.pl_record import PlRecordDim
                            db.query(PlRecordDim).filter(
                                (PlRecordDim.empresa == target_empresa) & 
                                (PlRecordDim.periodo == per_to_del)
                            ).delete()
                            
                            # 3. Borrar Memoria Histórica
                            db.query(HistoricalDataRecord).filter(
                                ((HistoricalDataRecord.empresa == target_empresa) | 
                                 (HistoricalDataRecord.empresa == target_empresa.replace("[GRUPO] ", ""))) & 
                                (HistoricalDataRecord.periodo == per_to_del)
                            ).delete()
                            from src.models.historical_data import HistoricalDetailRecord
                            db.query(HistoricalDetailRecord).filter(
                                (HistoricalDetailRecord.empresa == target_empresa) & 
                                (HistoricalDetailRecord.periodo == per_to_del)
                            ).delete()
                            
                            # 4. Borrar Asientos de Consolidación si aplica
                            try:
                                from src.models.consolidacion import JournalEntryModel, JournalEntryLineModel
                                je_entries = db.query(JournalEntryModel.id).filter(
                                    (JournalEntryModel.periodo == per_to_del) & 
                                    (JournalEntryModel.grupo_name == empresa_seleccionada)
                                ).all()
                                je_ids = [je[0] for je in je_entries]
                                if je_ids:
                                    db.query(JournalEntryLineModel).filter(JournalEntryLineModel.entry_id.in_(je_ids)).delete(synchronize_session=False)
                                    db.query(JournalEntryModel).filter(JournalEntryModel.id.in_(je_ids)).delete(synchronize_session=False)
                            except Exception:
                                pass
                                
                            db.commit()
                            db.close()
                        except Exception as e:
                            pass
                        
                        # Archivos Físicos P&L y tb temporal
                        pl_hist = os.path.join(empresa_path, f"pl_cubo_{per_to_del}.xlsx")
                        if os.path.exists(pl_hist): os.remove(pl_hist)
                        
                        # Limpieza de caché de sesión
                        st.session_state.pop('tb_df', None)
                        st.session_state.pop('pl_df', None)
                        
                        temp_tb = os.path.join(empresa_path, "temp_uploaded.xlsx")
                        if os.path.exists(temp_tb): os.remove(temp_tb)
                        temp_pl = os.path.join(empresa_path, "pl_cubo.xlsx")
                        if os.path.exists(temp_pl): os.remove(temp_pl)
                        
                        st.session_state['success_msg'] = f"✅ Datos transaccionales del periodo {format_periodo(per_to_del)} eliminados con éxito del servidor y memoria activa."
                        st.rerun()

