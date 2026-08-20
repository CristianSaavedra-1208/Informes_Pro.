import streamlit as st
import pandas as pd
import os
from src.core.excel_utils import df_to_excel_bytes, format_periodo

def render(empresa_seleccionada, empresa_path):
    # Vista simplificada exclusiva para el rol Analista de Reportes
    if st.session_state.get('auth_role') == "Analista de Reportes":
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
    
    tab_empresas, tab_erp, tab_roles, tab_danger = st.tabs(["Empresas y Entornos", "Conexiones ERP (API)", "Roles & Settings", "Eliminación de Data"])
    
    with tab_empresas:
        col1, col2 = st.columns(2)
        
        with col1:
            with st.expander("➕ Crear Nueva Empresa", expanded=True):
                nueva_empresa = st.text_input("Nombre de la Empresa")
                if st.button("Crear Empresa"):
                    if nueva_empresa.strip():
                        nueva_carpeta = os.path.join(empresas_dir, nueva_empresa.strip())
                        if not os.path.exists(nueva_carpeta):
                            os.makedirs(nueva_carpeta, exist_ok=True)
                            
                            # Copiar plantillas default desde templates a la nueva empresa
                            import shutil
                            templates_dir = "templates"
                            if os.path.exists(templates_dir):
                                for t_file in os.listdir(templates_dir):
                                    if t_file.endswith(".xlsx"):
                                        shutil.copy2(os.path.join(templates_dir, t_file), os.path.join(nueva_carpeta, t_file))
                                        
                            # Copiar archivos maestros globales de configuracion si existen
                            global_master_dir = os.path.join(empresas_dir, "Pacifico SpA")
                            if os.path.exists(global_master_dir):
                                for master_f in ["plan_cuentas.xlsx", "map_balance.xlsx", "map_pl.xlsx"]:
                                    src_master = os.path.join(global_master_dir, master_f)
                                    if os.path.exists(src_master):
                                        shutil.copy2(src_master, os.path.join(nueva_carpeta, master_f))

                            st.success(f"Empresa '{nueva_empresa}' creada exitosamente.")
                            st.rerun()
                        else:
                            st.error("La empresa ya existe.")

    with col2:
        with st.expander("✏️ Renombrar Empresa", expanded=True):
            if empresas:
                empresa_a_renombrar = st.selectbox("Selecciona empresa a modificar", empresas, key="rename_select")
                nuevo_nombre = st.text_input("Nuevo nombre de la empresa", key="rename_input")
                if st.button("Renombrar Empresa"):
                    if nuevo_nombre.strip():
                        if nuevo_nombre.strip() not in empresas:
                            old_path = os.path.join(empresas_dir, empresa_a_renombrar)
                            new_path = os.path.join(empresas_dir, nuevo_nombre.strip())
                            try:
                                os.rename(old_path, new_path)
                                # Limpiar cache local forzosamente
                                for key in ['plan_cuentas_df', 'tb_df', 'map_balance_df', 'map_pl_df', 'pl_df', 'er_preview_df']:
                                    if key in st.session_state:
                                        del st.session_state[key]
                                
                                if st.session_state.get('empresa_activa') == empresa_a_renombrar:
                                    st.session_state['empresa_activa'] = nuevo_nombre.strip()
                                    st.session_state['empresa_activa_prev'] = nuevo_nombre.strip()
                                    
                                st.success(f"Empresa renombrada a '{nuevo_nombre.strip()}' correctamente.")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Error al renombrar: {e}")
                        else:
                            st.error("El nombre ya existe.")
                    else:
                        st.warning("El nuevo nombre no puede estar vacío.")
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
                        nu_role = st.selectbox("Rol Asignado:", ["Administrador", "Analista Contable", "Analista de Reportes", "Auditor Lector"])
                        
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
                            roles_all = ["Administrador", "Analista Contable", "Analista de Reportes", "Auditor Lector"]
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

    with danger_col2:
        with st.expander("🗑️ Borrar Mes (Transaccional)"):
            st.warning("Extirpa los saldos del Balance y del Cubo P&L de un periodo específico.")
            from src.models.trial_balance_db import TrialBalanceDB
            from src.models.pl_cubo_db import PlCuboDB
            try:
                per_tb = TrialBalanceDB.get_available_periods(empresa_seleccionada)
                per_pl = PlCuboDB.get_available_periods(empresa_seleccionada)
                per_avail = sorted(list(set(per_tb + per_pl)), reverse=True)
            except Exception:
                per_avail = []
            
            if not per_avail:
                st.info("No hay meses guardados en el archivo histórico.")
            else:
                per_to_del = st.selectbox("Periodo a Eliminar", per_avail, key="sel_del_per", format_func=format_periodo)
                check_per = st.checkbox(f"Entiendo que borraré todos los datos de {format_periodo(per_to_del)}", key="check_del_per")
                if st.button(f"Aniquilar transacciones de {format_periodo(per_to_del)}", type="primary", disabled=not check_per):
                    # DB Trial Balance
                    try:
                        from src.models.database import SessionLocal
                        from src.models.trial_balance import TrialBalanceRecord
                        db = SessionLocal()
                        db.query(TrialBalanceRecord).filter(TrialBalanceRecord.empresa == empresa_seleccionada, TrialBalanceRecord.periodo == per_to_del).delete()
                        from src.models.pl_record import PlRecordDim
                        db.query(PlRecordDim).filter(PlRecordDim.empresa == empresa_seleccionada, PlRecordDim.periodo == per_to_del).delete()
                        db.commit()
                        db.close()
                    except Exception as e:
                        pass
                    
                    # Archivos Fisicos P&L y tb temporal
                    pl_hist = os.path.join(empresa_path, f"pl_cubo_{per_to_del}.xlsx")
                    if os.path.exists(pl_hist): os.remove(pl_hist)
                    
                    # Cache Amnesia
                    st.session_state.pop('tb_df', None)
                    st.session_state.pop('pl_df', None)
                    
                    temp_tb = os.path.join(empresa_path, "temp_uploaded.xlsx")
                    if os.path.exists(temp_tb): os.remove(temp_tb)
                    temp_pl = os.path.join(empresa_path, "pl_cubo.xlsx")
                    if os.path.exists(temp_pl): os.remove(temp_pl)
                    
                    st.session_state['success_msg'] = f"✅ Datos transaccionales del periodo {format_periodo(per_to_del)} eliminados con éxito del servidor y memoria activa."
                    st.rerun()

