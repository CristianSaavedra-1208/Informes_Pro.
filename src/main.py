# Streamlit hot-reload trigger - force reload 2
import streamlit as st
import pandas as pd
import os
import sys
import io

# Asegurar importabilidad de src y consistencia de rutas relativas
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(ROOT_DIR)
sys.path.append(ROOT_DIR)
from src.core.excel_utils import df_to_excel_bytes, sort_accounts, heal_mapping_fields, read_excel_cached
from src.ingestion.trial_balance import TrialBalanceIngestor
from src.core.mapping import MappingEngine
from src.core.rules import AccountingRulesEngine
from src.core.validation import ValidationEngine
from src.core.tie_out import TieOutEngine

def main():
    st.set_page_config(page_title="Informes Pro", layout="wide", page_icon="📊")

    # Inyección CSS para Menú Lateral Tailwind & Iconografía Monocromática
    st.markdown("""
        <style>
        /* 1. Encabezado de Categorías (OPERACIONES, REPORTES, ADMINISTRACIÓN) */
        div[data-testid="stSidebarNav"] span[data-testid="stSidebarNavHeader"] {
            font-size: 11px !important;
            font-weight: 700 !important;
            letter-spacing: 0.08em !important;
            color: #94a3b8 !important;
            text-transform: uppercase !important;
            padding-top: 14px !important;
            padding-bottom: 4px !important;
        }

        /* 2. Enlaces del Menú Lateral */
        div[data-testid="stSidebarNav"] a {
            display: flex !important;
            align-items: center !important;
            gap: 10px !important;
            padding: 8px 12px !important;
            border-radius: 8px !important;
            font-size: 13.5px !important;
            font-weight: 500 !important;
            color: #475569 !important;
            text-decoration: none !important;
            transition: all 0.15s ease-in-out !important;
        }

        /* 3. Iconos Monocromáticos (Slate-400 para inactivos) */
        div[data-testid="stSidebarNav"] a span[data-testid="stIconMaterial"] {
            color: #94a3b8 !important;
            font-size: 19px !important;
            transition: color 0.15s ease-in-out !important;
        }

        /* 4. Hover State */
        div[data-testid="stSidebarNav"] a:hover {
            background-color: #f1f5f9 !important;
            color: #0f172a !important;
        }
        div[data-testid="stSidebarNav"] a:hover span[data-testid="stIconMaterial"] {
            color: #475569 !important;
        }

        /* 5. Elemento Activo (bg-blue-50, text-blue-700, shadow-sm) */
        div[data-testid="stSidebarNav"] a[aria-current="page"] {
            background-color: #eff6ff !important;
            color: #1d4ed8 !important;
            font-weight: 600 !important;
            box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.06), 0 1px 2px 0 rgba(0, 0, 0, 0.04) !important;
        }
        div[data-testid="stSidebarNav"] a[aria-current="page"] span[data-testid="stIconMaterial"] {
            color: #1d4ed8 !important;
        }
        div[data-testid="stSidebarNav"] a[aria-current="page"] span {
            color: #1d4ed8 !important;
            font-weight: 600 !important;
        }

        /* Marca de Encabezado Superior (Blue Dot + INFORMES PRO) */
        .brand-header-title {
            display: flex;
            align-items: center;
            gap: 8px;
            font-size: 13.5px;
            font-weight: 800;
            color: #0f172a;
            letter-spacing: 0.05em;
            padding: 2px 4px 10px 4px;
            border-bottom: 1px solid #e2e8f0;
            margin-bottom: 10px;
        }
        .brand-dot {
            width: 8px;
            height: 8px;
            background-color: #2563eb;
            border-radius: 50%;
            display: inline-block;
        }
        </style>
    """, unsafe_allow_html=True)

    # Asegurar que el directorio de persistencia e inicializar base de datos
    os.makedirs("data", exist_ok=True)
    try:
        from src.models.database import init_db
        init_db()
    except Exception as e:
        print(f"Error al inicializar base de datos: {e}")

    # --- GUARDIA DE AUTENTICACION (LOGIN) ---
    if not st.session_state.get('is_authenticated', False):
        from src.ui_pages.login import render_login
        render_login()
        return

    # Migración automática de plantillas: agregar hoja _CONFIG_ si no existe
    try:
        from migrate_templates import run_migration_if_needed
        run_migration_if_needed(verbose=False)
    except Exception:
        pass  # No interrumpir la app si la migración falla

    # --- GESTOR DE COMPAÑÍAS ---
    empresas_dir = os.path.join("data", "empresas")
    os.makedirs(empresas_dir, exist_ok=True)

    # Migración de legado
    legacy_files = [f for f in os.listdir("data") if f.endswith(".xlsx") and os.path.isfile(os.path.join("data", f))]
    if legacy_files and not os.listdir(empresas_dir):
        default_dir = os.path.join(empresas_dir, "Pacifico")
        os.makedirs(default_dir, exist_ok=True)
        import shutil
        for f in legacy_files:
            try:
                shutil.move(os.path.join("data", f), os.path.join(default_dir, f))
            except:
                pass

    st.sidebar.markdown("""
        <div class="brand-header-title">
            <span class="brand-dot"></span>
            <span>INFORMES PRO</span>
        </div>
    """, unsafe_allow_html=True)
    
    # Usuario Logueado y Botón de Cerrar Sesión directo
    st.sidebar.markdown(f"**Usuario:** {st.session_state.get('auth_name', 'Usuario')} (`{st.session_state.get('auth_role')}`)")
    if st.sidebar.button("Cerrar Sesión", use_container_width=True, key="btn_logout_sidebar"):
        from src.core.security_engine import register_audit_log
        register_audit_log(st.session_state.get('auth_user'), "LOGOUT", detalles="Cierre de sesión manual.")
        for k in ['is_authenticated', 'auth_user', 'auth_role', 'auth_name']:
            st.session_state.pop(k, None)
        st.rerun()

    st.sidebar.subheader("🏢 Empresa Activa")
    
    # Cargar empresas para armar el selector principal
    real_empresas = sorted([d for d in os.listdir(empresas_dir) if os.path.isdir(os.path.join(empresas_dir, d))])
    
    if not real_empresas:
        st.sidebar.warning("Crea una empresa.")
        st.warning("No hay empresas creadas. Por favor, crea una nueva empresa desde el menú lateral.")
        st.stop()
        
    user_role = st.session_state.get('auth_role', 'Analista Contable')

    global_opt = "[GLOBAL] Configuración General"
    if user_role == "Analista de Reportes":
        empresas = real_empresas
        if st.session_state.get('empresa_activa') == global_opt or 'empresa_activa' not in st.session_state:
            st.session_state['empresa_activa'] = real_empresas[0] if real_empresas else ""
    else:
        empresas = [global_opt] + real_empresas
    
    last_active_path = os.path.join("data", "last_active.txt")
    if 'empresa_activa' not in st.session_state:
        if os.path.exists(last_active_path):
            with open(last_active_path, "r", encoding="utf-8") as f:
                last_co = f.read().strip()
            if last_co in empresas:
                st.session_state['empresa_activa'] = last_co
            else:
                st.session_state['empresa_activa'] = empresas[0]
        else:
            st.session_state['empresa_activa'] = empresas[0]

    empresa_activa_actual = st.session_state['empresa_activa']
    empresa_idx = empresas.index(empresa_activa_actual) if empresa_activa_actual in empresas else 0

    empresa_seleccionada_combo = st.sidebar.selectbox(
        "Selecciona la empresa de trabajo:",
        empresas,
        index=empresa_idx,
        key="selector_empresa"
    )

    if empresa_seleccionada_combo != st.session_state['empresa_activa']:
        st.sidebar.info(f"Selección pendiente:\n**{empresa_seleccionada_combo}**")
        if st.sidebar.button("Aplicar Cambio de Empresa", type="primary", use_container_width=True):
            auth_keys = {'is_authenticated', 'auth_user', 'auth_role', 'auth_name', 'auth_email', 'empresa_activa', 'selector_empresa'}
            for key in list(st.session_state.keys()):
                if key not in auth_keys:
                    del st.session_state[key]
            st.session_state['empresa_activa'] = empresa_seleccionada_combo
            with open(last_active_path, "w", encoding="utf-8") as f:
                f.write(empresa_seleccionada_combo)
            st.session_state['success_msg'] = f"Cambio de empresa efectuado exitosamente: **{empresa_seleccionada_combo}**"
            st.rerun()
    else:
        st.sidebar.success(f"**Empresa activa:**\n\n{st.session_state['empresa_activa']}")

    empresa_seleccionada = st.session_state['empresa_activa']

    if empresa_seleccionada == global_opt:
        empresa_path = os.path.join(empresas_dir, "Pacifico SpA")
    else:
        empresa_path = os.path.join(empresas_dir, empresa_seleccionada)

    # Inicializar datos persistentes desde disco si no están en sesión
    if 'plan_cuentas_df' not in st.session_state:
        plan_path = os.path.join(empresa_path, "plan_cuentas.xlsx")
        if not os.path.exists(plan_path):
            # Fallback automático al Plan de Cuentas Global Maestro
            plan_path = os.path.join(empresas_dir, "Pacifico SpA", "plan_cuentas.xlsx")
        if os.path.exists(plan_path):
            try:
                df_plan = read_excel_cached(plan_path, dtype=str)
                df_plan = sort_accounts(df_plan, 'Cuenta', 'Tipo')
                st.session_state['plan_cuentas_df'] = df_plan
            except Exception:
                pass

    if empresa_seleccionada != global_opt:
        if 'tb_df' not in st.session_state:
            from src.models.trial_balance_db import TrialBalanceDB
            try:
                periodos = TrialBalanceDB.get_available_periods(empresa_seleccionada)
                if periodos:
                    st.session_state['tb_df'] = TrialBalanceDB.get_trial_balance(empresa_seleccionada, periodos[-1])
            except Exception:
                pass

    if 'map_balance_df' not in st.session_state:
        map_bal_path = os.path.join(empresa_path, "map_balance.xlsx")
        if os.path.exists(map_bal_path):
            try:
                df_bal = read_excel_cached(map_bal_path, dtype=str, engine='openpyxl')
                df_bal = sort_accounts(df_bal, 'N° de Cuenta')
                
                # Sanar columna de flujo de efectivo si falta por completo
                cf_col = next((c for c in df_bal.columns if 'flujo' in c.lower() and 'efectivo' in c.lower()), None)
                if not cf_col:
                    ref_path = os.path.join("data", "empresas", "Pacifico SpA", "map_balance.xlsx")
                    if os.path.exists(ref_path):
                        ref_df = pd.read_excel(ref_path, dtype=str)
                        ref_cf_col = next((c for c in ref_df.columns if 'flujo' in c.lower() and 'efectivo' in c.lower()), "Clasificación Flujo Efectivo")
                        # Crear mapa de referencia
                        map_dict = {}
                        for _, r in ref_df.iterrows():
                            b_val = str(r.get("Clasificación balance", "")).strip().lower()
                            cf_val = str(r.get(ref_cf_col, "")).strip()
                            if b_val and b_val != "nan" and cf_val and cf_val != "nan" and cf_val != "":
                                map_dict[b_val] = cf_val
                        # Mapear
                        new_cf_vals = []
                        for _, r in df_bal.iterrows():
                            b_val = str(r.get("Clasificación balance", "")).strip().lower()
                            new_cf_vals.append(map_dict.get(b_val, ""))
                        df_bal["Clasificación Flujo Efectivo"] = new_cf_vals
                        df_bal = sort_accounts(df_bal, 'N° de Cuenta')

                # Sanar celdas de taxonomía y flujo vacías
                df_bal_healed = heal_mapping_fields(df_bal)
                has_changes = not df_bal_healed.equals(df_bal)
                df_bal = df_bal_healed
                
                st.session_state['map_balance_df'] = df_bal
                if has_changes:
                    df_bal.to_excel(map_bal_path, index=False)

                # Sincronizar taxonomía en base de datos
                from src.models.taxonomy_generator import process_mapping_for_taxonomy
                if empresa_seleccionada == global_opt:
                    for co in real_empresas:
                        process_mapping_for_taxonomy(st.session_state['map_balance_df'], "Balance", co)
                else:
                    process_mapping_for_taxonomy(st.session_state['map_balance_df'], "Balance", empresa_seleccionada)
            except Exception:
                pass

    if 'map_pl_df' not in st.session_state:
        map_pl_path = os.path.join(empresa_path, "map_pl.xlsx")
        if os.path.exists(map_pl_path):
            try:
                df_pl = read_excel_cached(map_pl_path, dtype=str)
                df_pl = sort_accounts(df_pl, 'N° de cuenta')
                st.session_state['map_pl_df'] = df_pl
                # Sincronizar taxonomía en base de datos
                from src.models.taxonomy_generator import process_mapping_for_taxonomy
                if empresa_seleccionada == global_opt:
                    for co in real_empresas:
                        process_mapping_for_taxonomy(df_pl, "PL", co)
                else:
                    process_mapping_for_taxonomy(df_pl, "PL", empresa_seleccionada)
            except Exception:
                pass
                
    if empresa_seleccionada != global_opt:
        if 'pl_df' not in st.session_state:
            from src.models.pl_cubo_db import PlCuboDB
            try:
                per_pl = PlCuboDB.get_available_periods(empresa_seleccionada)
                if per_pl:
                    st.session_state['pl_df'] = PlCuboDB.get_pl_cubo(empresa_seleccionada, per_pl[-1])
            except Exception:
                pass
    

    import src.core.excel_utils as excel_utils
    import src.ui_pages.inicio as pg_inicio
    import src.ui_pages.cargas_de_datos as pg_cargas
    import src.ui_pages.organizacion_de_cuentas as pg_map
    import src.ui_pages.ajustes_manuales as pg_ajustes
    import src.ui_pages.consolidacion as pg_cons
    import src.ui_pages.estados_financieros as pg_estados
    import src.ui_pages.informes_y_notas as pg_informes
    import src.ui_pages.cierre_y_memoria_historica as pg_cierre
    import src.ui_pages.validacion_saldos as pg_val
    import src.ui_pages.parametros_globales as pg_param
    import src.ui_pages.configuraciones as pg_conf
    import src.ui_pages.reportes_consolidados as pg_reportes_cons
    import src.ui_pages.reporte_corporativo as pg_reporte_corp
    import src.ui_pages.auditoria_diccionario as pg_audit

    # Definir wrappers para st.navigation
    def run_inicio(): pg_inicio.render(empresa_seleccionada, empresa_path)
    def run_cargas(): pg_cargas.render(empresa_seleccionada, empresa_path)
    def run_map(): pg_map.render(empresa_seleccionada, empresa_path)
    def run_ajustes(): pg_ajustes.render(empresa_seleccionada, empresa_path)
    def run_cons(): pg_cons.render(empresa_seleccionada, empresa_path)
    def run_estados(): pg_estados.render(empresa_seleccionada, empresa_path)
    def run_informes(): pg_informes.render(empresa_seleccionada, empresa_path)
    def run_val(): pg_val.render(empresa_seleccionada, empresa_path)
    def run_cierre(): pg_cierre.render(empresa_seleccionada, empresa_path)
    def run_audit(): pg_audit.render(empresa_seleccionada, empresa_path)
    def run_param(): pg_param.render(empresa_seleccionada, empresa_path)
    def run_conf(): pg_conf.render(empresa_seleccionada, empresa_path)
    def run_reportes_cons(): pg_reportes_cons.render(empresa_seleccionada, empresa_path)
    def run_reporte_corp(): pg_reporte_corp.render(empresa_seleccionada, empresa_path)

    es_grupo = empresa_seleccionada.startswith("[GRUPO]")
    user_role = st.session_state.get('auth_role', 'Analista Contable')
    
    # Operaciones dinámicas
    operaciones_list = []
    if user_role not in ["Auditor Lector", "Analista de Reportes"]:
        if es_grupo:
            operaciones_list = [
                st.Page(run_map, title="Clasificación Cuentas", icon=":material/format_list_bulleted:"),
                st.Page(run_cons, title="Consolidación", icon=":material/domain:")
            ]
        else:
            operaciones_list = [
                st.Page(run_cargas, title="Carga de Datos", icon=":material/upload:"),
                st.Page(run_map, title="Clasificación Cuentas", icon=":material/format_list_bulleted:"),
                st.Page(run_ajustes, title="Ajustes Auditoría", icon=":material/edit_note:")
            ]
        
    # Reportes dinámicos
    reportes_list = []
    if es_grupo:
        reportes_list.append(st.Page(run_reportes_cons, title="Estados Fin. Consolidados", icon=":material/bar_chart:"))
    else:
        reportes_list.append(st.Page(run_estados, title="Estados Fin. Individuales", icon=":material/bar_chart:"))
        
    reportes_list.extend([
        st.Page(run_informes, title="Informes & Notas", icon=":material/description:"),
        st.Page(run_val, title="Validaciones", icon=":material/check_circle:"),
        st.Page(run_reporte_corp, title="Reportes Word", icon=":material/file_download:"),
        st.Page(run_audit, title="Auditoría", icon=":material/search:")
    ])
    
    # Administración dinámicas
    admin_list = []
    if user_role != "Analista de Reportes":
        admin_list = [
            st.Page(run_cierre, title="Históricos", icon=":material/schedule:"),
            st.Page(run_param, title="Parámetros", icon=":material/settings:")
        ]
        
    if user_role == "Administrador":
        admin_list.append(st.Page(run_conf, title="Roles & Settings", icon=":material/group:"))
    elif user_role == "Analista de Reportes":
        admin_list.append(st.Page(run_conf, title="Perfil y Clave", icon=":material/key:"))

    pages = {}
    if operaciones_list:
        pages["Operaciones"] = operaciones_list
    pages["Reportes"] = reportes_list
    if admin_list:
        pages["Administración"] = admin_list
    
    pg = st.navigation(pages)
    
    # --- GLOBAL FLASH MESSAGES ---
    if 'success_msg' in st.session_state and st.session_state['success_msg']:
        st.success(st.session_state.pop('success_msg'))
    if 'error_msg' in st.session_state and st.session_state['error_msg']:
        st.error(st.session_state.pop('error_msg'))

    pg.run()


if __name__ == "__main__":
    main()
