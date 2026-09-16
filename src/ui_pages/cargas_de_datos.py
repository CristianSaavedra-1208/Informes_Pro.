import streamlit as st
import pandas as pd
import os
import re
from src.core.excel_utils import df_to_excel_bytes, sort_accounts, propagate_global_file
from src.models.pl_cubo_db import PlCuboDB

def parse_numeric_cell(val):
    """
    Convierte cualquier formato numérico (texto con puntos de miles, comas decimales,
    números negativos en paréntesis o con signo, símbolos de moneda) a float estándar.
    """
    if pd.isna(val) or val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if not s or s.lower() == 'nan' or s == '-':
        return 0.0
    is_neg = False
    if (s.startswith('(') and s.endswith(')')) or s.startswith('-'):
        is_neg = True
        s = s.replace('(', '').replace(')', '').replace('-', '')
    s = s.replace('$', '').replace('CLP', '').replace(' ', '').replace("'", "")
    
    # Manejar formatos de miles y decimales
    if '.' in s and ',' in s:
        if s.rfind(',') > s.rfind('.'):
            # Formato chileno/europeo: 1.234.567,89
            s = s.replace('.', '').replace(',', '.')
        else:
            # Formato anglosajón: 1,234,567.89
            s = s.replace(',', '')
    elif '.' in s and ',' not in s:
        parts = s.split('.')
        # Si tiene más de un punto (1.234.567) o un punto seguido de 3 dígitos (1.000)
        if len(parts) > 2 or (len(parts) == 2 and len(parts[1]) == 3):
            s = s.replace('.', '')
    elif ',' in s and '.' not in s:
        parts = s.split(',')
        if len(parts) == 2 and len(parts[1]) <= 2:
            s = s.replace(',', '.')
        else:
            s = s.replace(',', '')
            
    try:
        num = float(s)
        return -num if is_neg else num
    except Exception:
        return 0.0

def add_totals_row(df):
    if df is None or df.empty:
        return df
    
    # Identify key columns
    cuenta_col = next((c for c in df.columns if "cuenta" in str(c).lower() and "nombre" not in str(c).lower()), "N° de cuenta")
    desc_col = next((c for c in df.columns if "nombre" in str(c).lower()), "Nombre de la cuenta")
    
    # Filter out any existing TOTAL row
    df_clean = df[df[cuenta_col].astype(str).str.strip().str.upper() != "TOTAL"].copy()
    
    # Identify numeric columns to sum (all columns except account and description)
    non_numeric = [cuenta_col, desc_col]
    numeric_cols = [c for c in df.columns if c not in non_numeric]
    
    # Convert numeric columns to float/int in df_clean to ensure proper numeric type
    for col in numeric_cols:
        df_clean[col] = df_clean[col].apply(parse_numeric_cell)
    
    # Calculate sum
    totals = {}
    for col in numeric_cols:
        totals[col] = df_clean[col].sum()
        
    total_row = {
        cuenta_col: "TOTAL",
        desc_col: "TOTAL GENERAL"
    }
    for col in numeric_cols:
        total_row[col] = totals[col]
        
    df_total = pd.concat([df_clean, pd.DataFrame([total_row])], ignore_index=True)
    return df_total

def render_balance_vs_pl_validation_card(empresa: str, periodo: str, pl_df_override=None, tb_df_override=None):
    """
    Despliega la tarjeta de auditoría cruzada entre Resultados Acumulados del Balance y Cubo P&L con diseño compacto y tipografía ajustada.
    """
    try:
        from src.core.validation_tie_out import ValidationTieOutEngine
        if pl_df_override is None and 'pl_edit_df' in st.session_state:
            pl_df_override = st.session_state.get('pl_edit_df')
        val = ValidationTieOutEngine.validar_balance_vs_pl(empresa, periodo, pl_df_override=pl_df_override, tb_df_override=tb_df_override)
        if val.get("has_tb") and val.get("has_pl"):
            s_bal = f"${val['saldo_balance']:,.0f}".replace(",", ".")
            s_pl = f"${val['saldo_pl']:,.0f}".replace(",", ".")
            s_diff = f"${val['diferencia']:,.0f}".replace(",", ".")
            
            if val["is_cuadrado"]:
                badge = '<span style="background-color:#d4edda; color:#155724; padding:3px 9px; border-radius:10px; font-weight:600; font-size:12px;">🟢 100% CUADRADO</span>'
                border_color = '#28a745'
            else:
                badge = '<span style="background-color:#f8d7da; color:#721c24; padding:3px 9px; border-radius:10px; font-weight:600; font-size:12px;">🔴 DESCUADRE DETECTADO</span>'
                border_color = '#dc3545'

            st.markdown(f"""
            <div style="border: 1px solid {border_color}; border-radius: 8px; padding: 8px 14px; background-color: #fafbfc; margin: 8px 0 12px 0;">
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px;">
                    <span style="font-weight: 600; font-size: 13px; color: #2c3e50;">⚖️ Control de Cuadratura: Balance vs. Cubo P&L ({periodo})</span>
                    {badge}
                </div>
                <div style="display: flex; gap: 20px; font-size: 12px; color: #555; flex-wrap: wrap;">
                    <div><span style="color:#666;">Resultados Acum. (Balance):</span> <strong style="color:#222; font-size:12px;">{s_bal}</strong></div>
                    <div><span style="color:#666;">Total Cubo P&L:</span> <strong style="color:#222; font-size:12px;">{s_pl}</strong></div>
                    <div><span style="color:#666;">Diferencia:</span> <strong style="color:#222; font-size:12px;">{s_diff}</strong></div>
                </div>
            </div>
            """, unsafe_allow_html=True)
    except Exception:
        pass

def check_unmapped_accounts(empresa_path, accounts_to_check, plan_cuentas_df):
    """
    Verifica si las cuentas provistas están mapeadas en las plantillas.
    Retorna el set de cuentas sin clasificar usando operaciones vectorizadas y set algebra.
    """
    if plan_cuentas_df is None or plan_cuentas_df.empty:
        return set()
        
    plan_cuentas = set(plan_cuentas_df['Cuenta'].astype(str).str.strip())
    from src.core.excel_utils import read_excel_cached
    
    # 1. Cargar mapeos de Balance
    map_bal_path = os.path.join(empresa_path, "map_balance.xlsx")
    mapped_bal_accounts = set()
    if os.path.exists(map_bal_path):
        try:
            df_bal = read_excel_cached(map_bal_path, dtype=str)
            if df_bal is not None and not df_bal.empty:
                col_bal_cuenta = next((c for c in df_bal.columns if 'cuenta' in c.lower()), df_bal.columns[0])
                clasif_col = next((c for c in df_bal.columns if 'clasificaci' in c.lower() and 'balance' in c.lower()), None)
                if clasif_col:
                    valid_mask = df_bal[clasif_col].notna() & (df_bal[clasif_col].astype(str).str.strip() != "") & (df_bal[clasif_col].astype(str).str.lower() != "nan")
                    mapped_bal_accounts = set(df_bal.loc[valid_mask, col_bal_cuenta].astype(str).str.strip())
        except Exception:
            pass

    # 2. Cargar mapeos de P&L
    map_pl_path = os.path.join(empresa_path, "map_pl.xlsx")
    mapped_pl_accounts = set()
    # Add overrides
    pl_overrides = {"3105301", "3105302", "3105702", "3105703", "3105711", "3105834", "3105835", "3108112", "3103111", "3103113", "3103112", "3103122", "3105704"}
    mapped_pl_accounts.update(pl_overrides)
    if os.path.exists(map_pl_path):
        try:
            df_pl = read_excel_cached(map_pl_path, dtype=str)
            if df_pl is not None and not df_pl.empty:
                col_pl_cuenta = next((c for c in df_pl.columns if 'cuenta' in c.lower()), df_pl.columns[0])
                data_cols = [c for c in df_pl.columns if c != col_pl_cuenta]
                if data_cols:
                    clean_vals = df_pl[data_cols].fillna("").astype(str)
                    has_mapping_mask = clean_vals.apply(lambda col: (col.str.strip() != "") & (col.str.lower() != "nan")).any(axis=1)
                    mapped_pl_accounts.update(df_pl.loc[has_mapping_mask, col_pl_cuenta].astype(str).str.strip())
        except Exception:
            pass

    # 3. Filtrado vectorizado mediante álgebra de conjuntos
    clean_accounts = {
        str(acc).strip() for acc in accounts_to_check
        if str(acc).strip() and str(acc).strip().upper() != 'TOTAL' and str(acc).strip().lower() != 'nan'
    }
    return (clean_accounts & plan_cuentas) - (mapped_bal_accounts | mapped_pl_accounts)

def render(empresa_seleccionada, empresa_path):
    global_opt = "[GLOBAL] Configuración General"
    is_global = ("GLOBAL" in empresa_seleccionada or empresa_seleccionada == global_opt)
        
    st.title("Carga de Datos")
    st.write("Centraliza la carga de todos los insumos necesarios para el ciclo contable.")
    
    if "success_msg" in st.session_state:
        st.success(st.session_state.pop("success_msg"))
    
    # CSS para pestañas persistentes con estilo moderno
    st.markdown("""
        <style>
        div[data-testid="stRadio"] > div[role="radiogroup"] {
            display: flex;
            gap: 10px;
            background-color: #f1f5f9;
            padding: 6px;
            border-radius: 10px;
            border: 1px solid #e2e8f0;
            margin-bottom: 15px;
        }
        div[data-testid="stRadio"] > div[role="radiogroup"] > label {
            background-color: transparent;
            padding: 8px 18px;
            border-radius: 8px;
            font-weight: 600;
            color: #475569;
            cursor: pointer;
            transition: all 0.2s ease;
            border: none !important;
            margin: 0 !important;
        }
        div[data-testid="stRadio"] > div[role="radiogroup"] > label:hover {
            background-color: #e2e8f0;
            color: #0f172a;
        }
        div[data-testid="stRadio"] > div[role="radiogroup"] > label[data-checked="true"],
        div[data-testid="stRadio"] > div[role="radiogroup"] > label:has(input:checked) {
            background-color: #2563eb !important;
            color: #ffffff !important;
            box-shadow: 0 2px 6px rgba(37, 99, 235, 0.3);
        }
        div[data-testid="stRadio"] > div[role="radiogroup"] > label:has(input:checked) span {
            color: #ffffff !important;
        }
        </style>
    """, unsafe_allow_html=True)

    if is_global:
        selected_main_tab = "Plan de Cuentas"
    else:
        if "cargas_active_tab" not in st.session_state:
            st.session_state["cargas_active_tab"] = "Trial Balance"
            
        tab_options = ["⚖️ Trial Balance", "📊 P&L (Estado de Resultados)", "📋 Plan de Cuentas"]
        cur_tab = st.session_state.get("cargas_active_tab", "Trial Balance")
        if "Plan" in cur_tab:
            default_index = 2
        elif "P&L" in cur_tab:
            default_index = 1
        else:
            default_index = 0
        
        selected_option = st.radio(
            "Seleccionar tipo de reporte a cargar",
            tab_options,
            index=default_index,
            horizontal=True,
            key="cargas_main_tab_selector",
            label_visibility="collapsed"
        )
        if "Plan" in selected_option:
            selected_main_tab = "Plan de Cuentas"
        elif "P&L" in selected_option:
            selected_main_tab = "P&L"
        else:
            selected_main_tab = "Trial Balance"
        st.session_state["cargas_active_tab"] = selected_main_tab

    # =========================================================================
    # TAB 0: PLAN DE CUENTAS (GLOBAL)
    # =========================================================================
    if selected_main_tab == "Plan de Cuentas":
        st.subheader("Maestro de Plan de Cuentas")
        st.write("Sube el Plan de Cuentas Maestro de tu empresa. Esto servirá para auditar que el Trial Balance no traiga cuentas huérfanas o no reconocidas.")
        
        with st.expander("👀 Ver formato requerido (Plantilla)"):
            st.markdown("""
            El archivo Excel del Plan Maestro de Cuentas debe contener **obligatoriamente** dos columnas fundamentales: 
            - `Cuenta`: El código numérico o alfanumérico.
            - `Tipo`: Debe indicar si la cuenta pertenece a "Balance" o "Resultado".
            
            Las demás columnas sirven de apoyo referencial para lectura.
            """)
            
            example_plan = {
                "Cuenta": ["110101", "110201", "210101", "310101", "410101", "510101"],
                "Descripción (opcional)": ["Caja General", "Banco Nacional", "Proveedores", "Capital Social", "Ingresos Operacionales", "Costo de Ventas"],
                "Tipo": ["Balance", "Balance", "Balance", "Balance", "Resultado", "Resultado"]
            }
            st.dataframe(pd.DataFrame(example_plan))
            
            df_plan = pd.DataFrame(example_plan)
            excel_bytes_plan = df_to_excel_bytes(df_plan, 'Ejemplo Plan Cuentas')
            st.download_button(
                label="📥 Descargar Excel de Ejemplo",
                data=excel_bytes_plan,
                file_name="ejemplo_plan_cuentas.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_plan"
            )

        uploaded_plan = st.file_uploader("Cargar Plan de Cuentas Maestro (Excel)", type=["xlsx", "xls"], key="up_plan_file")
        if uploaded_plan is not None:
            plan_file_sig = f"{uploaded_plan.name}_{uploaded_plan.size}"
            if st.session_state.get('last_plan_upload_sig') != plan_file_sig:
                df_new = pd.read_excel(uploaded_plan, dtype=str)
                df_new.columns = [str(c).strip() for c in df_new.columns]
                
                c_cuenta = next((c for c in df_new.columns if "cuenta" in c.lower()), None)
                c_tipo = next((c for c in df_new.columns if "tipo" in c.lower()), None)
                
                if c_cuenta and c_tipo:
                    if c_cuenta != "Cuenta":
                        df_new.rename(columns={c_cuenta: "Cuenta"}, inplace=True)
                    if c_tipo != "Tipo":
                        df_new.rename(columns={c_tipo: "Tipo"}, inplace=True)
                        
                    df_new['Cuenta'] = df_new['Cuenta'].astype(str).str.strip()
                    df_new['Tipo'] = df_new['Tipo'].astype(str).str.strip().str.capitalize()
                    
                    # Descartar filas vacías o nan
                    df_new = df_new[df_new['Cuenta'].notna() & (df_new['Cuenta'] != '') & (df_new['Cuenta'].str.lower() != 'nan')].copy()
                    df_new = df_new.drop_duplicates(subset=['Cuenta'], keep='last')
                    
                    df_final = sort_accounts(df_new, 'Cuenta', 'Tipo')
                    st.session_state['plan_cuentas_df'] = df_final
                    st.session_state['last_plan_upload_sig'] = plan_file_sig
                    
                    plan_path = os.path.join(empresa_path, "plan_cuentas.xlsx")
                    df_final.to_excel(plan_path, index=False)
                    if is_global:
                        propagate_global_file("plan_cuentas.xlsx", os.path.dirname(empresa_path))
                    
                    from src.core.sabana_manager import SabanaManager
                    SabanaManager.clear_sabana_cache()
                    st.session_state['ed_plan_ver'] = st.session_state.get('ed_plan_ver', 1) + 1
                    st.session_state['success_msg'] = f"✅ Plan de Cuentas reemplazado y actualizado exitosamente. Total maestro: **{len(df_final)}** cuentas cargadas."
                    st.rerun()
                else:
                    st.error("❌ El archivo Excel debe contener las columnas obligatorias 'Cuenta' y 'Tipo'.")

        if 'plan_cuentas_df' in st.session_state and st.session_state['plan_cuentas_df'] is not None:
            if 'ed_plan_ver' not in st.session_state:
                st.session_state['ed_plan_ver'] = 1

            st.success(f"🟢 **Plan de Cuentas Activo y Validado:** Se encuentra cargado el Plan de Cuentas Maestro para **{empresa_seleccionada}** con un total de **{len(st.session_state['plan_cuentas_df'])}** cuentas.")

            with st.expander("✏️ Editor de Datos Profesional (Modificar/Eliminar sin resubir Excel)", expanded=True):
                st.info("💡 **Indicación importante**: Haz doble clic en cualquier celda para editar. Para eliminar una fila, selecciónala en el lateral izquierdo y presiona la tecla `Supr` o `Delete`. Para agregar una nueva cuenta, escribe directamente en la última fila vacía (*). Tras realizar tus modificaciones, haz clic en **Guardar Cambios Directos en la Base**.")
                
                edited_plan = st.data_editor(
                    st.session_state['plan_cuentas_df'],
                    num_rows="dynamic",
                    use_container_width=True,
                    key=f"ed_plan_{st.session_state['ed_plan_ver']}"
                )

                if st.button("💾 Guardar Cambios Directos en la Base", type="primary", key="save_ed_plan"):
                    c_col = next((c for c in edited_plan.columns if "cuenta" in str(c).lower()), "Cuenta")
                    t_col = next((c for c in edited_plan.columns if "tipo" in str(c).lower()), "Tipo")
                    
                    df_valid = edited_plan.dropna(subset=[c_col]).copy()
                    df_valid[c_col] = df_valid[c_col].astype(str).str.strip()
                    df_valid = df_valid[df_valid[c_col] != '']
                    df_valid = df_valid[df_valid[c_col].str.lower() != 'nan']
                    
                    if df_valid.duplicated(subset=[c_col], keep=False).any():
                        dups = df_valid[df_valid.duplicated(subset=[c_col], keep=False)][c_col].unique()
                        st.error(f"🚨 Siguen existiendo cuentas duplicadas: {list(dups)}. Elimínalas antes de guardar.")
                    else:
                        if t_col in df_valid.columns:
                            df_valid[t_col] = df_valid[t_col].astype(str).str.strip().str.capitalize()
                        df_final = sort_accounts(df_valid, c_col, t_col if t_col in df_valid.columns else None)
                        st.session_state['plan_cuentas_df'] = df_final
                        
                        plan_path = os.path.join(empresa_path, "plan_cuentas.xlsx")
                        df_final.to_excel(plan_path, index=False)
                        if is_global:
                            propagate_global_file("plan_cuentas.xlsx", os.path.dirname(empresa_path))
                            
                        from src.core.sabana_manager import SabanaManager
                        SabanaManager.clear_sabana_cache()
                        st.session_state['ed_plan_ver'] += 1
                        st.session_state['success_msg'] = f"✅ Plan de Cuentas guardado con éxito. Total: {len(df_final)} cuentas."
                        st.rerun()

            excel_data = df_to_excel_bytes(st.session_state['plan_cuentas_df'], "Plan de Cuentas")
            st.download_button(
                label="📥 Descargar Plan de Cuentas en Excel",
                data=excel_data,
                file_name="plan_de_cuentas_activo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_dl_plan_cuentas"
            )
        else:
            st.info("Aún no se ha cargado ningún Plan de Cuentas.")

    # =========================================================================
    # TAB 1: TRIAL BALANCE
    # =========================================================================
    elif selected_main_tab == "Trial Balance":
        st.subheader("Importación de Trial Balance")
        st.write("Sube el balance de comprobación exportado de tu ERP corporativo.")

        col_sub_tb1, col_sub_tb2 = st.columns([1, 1])
        sub_tb_choice = st.radio(
            "Vista TB",
            ["📄 Cargar / Extraer TB", "👁️ Ver Trial Balance Guardado"],
            horizontal=True,
            key=f"sub_tb_choice_{empresa_seleccionada}",
            label_visibility="collapsed"
        )

        from src.models.trial_balance_db import TrialBalanceDB
        TrialBalanceDB.initialize()

        col_y, col_m = st.columns(2)
        with col_y:
            upload_year = st.selectbox("Año a cargar", ["2023", "2024", "2025", "2026", "2027"], index=2, key="upl_year")
        with col_m:
            upload_month = st.selectbox("Mes a cargar", ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"], key="upl_month")

        periodo_str = f"{upload_year}-{upload_month}"

        if sub_tb_choice == "📄 Cargar / Extraer TB":
            with st.expander("👀 Ver formato requerido (Plantilla)"):
                st.markdown("""
                Para que el sistema procese el balance, el archivo Excel debe contener **ESTRICTAMENTE estas 3 columnas explícitas** (además, la columna `Saldo DR/CR` debe sumar cero para garantizar la cuadratura contable de tu ERP):
                """)

                example_data = {
                    "N° de Cuenta": ["110101", "110201", "210101"],
                    "Nombre de la cuenta": ["Caja", "Banco de Chile", "Proveedores"],
                    "Saldo DR/CR": [1000, 4000, -5000]
                }
                st.dataframe(pd.DataFrame(example_data))

                df_tb = pd.DataFrame(example_data)
                excel_bytes_tb = df_to_excel_bytes(df_tb, 'Ejemplo Trial Balance')
                st.download_button(
                    label="📥 Descargar Excel de Ejemplo",
                    data=excel_bytes_tb,
                    file_name="ejemplo_trial_balance.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_tb"
                )

            st.info(f"Los datos se asimilarán al periodo contable: **{periodo_str}**")

            from src.core.cierre_engine import es_periodo_cerrado
            cerrado = es_periodo_cerrado(empresa_seleccionada, periodo_str)

            if cerrado:
                st.error(f"⚠️ **Periodo Bloqueado:** El periodo {periodo_str} se encuentra actualmente **CERRADO** en el histórico. Para volver a cargar o extraer datos, debes reabrir el periodo desde la pestaña de **Históricos**.")
            else:
                uploaded_file = st.file_uploader("Selecciona un archivo Excel", type=["xlsx", "xls"], key=f"up_tb_{empresa_seleccionada}_{periodo_str}")

                st.markdown("---")
                if st.button("🔌 Extracción de data desde ERP", type="secondary"):
                    from src.integrations.erp_adapter import ErpAdapterLogger
                    if ErpAdapterLogger.is_configured(empresa_path):
                        with st.spinner("Conectando con la API del ERP..."):
                            import time
                            start_time = time.time()
                            time.sleep(1.5)
                            df_erp, erp_name = ErpAdapterLogger.fetch_trial_balance(empresa_path, upload_year, upload_month)

                            cuentas_huerfanas = set()
                            cuentas_no_mapeadas = set()
                            if 'plan_cuentas_df' in st.session_state:
                                plan_cuentas = set(st.session_state['plan_cuentas_df']['Cuenta'].astype(str).str.strip())
                                tb_cuentas = set(df_erp['cuenta_id'].astype(str).str.strip())
                                cuentas_huerfanas = tb_cuentas - plan_cuentas

                            if cuentas_huerfanas:
                                st.error(f"❌ ERROR DE AUDITORÍA: Se detectaron {len(cuentas_huerfanas)} cuentas en la extracción que NO existen en el Plan de Cuentas maestro.")
                                st.write("Cuentas huérfanas encontradas:", sorted(list(cuentas_huerfanas)))
                            else:
                                cuentas_no_mapeadas = check_unmapped_accounts(empresa_path, tb_cuentas, st.session_state.get('plan_cuentas_df'))
                                if cuentas_no_mapeadas:
                                    st.error(f"❌ ERROR DE AUDITORÍA: Se detectaron {len(cuentas_no_mapeadas)} cuentas que NO están mapeadas en el sistema.")
                                    st.write("Cuentas sin clasificar encontradas:", sorted(list(cuentas_no_mapeadas)))
                                else:
                                    TrialBalanceDB.save_trial_balance(empresa_seleccionada, periodo_str, df_erp)
                                    st.session_state['tb_df'] = TrialBalanceDB.get_trial_balance(empresa_seleccionada, periodo_str)
                                    elapsed_time = time.time() - start_time
                                    st.session_state['success_msg'] = f"✅ Data extraída exitosamente desde {erp_name} para el periodo {periodo_str} (Tiempo de ejecución: {elapsed_time:.2f} segundos)."
                                    st.rerun()
                    else:
                        st.error("⚠️ No has configurado las credenciales del ERP para esta empresa. Dirígete a ⚙️ Configuraciones.")

                if uploaded_file is not None:
                    file_sig = f"{uploaded_file.name}_{uploaded_file.size}_{periodo_str}_{empresa_seleccionada}"
                    temp_path = os.path.join(empresa_path, "temp_uploaded.xlsx")
                    with open(temp_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())

                    from src.ingestion.trial_balance import TrialBalanceIngestor
                    ingestor = TrialBalanceIngestor(temp_path)
                    try:
                        with st.spinner("Procesando y validando balance de comprobación..."):
                            import time
                            start_time = time.time()
                            df = ingestor.load_and_standardize()

                            cuentas_huerfanas = set()
                            cuentas_no_mapeadas = set()
                            if 'plan_cuentas_df' in st.session_state:
                                plan_cuentas = set(st.session_state['plan_cuentas_df']['Cuenta'].astype(str).str.strip())
                                tb_cuentas = set(df['cuenta_id'].astype(str).str.strip())
                                cuentas_huerfanas = tb_cuentas - plan_cuentas

                            if cuentas_huerfanas:
                                st.error(f"❌ ERROR DE AUDITORÍA: Se detectaron {len(cuentas_huerfanas)} cuentas en el Balance que NO existen en el Plan Maestro.")
                                st.write("Cuentas huérfanas encontradas:", sorted(list(cuentas_huerfanas)))
                                
                                # Auto-resolución en 1 clic
                                with st.container(border=True):
                                    st.markdown(f"#### ⚡ Auto-Resolución de Plan Maestro")
                                    st.write("Puedes registrar estas cuentas huérfanas en el Plan Maestro de inmediato sin salir de esta pantalla:")
                                    if st.button("➕ Registrar cuentas faltantes en el Plan de Cuentas Maestro", type="primary", key="btn_auto_add_tb_plan"):
                                        new_rows = []
                                        for acc in sorted(list(cuentas_huerfanas)):
                                            match = df[df['cuenta_id'].astype(str) == str(acc)]
                                            desc_val = str(match.iloc[0]['descripcion']) if not match.empty and 'descripcion' in match.columns else ""
                                            tipo_sug = "Resultado" if str(acc).startswith(('4', '5', '6', '7', '8')) else "Balance"
                                            new_rows.append({"Cuenta": str(acc), "Nombre": desc_val, "Tipo": tipo_sug})
                                        df_add = pd.DataFrame(new_rows)
                                        df_plan_cur = st.session_state.get('plan_cuentas_df')
                                        df_final = pd.concat([df_plan_cur, df_add], ignore_index=True).drop_duplicates(subset=['Cuenta']) if df_plan_cur is not None else df_add
                                        df_final = sort_accounts(df_final, 'Cuenta', 'Tipo')
                                        st.session_state['plan_cuentas_df'] = df_final
                                        df_final.to_excel(os.path.join(empresa_path, "plan_cuentas.xlsx"), index=False)
                                        if is_global:
                                            propagate_global_file("plan_cuentas.xlsx", os.path.dirname(empresa_path))
                                        st.session_state['success_msg'] = f"✅ Se registraron {len(cuentas_huerfanas)} cuentas en el Plan Maestro exitosamente."
                                        st.rerun()
                            else:
                                cuentas_no_mapeadas = check_unmapped_accounts(empresa_path, tb_cuentas, st.session_state.get('plan_cuentas_df'))
                                if cuentas_no_mapeadas:
                                    st.error(f"❌ ERROR DE AUDITORÍA: Se detectaron {len(cuentas_no_mapeadas)} cuentas que NO están mapeadas en el sistema.")
                                    st.write("Cuentas sin clasificar encontradas:", sorted(list(cuentas_no_mapeadas)))
                                    st.warning("⚠️ Debes clasificar estas cuentas en Organización de Cuentas o en el Mapeador para completar la auditoría.")
                                else:
                                    if st.session_state.get('last_tb_file_sig') != file_sig:
                                        TrialBalanceDB.save_trial_balance(empresa_seleccionada, periodo_str, df)
                                        elapsed_time = time.time() - start_time
                                        st.session_state['last_tb_file_sig'] = file_sig
                                        st.session_state['success_msg'] = f"✅ Archivo cargado e ingestado exitosamente en base de datos. Se guardaron {len(df)} registros para {periodo_str} (Tiempo: {elapsed_time:.2f}s). Auditoría superada: 100% de cuentas válidas y mapeadas."
                                        st.session_state['tb_df'] = TrialBalanceDB.get_trial_balance(empresa_seleccionada, periodo_str)
                                        for k in ['preview_df', 'balance_excel_binary', 'balance_word_binary', 'er_preview_df', 'er_excel_binary', 'flujo_preview_df', 'pat_preview_df', 'ori_preview_df']:
                                            if k in st.session_state:
                                                del st.session_state[k]
                                        st.rerun()
                    except Exception as e:
                        error_msg = str(e)
                        st.error(f"Error procesando el archivo: {error_msg}")

        else: # Ver TB
            tb_from_db = TrialBalanceDB.get_trial_balance(empresa_seleccionada, periodo_str)
            if tb_from_db is not None and not tb_from_db.empty:
                st.subheader(f"📊 Vista de Trial Balance Guardado ({periodo_str})")
                tb_df_to_show = tb_from_db.copy()
                cuenta_col_tb = next((c for c in tb_df_to_show.columns if "cuenta" in str(c).lower() and "nombre" not in str(c).lower()), "cuenta_id")
                desc_col_tb = next((c for c in tb_df_to_show.columns if "nombre" in str(c).lower() or "desc" in str(c).lower()), "descripcion")
                numeric_cols_tb = [c for c in tb_df_to_show.columns if c not in [cuenta_col_tb, desc_col_tb]]

                for col in numeric_cols_tb:
                    tb_df_to_show[col] = tb_df_to_show[col].apply(lambda x: f"{int(round(pd.to_numeric(x, errors='coerce') or 0.0)):,}".replace(",", "."))

                column_config_tb = {c: st.column_config.TextColumn(c) for c in numeric_cols_tb}
                column_config_tb[cuenta_col_tb] = st.column_config.TextColumn("N° de Cuenta")
                column_config_tb[desc_col_tb] = st.column_config.TextColumn("Nombre de la cuenta")

                st.dataframe(
                    tb_df_to_show,
                    use_container_width=True,
                    column_config=column_config_tb,
                    key=f"df_tb_show_{empresa_seleccionada}_{periodo_str}"
                )
                excel_data = df_to_excel_bytes(tb_from_db, "Trial Balance")
                st.download_button(
                    label="📥 Descargar Trial Balance en Excel",
                    data=excel_data,
                    file_name=f"trial_balance_{empresa_seleccionada}_{periodo_str}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"btn_dl_tb_{empresa_seleccionada}_{periodo_str}"
                )
                render_balance_vs_pl_validation_card(empresa_seleccionada, periodo_str)
            else:
                st.info(f"Aún no se ha cargado ningún Trial Balance para **{empresa_seleccionada}** en el periodo **{periodo_str}**.")

    # =========================================================================
    # TAB 2: CUBO P&L (ESTADO DE RESULTADOS)
    # =========================================================================
    elif selected_main_tab == "P&L":
        st.subheader("Cubo de Estado de Resultados (P&L)")
        st.write("Sube el cubo analítico de Pérdidas y Ganancias (ventas por centro de costo, unidad de negocio, etc.) o edita directamente los saldos guardados.")

        # Obtener columnas de P&L de taxonomía (con caché optimizado)
        from src.core.taxonomy_cache import get_pl_taxonomy_columns
        db_cols = get_pl_taxonomy_columns(empresa_seleccionada)

        default_pl_cols = [
            "Ingresos de arriendo fibra optica",
            "Ingresos de actividades ordinarias", 
            "Costo de ventas", 
            "Acceso a infraestructura fibra óptica",
            "Depreciación operacional", 
            "Gastos de administración", 
            "Depreciación y amortizaciones", 
            "Otros ingresos por función", 
            "Otros egresos por función", 
            "Ingresos financieros", 
            "Costos financieros", 
            "Diferencias de cambio", 
            "Resultados por unidades de reajuste", 
            "Resultado por impuestos a las ganancias"
        ]

        pl_rubros = db_cols if db_cols else default_pl_cols
        plantilla_cols_transaccional = ["N° de cuenta", "Nombre de la cuenta"] + pl_rubros

        sub_pl_choice = st.radio(
            "Vista P&L",
            ["📄 Cargar / Editar P&L", "👁️ Ver P&L Guardado"],
            horizontal=True,
            key=f"sub_pl_choice_{empresa_seleccionada}",
            label_visibility="collapsed"
        )

        col_y_pl, col_m_pl = st.columns(2)
        with col_y_pl:
            upload_year_pl = st.selectbox("Año a cargar", ["2023", "2024", "2025", "2026", "2027"], index=2, key="upl_year_pl")
        with col_m_pl:
            upload_month_pl = st.selectbox("Mes a cargar", ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"], key="upl_month_pl")

        periodo_str_pl = f"{upload_year_pl}-{upload_month_pl}"

        if sub_pl_choice == "📄 Cargar / Editar P&L":
            with st.expander("👀 Ver formato requerido (Plantilla)"):
                st.markdown("""
                Para procesar el cubo P&L de forma analítica, el Excel debe contener las siguientes columnas para estructurar los saldos:
                """)
                st.write(f"`{', '.join(plantilla_cols_transaccional)}`")

                mock_plcubo_df = pd.DataFrame(columns=plantilla_cols_transaccional)
                mock_plcubo_df.loc[0] = {c: None for c in plantilla_cols_transaccional}
                mock_plcubo_df.loc[0, "N° de cuenta"] = "410101"
                mock_plcubo_df.loc[0, "Nombre de la cuenta"] = "Ventas Consumidor Final"

                if "Ingresos de actividades ordinarias" in mock_plcubo_df.columns:
                    mock_plcubo_df.loc[0, "Ingresos de actividades ordinarias"] = "SERVICIOS MAYORISTAS"
                elif len(pl_rubros) > 0:
                    mock_plcubo_df.loc[0, pl_rubros[0]] = "EJEMPLO INGRESO"

                mock_plcubo_df.loc[1] = {c: None for c in plantilla_cols_transaccional}
                mock_plcubo_df.loc[1, "N° de cuenta"] = "510101"
                mock_plcubo_df.loc[1, "Nombre de la cuenta"] = "Costo Tráfico Local"

                if "Costo de ventas" in mock_plcubo_df.columns:
                    mock_plcubo_df.loc[1, "Costo de ventas"] = "SEÑALES NACIONALES"
                elif len(pl_rubros) > 1:
                    mock_plcubo_df.loc[1, pl_rubros[1]] = "EJEMPLO COSTO"

                st.dataframe(mock_plcubo_df)
                excel_data_cubo = df_to_excel_bytes(mock_plcubo_df, 'Plantilla Carga P&L')

                st.download_button(
                    label="📥 Descargar Plantilla P&L",
                    data=excel_data_cubo,
                    file_name="plantilla_carga_pl_cubo.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_plcubo"
                )

            st.info(f"Los datos de este Cubo P&L se archivarán en el periodo contable: **{periodo_str_pl}**")

            from src.core.cierre_engine import es_periodo_cerrado
            cerrado_pl = es_periodo_cerrado(empresa_seleccionada, periodo_str_pl)

            if cerrado_pl:
                st.error(f"⚠️ **Periodo Bloqueado:** El periodo {periodo_str_pl} se encuentra actualmente **CERRADO** en el histórico. Para volver a cargar datos de P&L, debes reabrir el periodo desde la pestaña de **Históricos**.")
            else:
                up_pl_cubo = st.file_uploader("Selecciona el Cubo P&L (Excel)", type=["xlsx", "xls"], key=f"up_pl_cubo_{empresa_seleccionada}_{periodo_str_pl}")

                # Cargar maestro de mapeos para la empresa
                map_pl_path = os.path.join(empresa_path, "map_pl.xlsx")
                if os.path.exists(map_pl_path):
                    df_map_master = pd.read_excel(map_pl_path)
                else:
                    df_map_master = None

                df_processed = None
                current_file_id = None
                df_raw_for_reprocess = None

                if up_pl_cubo is not None:
                    try:
                        file_sig = f"{up_pl_cubo.name}_{up_pl_cubo.size}_{empresa_seleccionada}_{periodo_str_pl}"
                        if st.session_state.get('pl_cached_file_sig') != file_sig or 'pl_cached_sheet_names' not in st.session_state:
                            with st.spinner("Analizando la estructura del archivo Excel..."):
                                xls = pd.ExcelFile(up_pl_cubo, engine='openpyxl')
                                st.session_state['pl_cached_sheet_names'] = xls.sheet_names
                                st.session_state['pl_cached_file_sig'] = file_sig

                        sheet_names = st.session_state['pl_cached_sheet_names']

                        if len(sheet_names) > 1:
                            def guess_sheet_index(sheets, company):
                                c_lower = company.lower()
                                for i, s in enumerate(sheets):
                                    s_lower = s.lower()
                                    if s_lower in c_lower or c_lower in s_lower:
                                        return i
                                if "pacifico" in c_lower:
                                    for i, s in enumerate(sheets):
                                        if "pacifico" in s.lower():
                                            return i
                                if "holdco" in c_lower or "terra" in c_lower:
                                    for i, s in enumerate(sheets):
                                        if "holdco" in s.lower() or "terra" in s.lower():
                                            return i
                                return 0

                            default_idx = guess_sheet_index(sheet_names, empresa_seleccionada)
                            selected_sheet = st.selectbox("Selecciona la pestaña del Excel a importar:", sheet_names, index=default_idx)
                        else:
                            selected_sheet = sheet_names[0]

                        full_sig = f"{file_sig}_{selected_sheet}"
                        current_file_id = full_sig

                        # Reutilizar resultado procesado en caché de sesión si no ha cambiado el archivo
                        if (
                            st.session_state.get('pl_last_processed_sig') == full_sig
                            and 'pl_cached_df_processed' in st.session_state
                        ):
                            df_processed = st.session_state['pl_cached_df_processed']
                            df_raw_for_reprocess = st.session_state.get('pl_cached_df_raw')
                        else:
                            import time
                            start_time = time.time()
                            with st.spinner(f"Cargando transacciones de la pestaña '{selected_sheet}'..."):
                                df_raw = pd.read_excel(up_pl_cubo, sheet_name=selected_sheet, dtype=str, engine='openpyxl')
                                df_raw.columns = [str(c).strip() for c in df_raw.columns]
                                df_raw_for_reprocess = df_raw.copy()

                            # Detectar formato: ¿Es Cubo de Odoo o Formato P&L ancho clásico?
                            has_eerr_col = any("eerr" in c.lower() or "informe_eerr" in c.lower() for c in df_raw.columns)

                            if has_eerr_col:
                                format_detected = "Cubo transaccional original de Odoo"
                                st.info(f"📂 Formato detectado: **{format_detected}** en la pestaña **'{selected_sheet}'**.")
                                if df_map_master is None:
                                    st.warning("⚠️ No se encontró el archivo maestro 'map_pl.xlsx' de la empresa activa. Se usará clasificación directa de Odoo sin mapeo específico.")

                                with st.spinner("Agrupando transacciones y aplicando reglas de mapeo IFRS (YTD)..."):
                                    from src.core.pl_cubo_processor import process_odoo_cubo
                                    df_processed, df_audit = process_odoo_cubo(df_raw, upload_year_pl, upload_month_pl, df_map_master, standard_categories=pl_rubros, return_audit_log=True)
                                    st.session_state['pl_cubo_audit_warnings'] = df_audit
                            else:
                                format_detected = "Formato de columnas P&L estructurado (Matriz)"
                                st.info(f"📂 Formato detectado: **{format_detected}**.")

                                cuenta_col = next((c for c in df_raw.columns if "Cuenta" in c or "cuenta" in c), None)
                                if cuenta_col:
                                    df_processed = df_raw.copy()
                                    if cuenta_col != "N° de cuenta":
                                        df_processed.rename(columns={cuenta_col: "N° de cuenta"}, inplace=True)
                                    for cat in pl_rubros:
                                        if cat not in df_processed.columns:
                                            df_processed[cat] = "0.0"
                                    for cat in pl_rubros:
                                        df_processed[cat] = pd.to_numeric(df_processed[cat], errors='coerce').fillna(0.0)
                                else:
                                    st.error("❌ Estructura inválida. No se detectó ninguna columna de 'Cuenta'.")

                            elapsed_time = time.time() - start_time
                            if df_processed is not None:
                                st.session_state['pl_cached_df_processed'] = df_processed
                                st.session_state['pl_cached_df_raw'] = df_raw_for_reprocess
                                st.session_state['pl_last_processed_sig'] = full_sig
                                st.success(f"✅ Archivo Excel procesado con éxito (Tiempo de ejecución: {elapsed_time:.2f} segundos).")
                    except Exception as e:
                        st.error(f"❌ Error al procesar el archivo Excel: {e}")

                # Verificar data guardada en base de datos
                db_pl_df = PlCuboDB.get_pl_cubo(empresa_seleccionada, periodo_str_pl)

                # Acciones si hay datos guardados y no se ha subido un archivo
                if up_pl_cubo is None and db_pl_df is not None and not db_pl_df.empty:
                    st.write("---")
                    st.info(f"💡 Se detectaron datos guardados en la base de datos para el periodo **{periodo_str_pl}** ({len(db_pl_df)} cuentas registradas).")
                    col_btn1, col_btn2 = st.columns([2, 1])
                    with col_btn1:
                        if st.button("✏️ Cargar y Editar datos guardados de este periodo", use_container_width=True, type="primary"):
                            st.session_state['pl_edit_df'] = add_totals_row(db_pl_df.copy())
                            st.session_state['pl_edit_file_id'] = f"saved_db_{periodo_str_pl}"
                            st.rerun()

                # Cargar en el editor cuando entra un archivo nuevo
                if current_file_id:
                    if st.session_state.get('pl_edit_file_id') != current_file_id:
                        st.session_state['pl_edit_df'] = add_totals_row(df_processed)
                        st.session_state['pl_edit_file_id'] = current_file_id

                is_editing_db = st.session_state.get('pl_edit_file_id') == f"saved_db_{periodo_str_pl}"
                has_active_editor = (up_pl_cubo is not None) or is_editing_db

                if has_active_editor and 'pl_edit_df' in st.session_state and st.session_state['pl_edit_df'] is not None:
                    col_hd1, col_hd2 = st.columns([3, 1])
                    with col_hd1:
                        if is_editing_db:
                            st.warning("⚠️ **Modo Edición Directa de BD**: Estás editando los saldos guardados para este periodo.")
                        else:
                            st.info("📂 **Modo Archivo Subido**: Puedes revisar o editar montos antes de confirmar el guardado.")
                    with col_hd2:
                        if st.button("❌ Cerrar / Limpiar Editor", use_container_width=True):
                            if 'pl_edit_df' in st.session_state:
                                del st.session_state['pl_edit_df']
                            if 'pl_edit_file_id' in st.session_state:
                                del st.session_state['pl_edit_file_id']
                            if 'pl_cubo_audit_warnings' in st.session_state:
                                del st.session_state['pl_cubo_audit_warnings']
                            st.rerun()

                    st.markdown("### 📝 Vista Previa y Edición del P&L")
                    st.caption("Puedes modificar los montos directamente haciendo doble clic en cualquier celda o agregar nuevas filas.")
                    render_balance_vs_pl_validation_card(empresa_seleccionada, periodo_str_pl)

                    # Mostrar advertencias de auditoría de clasificación si existen
                    audit_warn_df = st.session_state.get('pl_cubo_audit_warnings')
                    if audit_warn_df is not None and not audit_warn_df.empty:
                        st.warning(
                            f"⚠️ **Observaciones de Auditoría en Cubo Odoo ({len(audit_warn_df)} transacciones detectadas):** "
                            f"Se encontraron filas con la columna de clasificación vacía o no reconocida. "
                            f"El sistema las asignó automáticamente según las reglas de contingencia / diccionario."
                        )
                        with st.expander(f"🔍 Ver detalle de las {len(audit_warn_df)} transacciones asignadas automáticamente", expanded=False):
                            st.dataframe(
                                audit_warn_df,
                                use_container_width=True,
                                hide_index=True
                            )
                            import io
                            audit_buf = io.BytesIO()
                            with pd.ExcelWriter(audit_buf, engine='openpyxl') as writer:
                                audit_warn_df.to_excel(writer, index=False, sheet_name='Advertencias_Clasificacion')
                            st.download_button(
                                label="📥 Descargar Reporte de Advertencias (Excel)",
                                data=audit_buf.getvalue(),
                                file_name=f"advertencias_clasificacion_cubo_{periodo_str_pl}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"btn_download_audit_{periodo_str_pl}"
                            )

                    cuenta_col = next((c for c in st.session_state['pl_edit_df'].columns if "cuenta" in str(c).lower() and "nombre" not in str(c).lower()), "N° de cuenta")
                    desc_col = next((c for c in st.session_state['pl_edit_df'].columns if "nombre" in str(c).lower()), "Nombre de la cuenta")
                    numeric_cols = [c for c in st.session_state['pl_edit_df'].columns if c not in [cuenta_col, desc_col]]

                    # Preparar DF formateado de manera optimizada
                    df_to_edit = st.session_state['pl_edit_df'].copy()
                    for col in numeric_cols:
                        series = df_to_edit[col]
                        first_val = series.dropna().iloc[0] if not series.dropna().empty else None
                        if isinstance(first_val, str) and ('.' in first_val or first_val in ['0', '-']):
                            continue
                        df_to_edit[col] = series.apply(parse_numeric_cell).map(lambda x: f"{int(round(x)):,}".replace(",", "."))

                    # -------------------------------------------------------------
                    # AUDITORÍA Y VALIDACIONES
                    # -------------------------------------------------------------
                    pl_huerfanas = set()
                    pl_no_mapeadas = set()
                    if 'plan_cuentas_df' in st.session_state and st.session_state['plan_cuentas_df'] is not None:
                        plan_cuentas = set(st.session_state['plan_cuentas_df']['Cuenta'].astype(str).str.strip())
                        editor_accounts = set(
                            st.session_state['pl_edit_df'][
                                st.session_state['pl_edit_df'][cuenta_col].astype(str).str.strip().str.upper() != "TOTAL"
                            ][cuenta_col].astype(str).str.strip()
                        )
                        # Filtrar cuentas vacías o nan
                        editor_accounts = {acc for acc in editor_accounts if acc and acc.lower() != 'nan'}
                        pl_huerfanas = editor_accounts - plan_cuentas
                        if not pl_huerfanas:
                            pl_no_mapeadas = check_unmapped_accounts(empresa_path, editor_accounts, st.session_state['plan_cuentas_df'])

                    # -------------------------------------------------------------
                    # CASO 1: CUENTAS HUÉRFANAS (CON AUTO-RESOLUCIÓN EN 1 CLIC)
                    # -------------------------------------------------------------
                    if pl_huerfanas:
                        st.error(f"❌ ERROR DE AUDITORÍA: Se detectaron {len(pl_huerfanas)} cuentas en el P&L que NO existen en el Plan de Cuentas maestro.")
                        st.write("Cuentas huérfanas encontradas:", sorted(list(pl_huerfanas)))
                        
                        with st.container(border=True):
                            st.markdown("#### ⚡ Auto-Resolución: Agregar Cuentas al Plan Maestro en 1 Clic")
                            st.write("Agrega estas cuentas automáticamente a tu Plan Maestro sin necesidad de salir ni editar un Excel:")
                            if st.button("➕ Registrar cuentas faltantes en el Plan de Cuentas Maestro", type="primary", key="btn_auto_add_pl_plan"):
                                new_rows = []
                                for acc in sorted(list(pl_huerfanas)):
                                    acc_name = ""
                                    if desc_col in st.session_state['pl_edit_df'].columns:
                                        row_m = st.session_state['pl_edit_df'][st.session_state['pl_edit_df'][cuenta_col].astype(str) == str(acc)]
                                        if not row_m.empty:
                                            acc_name = str(row_m.iloc[0][desc_col])
                                    new_rows.append({"Cuenta": str(acc), "Nombre": acc_name, "Tipo": "Resultado"})
                                
                                df_add = pd.DataFrame(new_rows)
                                df_plan_cur = st.session_state.get('plan_cuentas_df')
                                df_final = pd.concat([df_plan_cur, df_add], ignore_index=True).drop_duplicates(subset=['Cuenta']) if df_plan_cur is not None else df_add
                                df_final = sort_accounts(df_final, 'Cuenta', 'Tipo')
                                st.session_state['plan_cuentas_df'] = df_final
                                df_final.to_excel(os.path.join(empresa_path, "plan_cuentas.xlsx"), index=False)
                                if is_global:
                                    propagate_global_file("plan_cuentas.xlsx", os.path.dirname(empresa_path))
                                st.session_state['success_msg'] = f"✅ Se registraron {len(pl_huerfanas)} cuentas en el Plan Maestro exitosamente."
                                st.rerun()

                    # -------------------------------------------------------------
                    # CASO 2: CUENTAS NO MAPEADAS (CON MAPEADOR RÁPIDO EN PANTALLA)
                    # -------------------------------------------------------------
                    elif pl_no_mapeadas:
                        st.error(f"❌ ERROR DE AUDITORÍA: Se detectaron {len(pl_no_mapeadas)} cuentas en el P&L que NO están mapeadas en el maestro de P&L.")
                        st.write("Cuentas sin clasificar:", sorted(list(pl_no_mapeadas)))
                        
                        with st.container(border=True):
                            st.markdown("#### ⚡ Mapeador Rápido de P&L (En 1 Clic)")
                            st.write("Selecciona el rubro IFRS para cada cuenta directamente en pantalla para desbloquear el guardado de inmediato:")
                            
                            with st.form("form_quick_map_pl"):
                                quick_map_dict = {}
                                for acc in sorted(list(pl_no_mapeadas)):
                                    acc_name = ""
                                    if desc_col in st.session_state['pl_edit_df'].columns:
                                        row_m = st.session_state['pl_edit_df'][st.session_state['pl_edit_df'][cuenta_col].astype(str) == str(acc)]
                                        if not row_m.empty:
                                            acc_name = str(row_m.iloc[0][desc_col])
                                    
                                    col_q1, col_q2 = st.columns([1, 2])
                                    with col_q1:
                                        st.markdown(f"**Cuenta `{acc}`**\n\n*{acc_name}*")
                                    with col_q2:
                                        quick_map_dict[acc] = st.selectbox(
                                            f"Rubro para {acc}",
                                            options=pl_rubros,
                                            key=f"qm_sel_{acc}"
                                        )
                                
                                btn_save_quick_map = st.form_submit_button("💾 Guardar Mapeo y Aplicar a P&L", type="primary")
                                if btn_save_quick_map:
                                    if os.path.exists(map_pl_path):
                                        df_map_to_update = pd.read_excel(map_pl_path, dtype=str)
                                    else:
                                        df_map_to_update = pd.DataFrame(columns=["Cuenta"] + pl_rubros)
                                    
                                    c_acc_m = df_map_to_update.columns[0]
                                    df_map_to_update[c_acc_m] = df_map_to_update[c_acc_m].astype(str).str.strip()
                                    
                                    for acc_key, rubro_val in quick_map_dict.items():
                                        acc_k_str = str(acc_key).strip()
                                        if acc_k_str in df_map_to_update[c_acc_m].values:
                                            if rubro_val in df_map_to_update.columns:
                                                df_map_to_update.loc[df_map_to_update[c_acc_m] == acc_k_str, rubro_val] = "1"
                                        else:
                                            new_row_m = {c: "" for c in df_map_to_update.columns}
                                            new_row_m[c_acc_m] = acc_k_str
                                            if rubro_val in df_map_to_update.columns:
                                                new_row_m[rubro_val] = "1"
                                            else:
                                                df_map_to_update[rubro_val] = ""
                                                new_row_m[rubro_val] = "1"
                                            df_map_to_update = pd.concat([df_map_to_update, pd.DataFrame([new_row_m])], ignore_index=True)
                                    
                                    df_map_to_update.to_excel(map_pl_path, index=False)
                                    st.session_state['map_pl_df'] = df_map_to_update
                                    if is_global:
                                        propagate_global_file("map_pl.xlsx", os.path.dirname(empresa_path))
                                    
                                    # Si hay un archivo cargado, re-procesar automáticamente con el nuevo mapeo
                                    if df_raw_for_reprocess is not None:
                                        from src.core.pl_cubo_processor import process_odoo_cubo
                                        df_reprocessed, df_audit = process_odoo_cubo(df_raw_for_reprocess, upload_year_pl, upload_month_pl, df_map_to_update, standard_categories=pl_rubros, return_audit_log=True)
                                        st.session_state['pl_cubo_audit_warnings'] = df_audit
                                        st.session_state['pl_edit_df'] = add_totals_row(df_reprocessed)
                                        # Guardar de inmediato en BD
                                        PlCuboDB.save_pl_cubo(empresa_seleccionada, periodo_str_pl, df_reprocessed)
                                        st.session_state['pl_df'] = PlCuboDB.get_pl_cubo(empresa_seleccionada, periodo_str_pl)

                                    st.session_state['success_msg'] = "✅ Cuentas mapeadas y P&L actualizado con éxito."
                                    st.rerun()

                    # -------------------------------------------------------------
                    # CASO 3: AUDITORÍA LIMPIA -> EDITOR ACTIVO Y GUARDADO DIRECTO
                    # -------------------------------------------------------------
                    else:
                        # Auto-guardar archivo subido la primera vez si la auditoría es limpia
                        if current_file_id and df_processed is not None:
                            if st.session_state.get('last_pl_saved_file_id') != current_file_id:
                                try:
                                    saved_cnt = PlCuboDB.save_pl_cubo(empresa_seleccionada, periodo_str_pl, df_processed)
                                    st.session_state['last_pl_saved_file_id'] = current_file_id
                                    st.session_state['pl_df'] = PlCuboDB.get_pl_cubo(empresa_seleccionada, periodo_str_pl)
                                    for k in ['preview_df', 'balance_excel_binary', 'balance_word_binary', 'er_preview_df', 'er_excel_binary', 'flujo_preview_df', 'pat_preview_df', 'ori_preview_df']:
                                        if k in st.session_state:
                                            del st.session_state[k]
                                    st.success(f"✅ Cubo P&L archivado e ingestado exitosamente en base de datos para el periodo {periodo_str_pl} ({saved_cnt} registros dimensionales).")
                                except Exception as save_err:
                                    st.error(f"❌ Error al auto-guardar en base de datos: {save_err}")

                        editor_key = f"pl_editor_act_{periodo_str_pl}"
                        
                        edited_df = st.data_editor(
                            df_to_edit,
                            num_rows="dynamic",
                            use_container_width=True,
                            key=editor_key
                        )

                        col_act1, col_act2 = st.columns([2, 1])
                        with col_act1:
                            if st.button("💾 Guardar Modificaciones en Base de Datos", type="primary", use_container_width=True):
                                try:
                                    df_to_save = edited_df.copy()
                                    # Descartar fila TOTAL si existe
                                    df_to_save = df_to_save[df_to_save[cuenta_col].astype(str).str.strip().str.upper() != "TOTAL"].copy()
                                    for col in numeric_cols:
                                        df_to_save[col] = df_to_save[col].apply(parse_numeric_cell)

                                    saved_count = PlCuboDB.save_pl_cubo(empresa_seleccionada, periodo_str_pl, df_to_save)
                                    st.session_state['pl_df'] = PlCuboDB.get_pl_cubo(empresa_seleccionada, periodo_str_pl)
                                    st.session_state['pl_edit_df'] = add_totals_row(df_to_save)
                                    msg = f"✅ Cubo P&L actualizado exitosamente en base de datos para {periodo_str_pl} ({saved_count} registros guardados)."
                                    st.session_state['success_msg'] = msg
                                    st.toast(msg, icon="✅")
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"❌ Error al guardar en base de datos: {e}")

                        with col_act2:
                            if st.button("🔄 Recalcular Totales en Pantalla", use_container_width=True):
                                df_recalc = edited_df.copy()
                                df_recalc = df_recalc[df_recalc[cuenta_col].astype(str).str.strip().str.upper() != "TOTAL"].copy()
                                for col in numeric_cols:
                                    df_recalc[col] = df_recalc[col].apply(parse_numeric_cell)
                                st.session_state['pl_edit_df'] = add_totals_row(df_recalc)
                                st.rerun()

        else: # Ver P&L Guardado
            db_pl_show = PlCuboDB.get_pl_cubo(empresa_seleccionada, periodo_str_pl)
            if db_pl_show is not None and not db_pl_show.empty:
                st.subheader(f"📊 Vista de Cubo P&L Guardado en Base de Datos ({periodo_str_pl})")
                pl_show_df = db_pl_show.copy()
                c_cuenta = next((c for c in pl_show_df.columns if "cuenta" in str(c).lower() and "nombre" not in str(c).lower()), "N° de cuenta")
                c_desc = next((c for c in pl_show_df.columns if "nombre" in str(c).lower() or "desc" in str(c).lower()), "Nombre de la cuenta")
                n_cols = [c for c in pl_show_df.columns if c not in [c_cuenta, c_desc]]

                for col in n_cols:
                    pl_show_df[col] = pl_show_df[col].apply(lambda x: f"{int(round(parse_numeric_cell(x))):,}".replace(",", "."))

                column_config_pl = {c: st.column_config.TextColumn(c) for c in n_cols}
                column_config_pl[c_cuenta] = st.column_config.TextColumn("N° de Cuenta")
                column_config_pl[c_desc] = st.column_config.TextColumn("Nombre de la cuenta")

                st.dataframe(
                    pl_show_df,
                    use_container_width=True,
                    column_config=column_config_pl,
                    key=f"df_pl_show_{empresa_seleccionada}_{periodo_str_pl}"
                )
                excel_data = df_to_excel_bytes(db_pl_show, "Cubo P&L")
                st.download_button(
                    label="📥 Descargar Cubo P&L en Excel",
                    data=excel_data,
                    file_name=f"cubo_pl_{empresa_seleccionada}_{periodo_str_pl}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_dl_pl_db"
                )
                render_balance_vs_pl_validation_card(empresa_seleccionada, periodo_str_pl)
            else:
                st.info(f"Aún no se ha guardado ningún Cubo P&L para el periodo **{periodo_str_pl}**.")
