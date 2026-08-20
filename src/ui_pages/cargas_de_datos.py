import streamlit as st
import pandas as pd
import os
from src.core.excel_utils import df_to_excel_bytes, sort_accounts, propagate_global_file
from src.models.pl_cubo_db import PlCuboDB

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
    
    # Convert numeric columns to int64 in df_clean to ensure proper numeric type
    for col in numeric_cols:
        df_clean[col] = pd.to_numeric(df_clean[col], errors='coerce').fillna(0.0).round().astype('int64')
    
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
    
    # Force numeric type (int64) on the combined dataframe as well
    for col in numeric_cols:
        df_total[col] = pd.to_numeric(df_total[col], errors='coerce').fillna(0.0).round().astype('int64')
        
    return df_total

def check_unmapped_accounts(empresa_path, accounts_to_check, plan_cuentas_df):
    """
    Checks if the provided accounts are mapped in the Excel templates.
    Returns:
        cuentas_no_mapeadas (set): Set of accounts that lack a mapping.
    """
    if plan_cuentas_df is None or plan_cuentas_df.empty:
        return set()
        
    plan_cuentas = set(plan_cuentas_df['Cuenta'].astype(str).str.strip())
    plan_tipo_map = dict(zip(plan_cuentas_df['Cuenta'].astype(str).str.strip(), plan_cuentas_df['Tipo'].astype(str).str.strip()))
    
    # 1. Load Balance mappings
    map_bal_path = os.path.join(empresa_path, "map_balance.xlsx")
    mapped_bal_accounts = set()
    if os.path.exists(map_bal_path):
        try:
            df_bal = pd.read_excel(map_bal_path, dtype=str)
            if not df_bal.empty:
                col_bal_cuenta = next((c for c in df_bal.columns if 'cuenta' in c.lower()), df_bal.columns[0])
                clasif_col = next((c for c in df_bal.columns if 'clasificaci' in c.lower() and 'balance' in c.lower()), None)
                if clasif_col:
                    for _, row in df_bal.iterrows():
                        acc_id = str(row[col_bal_cuenta]).strip()
                        val = row[clasif_col]
                        if pd.notna(val) and str(val).strip() != "" and str(val).strip().lower() != "nan":
                            mapped_bal_accounts.add(acc_id)
        except Exception:
            pass

    # 2. Load P&L mappings
    map_pl_path = os.path.join(empresa_path, "map_pl.xlsx")
    mapped_pl_accounts = set()
    # Add overrides
    pl_overrides = {"3105301", "3105302", "3105702", "3105703", "3105711", "3105834", "3105835", "3108112", "3103111", "3103113", "3103112", "3103122", "3105704"}
    mapped_pl_accounts.update(pl_overrides)
    if os.path.exists(map_pl_path):
        try:
            df_pl = pd.read_excel(map_pl_path, dtype=str)
            if not df_pl.empty:
                col_pl_cuenta = next((c for c in df_pl.columns if 'cuenta' in c.lower()), df_pl.columns[0])
                for _, row in df_pl.iterrows():
                    acc_id = str(row[col_pl_cuenta]).strip()
                    is_mapped = False
                    for col in df_pl.columns[1:]:
                        val = row[col]
                        if pd.notna(val) and str(val).strip() != "" and str(val).strip().lower() != "nan":
                            is_mapped = True
                            break
                    if is_mapped:
                        mapped_pl_accounts.add(acc_id)
        except Exception:
            pass

    # 3. Verify each account
    cuentas_no_mapeadas = set()
    for acc in accounts_to_check:
        acc_str = str(acc).strip()
        if acc_str in plan_cuentas:
            tipo = plan_tipo_map.get(acc_str)
            tipo_lower = str(tipo).strip().lower()
            if "balance" in tipo_lower:
                if acc_str not in mapped_bal_accounts:
                    cuentas_no_mapeadas.add(acc_str)
            elif "resultado" in tipo_lower:
                if acc_str not in mapped_pl_accounts:
                    cuentas_no_mapeadas.add(acc_str)
                    
    return cuentas_no_mapeadas

def render(empresa_seleccionada, empresa_path):
    global_opt = "🌐 [GLOBAL] Configuración General"
    is_global = (empresa_seleccionada == global_opt)
        
    st.title("Carga de Datos")
    st.write("Centraliza la carga de todos los insumos necesarios para el ciclo contable.")
    
    if "success_msg" in st.session_state:
        st.success(st.session_state.pop("success_msg"))
    
    if is_global:
        tab_plan = st.tabs(["Plan de Cuentas"])[0]
        tab_tb = None
        tab_pl = None
    else:
        tab_tb, tab_pl = st.tabs(["Trial Balance", "P&L"])
        tab_plan = None

    if tab_plan is not None:
        with tab_plan:
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

            uploaded_plan = st.file_uploader("Selecciona el Plan de Cuentas (Excel con columnas 'Cuenta' y 'Tipo')", type=["xlsx", "xls"], key="up_plan")
            if uploaded_plan is not None:
                df_new = pd.read_excel(uploaded_plan, dtype=str)
                df_new.columns = [str(c).strip() for c in df_new.columns]
                
                if 'Cuenta' in df_new.columns and 'Tipo' in df_new.columns:
                    if 'plan_cuentas_df' in st.session_state:
                        df_existente = st.session_state['plan_cuentas_df']
                        cuentas_existentes = set(df_existente['Cuenta'])
                        cuentas_nuevas = set(df_new['Cuenta'])
                        
                        duplicadas = cuentas_existentes.intersection(cuentas_nuevas)
                        nuevas_puras = cuentas_nuevas - cuentas_existentes
                        
                        # --- ALINEACIÓN DE COLUMNAS (Evitar duplicidad Nombre vs Name) ---
                        # Identificar la columna de descripción principal en el DF existente
                        desc_col_existente = next((c for c in df_existente.columns if c not in ['Cuenta', 'Tipo']), None)
                        
                        # Si hay una columna descriptiva en el DF nuevo, renombrarla para que coincida
                        if desc_col_existente:
                            desc_col_nuevo = next((c for c in df_new.columns if c not in ['Cuenta', 'Tipo']), None)
                            if desc_col_nuevo and desc_col_nuevo != desc_col_existente:
                                df_new.rename(columns={desc_col_nuevo: desc_col_existente}, inplace=True)
                        
                        if duplicadas:
                            st.warning(f"⚠️ Se detectaron {len(duplicadas)} cuentas que ya existen en el plan actual y {len(nuevas_puras)} cuentas totalmente nuevas.")
                            accion_dup = st.radio("¿Qué deseas hacer con las cuentas duplicadas?", 
                                                 ["Omitir (mantener datos antiguos)", "Reemplazar (actualizar con este archivo)"])
                            
                            if st.button("Confirmar Carga Adicional", type="primary"):
                                if "Omitir" in accion_dup:
                                    df_new_filtered = df_new[~df_new['Cuenta'].isin(duplicadas)]
                                    df_final = pd.concat([df_existente, df_new_filtered], ignore_index=True)
                                else:
                                    df_existente_filtered = df_existente[~df_existente['Cuenta'].isin(duplicadas)]
                                    df_final = pd.concat([df_existente_filtered, df_new], ignore_index=True)
                                
                                df_final = sort_accounts(df_final, 'Cuenta', 'Tipo')
                                st.session_state['plan_cuentas_df'] = df_final
                                df_final.to_excel(os.path.join(empresa_path, "plan_cuentas.xlsx"), index=False)
                                if is_global:
                                    propagate_global_file("plan_cuentas.xlsx", os.path.dirname(empresa_path))
                                st.session_state['success_msg'] = f"✅ Plan de Cuentas actualizado exitosamente para **{empresa_seleccionada}**. Total maestro: **{len(df_final)}** cuentas."
                                st.rerun()
                        else:
                            st.info(f"Se detectaron {len(df_new)} cuentas nuevas para añadir al plan existente.")
                            if st.button("Añadir Nuevas Cuentas", type="primary"):
                                df_final = pd.concat([df_existente, df_new], ignore_index=True)
                                df_final = sort_accounts(df_final, 'Cuenta', 'Tipo')
                                st.session_state['plan_cuentas_df'] = df_final
                                df_final.to_excel(os.path.join(empresa_path, "plan_cuentas.xlsx"), index=False)
                                if is_global:
                                    propagate_global_file("plan_cuentas.xlsx", os.path.dirname(empresa_path))
                                st.session_state['success_msg'] = f"✅ Plan de Cuentas actualizado exitosamente para **{empresa_seleccionada}**. Se añadieron **{len(df_new)}** cuentas (Total maestro: **{len(df_final)}** cuentas)."
                                st.rerun()
                    else:
                        df_new = sort_accounts(df_new, 'Cuenta', 'Tipo')
                        st.session_state['plan_cuentas_df'] = df_new
                        plan_path = os.path.join(empresa_path, "plan_cuentas.xlsx")
                        df_new.to_excel(plan_path, index=False)
                        if is_global:
                            propagate_global_file("plan_cuentas.xlsx", os.path.dirname(empresa_path))
                        st.session_state['success_msg'] = f"✅ Plan de Cuentas cargado exitosamente para **{empresa_seleccionada}**. Maestro inicial: **{len(df_new)}** cuentas."
                        st.rerun()
                else:
                    st.error("❌ El archivo Excel debe contener las columnas obligatorias 'Cuenta' y 'Tipo'.")

            # Mostrar tabs siempre
            sub_tab1, sub_tab2 = st.tabs(["📄 Cargar/Estado Plan", "👁️ Ver Plan de Cuentas"])
            with sub_tab1:
                if 'plan_cuentas_df' in st.session_state:
                    st.success(f"🟢 **Plan de Cuentas Activo y Validado:** Se encuentra cargado el Plan de Cuentas Maestro para la empresa **{empresa_seleccionada}** con un total de **{len(st.session_state['plan_cuentas_df'])}** cuentas.")
                else:
                    st.info(f"ℹ️ **Sin Plan de Cuentas:** Aún no se ha cargado ningún Plan de Cuentas para **{empresa_seleccionada}**.")
            with sub_tab2:
                if 'plan_cuentas_df' in st.session_state:
                    st.dataframe(st.session_state['plan_cuentas_df'])
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

    if tab_tb is not None:
        with tab_tb:
            st.subheader("Importación de Trial Balance")
            st.write("Sube el balance de comprobación exportado de tu ERP corporativo.")

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

            sub_tb1, sub_tb2 = st.tabs(["📄 Cargar/Estado TB", "👁️ Ver Trial Balance"])

            with sub_tb1:
                from src.models.trial_balance_db import TrialBalanceDB
                TrialBalanceDB.initialize()

                col_y, col_m = st.columns(2)
                with col_y:
                    upload_year = st.selectbox("Año a cargar", ["2023", "2024", "2025", "2026", "2027"], index=2, key="upl_year")
                with col_m:
                    upload_month = st.selectbox("Mes a cargar", ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"], key="upl_month")

                periodo_str = f"{upload_year}-{upload_month}"
                st.info(f"Los datos se asimilarán al periodo contable: {periodo_str}")

                from src.core.cierre_engine import es_periodo_cerrado
                cerrado = es_periodo_cerrado(empresa_seleccionada, periodo_str)

                if cerrado:
                    st.error(f"⚠️ **Periodo Bloqueado:** El periodo {periodo_str} se encuentra actualmente **CERRADO** en el histórico. Para volver a cargar o extraer datos, debes reabrir el periodo desde la pestaña de **Históricos**.")
                else:
                    uploaded_file = st.file_uploader("Selecciona un archivo Excel", type=["xlsx", "xls"], key="up_tb")

                    st.markdown("---")
                    if st.button("🔌 Extracción de data desde ERP", type="secondary"):
                        from src.integrations.erp_adapter import ErpAdapterLogger
                        if ErpAdapterLogger.is_configured(empresa_path):
                            with st.spinner("Conectando con la API del ERP..."):
                                import time
                                start_time = time.time()
                                time.sleep(1.5)
                                df_erp, erp_name = ErpAdapterLogger.fetch_trial_balance(empresa_path, upload_year, upload_month)

                                # Validación de cuentas huérfanas antes de guardar (Bloqueo Duro)
                                cuentas_huerfanas = set()
                                cuentas_no_mapeadas = set()
                                if 'plan_cuentas_df' in st.session_state:
                                    plan_cuentas = set(st.session_state['plan_cuentas_df']['Cuenta'].astype(str).str.strip())
                                    tb_cuentas = set(df_erp['cuenta_id'].astype(str).str.strip())
                                    cuentas_huerfanas = tb_cuentas - plan_cuentas

                                if cuentas_huerfanas:
                                    st.error(f"❌ ERROR DE AUDITORÍA: Se detectaron {len(cuentas_huerfanas)} cuentas en la extracción que NO existen en el Plan de Cuentas maestro.")
                                    st.write("Cuentas huérfanas encontradas:", sorted(list(cuentas_huerfanas)))
                                    st.error("❌ Extracción cancelada: Debes actualizar estas cuentas en el Plan de Cuentas maestro antes de poder importar la data de este periodo.")
                                else:
                                    # Validación de cuentas no mapeadas (Bloqueo Duro)
                                    cuentas_no_mapeadas = check_unmapped_accounts(empresa_path, tb_cuentas, st.session_state.get('plan_cuentas_df'))
                                    if cuentas_no_mapeadas:
                                        st.error(f"❌ ERROR DE AUDITORÍA: Se detectaron {len(cuentas_no_mapeadas)} cuentas que NO están mapeadas en el sistema.")
                                        st.write("Cuentas sin clasificar encontradas:", sorted(list(cuentas_no_mapeadas)))
                                        st.error("❌ Extracción cancelada: Debes clasificar estas cuentas en el maestro de Balance (map_balance.xlsx) o P&L (map_pl.xlsx) antes de poder importar este Balance.")
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
                        if st.session_state.get('last_tb_file_sig') != file_sig:
                            # Directorio data está garantizado desde nuestro backend setup
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

                                    # Validación de cuentas huérfanas antes de guardar (Bloqueo Duro)
                                    cuentas_huerfanas = set()
                                    cuentas_no_mapeadas = set()
                                    if 'plan_cuentas_df' in st.session_state:
                                        plan_cuentas = set(st.session_state['plan_cuentas_df']['Cuenta'].astype(str).str.strip())
                                        tb_cuentas = set(df['cuenta_id'].astype(str).str.strip())
                                        cuentas_huerfanas = tb_cuentas - plan_cuentas

                                    if cuentas_huerfanas:
                                        st.error(f"❌ ERROR DE AUDITORÍA: Se detectaron {len(cuentas_huerfanas)} cuentas en el Balance subido que NO existen en el Plan de Cuentas maestro.")
                                        st.write("Cuentas huérfanas encontradas:", sorted(list(cuentas_huerfanas)))
                                        st.error("❌ Carga bloqueada: Debes actualizar estas cuentas en el Plan de Cuentas maestro antes de poder importar este Balance.")
                                    else:
                                        # Validación de cuentas no mapeadas (Bloqueo Duro)
                                        cuentas_no_mapeadas = check_unmapped_accounts(empresa_path, tb_cuentas, st.session_state.get('plan_cuentas_df'))
                                        if cuentas_no_mapeadas:
                                            st.error(f"❌ ERROR DE AUDITORÍA: Se detectaron {len(cuentas_no_mapeadas)} cuentas que NO están mapeadas en el sistema.")
                                            st.write("Cuentas sin clasificar encontradas:", sorted(list(cuentas_no_mapeadas)))
                                            st.error("❌ Carga bloqueada: Debes clasificar estas cuentas en el maestro de Balance (map_balance.xlsx) o P&L (map_pl.xlsx) antes de poder importar este Balance.")
                                        else:
                                            TrialBalanceDB.save_trial_balance(empresa_seleccionada, periodo_str, df)
                                            elapsed_time = time.time() - start_time
                                            st.session_state['last_tb_file_sig'] = file_sig
                                            st.session_state['success_msg'] = f"✅ Archivo cargado e Ingestado exitosamente en base de datos. Se guardaron {len(df)} registros para {periodo_str} (Tiempo de ejecución: {elapsed_time:.2f} segundos). Balance de comprobación cuadrado y validado exitosamente. Auditoría superada: El 100% de las cuentas de este Balance existen en el Plan Maestro."

                                            # Guardamos en la memoria del navegador temporalmente
                                            st.session_state['tb_df'] = TrialBalanceDB.get_trial_balance(empresa_seleccionada, periodo_str)
                                            st.rerun()
                            except Exception as e:
                                error_msg = str(e)
                                if error_msg.startswith("🚨"):
                                    st.error(error_msg)
                                else:
                                    st.error(f"Error procesando el archivo: {error_msg}")

            with sub_tb2:
                if 'tb_df' in st.session_state and st.session_state['tb_df'] is not None:
                    tb_df_to_show = st.session_state['tb_df'].copy()
                    cuenta_col_tb = next((c for c in tb_df_to_show.columns if "cuenta" in str(c).lower() and "nombre" not in str(c).lower()), "cuenta_id")
                    desc_col_tb = next((c for c in tb_df_to_show.columns if "nombre" in str(c).lower() or "desc" in str(c).lower()), "descripcion")
                    numeric_cols_tb = [c for c in tb_df_to_show.columns if c not in [cuenta_col_tb, desc_col_tb]]

                    for col in numeric_cols_tb:
                        tb_df_to_show[col] = tb_df_to_show[col].apply(lambda x: f"{int(round(pd.to_numeric(x, errors='coerce') or 0.0)):,}".replace(",", "."))

                    column_config_tb = {
                        c: st.column_config.TextColumn(c)
                        for c in numeric_cols_tb
                    }
                    column_config_tb[cuenta_col_tb] = st.column_config.TextColumn("N° de Cuenta")
                    column_config_tb[desc_col_tb] = st.column_config.TextColumn("Nombre de la cuenta")

                    st.dataframe(
                        tb_df_to_show,
                        use_container_width=True,
                        column_config=column_config_tb,
                        key=f"df_tb_show_{empresa_seleccionada}_{periodo_str}"
                    )
                    excel_data = df_to_excel_bytes(st.session_state['tb_df'], "Trial Balance")
                    st.download_button(
                        label="📥 Descargar Trial Balance en Excel",
                        data=excel_data,
                        file_name="trial_balance_activo.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="btn_dl_tb"
                    )
                else:
                    st.info("Aún no se ha cargado ningún Trial Balance.")

    if tab_pl is not None:
        with tab_pl:
            st.subheader("Cubo de Estado de Resultados (P&L)")
            st.write("Sube el cubo analítico de Pérdidas y Ganancias (ventas por centro de costo, unidad de negocio, etc.).")

            # Obtener dinámicamente las columnas de P&L de la base de datos de taxonomía para esta empresa
            from src.models.database import SessionLocal
            from src.models.taxonomy_master import TaxonomyMasterRecord

            db_pl = SessionLocal()
            try:
                tax_recs = db_pl.query(TaxonomyMasterRecord.nombre_linea_es).filter_by(
                    empresa=empresa_seleccionada,
                    reporte_destino="P&L"
                ).order_by(TaxonomyMasterRecord.id_reporte).all()
                db_cols = [r[0] for r in tax_recs if r[0]]
                db_cols = list(dict.fromkeys(db_cols)) # Eliminar duplicados preservando orden

                # Ordenar los rubros según el orden estándar del reporte
                PL_ORDER_LIST = [
                    "ingresos de arriendo fibra optica",
                    "ingresos de actividades ordinarias",
                    "costo de ventas",
                    "acceso a infraestructura fibra optica",
                    "costos de uso fibra optica",
                    "depreciacion operacional",
                    "depreciacion y amortizacion operacional",
                    "otros ingresos por funcion",
                    "costos de distribucion",
                    "gastos de administracion",
                    "depreciacion y amortizaciones",
                    "otros egresos por funcion",
                    "resultado por inversion en empresas relacionadas",
                    "ingresos financieros",
                    "ingresos financieros con empresas relacionadas",
                    "ingresos financieros ic",
                    "costos financieros",
                    "diferencias de cambio",
                    "resultado por unidad de reajuste",
                    "resultados por unidades de reajuste",
                    "ganancia (perdida) por impuesto a las ganancias",
                    "resultado por impuestos a las ganancias"
                ]

                def get_pl_sort_key(name):
                    if not name:
                        return 9999
                    norm = name.lower().strip().replace('á','a').replace('é','e').replace('í','i').replace('ó','o').replace('ú','u')
                    if norm in PL_ORDER_LIST:
                        return PL_ORDER_LIST.index(norm)
                    for idx, item in enumerate(PL_ORDER_LIST):
                        if item in norm or norm in item:
                            return idx
                    return 9999

                db_cols = sorted(db_cols, key=get_pl_sort_key)
            except Exception:
                db_cols = []
            finally:
                db_pl.close()

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

            sub_pl1, sub_pl2 = st.tabs(["📄 Cargar/Estado P&L", "👁️ Ver P&L"])
            with sub_pl1:
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

                col_y_pl, col_m_pl = st.columns(2)
                with col_y_pl:
                    upload_year_pl = st.selectbox("Año a cargar", ["2023", "2024", "2025", "2026", "2027"], index=2, key="upl_year_pl")
                with col_m_pl:
                    upload_month_pl = st.selectbox("Mes a cargar", ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"], key="upl_month_pl")

                periodo_str_pl = f"{upload_year_pl}-{upload_month_pl}"
                st.info(f"Los datos de este Cubo P&L se archivarán en el periodo contable: {periodo_str_pl}")

                from src.core.cierre_engine import es_periodo_cerrado
                cerrado_pl = es_periodo_cerrado(empresa_seleccionada, periodo_str_pl)

                if cerrado_pl:
                    st.error(f"⚠️ **Periodo Bloqueado:** El periodo {periodo_str_pl} se encuentra actualmente **CERRADO** en el histórico. Para volver a cargar datos de P&L, debes reabrir el periodo desde la pestaña de **Históricos**.")
                else:
                    up_pl_cubo = st.file_uploader("Selecciona el Cubo P&L (Excel)", type=["xlsx", "xls"], key="up_pl_cubo")

                    # Cargar maestro de mapeos para la empresa
                    map_pl_path = os.path.join(empresa_path, "map_pl.xlsx")
                    if os.path.exists(map_pl_path):
                        df_map_master = pd.read_excel(map_pl_path)
                    else:
                        df_map_master = None

                    df_processed = None
                    current_file_id = None

                    if up_pl_cubo is not None:
                        try:
                            # 1. Leer hojas del Excel para permitir seleccionar pestaña si hay varias
                            with st.spinner("Analizando la estructura del archivo Excel..."):
                                xls = pd.ExcelFile(up_pl_cubo, engine='openpyxl')
                                sheet_names = xls.sheet_names

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

                            # 2. Leer la hoja seleccionada
                            import time
                            start_time = time.time()
                            with st.spinner(f"Cargando transacciones de la pestaña '{selected_sheet}' (esto puede tomar unos segundos)..."):
                                df_raw = pd.read_excel(up_pl_cubo, sheet_name=selected_sheet, dtype=str, engine='openpyxl')
                                df_raw.columns = [str(c).strip() for c in df_raw.columns]

                            # Detectar formato: ¿Es Cubo de Odoo o Formato P&L ancho clásico?
                            has_eerr_col = any("eerr" in c.lower() or "informe_eerr" in c.lower() for c in df_raw.columns)

                            if has_eerr_col:
                                format_detected = "Cubo transaccional original de Odoo"
                                st.info(f"📂 Formato detectado: **{format_detected}** en la pestaña **'{selected_sheet}'**.")
                                if df_map_master is None:
                                    st.warning("⚠️ No se encontró el archivo maestro 'map_pl.xlsx' de la empresa activa. Se usará clasificación directa de Odoo sin mapeo específico.")

                                with st.spinner("Agrupando transacciones y aplicando reglas de mapeo IFRS (YTD)..."):
                                    from src.core.pl_cubo_processor import process_odoo_cubo
                                    df_processed = process_odoo_cubo(df_raw, upload_year_pl, upload_month_pl, df_map_master, standard_categories=pl_rubros)
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
                                st.success(f"✅ Archivo Excel procesado con éxito (Tiempo de ejecución: {elapsed_time:.2f} segundos).")
                                current_file_id = f"{up_pl_cubo.name}_{selected_sheet}_{periodo_str_pl}_{up_pl_cubo.size}"
                        except Exception as e:
                            st.error(f"❌ Error al procesar el archivo Excel: {e}")

                    # Check for DB saved data if no file is uploaded
                    db_pl_df = PlCuboDB.get_pl_cubo(empresa_seleccionada, periodo_str_pl)

                    if up_pl_cubo is None and db_pl_df is not None and not db_pl_df.empty:
                        st.write("---")
                        st.info(f"💡 Se detectaron datos guardados en la base de datos para el periodo **{periodo_str_pl}**.")
                        if st.button("✏️ Cargar y Editar datos guardados de este periodo", use_container_width=True):
                            st.session_state['pl_edit_df'] = add_totals_row(db_pl_df.copy())
                            st.session_state['pl_edit_file_id'] = f"saved_db_{periodo_str_pl}"
                            st.rerun()

                    # Determine what to load into editor
                    if current_file_id:
                        if st.session_state.get('pl_edit_file_id') != current_file_id:
                            st.session_state['pl_edit_df'] = add_totals_row(df_processed)
                            st.session_state['pl_edit_file_id'] = current_file_id

                    # If we are editing either uploaded file or DB data
                    is_editing_db = st.session_state.get('pl_edit_file_id') == f"saved_db_{periodo_str_pl}"
                    has_active_editor = (up_pl_cubo is not None) or is_editing_db

                    if has_active_editor and 'pl_edit_df' in st.session_state:
                        if is_editing_db:
                            st.warning("⚠️ **Modo Edición BD**: Estás editando directamente los saldos guardados en la base de datos.")
                            if st.button("❌ Cancelar edición BD"):
                                del st.session_state['pl_edit_df']
                                del st.session_state['pl_edit_file_id']
                                st.rerun()

                        st.markdown("### 📝 Vista Previa y Edición Manual del P&L")
                        st.write("Puedes modificar los montos haciendo doble clic en las celdas antes de guardar en la base de datos.")

                        # Determine columns
                        cuenta_col = next((c for c in st.session_state['pl_edit_df'].columns if "cuenta" in str(c).lower() and "nombre" not in str(c).lower()), "N° de cuenta")
                        desc_col = next((c for c in st.session_state['pl_edit_df'].columns if "nombre" in str(c).lower()), "Nombre de la cuenta")
                        numeric_cols = [c for c in st.session_state['pl_edit_df'].columns if c not in [cuenta_col, desc_col]]

                        # Convert to string format with dots for editor view to bypass sprintf errors
                        df_to_edit = st.session_state['pl_edit_df'].copy()
                        for col in numeric_cols:
                            df_to_edit[col] = df_to_edit[col].apply(lambda x: f"{int(round(pd.to_numeric(x, errors='coerce') or 0.0)):,}".replace(",", "."))

                        # Validación cruzada con Plan de Cuentas (Bloqueo Duro)
                        pl_huerfanas = set()
                        pl_no_mapeadas = set()
                        if 'plan_cuentas_df' in st.session_state:
                            plan_cuentas = set(st.session_state['plan_cuentas_df']['Cuenta'].astype(str).str.strip())
                            editor_accounts = set(st.session_state['pl_edit_df'][st.session_state['pl_edit_df'][cuenta_col].astype(str).str.strip().str.upper() != "TOTAL"][cuenta_col].astype(str).str.strip())
                            pl_huerfanas = editor_accounts - plan_cuentas
                            if not pl_huerfanas:
                                pl_no_mapeadas = check_unmapped_accounts(empresa_path, editor_accounts, st.session_state['plan_cuentas_df'])

                        # Auto-guardar en BD tras carga exitosa si la auditoría es limpia
                        if current_file_id and not pl_huerfanas and not pl_no_mapeadas:
                            if st.session_state.get('last_pl_saved_file_id') != current_file_id and df_processed is not None:
                                try:
                                    saved_cnt = PlCuboDB.save_pl_cubo(empresa_seleccionada, periodo_str_pl, df_processed)
                                    st.session_state['last_pl_saved_file_id'] = current_file_id
                                    st.session_state['pl_df'] = PlCuboDB.get_pl_cubo(empresa_seleccionada, periodo_str_pl)
                                    st.success(f"✅ Cubo P&L archivado e ingestado exitosamente en base de datos para el periodo {periodo_str_pl} ({saved_cnt} registros dimensionales).")
                                except Exception as save_err:
                                    st.error(f"❌ Error al guardar en base de datos: {save_err}")

                        # Dynamic key to reset editor state when data/file/period changes
                        editor_key = f"pl_editor_{current_file_id}" if current_file_id else f"pl_editor_{periodo_str_pl}"

                        if pl_huerfanas:
                            st.error(f"❌ ERROR DE AUDITORÍA: Se detectaron {len(pl_huerfanas)} cuentas en el P&L que NO existen en el Plan de Cuentas maestro.")
                            st.write("Cuentas huérfanas encontradas:", sorted(list(pl_huerfanas)))
                            st.error("❌ Guardado y Edición bloqueados: Debes actualizar estas cuentas en el Plan de Cuentas maestro antes de poder trabajar con esta información.")

                            column_config = {
                                c: st.column_config.TextColumn(c, disabled=True)
                                for c in numeric_cols
                            }
                            column_config[cuenta_col] = st.column_config.TextColumn(cuenta_col, disabled=True)
                            column_config[desc_col] = st.column_config.TextColumn(desc_col, disabled=True)

                            st.data_editor(
                                df_to_edit,
                                num_rows="dynamic",
                                use_container_width=True,
                                column_config=column_config,
                                key=editor_key + "_blocked",
                                disabled=True
                            )
                            st.warning("⚠️ Debes actualizar el Plan de Cuentas Maestro en la Configuración Global y re-subir la data para desbloquear el guardado.")
                        elif pl_no_mapeadas:
                            st.error(f"❌ ERROR DE AUDITORÍA: Se detectaron {len(pl_no_mapeadas)} cuentas en el P&L que NO están mapeadas en el maestro de P&L.")
                            st.write("Cuentas no mapeadas encontradas:", sorted(list(pl_no_mapeadas)))
                            st.error("❌ Guardado y Edición bloqueados: Debes clasificar estas cuentas en el maestro de P&L (map_pl.xlsx) antes de poder guardar los datos de este periodo.")

                            column_config = {
                                c: st.column_config.TextColumn(c, disabled=True)
                                for c in numeric_cols
                            }
                            column_config[cuenta_col] = st.column_config.TextColumn(cuenta_col, disabled=True)
                            column_config[desc_col] = st.column_config.TextColumn(desc_col, disabled=True)

                            st.data_editor(
                                df_to_edit,
                                num_rows="dynamic",
                                use_container_width=True,
                                column_config=column_config,
                                key=editor_key + "_blocked_map",
                                disabled=True
                            )
                            st.warning("⚠️ Debes clasificar las cuentas en el maestro de P&L (map_pl.xlsx) y re-subir la data para desbloquear el guardado.")
                        else:
                            edited_df = st.data_editor(
                                df_to_edit,
                                num_rows="dynamic",
                                use_container_width=True,
                                key=editor_key
                            )

                            st.write("")
                            if st.button("💾 Guardar Cubo P&L en Base de Datos", type="primary", use_container_width=True):
                                try:
                                    df_to_save = edited_df.copy()
                                    # Descartar fila TOTAL si existe
                                    df_to_save = df_to_save[df_to_save[cuenta_col].astype(str).str.strip().str.upper() != "TOTAL"].copy()
                                    for col in numeric_cols:
                                        df_to_save[col] = df_to_save[col].astype(str).str.replace(".", "").str.replace(",", ".")
                                        df_to_save[col] = pd.to_numeric(df_to_save[col], errors='coerce').fillna(0.0)

                                    saved_count = PlCuboDB.save_pl_cubo(empresa_seleccionada, periodo_str_pl, df_to_save)
                                    st.session_state['pl_df'] = PlCuboDB.get_pl_cubo(empresa_seleccionada, periodo_str_pl)
                                    msg = f"✅ Cubo P&L guardado exitosamente en la base de datos para el periodo {periodo_str_pl} ({saved_count} registros dimensionales)."
                                    st.session_state['success_msg'] = msg
                                    st.toast(msg, icon="✅")
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"❌ Error al guardar en base de datos: {e}")

            with sub_pl2:
                db_pl_show = PlCuboDB.get_pl_cubo(empresa_seleccionada, periodo_str_pl)
                if db_pl_show is not None and not db_pl_show.empty:
                    st.subheader(f"📊 Vista de Cubo P&L Guardado en Base de Datos ({periodo_str_pl})")
                    pl_show_df = db_pl_show.copy()
                    c_cuenta = next((c for c in pl_show_df.columns if "cuenta" in str(c).lower() and "nombre" not in str(c).lower()), "N° de cuenta")
                    c_desc = next((c for c in pl_show_df.columns if "nombre" in str(c).lower() or "desc" in str(c).lower()), "Nombre de la cuenta")
                    n_cols = [c for c in pl_show_df.columns if c not in [c_cuenta, c_desc]]

                    for col in n_cols:
                        pl_show_df[col] = pl_show_df[col].apply(lambda x: f"{int(round(pd.to_numeric(x, errors='coerce') or 0.0)):,}".replace(",", "."))

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
                else:
                    st.info(f"Aún no se ha guardado ningún Cubo P&L para el periodo **{periodo_str_pl}**.")
