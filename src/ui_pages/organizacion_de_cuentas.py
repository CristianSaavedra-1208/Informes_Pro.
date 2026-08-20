import streamlit as st
import pandas as pd
import os
from src.core.excel_utils import df_to_excel_bytes, sort_accounts, heal_mapping_fields

def render(empresa_seleccionada, empresa_path):
    # Cargar automáticamente mapeos existentes desde el disco a st.session_state si existen y no están cargados
    if 'map_balance_df' not in st.session_state or st.session_state.get('map_balance_df_empresa') != empresa_seleccionada:
        map_bal_path = os.path.join(empresa_path, "map_balance.xlsx")
        if os.path.exists(map_bal_path):
            try:
                st.session_state['map_balance_df'] = pd.read_excel(map_bal_path, dtype=str)
                st.session_state['map_balance_df_empresa'] = empresa_seleccionada
            except:
                pass

    if 'map_pl_df' not in st.session_state or st.session_state.get('map_pl_df_empresa') != empresa_seleccionada:
        map_pl_path = os.path.join(empresa_path, "map_pl.xlsx")
        if os.path.exists(map_pl_path):
            try:
                # Use openpyxl engine to properly read special chars in headers on Windows
                st.session_state['map_pl_df'] = pd.read_excel(map_pl_path, dtype=str, engine='openpyxl')
                st.session_state['map_pl_df_empresa'] = empresa_seleccionada
            except:
                pass

    st.title("🔀 Motor de Organización de Cuentas")
    st.write("Configura la equivalencia de cuentas para Balance general y el Estado de Resultados (P&L).")
    
    global_opt = "🌐 [GLOBAL] Configuración General"
    is_global = (empresa_seleccionada == global_opt)
    real_empresas = sorted([d for d in os.listdir(os.path.dirname(empresa_path)) if os.path.isdir(os.path.join(os.path.dirname(empresa_path), d))])

    def propagate_global_file(file_name):
        import shutil
        source_file = os.path.join(empresa_path, file_name)
        if os.path.exists(source_file):
            for co in real_empresas:
                if co == "Pacifico SpA":
                    continue
                dest_dir = os.path.join(os.path.dirname(empresa_path), co)
                if os.path.isdir(dest_dir):
                    shutil.copy2(source_file, os.path.join(dest_dir, file_name))

    # Mostrar notificaciones persistidas desde el estado de sesión al inicio de la página
    if "rubro_creado_msg" in st.session_state:
        st.success(st.session_state.pop("rubro_creado_msg"))
    if "rubro_cambio_msg" in st.session_state:
        st.success(st.session_state.pop("rubro_cambio_msg"))
    if "success_msg" in st.session_state:
        st.success(st.session_state.pop("success_msg"))
        
    if is_global:
        tab_names = [
            "ℹ️ Guía de Uso",
            "🏛️ Mapeo y Clasificación Balance", 
            "📈 Mapeo y Clasificación P&L", 
            "🏛️ Diccionario Maestro (Nuevo)", 
            "✍️ Mapeo Manual"
        ]
        tabs = st.tabs(tab_names)
        tab_intro, tab_bal, tab_pl, tab_maestro, tab_manual = tabs
        tab_rubros = None
        tab_plantillas = None
        tab_sabana = None
    else:
        tab_names = [
            "ℹ️ Guía de Uso",
            "📋 Administrador de Rubros",
            "📄 Gestión de Plantillas",
            "📊 Sábanas de Auditoría"
        ]
        tabs = st.tabs(tab_names)
        tab_intro, tab_rubros, tab_plantillas, tab_sabana = tabs
        tab_bal = None
        tab_pl = None
        tab_maestro = None
        tab_manual = None
    
    with tab_intro:
        st.subheader("ℹ️ Guía de Uso y Utilidad del Motor")
        
        st.markdown("""
        El **Motor de Organización de Cuentas** es el núcleo de consistencia del sistema. Su propósito es conectar la contabilidad cruda (saldos transaccionales) con la presentación formal de los Estados Financieros estandarizados.
        """)
        
        col_intro1, col_intro2 = st.columns(2)
        
        with col_intro1:
            st.info("""
            ### 🎯 ¿Para qué sirve?
            Cuando cargas un *Trial Balance* desde tu ERP, las cuentas contables son sumamente detalladas. Este módulo te permite **agrupar y direccionar** esos saldos detallados hacia los rubros oficiales exigidos bajo las normas de presentación **IFRS/NIIF** u otras taxonomías locales.
            """)
            
            st.markdown("""
            ### 📋 Flujo de Trabajo Recomendado
            1. **Carga de Datos:** Importa tu Trial Balance y P&L del mes.
            2. **Mapeo / Clasificación:** Utiliza este módulo para asegurar que cada cuenta contable esté asignada a un rubro del Balance y/o P&L.
            3. **Ajustes:** Si hay descuadres o cuentas nuevas, utiliza la pestaña de **Mapeo Manual**.
            4. **Emisión:** Genera los reportes oficiales en Excel y Word.
            """)
            
        with col_intro2:
            st.write("### 🗂️ Utilidad de cada pestaña:")
            st.markdown("""
            * **🏛️ Mapeo y Clasificación Balance:** Sube planillas masivas para asociar cuentas de Activo, Pasivo y Patrimonio con sus clasificaciones en el Balance General (relación 1:1).
            * **📈 Mapeo y Clasificación P&L:** Configura la matriz multidimensional para estructurar las notas operacionales del Estado de Resultados.
            * **🏛️ Diccionario Maestro:** Administra el índice jerárquico global de reportes, regulando qué líneas aparecen en los informes de Excel/Word y qué notas explicativas tienen asignadas.
            * **✍️ Mapeo Manual:** Clasifica rápidamente cuentas nuevas o corrige asignaciones individuales sin necesidad de volver a subir planillas completas.
            * **📋 Administrador de Rubros:** Crea nuevos rubros (ej. *Activos Biológicos*) y autogenera códigos de taxonomía para guardarlos en la base de datos SQL.
            * **📄 Gestión de Plantillas:** Sube, descarga y elimina los archivos Excel base que sirven como esqueleto estructurado para generar y dar formato a los Estados Financieros individuales y consolidados.
            """)
            
        st.divider()

    if tab_bal is not None:
        with tab_bal:
            st.subheader("Mapeo de Balance")
            st.write("Sube el archivo Excel con la clasificación única para las cuentas de activo, pasivo y patrimonio.")
            with st.expander("Ver columnas requeridas y descargar plantilla"):
                st.write("Columnas clave requeridas en el Excel: `N° de Cuenta`, `Nombre cuenta`, `Clasificación balance`, `nota 1`, y `nota 2`.")
            
                mock_bal_path = os.path.join(empresa_path, "mock_map_balance.xlsx")
                if os.path.exists(mock_bal_path):
                    with open(mock_bal_path, "rb") as file:
                        st.download_button(
                            label="📥 Descargar Plantilla Mapeo Balance",
                            data=file,
                            file_name="plantilla_mapeo_balance.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="dl_map_bal"
                        )
            
            # Layout con 2 columnas para alinearse con Gestión de Plantillas
            col_dn, col_up = st.columns(2)
        
            with col_dn:
                st.write("##### 1. Descargar Mapeo Actual")
                map_bal_path = os.path.join(empresa_path, "map_balance.xlsx")
                if os.path.exists(map_bal_path):
                    with open(map_bal_path, "rb") as file:
                        st.download_button(
                            label="📥 Descargar mapeo_balance.xlsx",
                            data=file,
                            file_name="map_balance.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="btn_dl_map_bal_disk",
                            use_container_width=True
                        )
                    st.success("✅ Mapeo detectado en el servidor.")
                else:
                    st.warning("⚠️ No se detectó un mapeo de Balance en el servidor.")
                
            with col_up:
                st.write("##### 2. Subir Nuevo Mapeo (Sobreescribir)")
                up_bal = st.file_uploader("Arrastra tu Excel de Mapeo de Balance aquí", type=["xlsx", "xls"], key="up_map_bal_f")
            
            if up_bal is not None:
                # Use openpyxl engine to properly read 'N°' and 'ó' characters on Windows without crashing
                df_bal = pd.read_excel(up_bal, dtype=str, engine='openpyxl')
                # Cleanup headers
                df_bal.columns = [str(c).strip() for c in df_bal.columns]
                # Flexible case-insensitive column detection
                cuenta_col = next((c for c in df_bal.columns if "cuenta" in c.lower()), None)
                clasif_col = next((c for c in df_bal.columns if "clasifica" in c.lower() and "flujo" not in c.lower()), None)
                if not clasif_col:
                    clasif_col = next((c for c in df_bal.columns if "balance" in c.lower() and "flujo" not in c.lower()), None)
            
                if cuenta_col and clasif_col:
                    # Validacion estricta: Una cuenta de balance no puede mapearse dos veces
                    # Ignorar nulos para no fallar por filas vacías
                    df_valid = df_bal.dropna(subset=[cuenta_col])
                    duplicated_mask = df_valid.duplicated(subset=[cuenta_col], keep=False)
                    if duplicated_mask.any():
                        dup_accounts = df_valid[duplicated_mask][cuenta_col].unique()
                        st.error(f"🚨 ERROR DE INTEGRIDAD: Se detectaron cuentas duplicadas en tu Excel. A diferencia del P&L, en el Balance una cuenta tiene asignación estricta 1:1 y NO puede apuntar a dos rubros. Por favor elimina los duplicados de las siguientes cuentas: {', '.join(dup_accounts)}")
                    else:
                        from src.models.taxonomy_generator import process_mapping_for_taxonomy
                        new_codes = process_mapping_for_taxonomy(df_bal, "Balance", empresa_seleccionada)
                        if new_codes > 0:
                            st.info(f"✨ Auto-Descubrimiento Activo: Se auto-generaron e inscribieron {new_codes} nuevos códigos en la Bóveda Maestra de Taxonomía basándose en las palabras de tu mapeo de Balance.")

                        df_bal = sort_accounts(df_bal, cuenta_col)
                        df_bal = heal_mapping_fields(df_bal)
                        st.success(f"✅ Archivo procesado. {len(df_bal)} cuentas mapeadas para Balance.")
                        st.session_state['map_balance_df'] = df_bal
                        df_bal.to_excel(os.path.join(empresa_path, "map_balance.xlsx"), index=False)
                        if is_global: propagate_global_file("map_balance.xlsx")
                else:
                    st.error(f"❌ Estructura inválida. Asegúrate de tener las columnas requeridas. Columnas detectadas: {list(df_bal.columns)}")
                
            if 'map_balance_df' in st.session_state:
                if 'ed_bal_ver' not in st.session_state:
                    st.session_state['ed_bal_ver'] = 1

                with st.expander("✏️ Editor de Datos Profesional (Modificar/Eliminar sin resubir Excel)", expanded=True):
                    st.info("💡 **Indicación importante**: Haz doble clic en cualquier celda para editar. Tras realizar tus modificaciones en las casillas, **presiona Enter o haz clic fuera de la celda antes de pulsar el botón de guardar**.")
                    edited_bal = st.data_editor(
                        st.session_state['map_balance_df'], 
                        num_rows="dynamic", 
                        use_container_width=True, 
                        key=f"ed_bal_{st.session_state['ed_bal_ver']}"
                    )
                
                    if st.button("💾 Guardar Cambios Directos en la Base", type="primary", key="save_ed_bal"):
                        # Validación
                        c_col = next((c for c in edited_bal.columns if "Cuenta" in c), None)
                        if c_col:
                            df_valid = edited_bal.dropna(subset=[c_col])
                            if df_valid.duplicated(subset=[c_col], keep=False).any():
                                st.error("🚨 Siguen existiendo cuentas duplicadas. Elimínalas antes de guardar.")
                            else:
                                c_col = next((c for c in edited_bal.columns if "Cuenta" in c), "Cuenta")
                                edited_bal = sort_accounts(edited_bal, c_col)
                                edited_bal = heal_mapping_fields(edited_bal)
                                st.session_state['map_balance_df'] = edited_bal
                                st.session_state['map_balance_df_empresa'] = empresa_seleccionada
                                
                                map_bal_file = os.path.join(empresa_path, "map_balance.xlsx")
                                edited_bal.to_excel(map_bal_file, index=False)
                                if is_global: propagate_global_file("map_balance.xlsx")
                                
                                # Invalidar caché de lectura de Excel y sincronizar taxonomía
                                st.cache_data.clear()
                                from src.models.taxonomy_generator import process_mapping_for_taxonomy
                                process_mapping_for_taxonomy(edited_bal, "Balance", empresa_seleccionada)
                                
                                # Incrementar versión del widget para forzar render fresco
                                st.session_state['ed_bal_ver'] += 1
                                st.session_state['success_msg'] = "✅ Cambios en el mapeo de Balance guardados con éxito."
                                st.rerun()

                excel_data = df_to_excel_bytes(st.session_state['map_balance_df'], "Mapeo Balance")
                st.download_button(
                    label="📥 Descargar Mapeo Balance en Excel",
                    data=excel_data,
                    file_name="mapeo_balance_activo.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_dl_map_bal"
                )
            
    if tab_pl is not None:
        with tab_pl:
            st.subheader("Mapeo de P&L (Estado de Resultados)")
            st.write("Sube el archivo Excel con las definiciones operacionales y administrativas (e.g. Notas de Costos, Administración).")
        
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
            plantilla_cols = ["N° de cuenta", "Nombre de la cuenta"] + pl_rubros

            # Sincronizar st.session_state['map_pl_df'] para que tenga todas las columnas dinámicas (para el editor directo)
            if 'map_pl_df' in st.session_state and st.session_state['map_pl_df'] is not None:
                df_active = st.session_state['map_pl_df'].copy()
                cols_to_add = [c for c in pl_rubros if c not in df_active.columns]
                if cols_to_add:
                    for col in cols_to_add:
                        df_active[col] = None
                    st.session_state['map_pl_df'] = df_active

            with st.expander("Ver columnas requeridas y descargar plantilla"):
                st.write(f"Columnas requeridas: `{', '.join(plantilla_cols)}`.")

                mock_pl_df = pd.DataFrame(columns=plantilla_cols)
                excel_data = df_to_excel_bytes(mock_pl_df, 'Plantilla Mapeo PL')
            
                st.download_button(
                    label="📥 Descargar Plantilla Mapeo P&L",
                    data=excel_data,
                    file_name="plantilla_mapeo_pl.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_map_pl"
                )

            # Layout con 2 columnas para alinearse con Gestión de Plantillas
            col_dn, col_up = st.columns(2)
        
            with col_dn:
                st.write("##### 1. Descargar Mapeo Actual")
                map_pl_path = os.path.join(empresa_path, "map_pl.xlsx")
                if os.path.exists(map_pl_path):
                    with open(map_pl_path, "rb") as file:
                        st.download_button(
                            label="📥 Descargar mapeo_pl.xlsx",
                            data=file,
                            file_name="map_pl.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="btn_dl_map_pl_disk",
                            use_container_width=True
                        )
                    st.success("✅ Mapeo detectado en el servidor.")
                else:
                    st.warning("⚠️ No se detectó un mapeo de P&L en el servidor.")
                
            with col_up:
                st.write("##### 2. Subir Nuevo Mapeo (Sobreescribir)")
                up_pl = st.file_uploader("Arrastra tu Excel de Mapeo de P&L aquí", type=["xlsx", "xls"], key="up_map_pl_f")
            
            if up_pl is not None:
                # Use openpyxl engine to properly read 'N°' and 'ó' characters on Windows without crashing
                df_pl = pd.read_excel(up_pl, dtype=str, engine='openpyxl')
                df_pl.columns = [str(c).strip() for c in df_pl.columns]
            
                cuenta_col = next((c for c in df_pl.columns if "Cuenta" in c or "cuenta" in c), None)
            
                # Check that at least some of the specific P&L columns from the template are present
                expected_pl_cols = [c.strip().lower() for c in plantilla_cols if "cuenta" not in c.lower()]
                has_pl_cols = any(c.strip().lower() in expected_pl_cols for c in df_pl.columns)
            
                if cuenta_col and has_pl_cols:
                    from src.models.taxonomy_generator import process_mapping_for_taxonomy
                    new_codes = process_mapping_for_taxonomy(df_pl, "PL", empresa_seleccionada)
                    if new_codes > 0:
                        st.info(f"✨ Auto-Descubrimiento Activo: Se auto-generaron e inscribieron {new_codes} nuevos códigos en la Bóveda Maestra de Taxonomía basándose en las palabras de tu mapeo de P&L.")
                    
                    df_pl = sort_accounts(df_pl, cuenta_col)
                    st.success(f"✅ Archivo procesado. {len(df_pl)} cuentas mapeadas para P&L.")
                    st.session_state['map_pl_df'] = df_pl
                    df_pl.to_excel(os.path.join(empresa_path, "map_pl.xlsx"), index=False)
                    if is_global: propagate_global_file("map_pl.xlsx")
                else:
                    st.error(f"❌ Estructura inválida. Asegúrate de tener las columnas indicadas en la plantilla. Columnas detectadas: {list(df_pl.columns)}")

            if 'map_pl_df' in st.session_state:
                if 'ed_pl_ver' not in st.session_state:
                    st.session_state['ed_pl_ver'] = 1

                with st.expander("✏️ Editor de Datos Profesional (Modificar/Eliminar sin resubir Excel)", expanded=True):
                    st.info("💡 **Indicación importante**: Haz doble clic en cualquier celda para editar. Tras realizar tus modificaciones en las casillas, **presiona Enter o haz clic fuera de la celda antes de pulsar el botón de guardar**.")
                    edited_pl = st.data_editor(
                        st.session_state['map_pl_df'], 
                        num_rows="dynamic", 
                        use_container_width=True, 
                        key=f"ed_pl_{st.session_state['ed_pl_ver']}"
                    )
                
                    if st.button("💾 Guardar Cambios Directos en la Base", type="primary", key="save_ed_pl"):
                        c_col = next((c for c in edited_pl.columns if "Cuenta" in c or "cuenta" in c), "N° de cuenta")
                        edited_pl = sort_accounts(edited_pl, c_col)
                        st.session_state['map_pl_df'] = edited_pl
                        st.session_state['map_pl_df_empresa'] = empresa_seleccionada
                        
                        map_pl_file = os.path.join(empresa_path, "map_pl.xlsx")
                        edited_pl.to_excel(map_pl_file, index=False)
                        if is_global: propagate_global_file("map_pl.xlsx")
                        
                        # Invalidar caché de lectura de Excel y sincronizar taxonomía
                        st.cache_data.clear()
                        from src.models.taxonomy_generator import process_mapping_for_taxonomy
                        process_mapping_for_taxonomy(edited_pl, "PL", empresa_seleccionada)
                        
                        # Incrementar versión del widget para forzar render fresco
                        st.session_state['ed_pl_ver'] += 1
                        msg = "✅ Cambios en el mapeo de P&L guardados con éxito."
                        st.session_state['success_msg'] = msg
                        st.toast(msg, icon="✅")
                        st.rerun()

                excel_data = df_to_excel_bytes(st.session_state['map_pl_df'], "Mapeo PL")
                st.download_button(
                    label="📥 Descargar Mapeo P&L en Excel",
                    data=excel_data,
                    file_name="mapeo_pl_activo.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_dl_map_pl"
                )

    if tab_maestro is not None:
        with tab_maestro:
            st.subheader("Diccionario / Taxonomía Maestra de Reportes")
            st.write("Sube el archivo Excel con la estructura jerárquica de líneas para Balance, P&L y Notas.")
        
            sub_m1, sub_m2 = st.tabs(["📄 Cargar Maestro", "👁️ Ver Diccionario Maestro"])
        
            with sub_m1:
                with st.expander("Ver columnas requeridas y descargar plantilla"):
                    plantilla_maestro_cols = [
                        "ID_Reporte", 
                        "Reporte_Destino", 
                        "Nombre_Linea_ES", 
                        "ID_Nota_Asociada", 
                        "Desglose_Nota_ES", 
                        "Nombre_Idioma_1", 
                        "Nombre_Idioma_2"
                    ]
                    st.write(f"Columnas requeridas: `{', '.join(plantilla_maestro_cols)}`.")
                
                    mock_maestro_df = pd.DataFrame(columns=plantilla_maestro_cols)
                    excel_data_maestro = df_to_excel_bytes(mock_maestro_df, 'Plantilla Maestro Taxonomia')
                
                    st.download_button(
                        label="📥 Descargar Plantilla Maestro de Taxonomía",
                        data=excel_data_maestro,
                        file_name="plantilla_maestro_taxonomia.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_map_maestro"
                    )
                
                up_maestro = st.file_uploader("Selecciona el Excel de Maestro Estructural", type=["xlsx", "xls"], key="up_maestro_f")
            
                if up_maestro is not None:
                    df_maestro = pd.read_excel(up_maestro, dtype=str, engine='openpyxl')
                    df_maestro.columns = [str(c).strip() for c in df_maestro.columns]
                
                    if "ID_Reporte" in df_maestro.columns and "Nombre_Linea_ES" in df_maestro.columns:
                        st.success(f"✅ Archivo estructural válido. {len(df_maestro)} líneas detectadas.")
                        if st.button("💾 Procesar e Inyectar a Base de Datos Central (Upsert)", type="primary"):
                            try:
                                from src.models.database import SessionLocal
                                from src.models.taxonomy_master import TaxonomyMasterRecord
                                db = SessionLocal()
                            
                                count_upsert = 0
                                for _, row in df_maestro.iterrows():
                                    if pd.isna(row.get("ID_Reporte")) or str(row.get("ID_Reporte")).strip() == "":
                                        continue
                                    
                                    existing = db.query(TaxonomyMasterRecord).filter_by(
                                        empresa=empresa_seleccionada,
                                        id_reporte=str(row["ID_Reporte"]).strip()
                                    ).first()
                                
                                    if not existing:
                                        existing = TaxonomyMasterRecord(
                                            empresa=empresa_seleccionada,
                                            id_reporte=str(row["ID_Reporte"]).strip()
                                        )
                                        db.add(existing)
                                
                                    existing.reporte_destino = str(row.get("Reporte_Destino", ""))
                                    existing.nombre_linea_es = str(row.get("Nombre_Linea_ES", ""))
                                    existing.id_nota_asociada = str(row.get("ID_Nota_Asociada", "")) if not pd.isna(row.get("ID_Nota_Asociada")) else None
                                    existing.desglose_nota_es = str(row.get("Desglose_Nota_ES", "")) if not pd.isna(row.get("Desglose_Nota_ES")) else None
                                    existing.nombre_idioma_1 = str(row.get("Nombre_Idioma_1", "")) if not pd.isna(row.get("Nombre_Idioma_1")) else None
                                    existing.nombre_idioma_2 = str(row.get("Nombre_Idioma_2", "")) if not pd.isna(row.get("Nombre_Idioma_2")) else None
                                    count_upsert += 1
                                
                                db.commit()
                                db.close()
                                st.session_state['success_msg'] = f"✅ Éxito: {count_upsert} líneas del Diccionario Maestro guardadas/actualizadas inteligentemente en el Dicionario Central (SQL)."
                                st.rerun()
                            except Exception as e:
                                st.error(f"Error procesando base de datos: {e}")
                    else:
                        st.error("❌ El archivo no contiene las columnas clave 'ID_Reporte' o 'Nombre_Linea_ES'.")

            with sub_m2:
                try:
                    from src.models.database import SessionLocal
                    from src.models.taxonomy_master import TaxonomyMasterRecord
                    db = SessionLocal()
                    records = db.query(TaxonomyMasterRecord).filter_by(empresa=empresa_seleccionada).all()
                    db.close()
                    if records:
                        data_maestro = [{
                            'ID_Reporte': r.id_reporte,
                            'Reporte_Destino': r.reporte_destino,
                            'Nombre_Linea_ES': r.nombre_linea_es,
                            'ID_Nota_Asociada': r.id_nota_asociada
                        } for r in records]
                        df_maestro_db = pd.DataFrame(data_maestro)
                        st.dataframe(df_maestro_db)
                        excel_data_m = df_to_excel_bytes(df_maestro_db, "Diccionario Maestro")
                        st.download_button(
                            label="📥 Descargar Diccionario en Excel",
                            data=excel_data_m,
                            file_name="diccionario_maestro_activo.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="dl_maestro_db"
                        )
                    else:
                        st.info("Aún no hay Diccionario Maestro inyectado para esta empresa.")
                except Exception as e:
                    st.error(f"Error cargando Diccionario Maestro: {e}")

    if tab_manual is not None:
        with tab_manual:
            st.subheader("Configuración Manual / Correcciones")
            tipo_mapeo = st.radio("Tipo de Cuenta a Clasificar", ["Balance", "P&L (Estado de Resultados)"], horizontal=True)
        
            # Extraer opciones creadas dinámicamente
            opciones_bal = []
            opciones_pl = []
            try:
                from src.models.database import SessionLocal
                from src.models.taxonomy_master import TaxonomyMasterRecord
                db = SessionLocal()
                tax_bal = db.query(TaxonomyMasterRecord.nombre_linea_es).filter_by(empresa=empresa_seleccionada, reporte_destino="Balance").all()
                tax_pl = db.query(TaxonomyMasterRecord.nombre_linea_es).filter_by(empresa=empresa_seleccionada, reporte_destino="P&L").all()
                db.close()
                if tax_bal: opciones_bal = [r[0] for r in tax_bal]
                if tax_pl: opciones_pl = [r[0] for r in tax_pl]
            except Exception:
                pass
            
            # Fallbacks e inclusiones desde DF en sesión
            if 'map_balance_df' in st.session_state:
                df = st.session_state['map_balance_df']
                col_c = next((c for c in df.columns if "Clasificaci" in c and "balance" in c.lower()), None)
                if col_c:
                    opciones_bal.extend(df[col_c].dropna().unique().tolist())
                
            if 'map_pl_df' in st.session_state:
                df_pl = st.session_state['map_pl_df']
                pl_cols = [c for c in df_pl.columns if "cuenta" not in c.lower() and "detalle" not in c.lower() and "unnamed" not in c.lower()]
                opciones_pl.extend(pl_cols)
            
            # Limpiar y unificar
            opciones_bal = sorted(list(set([str(x).strip() for x in opciones_bal if str(x).strip()])))
            opciones_pl = sorted(list(set([str(x).strip() for x in opciones_pl if str(x).strip()])))
        
            # Fallback final de seguridad
            if not opciones_bal: opciones_bal = ["Activo Corriente", "Activo No Corriente", "Pasivo Corriente", "Pasivo No Corriente", "Patrimonio"]
            if not opciones_pl: opciones_pl = ["Ingresos de actividades ordinarias", "Costo de ventas", "Gastos de administración", "Depreciación y amortizaciones", "Resultado por impuestos a las ganancias"]

            # Quitamos el st.form para permitir selectores interactivos anidados
            col1, col2, col3 = st.columns([2, 3, 2])
            # 1. Cuentas dinámicas del programa (Plan Maestro + Trial Balance)
            cuentas_disponibles_set = set()
            if 'plan_cuentas_df' in st.session_state:
                cuentas_disponibles_set.update([str(x).strip() for x in st.session_state['plan_cuentas_df']['Cuenta'].dropna().unique()])
            
            if 'tb_df' in st.session_state:
                cuentas_disponibles_set.update([str(x).strip() for x in st.session_state['tb_df']['cuenta_id'].unique()])
            else:
                try:
                    from src.models.trial_balance_db import TrialBalanceDB
                    # get_available_periods and get_trial_balance
                    per = TrialBalanceDB.get_available_periods(empresa_seleccionada)
                    if per:
                        db_tb = TrialBalanceDB.get_trial_balance(empresa_seleccionada, per[-1])
                        if db_tb is not None:
                            cuentas_disponibles_set.update([str(r['cuenta_id']).strip() for _, r in db_tb.iterrows()])
                except Exception:
                    pass
                
            cuentas_disponibles = sorted(list(cuentas_disponibles_set))
            if not cuentas_disponibles:
                cuentas_disponibles = ["- No hay cuentas cargadas -"]

            cuenta_ejemplo = col1.selectbox("N° de Cuenta:", cuentas_disponibles, key="c_ej")
        
            if tipo_mapeo == "Balance":
                clasificacion = col2.selectbox("Clasificación en Balance:", opciones_bal, key="cl_bal")
            else:
                clasificacion = col2.selectbox("Naturaleza en EERR:", opciones_pl, key="cl_pl")
        
            # 2. Construir Selectores de Notas independientes para cada columna
            nota_ejemplo = ""
            nota_config_dict = {}
        
            # Helper local de normalización de acentos y caracteres de codificación
            def normalize_text_local(s):
                if not s: return ""
                s = str(s).strip().lower()
                import unicodedata
                s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
                s = s.replace('', 'o').replace('', 'a').replace('', 'e').replace('', 'i').replace('', 'u')
                s = s.replace('á','a').replace('é','e').replace('í','i').replace('ó','o').replace('ú','u').replace('ñ','n')
                return s

            if tipo_mapeo == "Balance" and 'map_balance_df' in st.session_state:
                df_b = st.session_state['map_balance_df']
                cls_col = next((c for c in df_b.columns if "clasificaci" in c.lower() and "balance" in c.lower()), "Clasificación balance")
            
                sub_df = df_b[df_b[cls_col] == clasificacion] if cls_col in df_b.columns else df_b
                nota_cols = [c for c in df_b.columns if "nota" in c.lower() and "id" not in c.lower()]
            
                for n_idx, nc in enumerate(nota_cols):
                    # Extraemos opciones: primero las específicas a la clasificación, luego las globales para no limitar al usuario
                    opciones_especificas = sorted([str(x).strip() for x in sub_df[nc].dropna().unique() if str(x).strip()])
                    opciones_globales = sorted([str(x).strip() for x in df_b[nc].dropna().unique() if str(x).strip()])
                
                    # Mantenemos orden y removemos duplicados
                    todas_opc = list(dict.fromkeys(opciones_especificas + opciones_globales))
                
                    # Añadir también las notas de base de datos registradas para esta clasificación
                    try:
                        from src.models.database import SessionLocal
                        from src.models.taxonomy_master import TaxonomyMasterRecord
                        db = SessionLocal()
                        db_notes = db.query(TaxonomyMasterRecord.nombre_linea_es, TaxonomyMasterRecord.desglose_nota_es).filter(
                            TaxonomyMasterRecord.reporte_destino == "Notas Balance"
                        ).filter(TaxonomyMasterRecord.desglose_nota_es.isnot(None), TaxonomyMasterRecord.desglose_nota_es != "").all()
                        db.close()
                    
                        target_norm = normalize_text_local(clasificacion)
                        db_note_list = []
                        for r in db_notes:
                            parent_name = str(r[0]).strip()
                            desc = str(r[1]).strip()
                            if normalize_text_local(parent_name) == target_norm and desc:
                                db_note_list.append(desc)
                        todas_opc.extend(db_note_list)
                    except Exception:
                        pass

                    todas_opc = sorted(list(set([x.strip() for x in todas_opc if x.strip()])))
                    todas_opc.insert(0, "")
                
                    sel_val = col3.selectbox(f"{nc} (Existente):", todas_opc, key=f"n_sel_{n_idx}")
                    new_val = col3.text_input(f"O escribir nueva {nc}:", key=f"n_new_{n_idx}")
                
                    final_val = new_val.strip() if new_val.strip() != "" else sel_val
                    nota_config_dict[nc] = final_val

                # Recuperar el ID de Nota Asociada genérico según la clasificación
                id_nota_c = next((c for c in df_b.columns if "id_nota" in c.lower()), "ID_Nota_Asociada")
                id_rep_c = next((c for c in df_b.columns if "id_reporte" in c.lower()), "ID_Reporte")
            
                if id_nota_c in df_b.columns: 
                    v = sub_df[id_nota_c].dropna().values
                    nota_config_dict[id_nota_c] = str(v[0]) if len(v) > 0 else ""
                if id_rep_c in df_b.columns: 
                    v = sub_df[id_rep_c].dropna().values
                    nota_config_dict[id_rep_c] = str(v[0]) if len(v) > 0 else ""
                
                import json
                nota_ejemplo = json.dumps(nota_config_dict)
            
            elif tipo_mapeo != "Balance":
                dict_notas = {}
                # Primero, obtener el parent code para esta clasificación en la empresa seleccionada
                parent_code = ""
                try:
                    from src.models.database import SessionLocal
                    from src.models.taxonomy_master import TaxonomyMasterRecord
                    db = SessionLocal()
                    parent_rec = db.query(TaxonomyMasterRecord.id_reporte).filter_by(
                        empresa=empresa_seleccionada,
                        nombre_linea_es=clasificacion,
                        reporte_destino="P&L"
                    ).first()
                    if parent_rec:
                        parent_code = parent_rec[0]
                except Exception:
                    pass

                try:
                    # Buscar notas de P&L de esta clasificación a nivel global (todas las empresas) para dar opciones completas
                    import json
                    db = SessionLocal()
                    tax_notas = db.query(TaxonomyMasterRecord.nombre_linea_es, TaxonomyMasterRecord.desglose_nota_es).filter(
                        TaxonomyMasterRecord.reporte_destino == "Notas P&L"
                    ).filter(TaxonomyMasterRecord.desglose_nota_es.isnot(None), TaxonomyMasterRecord.desglose_nota_es != "").all()
                    db.close()
                
                    target_norm = normalize_text_local(clasificacion)
                    for r in tax_notas:
                        parent_name = str(r[0]).strip()
                        desc = str(r[1]).strip()
                        if normalize_text_local(parent_name) == target_norm and desc:
                            # Guardamos asociándolo al código de la empresa activa
                            dict_notas[json.dumps({"ID_Nota_Asociada": parent_code, "Desglose_Nota_ES": desc})] = desc
                except Exception:
                    pass
                
                # Siempre permitir la opción de guardar sin nota para P&L
                dict_notas[json.dumps({"ID_Nota_Asociada": parent_code, "Desglose_Nota_ES": ""})] = "Sin Nota Asociada"
                
                llaves_notas = sorted(list(dict_notas.keys()), key=lambda k: dict_notas[k])
            
                nota_seleccionada = col3.selectbox(
                    "Nota Asociada Obligatoria:", 
                    llaves_notas, 
                    format_func=lambda x: dict_notas[x],
                    key="n_sel_pl"
                )
                nueva_nota = col3.text_input("O escribir una Nueva Nota:", key="n_new_pl")
            
                if nueva_nota.strip() != "":
                    nota_config_dict_temp = {"ID_Nota_Asociada": parent_code, "Desglose_Nota_ES": nueva_nota.strip()}
                    nota_ejemplo = json.dumps(nota_config_dict_temp)
                else:
                    nota_ejemplo = nota_seleccionada
        
            if st.button("Guardar Clasificación", type="primary"):

                if not cuenta_ejemplo or cuenta_ejemplo == "- No hay cuentas cargadas -":
                    st.warning("Debes seleccionar un N° de Cuenta válido.")
                elif not nota_ejemplo or nota_ejemplo == "":
                    st.warning("No hay notas disponibles para esta clasificación. Selecciona o configura al menos una válida.")
                else:
                    from src.core.mapping_updater import save_manual_mapping
                
                    # Conseguir el nombre de cuenta real
                    nombre_loc = cuenta_ejemplo
                    if 'tb_df' in st.session_state:
                        tb = st.session_state['tb_df']
                        n_match = tb.loc[tb['cuenta_id'].astype(str) == cuenta_ejemplo, 'descripcion']
                        if not n_match.empty:
                            nombre_loc = str(n_match.iloc[0])
                        
                    success, msg = save_manual_mapping(
                        cuenta_ejemplo=cuenta_ejemplo,
                        nombre_loc=nombre_loc,
                        clasificacion=clasificacion,
                        nota_config_str=nota_ejemplo,
                        tipo_mapeo=tipo_mapeo,
                        empresa_seleccionada=empresa_seleccionada,
                        empresa_path=empresa_path,
                        session_state=st.session_state
                    )
                
                    if success:
                        st.session_state['success_msg'] = f"✅ La cuenta '{cuenta_ejemplo}' ({nombre_loc}) ha sido emparejada a '{clasificacion}' con éxito."
                        st.rerun()
                    else:
                        st.error(f"Error técnico guardando mapeo: {msg}")

    if tab_rubros is not None:
        with tab_rubros:
            st.subheader("📋 Administrador de Rubros de Estados Financieros")
            st.write("Crea, edita o elimina rubros de forma rápida. Los códigos de taxonomía se autogeneran e inyectan automáticamente en la Bóveda Maestra de SQL.")
            
            from src.models.database import SessionLocal
            from src.models.taxonomy_master import TaxonomyMasterRecord
        
            db_rub = SessionLocal()
            try:
                # Consultar todos los rubros de esta empresa
                recs_rub = db_rub.query(TaxonomyMasterRecord).filter_by(empresa=empresa_seleccionada).filter(
                    TaxonomyMasterRecord.reporte_destino.in_(['Balance', 'P&L'])
                ).all()
            
                def get_tax_sort_key(r):
                    code = r.id_reporte or ""
                    parts = code.split('_')
                    if parts and parts[-1].isdigit():
                        return int(parts[-1])
                    return 99999
                
                recs_rub = sorted(recs_rub, key=get_tax_sort_key)
            
                if recs_rub:
                    df_rubros_exist = pd.DataFrame([{
                        "ID": r.id,
                        "Código Taxonomía": r.id_reporte,
                        "Reporte Destino": r.reporte_destino,
                        "Nombre Rubro (ES)": r.nombre_linea_es,
                        "Nombre Rubro (EN)": r.nombre_idioma_1 or "",
                        "Nota Asociada": r.id_nota_asociada or "",
                        "Desglose de Nota": r.desglose_nota_es or ""
                    } for r in recs_rub])
                
                    st.info("💡 **Edición Directa**: Puedes modificar cualquier celda (excepto el código autogenerado) o seleccionar y eliminar filas marcando el recuadro izquierdo y presionando 'Supr/Delete'. Haz clic en 'Guardar Cambios' para confirmar.")
                
                    original_ids_rub = set(df_rubros_exist["ID"])
                
                    edited_rubros_df = st.data_editor(
                        df_rubros_exist,
                        key="rubros_editor",
                        num_rows="dynamic",
                        disabled=["ID", "Código Taxonomía"],
                        use_container_width=True
                    )
                
                    col_rub_save1, col_rub_save2 = st.columns(2)
                    with col_rub_save1:
                        if st.button("💾 Guardar Cambios en Rubros", type="primary", use_container_width=True):
                            db_rub_save = SessionLocal()
                            try:
                                # 1. Detectar eliminados
                                edited_ids_rub = set(edited_rubros_df["ID"].dropna().astype(int))
                                deleted_ids_rub = original_ids_rub - edited_ids_rub
                                for d_id in deleted_ids_rub:
                                    db_rub_save.query(TaxonomyMasterRecord).filter_by(id=d_id).delete()
                                
                                # 2. Detectar y aplicar modificaciones
                                for _, row in edited_rubros_df.iterrows():
                                    row_id = row.get("ID")
                                    if pd.isna(row_id):
                                        continue # Las filas nuevas se agregan por el formulario para autogenerar su código
                                    
                                    row_id = int(row_id)
                                    orig_row = df_rubros_exist[df_rubros_exist["ID"] == row_id].iloc[0]
                                
                                    fields_changed = (
                                        str(row["Reporte Destino"]).strip() != str(orig_row["Reporte Destino"]).strip() or
                                        str(row["Nombre Rubro (ES)"]).strip() != str(orig_row["Nombre Rubro (ES)"]).strip() or
                                        str(row["Nombre Rubro (EN)"]).strip() != str(orig_row["Nombre Rubro (EN)"]).strip() or
                                        str(row["Nota Asociada"]).strip() != str(orig_row["Nota Asociada"]).strip() or
                                        str(row["Desglose de Nota"]).strip() != str(orig_row["Desglose de Nota"]).strip()
                                    )
                                
                                    if fields_changed:
                                        entry_rub = db_rub_save.query(TaxonomyMasterRecord).filter_by(id=row_id).first()
                                        if entry_rub:
                                            entry_rub.reporte_destino = str(row["Reporte Destino"]).strip()
                                            entry_rub.nombre_linea_es = str(row["Nombre Rubro (ES)"]).strip()
                                            entry_rub.nombre_idioma_1 = str(row["Nombre Rubro (EN)"]).strip() if str(row["Nombre Rubro (EN)"]).strip() else None
                                            entry_rub.id_nota_asociada = str(row["Nota Asociada"]).strip() if str(row["Nota Asociada"]).strip() else None
                                            entry_rub.desglose_nota_es = str(row["Desglose de Nota"]).strip() if str(row["Desglose de Nota"]).strip() else None
                                        
                                db_rub_save.commit()
                                st.session_state["success_msg"] = "✅ Cambios en la taxonomía/rubros guardados con éxito."
                                st.rerun()
                            except Exception as e:
                                st.error(f"Error al guardar cambios: {e}")
                            finally:
                                db_rub_save.close()
                    with col_rub_save2:
                        if st.button("🔄 Descartar Cambios en Rubros", type="secondary", use_container_width=True):
                            st.rerun()
                else:
                    st.info("No hay rubros de estados financieros registrados para esta empresa.")
            finally:
                db_rub.close()
            
            st.divider()
            st.write("### ➕ Agregar Nuevo Rubro de Estado Financiero")
        
            with st.form("form_nuevo_rubro", clear_on_submit=False):
                col_add1, col_add2 = st.columns(2)
                with col_add1:
                    new_nombre = st.text_input("Nombre del Rubro (ES) *", placeholder="Ej: Otros activos financieros corrientes")
                    new_reporte = st.selectbox("Reporte Destino *", ["Balance", "P&L"])
                with col_add2:
                    new_seccion = st.selectbox("Sección Contable (Solo para Balance) *", [
                        "Activos corrientes", 
                        "Activos no corrientes", 
                        "Pasivos corrientes", 
                        "Pasivos no corrientes", 
                        "Patrimonio",
                        "P&L"
                    ])
                    new_nombre_en = st.text_input("Traducción al Inglés (Opcional)", placeholder="Ej: Other current financial assets")
                
                col_add3, col_add4 = st.columns(2)
                with col_add3:
                    new_nota = st.text_input("Nota Asociada (Opcional)", placeholder="Ej: 8")
                with col_add4:
                    new_desglose = st.text_input("Desglose de Nota (Opcional)", placeholder="Ej: Detalle de activos financieros")
                
                if st.form_submit_button("➕ Registrar y Auto-generar Código de Taxonomía", type="primary"):
                    if not new_nombre:
                        st.error("El nombre del rubro es obligatorio.")
                    else:
                        db_add = SessionLocal()
                        try:
                            # 1. Determinar prefijo según reporte y sección
                            if new_reporte == "P&L":
                                prefix = "ER_OPE_ER_"
                            else:
                                if "activo" in new_seccion.lower():
                                    prefix = "ACT_C_ACT_"
                                elif "pasivo" in new_seccion.lower():
                                    prefix = "PAS_C_PAS_"
                                else:
                                    prefix = "PAT_C_PAT_"
                                
                            # 2. Buscar último código para ese prefijo en la base de datos
                            from sqlalchemy import text
                            result_max = db_add.execute(
                                text("SELECT id_reporte FROM taxonomy_master WHERE empresa = :emp AND id_reporte LIKE :pref"),
                                {"emp": empresa_seleccionada, "pref": f"{prefix}%"}
                            ).fetchall()
                        
                            max_seq = 0
                            for r in result_max:
                                code_str = r[0]
                                parts = code_str.split('_')
                                if parts and parts[-1].isdigit():
                                    seq = int(parts[-1])
                                    if seq > max_seq:
                                        max_seq = seq
                                    
                            next_seq = max_seq + 100 if max_seq > 0 else 100
                            new_code = f"{prefix}{next_seq}"
                        
                            # 3. Crear el registro en la base de datos
                            new_tax = TaxonomyMasterRecord(
                                empresa=empresa_seleccionada,
                                id_reporte=new_code,
                                reporte_destino=new_reporte,
                                nombre_linea_es=new_nombre.strip(),
                                nombre_idioma_1=new_nombre_en.strip() if new_nombre_en else None,
                                id_nota_asociada=new_nota.strip() if new_nota else None,
                                desglose_nota_es=new_desglose.strip() if new_desglose else None
                            )
                            db_add.add(new_tax)
                            db_add.commit()
                            st.session_state["success_msg"] = f"✅ Rubro registrado exitosamente con el código autogenerado: **{new_code}**"
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error al registrar: {e}")
                        finally:
                            db_add.close()

    if tab_plantillas is not None:
        with tab_plantillas:
            st.subheader("📄 Gestión de Plantillas de Reportes")
            st.write("Sube y administra las plantillas Excel base (esqueletos estructurales) utilizadas para inyectar los saldos y generar los Estados Financieros.")
        
            st.warning("⚠️ **Nota de Advertencia:** Si vas a incorporar un nuevo rubro en las plantillas Excel, este **debe estar previamente registrado en la pestaña 'Administrador de Rubros'** (con su respectiva Nota e ID de taxonomía). De lo contrario, las cuentas contables no podrán mapearse ni asociarse correctamente en el sistema.")
        
            is_group_selected = empresa_seleccionada.startswith("[GRUPO]")
        
            if is_group_selected:
                grupo_name = empresa_seleccionada.replace("[GRUPO] ", "").strip()
                grupo_folder = os.path.join("data", "empresas", f"[GRUPO] {grupo_name}")
                os.makedirs(grupo_folder, exist_ok=True)
            
                st.write(f"Gestiona las plantillas Excel base para el Grupo de Consolidación activo: **{grupo_name}**.")
                st.info(f"📁 **Directorio del Grupo**: `{grupo_folder}`")
            
                plantillas_mapping = {
                    "Balance Clasificado": "Balance clasificado.xlsx",
                    "Estado de Resultados (P&L)": "Estado de Resultados Clasificados.xlsx",
                    "Flujo de Efectivo": "Estado de Flujos de Efectivo.xlsx",
                    "Estado de Cambios en el Patrimonio": "Estado de Cambios en el Patrimonio.xlsx",
                    "Resultados Integrales (ORI)": "Estado de Resultados Integrales.xlsx"
                }
            
                sel_rep = st.selectbox("Selecciona la plantilla del Grupo a modificar", list(plantillas_mapping.keys()), key="sel_rep_group")
                plantilla_filename = plantillas_mapping[sel_rep]
                plantilla_path = os.path.join(grupo_folder, plantilla_filename)
            
                col_g_dl, col_g_up = st.columns(2)
                with col_g_dl:
                    st.write("**1. Descargar Plantilla Actual**")
                    if os.path.exists(plantilla_path):
                        with open(plantilla_path, "rb") as f:
                            file_data = f.read()
                        st.download_button(
                            label=f"📥 Descargar {plantilla_filename}",
                            data=file_data,
                            file_name=plantilla_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"dl_group_tpl_{sel_rep}"
                        )
                        st.success(f"✅ Plantilla detectada en el servidor ({len(file_data)} bytes).")
                    else:
                        st.error(f"❌ La plantilla '{plantilla_filename}' no existe aún para este grupo.")
                    
                with col_g_up:
                    st.write("**2. Subir Nueva Plantilla**")
                    uploaded_tpl = st.file_uploader(f"Arrastra tu Excel para {sel_rep} aquí", type=["xlsx"], key=f"up_group_tpl_{sel_rep}")
                    if uploaded_tpl is not None:
                        try:
                            pd.read_excel(uploaded_tpl)
                            with open(plantilla_path, "wb") as f:
                                f.write(uploaded_tpl.getbuffer())
                            st.session_state['success_msg'] = f"✅ ¡Plantilla de {sel_rep} del Grupo actualizada con éxito!"
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error al subir: {e}")
            else:
                st.write(f"Gestiona las plantillas Excel base para la empresa activa: **{empresa_seleccionada}**.")
                st.info("Descarga la plantilla actual, modifícala en Excel y vuelve a subirla para sobreescribir el diseño.")
            
                plantillas_mapping = {
                    "Balance Clasificado": "Balance clasificado.xlsx",
                    "Estado de Resultados (P&L)": "Estado de Resultados Clasificados.xlsx",
                    "Flujo de Efectivo": "Estado de Flujos de Efectivo.xlsx",
                    "Estado de Cambios en el Patrimonio": "Estado de Cambios en el Patrimonio.xlsx",
                    "Resultados Integrales (ORI)": "Estado de Resultados Integrales.xlsx",
                    "Nota Efectivo y Equivalentes": "Nota Efectivo y equivalentes.xlsx"
                }
            
                sel_rep = st.selectbox("Selecciona la plantilla de la Empresa a modificar", list(plantillas_mapping.keys()), key="sel_rep_co")
                plantilla_filename = plantillas_mapping[sel_rep]
                plantilla_path = os.path.join(empresa_path, plantilla_filename)
            
                col_co_dl, col_co_up = st.columns(2)
                with col_co_dl:
                    st.write("**1. Descargar Plantilla Actual**")
                    if os.path.exists(plantilla_path):
                        with open(plantilla_path, "rb") as f:
                            file_data = f.read()
                        st.download_button(
                            label=f"📥 Descargar {plantilla_filename}",
                            data=file_data,
                            file_name=plantilla_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"dl_co_tpl_{sel_rep}"
                        )
                        st.success(f"✅ Plantilla detectada en el servidor ({len(file_data)} bytes).")
                    else:
                        st.error(f"❌ La plantilla '{plantilla_filename}' no existe aún para esta empresa.")
                    
                with col_co_up:
                    st.write("**2. Subir Nueva Plantilla**")
                    uploaded_tpl = st.file_uploader(f"Arrastra tu Excel para {sel_rep} aquí", type=["xlsx"], key=f"up_co_tpl_{sel_rep}")
                    if uploaded_tpl is not None:
                        try:
                            # Test if excel can be read
                            pd.read_excel(uploaded_tpl)
                            with open(plantilla_path, "wb") as f:
                                f.write(uploaded_tpl.getbuffer())
                            st.session_state['success_msg'] = f"✅ ¡Plantilla de {sel_rep} actualizada con éxito!"
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error al subir: {e}")

    if tab_sabana is not None:
        with tab_sabana:
            st.subheader("📊 Sábanas de Datos Unificadas (Auditoría)")
            st.write("Genera y descarga tablas planas completas que combinan la contabilidad cruda del ERP con tus reglas de clasificación y mapeo actual.")
        
            from src.models.trial_balance_db import TrialBalanceDB
            from src.models.pl_cubo_db import PlCuboDB
            import importlib
            import src.core.sabana_builder
            importlib.reload(src.core.sabana_builder)
            from src.core.sabana_builder import (
                build_balance_sabana, build_pl_sabana,
                build_consolidated_balance_sabana, build_consolidated_pl_sabana
            )
            from src.core.excel_utils import format_periodo
            from src.models.historical_data import HistoricalDataRecord
            from src.models.database import SessionLocal
        
            is_consolidated = empresa_seleccionada.startswith("[GRUPO]")
            
            if is_consolidated:
                db = SessionLocal()
                per_recs = db.query(HistoricalDataRecord.periodo).distinct().all()
                db.close()
                available_periods = sorted([r[0] for r in per_recs], reverse=True)
                if not available_periods: available_periods = ["2026-03", "2025-12"]
            else:
                available_periods = TrialBalanceDB.get_available_periods(empresa_seleccionada)
                
            if not available_periods:
                st.warning("⚠️ No se encontraron periodos guardados en el sistema.")
            else:
                col_sab1, col_sab2 = st.columns(2)
                with col_sab1:
                    sab_periodo = st.selectbox("Selecciona el periodo", available_periods, format_func=format_periodo, key="sab_periodo_sel")
                with col_sab2:
                    sab_tipo = st.selectbox("Tipo de Sábana", ["Balance General (Sábana)", "Estado de Resultados / P&L (Sábana)"], key="sab_tipo_sel")
                
                if st.button("🔨 Construir Sábana", type="primary", key="btn_build_sab"):
                    with st.spinner("Procesando datos y cruzando mapeos..."):
                        try:
                            # Cargar mapeos maestros (local o global)
                            map_bal_path = os.path.join(empresa_path, "map_balance.xlsx")
                            if not os.path.exists(map_bal_path): map_bal_path = "map_balance.xlsx"
                            map_bal_df = pd.read_excel(map_bal_path, dtype=str) if os.path.exists(map_bal_path) else None
                            
                            map_pl_path = os.path.join(empresa_path, "map_pl.xlsx")
                            if not os.path.exists(map_pl_path): map_pl_path = "map_pl.xlsx"
                            map_pl_df = pd.read_excel(map_pl_path, dtype=str) if os.path.exists(map_pl_path) else None
                            
                            global_tpl = "Plantilla de notas_v1.xlsx"
                            if os.path.exists(global_tpl):
                                if map_bal_df is None or map_bal_df.empty:
                                    try: map_bal_df = pd.read_excel(global_tpl, sheet_name="Mapeo Balance", dtype=str)
                                    except: pass
                                if map_pl_df is None or map_pl_df.empty:
                                    try: map_pl_df = pd.read_excel(global_tpl, sheet_name="Mapeo Ctas P&L Cubo", dtype=str)
                                    except: pass

                            if is_consolidated:
                                if "balance" in sab_tipo.lower():
                                    sab_df = build_consolidated_balance_sabana(empresa_seleccionada, sab_periodo, map_bal_df)
                                else:
                                    sab_df = build_consolidated_pl_sabana(empresa_seleccionada, sab_periodo, map_pl_df)
                            else:
                                if "balance" in sab_tipo.lower():
                                    tb_df = TrialBalanceDB.get_trial_balance(empresa_seleccionada, sab_periodo)
                                    if map_bal_df is None or map_bal_df.empty:
                                        st.error("❌ No se encontró el archivo de Mapeo de Balance.")
                                        sab_df = None
                                    else:
                                        sab_df = build_balance_sabana(tb_df, map_bal_df)
                                else:
                                    pl_df = PlCuboDB.get_pl_cubo(empresa_seleccionada, sab_periodo)
                                    tb_df = TrialBalanceDB.get_trial_balance(empresa_seleccionada, sab_periodo)
                                    if map_pl_df is None or map_pl_df.empty:
                                        st.error("❌ No se encontró el archivo de Mapeo de P&L.")
                                        sab_df = None
                                    else:
                                        sab_df = build_pl_sabana(pl_df, map_pl_df, tb_df)
                                
                            if sab_df is not None and not sab_df.empty:
                                clean_name = empresa_seleccionada.replace("[GRUPO] ", "").replace(" ", "_")
                                st.session_state['generated_sab_df'] = sab_df
                                st.session_state['generated_sab_name'] = f"sabana_{'balance' if 'balance' in sab_tipo.lower() else 'pl'}_{clean_name}_{sab_periodo}.xlsx"
                                st.success("✅ ¡Sábana construida con éxito!")
                            else:
                                st.warning("⚠️ No se generaron registros para la sábana en este período.")
                        except Exception as e:
                            st.error(f"Error construyendo la sábana: {e}")
                        
                if 'generated_sab_df' in st.session_state:
                    sab_df = st.session_state['generated_sab_df']
                    sab_filename = st.session_state['generated_sab_name']
                
                    # Excel Download
                    excel_bytes = df_to_excel_bytes(sab_df, sheet_name="Sabana_Auditoria")
                    st.download_button(
                        label="📥 Descargar Sábana en Excel",
                        data=excel_bytes,
                        file_name=sab_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_btn_sab"
                    )
                
                    st.write("**Vista previa de los registros:**")
                    st.dataframe(sab_df, use_container_width=True)
