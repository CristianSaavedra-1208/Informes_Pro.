import streamlit as st
import pandas as pd
import os
from io import BytesIO
from src.core.excel_utils import format_periodo
from src.core.word_template_engine import WordTemplateEngine

def generate_sample_template():
    """Genera en memoria una plantilla Word de ejemplo con las etiquetas soportadas."""
    from docx import Document
    doc = Document()
    doc.add_heading("INFORME FINANCIERO CORPORATIVO", level=0)
    
    p = doc.add_paragraph()
    p.add_run("Este documento ha sido generado automáticamente utilizando la plantilla de ejemplo y el motor de reportes de ").italic = True
    p.add_run("Informes Pro").bold = True
    p.add_run(".")
    
    doc.add_heading("1. Resumen General", level=1)
    doc.add_paragraph("Empresa Activa: {{EMPRESA_ACTIVA}}")
    doc.add_paragraph("Período de Análisis YTD: {{PERIODO_ACTUAL}}")
    doc.add_paragraph("Período Comparativo: {{PERIODO_COMPARATIVO}}")
    doc.add_paragraph("Utilidad Neta de las Operaciones: {{UTILIDAD_NETA}}")
    
    doc.add_heading("2. Estado de Situación Financiera (Balance)", level=1)
    doc.add_paragraph("El siguiente cuadro detalla la situación financiera consolidada/individual de la compañía:")
    doc.add_paragraph("{{#BALANCE}}")
    
    doc.add_heading("3. Estado de Resultados (P&L)", level=1)
    doc.add_paragraph("Desempeño de ingresos, costos y gastos acumulados para el ejercicio:")
    doc.add_paragraph("{{#ER}}")
    
    doc.add_heading("4. Estado de Flujos de Efectivo (EFE)", level=1)
    doc.add_paragraph("Origen y aplicación de recursos de efectivo:")
    doc.add_paragraph("{{#EFE}}")
    
    doc.add_heading("5. Notas Explicativas a los Estados Financieros", level=1)
    
    doc.add_heading("Nota 4: Efectivo y Equivalentes de Efectivo", level=2)
    doc.add_paragraph("{{#N04}}")
    
    doc.add_heading("Nota 14: Partes Relacionadas", level=2)
    doc.add_paragraph("{{#N14}}")
    
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

def render(empresa_seleccionada, empresa_path):
    if empresa_seleccionada == "[GLOBAL] Configuración General" or "GLOBAL" in empresa_seleccionada:
        st.warning("Módulo de Sociedad Activa: Por favor, selecciona una empresa de trabajo específica (ej. Pacifico SpA) en la barra lateral izquierda para acceder a esta sección.")
        st.stop()
        
    # Detectar cambio de empresa activa y resetear la plantilla del reporte anterior
    if st.session_state.get('corp_empresa_activa') != empresa_seleccionada:
        for k in ['corp_word_output_bytes', 'corp_word_filename', 'rep_corp_file']:
            if k in st.session_state:
                del st.session_state[k]
        st.session_state['corp_empresa_activa'] = empresa_seleccionada
        
    st.title("Generación de reportes Word")
    st.markdown("""
    Genera informes corporativos y de gestión a partir de **plantillas personalizadas de Word (`.docx`)**. 
    El motor inyectará automáticamente los estados financieros consolidados o individuales y las notas explicativas en las marcas correspondientes.
    """)
    
    # Acordeón de Guía de Etiquetas
    with st.expander("Guía de Diseño de Plantilla (Etiquetas Soportadas)", expanded=False):
        st.markdown("""
        Puedes diseñar tu plantilla de Word en cualquier formato, incluir portadas, logotipos corporativos, 
        encabezados y tablas estilizadas. Para inyectar datos dinámicos, utiliza las siguientes etiquetas en el texto:
        
        #### Variables Simples (Párrafos o Celdas):
        * `{{EMPRESA_ACTIVA}}` : Nombre de la empresa o grupo seleccionado.
        * `{{PERIODO_ACTUAL}}` : Período actual formateado (ej. *31 de Diciembre de 2025*).
        * `{{PERIODO_COMPARATIVO}}` : Período comparativo formateado (ej. *31 de Diciembre de 2024* o vacío si no se selecciona).
        * `{{UTILIDAD_NETA}}` : Utilidad neta formateada con el factor de escala seleccionado.
        
        #### 📊 Reportes Principales:
        Inserta la etiqueta en un párrafo individual:
        
        **Opción A: Con Encabezado Institucional Completo** (Empresa, Título formal, Período, Unidad y Tabla):
        * `{{#BALANCE}}` (o `{{#ESF}}`) : Balance Clasificado completo.
        * `{{#ER}}` : Estado de Resultados completo.
        * `{{#EFE}}` : Estado de Flujos de Efectivo completo.
        * `{{#PATRIMONIO}}` : Estado de Cambios en el Patrimonio completo.
        * `{{#ORI}}` : Estado de Resultados Integrales completo.
        
        **Opción B: Solo la Tabla** (Sin encabezado superior, para cuando uses tus propios títulos):
        * `{{#TABLA_BALANCE}}` (o `{{#BALANCE_TABLA}}`) : Únicamente la tabla de Balance.
        * `{{#TABLA_ER}}` (o `{{#ER_TABLA}}`) : Únicamente la tabla de Resultados.
        * `{{#TABLA_EFE}}` (o `{{#EFE_TABLA}}`) : Únicamente la tabla de Flujos.
        * `{{#TABLA_PATRIMONIO}}` (o `{{#PATRIMONIO_TABLA}}`) : Únicamente la tabla de Patrimonio.
        * `{{#TABLA_ORI}}` (o `{{#ORI_TABLA}}`) : Únicamente la tabla de ORI.
        
        #### 📑 Notas Explicativas:
        Inserta la etiqueta de la nota para inyectar su respectiva explicación de texto y cuadros asociados:
        * `{{#N04}}` : Efectivo y Equivalentes.
        * `{{#N06}}` : Deudores Comerciales.
        * `{{#N14}}` : Partes Relacionadas.
        * *(Cualquier nota registrada con su respectivo prefijo `#NXX`)*
        """)
        
        # Ofrecer descarga de la plantilla de ejemplo
        sample_buf = generate_sample_template()
        st.download_button(
            label="📥 Descargar Plantilla de Ejemplo (.docx)",
            data=sample_buf,
            file_name="Plantilla_Ejemplo_Reporte.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
            key="dl_sample_template"
        )
        
    st.write("")
    
    # 1. Obtener períodos disponibles unificados (Memoria Activa + Histórico)
    from src.models.trial_balance_db import TrialBalanceDB
    periodos_hist = TrialBalanceDB.get_available_periods(empresa_seleccionada)
    if not periodos_hist:
        periodos_hist = ["2026-12", "2026-07", "2026-06", "2026-03", "2025-12"]
        
    # 2. Configuración de parámetros en columnas
    st.subheader("⚙️ Configuración del Informe")
    col1, col2, col3 = st.columns([2, 2, 2])
    
    with col1:
        periodo_actual = st.selectbox(
            "Período Actual:",
            periodos_hist,
            index=0,
            format_func=format_periodo,
            key="rep_corp_act"
        )
    with col2:
        periodos_comp_opts = ["Ninguno"] + periodos_hist
        # Pre-seleccionar el segundo periodo si existe
        default_comp_idx = 1 if len(periodos_hist) > 1 else 0
        periodo_comp = st.selectbox(
            "Período Comparativo (Opcional):",
            periodos_comp_opts,
            index=default_comp_idx,
            format_func=lambda x: "Ninguno" if x == "Ninguno" else format_periodo(x),
            key="rep_corp_comp"
        )
    with col3:
        unidad = st.selectbox(
            "Unidad de Medida / Escala:",
            ["Miles de pesos (M$)", "Unidades ($)", "Millones de pesos (MM$)"],
            index=0,
            key="rep_corp_scale"
        )
        
        if "Miles" in unidad:
            scale_factor = 1000.0
        elif "Millones" in unidad:
            scale_factor = 1000000.0
        else:
            scale_factor = 1.0

    from src.core.snapshot_vault_manager import SnapshotVaultManager
    is_closed_snap = SnapshotVaultManager.has_active_snapshot(empresa_seleccionada, periodo_actual)
    if is_closed_snap:
        st.caption("🔒 **Modo Bóveda Oficial**: Este período se encuentra cerrado y congelado. La inyección en Word se realizará instantáneamente (< 1s) garantizando inmutabilidad y auditoría.")
    else:
        st.caption("🔓 **Modo Dinámico Vivo**: Este período está abierto en elaboración. Los estados financieros y notas se calcularán en vivo.")

    st.write("")
    
    # Opciones de configuración
    col_opt1, col_opt2 = st.columns([1, 1])
    with col_opt1:
        use_ifrs_auto = st.checkbox(
            "Reglas comparativas IFRS automáticas", 
            value=True, 
            help="Si está activo: el Balance comparará con el cierre de diciembre del año anterior, y el Estado de Resultados / Flujo de Efectivo / Notas compararán con el mismo mes del año anterior. Si se desactiva: se utilizará la opción del selector 'Período Comparativo' para todos los rubros.",
            key="rep_corp_ifrs_auto"
        )
    with col_opt2:
        incluir_encabezados = st.checkbox(
            "Incluir encabezados institucionales en tablas", 
            value=True, 
            help="Si está activo: añade automáticamente el bloque superior centrado (Nombre Empresa, Título del Estado, Períodos y Unidad) antes de cada tabla. Si se desmarca: se generarán las tablas limpias sin encabezado superior.",
            key="rep_corp_incluir_encabezados"
        )
    
    st.write("")
    
    # 3. Panel de Carga de Plantilla
    st.subheader("📂 Plantilla Word")
    
    def reset_corp_report_state():
        for k in ['corp_word_output_bytes', 'corp_word_filename', 'rep_corp_file', 'corp_word_time_str', 'corp_word_tags_count']:
            if k in st.session_state:
                del st.session_state[k]

    uploaded_file = st.file_uploader(
        "Carga tu archivo de plantilla Word (.docx):",
        type=["docx"],
        key="rep_corp_file"
    )
    
    if uploaded_file is not None:
        st.success("Plantilla cargada con éxito. Listo para generar el reporte.")
        
        # Botón para procesar
        if st.button("Generar Reporte Corporativo", type="primary", use_container_width=True):
            import time
            start_time = time.time()

            progress_bar = st.progress(0.0)
            status_box = st.empty()
            status_box.info("⏳ Inicializando motor y analizando etiquetas en la plantilla...")

            def on_progress(current, total, current_tag):
                pct = min(1.0, current / max(1, total))
                progress_bar.progress(pct)
                elapsed_now = time.time() - start_time
                status_box.info(f"⏳ Inyectando: **`{current_tag}`** ({current}/{total}) • Tiempo transcurrido: **{elapsed_now:.1f}s**...")

            try:
                from src.core.word_template_engine import WordTemplateEngine
                
                # Crear una copia de los bytes subidos
                uploaded_file.seek(0)
                template_bytes = BytesIO(uploaded_file.read())
                
                # Instanciar el motor y ejecutar
                engine = WordTemplateEngine(template_bytes)
                output_buffer = engine.process(
                    empresa_activa=empresa_seleccionada,
                    periodo_actual=periodo_actual,
                    periodo_comp=periodo_comp,
                    scale_factor=scale_factor,
                    use_ifrs_auto=use_ifrs_auto,
                    include_headers=incluir_encabezados,
                    progress_callback=on_progress
                )
                
                total_elapsed = time.time() - start_time
                progress_bar.progress(1.0)
                status_box.empty()
                
                out_bytes = output_buffer.getvalue() if hasattr(output_buffer, 'getvalue') else output_buffer
                st.session_state['corp_word_output_bytes'] = out_bytes
                st.session_state['corp_word_filename'] = f"Reporte_Corporativo_{empresa_seleccionada.replace('[GRUPO] ', '').replace(' ', '_')}_{periodo_actual}.docx"
                
                # Formatear tiempo de ejecución
                mins = int(total_elapsed // 60)
                secs = total_elapsed % 60
                time_str = f"{mins} min {secs:.1f} s" if mins > 0 else f"{total_elapsed:.1f} segundos"
                st.session_state['corp_word_time_str'] = time_str
                st.session_state['corp_word_tags_count'] = getattr(engine, 'total_tags_found', None)
                
                tags_msg = f" ({st.session_state['corp_word_tags_count']} etiquetas inyectadas)" if st.session_state['corp_word_tags_count'] else ""
                st.success(f"✅ ¡Reporte generado con éxito en **{time_str}**{tags_msg}! Presione descargar reporte para emitir Word actualizado.")
            except Exception as e:
                progress_bar.empty()
                status_box.empty()
                st.warning(f"⚠️ **Información del Sistema:** {str(e)}")
                    
        if 'corp_word_output_bytes' in st.session_state:
            st.write("---")
            if 'corp_word_time_str' in st.session_state:
                col_info1, col_info2 = st.columns(2)
                with col_info1:
                    st.info(f"⏱️ **Tiempo de procesamiento:** {st.session_state['corp_word_time_str']}")
                with col_info2:
                    size_kb = len(st.session_state['corp_word_output_bytes']) / 1024
                    tags_txt = f" • {st.session_state.get('corp_word_tags_count')} etiquetas" if st.session_state.get('corp_word_tags_count') else ""
                    st.info(f"📄 **Documento generado:** {size_kb:.1f} KB{tags_txt}")
                    
            col_dl1, col_dl2 = st.columns([3, 1])
            with col_dl1:
                st.download_button(
                    label="Descargar Reporte Final (.docx)",
                    data=st.session_state['corp_word_output_bytes'],
                    file_name=st.session_state.get('corp_word_filename', 'Reporte_Corporativo.docx'),
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary",
                    use_container_width=True,
                    on_click=reset_corp_report_state,
                    key="dl_processed_report"
                )
            with col_dl2:
                if st.button("Limpiar / Listo para otro reporte", use_container_width=True, key="btn_reset_corp"):
                    reset_corp_report_state()
                    st.rerun()
    else:
        st.info("Sube una plantilla en formato Word (.docx) o descarga nuestra plantilla de ejemplo para comenzar a generar reportes.")
