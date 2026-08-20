import streamlit as st
import pandas as pd
import os
from src.core.excel_utils import df_to_excel_bytes, format_periodo

def get_all_note_options(rubro_name: str) -> list:
    """
    Recupera exhaustivamente todas las líneas/desgloses de nota asociados a un Rubro
    de Balance o P&L desde todas las fuentes del sistema (TaxonomyMasterRecord,
    snapshots_control.json, etc.).
    """
    if not rubro_name:
        return ["-- Sin Detalle Específico (Solo Rubro Principal) --"]
        
    def _get_norm(t):
        if not t: return ""
        import unicodedata, re
        return re.sub(r'\s+', ' ', ''.join(c for c in unicodedata.normalize('NFD', str(t).strip().lower()) if unicodedata.category(c) != 'Mn'))

    norm_target = _get_norm(rubro_name)
    options = set()
    
    # 1. Búsqueda en TaxonomyMasterRecord (todas las empresas y global)
    try:
        from src.models.database import SessionLocal
        from src.models.taxonomy_master import TaxonomyMasterRecord
        db = SessionLocal()
        recs = db.query(TaxonomyMasterRecord).filter(
            TaxonomyMasterRecord.desglose_nota_es.isnot(None), 
            TaxonomyMasterRecord.desglose_nota_es != ''
        ).all()
        
        for r in recs:
            n_line = _get_norm(r.nombre_linea_es)
            if n_line and (n_line == norm_target or norm_target in n_line or n_line in norm_target):
                if r.desglose_nota_es:
                    options.add(r.desglose_nota_es.strip())
        db.close()
    except Exception:
        pass

    # 2. Búsqueda en snapshots_control.json (todas las empresas y periodos)
    try:
        import os, json
        base_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        snap_path = os.path.join(base_dir, 'data', 'snapshots_control.json')
        if os.path.exists(snap_path):
            with open(snap_path, 'r', encoding='utf-8') as f:
                snap_data = json.load(f)
                for emp_val in snap_data.values():
                    if not isinstance(emp_val, dict): continue
                    for per_val in emp_val.values():
                        if not isinstance(per_val, dict): continue
                        nctx = per_val.get('notes_ctx', {})
                        if isinstance(nctx, dict):
                            for nkey in ['nota1', 'nota2', 'pl']:
                                ndict = nctx.get(nkey, {})
                                if isinstance(ndict, dict):
                                    for k, v in ndict.items():
                                        norm_k = _get_norm(k)
                                        if norm_k and (norm_k == norm_target or norm_target in norm_k or norm_k in norm_target):
                                            if isinstance(v, list):
                                                for item in v:
                                                    if item and str(item).lower() not in ('nan', 'none', 'acct_details', 'acct_names', 'accts'):
                                                        options.add(str(item).strip())
                                            elif isinstance(v, dict):
                                                for item in v.keys():
                                                    if item and str(item).lower() not in ('nan', 'none', 'acct_details', 'acct_names', 'accts'):
                                                        options.add(str(item).strip())
    except Exception:
        pass

    # Clean, capitalize and filter out metadata noise
    cleaned = set()
    for o in options:
        if not o or o.lower() in ('nan', 'none', 'x', 'acct_details', 'acct_names', 'accts'):
            continue
        item_str = str(o).strip()
        if item_str.islower():
            item_str = item_str.capitalize()
        cleaned.add(item_str)
        
    sorted_opts = sorted(list(cleaned))
    return ["-- Sin Detalle Específico (Solo Rubro Principal) --"] + sorted_opts

def render(empresa_seleccionada, empresa_path):
    st.title("🏢 Consolidación Financiera")
    st.write("Módulo para unificar los saldos de compañías filiales y matriz mediante ajustes trazables.")
    
    # Init DB
    from src.models.database import SessionLocal
    from src.models.consolidacion import ConsolidationGroup, ConsolidationJournalEntry
    from src.models.historical_data import HistoricalDataRecord
    
    tab_conf, tab_asientos, tab_concil, tab_hoja = st.tabs([
        "⚙️ Configurar Perímetro", 
        "✍️ Comprobantes de Ajuste", 
        "🔍 Conciliación Intercompany", 
        "📊 Hoja de Trabajo Consolidada"
    ])
    
    empresas_dir = os.path.join("data", "empresas")
    empresas_disp = sorted([d for d in os.listdir(empresas_dir) if os.path.isdir(os.path.join(empresas_dir, d))])
    
    with tab_conf:
        st.subheader("Configurar Grupo de Consolidación")
        
        # Obtener grupos existentes para el selector de filiales
        db = SessionLocal()
        grupos_existentes = db.query(ConsolidationGroup).all()
        db.close()
        
        opciones_filial = list(empresas_disp)
        for g in grupos_existentes:
            opciones_filial.append(f"[GRUPO] {g.nombre_grupo}")
            
        col1, col2 = st.columns(2)
        matriz = col1.selectbox("Empresa Matriz (Holding)", empresas_disp)
        filial_str = col2.selectbox("Empresa Filial (Subsidiaria)", opciones_filial)
        nombre_grupo = st.text_input("Nombre del Grupo (ej. Grupo Pacífico)")
        
        if st.button("Guardar Perímetro", type="primary"):
            if matriz == filial_str:
                st.error("La matriz y la filial no pueden ser la misma empresa.")
            elif not nombre_grupo:
                st.error("Debes ingresar un nombre para el grupo.")
            else:
                db = SessionLocal()
                try:
                    filial_is_g = False
                    val_filial = filial_str
                    if filial_str.startswith("[GRUPO] "):
                        nombre_g_buscado = filial_str.replace("[GRUPO] ", "")
                        g_obj = db.query(ConsolidationGroup).filter_by(nombre_grupo=nombre_g_buscado).first()
                        if g_obj:
                            filial_is_g = True
                            val_filial = str(g_obj.id)
                            
                    exists = db.query(ConsolidationGroup).filter_by(nombre_grupo=nombre_grupo).first()
                    if exists:
                        exists.empresa_matriz = matriz
                        exists.empresa_filial = val_filial
                        exists.filial_is_group = filial_is_g
                    else:
                        g = ConsolidationGroup(
                            nombre_grupo=nombre_grupo, 
                            empresa_matriz=matriz, 
                            empresa_filial=val_filial,
                            filial_is_group=filial_is_g
                        )
                    db.commit()
                    
                    # Copiar plantillas default desde templates al grupo consolidado si no existen
                    grupo_folder = os.path.join(empresas_dir, f"[GRUPO] {nombre_grupo}")
                    os.makedirs(grupo_folder, exist_ok=True)
                    import shutil
                    templates_dir = "templates"
                    if os.path.exists(templates_dir):
                        for t_file in ["Balance clasificado.xlsx", "Estado de Resultados Clasificados.xlsx", "Estado de Flujos de Efectivo.xlsx"]:
                            src_file = os.path.join(templates_dir, t_file)
                            dest_file = os.path.join(grupo_folder, t_file)
                            if os.path.exists(src_file) and not os.path.exists(dest_file):
                                shutil.copy2(src_file, dest_file)
                                
                    st.success("✅ Perímetro guardado correctamente.")
                except Exception as e:
                    st.error(f"Error: {e}")
                finally:
                    db.close()
                    
        st.divider()
        st.write("Grupos Configurados:")
        try:
            db = SessionLocal()
            grupos = db.query(ConsolidationGroup).all()
            if grupos:
                # Mostrar nombre bonito si es grupo
                display_data = []
                for g in grupos:
                    f_name = g.empresa_filial
                    if g.filial_is_group:
                        sub = db.query(ConsolidationGroup).filter_by(id=int(g.empresa_filial)).first()
                        if sub:
                            f_name = f"[GRUPO] {sub.nombre_grupo}"
                    display_data.append({
                        "ID": g.id, 
                        "Nombre Grupo": g.nombre_grupo, 
                        "Matriz": g.empresa_matriz, 
                        "Filial": f_name
                    })
                st.dataframe(display_data, key="df_grupos_configurados")
            else:
                st.info("No hay grupos configurados aún.")
        finally:
            db.close()

    with tab_asientos:
        st.subheader("Ingresar Comprobantes de Ajuste y Eliminación")
        
        # Load Grupos
        db = SessionLocal()
        grupos_disp = db.query(ConsolidationGroup).all()
        db.close()
        
        if not grupos_disp:
            st.warning("Debes configurar un grupo primero en la pestaña anterior.")
        else:
            grupo_dict = {g.id: g.nombre_grupo for g in grupos_disp}
            sel_grupo = st.selectbox("Seleccionar Grupo", options=list(grupo_dict.keys()), format_func=lambda x: grupo_dict[x])
            
            # Callback functions for managing the form state safely
            if 'temp_asiento_lineas' not in st.session_state:
                st.session_state['temp_asiento_lineas'] = []

            def cb_agregar_linea():
                rubro_val = st.session_state.get("asiento_rubro")
                nota_sel = st.session_state.get("asiento_linea_nota")
                nota_val = None if (not nota_sel or str(nota_sel).startswith("--")) else str(nota_sel).strip()
                debe_val = st.session_state.get("asiento_debe", 0.0)
                haber_val = st.session_state.get("asiento_haber", 0.0)
                elimina_saldo_val = st.session_state.get("asiento_elimina_saldo", False)
                
                has_any_dynamic = elimina_saldo_val or any(l.get("elimina_saldo_total", False) for l in st.session_state.get('temp_asiento_lineas', []))
                if not has_any_dynamic and debe_val == 0.0 and haber_val == 0.0:
                    st.session_state["asiento_msg_error"] = "La línea debe tener un monto en el Debe o en el Haber si el comprobante no tiene una cuenta dinámica."
                    st.session_state["asiento_msg_success"] = None
                else:
                    lbl_nota = f" [Nota: {nota_val}]" if nota_val else ""
                    st.session_state['temp_asiento_lineas'].append({
                        "linea_item": rubro_val,
                        "linea_nota": nota_val,
                        "debe": 0.0 if elimina_saldo_val else debe_val,
                        "haber": 0.0 if elimina_saldo_val else haber_val,
                        "elimina_saldo_total": elimina_saldo_val
                    })
                    st.session_state['asiento_debe'] = 0.0
                    st.session_state['asiento_haber'] = 0.0
                    st.session_state['asiento_elimina_saldo'] = False
                    st.session_state["asiento_msg_error"] = None
                    st.session_state["asiento_msg_success"] = f"✔️ Línea agregada: {rubro_val}{lbl_nota} (Debe: {debe_val:,.0f} | Haber: {haber_val:,.0f})"
                    st.session_state['asiento_draft_version'] = st.session_state.get('asiento_draft_version', 0) + 1

            def cb_limpiar_borrador():
                st.session_state['temp_asiento_lineas'] = []
                st.session_state['asiento_glosa'] = ""
                st.session_state['asiento_debe'] = 0.0
                st.session_state['asiento_haber'] = 0.0
                st.session_state['asiento_elimina_saldo'] = False
                st.session_state['asiento_es_rec'] = False
                st.session_state["asiento_msg_error"] = None
                st.session_state["asiento_msg_success"] = "Borrador/Edición cancelada."
                st.session_state['asiento_editando_original_key'] = None
                st.session_state['sel_comprobante_activo'] = "-- Selecciona un comprobante --"
                st.session_state['asiento_draft_version'] = st.session_state.get('asiento_draft_version', 0) + 1

            def cb_toggle_dinamico_todas():
                val = st.session_state.get("asiento_elimina_saldo", False)
                for l in st.session_state.get('temp_asiento_lineas', []):
                    l['elimina_saldo_total'] = val
                st.session_state['asiento_draft_version'] = st.session_state.get('asiento_draft_version', 0) + 1

            def cb_guardar_asiento(sel_grupo, periodo_a, col_ajuste):
                glosa_val = st.session_state.get("asiento_glosa", "").strip()
                es_rec_val = st.session_state.get("asiento_es_rec", False)
                draft_lines = st.session_state.get("temp_asiento_lineas", [])
                
                if not glosa_val:
                    st.session_state["asiento_msg_error"] = "La glosa explicativa es obligatoria para registrar el asiento."
                    st.session_state["asiento_msg_success"] = None
                    return
                    
                tot_d = sum(l["debe"] for l in draft_lines)
                tot_h = sum(l["haber"] for l in draft_lines)
                diff = tot_d - tot_h
                has_dynamic = any(l["elimina_saldo_total"] for l in draft_lines)
                
                if not has_dynamic and abs(diff) > 0.01:
                    st.session_state["asiento_msg_error"] = f"❌ No se puede guardar: El asiento no está cuadrado. La diferencia debe ser cero (Diferencia actual: {diff:,.0f})."
                    st.session_state["asiento_msg_success"] = None
                    return
                    
                db_save = SessionLocal()
                try:
                    from src.core.consolidacion_engine import generar_siguiente_codigo_asiento

                    # Si estamos editando, obtener el asiento_codigo original si existe, o generar uno nuevo
                    edit_key = st.session_state.get('asiento_editando_original_key')
                    codigo_voucher = None
                    if edit_key:
                        orig_per, orig_col, orig_glosa = edit_key
                        orig_entry = db_save.query(ConsolidationJournalEntry).filter_by(
                            grupo_id=sel_grupo,
                            periodo=orig_per,
                            columna_ajuste=orig_col,
                            glosa=orig_glosa
                        ).first()
                        orig_code = getattr(orig_entry, 'asiento_codigo', None) if orig_entry else None
                        if orig_code:
                            codigo_voucher = orig_code
                            
                        db_save.query(ConsolidationJournalEntry).filter_by(
                            grupo_id=sel_grupo,
                            periodo=orig_per,
                            columna_ajuste=orig_col,
                            glosa=orig_glosa
                        ).delete()
                        
                    if not codigo_voucher:
                        codigo_voucher = generar_siguiente_codigo_asiento(sel_grupo, periodo_a, db_save)

                    user_curr = st.session_state.get('user_name', 'usuario_sistema')
                    for idx, l in enumerate(draft_lines, start=1):
                        a = ConsolidationJournalEntry(
                            grupo_id=sel_grupo,
                            periodo=periodo_a,
                            glosa=glosa_val,
                            columna_ajuste=col_ajuste,
                            linea_item=l["linea_item"],
                            linea_nota=l.get("linea_nota"),
                            debe=l["debe"],
                            haber=l["haber"],
                            es_recurrente=es_rec_val,
                            elimina_saldo_total=l["elimina_saldo_total"],
                            asiento_codigo=codigo_voucher,
                            num_linea=idx,
                            created_by=user_curr,
                            updated_by=user_curr
                        )
                        db_save.add(a)
                    db_save.commit()
                    
                    if edit_key:
                        st.session_state["asiento_msg_success"] = f"✅ Cambios del asiento [{codigo_voucher}] '{glosa_val}' guardados con éxito."
                    else:
                        st.session_state["asiento_msg_success"] = f"✅ Asiento contable [{codigo_voucher}] guardado y posteado con éxito ({len(draft_lines)} líneas)."
                        
                    # Limpiar caché de contextos de entidades para forzar recarga en la pestaña de Notas
                    for k_ctx in list(st.session_state.keys()):
                        if k_ctx.startswith("_entity_ctx__"):
                            del st.session_state[k_ctx]

                    st.session_state["asiento_msg_error"] = None
                    st.session_state['temp_asiento_lineas'] = []
                    st.session_state['asiento_glosa'] = ""
                    st.session_state['asiento_debe'] = 0.0
                    st.session_state['asiento_haber'] = 0.0
                    st.session_state['asiento_elimina_saldo'] = False
                    st.session_state['asiento_es_rec'] = False
                    st.session_state['asiento_editando_original_key'] = None
                    st.session_state['sel_comprobante_activo'] = "-- Selecciona un comprobante --"
                    st.session_state['asiento_draft_version'] = st.session_state.get('asiento_draft_version', 0) + 1
                except Exception as e:
                    st.session_state["asiento_msg_error"] = f"Error: {e}"
                    st.session_state["asiento_msg_success"] = None
                finally:
                    db_save.close()

            def cb_cargar_para_editar(periodo, columna, glosa, es_rec, lines):
                st.session_state['temp_asiento_lineas'] = [
                    {
                        "linea_item": l.linea_item,
                        "linea_nota": getattr(l, 'linea_nota', None),
                        "debe": float(l.debe or 0.0),
                        "haber": float(l.haber or 0.0),
                        "elimina_saldo_total": bool(l.elimina_saldo_total)
                    } for l in lines
                ]
                st.session_state['asiento_glosa'] = glosa
                st.session_state['asiento_col_ajuste'] = columna
                st.session_state['per_asiento'] = periodo
                st.session_state['asiento_es_rec'] = bool(es_rec)
                st.session_state['asiento_elimina_saldo'] = any(bool(l.elimina_saldo_total) for l in lines)
                st.session_state['asiento_editando_original_key'] = (periodo, columna, glosa)
                st.session_state['asiento_msg_success'] = f"✏️ Asiento '{glosa}' cargado en el formulario de arriba para editar. Modifica lo que necesites y haz clic en 'Guardar Cambios del Asiento'."
                st.session_state['edit_msg_success'] = f"✏️ **Asiento cargado para editar**: Desplázate hacia arriba 👆 al formulario superior ('Editar Asiento Contable') para modificar los números."
                st.session_state['asiento_msg_error'] = None
                st.session_state['asiento_draft_version'] = st.session_state.get('asiento_draft_version', 0) + 1

            def cb_copiar_asiento(periodo, columna, glosa, es_rec, lines):
                st.session_state['temp_asiento_lineas'] = [
                    {
                        "linea_item": l.linea_item,
                        "linea_nota": getattr(l, 'linea_nota', None),
                        "debe": float(l.debe or 0.0),
                        "haber": float(l.haber or 0.0),
                        "elimina_saldo_total": bool(l.elimina_saldo_total)
                    } for l in lines
                ]
                st.session_state['asiento_glosa'] = f"Copia de {glosa}"
                st.session_state['asiento_col_ajuste'] = columna
                st.session_state['asiento_es_rec'] = bool(es_rec)
                st.session_state['asiento_editando_original_key'] = None
                st.session_state['asiento_msg_success'] = f"📋 Estructura de '{glosa}' copiada al borrador superior. Cambia el periodo/mes, modifica los montos y haz clic en 'Guardar Asiento Completo' para guardarlo como un nuevo registro."
                st.session_state['edit_msg_success'] = f"📋 **Estructura copiada**: Desplázate hacia arriba 👆 al formulario superior ('Nuevo Asiento Contable') para cambiar el mes, cambiar números y guardar."
                st.session_state['asiento_msg_error'] = None
                st.session_state['asiento_draft_version'] = st.session_state.get('asiento_draft_version', 0) + 1
                st.session_state['asiento_glosa'] = f"Copia de {glosa}"
                st.session_state['asiento_col_ajuste'] = columna
                st.session_state['asiento_es_rec'] = bool(es_rec)
                st.session_state['asiento_editando_original_key'] = None
                st.session_state['asiento_msg_success'] = f"📋 Estructura de '{glosa}' copiada al borrador superior. Cambia el periodo/mes, modifica los montos y haz clic en 'Guardar Asiento Completo' para guardarlo como un nuevo registro."
                st.session_state['edit_msg_success'] = f"📋 **Estructura copiada**: Desplázate hacia arriba 👆 al formulario superior ('Nuevo Asiento Contable') para cambiar el mes, cambiar números y guardar."
                st.session_state['asiento_msg_error'] = None
                st.session_state['asiento_draft_version'] = st.session_state.get('asiento_draft_version', 0) + 1

            def cb_eliminar_asiento(periodo, columna, glosa):
                db_del = SessionLocal()
                try:
                    db_del.query(ConsolidationJournalEntry).filter_by(
                        grupo_id=sel_grupo,
                        periodo=periodo,
                        columna_ajuste=columna,
                        glosa=glosa
                    ).delete()
                    db_del.commit()
                    st.session_state["asiento_msg_success"] = f"🗑️ Asiento '{glosa}' eliminado por completo."
                    st.session_state["asiento_msg_error"] = None
                    
                    # Si se estaba editando este asiento, limpiar el borrador
                    if st.session_state.get('asiento_editando_original_key') == (periodo, columna, glosa):
                        st.session_state['temp_asiento_lineas'] = []
                        st.session_state['asiento_glosa'] = ""
                        st.session_state['asiento_debe'] = 0.0
                        st.session_state['asiento_haber'] = 0.0
                        st.session_state['asiento_elimina_saldo'] = False
                        st.session_state['asiento_es_rec'] = False
                        st.session_state['asiento_editando_original_key'] = None
                    st.session_state['sel_comprobante_activo'] = "-- Selecciona un comprobante --"
                    st.session_state['asiento_draft_version'] = st.session_state.get('asiento_draft_version', 0) + 1
                except Exception as e:
                    st.session_state["asiento_msg_error"] = f"Error al eliminar: {e}"
                    st.session_state["asiento_msg_success"] = None
                finally:
                    db_del.close()

            edit_key = st.session_state.get('asiento_editando_original_key')
            expander_title = "✏️ Editando Asiento Contable" if edit_key else "➕ Nuevo Asiento Contable de Ajuste"
            with st.expander(expander_title, expanded=True):
                # Display validation messages from callback actions
                if st.session_state.get("asiento_msg_error"):
                    st.error(st.session_state["asiento_msg_error"])
                if st.session_state.get("asiento_msg_success"):
                    st.success(st.session_state["asiento_msg_success"])

                # Obtener periodos disponibles en históricos
                db = SessionLocal()
                per_recs = db.query(HistoricalDataRecord.periodo).distinct().all()
                db.close()
                periodos_hist = sorted([r[0] for r in per_recs], reverse=True)
                if not periodos_hist: periodos_hist = ["2026-12", "2026-11", "2026-10", "2026-09", "2026-08", "2026-07", "2026-06", "2026-05", "2026-04", "2026-03", "2026-02", "2026-01", "2025-12"]
                
                # Asegurar que el periodo que estamos editando esté en las opciones
                if edit_key and edit_key[0] not in periodos_hist:
                    periodos_hist.append(edit_key[0])
                    periodos_hist = sorted(periodos_hist, reverse=True)
                
                columnas_destino = ["Elim inversión", "Elim Ctas IC", "reversa reclas Plusvalia", "PPA", "Amortizaciones", "Reclasificaciones", "Otras Eliminaciones"]
                # Asegurar que la columna que estamos editando esté en las opciones
                if edit_key and edit_key[1] not in columnas_destino:
                    columnas_destino.append(edit_key[1])
                
                col_p, col_f, col_c = st.columns(3)
                periodo_a = col_p.selectbox("Periodo", periodos_hist, key="per_asiento", format_func=format_periodo)
                col_ajuste = col_c.selectbox("Columna Destino en Hoja de Trabajo", columnas_destino, key="asiento_col_ajuste")
                glosa = st.text_input("Glosa Explicativa del Ajuste (Detalle)", key="asiento_glosa")
                
                # Fetch line items for the Rubro selector from the group's taxonomy and history
                lineas = []
                try:
                    db = SessionLocal()
                    from src.models.taxonomy_master import TaxonomyMasterRecord
                    grupo_obj = db.query(ConsolidationGroup).filter_by(id=sel_grupo).first()
                    if grupo_obj:
                        entidades_grupo = [grupo_obj.empresa_matriz]
                        if grupo_obj.filial_is_group:
                            def get_subgroup_companies(sub_g_id):
                                comps = []
                                sub_g = db.query(ConsolidationGroup).filter_by(id=sub_g_id).first()
                                if sub_g:
                                    comps.append(sub_g.empresa_matriz)
                                    if sub_g.filial_is_group:
                                        comps.extend(get_subgroup_companies(int(sub_g.empresa_filial)))
                                    else:
                                        comps.append(sub_g.empresa_filial)
                                return comps
                            entidades_grupo.extend(get_subgroup_companies(int(grupo_obj.empresa_filial)))
                        else:
                            entidades_grupo.append(grupo_obj.empresa_filial)
                            
                        # Query unique line names from taxonomy and history of these companies
                        tax_items = db.query(TaxonomyMasterRecord.nombre_linea_es).filter(
                            TaxonomyMasterRecord.empresa.in_(entidades_grupo),
                            TaxonomyMasterRecord.reporte_destino.in_(['Balance', 'P&L'])
                        ).distinct().all()
                        
                        hist_items = db.query(HistoricalDataRecord.linea_item).filter(
                            HistoricalDataRecord.empresa.in_(entidades_grupo)
                        ).distinct().all()
                        
                        lineas_set = set(r[0] for r in tax_items).union(set(r[0] for r in hist_items))
                        lineas_set.add("Ganancias (Pérdida) del Ejercicio")
                        lineas = sorted(list(lineas_set))
                    else:
                        tax_recs = db.query(HistoricalDataRecord.linea_item).filter_by(empresa=empresa_seleccionada).distinct().all()
                        lineas_set = set(r[0] for r in tax_recs)
                        lineas_set.add("Ganancias (Pérdida) del Ejercicio")
                        lineas = sorted(list(lineas_set))
                except Exception as e:
                    st.error(f"Error cargando rubros: {e}")
                finally:
                    db.close()
                if not lineas: lineas = ["Activos", "Pasivos", "Patrimonio", "Ingresos", "Gastos"]
                
                col_r1, col_r2 = st.columns(2)
                with col_r1:
                    rubro = st.selectbox("Rubro (Línea de Balance/P&L) a afectar", lineas, key="asiento_rubro")
                
                opciones_nota = get_all_note_options(rubro)
                with col_r2:
                    st.selectbox("Detalle / Nota Afectada (Opcional)", opciones_nota, key="asiento_linea_nota")

                col_d, col_h = st.columns(2)
                elimina_saldo = st.checkbox("🔮 Eliminar saldo total automáticamente (Dinámico - calcula la reversa según balance mensual)", key="asiento_elimina_saldo", on_change=cb_toggle_dinamico_todas)
                
                debe = col_d.number_input("Debe (Monto +)", min_value=0.0, step=1000.0, disabled=elimina_saldo, key="asiento_debe")
                haber = col_h.number_input("Haber (Monto -)", min_value=0.0, step=1000.0, disabled=elimina_saldo, key="asiento_haber")
                
                es_rec = st.checkbox("🔄 Es Recurrente (Repetir automáticamente este asiento en consolidaciones de meses futuros)", key="asiento_es_rec")
                
                st.button("➕ Agregar Línea al Borrador", type="secondary", use_container_width=True, on_click=cb_agregar_linea)
                
                # Mostrar Borrador si tiene líneas
                if st.session_state['temp_asiento_lineas']:
                    st.divider()
                    st.write("📋 **Borrador de Asiento Actual (Editable)**")
                    st.caption("💡 Puedes hacer doble clic en cualquier celda para corregir los montos o cambiar de rubro antes de guardar. También puedes seleccionar filas y presionar 'Supr' (Delete) para eliminarlas.")
                    
                    # Calcular saldos y ajustes resueltos en tiempo real para el borrador
                    from src.core.consolidacion_engine import resolver_montos_asiento
                    try:
                        resolved_draft = resolver_montos_asiento(sel_grupo, periodo_a, st.session_state['temp_asiento_lineas'], columna_destino=col_ajuste)
                    except Exception as e:
                        # Si no hay saldos aún o periodo inválido, usar valores manuales como fallback
                        resolved_draft = [{
                            "linea_item": l["linea_item"],
                            "debe_calculado": l["debe"],
                            "haber_calculado": l["haber"],
                            "elimina_saldo_total": l["elimina_saldo_total"],
                            "saldo_base": 0.0
                        } for l in st.session_state['temp_asiento_lineas']]
                    
                    df_draft_raw = pd.DataFrame([{
                        "Rubro": l["linea_item"],
                        "Nota Afectada": l.get("linea_nota") or "Sin Detalle",
                        "Saldo Base": f"{int(round(resolved_draft[i]['saldo_base'])):,}".replace(",", "."),
                        "Debe (Ajuste Real)": f"{int(round(resolved_draft[i]['debe_calculado'])):,}".replace(",", "."),
                        "Haber (Ajuste Real)": f"{int(round(resolved_draft[i]['haber_calculado'])):,}".replace(",", "."),
                        "Debe (Manual)": f"{int(round(l['debe'] or 0.0)):,}".replace(",", "."),
                        "Haber (Manual)": f"{int(round(l['haber'] or 0.0)):,}".replace(",", "."),
                        "Dinámico": bool(l["elimina_saldo_total"])
                    } for i, l in enumerate(st.session_state['temp_asiento_lineas'])])
                    
                    # Dynamic key to reset editor state when group, period, or external changes occur
                    draft_editor_version = st.session_state.get('asiento_draft_version', 0)
                    draft_editor_key = f"asiento_draft_editor_{sel_grupo}_{periodo_a}_{draft_editor_version}"

                    edited_draft_df = st.data_editor(
                        df_draft_raw,
                        key=draft_editor_key,
                        num_rows="dynamic",
                        disabled=["Saldo Base", "Debe (Ajuste Real)", "Haber (Ajuste Real)"],
                        use_container_width=True,
                        column_config={
                            "Rubro": st.column_config.SelectboxColumn(
                                "Rubro",
                                help="Rubro a afectar",
                                width="medium",
                                options=lineas,
                                required=True
                            ),
                            "Saldo Base": st.column_config.TextColumn(
                                "Saldo Base"
                            ),
                            "Debe (Ajuste Real)": st.column_config.TextColumn(
                                "Debe (Ajuste Real)"
                            ),
                            "Haber (Ajuste Real)": st.column_config.TextColumn(
                                "Haber (Ajuste Real)"
                            ),
                            "Debe (Manual)": st.column_config.TextColumn(
                                "Debe (Manual)"
                            ),
                            "Haber (Manual)": st.column_config.TextColumn(
                                "Haber (Manual)"
                            ),
                            "Dinámico": st.column_config.CheckboxColumn(
                                "Dinámico",
                                help="Eliminar saldo total automáticamente (calcula la reversa según balance)",
                                default=False
                            )
                        }
                    )
                    
                    # Actualizar st.session_state con las líneas editadas en tiempo real
                    new_lines = []
                    for _, row in edited_draft_df.iterrows():
                        rubro_val = str(row.get("Rubro", "")).strip()
                        nota_raw  = str(row.get("Nota Afectada", "")).strip()
                        nota_val  = None if (not nota_raw or nota_raw in ("Sin Detalle", "-- Sin Detalle / General --", "")) else nota_raw
                        if rubro_val:
                            debe_str = str(row.get("Debe (Manual)", "0")).replace(".", "").replace(",", "")
                            haber_str = str(row.get("Haber (Manual)", "0")).replace(".", "").replace(",", "")
                            new_lines.append({
                                "linea_item": rubro_val,
                                "linea_nota": nota_val,
                                "debe": float(pd.to_numeric(debe_str, errors='coerce') or 0.0),
                                "haber": float(pd.to_numeric(haber_str, errors='coerce') or 0.0),
                                "elimina_saldo_total": bool(row.get("Dinámico", False))
                            })
                    
                    if new_lines != st.session_state['temp_asiento_lineas']:
                        st.session_state['temp_asiento_lineas'] = new_lines
                        st.rerun()
                    
                    # Recalcular totales reales sobre el borrador resuelto
                    tot_d = sum(r["debe_calculado"] for r in resolved_draft)
                    tot_h = sum(r["haber_calculado"] for r in resolved_draft)
                    diff = tot_d - tot_h
                    
                    col_t1, col_t2 = st.columns(2)
                    with col_t1:
                        st.info(f"**Totales Resueltos (Incluye Dinámicos)**:\n* Debe: {tot_d:,.0f}\n* Haber: {tot_h:,.0f}\n* Diferencia: {diff:,.0f}")
                    with col_t2:
                        if abs(diff) > 0.01:
                            st.warning("⚠️ El asiento resuelto no está cuadrado.")
                        else:
                            st.success("✅ Asiento cuadrado (incluyendo saldos dinámicos).")
                    
                    is_editing = st.session_state.get('asiento_editando_original_key') is not None
                    btn_save_lbl = "💾 Guardar Cambios del Asiento" if is_editing else "💾 Guardar Asiento Completo"
                    btn_cancel_lbl = "❌ Cancelar Edición" if is_editing else "🗑️ Limpiar Borrador"
                    
                    col_btn1, col_btn2 = st.columns(2)
                    col_btn1.button(btn_save_lbl, type="primary", use_container_width=True, on_click=cb_guardar_asiento, args=(sel_grupo, periodo_a, col_ajuste))
                    col_btn2.button(btn_cancel_lbl, type="secondary", use_container_width=True, on_click=cb_limpiar_borrador)
            
            st.divider()
            st.write("### 📖 Libro de Comprobantes de Ajuste (Resumen)")
            try:
                db = SessionLocal()
                # Consultar asientos agrupados por comprobante (periodo, columna, glosa)
                from sqlalchemy import func
                asientos_resumen = db.query(
                    ConsolidationJournalEntry.asiento_codigo,
                    ConsolidationJournalEntry.periodo,
                    ConsolidationJournalEntry.columna_ajuste,
                    ConsolidationJournalEntry.glosa,
                    func.count(ConsolidationJournalEntry.id).label("lineas"),
                    func.sum(ConsolidationJournalEntry.debe).label("total_debe"),
                    func.sum(ConsolidationJournalEntry.haber).label("total_haber"),
                    func.max(ConsolidationJournalEntry.es_recurrente).label("es_recurrente")
                ).filter_by(grupo_id=sel_grupo).group_by(
                    ConsolidationJournalEntry.asiento_codigo,
                    ConsolidationJournalEntry.periodo,
                    ConsolidationJournalEntry.columna_ajuste,
                    ConsolidationJournalEntry.glosa
                ).order_by(
                    ConsolidationJournalEntry.periodo.desc(),
                    ConsolidationJournalEntry.asiento_codigo.desc()
                ).all()
                
                if asientos_resumen:
                    # --- BARRA DE FILTROS SUPERIOR ---
                    st.write("#### 🔍 Filtros y Búsqueda de Comprobantes")
                    col_f1, col_f2, col_f3, col_f4 = st.columns([2, 2, 3, 1.5])
                    
                    periodos_unicos = sorted(list(set(r.periodo for r in asientos_resumen)), reverse=True)
                    opciones_periodo = ["Todos los períodos"]
                    if periodo_a in periodos_unicos:
                        opciones_periodo.append(f"Período Activo ({format_periodo(periodo_a)})")
                    opciones_periodo += [format_periodo(p) for p in periodos_unicos]
                    
                    sel_filtro_periodo = col_f1.selectbox(
                        "📅 Período", 
                        options=opciones_periodo,
                        key=f"f_periodo_{sel_grupo}"
                    )
                    
                    columnas_unicas = sorted(list(set(r.columna_ajuste for r in asientos_resumen)))
                    sel_filtro_columna = col_f2.selectbox(
                        "🏷️ Tipo / Columna", 
                        options=["Todas"] + columnas_unicas,
                        key=f"f_columna_{sel_grupo}"
                    )
                    
                    txt_busqueda = col_f3.text_input(
                        "🔎 Buscar en Código o Glosa", 
                        placeholder="Ej: AST-202605-001, Servicios, PPA...",
                        key=f"f_glosa_{sel_grupo}"
                    ).strip().lower()
                    
                    solo_recurrentes = col_f4.checkbox("🔄 Recurrentes", key=f"f_rec_{sel_grupo}")
                    
                    # Filtrar lista de asientos_resumen según los controles
                    asientos_filtrados = []
                    for r in asientos_resumen:
                        # Filtro Período
                        if sel_filtro_periodo.startswith("Período Activo"):
                            if r.periodo != periodo_a:
                                continue
                        elif sel_filtro_periodo != "Todos los períodos":
                            if format_periodo(r.periodo) != sel_filtro_periodo and r.periodo != sel_filtro_periodo:
                                continue
                        
                        # Filtro Columna
                        if sel_filtro_columna != "Todas" and r.columna_ajuste != sel_filtro_columna:
                            continue
                        
                        # Filtro Glosa / Código / Explicación
                        if txt_busqueda:
                            code_val = getattr(r, 'asiento_codigo', None)
                            code_match = code_val and txt_busqueda in code_val.lower()
                            glosa_match = txt_busqueda in r.glosa.lower()
                            if not (code_match or glosa_match):
                                continue
                        
                        # Filtro Recurrente
                        if solo_recurrentes and not r.es_recurrente:
                            continue
                            
                        asientos_filtrados.append(r)
                    
                    if not asientos_filtrados:
                        st.info("ℹ️ No se encontraron comprobantes que coincidan con los filtros aplicados.")
                    else:
                        # Agrupar por período para desplegar acordeones
                        asientos_por_periodo = {}
                        for r in asientos_filtrados:
                            asientos_por_periodo.setdefault(r.periodo, []).append(r)
                            
                        st.caption(f"Mostrando **{len(asientos_filtrados)}** de **{len(asientos_resumen)}** comprobante(s) registrado(s).")
                        
                        for p_key, p_asientos in asientos_por_periodo.items():
                            p_fmt = format_periodo(p_key)
                            tot_debe_mes = sum((r.total_debe or 0.0) for r in p_asientos)
                            tot_haber_mes = sum((r.total_haber or 0.0) for r in p_asientos)
                            lbl_debe = f"{int(round(tot_debe_mes)):,}".replace(",", ".")
                            lbl_haber = f"{int(round(tot_haber_mes)):,}".replace(",", ".")
                            
                            # Expandir si coincide con el periodo activo o si se filtró específicamente ese período
                            default_expanded = (p_key == periodo_a) or (sel_filtro_periodo != "Todos los períodos")
                            
                            with st.expander(
                                f"📅 **{p_fmt}** — {len(p_asientos)} comprobante(s) | Total Debe: ${lbl_debe} | Total Haber: ${lbl_haber}", 
                                expanded=default_expanded
                            ):
                                df_resumen_p = pd.DataFrame([{
                                    "Código Folio": getattr(r, 'asiento_codigo', "N/A") or "N/A",
                                    "Columna de Ajuste": r.columna_ajuste,
                                    "Glosa / Explicación": r.glosa,
                                    "N° Líneas": r.lineas,
                                    "Total Debe": f"{int(round(r.total_debe or 0.0)):,}".replace(",", "."),
                                    "Total Haber": f"{int(round(r.total_haber or 0.0)):,}".replace(",", "."),
                                    "Recurrente": "Sí" if r.es_recurrente else "No"
                                } for r in p_asientos])
                                
                                st.dataframe(
                                    df_resumen_p,
                                    use_container_width=True,
                                    column_config={
                                        "Código Folio": st.column_config.TextColumn("Código Folio"),
                                        "Columna de Ajuste": st.column_config.TextColumn("Columna de Ajuste"),
                                        "Glosa / Explicación": st.column_config.TextColumn("Glosa / Explicación"),
                                        "N° Líneas": st.column_config.NumberColumn("N° Líneas", format="%d"),
                                        "Total Debe": st.column_config.TextColumn("Total Debe"),
                                        "Total Haber": st.column_config.TextColumn("Total Haber"),
                                        "Recurrente": st.column_config.TextColumn("Recurrente")
                                    },
                                    key=f"df_resumen_v3_{sel_grupo}_{p_key}_{len(p_asientos)}"
                                )
                                
                                st.write("---")
                                st.write(f"🔍 **Consultar / Editar / Eliminar Comprobante de {p_fmt}:**")
                                
                                options_comp_p = [f"[{getattr(r, 'asiento_codigo', 'N/A') or 'N/A'}] {r.columna_ajuste} - {r.glosa}" for r in p_asientos]
                                sel_comp_p = st.selectbox(
                                    f"Selecciona un comprobante de {p_fmt} para trabajar",
                                    options=["-- Selecciona un comprobante --"] + options_comp_p,
                                    key=f"sel_comp_p_{sel_grupo}_{p_key}"
                                )
                                
                                if sel_comp_p != "-- Selecciona un comprobante --":
                                    idx_c = options_comp_p.index(sel_comp_p)
                                    chosen = p_asientos[idx_c]
                                    codigo_actual = getattr(chosen, 'asiento_codigo', 'N/A') or "N/A"
                                    
                                    st.info(f"📌 **Comprobante Activo:** `{codigo_actual}` | **Tipo:** {chosen.columna_ajuste} | **Glosa:** {chosen.glosa}")
                                    
                                    # Informar visualmente si este comprobante se está editando actualmente
                                    edit_key = st.session_state.get('asiento_editando_original_key')
                                    if edit_key and edit_key == (chosen.periodo, chosen.columna_ajuste, chosen.glosa):
                                        st.info("✏️ **Este comprobante ya está cargado en el formulario superior para editar.**\nPor favor, desplázate hacia arriba 👆 en esta pestaña para modificar los datos y guardar.")
                                    
                                    # Cargar las líneas detalladas de este comprobante
                                    lineas_comp = db.query(ConsolidationJournalEntry).filter_by(
                                        grupo_id=sel_grupo,
                                        periodo=chosen.periodo,
                                        columna_ajuste=chosen.columna_ajuste,
                                        glosa=chosen.glosa
                                    ).all()
                                    
                                    st.write(f"**Previsualización de Líneas del Comprobante `{codigo_actual}`:**")
                                    from src.core.consolidacion_engine import resolver_montos_asiento
                                    try:
                                        resolved_preview = resolver_montos_asiento(sel_grupo, chosen.periodo, lineas_comp, columna_destino=chosen.columna_ajuste)
                                        df_preview = pd.DataFrame([{
                                            "Línea #": getattr(l, 'num_linea', i + 1),
                                            "Código Folio": getattr(l, 'asiento_codigo', 'N/A') or "N/A",
                                            "Rubro": l.linea_item,
                                            "Nota Afectada": getattr(l, 'linea_nota', None) or "Sin Detalle",
                                            "Saldo Base": f"{int(round(resolved_preview[i]['saldo_base'])):,}".replace(",", "."),
                                            "Debe (Ajuste Real)": f"{int(round(resolved_preview[i]['debe_calculado'])):,}".replace(",", "."),
                                            "Haber (Ajuste Real)": f"{int(round(resolved_preview[i]['haber_calculado'])):,}".replace(",", "."),
                                            "Dinámico (Elimina Saldo)": "Sí" if l.elimina_saldo_total else "No"
                                        } for i, l in enumerate(lineas_comp)])
                                    except Exception as e:
                                        df_preview = pd.DataFrame([{
                                            "Línea #": getattr(a, 'num_linea', i + 1),
                                            "Código Folio": getattr(a, 'asiento_codigo', 'N/A') or "N/A",
                                            "Rubro": a.linea_item,
                                            "Nota Afectada": getattr(a, 'linea_nota', None) or "Sin Detalle",
                                            "Saldo Base": "0",
                                            "Debe (Ajuste Real)": f"{int(round(a.debe or 0.0)):,}".replace(",", "."),
                                            "Haber (Ajuste Real)": f"{int(round(a.haber or 0.0)):,}".replace(",", "."),
                                            "Dinámico (Elimina Saldo)": "Sí" if a.elimina_saldo_total else "No"
                                        } for i, a in enumerate(lineas_comp)])
                                    
                                    st.dataframe(
                                        df_preview,
                                        use_container_width=True,
                                        column_config={
                                            "Línea #": st.column_config.NumberColumn("Línea #", format="%d"),
                                            "Código Folio": st.column_config.TextColumn("Código Folio"),
                                            "Rubro": st.column_config.TextColumn("Rubro"),
                                            "Saldo Base": st.column_config.TextColumn("Saldo Base"),
                                            "Debe (Ajuste Real)": st.column_config.TextColumn("Debe (Ajuste Real)"),
                                            "Haber (Ajuste Real)": st.column_config.TextColumn("Haber (Ajuste Real)"),
                                            "Dinámico (Elimina Saldo)": st.column_config.TextColumn("Dinámico (Elimina Saldo)")
                                        },
                                        key=f"df_preview_v3_{sel_grupo}_{p_key}_{chosen.columna_ajuste}_{idx_c}"
                                    )
                                    
                                    col_act1, col_act2, col_act3 = st.columns(3)
                                    
                                    col_act1.button(
                                        "✏️ Cargar para Editar", 
                                        type="primary", 
                                        use_container_width=True, 
                                        on_click=cb_cargar_para_editar, 
                                        args=(chosen.periodo, chosen.columna_ajuste, chosen.glosa, chosen.es_recurrente, lineas_comp),
                                        key=f"btn_edit_{p_key}_{idx_c}"
                                    )
                                    
                                    col_act2.button(
                                        "📋 Copiar al Borrador", 
                                        type="secondary", 
                                        use_container_width=True, 
                                        on_click=cb_copiar_asiento, 
                                        args=(chosen.periodo, chosen.columna_ajuste, chosen.glosa, chosen.es_recurrente, lineas_comp),
                                        key=f"btn_copy_{p_key}_{idx_c}"
                                    )
                                    
                                    col_act3.button(
                                        "🗑️ Eliminar Asiento", 
                                        type="secondary", 
                                        use_container_width=True, 
                                        on_click=cb_eliminar_asiento, 
                                        args=(chosen.periodo, chosen.columna_ajuste, chosen.glosa),
                                        key=f"btn_del_{p_key}_{idx_c}"
                                    )
                else:
                    if "last_sel_comp" in st.session_state:
                        del st.session_state["last_sel_comp"]
                    st.info("No hay asientos registrados para este grupo.")
            finally:
                db.close()


    with tab_concil:
        st.subheader("Conciliación y Cuadratura Intercompany")
        st.write("Verifica la simetría de saldos por cobrar/pagar e ingresos/gastos intercompañía antes de consolidar.")
        
        # Load Grupos
        db = SessionLocal()
        grupos_disp_c = db.query(ConsolidationGroup).all()
        db.close()
        
        if not grupos_disp_c:
            st.warning("Debes configurar un grupo de consolidación primero en la pestaña '⚙️ Configurar Perímetro'.")
        else:
            grupo_dict_c = {g.id: g.nombre_grupo for g in grupos_disp_c}
            col_gc, col_pc = st.columns(2)
            sel_g_c = col_gc.selectbox("Seleccionar Grupo a Conciliar", options=list(grupo_dict_c.keys()), format_func=lambda x: grupo_dict_c[x], key="sel_g_c")
            
            db = SessionLocal()
            per_recs_c = db.query(HistoricalDataRecord.periodo).distinct().all()
            db.close()
            periodos_hist_c = sorted([r[0] for r in per_recs_c], reverse=True)
            if not periodos_hist_c: periodos_hist_c = ["2026-12", "2025-12"]
            
            periodo_concil = col_pc.selectbox("Periodo a Conciliar", periodos_hist_c, key="per_concil", format_func=format_periodo)
            
            custom_kws = st.text_input("Palabras clave adicionales (ej: fibra, servicio - separadas por coma)", key="custom_kws")
            
            if st.button("Ejecutar Conciliación", type="primary", key="btn_ejecutar_concil"):
                with st.spinner("Buscando saldos intercompañía y calculando diferencias..."):
                    # 1. Recuperar info del grupo
                    db = SessionLocal()
                    try:
                        grupo_obj = db.query(ConsolidationGroup).filter_by(id=sel_g_c).first()
                        if not grupo_obj:
                            st.error("Grupo no encontrado.")
                            return
                        
                        matriz_name = grupo_obj.empresa_matriz
                        filial_name = grupo_obj.empresa_filial
                        if grupo_obj.filial_is_group:
                            sub_g_obj = db.query(ConsolidationGroup).filter_by(id=int(grupo_obj.empresa_filial)).first()
                            if sub_g_obj:
                                filial_name = f"Consolidado {sub_g_obj.nombre_grupo}"
                                
                        from src.core.consolidacion_engine import obtener_saldos_base, is_pl_account
                        base_data = obtener_saldos_base(grupo_obj.id, periodo_concil, db)
                    finally:
                        db.close()
                        
                    if not base_data:
                        st.warning("⚠️ No se encontraron saldos base (Matriz ni Filial) para el periodo seleccionado.")
                    else:
                        # 2. Configurar palabras clave de búsqueda (excluyendo "ic" que se busca exacto)
                        KEYWORDS_IC = ["relacionada", "relacionadas", "filial", "filiales", "matriz", "holding", "intercompany"]
                        if custom_kws.strip():
                            extra_kws = [k.strip().lower() for k in custom_kws.split(",") if k.strip()]
                            KEYWORDS_IC.extend(extra_kws)
                            
                        # Palabras clave del nombre de las compañías
                        def extract_company_words(name):
                            words = []
                            clean = name.lower().replace("spa", "").replace("sa", "").replace("individual", "").replace("consolidado", "").replace("sp", "")
                            for w in clean.split():
                                w_clean = "".join(c for c in w if c.isalnum())
                                if len(w_clean) > 2:
                                    words.append(w_clean)
                            return words
                            
                        company_keywords = set(extract_company_words(matriz_name) + extract_company_words(filial_name))
                        
                        # Detección
                        def is_ic_line(line_name):
                            norm = line_name.lower().strip().replace('á','a').replace('é','e').replace('í','i').replace('ó','o').replace('ú','u')
                            words = [w.strip(".,()/-") for w in norm.split()]
                            
                            if "ic" in words or "i/c" in words:
                                return True
                                
                            for kw in KEYWORDS_IC:
                                if kw in norm:
                                    return True
                                    
                            for kw in company_keywords:
                                if len(kw) > 3 and kw in norm:
                                    return True
                                    
                            return False
                            
                        # Clasificación
                        def classify_ic_line(line_name):
                            norm = line_name.lower().strip().replace('á','a').replace('é','e').replace('í','i').replace('ó','o').replace('ú','u')
                            is_pl = is_pl_account(line_name)
                            
                            if is_pl:
                                if any(x in norm for x in ["ingreso", "venta", "arriendo", "prestacion", "ganancia"]):
                                    return "Ingresos IC"
                                else:
                                    return "Gastos IC"
                            else:
                                if "inversion" in norm:
                                    return "Inversiones IC"
                                if any(x in norm for x in ["cobrar", "cxc", "deudor", "deudores", "activo", "anticipo", "prestamo", "mutuo"]):
                                    return "Activos IC"
                                else:
                                    return "Pasivos IC"
                                    
                        # Procesar
                        matched_rows = []
                        for line, saldos in base_data.items():
                            if is_ic_line(line):
                                cat = classify_ic_line(line)
                                matched_rows.append({
                                    "Cuenta": line,
                                    "Categoría": cat,
                                    f"Matriz ({matriz_name})": saldos["Matriz"],
                                    f"Filial ({filial_name})": saldos["Filial"]
                                })
                                
                        if not matched_rows:
                            st.info("No se detectaron cuentas con saldos intercompany para este periodo bajo los criterios de búsqueda.")
                        else:
                            df_matched = pd.DataFrame(matched_rows)
                            
                            col_m_name = f"Matriz ({matriz_name})"
                            col_f_name = f"Filial ({filial_name})"
                            
                            # Totales
                            m_act = df_matched[df_matched['Categoría'] == 'Activos IC'][col_m_name].sum()
                            f_pas = df_matched[df_matched['Categoría'] == 'Pasivos IC'][col_f_name].sum()
                            
                            f_act = df_matched[df_matched['Categoría'] == 'Activos IC'][col_f_name].sum()
                            m_pas = df_matched[df_matched['Categoría'] == 'Pasivos IC'][col_m_name].sum()
                            
                            m_ing = df_matched[df_matched['Categoría'] == 'Ingresos IC'][col_m_name].sum()
                            f_gas = df_matched[df_matched['Categoría'] == 'Gastos IC'][col_f_name].sum()
                            
                            f_ing = df_matched[df_matched['Categoría'] == 'Ingresos IC'][col_f_name].sum()
                            m_gas = df_matched[df_matched['Categoría'] == 'Gastos IC'][col_m_name].sum()
                            
                            # Mostrar Resumen de Cuadraturas
                            st.write("### 📊 Resumen de Conciliaciones Intercompany")
                            
                            col_met1, col_met2 = st.columns(2)
                            
                            with col_met1:
                                st.write("**Cruces de Situación Financiera (Balance)**")
                                # Cruce 1: Matriz Activo vs Filial Pasivo
                                diff_1 = abs(m_act) - abs(f_pas)
                                m_act_f = f"{m_act:,.0f}".replace(",", ".")
                                f_pas_f = f"{abs(f_pas):,.0f}".replace(",", ".")
                                diff_1_f = f"{diff_1:,.0f}".replace(",", ".")
                                if abs(diff_1) <= 999.0:
                                    st.success(f"**Matriz Activos vs Filial Pasivos**\n\nCuadrado (Dif: ${diff_1_f})\n* Matriz Activos: ${m_act_f}\n* Filial Pasivos: ${f_pas_f}")
                                else:
                                    st.warning(f"**Matriz Activos vs Filial Pasivos**\n\n⚠️ Descuadre: ${diff_1_f}\n* Matriz Activos: ${m_act_f}\n* Filial Pasivos: ${f_pas_f}")
                                    
                                # Cruce 2: Filial Activo vs Matriz Pasivo
                                diff_2 = abs(f_act) - abs(m_pas)
                                f_act_f = f"{f_act:,.0f}".replace(",", ".")
                                m_pas_f = f"{abs(m_pas):,.0f}".replace(",", ".")
                                diff_2_f = f"{diff_2:,.0f}".replace(",", ".")
                                if abs(diff_2) <= 999.0:
                                    st.success(f"**Filial Activos vs Matriz Pasivos**\n\nCuadrado (Dif: ${diff_2_f})\n* Filial Activos: ${f_act_f}\n* Matriz Pasivos: ${m_pas_f}")
                                else:
                                    st.warning(f"**Filial Activos vs Matriz Pasivos**\n\n⚠️ Descuadre: ${diff_2_f}\n* Filial Activos: ${f_act_f}\n* Matriz Pasivos: ${m_pas_f}")
                                    
                            with col_met2:
                                st.write("**Cruces de Estado de Resultados (P&L)**")
                                # Cruce 3: Matriz Ingreso vs Filial Gasto
                                diff_3 = abs(m_ing) - abs(f_gas)
                                m_ing_f = f"{abs(m_ing):,.0f}".replace(",", ".")
                                f_gas_f = f"{f_gas:,.0f}".replace(",", ".")
                                diff_3_f = f"{diff_3:,.0f}".replace(",", ".")
                                if abs(diff_3) <= 999.0:
                                    st.success(f"**Matriz Ingresos vs Filial Gastos**\n\nCuadrado (Dif: ${diff_3_f})\n* Matriz Ingresos: ${m_ing_f}\n* Filial Gastos: ${f_gas_f}")
                                else:
                                    st.warning(f"**Matriz Ingresos vs Filial Gastos**\n\n⚠️ Descuadre: ${diff_3_f}\n* Matriz Ingresos: ${m_ing_f}\n* Filial Gastos: ${f_gas_f}")
                                    
                                # Cruce 4: Filial Ingreso vs Matriz Gasto
                                diff_4 = abs(f_ing) - abs(m_gas)
                                f_ing_f = f"{abs(f_ing):,.0f}".replace(",", ".")
                                m_gas_f = f"{m_gas:,.0f}".replace(",", ".")
                                diff_4_f = f"{diff_4:,.0f}".replace(",", ".")
                                if abs(diff_4) <= 999.0:
                                    st.success(f"**Filial Ingresos vs Matriz Gastos**\n\nCuadrado (Dif: ${diff_4_f})\n* Filial Ingresos: ${f_ing_f}\n* Matriz Gastos: ${m_gas_f}")
                                else:
                                    st.warning(f"**Filial Ingresos vs Matriz Gastos**\n\n⚠️ Descuadre: ${diff_4_f}\n* Filial Ingresos: ${f_ing_f}\n* Matriz Gastos: ${m_gas_f}")
                                    
                            # Detalle desglosado
                            st.write("---")
                            st.write("### 📋 Detalle de Saldos Detectados")
                            
                            df_balance = df_matched[df_matched['Categoría'].isin(['Activos IC', 'Pasivos IC', 'Inversiones IC'])].copy()
                            df_pl = df_matched[df_matched['Categoría'].isin(['Ingresos IC', 'Gastos IC'])].copy()
                            
                            # Formateadores para visualización estética
                            def format_val(x):
                                return f"{int(round(x)):,}".replace(",", ".") if pd.notna(x) else "0"
                                
                            if not df_balance.empty:
                                st.write("**Saldos de Situación Financiera (Balance)**")
                                df_bal_disp = df_balance.copy()
                                df_bal_disp[col_m_name] = df_bal_disp[col_m_name].apply(format_val)
                                df_bal_disp[col_f_name] = df_bal_disp[col_f_name].apply(format_val)
                                st.dataframe(df_bal_disp, use_container_width=True, key="df_bal_concil_disp")
                                
                            if not df_pl.empty:
                                st.write("**Saldos de Resultados (P&L)**")
                                df_pl_disp = df_pl.copy()
                                df_pl_disp[col_m_name] = df_pl_disp[col_m_name].apply(format_val)
                                df_pl_disp[col_f_name] = df_pl_disp[col_f_name].apply(format_val)
                                st.dataframe(df_pl_disp, use_container_width=True, key="df_pl_concil_disp")

    with tab_hoja:
        st.subheader("Hoja de Trabajo Consolidada (Read-Only)")
        
        st.info("💡 **Recordatorio:** Si el periodo que necesitas consolidar no aparece en el listado, verifica que los periodos históricos de las filiales a consolidar hayan sido congelados/cerrados en el módulo **🔒 Cierre de Periodo**.")
        
        db = SessionLocal()
        grupos_disp2 = db.query(ConsolidationGroup).all()
        db.close()
        
        if not grupos_disp2:
            st.warning("Debes configurar un grupo primero.")
        else:
            grupo_dict2 = {g.id: g.nombre_grupo for g in grupos_disp2}
            col_g, col_p2 = st.columns(2)
            sel_g2 = col_g.selectbox("Seleccionar Grupo a Consolidar", options=list(grupo_dict2.keys()), format_func=lambda x: grupo_dict2[x], key="sel_g2")
            
            db = SessionLocal()
            per_recs2 = db.query(HistoricalDataRecord.periodo).distinct().all()
            db.close()
            periodos_hist2 = sorted([r[0] for r in per_recs2], reverse=True)
            if not periodos_hist2: periodos_hist2 = ["2026-12", "2025-12"]
            
            periodo_cons = col_p2.selectbox("Periodo a Consolidar", periodos_hist2, key="per_cons", format_func=format_periodo)
            
            if st.button("Generar Consolidación", type="primary"):
                with st.spinner("Compilando matrices y aplicando ajustes recurrentes y del periodo..."):
                    import sys
                    import importlib
                    import time
                    start_time = time.time()
                    import src.core.consolidacion_engine
                    importlib.reload(src.core.consolidacion_engine)
                    from src.core.consolidacion_engine import generar_hoja_trabajo
                    df_hoja, msg = generar_hoja_trabajo(sel_g2, periodo_cons)
                    elapsed_time = time.time() - start_time
                    
                    if df_hoja is not None:
                        st.success(f"✅ {msg} (Tiempo de ejecución: {elapsed_time:.2f} segundos)")
                        
                        try:
                            from src.reporting.formatting import apply_corporate_style
                            styled_df = apply_corporate_style(df_hoja)
                            st.markdown(styled_df.to_html(index=False), unsafe_allow_html=True)
                        except Exception as e:
                            # Fallback if styling fails
                            st.dataframe(df_hoja, height=500, key=f"df_hoja_fallback_{sel_g2}_{periodo_cons}")
                        
                        # --- VALIDACIÓN DE ECUACIÓN CONTABLE ---
                        row_activos = df_hoja[df_hoja['Balance clasificado'].astype(str).str.lower().str.strip() == "total activos"]
                        row_pat_pas = df_hoja[df_hoja['Balance clasificado'].astype(str).str.lower().str.strip() == "total patrimonio y pasivos"]
                        
                        if not row_activos.empty and not row_pat_pas.empty:
                            st.write("")
                            st.subheader("🔍 Validación de Ecuación Contable: Activos - (Pasivos + Patrimonio)")
                            
                            cols_verificar = df_hoja.columns[1:] # Excluir 'Balance clasificado'
                            cols_metrics = st.columns(len(cols_verificar))
                            
                            for idx, col in enumerate(cols_verificar):
                                try:
                                    val_act = float(row_activos[col].values[0])
                                except:
                                    val_act = 0.0
                                    
                                try:
                                    # En la base, total patrimonio y pasivos está negativo
                                    val_pp = float(row_pat_pas[col].values[0])
                                except:
                                    val_pp = 0.0
                                    
                                diff = abs(val_act) - abs(val_pp)
                                
                                with cols_metrics[idx]:
                                    if abs(diff) < 1.0:
                                        st.success(f"**{col}**\n\nCuadrado (0)")
                                    else:
                                        diff_fmt = f"{diff:,.0f}".replace(",", ".")
                                        st.error(f"**{col}**\n\nDescuadre: {diff_fmt}")
                        
                        excel_data = df_to_excel_bytes(df_hoja, "Consolidacion")
                        st.download_button(
                            label=f"📥 Descargar WP Consolidación {format_periodo(periodo_cons)} (Excel)",
                            data=excel_data,
                            file_name=f"WP_Consolidacion_{periodo_cons}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
                    else:
                        st.error(f"❌ Error al generar: {msg}")


