import streamlit as st
import pandas as pd
import os
from src.core.excel_utils import format_periodo
from src.core.consolidacion_account_mapper import get_consolidation_accounts_list
from src.core.cierre_engine import es_periodo_cerrado
from src.core.audit_adjustments_engine import (
    guardar_asiento_ajuste,
    obtener_asientos_ajuste,
    eliminar_asiento_ajuste
)

def render(empresa_seleccionada, empresa_path):
    if empresa_seleccionada == "[GLOBAL] Configuración General" or "GLOBAL" in empresa_seleccionada or str(empresa_seleccionada).startswith("[GRUPO]"):
        st.warning("⚠️ **Módulo de Empresa Individual:** Los ajustes de auditoría extracontables solo aplican para empresas individuales. Para realizar eliminaciones o reclasificaciones consolidadas, dirígete al módulo de **Consolidación**.")
        st.stop()
        
    st.title("⚖️ Ajustes de Auditoría (Individuales)")
    st.write(
        "Registra asientos de ajuste post-cierre o extracontables para los **Estados Financieros Individuales** de la empresa activa. "
        "Utilizan partida doble obligatoria (Debe = Haber) con el catálogo oficial de cuentas de ajuste (serie 9999xxx) y afectan exclusivamente los reportes individuales sin alterar la consolidación."
    )
    
    from src.models.trial_balance_db import TrialBalanceDB
    TrialBalanceDB.initialize()
    available_periods = TrialBalanceDB.get_available_periods(empresa_seleccionada)
    
    if not available_periods:
        st.warning("⚠️ No hay periodos activos ni históricos para esta empresa. Carga primero un Trial Balance desde el módulo de Carga de Datos.")
        return

    col_per, col_info = st.columns([1, 2])
    with col_per:
        periodo_sel = st.selectbox("Seleccionar Periodo Contable", available_periods, format_func=format_periodo, key="adj_sel_periodo")
    
    cerrado = es_periodo_cerrado(empresa_seleccionada, periodo_sel)
    with col_info:
        if cerrado:
            st.warning("🔒 **Periodo Cerrado:** Este periodo está cerrado en el histórico. Los ajustes están en modo lectura. Para agregar o modificar ajustes, reversa el cierre en la sección de **Históricos**.")
        else:
            st.success("🟢 **Periodo Activo:** Puedes registrar nuevos asientos de ajuste, editarlos o eliminarlos.")

    st.markdown("---")

    # Iniciar estado de borrador para el comprobante
    if 'draft_adj_lineas' not in st.session_state:
        st.session_state['draft_adj_lineas'] = []
    if 'draft_adj_periodo' not in st.session_state or st.session_state['draft_adj_periodo'] != periodo_sel:
        st.session_state['draft_adj_periodo'] = periodo_sel
        st.session_state['draft_adj_lineas'] = []
        st.session_state['draft_adj_editando_codigo'] = None
        st.session_state['draft_adj_glosa'] = ""

    # Cargar cuentas de ajuste (serie 9999xxx)
    accounts = get_consolidation_accounts_list(empresa_seleccionada)
    accounts_dict = {a["codigo"]: a for a in accounts}
    accounts_codes = [a["codigo"] for a in accounts]

    edit_codigo = st.session_state.get('draft_adj_editando_codigo')

    if not cerrado:
        if edit_codigo:
            st.subheader(f"✏️ Modificando Asiento de Auditoría: {edit_codigo}")
            st.info(f"💡 Estás editando el comprobante **{edit_codigo}**. Puedes agregar nuevas líneas, eliminar líneas erróneas, cambiar montos o la glosa explicativa. Al guardar, se actualizará el comprobante existente conservando su folio.")
        else:
            st.subheader("📝 Nuevo Asiento de Auditoría")
        
        if not accounts_codes:
            st.error("No se encontraron cuentas de ajuste (serie 9999xxx) configuradas en los mapeos de esta empresa.")
        else:
            with st.container():
                col_acc, col_debe, col_haber = st.columns([3, 1, 1])
                
                with col_acc:
                    sel_cod = st.selectbox(
                        "Cuenta de Ajuste / Rubro",
                        options=accounts_codes,
                        format_func=lambda c: accounts_dict[c]["display_label"],
                        key="adj_sel_acc"
                    )
                
                info_cuenta = accounts_dict.get(sel_cod, {})
                rubro_val = info_cuenta.get("linea_item", "")
                nota_val = info_cuenta.get("linea_nota")
                tipo_val = info_cuenta.get("tipo", "Balance")
                lbl_nota = f" • **Nota:** `{nota_val}`" if nota_val else " • **Nota:** *Sin Desglose*"
                
                with col_debe:
                    monto_debe = st.number_input("Debe (+)", min_value=0.0, step=1000.0, key="adj_input_debe")
                with col_haber:
                    monto_haber = st.number_input("Haber (-)", min_value=0.0, step=1000.0, key="adj_input_haber")

                st.caption(f"📌 **Clasificación Automática:** EEFF: `{tipo_val}` • **Rubro:** `{rubro_val}`{lbl_nota}")

                col_btn_add, _ = st.columns([1, 3])
                with col_btn_add:
                    if st.button("➕ Agregar Línea al Asiento", use_container_width=True, type="secondary"):
                        if monto_debe == 0.0 and monto_haber == 0.0:
                            st.error("Debes ingresar un monto mayor a cero en el Debe o en el Haber.")
                        elif monto_debe > 0.0 and monto_haber > 0.0:
                            st.error("Una misma línea no puede tener montos simultáneos al Debe y al Haber. Ingresa líneas separadas.")
                        else:
                            st.session_state['draft_adj_version'] = st.session_state.get('draft_adj_version', 0) + 1
                            st.session_state['draft_adj_lineas'].append({
                                "cuenta_codigo": sel_cod,
                                "nombre_cuenta": info_cuenta.get("nombre_cuenta", ""),
                                "linea_item": rubro_val,
                                "linea_nota": nota_val,
                                "debe": float(monto_debe),
                                "haber": float(monto_haber)
                            })
                            st.rerun()

            # Mostrar tabla borrador editable si tiene líneas
            if st.session_state['draft_adj_lineas']:
                st.write("### 📋 Líneas del Asiento (Editable en Pantalla)")
                st.caption("💡 **Tip:** Haz doble clic en la celda de **Cuenta / Rubro** para seleccionar otra cuenta desde el desplegable del catálogo. También puedes hacer doble clic en **Debe** o **Haber** para cambiar los montos, o seleccionar una fila y presionar **Supr** (Delete) para eliminarla.")
                
                # Diccionario de etiquetas para el selector de cuentas dentro de la tabla
                label_to_acc = {a["display_label"]: a for a in accounts}
                
                table_rows = []
                for idx, l in enumerate(st.session_state['draft_adj_lineas']):
                    c_cod = l["cuenta_codigo"]
                    if c_cod in accounts_dict:
                        acc_lbl = accounts_dict[c_cod]["display_label"]
                    else:
                        acc_lbl = f"{c_cod} — {l.get('linea_item', '')}"
                        if acc_lbl not in label_to_acc:
                            label_to_acc[acc_lbl] = {
                                "codigo": c_cod,
                                "nombre_cuenta": l.get("nombre_cuenta", ""),
                                "linea_item": l.get("linea_item", ""),
                                "linea_nota": l.get("linea_nota")
                            }
                    table_rows.append({
                        "Cuenta / Rubro": acc_lbl,
                        "Rubro EEFF": l.get("linea_item", ""),
                        "Nota": l.get("linea_nota") or "—",
                        "Debe ($)": float(l.get("debe", 0.0)),
                        "Haber ($)": float(l.get("haber", 0.0))
                    })
                
                df_draft_raw = pd.DataFrame(table_rows)
                all_account_options = list(label_to_acc.keys())
                
                editor_version = st.session_state.get('draft_adj_version', 0)
                editor_key = f"adj_editor_{empresa_seleccionada}_{periodo_sel}_{editor_version}"
                
                edited_df = st.data_editor(
                    df_draft_raw,
                    key=editor_key,
                    num_rows="dynamic",
                    use_container_width=True,
                    disabled=["Rubro EEFF", "Nota"],
                    column_config={
                        "Cuenta / Rubro": st.column_config.SelectboxColumn(
                            "Cuenta / Rubro",
                            help="Haz doble clic para seleccionar otra cuenta del catálogo.",
                            options=all_account_options,
                            required=True,
                            width="large"
                        ),
                        "Rubro EEFF": st.column_config.TextColumn("Rubro EEFF", disabled=True, width="medium"),
                        "Nota": st.column_config.TextColumn("Nota", disabled=True, width="small"),
                        "Debe ($)": st.column_config.NumberColumn(
                            "Debe ($)",
                            help="Monto al Debe. Haz doble clic para editar.",
                            min_value=0.0,
                            step=1000.0,
                            format="%d"
                        ),
                        "Haber ($)": st.column_config.NumberColumn(
                            "Haber ($)",
                            help="Monto al Haber. Haz doble clic para editar.",
                            min_value=0.0,
                            step=1000.0,
                            format="%d"
                        )
                    }
                )

                # Sincronizar cambios del editor con el estado de sesión
                updated_lines = []
                for idx_row, row in edited_df.iterrows():
                    sel_lbl = str(row.get("Cuenta / Rubro", "")).strip()
                    if not sel_lbl or sel_lbl.lower() == 'nan':
                        continue
                    orig = st.session_state['draft_adj_lineas'][idx_row] if idx_row < len(st.session_state['draft_adj_lineas']) else {}
                    acc_info = label_to_acc.get(sel_lbl, {})
                    c_cod = acc_info.get("codigo", orig.get("cuenta_codigo", ""))
                    nom_cta = acc_info.get("nombre_cuenta", orig.get("nombre_cuenta", ""))
                    rubro_val = acc_info.get("linea_item", orig.get("linea_item", ""))
                    nota_val = acc_info.get("linea_nota", orig.get("linea_nota"))

                    d_val = float(pd.to_numeric(row.get("Debe ($)", 0.0), errors='coerce') or 0.0)
                    h_val = float(pd.to_numeric(row.get("Haber ($)", 0.0), errors='coerce') or 0.0)
                    updated_lines.append({
                        "cuenta_codigo": c_cod,
                        "nombre_cuenta": nom_cta,
                        "linea_item": rubro_val,
                        "linea_nota": nota_val,
                        "debe": d_val,
                        "haber": h_val
                    })

                # Detectar si cambiaron cuentas, valores o filas
                lines_changed = False
                if len(updated_lines) != len(st.session_state['draft_adj_lineas']):
                    lines_changed = True
                else:
                    for u, cur in zip(updated_lines, st.session_state['draft_adj_lineas']):
                        if (u["cuenta_codigo"] != cur["cuenta_codigo"] or
                            abs(u["debe"] - cur["debe"]) > 0.01 or 
                            abs(u["haber"] - cur["haber"]) > 0.01):
                            lines_changed = True
                            break

                if lines_changed:
                    st.session_state['draft_adj_version'] = st.session_state.get('draft_adj_version', 0) + 1
                    st.session_state['draft_adj_lineas'] = updated_lines
                    st.rerun()

                tot_d = sum(l["debe"] for l in st.session_state['draft_adj_lineas'])
                tot_h = sum(l["haber"] for l in st.session_state['draft_adj_lineas'])
                diff = round(tot_d - tot_h, 2)

                c_m1, c_m2, c_m3 = st.columns(3)
                c_m1.metric("Total Debe", f"${tot_d:,.0f}")
                c_m2.metric("Total Haber", f"${tot_h:,.0f}")
                if abs(diff) < 0.01:
                    c_m3.metric("Diferencia (Cuadrado)", "$0", delta="Cuadrado", delta_color="normal")
                else:
                    c_m3.metric("Diferencia (Descuadrado)", f"${diff:,.0f}", delta=f"${diff:,.0f}", delta_color="inverse")

                # Glosa
                valor_glosa_inicial = st.session_state.get('draft_adj_glosa', '')
                glosa_asiento = st.text_input(
                    "Glosa / Explicación del Asiento de Auditoría *",
                    value=valor_glosa_inicial,
                    placeholder="Ej: Reclasificación de provisiones auditoría financiera BDO periodo 2025",
                    key="adj_input_glosa"
                )

                c_save, c_discard, _ = st.columns([1.8, 1.2, 1.5])
                with c_save:
                    puede_guardar = abs(diff) < 0.01 and tot_d > 0 and bool((glosa_asiento or "").strip())
                    btn_save_label = f"💾 Actualizar Asiento {edit_codigo}" if edit_codigo else "💾 Registrar Asiento en Base de Datos"
                    
                    if st.button(btn_save_label, type="primary", disabled=not puede_guardar, use_container_width=True):
                        ok, msg, cod_nuevo = guardar_asiento_ajuste(
                            empresa=empresa_seleccionada,
                            periodo=periodo_sel,
                            glosa=glosa_asiento.strip(),
                            lineas=st.session_state['draft_adj_lineas'],
                            asiento_codigo=edit_codigo
                        )
                        if ok:
                            st.session_state['draft_adj_lineas'] = []
                            st.session_state['draft_adj_editando_codigo'] = None
                            st.session_state['draft_adj_glosa'] = ""
                            # Invalidar firmas de memoria en estados_financieros para recarga inmediata
                            for k in ['_loaded_data_sig', 'tb_df', 'pl_df', 'preview_df', 'balance_excel_binary', 'balance_word_binary', 'er_preview_df', 'er_excel_binary']:
                                if k in st.session_state:
                                    del st.session_state[k]
                            st.success(f"✔️ {msg}")
                            st.rerun()
                        else:
                            st.error(msg)
                            
                with c_discard:
                    btn_discard_label = "❌ Cancelar Edición" if edit_codigo else "🗑️ Descartar Borrador"
                    if st.button(btn_discard_label, use_container_width=True):
                        st.session_state['draft_adj_version'] = st.session_state.get('draft_adj_version', 0) + 1
                        st.session_state['draft_adj_lineas'] = []
                        st.session_state['draft_adj_editando_codigo'] = None
                        st.session_state['draft_adj_glosa'] = ""
                        st.rerun()

    st.markdown("---")

    # Historial de asientos guardados
    st.subheader(f"📚 Asientos de Auditoría Registrados ({format_periodo(periodo_sel)})")
    vouchers = obtener_asientos_ajuste(empresa_seleccionada, periodo_sel)

    if not vouchers:
        st.info(f"ℹ️ No hay asientos de auditoría registrados para el periodo {format_periodo(periodo_sel)}.")
    else:
        st.write(f"Se encontraron **{len(vouchers)}** comprobante(s) de ajuste registrados para este periodo:")
        
        for v in vouchers:
            cod = v["asiento_codigo"]
            glosa = v["glosa"]
            tot_v = v["total_debe"]
            fecha_str = v["fecha_creacion"].strftime("%d/%m/%Y %H:%M") if hasattr(v["fecha_creacion"], "strftime") else str(v["fecha_creacion"])
            
            is_being_edited = (edit_codigo == cod)
            prefix = "✏️ [EN EDICIÓN] " if is_being_edited else ""
            
            with st.expander(f"{prefix}🧾 Comprobante **{cod}** — *{glosa}* (Total: ${tot_v:,.0f})", expanded=is_being_edited):
                st.caption(f"📅 Fecha de Registro: {fecha_str} | Empresa: **{empresa_seleccionada}** | Periodo: **{format_periodo(periodo_sel)}**")
                
                df_v_lines = pd.DataFrame([{
                    "Línea": l["num_linea"],
                    "Cuenta": l["cuenta_codigo"],
                    "Nombre Cuenta": l["nombre_cuenta"],
                    "Rubro EEFF": l["linea_item"],
                    "Nota": l["linea_nota"] or "—",
                    "Debe ($)": l["debe"],
                    "Haber ($)": l["haber"]
                } for l in v["lineas"]])
                
                st.dataframe(
                    df_v_lines.style.format({
                        "Debe ($)": "{:,.0f}",
                        "Haber ($)": "{:,.0f}"
                    }),
                    use_container_width=True,
                    hide_index=True
                )
                
                if not cerrado:
                    col_btn_edit, col_btn_del, _ = st.columns([1.2, 1.2, 3])
                    with col_btn_edit:
                        if st.button(f"✏️ Editar {cod}", key=f"btn_edit_{cod}", type="primary"):
                            st.session_state['draft_adj_version'] = st.session_state.get('draft_adj_version', 0) + 1
                            st.session_state['draft_adj_editando_codigo'] = cod
                            st.session_state['draft_adj_glosa'] = glosa
                            st.session_state['draft_adj_lineas'] = [
                                {
                                    "cuenta_codigo": l["cuenta_codigo"],
                                    "nombre_cuenta": l["nombre_cuenta"],
                                    "linea_item": l["linea_item"],
                                    "linea_nota": l["linea_nota"],
                                    "debe": float(l["debe"]),
                                    "haber": float(l["haber"])
                                } for l in v["lineas"]
                            ]
                            st.rerun()
                            
                    with col_btn_del:
                        if st.button(f"🗑️ Eliminar {cod}", key=f"btn_del_{cod}", type="secondary"):
                            ok, msg = eliminar_asiento_ajuste(empresa_seleccionada, periodo_sel, cod)
                            if ok:
                                if st.session_state.get('draft_adj_editando_codigo') == cod:
                                    st.session_state['draft_adj_editando_codigo'] = None
                                    st.session_state['draft_adj_lineas'] = []
                                    st.session_state['draft_adj_glosa'] = ""
                                for k in ['_loaded_data_sig', 'tb_df', 'pl_df', 'preview_df', 'balance_excel_binary', 'balance_word_binary', 'er_preview_df', 'er_excel_binary']:
                                    if k in st.session_state:
                                        del st.session_state[k]
                                st.success(msg)
                                st.rerun()
                            else:
                                st.error(msg)
