import streamlit as st
import pandas as pd

from src.core.snapshot_vault_manager import SnapshotVaultManager
from src.core.validation_tie_out import ValidationTieOutEngine

_VAL_CACHE_VERSION = "v4_rounding_reconciler_hierarchical"


def _fmt_m(val: int) -> str:
    try:
        return f"{int(round(val)):,}".replace(",", ".")
    except Exception:
        return "0"


def render(empresa_seleccionada, empresa_path):
    if empresa_seleccionada == "[GLOBAL] Configuración General" or "GLOBAL" in str(empresa_seleccionada):
        st.info("Por favor, selecciona una empresa de trabajo o grupo consolidado en el menú lateral izquierdo para acceder al Módulo de Validación.")
        st.stop()

    is_consolidated = str(empresa_seleccionada).startswith("[GRUPO] ")
    clean_empresa_name = str(empresa_seleccionada).replace("[GRUPO] ", "").strip()

    st.title("Validaciones y Cuadratura de Notas (M$)")
    st.caption(
        "Comparación directa entre líneas de Estados Financieros, cuadros principales de Notas y sub-cuadros internos en Miles de Pesos (M\\$) — **Tolerancia Estricta Cero (`Diferencia = 0 M$`)**."
    )

    available_periods = ValidationTieOutEngine.get_available_validation_periods(empresa_seleccionada)

    def _fmt_periodo_opt(p: str) -> str:
        if SnapshotVaultManager.has_active_snapshot(empresa_seleccionada, p):
            return f"🔒 {p} (Snapshot Cerrado)"
        return f"⚡ {p} (Período Abierto)"

    col_p1, col_p2, col_p3 = st.columns([3, 2, 3])
    with col_p1:
        per_key = f"val_sel_periodo_{clean_empresa_name}_{'cons' if is_consolidated else 'ind'}"
        sel_periodo = st.selectbox(
            "Período a Validar / Auditar:",
            available_periods,
            format_func=_fmt_periodo_opt,
            key=per_key,
        )

    cache_store = st.session_state.setdefault("_tie_out_matrix_cache", {})
    cache_key = (_VAL_CACHE_VERSION, empresa_seleccionada, sel_periodo)

    with col_p2:
        st.write(" ")
        btn_recalcular = st.button(
            "🔄 Recalcular Validación",
            type="primary",
            use_container_width=True,
            key=f"btn_val_recalc_{clean_empresa_name}_{sel_periodo}",
        )
        if btn_recalcular and cache_key in cache_store:
            del cache_store[cache_key]

    if cache_key not in cache_store:
        with st.spinner(f"Ejecutando cuadratura directa en M$ para {empresa_seleccionada} ({sel_periodo})..."):
            cache_store[cache_key] = ValidationTieOutEngine.obtener_matriz_tie_out(
                empresa_o_grupo=clean_empresa_name,
                periodo=sel_periodo,
                is_consolidated=is_consolidated,
            )

    df_matrix, health = cache_store[cache_key]

    with col_p3:
        st.caption(f"Entidad: **{empresa_seleccionada}**")
        st.caption(f"Fuente: **{health.get('source_label', '')}** ({health.get('elapsed_sec', 0):.2f} s)")

    # =========================================================================
    # SEMÁFOROS / KPIS SUPERIORES DE SALUD FINANCIERA (TOLERANCIA CERO)
    # =========================================================================
    st.divider()
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)

    with kpi1:
        tot_reg = health.get("total_reglas", 0)
        tot_ok = health.get("total_cuadrados", 0)
        tot_err = health.get("total_descuadres", 0)
        if health.get("is_valid") and tot_reg > 0:
            st.metric(
                label="Tie-Out EEFF vs. Notas",
                value="100% Cuadrado",
                delta=f"{tot_ok}/{tot_reg} controles en 0 M$",
            )
        else:
            st.metric(
                label="Tie-Out EEFF vs. Notas",
                value=f"{tot_err} Descuadre(s)",
                delta=f"{tot_ok}/{tot_reg} cuadrados (Tol: 0 M$)",
                delta_color="inverse",
            )

    with kpi2:
        eq = health.get("ecuacion_patrimonial", {})
        if eq.get("ok", False):
            st.metric(label="Ecuación Patrimonial", value="A = P + Pt", delta="Cuadrado (0 M$)")
        else:
            st.metric(
                label="Ecuación Patrimonial",
                value=f"Diff: {_fmt_m(eq.get('diff', 0))} M$",
                delta="Descuadre Balance",
                delta_color="inverse",
            )

    with kpi3:
        cr = health.get("cruce_resultado", {})
        if cr.get("ok", False):
            st.metric(label="Resultado del Ejercicio", value="ER = Patrimonio", delta="Cuadrado (0 M$)")
        else:
            st.metric(
                label="Resultado del Ejercicio",
                value=f"Diff: {_fmt_m(cr.get('diff', 0))} M$",
                delta="ER vs. Patrimonio",
                delta_color="inverse",
            )

    with kpi4:
        ce = health.get("cruce_efectivo", {})
        if ce.get("ok", False):
            st.metric(label="Efectivo Final", value="Balance = EFE", delta="Cuadrado (0 M$)")
        else:
            st.metric(
                label="Efectivo Final",
                value=f"Diff: {_fmt_m(ce.get('diff', 0))} M$",
                delta="Balance vs. EFE",
                delta_color="inverse",
            )

    st.divider()

    if df_matrix is None or df_matrix.empty:
        st.info(f"No se encontraron registros de Estados Financieros para **{empresa_seleccionada}** en el período **{sel_periodo}**.")
        return

    st.subheader("Matriz de Validación Directa (M$ — Tolerancia Cero)")

    col_f1, col_f2, col_f3, col_f4 = st.columns([2, 2, 2, 3])
    with col_f1:
        sel_nivel = st.selectbox(
            "Nivel de Control",
            ["Todos", "EEFF vs. Nota", "Cuadratura Interna Nota", "Cruce entre EEFF"],
            key="val_f_nivel",
        )
    with col_f2:
        sel_filtro_rep = st.selectbox(
            "Estado Financiero",
            ["Todos", "Balance Clasificado", "Estado de Resultados"],
            key="val_f_rep",
        )
    with col_f3:
        sel_filtro_est = st.selectbox(
            "Estado de Cuadratura",
            ["Todos", "❌ Solo Descuadres (≠ 0)", "✅ Solo Cuadrados (= 0)"],
            key="val_f_est",
        )
    with col_f4:
        txt_busqueda = (
            st.text_input(
                "Buscar por Código, Rubro o Nota",
                placeholder="Ej: #N04, Efectivo, Derechos de uso, Fondos mutuos...",
                key="val_f_txt",
            )
            .strip()
            .lower()
        )

    ocultar_ceros = st.checkbox(
        "Ocultar líneas sin saldo en ambos lados (0 M\\$ = 0 M\\$)",
        value=False,
        key="val_chk_hide_zeros",
    )

    df_filtered = df_matrix.copy()
    if sel_nivel != "Todos":
        df_filtered = df_filtered[df_filtered["Nivel"] == sel_nivel]

    if sel_filtro_rep != "Todos":
        df_filtered = df_filtered[df_filtered["Reporte"] == sel_filtro_rep]

    if sel_filtro_est.startswith("❌"):
        df_filtered = df_filtered[df_filtered["Diferencia (M$)"] != 0]
    elif sel_filtro_est.startswith("✅"):
        df_filtered = df_filtered[df_filtered["Diferencia (M$)"] == 0]

    if ocultar_ceros:
        df_filtered = df_filtered[
            (df_filtered["Saldo EEFF (M$)"] != 0)
            | (df_filtered["Saldo Nota (M$)"] != 0)
        ]

    if txt_busqueda:
        mask = (
            df_filtered["Código"].astype(str).str.lower().str.contains(txt_busqueda, regex=False)
            | df_filtered["Rubro EEFF"].astype(str).str.lower().str.contains(txt_busqueda, regex=False)
            | df_filtered["Nota"].astype(str).str.lower().str.contains(txt_busqueda, regex=False)
        )
        df_filtered = df_filtered[mask]

    n_desc_vis = int((df_filtered["Diferencia (M$)"] != 0).sum())
    n_rond_vis = int((df_filtered["Diferencia (M$)"].abs().isin([1, 2])).sum())
    st.caption(
        f"Mostrando **{len(df_filtered)}** de **{len(df_matrix)}** controles — "
        f"**{n_desc_vis}** con diferencia ≠ 0 M\\$"
        + (f" *(de los cuales **{n_rond_vis}** corresponden a diferencias de redondeo de ±1 a ±2 M\\$)*." if n_rond_vis > 0 else ".")
    )

    df_display = df_filtered.copy().reset_index(drop=True)
    raw_diffs = df_display["Diferencia (M$)"].tolist()
    df_display["Saldo EEFF (M$)"] = df_display["Saldo EEFF (M$)"].apply(_fmt_m)
    df_display["Saldo Nota (M$)"] = df_display["Saldo Nota (M$)"].apply(_fmt_m)
    df_display["Diferencia (M$)"] = df_display["Diferencia (M$)"].apply(_fmt_m)

    def _style_row(row):
        idx = row.name
        is_err = raw_diffs[idx] != 0 if idx < len(raw_diffs) else False
        if is_err:
            return [
                "background-color: #FFF5F5; color: #9B1C1C; font-weight: 600;" if col in ("Diferencia (M$)", "Estado")
                else "background-color: #FFF9F9;"
                for col in row.index
            ]
        return ["" for _ in row.index]

    styled_df = df_display.style.apply(_style_row, axis=1)

    st.dataframe(
        styled_df,
        use_container_width=True,
        hide_index=True,
        height=min(max(240, 38 * (len(df_display) + 1)), 760),
        column_config={
            "Código": st.column_config.TextColumn("Código", width="small"),
            "Rubro EEFF": st.column_config.TextColumn("Rubro EEFF", width="large"),
            "Nota": st.column_config.TextColumn("Nota", width="large"),
            "Saldo EEFF (M$)": st.column_config.TextColumn("Saldo EEFF (M$)", width="medium"),
            "Saldo Nota (M$)": st.column_config.TextColumn("Saldo Nota (M$)", width="medium"),
            "Diferencia (M$)": st.column_config.TextColumn("Diferencia (M$)", width="small"),
            "Estado": st.column_config.TextColumn("Estado", width="small"),
            "Reporte": st.column_config.TextColumn("Reporte", width="medium"),
            "Nivel": st.column_config.TextColumn("Nivel", width="medium"),
        },
        key=f"df_tie_out_v7_{clean_empresa_name}_{sel_periodo}_{sel_nivel}_{sel_filtro_rep}_{sel_filtro_est}_{ocultar_ceros}",
    )

    st.divider()

    excel_bytes = ValidationTieOutEngine.generar_excel_tie_out(
        df_matrix=df_matrix,
        health_checks=health,
        empresa_o_grupo=empresa_seleccionada,
        periodo=sel_periodo,
    )

    clean_file_name = clean_empresa_name.replace(" ", "_")
    st.download_button(
        label=f"📥 Descargar Matriz de Validación Directa en M$ (Excel) — {sel_periodo}",
        data=excel_bytes,
        file_name=f"Matriz_Validacion_Directa_M$_{clean_file_name}_{sel_periodo}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
