import pandas as pd
import openpyxl
from io import BytesIO
import unicodedata
import re
import datetime
import calendar
from copy import copy

from src.core.ifrs_glossary import translate_ifrs_term, extract_mapping_en_overrides


def format_period_date(period_str, lang='es'):
    if not period_str:
        return ""
    try:
        from src.core.pl_cubo_processor import parse_month_to_num, parse_year_to_num
        year = parse_year_to_num(period_str)
        month = parse_month_to_num(period_str)
        if year and month:
            last_day = calendar.monthrange(year, month)[1]
            if str(lang).lower() == 'en':
                months_en = {
                    1: "January", 2: "February", 3: "March", 4: "April",
                    5: "May", 6: "June", 7: "July", 8: "August",
                    9: "September", 10: "October", 11: "November", 12: "December"
                }
                return f"{months_en.get(month, '')} {last_day}, {year}"
            else:
                months_es = {
                    1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
                    5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
                    9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
                }
                return f"{last_day} de {months_es.get(month, '')} {year}"
    except Exception:
        pass
    return str(period_str)


format_period_to_spanish_date = format_period_date


class BalanceGenerator:
    def __init__(self, template_path):
        self.template_path = template_path

    def generate(
        self,
        tb_df,
        map_balance_df,
        scale_factor=1.0,
        tb_df_comp=None,
        periodo_actual_str=None,
        periodo_comp_str=None,
        target_lang=None,
        empresa=None,
        er_net_m=None,
    ):
        if target_lang is None:
            try:
                import streamlit as st
                target_lang = st.session_state.get('idioma_reporte', 'es')
            except Exception:
                target_lang = 'es'

        overrides = extract_mapping_en_overrides(map_balance_df)
        """
        1. Cruza tb_df con map_balance_df para obtener la 'Clasificación balance' de cada cuenta.
        2. Suma los saldos por cada 'Clasificación balance'.
        3. Verifica guardarraíl contable en Pesos (<= $4.000) y aplica Redondeo Jerárquico por Mayor Residuo en M$.
        4. Inyecta los valores en el template de Excel preservando formato y fórmulas adicionales.
        """
        tb_df = tb_df.copy()
        acc_col_tb = next(
            (c for c in tb_df.columns if str(c).strip().lower() in ['cuenta_id', 'cuenta', 'n° de cuenta', 'n de cuenta', 'n° de cuenta \n']),
            tb_df.columns[0],
        )
        sf_col_tb = next(
            (c for c in tb_df.columns if str(c).strip().lower() in ['saldo_final', 'saldo final', 'saldo']),
            None,
        )

        tb_df['cuenta_id'] = tb_df[acc_col_tb].astype(str).str.strip()
        if sf_col_tb:
            tb_df['saldo_final'] = pd.to_numeric(tb_df[sf_col_tb], errors='coerce').fillna(0.0)
        elif 'saldo_final' in tb_df.columns:
            tb_df['saldo_final'] = pd.to_numeric(tb_df['saldo_final'], errors='coerce').fillna(0.0)
        else:
            tb_df['saldo_final'] = 0.0

        tb_total_pesos = float(tb_df['saldo_final'].sum()) if not tb_df.empty else 0.0

        map_balance_df = map_balance_df.copy()
        col_map_acc = next(
            (c for c in map_balance_df.columns if str(c).strip().lower() in ['n° de cuenta', 'n de cuenta', 'cuenta', 'cuenta_id']),
            map_balance_df.columns[0],
        )
        map_balance_df['N° de Cuenta'] = map_balance_df[col_map_acc].astype(str).str.strip()

        cls_col = next(
            (c for c in map_balance_df.columns if "clasificaci" in c.lower() and "balance" in c.lower()),
            "Clasificación balance",
        )
        if cls_col in map_balance_df.columns:
            map_balance_df[cls_col] = map_balance_df[cls_col].apply(lambda x: str(x).strip() if pd.notna(x) else x)

        from src.core.sabana_builder import build_balance_sabana
        merged_df = build_balance_sabana(tb_df, map_balance_df)

        mapped_mask = merged_df[cls_col].notna()
        agg_mapped = merged_df[mapped_mask].groupby(cls_col, as_index=False)['saldo_final'].sum()
        pl_residual = float(merged_df[~mapped_mask]['saldo_final'].sum())
        sums_dict = dict(zip(agg_mapped[cls_col].str.strip(), agg_mapped['saldo_final']))

        # Inferir nombre de empresa desde parámetro explícito o template_path
        tp_norm = str(self.template_path or "").replace("\\", "/")
        emp_inferred = str(empresa).strip() if empresa else None
        if not emp_inferred and "data/empresas/" in tp_norm:
            parts = tp_norm.split("data/empresas/")[1].split("/")
            if len(parts) >= 2:
                emp_inferred = parts[0].strip()

        sums_comp_dict = {}
        pl_residual_comp = 0.0
        tb_comp_total_pesos = 0.0
        if tb_df_comp is not None and not tb_df_comp.empty:
            tb_df_comp = tb_df_comp.copy()
            acc_col_comp = next(
                (c for c in tb_df_comp.columns if str(c).strip().lower() in ['cuenta_id', 'cuenta', 'n° de cuenta', 'n de cuenta', 'n° de cuenta \n']),
                tb_df_comp.columns[0],
            )
            sf_col_comp = next(
                (c for c in tb_df_comp.columns if str(c).strip().lower() in ['saldo_final', 'saldo final', 'saldo']),
                None,
            )
            tb_df_comp['cuenta_id'] = tb_df_comp[acc_col_comp].astype(str).str.strip()
            if sf_col_comp:
                tb_df_comp['saldo_final'] = pd.to_numeric(tb_df_comp[sf_col_comp], errors='coerce').fillna(0.0)
            elif 'saldo_final' in tb_df_comp.columns:
                tb_df_comp['saldo_final'] = pd.to_numeric(tb_df_comp['saldo_final'], errors='coerce').fillna(0.0)
            else:
                tb_df_comp['saldo_final'] = 0.0

            tb_comp_total_pesos = float(tb_df_comp['saldo_final'].sum())

            map_balance_comp_df = map_balance_df
            try:
                from src.core.snapshot_vault_manager import SnapshotVaultManager
                if emp_inferred and periodo_comp_str and SnapshotVaultManager.has_active_snapshot(emp_inferred, periodo_comp_str):
                    f_mb_c, _ = SnapshotVaultManager.get_frozen_mappings(emp_inferred, periodo_comp_str)
                    if f_mb_c is not None and not f_mb_c.empty:
                        map_balance_comp_df = f_mb_c.copy()
                        col_map_acc_c = next(
                            (c for c in map_balance_comp_df.columns if str(c).strip().lower() in ['n° de cuenta', 'n de cuenta', 'cuenta', 'cuenta_id']),
                            map_balance_comp_df.columns[0],
                        )
                        map_balance_comp_df['N° de Cuenta'] = map_balance_comp_df[col_map_acc_c].astype(str).str.strip()
                        if cls_col in map_balance_comp_df.columns:
                            map_balance_comp_df[cls_col] = map_balance_comp_df[cls_col].apply(lambda x: str(x).strip() if pd.notna(x) else x)
            except Exception:
                pass

            merged_comp = build_balance_sabana(tb_df_comp, map_balance_comp_df)
            mapped_mask_comp = merged_comp[cls_col].notna()
            agg_mapped_comp = merged_comp[mapped_mask_comp].groupby(cls_col, as_index=False)['saldo_final'].sum()
            pl_residual_comp = float(merged_comp[~mapped_mask_comp]['saldo_final'].sum())
            sums_comp_dict = dict(zip(agg_mapped_comp[cls_col].str.strip(), agg_mapped_comp['saldo_final']))

        # Load Template with openpyxl to preserve formatting
        wb = openpyxl.load_workbook(self.template_path)
        ws = wb.active

        from src.core.excel_utils import detect_balance_columns
        name_col_idx, nota_col_idx, val25_col_idx, val24_col_idx = detect_balance_columns(ws)

        # Relocate "Activos" header if needed
        r_activos = None
        r_corrientes = None
        for r in range(1, 15):
            val = ws.cell(row=r, column=name_col_idx).value
            if val and str(val).strip().lower() == "activos":
                r_activos = r
            if val and str(val).strip().lower() == "activos corrientes":
                r_corrientes = r

        if r_activos is not None and r_corrientes is not None:
            target_row = r_corrientes - 1
            if target_row > r_activos:
                orig_cell = ws.cell(row=r_activos, column=name_col_idx)
                target_cell = ws.cell(row=target_row, column=name_col_idx)
                target_cell.value = orig_cell.value
                if orig_cell.font:
                    target_cell.font = copy(orig_cell.font)
                orig_cell.value = None

        has_comp = bool(
            tb_df_comp is not None
            and not tb_df_comp.empty
            and str(periodo_comp_str).strip().lower() not in ["ninguno", "none", "", "comp"]
        )

        if periodo_actual_str:
            formatted_actual = format_period_date(periodo_actual_str, lang=target_lang)
            formatted_comp = format_period_date(periodo_comp_str, lang=target_lang) if (periodo_comp_str and has_comp) else ""

            for r in range(1, 15):
                c25 = ws.cell(row=r, column=val25_col_idx)
                c24 = ws.cell(row=r, column=val24_col_idx)

                v25 = c25.value
                if v25 is not None:
                    is_date_or_year = (
                        isinstance(v25, (datetime.datetime, datetime.date))
                        or (isinstance(v25, str) and ("20" in v25 or "Actual" in v25))
                    )
                    if is_date_or_year:
                        c25.value = formatted_actual

                if has_comp:
                    v24 = c24.value
                    if v24 is not None:
                        is_date_or_year_comp = (
                            isinstance(v24, (datetime.datetime, datetime.date))
                            or (isinstance(v24, str) and ("20" in v24 or "Anterior" in v24))
                        )
                        if is_date_or_year_comp:
                            c24.value = formatted_comp

        def sanitize(text):
            if pd.isna(text):
                return ""
            clean_str = str(text).strip().lower()
            clean_str = ''.join(c for c in unicodedata.normalize('NFD', clean_str) if unicodedata.category(c) != 'Mn')
            clean_str = re.sub(r'\s+', ' ', clean_str)
            return clean_str

        def _is_total_activos_corrientes(s):
            return ('activo' in s or 'asset' in s) and 'corrien' in s and ('total' in s or 'subtotal' in s) and 'no corrien' not in s and 'nocorrien' not in s
        def _is_total_activos_no_corrientes(s):
            return ('activo' in s or 'asset' in s) and ('no corrien' in s or 'nocorrien' in s or 'largo' in s) and ('total' in s or 'subtotal' in s)
        def _is_total_activos(s):
            return ('total' in s or 'subtotal' in s) and ('activo' in s or 'asset' in s) and 'corrien' not in s and 'no corrien' not in s
        def _is_total_pasivos_corrientes(s):
            return ('pasivo' in s or 'liabilit' in s) and 'corrien' in s and ('total' in s or 'subtotal' in s) and 'no corrien' not in s and 'nocorrien' not in s
        def _is_total_pasivos_no_corrientes(s):
            return ('pasivo' in s or 'liabilit' in s) and ('no corrien' in s or 'nocorrien' in s or 'largo' in s) and ('total' in s or 'subtotal' in s)
        def _is_total_pasivos(s):
            return ('total' in s or 'subtotal' in s) and ('pasivo' in s or 'liabilit' in s) and 'corrien' not in s and 'no corrien' not in s and 'patrimonio' not in s
        def _is_patrimonio_controladora(s):
            return 'patrimonio' in s and ('controlad' in s or 'propietari' in s or 'atribuible' in s)
        def _is_participaciones_no_controladoras(s):
            return ('no controlad' in s or 'nocontrolad' in s or 'minoritari' in s) and 'patrimonio' not in s
        def _is_total_patrimonio(s):
            return ('total' in s or 'subtotal' in s) and 'patrimonio' in s and 'pasivo' not in s and 'controlad' not in s and 'propietari' not in s and 'atribuible' not in s
        def _is_total_pasivos_patrimonio(s):
            return ('total' in s or 'subtotal' in s) and 'patrimonio' in s and 'pasivo' in s

        def _is_any_total_fn(s):
            return (
                _is_total_activos_corrientes(s)
                or _is_total_activos_no_corrientes(s)
                or _is_total_activos(s)
                or _is_total_pasivos_corrientes(s)
                or _is_total_pasivos_no_corrientes(s)
                or _is_total_pasivos(s)
                or _is_patrimonio_controladora(s)
                or _is_total_patrimonio(s)
                or _is_total_pasivos_patrimonio(s)
            )

        # Pre-escanear rubros de detalle presentes en la plantilla por sección
        res_acum_labels = {"Resultados acumulados", "Ganancias (pérdidas) acumuladas"}
        section_headers = {
            "activos", "activos corrientes", "activos no corrientes",
            "patrimonio y pasivos", "pasivos corrientes", "pasivos no corrientes", "patrimonio",
        }
        asset_keys = []
        liab_keys = []
        pat_other_keys = []
        scan_sec = "None"
        for r_scan in range(1, ws.max_row + 1):
            c_val = ws.cell(row=r_scan, column=name_col_idx).value
            if c_val and isinstance(c_val, str):
                c_clean = c_val.strip()
                c_low = sanitize(c_val)
                if c_low == "activos corrientes":
                    scan_sec = "Activos Corrientes"
                    continue
                elif c_low == "activos no corrientes":
                    scan_sec = "Activos No Corrientes"
                    continue
                elif c_low == "pasivos corrientes":
                    scan_sec = "Pasivos Corrientes"
                    continue
                elif c_low == "pasivos no corrientes":
                    scan_sec = "Pasivos No Corrientes"
                    continue
                elif c_low == "patrimonio":
                    scan_sec = "Patrimonio"
                    continue

                if c_low in section_headers or _is_any_total_fn(c_low) or c_clean in res_acum_labels:
                    continue
                if scan_sec in ("Activos Corrientes", "Activos No Corrientes"):
                    if c_clean not in asset_keys:
                        asset_keys.append(c_clean)
                elif scan_sec in ("Pasivos Corrientes", "Pasivos No Corrientes"):
                    if c_clean not in liab_keys:
                        liab_keys.append(c_clean)
                elif scan_sec == "Patrimonio":
                    if c_clean not in pat_other_keys:
                        pat_other_keys.append(c_clean)

        template_bs_keys = set(asset_keys) | set(liab_keys) | set(pat_other_keys)

        from src.core.rounding_reconciler import (
            MAX_TOLERANCE_M,
            MAX_TOLERANCE_PESOS,
            _extract_stmt_target,
            is_pl_classification_label,
            largest_remainder_round,
        )

        def _solve_bs_column(sums_map, pl_unmapped_pesos, total_tb_pesos, emp_name_col, periodo_col, expected_res_acum_m=None):
            raw_res_acum_mapped = sum(float(sums_map.get(lbl, 0.0) or 0.0) for lbl in res_acum_labels)
            pl_mapped_pesos = 0.0
            missing_bs_pesos = 0.0
            for k_s, v_s in sums_map.items():
                vf = 0.0 if (v_s is None or pd.isna(v_s)) else float(v_s)
                if k_s in res_acum_labels or k_s in template_bs_keys:
                    continue
                if is_pl_classification_label(k_s):
                    pl_mapped_pesos += vf
                else:
                    missing_bs_pesos += vf

            asset_pesos_vis = {k: float(sums_map.get(k, 0.0) or 0.0) for k in asset_keys}
            liab_pesos_vis = {k: float(-(sums_map.get(k, 0.0) or 0.0)) for k in liab_keys}
            pat_other_pesos_vis = {k: float(-(sums_map.get(k, 0.0) or 0.0)) for k in pat_other_keys}

            sum_assets_vis = sum(asset_pesos_vis.values())
            sum_liab_vis = sum(liab_pesos_vis.values())
            sum_pat_other_vis = sum(pat_other_pesos_vis.values())
            effective_pl_pesos = float(pl_unmapped_pesos) + pl_mapped_pesos
            res_acum_from_tb_vis = float(-(raw_res_acum_mapped + effective_pl_pesos))

            # Verificar ecuación fundamental en Pesos (Guardarraíl estricto <= $4.000)
            pl_cubo_sum = 0.0
            if (
                abs(total_tb_pesos) > MAX_TOLERANCE_PESOS
                and emp_name_col
                and periodo_col
                and not str(emp_name_col).startswith("[GRUPO]")
            ):
                try:
                    from src.models.pl_cubo_db import PlCuboDB
                    pl_cubo_sum = float(PlCuboDB.get_pl_cubo_total_sum(str(emp_name_col).strip(), str(periodo_col).strip()) or 0.0)
                except Exception:
                    pl_cubo_sum = 0.0

            if abs(total_tb_pesos) <= MAX_TOLERANCE_PESOS + 1e-6:
                tb_balanced_ok = True
                true_res_acum_vis_pesos = res_acum_from_tb_vis
            elif pl_cubo_sum != 0.0 and abs(total_tb_pesos - pl_cubo_sum) <= MAX_TOLERANCE_PESOS + 1e-6:
                tb_balanced_ok = True
                true_res_acum_vis_pesos = sum_assets_vis - sum_liab_vis - sum_pat_other_vis
            else:
                tb_balanced_ok = False
                true_res_acum_vis_pesos = res_acum_from_tb_vis

            missing_bs_ok = abs(missing_bs_pesos) <= MAX_TOLERANCE_PESOS + 1e-6
            guardrail_ok = tb_balanced_ok and missing_bs_ok

            # Si supera el guardarraíl de $4.000, NO forzar cuadratura en Resultados Acumulados para exponer el descuadre real
            if not guardrail_ok:
                raw_lines = {}
                for k in asset_keys:
                    raw_lines[k] = int(round(asset_pesos_vis[k] / scale_factor))
                for k in liab_keys:
                    raw_lines[k] = int(round(liab_pesos_vis[k] / scale_factor))
                for k in pat_other_keys:
                    raw_lines[k] = int(round(pat_other_pesos_vis[k] / scale_factor))
                return raw_lines, int(round(true_res_acum_vis_pesos / scale_factor))

            # Escala distinta de M$ (ej. Pesos exactos scale_factor == 1.0)
            if abs(float(scale_factor) - 1000.0) >= 0.01:
                pesos_lines = {}
                for k in asset_keys:
                    pesos_lines[k] = int(round(asset_pesos_vis[k] / scale_factor))
                for k in liab_keys:
                    pesos_lines[k] = int(round(liab_pesos_vis[k] / scale_factor))
                for k in pat_other_keys:
                    pesos_lines[k] = int(round(pat_other_pesos_vis[k] / scale_factor))
                sum_a_s = sum(pesos_lines[k] for k in asset_keys)
                sum_l_s = sum(pesos_lines[k] for k in liab_keys)
                sum_p_s = sum(pesos_lines[k] for k in pat_other_keys)
                return pesos_lines, int(round(sum_a_s - sum_l_s - sum_p_s))

            # Escala M$ (scale_factor == 1000.0): Redondeo Jerárquico por Mayor Residuo
            target_assets_m = int(round(sum_assets_vis / 1000.0))
            locked_assets = {
                k for k in asset_keys
                if sanitize(k) in {"efectivo y efectivo equivalente", "plusvalia", "inversion en empresas relacionadas"}
            }
            rounded_assets, ok_a = largest_remainder_round(
                asset_pesos_vis,
                target_total_m=target_assets_m,
                scale_factor=1000.0,
                max_tol_pesos=MAX_TOLERANCE_PESOS,
                locked_keys=locked_assets,
            )
            if not ok_a:
                cash_only = {k for k in asset_keys if sanitize(k) == "efectivo y efectivo equivalente"}
                rounded_assets, ok_a = largest_remainder_round(
                    asset_pesos_vis,
                    target_total_m=target_assets_m,
                    scale_factor=1000.0,
                    max_tol_pesos=MAX_TOLERANCE_PESOS,
                    locked_keys=cash_only,
                )
            if not ok_a:
                rounded_assets, _ = largest_remainder_round(
                    asset_pesos_vis,
                    target_total_m=target_assets_m,
                    scale_factor=1000.0,
                    max_tol_pesos=MAX_TOLERANCE_PESOS,
                )

            rounded_pat_other = {k: int(round(pat_other_pesos_vis[k] / 1000.0)) for k in pat_other_keys}

            if (
                expected_res_acum_m is not None
                and abs(float(expected_res_acum_m) * 1000.0 - true_res_acum_vis_pesos) <= MAX_TOLERANCE_PESOS + 1e-6
            ):
                res_acum_m = int(expected_res_acum_m)
            else:
                res_acum_m = int(round(true_res_acum_vis_pesos / 1000.0))

            pat_total_m = sum(rounded_pat_other.values()) + res_acum_m
            target_liab_m = int(target_assets_m - pat_total_m)

            rounded_liab, ok_l = largest_remainder_round(
                liab_pesos_vis,
                target_total_m=target_liab_m,
                scale_factor=1000.0,
                max_tol_pesos=MAX_TOLERANCE_PESOS,
            )
            if not ok_l:
                if (
                    all(abs(v) < 1e-9 for v in liab_pesos_vis.values())
                    and abs(sum_assets_vis - float(pat_total_m) * 1000.0) <= MAX_TOLERANCE_PESOS + 1e-6
                ):
                    alt_assets, ok_alt = largest_remainder_round(
                        asset_pesos_vis,
                        target_total_m=pat_total_m,
                        scale_factor=1000.0,
                        max_tol_pesos=MAX_TOLERANCE_PESOS,
                        locked_keys=locked_assets,
                    )
                    if ok_alt:
                        rounded_assets = alt_assets
                        target_assets_m = pat_total_m
                        rounded_liab = {k: 0 for k in liab_keys}
                        ok_l = True
                if not ok_l and abs(target_assets_m - sum(rounded_liab.values()) - pat_total_m) <= MAX_TOLERANCE_M:
                    res_acum_m = int(target_assets_m - sum(rounded_liab.values()) - sum(rounded_pat_other.values()))

            merged_lines = {}
            merged_lines.update(rounded_assets)
            merged_lines.update(rounded_liab)
            merged_lines.update(rounded_pat_other)
            return merged_lines, res_acum_m

        # 1. Resolver primero columna comparativa (si existe tb_df_comp)
        comp_line_vals_m = {}
        comp_res_acum_m = 0
        if has_comp and tb_df_comp is not None and not tb_df_comp.empty:
            comp_line_vals_m, comp_res_acum_m = _solve_bs_column(
                sums_comp_dict,
                pl_residual_comp,
                tb_comp_total_pesos,
                emp_inferred,
                periodo_comp_str,
                expected_res_acum_m=None,
            )

        # 2. Calcular meta de Resultados Acumulados alineada con #PATRIMONIO y #ER (si aplica en M$)
        expected_res_acum_act_m = None
        if abs(float(scale_factor) - 1000.0) < 0.01 and periodo_actual_str and tb_df is not None and not tb_df.empty:
            try:
                from src.reporting.patrimonio_generator import get_patrimonio_balances_from_db, get_prior_december_period
                from src.core.snapshot_vault_manager import SnapshotVaultManager

                prior_dec = get_prior_december_period(str(periodo_actual_str))
                gan_ini_m = None
                emp_candidates = []
                if empresa:
                    emp_candidates.append(str(empresa).strip())
                if emp_inferred:
                    if emp_inferred not in emp_candidates:
                        emp_candidates.append(emp_inferred)
                    grp_cand = f"[GRUPO] {emp_inferred}"
                    if grp_cand not in emp_candidates:
                        emp_candidates.append(grp_cand)

                if prior_dec:
                    for ec in emp_candidates:
                        if SnapshotVaultManager.has_active_snapshot(ec, prior_dec):
                            snap_b = SnapshotVaultManager.get_statement(ec, prior_dec, "#BALANCE", scale_factor=1000.0)
                            if snap_b and snap_b[0] is not None:
                                v_snap_ra, ok_snap_ra = _extract_stmt_target(
                                    snap_b[0],
                                    prior_dec,
                                    ["Resultados acumulados", "Ganancias (pérdidas) acumuladas", "Ganancias acumuladas"],
                                )
                                if ok_snap_ra:
                                    gan_ini_m = v_snap_ra
                                    break

                if gan_ini_m is None and has_comp and tb_df_comp is not None and not tb_df_comp.empty and prior_dec and str(periodo_comp_str) == prior_dec:
                    gan_ini_m = comp_res_acum_m
                elif gan_ini_m is None and prior_dec and emp_candidates:
                    for ec in emp_candidates:
                        _, g_db, _ = get_patrimonio_balances_from_db(ec, prior_dec)
                        if g_db != 0.0:
                            gan_ini_m = int(round(g_db))
                            break

                resolved_er_m = None
                if er_net_m is not None:
                    resolved_er_m = int(round(er_net_m))
                elif emp_inferred and not str(emp_inferred).startswith("[GRUPO]"):
                    try:
                        from src.models.pl_cubo_db import PlCuboDB
                        sum_p = float(PlCuboDB.get_pl_cubo_total_sum(emp_inferred, str(periodo_actual_str)) or 0.0)
                        if sum_p != 0.0:
                            resolved_er_m = int(round(-sum_p / 1000.0))
                    except Exception:
                        pass

                if gan_ini_m is not None and resolved_er_m is not None:
                    expected_res_acum_act_m = int(gan_ini_m + resolved_er_m)
            except Exception:
                expected_res_acum_act_m = None

        # 3. Resolver columna actual
        act_line_vals_m = {}
        act_res_acum_m = 0
        if tb_df is not None and not tb_df.empty:
            act_line_vals_m, act_res_acum_m = _solve_bs_column(
                sums_dict,
                pl_residual,
                tb_total_pesos,
                emp_inferred,
                periodo_actual_str,
                expected_res_acum_m=expected_res_acum_act_m,
            )

        # Trackers for 2025 totals
        subtotal_corriente = 0
        subtotal_no_corriente = 0
        activos_totales = 0
        pasivos_totales = 0
        patrimonio_controladora = 0
        patrimonio_no_controladora = 0
        patrimonio_totales = 0

        # Trackers for 2024 totals
        subtotal_corriente_24 = 0
        subtotal_no_corriente_24 = 0
        activos_totales_24 = 0
        pasivos_totales_24 = 0
        patrimonio_controladora_24 = 0
        patrimonio_no_controladora_24 = 0
        patrimonio_totales_24 = 0

        current_section = "None"

        for row in range(1, ws.max_row + 1):
            cell_name = ws.cell(row=row, column=name_col_idx).value

            if cell_name and isinstance(cell_name, str):
                cell_name_clean = cell_name.strip()
                cell_lower = sanitize(cell_name)

                if cell_lower == "activos corrientes":
                    current_section = "Activos Corrientes"
                elif cell_lower == "activos no corrientes":
                    current_section = "Activos No Corrientes"
                elif cell_lower == "pasivos corrientes":
                    current_section = "Pasivos Corrientes"
                elif cell_lower == "pasivos no corrientes":
                    current_section = "Pasivos No Corrientes"
                elif cell_lower == "patrimonio":
                    current_section = "Patrimonio"

                val24 = 0
                if has_comp and tb_df_comp is not None and not tb_df_comp.empty:
                    if cell_name_clean in sums_comp_dict and cell_name_clean not in res_acum_labels:
                        val24 = int(comp_line_vals_m.get(cell_name_clean, 0))
                        ws.cell(row=row, column=val24_col_idx).value = val24
                else:
                    val24_dirty = ws.cell(row=row, column=val24_col_idx).value
                    if val24_dirty is not None:
                        try:
                            clean_str_val = str(val24_dirty).replace(',', '').replace(' ', '').strip()
                            if clean_str_val:
                                val24 = round(float(clean_str_val))
                        except (ValueError, TypeError):
                            val24 = 0

                if cell_name_clean in sums_dict and cell_name_clean not in res_acum_labels:
                    visual_val = int(act_line_vals_m.get(cell_name_clean, 0))
                    ws.cell(row=row, column=val25_col_idx).value = visual_val

                    if current_section == "Activos Corrientes":
                        subtotal_corriente += visual_val
                    elif current_section == "Activos No Corrientes":
                        subtotal_no_corriente += visual_val
                    elif current_section == "Pasivos Corrientes":
                        subtotal_corriente += visual_val
                    elif current_section == "Pasivos No Corrientes":
                        subtotal_no_corriente += visual_val
                    elif current_section == "Patrimonio":
                        if _is_participaciones_no_controladoras(cell_lower):
                            patrimonio_no_controladora += visual_val
                        else:
                            patrimonio_controladora += visual_val
                        patrimonio_totales += visual_val

                if cell_name_clean in res_acum_labels:
                    if tb_df is not None and not tb_df.empty:
                        visual_final_resid = int(act_res_acum_m)
                        ws.cell(row=row, column=val25_col_idx).value = visual_final_resid
                        patrimonio_controladora += visual_final_resid
                        patrimonio_totales += visual_final_resid

                    if has_comp and tb_df_comp is not None and not tb_df_comp.empty:
                        val24 = int(comp_res_acum_m)
                        ws.cell(row=row, column=val24_col_idx).value = val24
                        patrimonio_controladora_24 += val24
                        patrimonio_totales_24 += val24

                _is_any_total = _is_any_total_fn(cell_lower)

                if not _is_any_total and cell_name_clean not in res_acum_labels:
                    val24_round = round(val24)
                    if current_section == "Activos Corrientes":
                        subtotal_corriente_24 += val24_round
                    elif current_section == "Activos No Corrientes":
                        subtotal_no_corriente_24 += val24_round
                    elif current_section == "Pasivos Corrientes":
                        subtotal_corriente_24 += val24_round
                    elif current_section == "Pasivos No Corrientes":
                        subtotal_no_corriente_24 += val24_round
                    elif current_section == "Patrimonio":
                        if _is_participaciones_no_controladoras(cell_lower):
                            patrimonio_no_controladora_24 += val24_round
                        else:
                            patrimonio_controladora_24 += val24_round
                        patrimonio_totales_24 += val24_round

                if _is_total_activos_corrientes(cell_lower):
                    ws.cell(row=row, column=val25_col_idx).value = round(subtotal_corriente)
                    if has_comp:
                        ws.cell(row=row, column=val24_col_idx).value = round(subtotal_corriente_24)
                    activos_totales += round(subtotal_corriente)
                    activos_totales_24 += round(subtotal_corriente_24)
                    subtotal_corriente = 0
                    subtotal_corriente_24 = 0
                elif _is_total_activos_no_corrientes(cell_lower):
                    ws.cell(row=row, column=val25_col_idx).value = round(subtotal_no_corriente)
                    if has_comp:
                        ws.cell(row=row, column=val24_col_idx).value = round(subtotal_no_corriente_24)
                    activos_totales += round(subtotal_no_corriente)
                    activos_totales_24 += round(subtotal_no_corriente_24)
                    subtotal_no_corriente = 0
                    subtotal_no_corriente_24 = 0
                elif _is_total_activos(cell_lower):
                    ws.cell(row=row, column=val25_col_idx).value = round(activos_totales)
                    if has_comp:
                        ws.cell(row=row, column=val24_col_idx).value = round(activos_totales_24)
                elif _is_total_pasivos_corrientes(cell_lower):
                    ws.cell(row=row, column=val25_col_idx).value = round(subtotal_corriente)
                    if has_comp:
                        ws.cell(row=row, column=val24_col_idx).value = round(subtotal_corriente_24)
                    pasivos_totales += round(subtotal_corriente)
                    pasivos_totales_24 += round(subtotal_corriente_24)
                    subtotal_corriente = 0
                    subtotal_corriente_24 = 0
                elif _is_total_pasivos_no_corrientes(cell_lower):
                    ws.cell(row=row, column=val25_col_idx).value = round(subtotal_no_corriente)
                    if has_comp:
                        ws.cell(row=row, column=val24_col_idx).value = round(subtotal_no_corriente_24)
                    pasivos_totales += round(subtotal_no_corriente)
                    pasivos_totales_24 += round(subtotal_no_corriente_24)
                    subtotal_no_corriente = 0
                    subtotal_no_corriente_24 = 0
                elif _is_total_pasivos(cell_lower):
                    ws.cell(row=row, column=val25_col_idx).value = round(pasivos_totales)
                    if has_comp:
                        ws.cell(row=row, column=val24_col_idx).value = round(pasivos_totales_24)
                elif _is_patrimonio_controladora(cell_lower):
                    ws.cell(row=row, column=val25_col_idx).value = round(patrimonio_controladora)
                    if has_comp:
                        ws.cell(row=row, column=val24_col_idx).value = round(patrimonio_controladora_24)
                elif _is_total_pasivos_patrimonio(cell_lower):
                    ws.cell(row=row, column=val25_col_idx).value = round(pasivos_totales + patrimonio_totales)
                    if has_comp:
                        ws.cell(row=row, column=val24_col_idx).value = round(pasivos_totales_24 + patrimonio_totales_24)
                elif _is_total_patrimonio(cell_lower):
                    ws.cell(row=row, column=val25_col_idx).value = round(patrimonio_totales)
                    if has_comp:
                        ws.cell(row=row, column=val24_col_idx).value = round(patrimonio_totales_24)

                if str(target_lang).lower() == 'en':
                    ws.cell(row=row, column=name_col_idx).value = translate_ifrs_term(cell_name, target_lang='en', overrides_dict=overrides)

        output = BytesIO()
        wb.save(output)
        output.seek(0)

        try:
            from src.ui_pages.informes_y_notas import evaluate_formulas_in_workbook
            output = evaluate_formulas_in_workbook(output)
        except Exception:
            pass

        return output
