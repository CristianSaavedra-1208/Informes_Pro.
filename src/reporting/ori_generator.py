import pandas as pd
import openpyxl
from io import BytesIO

from src.core.ifrs_glossary import translate_ifrs_term

class OriGenerator:
    def __init__(self, template_path):
        self.template_path = template_path

    def generate(self, pl_preview_df=None, periodo_actual_str=None, periodo_comp_str=None, bal_preview_df=None, empresa=None, target_lang='es'):
        """
        Genera el formato de Resultados Integrales (ORI) inyectando la ganancia proveniente del ER y sumando líneas de Otros Resultados Integrales.
        pl_preview_df: DataFrame generado por ERGenerator (con 'Clasificación', col_actual, col_comp)
        bal_preview_df: DataFrame generado por BalanceGenerator (con 'Clasificación', col_actual, col_comp)
        """
        col_actual = periodo_actual_str if periodo_actual_str else '2025'
        col_comp = periodo_comp_str if periodo_comp_str else '2024'
        
        pl_25 = 0.0
        pl_24 = 0.0
        
        # Extraer Ganancia del Ejercicio (P&L Base) usando un extractor robusto
        if pl_preview_df is not None and not pl_preview_df.empty:
            cols = [str(c).strip() for c in pl_preview_df.columns]
            col_act_use = col_actual if col_actual in cols else (cols[1] if len(cols) > 1 else None)
            col_comp_use = col_comp if col_comp in cols else (cols[2] if len(cols) > 2 else None)
            
            # Prioridad 1: Coincidencia precisa con la ganancia/pérdida final del ejercicio
            target_exact_keywords = [
                'ganancia (pérdida) del ejercicio', 'ganancias (pérdida) del ejercicio',
                'ganancia (perdida) del ejercicio', 'ganancias (perdidas) del ejercicio',
                'resultado del ejercicio', 'utilidad del ejercicio', 'pérdida del ejercicio',
                'ganancia neta', 'utilidad neta', 'pérdida neta',
                'ganancia (pérdida) procedente de operaciones continuadas',
                'resultado de operaciones continuadas'
            ]
            exclude_keywords = [
                'unidad de reajuste', 'reajuste', 'financiero', 'financieros',
                'antes de impuesto', 'antes de impuestos', 'bruta', 'bruto', 'operacional'
            ]
            
            matched_row = None
            for idx, row in pl_preview_df.iterrows():
                clasif = str(row.iloc[0]).lower().strip()
                if any(k in clasif for k in target_exact_keywords) and not any(ex in clasif for ex in exclude_keywords):
                    matched_row = row
                    break
            
            # Fallback si no hubo coincidencia exacta
            if matched_row is None:
                for idx, row in pl_preview_df.iterrows():
                    clasif = str(row.iloc[0]).lower().strip()
                    if not any(ex in clasif for ex in exclude_keywords):
                        if 'continuadas' in clasif or ('ganancia' in clasif and 'impuesto' not in clasif):
                            matched_row = row
                            break

            if matched_row is not None:
                if col_act_use:
                    v25 = matched_row[col_act_use]
                    pl_25 = float(v25) if pd.notna(v25) and str(v25).strip() != '' else 0.0
                if col_comp_use and periodo_comp_str and str(periodo_comp_str).strip().lower() != "ninguno":
                    v24 = matched_row[col_comp_use]
                    pl_24 = float(v24) if pd.notna(v24) and str(v24).strip() != '' else 0.0
        
        # ------------------ RESOLVER SALDOS DE OTRAS RESERVAS (ALINEADO 100% CON PATRIMONIO) ------------------
        from src.reporting.patrimonio_generator import (
            get_prior_december_period,
            get_patrimonio_balances_from_db,
            parse_numeric_value,
        )

        res_ext_actual = None
        res_ext_comp = None

        # Intentar extraer de bal_preview_df soportando cualquier columna de clasificación y formato numérico
        def get_bal_val(name, col):
            if bal_preview_df is None or bal_preview_df.empty or not col:
                return None
            try:
                target_col = col if col in bal_preview_df.columns else None
                if not target_col:
                    for c in bal_preview_df.columns[1:]:
                        if str(col) in str(c) or (len(str(col)) >= 4 and str(col)[:4] in str(c)):
                            target_col = c
                            break
                if not target_col:
                    return None
                clasif_col = bal_preview_df.columns[0]
                row = bal_preview_df[bal_preview_df[clasif_col].astype(str).str.strip().str.lower() == name.lower()]
                if not row.empty:
                    v = row[target_col].iloc[0]
                    return parse_numeric_value(v)
            except Exception:
                pass
            return None

        if bal_preview_df is not None and not bal_preview_df.empty:
            res_ext_actual = get_bal_val("Otras reservas", col_actual)
            res_ext_comp = get_bal_val("Otras reservas", col_comp)

        if res_ext_actual is None and empresa and col_actual:
            _, _, r_act = get_patrimonio_balances_from_db(empresa, str(col_actual))
            res_ext_actual = r_act
        if res_ext_comp is None and empresa and col_comp:
            _, _, r_cmp = get_patrimonio_balances_from_db(empresa, str(col_comp))
            res_ext_comp = r_cmp

        res_ext_actual = res_ext_actual if res_ext_actual is not None else 0.0
        res_ext_comp = res_ext_comp if res_ext_comp is not None else 0.0

        # 1. Saldo inicial de reservas para el ejercicio actual (cierre diciembre año anterior)
        periodo_ini_actual = get_prior_december_period(str(col_actual))
        res_ini_actual = 0.0
        if col_comp and periodo_ini_actual and (periodo_ini_actual in str(col_comp) or str(col_comp).startswith(periodo_ini_actual[:4])):
            res_ini_actual = res_ext_comp
        elif periodo_ini_actual and empresa:
            cap_i, gan_i, res_i = get_patrimonio_balances_from_db(empresa, periodo_ini_actual)
            if cap_i == 0.0 and gan_i == 0.0 and res_i == 0.0:
                res_ini_actual = res_ext_comp
            else:
                res_ini_actual = res_i
        else:
            res_ini_actual = res_ext_comp

        # 2. Saldo inicial de reservas para el ejercicio comparativo (cierre diciembre año previo al comparativo)
        periodo_ini_comp = get_prior_december_period(str(col_comp))
        res_ini_comp = res_ext_comp
        if periodo_ini_comp and empresa:
            cap_ic, gan_ic, res_ic = get_patrimonio_balances_from_db(empresa, periodo_ini_comp)
            if not (cap_ic == 0.0 and gan_ic == 0.0 and res_ic == 0.0):
                res_ini_comp = res_ic

        # Calcular variaciones netas de reservas (Net ORI), 100% consistente con Estado de Cambios en el Patrimonio
        net_ori_25_raw = res_ext_actual - res_ini_actual
        net_ori_24_raw = res_ext_comp - res_ini_comp

        # Calcular efectos brutos e impuestos (Tasa del 27%)
        # Neto = Bruto * 0.73  =>  Bruto = Neto / 0.73
        # Impuesto = -0.27 * Bruto
        gross_ori_25 = round(net_ori_25_raw / 0.73)
        tax_ori_25 = round(-0.27 * gross_ori_25)
        net_ori_25 = gross_ori_25 + tax_ori_25

        gross_ori_24 = round(net_ori_24_raw / 0.73)
        tax_ori_24 = round(-0.27 * gross_ori_24)
        net_ori_24 = gross_ori_24 + tax_ori_24

        total_integral_25 = round(pl_25) + net_ori_25
        total_integral_24 = round(pl_24) + net_ori_24
        
        # Inyectando a Excel Abierto
        wb = openpyxl.load_workbook(self.template_path)
        ws = wb.active
        
        # Encontrar columna de conceptos/nombres y de valores de forma dinámica
        name_col_idx = 1
        for col in range(1, 5):
            for row in range(1, 15):
                val = ws.cell(row=row, column=col).value
                if val and str(val).strip().lower() in ["concepto", "descripcion", "detalle", "flujos", "origen/aplicacion", "clasificacion", "otros resultados"]:
                    name_col_idx = col
                    break
                    
        date_cols = []
        import datetime
        import re
        for col in range(name_col_idx + 1, ws.max_column + 1):
            for row in range(1, 10):
                val = ws.cell(row=row, column=col).value
                if val is not None:
                    val_str = str(val).strip().lower()
                    is_date = (
                        isinstance(val, (datetime.datetime, datetime.date)) or
                        re.search(r'20\d{2}', val_str) or
                        any(k in val_str for k in ["actual", "anterior", "comparativ", "auditado", "31-12", "31 de"])
                    )
                    if is_date:
                        date_cols.append(col)
                        break
        date_cols = sorted(list(set(date_cols)))
        
        val25_col_idx = date_cols[0] if len(date_cols) >= 1 else 2
        val24_col_idx = date_cols[1] if len(date_cols) >= 2 else 3
        
        # Row map based on our template script
        row_ganancia_base = 8
        row_coberturas = 12
        row_ori_pre_imp = 14
        row_imp_ori = 16
        row_sub_ori = 18
        row_total_integral = 19
        
        for r in range(1, ws.max_row + 1):
            val = ws.cell(row=r, column=name_col_idx).value
            if val and isinstance(val, str):
                val_clean = val.lower().strip()
                if any(k in val_clean for k in ["operaciones continuadas", "ganancia (perdida) del ejercicio", "utilidad (perdida) del ejercicio", "ganancia del ejercicio"]) and "componentes" not in val_clean and "total" not in val_clean:
                    row_ganancia_base = r
                elif "coberturas de flujo de efectivo" in val_clean or "coberturas" in val_clean:
                    row_coberturas = r
                elif "antes de impuestos" in val_clean and ("otros" in val_clean or "otro resultado integral" in val_clean):
                    row_ori_pre_imp = r
                elif "impuesto" in val_clean and ("otros" in val_clean or "otro resultado integral" in val_clean or r > row_ori_pre_imp):
                    row_imp_ori = r
                elif "total de otros resultados integrales" in val_clean or ("neto de impuestos" in val_clean and "otros" in val_clean) or val_clean in ["otro resultado integral", "otros resultados integrales"]:
                    row_sub_ori = r
                elif "resultado integral total" in val_clean or "total resultado integral" in val_clean or "total de resultado integral" in val_clean:
                    row_total_integral = r
        
        def inyectar(r, v25, v24):
            ws.cell(row=r, column=val25_col_idx, value=v25)
            if periodo_comp_str and str(periodo_comp_str).strip().lower() != "ninguno":
                ws.cell(row=r, column=val24_col_idx, value=v24)
            
        inyectar(row_ganancia_base, pl_25, pl_24)
        inyectar(row_coberturas, gross_ori_25, gross_ori_24)
        inyectar(row_ori_pre_imp, gross_ori_25, gross_ori_24)
        inyectar(row_imp_ori, tax_ori_25, tax_ori_24)
        inyectar(row_sub_ori, net_ori_25, net_ori_24)
        inyectar(row_total_integral, total_integral_25, total_integral_24)
 
        if periodo_actual_str and periodo_comp_str:
            for r in range(1, 10):
                c25 = ws.cell(row=r, column=val25_col_idx)
                c24 = ws.cell(row=r, column=val24_col_idx)
                if c25.value and isinstance(c25.value, str) and ("20" in c25.value or "Actual" in c25.value):
                    c25.value = periodo_actual_str
                if c24.value and isinstance(c24.value, str) and any(k in c24.value for k in ["20", "Anterior", "Comparativo", "comparativo"]):
                    c24.value = periodo_comp_str
        
        if str(target_lang).lower() == 'en':
            for r in range(1, ws.max_row + 1):
                cell = ws.cell(row=r, column=name_col_idx)
                if cell.value and isinstance(cell.value, str):
                    cell.value = translate_ifrs_term(cell.value.strip(), target_lang='en')
        
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        try:
            from src.ui_pages.informes_y_notas import evaluate_formulas_in_workbook
            output = evaluate_formulas_in_workbook(output)
        except Exception:
            pass
            
        return output
