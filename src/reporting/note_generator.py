import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
import unicodedata
import re
import datetime
import calendar

def normalize_company_name(name):
    if not name:
        return ""
    n = name.strip().lower()
    if 'pacifico' in n:
        return 'pacifico'
    if 'parent' in n:
        return 'parent'
    if 'holdco' in n or 'holco' in n:
        return 'holdco'
    if 'consolidado' in n or 'grupo' in n:
        return 'consolidado'
    return n

def clean_label(text):
    if text is None:
        return ""
    text_str = str(text).strip().lower()
    text_str = ''.join(c for c in unicodedata.normalize('NFD', text_str) if unicodedata.category(c) != 'Mn')
    return " ".join(text_str.split())

def determine_factor_for_accounts(accounts):
    factors = []
    for acc in accounts:
        acc_str = str(acc).strip()
        if acc_str.startswith('2') or acc_str.startswith('3101'):
            factors.append(-1)
        else:
            factors.append(1)
    # Si la mayoría o al menos la mitad de las cuentas son acreedoras (pasivo/patrimonio), el factor global es -1
    neg_count = sum(1 for f in factors if f == -1)
    if neg_count >= len(factors) / 2:
        return -1
    return 1

def check_extra_word_match(ew, acct_names):
    if not ew or ew in ['$', '%', 'm$', 'm', 'detalle', 'total']:
        return True
    
    acct_names_clean = [clean_label(n) for n in acct_names]
    ew_tokens = ew.split()
    
    keywords = ['bci', 'scotiabank', 'santander', 'estado', 'itau', 'chile', 'bice', 'security']
    for kw in keywords:
        if kw in ew_tokens:
            if not any(kw in name for name in acct_names_clean):
                return False
                
    return True

def detect_row_corriente_status(ws, row, col=None):
    # 1. Si se especifica columna, verificar primero las cabeceras de esa columna dentro del bloque actual
    if col is not None:
        for r in range(row - 1, 0, -1):
            # Detenerse si cruzamos el límite de la tabla actual (#Nxx.y o Total)
            is_boundary = False
            for c_chk in range(1, min(6, ws.max_column + 1)):
                v_chk = ws.cell(row=r, column=c_chk).value
                if v_chk is not None:
                    s_chk = str(v_chk).strip()
                    if re.search(r'#N\d+\.\d+', s_chk) or (s_chk.lower().startswith("total") and r < row - 1):
                        is_boundary = True
                        break
            for c_hdr in ([col, col - 1] if col > 1 else [col]):
                if c_hdr != col and ws.cell(row=r, column=col).value is not None:
                    continue
                vh = ws.cell(row=r, column=c_hdr).value
                if vh is not None and isinstance(vh, str):
                    vh_low = vh.strip().lower()
                    if len(vh_low) <= 35:
                        if "no corriente" in vh_low or "no-corriente" in vh_low or "largo plazo" in vh_low:
                            return "no_corriente"
                        elif "corriente" in vh_low or "corto plazo" in vh_low:
                            return "corriente"
            if is_boundary:
                break

    # 2. Escanear hacia atrás las celdas en Col 1 y 2 dentro del bloque actual
    for r in range(row, 0, -1):
        for c in [1, 2]:
            val = ws.cell(row=r, column=c).value
            if val is not None:
                val_str = str(val).strip().lower()
                if r < row and (re.search(r'#n\d+\.\d+', val_str) or val_str.startswith("total")):
                    return "unknown"
                # Ignorar párrafos narrativos largos (ej: "La composición del rubro de cuentas por cobrar corrientes...")
                if len(val_str) > 45:
                    continue
                if any(kw in val_str for kw in ["moneda", "vencimiento", "vencimientos"]):
                    return "unknown"
                if "no corriente" in val_str or "no-corriente" in val_str or "largo plazo" in val_str:
                    return "no_corriente"
                elif "corriente" in val_str:
                    return "corriente"
    return "unknown"

SHEET_TO_PL_RUBROS = {
    "Ingresos Ctos operacion": [
        "ingresos de actividades ordinarias",
        "costo de ventas",
        "acceso a infraestructura fibra optica",
        "costos de uso fibra optica",
        "depreciacion operacional"
    ],
    "Gtos Adm": [
        "gastos de administracion",
        "depreciacion y amortizaciones"
    ],
    "DC y Reajustes": [
        "diferencias de cambio",
        "resultados por unidades de reajuste"
    ],
    "Costos e ingresos Financieros": [
        "ingresos financieros",
        "costos financieros"
    ],
    "Otros ingresos por funcion": [
        "otros ingresos por funcion"
    ],
    "Otros gastos por funcion": [
        "otros egresos por funcion"
    ]
}

SECTION_TITLE_TO_PL_RUBRO = {
    "diferencia de cambios": "diferencias de cambio",
    "diferencia de cambio": "diferencias de cambio",
    "diferencias de cambio": "diferencias de cambio",
    "resultados por unidad de reajustes": "resultados por unidades de reajuste",
    "resultados por unidad de reajuste": "resultados por unidades de reajuste",
    "resultado por unidad de reajuste": "resultados por unidades de reajuste",
    "resultados por unidades de reajuste": "resultados por unidades de reajuste",
    "ingresos de la operacion": "ingresos de actividades ordinarias",
    "ingresos de la explotacion": "ingresos de actividades ordinarias",
    "ingresos de actividades ordinarias": "ingresos de actividades ordinarias",
    "costos de la operacion": "costo de ventas",
    "costo de ventas": "costo de ventas",
    "costos financieros": "costos financieros",
    "ingresos financieros": "ingresos financieros",
    "gastos de administracion": "gastos de administracion",
    "depreciacion y amortizaciones": "depreciacion y amortizaciones",
    "depreciacion operacional": "depreciacion operacional",
    "otros gastos por funcion": "otros egresos por funcion",
    "otros egresos por funcion": "otros egresos por funcion",
    "otros ingresos por funcion": "otros ingresos por funcion",
    "arriendo de fibra ic": "acceso a infraestructura fibra optica",
    "arriendo de fibra": "acceso a infraestructura fibra optica",
    "costos uso fibra optica": "acceso a infraestructura fibra optica",
    "costos de uso fibra optica": "acceso a infraestructura fibra optica",
    "acceso a infraestructura fibra optica": "acceso a infraestructura fibra optica",
    "acceso a infraestructura fibra": "acceso a infraestructura fibra optica",
}

def shift_formula_rows(formula_str, row_shift, min_row=None, max_row=None):
    if not isinstance(formula_str, str) or not formula_str.startswith('='):
        return formula_str
    def replace_cell_ref(match):
        col = match.group(1)
        row = int(match.group(2))
        if min_row is not None and row < min_row:
            return match.group(0)
        if max_row is not None and row > max_row:
            return match.group(0)
        return f"{col}{row + row_shift}"
    return re.sub(r'([A-Za-z]+)(\d+)', replace_cell_ref, formula_str)

def _norm_movement_concept(txt):
    if not txt:
        return ""
    import unicodedata
    s = unicodedata.normalize('NFKD', str(txt)).encode('ASCII', 'ignore').decode('utf-8').lower()
    return re.sub(r'[^a-z0-9]', '', s)

def _get_movement_concept_key(norm_lbl):
    if 'saldoinicial' in norm_lbl or 'al01' in norm_lbl:
        return 'saldo_inicial'
    elif 'adicion' in norm_lbl or 'adquisici' in norm_lbl:
        return 'adiciones'
    elif 'desapropia' in norm_lbl or 'enajena' in norm_lbl or 'disposal' in norm_lbl:
        return 'desapropiaciones'
    elif 'baja' in norm_lbl or 'retiro' in norm_lbl:
        return 'bajas'
    elif 'depreciaci' in norm_lbl:
        return 'depreciacion'
    elif 'deterioro' in norm_lbl:
        return 'deterioro'
    elif 'transferencia' in norm_lbl:
        return 'transferencias'
    elif 'saldofinal' in norm_lbl or 'al31' in norm_lbl or 'al30' in norm_lbl:
        return 'saldo_final'
    return norm_lbl

def _is_movement_header_or_title_row(ws_obj, r, label_col=1):
    lbl_val = ws_obj.cell(r, label_col).value
    lbl = str(lbl_val).strip() if lbl_val is not None else ""
    lbl_low = lbl.lower()

    if re.search(r'#N\d+\.\d+', lbl) or "[comparativo]" in lbl_low:
        return True
    if re.match(r'^20\d{2}$', lbl):
        return True
    if lbl_low in ("tipo", "concepto", "detalle", "m$", "ms", "r.u.t.", "rut") or lbl_low.startswith("detalle al") or lbl_low.startswith("concepto al"):
        return True
    if any(k in lbl_low for k in ("vencimiento de pasivos", "cambios en pasivo", "transacciones significativas", "detalle leasing", "detalle prestamos", "bonos emitidos")):
        return True

    right_strs = []
    for c in range(label_col + 1, min(label_col + 15, ws_obj.max_column + 1)):
        v = ws_obj.cell(r, c).value
        if v is not None and str(v).strip() != "":
            right_strs.append(str(v).strip().lower())

    if not lbl:
        if not right_strs:
            return True
        if all(
            s in ("m$", "ms", "pagos") or
            any(h in s for h in ("edificios", "redes de", "plantas", "terrenos", "construcciones", "transmisión", "transmision", "equipos", "equipo", "vehículos", "vehiculos", "muebles", "instalación", "instalacion", "otros", "activos", "obras", "naturaleza de", "descripción de", "descripcion de"))
            for s in right_strs
        ):
            return True

    if right_strs and any(s in ("m$", "ms", "90 días", "90 dias", "nuevos contratos", "nuevos créditos", "nuevos creditos", "terrenos", "concesiones", "bruto", "nombre entidad") for s in right_strs):
        return True

    return False

def _is_movement_total_row(ws_obj, r, label_col=1):
    lbl_val = ws_obj.cell(r, label_col).value
    lbl_low = str(lbl_val).strip().lower() if lbl_val is not None else ""
    return any(lbl_low.startswith(k) for k in ("total", "sub total", "subtotal", "sub- total", "valores netos al final"))

def _extract_movement_table_rows(source_ws, start_r, label_col=1):
    concept_map = {}
    seen_data_rows = 0
    for r in range(start_r + 1, min(start_r + 35, source_ws.max_row + 1)):
        # Si aparece otro marcador #Nxx.y en cualquier columna, fin de la tabla
        hit_next_tag = False
        for c_chk in range(1, min(6, source_ws.max_column + 1)):
            v_chk = source_ws.cell(r, c_chk).value
            if v_chk and isinstance(v_chk, str) and re.search(r'#N\d+\.\d+', v_chk):
                hit_next_tag = True
                break
        if hit_next_tag:
            break

        if _is_movement_header_or_title_row(source_ws, r, label_col=label_col):
            continue

        if _is_movement_total_row(source_ws, r, label_col=label_col):
            if seen_data_rows > 0:
                break
            continue

        v = source_ws.cell(r, label_col).value
        label = str(v).strip() if (v is not None and isinstance(v, str)) else ""
        if not label:
            continue

        norm_lbl = _norm_movement_concept(label)
        key = _get_movement_concept_key(norm_lbl)
        if key:
            concept_map[key] = r
            seen_data_rows += 1
        if key == 'saldo_final' and seen_data_rows > 1:
            break
    return concept_map

def process_comparative_tables_in_ws(ws, periodo_actual_str="2026-Ago", periodo_comp_str="2025-Dic", ext_cmp_ws=None, ext_act_ws=None):
    if not periodo_comp_str or str(periodo_comp_str).strip() in ["None", "Ninguno", ""]:
        for r in range(1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                cell = ws.cell(r, c)
                if cell.value and isinstance(cell.value, str) and "[COMPARATIVO]" in cell.value.upper():
                    cleaned = re.sub(r'\[COMPARATIVO\]', '', cell.value, flags=re.IGNORECASE).strip()
                    cell.value = cleaned if cleaned else None
        return

    act_year = re.search(r'20\d{2}', str(periodo_actual_str))
    act_yr_str = act_year.group(0) if act_year else "2026"

    cmp_year = re.search(r'20\d{2}', str(periodo_comp_str))
    cmp_yr_str = cmp_year.group(0) if cmp_year else "2025"

    act_date_str = format_date_dots(periodo_actual_str)
    cmp_date_str = format_date_dots(periodo_comp_str)
    act_init_str = f"01.01.{act_yr_str}"
    cmp_init_str = f"01.01.{cmp_yr_str}"

    comp_locations = []
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cell_val = ws.cell(r, c).value
            if cell_val and isinstance(cell_val, str) and "[COMPARATIVO]" in cell_val.upper():
                comp_locations.append((r, c))

    for comp_r, comp_c in reversed(comp_locations):
        start_row = comp_r

        # Extraer tag de código (ej: #N09.2)
        marker_str = str(ws.cell(comp_r, comp_c).value or '')
        m_code = re.search(r'#N\d+\.\d+', marker_str)
        code_tag = m_code.group(0) if m_code else None

        # Localizar tabla en ext_cmp_ws si existe
        ext_cmp_r0 = None
        ext_cmp_c0 = None
        cmp_concept_map = {}
        if ext_cmp_ws and code_tag:
            for r_scan in range(1, ext_cmp_ws.max_row + 1):
                for c_scan in range(1, ext_cmp_ws.max_column + 1):
                    v_scan = ext_cmp_ws.cell(r_scan, c_scan).value
                    if v_scan and isinstance(v_scan, str) and code_tag in v_scan:
                        ext_cmp_r0 = r_scan
                        ext_cmp_c0 = c_scan
                        break
                if ext_cmp_r0 is not None:
                    break
            if ext_cmp_r0 is not None:
                cmp_concept_map = _extract_movement_table_rows(ext_cmp_ws, ext_cmp_r0, label_col=ext_cmp_c0)

        # Localizar tabla en ext_act_ws si existe
        ext_act_r0 = None
        ext_act_c0 = None
        act_concept_map = {}
        if ext_act_ws and code_tag:
            for r_scan in range(1, ext_act_ws.max_row + 1):
                for c_scan in range(1, ext_act_ws.max_column + 1):
                    v_scan = ext_act_ws.cell(r_scan, c_scan).value
                    if v_scan and isinstance(v_scan, str) and code_tag in v_scan:
                        ext_act_r0 = r_scan
                        ext_act_c0 = c_scan
                        break
                if ext_act_r0 is not None:
                    break
            if ext_act_r0 is not None:
                act_concept_map = _extract_movement_table_rows(ext_act_ws, ext_act_r0, label_col=ext_act_c0)

        act_col_offset = (ext_act_c0 - comp_c) if ext_act_c0 is not None else 0
        cmp_col_offset = (ext_cmp_c0 - comp_c) if ext_cmp_c0 is not None else 0

        end_row = start_row
        for r in range(start_row + 2, min(start_row + 35, ws.max_row + 1)):
            v_lbl = ws.cell(r, comp_c).value
            label_cell = str(v_lbl).strip().lower() if v_lbl is not None else ""
            if "saldo final" in label_cell or "valores netos al final" in label_cell or (label_cell.startswith("total") and "sub" not in label_cell):
                end_row = r
                break
        
        if end_row <= start_row:
            end_row = min(start_row + 15, ws.max_row)

        # Determinar la última columna real del cuadro en ws para no escribir ceros fuera de la tabla
        max_table_col = comp_c
        for r_scan in range(start_row, end_row + 1):
            for c_scan in range(comp_c + 1, ws.max_column + 1):
                if ws.cell(r_scan, c_scan).value is not None:
                    if c_scan > max_table_col:
                        max_table_col = c_scan

        # 1. Actualizar fechas e inyectar ext_act_ws en Bloque 1 (Actual)
        for r_orig in range(start_row, end_row + 1):
            is_hdr_or_title = _is_movement_header_or_title_row(ws, r_orig, label_col=comp_c)
            is_tot_row = _is_movement_total_row(ws, r_orig, label_col=comp_c)
            v_lbl_orig = ws.cell(r_orig, comp_c).value
            label_r = str(v_lbl_orig).strip() if (v_lbl_orig is not None and isinstance(v_lbl_orig, str)) else ""
            concept_k = _get_movement_concept_key(_norm_movement_concept(label_r)) if (not is_hdr_or_title and not is_tot_row and label_r) else None

            for c_orig in range(1, max_table_col + 1):
                cell_o = ws.cell(r_orig, c_orig)
                if isinstance(cell_o.value, (int, float)) and cell_o.value in (2024, 2025, 2026, 2027) and c_orig == comp_c:
                    cell_o.value = int(act_yr_str)
                elif cell_o.value and isinstance(cell_o.value, str):
                    val_low = cell_o.value.strip().lower()
                    if "saldo final" in val_low or "detalle al" in val_low or "concepto al" in val_low:
                        cell_o.value = re.sub(r'\b\d{2}[\.\/]\d{2}[\.\/]20\d{2}\b', act_date_str, cell_o.value)
                    elif "saldo inicial" in val_low:
                        cell_o.value = re.sub(r'\b\d{2}[\.\/]\d{2}[\.\/]20\d{2}\b', act_init_str, cell_o.value)
                    elif re.match(r'^20\d{2}$', val_low) and c_orig == comp_c:
                        cell_o.value = act_yr_str

            # Inyectar datos de ext_act_ws si existe para este concepto
            if concept_k and concept_k in act_concept_map and ext_act_ws is not None:
                ext_r = act_concept_map[concept_k]
                for c_data in range(comp_c + 1, max_table_col + 1):
                    target_c = ws.cell(r_orig, c_data)
                    if isinstance(target_c.value, str) and target_c.value.startswith('='):
                        continue
                    c_ext = c_data + act_col_offset
                    v_ext = ext_act_ws.cell(ext_r, c_ext).value if (1 <= c_ext <= ext_act_ws.max_column) else None
                    if v_ext is None or str(v_ext).strip() in ['', '-', '-   ']:
                        target_c.value = 0
                    elif isinstance(v_ext, (int, float)):
                        target_c.value = float(v_ext) if isinstance(v_ext, float) else int(v_ext)
                    elif isinstance(v_ext, str) and not v_ext.startswith('='):
                        try:
                            clean_s = v_ext.replace('.', '').replace(',', '.') if (',' in v_ext and '.' in v_ext) else v_ext.replace(',', '')
                            target_c.value = float(clean_s)
                        except Exception:
                            target_c.value = v_ext

        block_height = end_row - start_row + 1
        insert_at = end_row + 2
        row_shift = insert_at - start_row
        shift_amt = block_height + 1

        ws.insert_rows(insert_at, amount=shift_amt)

        # Desplazar las referencias de fila en todas las fórmulas que quedaron debajo del bloque insertado
        for r_below in range(insert_at + shift_amt, ws.max_row + 1):
            for c_below in range(1, ws.max_column + 1):
                val_below = ws.cell(r_below, c_below).value
                if isinstance(val_below, str) and val_below.startswith('='):
                    ws.cell(r_below, c_below).value = shift_formula_rows(val_below, shift_amt, min_row=insert_at)

        # 2. Construir el bloque comparativo (Block 2)
        for offset in range(block_height):
            orig_r = start_row + offset
            target_r = insert_at + offset

            is_hdr_or_title = _is_movement_header_or_title_row(ws, orig_r, label_col=comp_c)
            is_tot_row = _is_movement_total_row(ws, orig_r, label_col=comp_c)
            v_lbl_orig = ws.cell(orig_r, comp_c).value
            label_orig = str(v_lbl_orig).strip() if (v_lbl_orig is not None and isinstance(v_lbl_orig, str)) else ""
            concept_k = _get_movement_concept_key(_norm_movement_concept(label_orig)) if (not is_hdr_or_title and not is_tot_row and label_orig) else None

            for c in range(1, max_table_col + 1):
                orig_cell = ws.cell(orig_r, c)
                target_cell = ws.cell(target_r, c)
                
                val_o = orig_cell.value
                if isinstance(val_o, str) and val_o.upper().startswith('=SUM(') and ':' in val_o and not is_tot_row:
                    target_cell.value = re.sub(r'([A-Za-z]+)\d+', rf'\g<1>{target_r}', val_o)
                elif c >= comp_c + 1 and concept_k and concept_k in cmp_concept_map and ext_cmp_ws is not None:
                    ext_r = cmp_concept_map[concept_k]
                    c_ext = c + cmp_col_offset
                    if 1 <= c_ext <= ext_cmp_ws.max_column:
                        val_ext = ext_cmp_ws.cell(ext_r, c_ext).value
                        if val_ext is None or str(val_ext).strip() in ['', '-', '-   ']:
                            target_cell.value = 0
                        elif isinstance(val_ext, (int, float)):
                            target_cell.value = float(val_ext) if isinstance(val_ext, float) else int(val_ext)
                        elif isinstance(val_ext, str) and not val_ext.startswith('='):
                            try:
                                clean_s = val_ext.replace('.', '').replace(',', '.') if (',' in val_ext and '.' in val_ext) else val_ext.replace(',', '')
                                target_cell.value = float(clean_s)
                            except Exception:
                                target_cell.value = val_ext
                        else:
                            target_cell.value = 0
                    else:
                        target_cell.value = 0
                elif c >= comp_c + 1 and concept_k and concept_k in act_concept_map and concept_k not in cmp_concept_map:
                    # Si el bloque actual provino de un anexo externo pero no existe anexo para el período comparativo,
                    # no arrastrar los montos del período actual al período comparativo.
                    if isinstance(val_o, str) and val_o.startswith('='):
                        target_cell.value = shift_formula_rows(val_o, row_shift, min_row=start_row)
                    else:
                        target_cell.value = 0
                else:
                    val = orig_cell.value
                    if val is not None:
                        if isinstance(val, (int, float)) and val == int(act_yr_str) and c == comp_c:
                            target_cell.value = int(cmp_yr_str)
                        elif isinstance(val, str):
                            if val.startswith('='):
                                target_cell.value = shift_formula_rows(val, row_shift, min_row=start_row)
                            else:
                                val_new = val
                                val_lower = val.strip().lower()
                                if "saldo inicial" in val_lower:
                                    val_new = re.sub(r'\b\d{2}[\.\/]\d{2}[\.\/]20\d{2}\b', cmp_init_str, val_new)
                                elif "saldo final" in val_lower or "detalle al" in val_lower or "concepto al" in val_lower:
                                    val_new = re.sub(r'\b\d{2}[\.\/]\d{2}[\.\/]20\d{2}\b', cmp_date_str, val_new)
                                else:
                                    val_new = re.sub(r'\b\d{2}[\.\/]\d{2}[\.\/]20\d{2}\b', cmp_date_str, val_new)

                                val_new = val_new.replace(act_yr_str, cmp_yr_str)
                                val_new = re.sub(r'\[COMPARATIVO\]', '', val_new, flags=re.IGNORECASE).strip()
                                target_cell.value = val_new if val_new else None
                        else:
                            target_cell.value = val if isinstance(val, (int, float)) and val != 0 else None

                if orig_cell.has_style:
                    target_cell.number_format = orig_cell.number_format
                    try:
                        import openpyxl
                        target_cell.font = openpyxl.styles.Font(
                            name=orig_cell.font.name,
                            size=orig_cell.font.size,
                            bold=orig_cell.font.bold,
                            italic=orig_cell.font.italic,
                            color=orig_cell.font.color
                        )
                    except Exception:
                        pass

        marker_cell = ws.cell(comp_r, comp_c)
        if marker_cell.value and isinstance(marker_cell.value, str):
            cleaned = re.sub(r'\[COMPARATIVO\]', '', marker_cell.value, flags=re.IGNORECASE).strip()
            marker_cell.value = cleaned if cleaned else None

def detect_row_section_rubro(ws, current_row):
    for r in range(current_row - 1, 0, -1):
        for c in range(1, min(6, ws.max_column + 1)):
            val = ws.cell(row=r, column=c).value
            if val is not None and isinstance(val, str):
                val_clean = clean_label(val)
                if val_clean in SECTION_TITLE_TO_PL_RUBRO:
                    return SECTION_TITLE_TO_PL_RUBRO[val_clean]
                for title_kw, pl_key in SECTION_TITLE_TO_PL_RUBRO.items():
                    if title_kw in val_clean or val_clean in title_kw:
                        return pl_key
    return None

def get_filtered_value(period_ctx, matched_key, match_type, flow_key, row, ws, section_rubro=None, filter_sign=None, col=None):
    """
    Retorna el valor filtrado (corriente vs no corriente y bruto vs depreciación) para el matched_key.
    Si el estado de la fila es corriente o no corriente, se filtran las cuentas
    en el contexto según su Clasificación balance.

    Para notas de P&L, si matched_key coincide con el rubro en sí (y no es un sub-ítem),
    se suman todos los sub-ítems del rubro (comportamiento de Total).
    """
    if match_type == 'pl':
        pl_ctx = period_ctx.get('pl', {})

        def _sum_rubro(rdata, fkey):
            """Suma todos los sub-ítems de un rubro para obtener el total."""
            total = 0.0
            for sub_val in rdata.values():
                if isinstance(sub_val, dict):
                    total += float(sub_val.get(fkey, 0.0))
            return total

        # Si se especificó el rubro exacto de la sección/tabla activa
        if section_rubro and section_rubro in pl_ctx:
            rubro_data = pl_ctx[section_rubro]
            if matched_key in rubro_data:
                return float(rubro_data[matched_key].get(flow_key, 0.0))
            # Si la clave buscada ES el rubro mismo, sumar todos los sub-ítems
            if matched_key == section_rubro:
                return _sum_rubro(rubro_data, flow_key)
            return 0.0

        sheet_title = str(ws.title).strip()
        allowed_rubros = SHEET_TO_PL_RUBROS.get(sheet_title, [])

        # Si la pestaña no está en el mapa, buscar en todos los rubros como fallback
        if not allowed_rubros:
            for r_key, rubro_data in pl_ctx.items():
                if matched_key in rubro_data:
                    return float(rubro_data[matched_key].get(flow_key, 0.0))
                if matched_key == r_key:
                    return _sum_rubro(rubro_data, flow_key)
            return 0.0

        # Buscar en los rubros específicos permitidos para esta pestaña
        for rubro in allowed_rubros:
            rubro_clean = ''.join(c for c in unicodedata.normalize('NFD', rubro.lower()) if unicodedata.category(c) != 'Mn')
            rubro_clean = re.sub(r'\s+', ' ', rubro_clean).strip()

            rubro_data = pl_ctx.get(rubro_clean)
            if rubro_data:
                if matched_key in rubro_data:
                    return float(rubro_data[matched_key].get(flow_key, 0.0))
                # Si la clave buscada ES el rubro mismo, sumar todos los sub-ítems
                if matched_key == rubro_clean:
                    return _sum_rubro(rubro_data, flow_key)
        return 0.0

    group_data = period_ctx.get(match_type, {}).get(matched_key)
    if not group_data and filter_sign:
        other_type = 'nota2' if match_type == 'nota1' else ('nota1' if match_type == 'nota2' else None)
        if other_type:
            group_data = period_ctx.get(other_type, {}).get(matched_key)

    if not group_data:
        return 0.0
        
    # Si no tiene acct_details (o es P&L), devolvemos el valor agregado directamente
    if 'acct_details' not in group_data:
        return float(group_data.get(flow_key, 0.0))
        
    status = detect_row_corriente_status(ws, row, col=col)
        
    # Filtrar detalles de cuentas por corriente/no corriente y signo (bruto vs depreciación)
    filtered_sum = 0.0
    for detail in group_data['acct_details']:
        val = detail.get(flow_key, 0.0)
        if filter_sign == 'gross' and val < 0:
            continue
        if filter_sign == 'depr' and val > 0:
            continue

        sheet_title_low = str(ws.title).strip().lower()
        if sheet_title_low in ['intangibles', 'intangible']:
            clasif_lbl = str(detail.get('clasificacion_bal', '')).lower()
            if clasif_lbl and 'intangib' not in clasif_lbl:
                continue
        elif sheet_title_low in ['efectivo']:
            clasif_lbl = str(detail.get('clasificacion_bal', '')).lower()
            if clasif_lbl and 'efectivo' not in clasif_lbl:
                continue
        elif sheet_title_low in ['cuentas por pagar']:
            clasif_lbl = str(detail.get('clasificacion_bal', '')).lower()
            if clasif_lbl and not any(kw in clasif_lbl for kw in ['pagar', 'comercial', 'proveedor']):
                continue

        if status != "unknown":
            clasif = detail['clasificacion_bal'].lower()
            is_nc = any(nc_kw in clasif for nc_kw in ["no corriente", "no-corriente", "largo plazo", "propiedad", "activo fijo", "intangib", "derecho", "plusvalia", "diferid"])
            if status == "no_corriente" and not is_nc:
                continue
            elif status == "corriente" and is_nc:
                continue

        filtered_sum += val
            
    return filtered_sum


def format_date_dots(period_str):
    if not period_str or str(period_str).strip() in ["Ninguno", "None", ""]:
        return ""
    try:
        parts = str(period_str).strip().split('-')
        if len(parts) >= 2:
            year = int(parts[0])
            m_part = parts[1].strip()
            months = {'ene':1,'feb':2,'mar':3,'abr':4,'may':5,'jun':6,'jul':7,'ago':8,'sep':9,'oct':10,'nov':11,'dic':12,
                      'jan':1,'apr':4,'aug':8,'dec':12}
            if m_part.isdigit():
                month = int(m_part)
            else:
                month = months.get(m_part.lower()[:3], 12)
            last_day = str(calendar.monthrange(year, month)[1]).zfill(2)
            m_two = str(month).zfill(2)
            return f"{last_day}.{m_two}.{year}"
    except Exception:
        pass
    return str(period_str)

def format_period_to_spanish_date(period_str, target_lang='es'):
    if not period_str or period_str == "Ninguno":
        return ""
    try:
        parts = str(period_str).strip().split('-')
        if len(parts) >= 2:
            year = int(parts[0])
            month = int(parts[1])
            last_day = calendar.monthrange(year, month)[1]
            if str(target_lang).lower() == 'en':
                months_en = {
                    1: "January", 2: "February", 3: "March", 4: "April",
                    5: "May", 6: "June", 7: "July", 8: "August",
                    9: "September", 10: "October", 11: "November", 12: "December"
                }
                return f"{months_en[month]} {last_day}, {year}"
            else:
                months_es = {
                    1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
                    5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
                    9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
                }
                return f"{last_day} de {months_es[month]} {year}"
    except Exception:
        pass
    return period_str

def build_entity_context(tb_df, map_balance_df, map_pl_df=None, empresa_name=None, periodo_str=None):
    """
    Agrupa los saldos de la TB para una entidad según las descripciones
    de mapeo de balance y P&L.
    """
    if tb_df is None or tb_df.empty:
        return {'nota1': {}, 'nota2': {}, 'pl': {}}
        
    mapping_cols = ['N° de Cuenta \n', 'N de Cuenta', 'N° de Cuenta', 'Cuenta', 'cuenta_id']
    tb_cuenta_col = next((c for c in tb_df.columns if c.strip() in [mc.strip() for mc in mapping_cols]), None)
    
    saldo_cols = ['Saldo DR/CR\n', 'Saldo DR/CR', 'Saldos', 'Saldo', 'saldo_final']
    saldo_col = next((c for c in tb_df.columns if c.strip() in [sc.strip() for sc in saldo_cols]), None)
    
    tb_nombre_cols = ['Nombre de la cuenta', 'Nombre cuenta', 'Nombre', 'descripcion', 'Detalle']
    tb_nombre_col = next((c for c in tb_df.columns if c.strip() in [nc.strip() for nc in tb_nombre_cols]), None)

    if not tb_cuenta_col or not saldo_col:
        return {'nota1': {}, 'nota2': {}, 'pl': {}}

    tb_df = tb_df.copy()
    tb_df[tb_cuenta_col] = tb_df[tb_cuenta_col].astype(str).str.strip()
    tb_df[saldo_col] = pd.to_numeric(tb_df[saldo_col], errors='coerce').fillna(0.0)

    sums_nota1 = {}
    sums_nota2 = {}
    pl_mappings = {}

    # Mapeo de Balance
    if map_balance_df is not None and not map_balance_df.empty:
        map_balance_copy = map_balance_df.copy()
        nota1_col = next((c for c in map_balance_copy.columns if "nota 1" in c.lower()), None)
        nota2_col = next((c for c in map_balance_copy.columns if "nota 2" in c.lower()), None)
        if nota1_col:
            map_balance_copy[nota1_col] = map_balance_copy[nota1_col].astype(str).str.strip()
        if nota2_col:
            map_balance_copy[nota2_col] = map_balance_copy[nota2_col].astype(str).str.strip()
            
        from src.core.sabana_builder import build_balance_sabana
        merged_df = build_balance_sabana(tb_df, map_balance_copy)
        
        map_cuenta_col = next((c for c in map_balance_copy.columns if c.strip() in [mc.strip() for mc in mapping_cols]), None)
        if map_cuenta_col:
            def get_grouped_sums(merged_df_obj, group_col):
                if not group_col or group_col not in merged_df_obj.columns:
                    return {}
                grouped = merged_df_obj.groupby(group_col)
                result = {}
                for name, gp in grouped:
                    records = gp.to_dict('records')
                    accts = [str(r.get(tb_cuenta_col, '')).strip() for r in records if tb_cuenta_col in r]
                    acct_names = [r.get(tb_nombre_col, '') for r in records if tb_nombre_col and tb_nombre_col in r]
                    
                    credit_abs = 0.0
                    debit_abs = 0.0
                    acct_details = []
                    
                    for r_val in records:
                        acc_id = str(r_val.get(tb_cuenta_col, '')).strip()
                        bal_raw = r_val.get(saldo_col, 0.0) if saldo_col in r_val else 0.0
                        bal = abs(float(bal_raw)) if pd.notna(bal_raw) else 0.0
                        if acc_id.startswith('2') or acc_id.startswith('3101'):
                            credit_abs += bal
                        else:
                            debit_abs += bal

                    if credit_abs == 0.0 and debit_abs == 0.0:
                        factor = determine_factor_for_accounts(accts)
                    else:
                        factor = -1 if credit_abs >= debit_abs else 1
                    
                    sum_val = float(gp[saldo_col].sum()) if saldo_col in gp.columns else 0.0
                    sum_ini = float(gp['saldo_inicial'].sum()) if 'saldo_inicial' in gp.columns else 0.0
                    sum_deb = float(gp['debitos'].sum()) if 'debitos' in gp.columns else 0.0
                    sum_cred = float(gp['creditos'].sum()) if 'creditos' in gp.columns else 0.0

                    for r_val in records:
                        acc_id = str(r_val.get(tb_cuenta_col, '')).strip()
                        val_raw = r_val.get(saldo_col, 0.0) if saldo_col in r_val else 0.0
                        val_num = float(val_raw) if pd.notna(val_raw) else 0.0
                        ini_num = float(r_val.get('saldo_inicial', 0.0)) if pd.notna(r_val.get('saldo_inicial')) else 0.0
                        deb_num = float(r_val.get('debitos', 0.0)) if pd.notna(r_val.get('debitos')) else 0.0
                        cred_num = float(r_val.get('creditos', 0.0)) if pd.notna(r_val.get('creditos')) else 0.0
                        
                        acct_details.append({
                            'cuenta_id': acc_id,
                            'nombre': r_val.get(tb_nombre_col, '') if tb_nombre_col else '',
                            'val': val_num * factor,
                            'inicial': ini_num * factor,
                            'debitos': deb_num * factor,
                            'creditos': cred_num * factor,
                            'clasificacion_bal': str(r_val.get('Clasificación balance', '') or '').strip()
                        })
                    
                    result[name] = {
                        'val': sum_val * factor,
                        'inicial': sum_ini * factor,
                        'debitos': sum_deb * factor,
                        'creditos': sum_cred * factor,
                        'accts': accts,
                        'acct_names': acct_names,
                        'acct_details': acct_details
                    }
                return result

            if nota1_col:
                sums_nota1 = get_grouped_sums(merged_df, nota1_col)
            if nota2_col:
                sums_nota2 = get_grouped_sums(merged_df, nota2_col)

    # Mapeo de P&L
    # Columnas a excluir del mapeo de notas en map_pl
    # Se excluyen columnas de identificación de cuenta (código y nombre) así como clasificaciones
    # que no son rubros de P&L, para que no aparezcan como claves espurias en el contexto.
    _PL_EXCLUDE_COLS = {
        'Cuenta', 'Detalle', 'Clasificación Flujo Efectivo', 'Clasificacion',
        'Clasificación Flujo Efectivo', 'Clasificaci\u00f3n Flujo Efectivo'
    }
    # Palabras clave que identifican columnas de ID/nombre de cuenta (se excluyen automáticamente)
    _PL_EXCLUDE_KEYWORDS = [
        'n° de cuenta', 'n de cuenta', 'numero de cuenta', 'cod cuenta', 'codigo cuenta',
        'nombre de la cuenta', 'nombre cuenta', 'nombre_cuenta',
        'clasificacion flujo', 'flujo efectivo', 'clasificacion',
    ]
    if map_pl_df is not None and not map_pl_df.empty:
        from src.models.pl_cubo_db import PlCuboDB
        from src.core.sabana_builder import build_pl_sabana
        import unicodedata
        import re
        
        # Cargar pl_df combinando todas las empresas involucradas
        pl_df_list = []
        cos = [empresa_name] if isinstance(empresa_name, str) else empresa_name
        for co in cos:
            df_pl = PlCuboDB.get_pl_cubo(co, periodo_str) if periodo_str else None
            if df_pl is not None and not df_pl.empty:
                pl_df_list.append(df_pl)
                
        if pl_df_list:
            pl_df_comb = pd.concat(pl_df_list, ignore_index=True)
            pl_cuenta_col = next((c for c in pl_df_comb.columns if "cuenta" in str(c).lower() and "nombre" not in str(c).lower()), pl_df_comb.columns[0])
            pl_nombre_col = next((c for c in pl_df_comb.columns if "nombre" in str(c).lower()), pl_df_comb.columns[1] if len(pl_df_comb.columns) > 1 else None)
            
            ignorar_cols = [pl_cuenta_col]
            if pl_nombre_col:
                ignorar_cols.append(pl_nombre_col)
                
            agg_dict = {c: 'sum' for c in pl_df_comb.columns if c not in ignorar_cols}
            if pl_nombre_col:
                agg_dict[pl_nombre_col] = 'first'
                
            pl_df = pl_df_comb.groupby(pl_cuenta_col, as_index=False).agg(agg_dict)
        else:
            pl_df = None
            
        merged_pl = build_pl_sabana(None, map_pl_df, tb_df)
        if tb_df is not None and not tb_df.empty:
            tb_acct_set = set(tb_df[tb_cuenta_col].astype(str).str.strip())
            map_pl_cuenta_col = next((c for c in map_pl_df.columns if c.strip() in [mc.strip() for mc in mapping_cols]), None)
            acct_col_in_merged = tb_cuenta_col if tb_cuenta_col in merged_pl.columns else map_pl_cuenta_col
            if acct_col_in_merged and acct_col_in_merged in merged_pl.columns:
                merged_pl = merged_pl[merged_pl[acct_col_in_merged].astype(str).str.strip().isin(tb_acct_set)]
        
        # Pre-cargar datos de P&L Cubo si existen para cuadrar saldos al centavo
        pl_cubo_data = {}  # {col_sanitized: {acc_id: sum_value}}
        has_pl_cubo = pl_df is not None
        
        def pl_sanitize(text):
            if pd.isna(text): return ""
            clean_str = str(text).strip().lower()
            clean_str = ''.join(c for c in unicodedata.normalize('NFD', clean_str) if unicodedata.category(c) != 'Mn')
            clean_str = re.sub(r'\s+', ' ', clean_str)
            return clean_str
            
        if has_pl_cubo:
            ignorar_rubros = ["n° de cuenta", "nombre de la cuenta", "cuenta", "nombre", "unnamed: 0"]
            for col_pl in pl_df.columns:
                if str(col_pl).strip().lower() in ignorar_rubros:
                    continue
                col_san = pl_sanitize(col_pl)
                if col_san not in pl_cubo_data:
                    pl_cubo_data[col_san] = {}
                for _, r_val in pl_df.iterrows():
                    acc_id = str(r_val[pl_df.columns[0]]).strip()
                    val_str = r_val.get(col_pl)
                    if val_str is not None:
                        try:
                            pl_cubo_data[col_san][acc_id] = float(val_str)
                        except ValueError:
                            pass

        map_pl_cuenta_col = next((c for c in map_pl_df.columns if c.strip().lower() in [mc.strip().lower() for mc in mapping_cols]), None)
        if map_pl_cuenta_col:
            split_accounts = set()
            exclude_from_mapping = ['cuenta', 'detalle', 'nombre', 'descripcion', 'flujo', 'unnamed']
            mapping_cols_to_check = [
                c for c in map_pl_df.columns 
                if not any(x in c.lower() for x in exclude_from_mapping)
            ]
            for _, row_val in map_pl_df.iterrows():
                acc_id = str(row_val[map_pl_cuenta_col]).strip()
                non_empty_count = 0
                for c_col in mapping_cols_to_check:
                    val_cell = row_val.get(c_col)
                    if val_cell is not None and str(val_cell).strip().lower() not in ('nan', 'none', ''):
                        non_empty_count += 1
                if non_empty_count > 1:
                    split_accounts.add(acc_id)

            for col in map_pl_df.columns:
                col_lower_norm = ''.join(
                    c for c in unicodedata.normalize('NFD', col.replace('\xa0', ' ').strip().lower())
                    if unicodedata.category(c) != 'Mn'
                )
                col_lower_norm = re.sub(r'\s+', ' ', col_lower_norm)
                if col in _PL_EXCLUDE_COLS or col.lower().startswith('unnamed'):
                    continue
                # Excluir columnas de código/nombre de cuenta y clasificaciones no-P&L
                if any(kw in col_lower_norm for kw in _PL_EXCLUDE_KEYWORDS):
                    continue
                
                col_san = col.replace("\xa0", " ").strip().lower()
                clean_col = ''.join(c for c in unicodedata.normalize('NFD', col_san) if unicodedata.category(c) != 'Mn')
                clean_col = re.sub(r'\s+', ' ', clean_col)
                cubo_col_dict = pl_cubo_data.get(clean_col)
                
                if clean_col not in pl_mappings:
                    pl_mappings[clean_col] = {}
                
                grouped = merged_pl.groupby(col)
                for name, gp in grouped:
                    k_str = str(name).strip()
                    if k_str and k_str.lower() != 'nan':
                        ck = clean_label(k_str)
                        accts = gp[tb_cuenta_col].tolist()
                        acct_names = gp[tb_nombre_col].tolist() if tb_nombre_col else []
                        
                        credit_abs = 0.0
                        debit_abs = 0.0
                        for _, r_val in gp.iterrows():
                            acc_id = str(r_val[tb_cuenta_col]).strip()
                            bal = abs(float(r_val[saldo_col])) if (saldo_col in gp.columns and pd.notna(r_val[saldo_col])) else 0.0
                            if 'ingreso' in clean_col:
                                credit_abs += bal
                            elif acc_id.startswith('2') or acc_id.startswith('3101'):
                                credit_abs += bal
                            else:
                                debit_abs += bal
                        
                        if credit_abs == 0.0 and debit_abs == 0.0:
                            factor = -1 if ('ingreso' in clean_col or any(a.startswith('2') or a.startswith('3101') for a in accts)) else determine_factor_for_accounts(accts)
                        else:
                            factor = -1 if credit_abs >= debit_abs else 1
                        
                        if ck not in pl_mappings[clean_col]:
                            pl_mappings[clean_col][ck] = {
                                'val': 0.0, 'inicial': 0.0, 'debitos': 0.0, 'creditos': 0.0,
                                'accts': [], 'acct_names': []
                            }
                        
                        sum_val = 0.0
                        sum_ini = 0.0
                        sum_deb = 0.0
                        sum_cred = 0.0
                        visited_cubo_accs = set()
                        
                        for _, r_val in gp.iterrows():
                            acc_id = str(r_val[tb_cuenta_col]).strip()
                            
                            use_cubo = False
                            tb_val = float(r_val[saldo_col]) if (saldo_col in gp.columns and pd.notna(r_val[saldo_col])) else 0.0
                            tb_deb = float(r_val.get('debitos', 0.0)) if pd.notna(r_val.get('debitos', 0.0)) else 0.0
                            tb_cred = float(r_val.get('creditos', 0.0)) if pd.notna(r_val.get('creditos', 0.0)) else 0.0
                            tb_ini = float(r_val.get('saldo_inicial', 0.0)) if pd.notna(r_val.get('saldo_inicial', 0.0)) else 0.0
                            
                            has_tb_movement = (abs(tb_val) > 0.001 or abs(tb_deb) > 0.001 or abs(tb_cred) > 0.001 or abs(tb_ini) > 0.001)
                            is_split = (acc_id in split_accounts)
                            
                            if has_pl_cubo:
                                if acc_id not in visited_cubo_accs:
                                    cubo_val = cubo_col_dict.get(acc_id, 0.0) if cubo_col_dict else 0.0
                                    sum_val += cubo_val * factor
                                    sum_ini += cubo_val * factor
                                    visited_cubo_accs.add(acc_id)
                                use_cubo = True
                                    
                            if not use_cubo:
                                sum_val += tb_val * factor
                                sum_ini += tb_ini * factor
                                sum_deb += tb_deb * factor
                                sum_cred += tb_cred * factor
                                
                        pl_mappings[clean_col][ck]['val'] += sum_val
                        pl_mappings[clean_col][ck]['inicial'] += sum_ini
                        pl_mappings[clean_col][ck]['debitos'] += sum_deb
                        pl_mappings[clean_col][ck]['creditos'] += sum_cred
                        pl_mappings[clean_col][ck]['accts'].extend(accts)
                        pl_mappings[clean_col][ck]['acct_names'].extend(acct_names)

    # Limpiar y agregar etiquetas duplicadas después de normalizar espacios
    sums_nota1_clean = {}
    for k, v in sums_nota1.items():
        if k is not None:
            ck = clean_label(k)
            if ck not in sums_nota1_clean:
                sums_nota1_clean[ck] = {
                    'val': 0.0, 'inicial': 0.0, 'debitos': 0.0, 'creditos': 0.0,
                    'accts': [], 'acct_names': [], 'acct_details': []
                }
            sums_nota1_clean[ck]['val'] += v['val']
            sums_nota1_clean[ck]['inicial'] += v['inicial']
            sums_nota1_clean[ck]['debitos'] += v['debitos']
            sums_nota1_clean[ck]['creditos'] += v['creditos']
            sums_nota1_clean[ck]['accts'].extend(v['accts'])
            sums_nota1_clean[ck]['acct_names'].extend(v['acct_names'])
            sums_nota1_clean[ck]['acct_details'].extend(v['acct_details'])

    sums_nota2_clean = {}
    for k, v in sums_nota2.items():
        if k is not None:
            ck = clean_label(k)
            if ck not in sums_nota2_clean:
                sums_nota2_clean[ck] = {
                    'val': 0.0, 'inicial': 0.0, 'debitos': 0.0, 'creditos': 0.0,
                    'accts': [], 'acct_names': [], 'acct_details': []
                }
            sums_nota2_clean[ck]['val'] += v['val']
            sums_nota2_clean[ck]['inicial'] += v['inicial']
            sums_nota2_clean[ck]['debitos'] += v['debitos']
            sums_nota2_clean[ck]['creditos'] += v['creditos']
            sums_nota2_clean[ck]['accts'].extend(v['accts'])
            sums_nota2_clean[ck]['acct_names'].extend(v['acct_names'])
            sums_nota2_clean[ck]['acct_details'].extend(v['acct_details'])

    return {
        'nota1': sums_nota1_clean,
        'nota2': sums_nota2_clean,
        'pl': pl_mappings
    }


def fix_formulas_column_letters(ws, col_letter_map):
    import re
    for row in range(1, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=row, column=col)
            val = cell.value
            if isinstance(val, str) and val.startswith('='):
                def replace_col(match):
                    c_let = match.group(1)
                    r_num = match.group(2)
                    new_c_let = col_letter_map.get(c_let, c_let)
                    return f"{new_c_let}{r_num}"
                new_val = re.sub(r'([A-Z]+)(\d+)', replace_col, val)
                cell.value = new_val

class NoteGenerator:
    def __init__(self, template_path):
        self.template_path = template_path

    def generate(self,
                 sheet_names,
                 entity_contexts,
                 active_entity_name,
                 is_consolidated,
                 scale_factor=1.0,
                 periodo_actual_str=None,
                 periodo_comp_str=None,
                 map_balance_df=None,
                 map_pl_df=None,
                 target_lang='es',
                 evaluate_formulas=True):
        """
        Genera un Excel de notas filtrando y llenando las pestañas indicadas 
        según el contexto consolidado o individual.
        """
        if not sheet_names:
            raise ValueError("No sheet names specified for note generation. The master template must contain at least one sheet for this note.")

        from src.reporting.notes import NOTE_REGISTRY
        wb = load_workbook(self.template_path)
        
        # 1. Eliminar hojas no relacionadas (preservando DB_DATA)
        matching_sheets = [s for s in sheet_names if s in wb.sheetnames]
        fallback_name = None
        if not matching_sheets and "DB_DATA" not in wb.sheetnames:
            fallback_name = sheet_names[0] if sheet_names else "Nota"
            ws_fb = wb.create_sheet(title=fallback_name)
            ws_fb.cell(row=1, column=1, value=f"La nota '{fallback_name}' no se encuentra en la plantilla actual.")

        for name in list(wb.sheetnames):
            if name not in sheet_names and name != "DB_DATA" and name != fallback_name:
                wb.remove(wb[name])

        # 1.5 Volcar base de datos plana a DB_DATA
        try:
            from src.models.trial_balance_db import TrialBalanceDB
            from openpyxl.utils.dataframe import dataframe_to_rows
            
            companies = []
            if is_consolidated:
                from src.models.database import SessionLocal
                from src.models.consolidacion import ConsolidationGroup
                grupo_name = active_entity_name.replace("[GRUPO] ", "").strip()
                db = SessionLocal()
                try:
                    grupo_obj = db.query(ConsolidationGroup).filter_by(nombre_grupo=grupo_name).first()
                    if grupo_obj:
                        companies.append(grupo_obj.empresa_matriz)
                        if grupo_obj.filial_is_group:
                            def get_sub_companies(sub_g_id):
                                sub_g = db.query(ConsolidationGroup).filter_by(id=sub_g_id).first()
                                if sub_g:
                                    c = [sub_g.empresa_matriz]
                                    if sub_g.filial_is_group:
                                        c.extend(get_sub_companies(int(sub_g.empresa_filial)))
                                    else:
                                        c.append(sub_g.empresa_filial)
                                    return c
                                return []
                            companies.extend(get_sub_companies(int(grupo_obj.empresa_filial)))
                        else:
                            companies.append(grupo_obj.empresa_filial)
                except Exception as e:
                    print(f"Error al obtener filiales para DB_DATA: {e}")
                finally:
                    db.close()
            else:
                companies = [active_entity_name]
                
            if not companies:
                companies = [active_entity_name]
                
            if map_balance_df is None:
                try:
                    map_balance_df = pd.read_excel(self.template_path, sheet_name="Mapeo Balance")
                except:
                    pass
            if map_pl_df is None:
                try:
                    map_pl_df = pd.read_excel(self.template_path, sheet_name="Mapeo Ctas P&L Cubo")
                except:
                    pass
                    
            rows = []
            periods = [periodo_actual_str]
            if periodo_comp_str and periodo_comp_str != "Ninguno":
                periods.append(periodo_comp_str)
                
            def _clean_map_df(m_df):
                if m_df is None or m_df.empty:
                    return None
                m_clean = m_df.copy()
                mc_cols = ['N° de Cuenta \n', 'N de Cuenta', 'N° de Cuenta', 'Cuenta', 'cuenta_id']
                mc_col = next((c for c in m_clean.columns if c.strip() in [x.strip() for x in mc_cols]), None)
                if mc_col:
                    m_clean[mc_col] = m_clean[mc_col].astype(str).str.strip()
                    m_clean.rename(columns={mc_col: 'cuenta_key'}, inplace=True)
                return m_clean

            map_bal_clean = _clean_map_df(map_balance_df)
            map_pl_clean = _clean_map_df(map_pl_df)
            
            from src.core.snapshot_vault_manager import SnapshotVaultManager
            for co in companies:
                for per in periods:
                    tb = TrialBalanceDB.get_trial_balance(co, per)
                    if tb is not None and not tb.empty:
                        mb_per_clean = map_bal_clean
                        mp_per_clean = map_pl_clean
                        if per and (SnapshotVaultManager.has_active_snapshot(co, per) or SnapshotVaultManager.has_active_snapshot(active_entity_name, per)):
                            f_mb, f_mp = SnapshotVaultManager.get_frozen_mappings(co, per)
                            if f_mb is None:
                                f_mb, f_mp = SnapshotVaultManager.get_frozen_mappings(active_entity_name, per)
                            if f_mb is not None:
                                mb_per_clean = _clean_map_df(f_mb)
                            if f_mp is not None:
                                mp_per_clean = _clean_map_df(f_mp)

                        tb_cols = ['N° de Cuenta \n', 'N de Cuenta', 'N° de Cuenta', 'Cuenta', 'cuenta_id']
                        tb_cuenta_col = next((c for c in tb.columns if c.strip() in [x.strip() for x in tb_cols]), None)
                        
                        tb_nombre_cols = ['Nombre de la cuenta', 'Nombre cuenta', 'Nombre', 'descripcion', 'Detalle']
                        tb_nombre_col = next((c for c in tb.columns if c.strip() in [x.strip() for x in tb_nombre_cols]), None)
                        
                        saldo_cols = ['Saldo DR/CR\n', 'Saldo DR/CR', 'Saldos', 'Saldo', 'saldo_final']
                        saldo_col = next((c for c in tb.columns if c.strip() in [sc.strip() for sc in saldo_cols]), None)
                        
                        if tb_cuenta_col and saldo_col:
                            tb = tb.copy()
                            tb[tb_cuenta_col] = tb[tb_cuenta_col].astype(str).str.strip()
                            
                            if mb_per_clean is not None and 'cuenta_key' in mb_per_clean.columns:
                                tb = pd.merge(tb, mb_per_clean, left_on=tb_cuenta_col, right_on='cuenta_key', how='left')
                                
                            if mp_per_clean is not None and 'cuenta_key' in mp_per_clean.columns:
                                tb = pd.merge(tb, mp_per_clean, left_on=tb_cuenta_col, right_on='cuenta_key', how='left')
                                
                            for _, r_val in tb.iterrows():
                                c_id = r_val[tb_cuenta_col]
                                c_name = r_val[tb_nombre_col] if tb_nombre_col in r_val else ""
                                s_ini = r_val.get('saldo_inicial', 0.0)
                                deb = r_val.get('debitos', 0.0)
                                cred = r_val.get('creditos', 0.0)
                                s_fin = r_val.get(saldo_col, 0.0)
                                
                                clas_bal = r_val.get('Clasificación balance', '')
                                n1 = r_val.get('nota 1', '')
                                n2 = r_val.get('nota 2', '')
                                clas_pl = r_val.get('Clasificacion estado de resultados', '')
                                
                                rows.append({
                                    'empresa': co,
                                    'periodo': per,
                                    'tipo_periodo': 'actual' if per == periodo_actual_str else 'comparativo',
                                    'cuenta_id': c_id,
                                    'nombre_cuenta': c_name,
                                    'saldo_inicial': s_ini,
                                    'debitos': deb,
                                    'creditos': cred,
                                    'saldo_final': s_fin,
                                    'rubro_balance': clas_bal,
                                    'nota_1_linea': n1,
                                    'nota_2_linea': n2,
                                    'clasificacion_pl': clas_pl
                                })
            
            if rows:
                df_db_data = pd.DataFrame(rows)
                if "DB_DATA" in wb.sheetnames:
                    wb.remove(wb["DB_DATA"])
                ws_db = wb.create_sheet("DB_DATA")
                    
                for r_idx, row_data in enumerate(dataframe_to_rows(df_db_data, index=False, header=True)):
                    ws_db.append(row_data)
        except Exception as e:
            print(f"Error generando hoja DB_DATA: {e}")

        # Caché de celdas combinadas y lista de empresas para O(1) lookups y cero I/O redundante
        import os, datetime as _dt
        empresas_dir = os.path.join("data", "empresas")
        real_companies = [d for d in os.listdir(empresas_dir) if os.path.isdir(os.path.join(empresas_dir, d))] if os.path.exists(empresas_dir) else []
        
        _ws_merged_cache = {}

        def get_merged_cell_value(ws_obj, row, col):
            if ws_obj not in _ws_merged_cache:
                m_map = {}
                for merged_range in ws_obj.merged_cells.ranges:
                    top_left_val = ws_obj.cell(row=merged_range.min_row, column=merged_range.min_col).value
                    for r in range(merged_range.min_row, merged_range.max_row + 1):
                        for c in range(merged_range.min_col, merged_range.max_col + 1):
                            m_map[(r, c)] = top_left_val
                _ws_merged_cache[ws_obj] = m_map
            m_map = _ws_merged_cache[ws_obj]
            if (row, col) in m_map:
                return m_map[(row, col)]
            return ws_obj.cell(row=row, column=col).value

        def analyze_columns(ws):
            col_mapping = {} # col_idx -> {'entity': str, 'year': int, 'period_type': str}
            
            for c in range(2, ws.max_column + 1):

                col_entity = None
                col_year = None
                col_period_type = None
                
                # Buscar en filas 1 a 6 (cabeceras legítimas de columnas)
                for r in range(1, 7):
                    val = get_merged_cell_value(ws, r, c)
                    if val is None:
                        continue

                    # --- Manejar datetime de Excel ---
                    if isinstance(val, (_dt.datetime, _dt.date)):
                        col_year = val.year
                        # Primera columna de fecha = actual, siguiente = comp
                        # Se resuelve abajo por posición
                        continue

                    val_lower = str(val).strip().lower()
                    
                    # Buscar coincidencia con la empresa activa primero
                    act_norm = normalize_company_name(active_entity_name)
                    if act_norm and act_norm in val_lower:
                        col_entity = active_entity_name
                    else:
                        for co in real_companies:
                            co_norm = normalize_company_name(co)
                            if co_norm and co_norm in val_lower:
                                col_entity = co
                                break
                            
                    if not col_entity:
                        if 'consolidado' in val_lower or 'grupo' in val_lower:
                            col_entity = 'Consolidado'
                        elif 'pacifico' in val_lower:
                            col_entity = active_entity_name if 'pacifico' in active_entity_name.lower() else 'Pacifico Cable SpA'
                        elif 'parent' in val_lower:
                            col_entity = active_entity_name if 'parent' in active_entity_name.lower() else 'Db Terra Chile Parent SpA'
                        elif 'holdco' in val_lower or 'holco' in val_lower:
                            col_entity = active_entity_name if 'holdco' in active_entity_name.lower() else 'Db Terra Chile Holdco SpA'
                    
                    # Buscar año en texto (no confundir filas descriptivas de movimientos ni notas de auditoría con cabeceras)
                    if col_year is None:
                        if not any(k in val_lower for k in ['saldo inicial', 'saldo final', 'adiciones', 'desapropiaciones', 'bajas', 'transferencias', 'informe', 'validacion']):
                            match = re.search(r'\b(202\d)\b', val_lower)
                            if match:
                                col_year = int(match.group(1))
                        
                    # Usar límites de palabra o regex más estrictos para evitar falsos positivos con 'composición', 'compañía', etc.
                    if re.search(r'\bcomparativ[ao]s?\b|\bcomp\b|\bcomparativa\b', val_lower):
                        col_period_type = 'comp'
                    elif re.search(r'\bactual\b', val_lower):
                        col_period_type = 'actual'
                        
                # --- Validar si es columna descriptiva de texto (no de datos) ---
                # Si en las cabeceras de la primera tabla (filas 1-6) hay palabras descriptivas o de movimientos, no es columna de datos.
                has_data_header = False
                for r in range(1, 7):
                    vh = ws.cell(row=r, column=c).value
                    if vh is not None and str(vh).strip() != '':
                        vhs = str(vh).strip().lower()
                        vh_clean = clean_label(vh)
                        if vh_clean in SECTION_TITLE_TO_PL_RUBRO or re.search(r'#n\d+', vhs):
                            continue
                        if any(k in vhs for k in ['bruto', 'depreciaci', 'acumulada', 'neto', 'm$', '$', 'costo']) or isinstance(vh, (_dt.datetime, _dt.date)) or re.search(r'\b202\d\b', vhs):
                            has_data_header = True
                            break

                is_label_col = False
                for r in range(1, 7):
                    val_h = get_merged_cell_value(ws, r, c)
                    if val_h is not None and not isinstance(val_h, (_dt.datetime, _dt.date)):
                        val_h_clean = clean_label(val_h)
                        if re.search(r'#n\d+\.\d+', val_h_clean) or val_h_clean in SECTION_TITLE_TO_PL_RUBRO:
                            is_label_col = True
                            break
                        # Ignorar párrafos narrativos largos
                        if len(val_h_clean) > 45:
                            continue
                        # Usar límites de palabra \b para evitar falsos positivos como 'rut' dentro de 'bruto'
                        if re.search(r'\b(detalle|nombre|concepto|cuenta|descripcion|nombre entidad|rut|codigo|especificacion|glosa|item|tipo de relacion|relacion|saldo inicial|saldo final|adiciones|desapropiaciones|transferencias|deterioro|gasto por depreciacion)\b', val_h_clean):
                            is_label_col = True
                            break

                if is_label_col:
                    continue

                if not has_data_header:
                    continue

                if col_year is None:
                    for c_adj in [c + 1, c - 1, c + 2, c - 2]:
                        if 1 <= c_adj <= ws.max_column:
                            for r_adj in range(1, 7):
                                v_adj = get_merged_cell_value(ws, r_adj, c_adj)
                                if v_adj is not None:
                                    if isinstance(v_adj, (_dt.datetime, _dt.date)):
                                        col_year = v_adj.year
                                        break
                                    m_yr = re.search(r'\b(202\d)\b', str(v_adj))
                                    if m_yr:
                                        col_year = int(m_yr.group(1))
                                        break
                            if col_year:
                                break

                if col_entity or col_year or col_period_type:
                    # Si hay año pero no entidad, asignar la entidad activa
                    if not col_entity:
                        col_entity = active_entity_name
                    col_mapping[c] = {
                        'entity': col_entity,
                        'year': col_year,
                        'period_type': col_period_type
                    }

            # --- Fallback para plantillas individuales sin etiqueta de empresa ---
            # Si no se detectó ninguna columna con datos, buscar columnas que tienen
            # encabezados como "M$", "$", "%" o fechas y asignarlas a la entidad activa
            if not col_mapping:
                numeric_header_cols = []
                for c in range(2, ws.max_column + 1):
                    for r in range(1, 13):
                        val = get_merged_cell_value(ws, r, c)
                        if val is None:
                            continue
                        if isinstance(val, (_dt.datetime, _dt.date)):
                            numeric_header_cols.append((c, val.year))
                            break
                        val_str = str(val).strip().lower()
                        if val_str in ['m$', '$', '%', 'm', 'miles'] or re.search(r'202\d', val_str):
                            yr_match = re.search(r'(202\d)', val_str)
                            numeric_header_cols.append((c, int(yr_match.group(1)) if yr_match else None))
                            break

                for idx, (c, yr) in enumerate(numeric_header_cols):
                    period_type = 'actual' if idx == 0 else 'comp'
                    col_mapping[c] = {
                        'entity': active_entity_name,
                        'year': yr,
                        'period_type': period_type
                    }

            return col_mapping


        def find_best_context(col_entity, entity_contexts):
            if not col_entity:
                return None
            col_ent_lower = col_entity.lower()
            for k in entity_contexts.keys():
                k_lower = k.lower()
                if col_ent_lower in k_lower or k_lower in col_ent_lower:
                    return entity_contexts[k]
            if 'consolidado' in col_ent_lower:
                for k in entity_contexts.keys():
                    if 'grupo' in k.lower() or 'consolidado' in k.lower():
                        return entity_contexts[k]
            return None

        # 2. Formatear fechas para reemplazo en cabeceras
        formatted_actual = format_period_to_spanish_date(periodo_actual_str) if periodo_actual_str else ""
        formatted_comp = format_period_to_spanish_date(periodo_comp_str) if periodo_comp_str and periodo_comp_str != "Ninguno" else ""
        
        actual_yr = periodo_actual_str.split('-')[0] if periodo_actual_str else "2025"
        comp_yr = periodo_comp_str.split('-')[0] if periodo_comp_str and periodo_comp_str != "Ninguno" else "2024"

        # 3. Procesar cada hoja seleccionada
        for sheet_name in sheet_names:
            if sheet_name not in wb.sheetnames:
                continue
            ws = wb[sheet_name]
            
            clean_s = sheet_name.strip().lower()
            is_results_note = False
            for code, info in NOTE_REGISTRY.items():
                reg_sheets = [rs.strip().lower() for rs in info.get('sheets', [])]
                if clean_s in reg_sheets:
                    if info.get('category') == 'resultados':
                        is_results_note = True
                        break
            
            col_mapping = analyze_columns(ws)
            
            # Detectar todos los años (202x) en toda la hoja para tener la lista de años de la plantilla
            all_years = set()
            for r_y in range(1, ws.max_row + 1):
                for c_y in range(1, ws.max_column + 1):
                    val_y = ws.cell(row=r_y, column=c_y).value
                    if val_y is not None:
                        if isinstance(val_y, (datetime.datetime, datetime.date)):
                            all_years.add(val_y.year)
                        else:
                            match_y = re.search(r'\b(202\d)\b', str(val_y))
                            if match_y:
                                all_years.add(int(match_y.group(1)))
            years_found = sorted(list(all_years))
            template_year_to_period_type = {}
            if len(years_found) >= 2:
                template_year_to_period_type[years_found[-1]] = 'actual'
                template_year_to_period_type[years_found[-2]] = 'comp'
            elif len(years_found) == 1:
                template_year_to_period_type[years_found[0]] = 'actual'

            def detect_block_year(ws_obj, current_row, col_idx):
                for r_b in range(current_row - 1, 0, -1):
                    is_boundary_b = False
                    for c_lbl in range(1, min(5, ws_obj.max_column + 1)):
                        cell_lbl = ws_obj.cell(row=r_b, column=c_lbl).value
                        if cell_lbl is not None:
                            cell_lbl_clean = clean_label(cell_lbl)
                            if ("total" in cell_lbl_clean and "sub" not in cell_lbl_clean) or "transaccion" in cell_lbl_clean or "las principales" in cell_lbl_clean or "la composicion" in cell_lbl_clean:
                                is_boundary_b = True
                                break
                    
                    val_found = None
                    for c_b in [col_idx, col_idx + 1, col_idx - 1]:
                        if 1 <= c_b <= ws_obj.max_column:
                            val_b = get_merged_cell_value(ws_obj, r_b, c_b)
                            if val_b is not None:
                                if isinstance(val_b, (_dt.datetime, _dt.date)):
                                    val_found = val_b
                                    break
                                val_str_b = str(val_b).strip()
                                match_b = re.search(r'\b(202\d|199\d)\b', val_str_b)
                                if match_b:
                                    val_found = val_b
                                    break
                    
                    if val_found is not None:
                        if isinstance(val_found, (datetime.datetime, datetime.date)):
                            return val_found.year
                        val_str_b = str(val_found).strip()
                        match_b = re.search(r'\b(202\d|199\d)\b', val_str_b)
                        if match_b:
                            return int(match_b.group(1))
                    if is_boundary_b:
                        break
                return None

            # Reemplazar cabeceras de fechas en filas 1-8
            # Construir fecha del periodo actual en varios formatos para comparación
            _actual_dot = format_date_dots(periodo_actual_str)   # ej. "30.06.2026"
            _comp_dot   = format_date_dots(periodo_comp_str) if periodo_comp_str and periodo_comp_str != 'Ninguno' else ''
            _actual_yr  = periodo_actual_str.split('-')[0] if periodo_actual_str else actual_yr
            _actual_mo  = periodo_actual_str.split('-')[1] if periodo_actual_str and '-' in periodo_actual_str else ''
            _comp_yr    = periodo_comp_str.split('-')[0] if periodo_comp_str and '-' in (periodo_comp_str or '') else comp_yr
            _comp_mo    = periodo_comp_str.split('-')[1] if periodo_comp_str and '-' in (periodo_comp_str or '') else ''

            is_no_comp_notes = (not periodo_comp_str) or str(periodo_comp_str).strip().lower() in ["ninguno", "none", ""]

            for r in range(1, ws.max_row + 1):
                # 1. Recolectar celdas con fechas o palabras clave de fecha en la fila r
                row_date_cells = []
                for c in range(1, ws.max_column + 1):
                    val = ws.cell(row=r, column=c).value
                    if val is None:
                        continue
                    is_d = False
                    if isinstance(val, (datetime.datetime, datetime.date)):
                        is_d = True
                    elif isinstance(val, str):
                        v_low = val.strip().lower()
                        if len(v_low) > 25 or re.match(r'^(?:\d+|[a-z])[\)\.\-]\s+', v_low):
                            is_d = False
                        elif any(kw in v_low for kw in ['saldo', 'detalle', 'concepto', 'tipo', 'adicion', 'amortiz', 'desapropia', 'gasto', 'transferen', 'total', 'cambios', 'vencimiento', 'movimiento', 'reconciliacion', 'pasivo', 'activo', 'instrumentos', 'composicion', 'efectos', 'los saldos', 'conciliacion', 'la composicion']):
                            is_d = False
                        elif v_low in ["fecha  actual", "fecha actual", "fecha_actual", "actual", "fecha  comparativa", "fecha comparativa", "fecha_comparativa", "comparativa", "comparativo"]:
                            is_d = True
                        elif re.search(r'(\d{4})[.\-,](\d{1,2})[.\-,](\d{1,2})', val) or re.search(r'(\d{1,2})[.\-,](\d{1,2})[.\-,](\d{4})', val) or re.search(r'^\s*202\d\s*$', val):
                            is_d = True
                    if is_d:
                        row_date_cells.append(c)

                # 2. Si hay 2 o más columnas con fechas en la fila, determinar esquema de pares (Actual / Comparativo)
                if len(row_date_cells) >= 2:
                    row_date_cells.sort()
                    if len(row_date_cells) >= 4 and len(row_date_cells) % 2 == 0:
                        # Verificar si la fila inferior (r+1) tiene subcabeceras repetidas (ej: Corriente / No corriente bajo cada fecha en #N16.1)
                        sub_hdrs = [str(ws.cell(row=r + 1, column=col_d).value or "").strip().lower() for col_d in row_date_cells] if r < ws.max_row else []
                        has_grouped_subheaders = sum(1 for sh in sub_hdrs if "corriente" in sh) >= 2
                        if has_grouped_subheaders:
                            half = len(row_date_cells) // 2
                            act_cols = row_date_cells[:half]
                            cmp_cols = row_date_cells[half:]
                        else:
                            # Pares alternados (Actual, Comp) bajo encabezados de grupo (ej: Corriente [Actual, Comp] | No corriente [Actual, Comp] en #N14.1)
                            act_cols = row_date_cells[0::2]
                            cmp_cols = row_date_cells[1::2]
                    else:
                        act_cols = [row_date_cells[0]]
                        cmp_cols = [row_date_cells[1]]

                    for c_act in act_cols:
                        ws.cell(row=r, column=c_act).value = _actual_dot
                        if c_act in col_mapping:
                            col_mapping[c_act]['period_type'] = 'actual'
                    for c_cmp in cmp_cols:
                        if c_cmp in col_mapping:
                            col_mapping[c_cmp]['period_type'] = 'comp'
                        if not is_no_comp_notes:
                            ws.cell(row=r, column=c_cmp).value = _comp_dot
                        else:
                            v_cmp_orig = ws.cell(row=r, column=c_cmp).value
                            if isinstance(v_cmp_orig, (datetime.datetime, datetime.date)):
                                ws.cell(row=r, column=c_cmp).value = v_cmp_orig.strftime("%d.%m.%Y")
                elif len(row_date_cells) == 1:
                    c = row_date_cells[0]
                    val = ws.cell(row=r, column=c).value
                    col_info = col_mapping.get(c)
                    target_dot = _actual_dot
                    is_comp_cell = False
                    if col_info and col_info.get('period_type') == 'comp':
                        target_dot = _comp_dot
                        is_comp_cell = True
                    elif isinstance(val, (datetime.datetime, datetime.date)):
                        if template_year_to_period_type.get(val.year) == 'comp':
                            target_dot = _comp_dot
                            is_comp_cell = True
                    elif isinstance(val, str):
                        v_low = val.strip().lower()
                        if v_low in ["fecha  comparativa", "fecha comparativa", "fecha_comparativa", "comparativa", "comparativo"]:
                            target_dot = _comp_dot
                            is_comp_cell = True
                    if not (is_comp_cell and is_no_comp_notes):
                        ws.cell(row=r, column=c).value = target_dot
                    elif isinstance(val, (datetime.datetime, datetime.date)):
                        ws.cell(row=r, column=c).value = val.strftime("%d.%m.%Y")

            # Asegurar que template_year_to_period_type incluya los años efectivos tras reemplazo de cabeceras
            try:
                if _actual_yr and str(_actual_yr).isdigit():
                    template_year_to_period_type[int(_actual_yr)] = 'actual'
                if not is_no_comp_notes and _comp_yr and str(_comp_yr).isdigit() and int(_comp_yr) != int(_actual_yr):
                    template_year_to_period_type[int(_comp_yr)] = 'comp'
            except Exception:
                pass

            # Si es reporte individual, remapear todas las columnas de 'Consolidado' a la empresa activa
            if not is_consolidated:
                for c, info in col_mapping.items():
                    if info.get('entity') in ['Consolidado', None]:
                        info['entity'] = active_entity_name
                    # Renombrar visualmente en las cabeceras
                    for r in range(1, min(ws.max_row + 1, 10)):
                        val = ws.cell(row=r, column=c).value
                        if val is not None and isinstance(val, str):
                            val_lower = val.strip().lower()
                            if "consolidado" in val_lower:
                                ws.cell(row=r, column=c).value = val.replace("Consolidado", active_entity_name).replace("CONSOLIDADO", active_entity_name)

            # 3.5 Limpiar columnas activas en filas de datos para evitar números obsoletos de la plantilla
            if ws.title.strip().lower() == 'impuestos diferidos':
                # Impuestos Diferidos (#N13.1..#N13.5) se alimenta íntegramente desde el Anexo Externo.
                # Limpiamos residuos estáticos de la plantilla (ej. fila 73 con valores de ejemplo,
                # celdas 0 huérfanas entre cuadros en C27/C67 y arrastre =+D42 en C33) y omitimos
                # el cruce automático de cuentas contables para no contaminar #N13.1.
                for r_id in range(1, ws.max_row + 1):
                    v_lbl = ws.cell(row=r_id, column=2).value
                    lbl_str = str(v_lbl).strip().lower() if v_lbl is not None else ""
                    for c_id in range(3, min(ws.max_column + 1, 8)):
                        cell_id = ws.cell(row=r_id, column=c_id)
                        if cell_id.__class__.__name__ == 'MergedCell':
                            continue
                        v_id = cell_id.value
                        if v_id is None:
                            continue
                        if not lbl_str and v_id == 0:
                            cell_id.value = None
                        elif isinstance(v_id, str) and v_id.strip() in ("-", "—", "–"):
                            cell_id.value = 0
                        elif isinstance(v_id, (int, float)) and v_id != 0 and not (2020 <= v_id <= 2035):
                            cell_id.value = 0
                        elif isinstance(v_id, str) and v_id.startswith("=") and "SUM" not in v_id.upper() and "tasa efectiva" not in lbl_str:
                            cell_id.value = 0
                continue

            is_no_comp_notes = (not periodo_comp_str) or str(periodo_comp_str).strip().lower() in ["ninguno", "none", ""]
            if is_no_comp_notes:
                active_cols = [
                    c for c, info in col_mapping.items() 
                    if (info.get('entity') is not None or info.get('year') is not None or info.get('period_type') is not None)
                    and info.get('period_type') != 'comp'
                    and (info.get('year') is None or str(info.get('year')) == str(_actual_yr))
                ]
            else:
                active_cols = [c for c, info in col_mapping.items() if info.get('entity') is not None or info.get('year') is not None or info.get('period_type') is not None]
            
            # Conjunto de palabras clave que identifican cabeceras o filas no-datos
            HEADER_KEYWORDS = {
                'detalle', 'm$', '$', '%', 'neto', 'bruto', 
                'concepto', 'nombre', 'validacion', 'cuadro', 'fecha', 'actual', 
                'comparativa', 'unidades', 'codigo', 'cuenta', 'nombre cuenta', 'detalle de'
            }
            # En notas de resultados (P&L), 'total' puede ser una fila de datos válida
            # (e.g., 'Total gastos de administración'), por lo que NO se excluye.
            # En notas de Balance, sí se excluye para no procesar totales como ítems.
            if not is_results_note:
                HEADER_KEYWORDS.add('total')
                HEADER_KEYWORDS.add('amortizacion')
                HEADER_KEYWORDS.add('depreciacion')
            
            for row in range(1, ws.max_row + 1):
                # Evitar limpiar filas de cabecera (1 a 4 siempre protegidas)
                if row < 5:
                    continue
                
                # Identificar el primer texto no vacío de la fila
                first_text = None
                for c_idx in range(1, min(5, ws.max_column + 1)):
                    val = ws.cell(row=row, column=c_idx).value
                    if val is not None and str(val).strip() != "":
                        first_text = clean_label(val)
                        break
                
                # Si la fila es cabecera, total o validación, no se limpian sus columnas de datos
                has_year_in_row = False
                for c_idx in active_cols:
                    val_c = ws.cell(row=row, column=c_idx).value
                    if val_c is not None:
                        if isinstance(val_c, (datetime.datetime, datetime.date)):
                            has_year_in_row = True
                            break
                        if re.search(r'202\d', str(val_c)):
                            has_year_in_row = True
                            break

                # Contar celdas de texto en las columnas activas y totales para identificar cabeceras apiladas o de matriz
                string_cells_count = 0
                for c_temp in active_cols:
                    val_temp = ws.cell(row=row, column=c_temp).value
                    if val_temp is not None and isinstance(val_temp, str) and not val_temp.startswith('='):
                        try:
                            float(val_temp)
                        except ValueError:
                            string_cells_count += 1

                total_string_cells = sum(1 for c_temp in range(1, ws.max_column + 1) if ws.cell(row=row, column=c_temp).value is not None and isinstance(ws.cell(row=row, column=c_temp).value, str) and not str(ws.cell(row=row, column=c_temp).value).startswith('='))

                is_intro_text = (
                    first_text.startswith(('a continuacion', 'los pagos', 'las notas', 'el siguiente cuadro', 'la composicion', 'al 31 de', 'durante el')) or
                    (len(first_text.split()) > 10 and len(first_text) > 60)
                ) if first_text else False

                if first_text is None or any(kw in first_text for kw in HEADER_KEYWORDS) or has_year_in_row or string_cells_count > 1 or total_string_cells >= 3 or is_intro_text:
                    continue

                # Si es una fila de datos legítima, verificar si pertenece a un bloque con cabecera de fechas
                # y no a una tabla de revelación manual (ej. vencimientos de arriendos)
                has_date_hdr = False
                for c_chk in active_cols:
                    for r_hdr_chk in range(row - 1, max(0, row - 12), -1):
                        v_hdr = get_merged_cell_value(ws, r_hdr_chk, c_chk)
                        if v_hdr is not None:
                            if isinstance(v_hdr, (_dt.datetime, _dt.date)) or re.search(r'\b(202\d)\b', str(v_hdr)):
                                has_date_hdr = True
                                break
                    if has_date_hdr:
                        break

                has_share_hdr = False
                if ws.title.lower() in ['patrimonio', 'patrimonio'] and row > 10:
                    has_share_hdr = True
                else:
                    for c_chk in range(1, min(6, ws.max_column + 1)):
                        v_chk = get_merged_cell_value(ws, row, c_chk)
                        if v_chk is not None:
                            v_chk_clean = clean_label(v_chk)
                            if any(kw in v_chk_clean for kw in ['accionista', 'acciones', 'participacion', 'parent', 'holdco', 'adiciones', 'saldo final', 'saldo inicial', 'movimiento']):
                                has_share_hdr = True
                                break

                is_manual_disclosure = False
                MANUAL_HEADER_KEYWORDS = [
                    'traspaso', 'traspasos', 'vencimiento', 'vencimientos', 'sin descontar', 
                    'compromiso', 'compromisos', 'garantia', 'garantias', 'hipoteca', 'prenda', 
                    'litigio', 'litigios', 'contingencia', 'contingencias', 'tasa de interes', 
                    'vida util', 'vidas utiles'
                ]
                for r_hdr_chk in range(row - 1, max(0, row - 15), -1):
                    for c_lbl in range(1, min(6, ws.max_column + 1)):
                        v_hdr = ws.cell(row=r_hdr_chk, column=c_lbl).value
                        if v_hdr is not None:
                            v_hdr_clean = clean_label(v_hdr)
                            if any(kw in v_hdr_clean for kw in MANUAL_HEADER_KEYWORDS):
                                is_manual_disclosure = True
                                break
                    if is_manual_disclosure:
                        break

                if not has_date_hdr or has_share_hdr or is_manual_disclosure:
                    # Es una tabla con cabeceras propias, revelación manual o de acciones fijas -> PRESERVAR
                    continue

                # Si es una fila de datos de un bloque con fechas, limpiar valores obsoletos que no sean fórmulas
                for c in active_cols:
                    cell = ws.cell(row=row, column=c)
                    if cell.value is not None:
                        val_str = str(cell.value).strip()
                        if not val_str.startswith('='):
                            cell.value = None

            last_total_row = 0
            running_sums = {}

            # Rellenar filas de datos
            for row in range(1, ws.max_row + 1):
                # Validar fila de Total
                is_total_row = False
                total_col = None
                for c_idx in range(1, min(5, ws.max_column + 1)):
                    cell_val = ws.cell(row=row, column=c_idx).value
                    if cell_val is not None and isinstance(cell_val, str):
                        cell_clean = cell_val.strip().lower()
                        if cell_clean.startswith("total") or "saldo final" in cell_clean:
                            is_total_row = True
                            total_col = c_idx
                            break
                            
                if is_total_row:
                    for col_idx in range(1, ws.max_column + 1):
                        if col_idx > (total_col or 2) and col_idx in col_mapping:
                            # Calcular la suma leyendo las celdas en el rango (last_total_row + 1) a (row - 1)
                            sum_val = 0.0
                            has_vals_in_col = False
                            for r in range(last_total_row + 1, row):
                                val_cell = ws.cell(row=r, column=col_idx).value
                                if val_cell is not None:
                                    try:
                                        sum_val += round(float(val_cell))
                                        has_vals_in_col = True
                                    except:
                                        pass
                            
                            cell = ws.cell(row=row, column=col_idx)
                            # Conservar la fórmula en la celda si ya existe en la plantilla, de lo contrario escribir la suma solo si hay datos en la columna
                            if isinstance(cell.value, str) and cell.value.startswith('='):
                                pass
                            elif has_vals_in_col:
                                cell.value = round(sum_val)
                                cell.number_format = r'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
                    last_total_row = row
                    running_sums = {}
                    continue

                # Si la fila tiene textos de cabecera como 'M$', '$', '%' en alguna de sus columnas de datos, no la procesamos como datos
                is_header_row = False
                for c_idx in range(1, ws.max_column + 1):
                    val_c = ws.cell(row=row, column=c_idx).value
                    if val_c is not None and isinstance(val_c, str):
                        val_c_clean = val_c.strip().lower()
                        if val_c_clean in ['m$', '$', '%', 'miles']:
                            is_header_row = True
                            break
                if is_header_row:
                    continue

                sheet_title_lower = str(ws.title).strip().lower()

                # Buscar etiqueta descriptiva
                desc_col = None
                matched_key = None
                match_type = None

                # 1. Encontrar la columna del primer texto sustancial (la columna de etiqueta)
                for c_idx in range(1, min(6, ws.max_column + 1)):
                    cell_val = ws.cell(row=row, column=c_idx).value
                    if cell_val is not None and isinstance(cell_val, str):
                        cell_clean = clean_label(cell_val)
                        # Contar celdas de texto en las columnas de datos para identificar cabeceras apiladas
                        string_cells_count = 0
                        for c_temp in col_mapping.keys():
                            val_temp = ws.cell(row=row, column=c_temp).value
                            if val_temp is not None and isinstance(val_temp, str) and not val_temp.startswith('='):
                                try:
                                    float(val_temp)
                                except ValueError:
                                    string_cells_count += 1
                        
                        is_in_ctx = False
                        if entity_contexts:
                            for ent_dict in entity_contexts.values():
                                for p_ctx in [ent_dict.get('actual', {}), ent_dict.get('comp', {})]:
                                    if not p_ctx:
                                        continue
                                    if any(cell_clean in p_ctx.get(k, {}) for k in ['nota1', 'nota2']):
                                        is_in_ctx = True
                                        break
                                    if p_ctx.get('pl'):
                                        for r_data in p_ctx['pl'].values():
                                            if isinstance(r_data, dict) and cell_clean in r_data:
                                                is_in_ctx = True
                                                break
                                    if is_in_ctx:
                                        break
                                if is_in_ctx:
                                    break

                        # Identificar si la fila es un título de sección de P&L previo a la tabla de datos
                        is_section_header = False
                        if is_results_note and cell_clean in SECTION_TITLE_TO_PL_RUBRO:
                            for r_next in range(row + 1, min(row + 4, ws.max_row + 1)):
                                for c_chk in range(1, ws.max_column + 1):
                                    v_chk = ws.cell(row=r_next, column=c_chk).value
                                    if v_chk is not None:
                                        v_chk_str = str(v_chk).strip().lower()
                                        if isinstance(v_chk, (datetime.datetime, datetime.date)) or re.search(r'202\d', v_chk_str) or v_chk_str in ['m$', '$', '%']:
                                            is_section_header = True
                                            break
                                if is_section_header:
                                    break

                        is_intro_text = (
                            cell_clean.startswith(('a continuacion', 'los pagos', 'las notas', 'el siguiente cuadro', 'la composicion', 'al 31 de', 'durante el')) or
                            (len(cell_clean.split()) > 10 and len(cell_clean) > 60)
                        )
                        total_string_cells = sum(1 for c_temp in range(1, ws.max_column + 1) if ws.cell(row=row, column=c_temp).value is not None and isinstance(ws.cell(row=row, column=c_temp).value, str) and not str(ws.cell(row=row, column=c_temp).value).startswith('='))
                        is_header = False
                        if is_section_header:
                            is_header = True
                        elif not is_in_ctx:
                            if is_intro_text or total_string_cells >= 3 or cell_val.isupper():
                                is_header = True
                            elif cell_clean in HEADER_KEYWORDS or any(cell_clean.startswith(kw + ' ') for kw in HEADER_KEYWORDS) or string_cells_count > 1:
                                is_header = True
                        if is_header or not cell_clean or len(cell_clean) <= 2:
                            continue
                        desc_col = c_idx
                        break

                # 2. Si encontramos la columna de etiqueta, verificar si tiene correspondencia en los datos
                if desc_col is not None:
                    cell_val = ws.cell(row=row, column=desc_col).value
                    cell_clean = clean_label(cell_val)
                    
                    if is_results_note:
                        PL_ALIASES = {
                            'depreciacion operacional': 'depreciacion activos fijos de operacion',
                            'amortizacion derechos de uso': 'amortizacion por derechos de uso',
                            'costos arriendos ic': 'costos arriendos intercompania',
                            'depreciacion en gtos de adm': 'depreciacion del ejercicio',
                            'amortizacion intangibles': 'amortizacion de intangibles',
                            'costo financiero por instrumentos de cobertura': 'instrumentos de cobertura (2)',
                            'ganancia en instrumentos de cobertura*': 'instrumentos de cobertura (1)',
                            'ganancia en instrumentos de cobertura': 'instrumentos de cobertura (1)',
                        }
                        if cell_clean in PL_ALIASES:
                            alias_candidate = PL_ALIASES[cell_clean]
                            found_in_ctx = False
                            if entity_contexts:
                                for ent_dict in entity_contexts.values():
                                    for p_ctx in [ent_dict.get('actual', {}), ent_dict.get('comp', {})]:
                                        if p_ctx and p_ctx.get('pl'):
                                            for r_data in p_ctx['pl'].values():
                                                if isinstance(r_data, dict) and cell_clean in r_data:
                                                    found_in_ctx = True
                                                    break
                                        if found_in_ctx:
                                            break
                                    if found_in_ctx:
                                        break
                            if not found_in_ctx:
                                cell_clean = alias_candidate
                    
                    sample_context = None
                    for k in [active_entity_name, 'Consolidado'] + list(entity_contexts.keys()):
                        if k in entity_contexts:
                            act_c = entity_contexts[k].get('actual', {}) or {}
                            cmp_c = entity_contexts[k].get('comp', {}) or {}
                            merged_n1 = dict(cmp_c.get('nota1', {}))
                            merged_n1.update(act_c.get('nota1', {}))
                            merged_n2 = dict(cmp_c.get('nota2', {}))
                            merged_n2.update(act_c.get('nota2', {}))
                            merged_pl = {}
                            for p_src in [cmp_c.get('pl', {}), act_c.get('pl', {})]:
                                for r_k, r_v in (p_src or {}).items():
                                    if r_k not in merged_pl:
                                        merged_pl[r_k] = {}
                                    if isinstance(r_v, dict):
                                        merged_pl[r_k].update(r_v)
                            sample_context = {'nota1': merged_n1, 'nota2': merged_n2, 'pl': merged_pl}
                            break
                    
                    if sample_context:
                        matched_rubro = None
                        # 2.1 Búsqueda exacta
                        if is_results_note:
                            section_rubro = detect_row_section_rubro(ws, row)
                            # Obtener rubros permitidos para esta hoja
                            sheet_title = str(ws.title).strip()
                            allowed_rubros_for_sheet = SHEET_TO_PL_RUBROS.get(sheet_title, [])

                            if section_rubro:
                                rubro_data = sample_context.get('pl', {}).get(section_rubro, {})
                                if cell_clean in rubro_data:
                                    matched_key = cell_clean
                                    match_type = 'pl'
                                    matched_rubro = section_rubro
                            
                            if matched_key is None:
                                rubros_to_search = allowed_rubros_for_sheet if allowed_rubros_for_sheet else list(sample_context.get('pl', {}).keys())
                                for rubro in rubros_to_search:
                                    rubro_clean = ''.join(c for c in unicodedata.normalize('NFD', rubro.lower()) if unicodedata.category(c) != 'Mn')
                                    rubro_clean = re.sub(r'\s+', ' ', rubro_clean).strip()
                                    rubro_data = sample_context.get('pl', {}).get(rubro_clean, {})
                                    if cell_clean in rubro_data:
                                        matched_key = cell_clean
                                        match_type = 'pl'
                                        matched_rubro = rubro_clean
                                        break

                        sheet_title_lower = str(ws.title).strip().lower()
                        LABEL_ALIASES = {
                            'bodega central (*)': 'mercaderias',
                            'bodega central': 'mercaderias',
                            'otras bodegas (**)': 'mercaderias',
                            'bodega tecnica': 'mercaderias',
                            'bodega de contratistas': 'mercaderias',
                            'bodega de operaciones': 'mercaderias',
                            'licencias': 'licencias',
                            'derecho de uso (1)': 'derecho de uso (1)',
                            'saldo inicial': 'plusvalia por adquisiciones de empresas' if sheet_title_lower in ['plusvalia', 'plusvalía'] else 'saldo inicial',
                            'saldo final': 'plusvalia por adquisiciones de empresas' if sheet_title_lower in ['plusvalia', 'plusvalía'] else 'saldo final',
                            'plusvalia': 'plusvalia por adquisiciones de empresas',
                            'plusvalía': 'plusvalia por adquisiciones de empresas',
                            'capital emitido': 'capital emitido',
                            'otras reservas': 'otras reservas',
                            'resultados acumulados': 'ganancias (perdidas) acumuladas',
                            'ganancias perdidas acumuladas': 'ganancias (perdidas) acumuladas',
                        }
                        if matched_key is None and not is_results_note:
                            if cell_clean in sample_context.get('nota1', {}):
                                matched_key = cell_clean
                                match_type = 'nota1'
                            elif cell_clean in sample_context.get('nota2', {}):
                                matched_key = cell_clean
                                match_type = 'nota2'
                            elif cell_clean in LABEL_ALIASES:
                                alias_target = LABEL_ALIASES[cell_clean]
                                for ctx_t in ['nota1', 'nota2']:
                                    for k_ctx in sample_context.get(ctx_t, {}):
                                        if clean_label(k_ctx) == clean_label(alias_target) or k_ctx == alias_target:
                                            matched_key = k_ctx
                                            match_type = ctx_t
                                            break
                                    if matched_key is not None:
                                        break
                        
                        # 2.2 Búsqueda tolerante por prefijo/sufijo para cortes menores (máx 5 de diferencia)
                        if matched_key is None:
                            ctx_types = ['pl'] if is_results_note else ['nota1', 'nota2', 'pl']
                            for ctx_type in ctx_types:
                                if ctx_type == 'pl':
                                    section_rubro = detect_row_section_rubro(ws, row)
                                    if section_rubro:
                                        rubros_to_inspect = [section_rubro]
                                    else:
                                        sheet_title = str(ws.title).strip()
                                        allowed_rubros = SHEET_TO_PL_RUBROS.get(sheet_title, [])
                                        if not allowed_rubros:
                                            rubros_to_inspect = list(sample_context.get('pl', {}).keys())
                                        else:
                                            rubros_to_inspect = []
                                            for rubro in allowed_rubros:
                                                rubro_clean = ''.join(c for c in unicodedata.normalize('NFD', rubro.lower()) if unicodedata.category(c) != 'Mn')
                                                rubro_clean = re.sub(r'\s+', ' ', rubro_clean).strip()
                                                rubros_to_inspect.append(rubro_clean)
                                                
                                    for rubro in rubros_to_inspect:
                                        rubro_data = sample_context.get('pl', {}).get(rubro, {})
                                        for k_ctx in rubro_data.keys():
                                            if (k_ctx.startswith(cell_clean) or cell_clean.startswith(k_ctx)) and abs(len(k_ctx) - len(cell_clean)) <= 5:
                                                matched_key = k_ctx
                                                match_type = 'pl'
                                                matched_rubro = rubro
                                                break
                                        if matched_key is not None:
                                            break
                                else:
                                    for k_ctx in sample_context.get(ctx_type, {}):
                                        if {cell_clean, k_ctx} == {'euro', 'euros'}:
                                            continue
                                        if (k_ctx.startswith(cell_clean) or cell_clean.startswith(k_ctx)) and abs(len(k_ctx) - len(cell_clean)) <= 5:
                                            matched_key = k_ctx
                                            match_type = ctx_type
                                            break
                                if matched_key is not None:
                                    break

                        if sheet_title_lower in ['patrimonio', 'patrimonio']:
                            # Si la fila pertenece a la subtabla de estructura/movimiento de acciones, NO matchear cuentas contables monetarias
                            is_share_row = False
                            if row > 10:
                                is_share_row = True
                            else:
                                for c_chk in range(1, min(6, ws.max_column + 1)):
                                    v_chk = get_merged_cell_value(ws, row, c_chk)
                                    if v_chk is not None:
                                        v_chk_clean = clean_label(v_chk)
                                        if any(kw in v_chk_clean for kw in ['accionista', 'acciones', 'participacion', 'parent', 'holdco', 'adiciones', 'saldo final', 'saldo inicial', 'movimiento']):
                                            is_share_row = True
                                            break
                            if is_share_row and cell_clean not in ['capital emitido', 'otras reservas', 'resultados acumulados', 'total']:
                                matched_key = None

                if desc_col is not None and matched_key is not None:
                    extra_words = []
                    for c_idx in range(1, min(5, ws.max_column + 1)):
                        if c_idx != desc_col and c_idx not in col_mapping:
                            val_temp = ws.cell(row=row, column=c_idx).value
                            if val_temp is not None and isinstance(val_temp, str):
                                extra_words.append(clean_label(val_temp))
                                
                    # Obtener nombres de cuentas del contexto activo (o comparativo como respaldo) para validación de palabras extra
                    active_ctx = sample_context if sample_context else entity_contexts.get(active_entity_name, {}).get('actual', {})
                    acct_names = []
                    if active_ctx:
                        if match_type == 'nota1' and matched_key in active_ctx.get('nota1', {}):
                            acct_names = active_ctx['nota1'][matched_key].get('acct_names', [])
                        elif match_type == 'nota2' and matched_key in active_ctx.get('nota2', {}):
                            acct_names = active_ctx['nota2'][matched_key].get('acct_names', [])
                        elif match_type == 'pl':
                            sheet_title = str(ws.title).strip()
                            allowed_rubros = [matched_rubro] if matched_rubro else SHEET_TO_PL_RUBROS.get(sheet_title, [])
                            found_data = None
                            if not allowed_rubros:
                                for rubro_data in active_ctx.get('pl', {}).values():
                                    if matched_key in rubro_data:
                                        found_data = rubro_data[matched_key]
                                        break
                            else:
                                for rubro in allowed_rubros:
                                    rubro_clean = ''.join(c for c in unicodedata.normalize('NFD', rubro.lower()) if unicodedata.category(c) != 'Mn')
                                    rubro_clean = re.sub(r'\s+', ' ', rubro_clean).strip()
                                    rubro_data = active_ctx.get('pl', {}).get(rubro_clean, {})
                                    if matched_key in rubro_data:
                                        found_data = rubro_data[matched_key]
                                        break
                            if found_data:
                                acct_names = found_data.get('acct_names', [])

                    skip_row = False
                    for ew in extra_words:
                        if not check_extra_word_match(ew, acct_names):
                            skip_row = True
                            break
                    if skip_row:
                        continue

                    # Rellenar primero columnas individuales
                    for c in range(desc_col + 1, ws.max_column + 1):
                        col_info = col_mapping.get(c)
                        if not col_info or col_info['entity'] == 'Consolidado':
                            continue
                            
                        entity_ctx = find_best_context(col_info['entity'], entity_contexts)
                        if not entity_ctx:
                            continue
                            
                        # Intentar detectar el año del bloque de filas actual
                        block_yr = None if is_results_note else detect_block_year(ws, row, c)
                        if block_yr and block_yr in template_year_to_period_type:
                            period_type = template_year_to_period_type[block_yr]
                        else:
                            period_type = col_info.get('period_type') or template_year_to_period_type.get(col_info['year'], 'actual')
                        if is_no_comp_notes and period_type == 'comp':
                            continue
                        period_ctx = entity_ctx.get(period_type)
                        if not period_ctx:
                            continue
                            
                        # --- Validar si la columna pertenece al bloque de tabla activo de esta fila ---
                        has_header_in_block = False
                        if is_results_note:
                            has_header_in_block = True
                        else:
                            for r_idx in range(row - 1, 0, -1):
                                # Comprobar si cruzamos el límite de la subtabla
                                is_boundary = False
                                for c_lbl in range(1, min(5, ws.max_column + 1)):
                                    cell_lbl = ws.cell(row=r_idx, column=c_lbl).value
                                    if cell_lbl is not None:
                                        cell_lbl_clean = clean_label(cell_lbl)
                                        if ("total" in cell_lbl_clean and "sub" not in cell_lbl_clean) or "transaccion" in cell_lbl_clean or "las principales" in cell_lbl_clean or "la composicion" in cell_lbl_clean or "saldo final" in cell_lbl_clean or "a continuacion" in cell_lbl_clean or "los pagos" in cell_lbl_clean:
                                            is_boundary = True
                                            break
                                if is_boundary:
                                    break
                                
                                val_h = get_merged_cell_value(ws, r_idx, c)
                                if val_h is None and c in col_mapping:
                                    # Solo buscar en columnas adyacentes si la columna 'c' tiene alguna celda de encabezado en este bloque
                                    has_any_hdr_in_c = any(
                                        ws.cell(row=chk_r, column=c).value is not None 
                                        for chk_r in range(r_idx, row)
                                    )
                                    if has_any_hdr_in_c:
                                        for c_adj in [c + 1, c + 2, c - 1, c - 2]:
                                            if 1 <= c_adj <= ws.max_column:
                                                val_adj_h = get_merged_cell_value(ws, r_idx, c_adj)
                                                if val_adj_h is not None and (isinstance(val_adj_h, (datetime.datetime, datetime.date)) or re.search(r'(202\d)', str(val_adj_h))):
                                                    val_h = val_adj_h
                                                    break
                                if val_h is not None:
                                    if isinstance(val_h, (datetime.datetime, datetime.date)) or re.search(r'(202\d)', str(val_h)):
                                        has_header_in_block = True
                                        break
                        
                        if not has_header_in_block:
                            continue

                        is_manual_disc_block = False
                        MANUAL_HEADER_KEYWORDS = [
                            'traspaso', 'traspasos', 'vencimiento', 'vencimientos', 'sin descontar', 
                            'compromiso', 'compromisos', 'garantia', 'garantias', 'hipoteca', 'prenda', 
                            'litigio', 'litigios', 'contingencia', 'contingencias', 'tasa de interes', 
                            'vida util', 'vidas utiles'
                        ]
                        for r_hdr_chk in range(row - 1, max(0, row - 15), -1):
                            for c_lbl in range(1, min(6, ws.max_column + 1)):
                                v_hdr = ws.cell(row=r_hdr_chk, column=c_lbl).value
                                if v_hdr is not None:
                                    v_hdr_clean = clean_label(v_hdr)
                                    if any(kw in v_hdr_clean for kw in MANUAL_HEADER_KEYWORDS):
                                        is_manual_disc_block = True
                                        break
                            if is_manual_disc_block:
                                break

                        if is_manual_disc_block:
                            continue
                            
                        col_headers = []
                        for r_idx in range(1, row):
                            h_val = get_merged_cell_value(ws, r_idx, c)
                            if h_val is not None:
                                col_headers.append(str(h_val).strip())
                                
                        is_gross = any("bruto" in h.lower() or "costo" in h.lower() for h in col_headers)
                        is_depr = any("depreciaci" in h.lower() or "acumulada" in h.lower() or "amortizaci" in h.lower() for h in col_headers)
                        is_net = any("neto" in h.lower() for h in col_headers)
                        
                        # Detectar el flujo según el texto de cabecera
                        is_inicial = any("inicial" in h.lower() or "apertura" in h.lower() or "comienzo" in h.lower() for h in col_headers)
                        is_debits = any("debit" in h.lower() or "debe" in h.lower() or "cargo" in h.lower() or "adicion" in h.lower() or "aumento" in h.lower() for h in col_headers)
                        is_credits = any("credit" in h.lower() or "haber" in h.lower() or "abono" in h.lower() or "retiro" in h.lower() or "disminuc" in h.lower() for h in col_headers)
                        
                        flow_key = 'val'
                        if is_inicial:
                            flow_key = 'inicial'
                        elif is_debits:
                            flow_key = 'debitos'
                        elif is_credits:
                            flow_key = 'creditos'
                            
                        val_to_write = None
                        if is_gross:
                            val_to_write = get_filtered_value(period_ctx, matched_key, 'nota1', flow_key, row, ws, filter_sign='gross', col=c) / scale_factor
                        elif is_depr:
                            val_to_write = get_filtered_value(period_ctx, matched_key, 'nota2', flow_key, row, ws, filter_sign='depr', col=c) / scale_factor
                        elif is_net:
                            val_to_write = (get_filtered_value(period_ctx, matched_key, 'nota1', flow_key, row, ws, filter_sign='gross', col=c) + 
                                            get_filtered_value(period_ctx, matched_key, 'nota2', flow_key, row, ws, filter_sign='depr', col=c)) / scale_factor
                        else:
                            if match_type == 'nota1':
                                val_to_write = get_filtered_value(period_ctx, matched_key, 'nota1', flow_key, row, ws, col=c) / scale_factor
                                # Fallback a nota2 si el resultado da cero
                                if val_to_write == 0.0:
                                    val_to_write = get_filtered_value(period_ctx, matched_key, 'nota2', flow_key, row, ws, col=c) / scale_factor
                            elif match_type == 'nota2':
                                val_to_write = get_filtered_value(period_ctx, matched_key, 'nota2', flow_key, row, ws, col=c) / scale_factor
                                # Fallback a nota1 si el resultado da cero
                                if val_to_write == 0.0:
                                    val_to_write = get_filtered_value(period_ctx, matched_key, 'nota1', flow_key, row, ws, col=c) / scale_factor
                            elif match_type == 'pl':
                                val_to_write = get_filtered_value(period_ctx, matched_key, 'pl', flow_key, row, ws, section_rubro=matched_rubro, col=c) / scale_factor
                                
                        if val_to_write is not None:
                            val_to_write = round(val_to_write)
                            cell = ws.cell(row=row, column=c)
                            if cell.__class__.__name__ == 'MergedCell':
                                continue
                            if not (isinstance(cell.value, str) and cell.value.startswith('=')):
                                cell.value = val_to_write
                                cell.number_format = r'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
                            running_sums[c] = running_sums.get(c, 0.0) + val_to_write

                    # Rellenar columnas de Consolidado por suma de las individuales o fallback
                    for c in range(desc_col + 1, ws.max_column + 1):
                        col_info = col_mapping.get(c)
                        if not col_info or col_info['entity'] != 'Consolidado':
                            continue
                            
                        block_yr = detect_block_year(ws, row, c)
                        if block_yr and block_yr in template_year_to_period_type:
                            period_type = template_year_to_period_type[block_yr]
                        else:
                            period_type = col_info.get('period_type') or template_year_to_period_type.get(col_info['year'], 'actual')
                        if is_no_comp_notes and period_type == 'comp':
                            continue
                        
                        # Buscar columnas individuales en esta hoja que tengan el mismo periodo
                        indiv_cols = [col_idx for col_idx, info in col_mapping.items()                                       if info['entity'] not in ['Consolidado', None] and (info.get('period_type') or template_year_to_period_type.get(info['year'], 'actual')) == period_type]
                                       
                        if indiv_cols:
                            sum_val = 0.0
                            for col_idx in indiv_cols:
                                val_cell = ws.cell(row=row, column=col_idx).value
                                if val_cell is not None:
                                    try:
                                        sum_val += round(float(val_cell))
                                    except:
                                        pass
                            sum_val = round(sum_val)
                            cell = ws.cell(row=row, column=c)
                            if cell.__class__.__name__ == 'MergedCell':
                                continue
                            if not (isinstance(cell.value, str) and cell.value.startswith('=')):
                                cell.value = sum_val
                                cell.number_format = r'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
                            running_sums[c] = running_sums.get(c, 0.0) + sum_val
                        else:
                            # Fallback: usar contexto consolidado/activo directo
                            entity_ctx = find_best_context(active_entity_name, entity_contexts)
                            if entity_ctx:
                                period_ctx = entity_ctx.get(period_type)
                                if period_ctx:
                                    col_headers = []
                                    for r_idx in range(1, row):
                                        h_val = ws.cell(row=r_idx, column=c).value
                                        if h_val is not None:
                                            col_headers.append(str(h_val).strip())
                                            
                                    is_gross = any("bruto" in h.lower() or "costo" in h.lower() for h in col_headers)
                                    is_depr = any("depreciaci" in h.lower() or "acumulada" in h.lower() or "amortizaci" in h.lower() for h in col_headers)
                                    is_net = any("neto" in h.lower() for h in col_headers)
                                    
                                    val_to_write = None
                                    if is_gross:
                                        val_to_write = period_ctx['nota1'].get(matched_key, {}).get('val', 0.0) / scale_factor
                                    elif is_depr:
                                        val_to_write = period_ctx['nota2'].get(matched_key, {}).get('val', 0.0) / scale_factor
                                    elif is_net:
                                        val_to_write = (period_ctx['nota1'].get(matched_key, {}).get('val', 0.0) + period_ctx['nota2'].get(matched_key, {}).get('val', 0.0)) / scale_factor
                                    else:
                                        if match_type == 'nota1':
                                            val_to_write = period_ctx['nota1'].get(matched_key, {}).get('val', 0.0) / scale_factor
                                            # Fallback a nota2 si el resultado da cero
                                            if val_to_write == 0.0:
                                                val_to_write = period_ctx['nota2'].get(matched_key, {}).get('val', 0.0) / scale_factor
                                        elif match_type == 'nota2':
                                            val_to_write = period_ctx['nota2'].get(matched_key, {}).get('val', 0.0) / scale_factor
                                            # Fallback a nota1 si el resultado da cero
                                            if val_to_write == 0.0:
                                                val_to_write = period_ctx['nota1'].get(matched_key, {}).get('val', 0.0) / scale_factor
                                        elif match_type == 'pl':
                                            val_to_write = get_filtered_value(period_ctx, matched_key, 'pl', 'val', row, ws, section_rubro=matched_rubro) / scale_factor
                                            
                                    if val_to_write is not None:
                                        val_to_write = round(val_to_write)
                                        cell = ws.cell(row=row, column=c)
                                        if cell.__class__.__name__ == 'MergedCell':
                                            continue
                                        if not (isinstance(cell.value, str) and cell.value.startswith('=')):
                                            cell.value = val_to_write
                                            cell.number_format = r'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
                                        running_sums[c] = running_sums.get(c, 0.0) + val_to_write
            # --- Regla especial para la nota de Deudores (Cuadro #N06.3 - Movimiento de Incobrables) ---
            if sheet_name == 'Deudores':
                r_ini = None
                r_per = None
                r_rev = None
                r_cas = None
                r_fin = None
                
                # Buscamos a partir de la fila 35 para enfocarnos en el cuadro #N06.3
                for r in range(35, min(ws.max_row + 1, 60)):
                    val_c = ws.cell(row=r, column=2).value
                    if val_c is not None and isinstance(val_c, str):
                        clean_lbl = clean_label(val_c)
                        if 'saldo inicial' in clean_lbl:
                            r_ini = r
                        elif 'perdida reconocida por deterioro' in clean_lbl or 'perdida reconocida' in clean_lbl:
                            r_per = r
                        elif 'reverso de deterioro' in clean_lbl:
                            r_rev = r
                        elif 'castigos' in clean_lbl:
                            r_cas = r
                        elif 'saldo final' in clean_lbl:
                            r_fin = r
                
                if r_ini is not None and r_fin is not None:
                    c_act = None
                    c_comp = None
                    for col_idx, info in col_mapping.items():
                        if info['entity'] == active_entity_name:
                            p_type = info.get('period_type') or template_year_to_period_type.get(info['year'])
                            if p_type == 'actual':
                                c_act = col_idx
                            elif p_type == 'comp':
                                c_comp = col_idx
                                
                    entity_ctx = find_best_context(active_entity_name, entity_contexts)
                    val_act = 0.0
                    val_comp = 0.0
                    if entity_ctx:
                        ctx_act = entity_ctx.get('actual', {})
                        ctx_cmp = entity_ctx.get('comp', {})
                        
                        for k_ctx in ctx_act.get('nota1', {}):
                            if 'estimacion de incobrables' in k_ctx or 'estimacion de incobr' in k_ctx:
                                val_act = abs(ctx_act['nota1'][k_ctx].get('val', 0.0)) / scale_factor
                                break
                        for k_ctx in ctx_cmp.get('nota1', {}):
                            if 'estimacion de incobrables' in k_ctx or 'estimacion de incobr' in k_ctx:
                                val_comp = abs(ctx_cmp['nota1'][k_ctx].get('val', 0.0)) / scale_factor
                                break
                                
                    if c_act is not None:
                        saldo_fin_act = round(val_act)
                        saldo_ini_act = round(val_comp)
                        
                        loss_act = max(0.0, saldo_fin_act - saldo_ini_act)
                        reverso_act = max(0.0, saldo_ini_act - saldo_fin_act)
                        castigos_act = 0.0
                        
                        ws.cell(row=r_ini, column=c_act).value = saldo_ini_act
                        ws.cell(row=r_ini, column=c_act).number_format = r'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
                        
                        if r_per is not None:
                            ws.cell(row=r_per, column=c_act).value = loss_act
                            ws.cell(row=r_per, column=c_act).number_format = r'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
                        if r_rev is not None:
                            ws.cell(row=r_rev, column=c_act).value = reverso_act
                            ws.cell(row=r_rev, column=c_act).number_format = r'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
                        if r_cas is not None:
                            ws.cell(row=r_cas, column=c_act).value = castigos_act
                            ws.cell(row=r_cas, column=c_act).number_format = r'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
                            
                        ws.cell(row=r_fin, column=c_act).value = saldo_fin_act
                        ws.cell(row=r_fin, column=c_act).number_format = r'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
                        
                    if c_comp is not None and not is_no_comp_notes:
                        saldo_fin_comp = round(val_comp)
                        saldo_ini_comp = 0.0
                        
                        loss_comp = max(0.0, saldo_fin_comp - saldo_ini_comp)
                        reverso_comp = max(0.0, saldo_ini_comp - saldo_fin_comp)
                        castigos_comp = 0.0
                        
                        ws.cell(row=r_ini, column=c_comp).value = saldo_ini_comp
                        ws.cell(row=r_ini, column=c_comp).number_format = r'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
                        
                        if r_per is not None:
                            ws.cell(row=r_per, column=c_comp).value = loss_comp
                            ws.cell(row=r_per, column=c_comp).number_format = r'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
                        if r_rev is not None:
                            ws.cell(row=r_rev, column=c_comp).value = reverso_comp
                            ws.cell(row=r_rev, column=c_comp).number_format = r'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
                        if r_cas is not None:
                            ws.cell(row=r_cas, column=c_comp).value = castigos_comp
                            ws.cell(row=r_cas, column=c_comp).number_format = r'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
                            
                        ws.cell(row=r_fin, column=c_comp).value = saldo_fin_comp
                        ws.cell(row=r_fin, column=c_comp).number_format = r'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'

            # --- Regla especial para la nota de Intangibles (Cuadro #N08.3 - Movimiento de Intangibles) ---
            if sheet_name == 'Intangibles':
                r_ini = None
                r_adi = None
                r_amo = None
                r_fin = None
                r_hdr = None
                
                # Buscar las filas en el cuadro #N08.3
                for r in range(25, min(ws.max_row + 1, 45)):
                    val_c = ws.cell(row=r, column=3).value
                    if val_c is not None and isinstance(val_c, str):
                        clean_lbl = clean_label(val_c)
                        if 'inicio del periodo' in clean_lbl or 'inicial' in clean_lbl:
                            r_ini = r
                        elif 'adiciones' in clean_lbl:
                            r_adi = r
                        elif 'amortizacion' in clean_lbl:
                            r_amo = r
                        elif 'final del periodo' in clean_lbl or 'saldo final' in clean_lbl:
                            r_fin = r
                    
                    if r_hdr is None:
                        row_vals = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
                        if any(val and isinstance(val, str) and clean_label(val) in ['licencias', 'concesiones', 'licencia'] for val in row_vals):
                            r_hdr = r

                if r_hdr is not None:
                    entity_ctx = find_best_context(active_entity_name, entity_contexts)
                    ctx_act = entity_ctx.get('actual', {}) if entity_ctx else None
                    ctx_cmp = entity_ctx.get('comp', {}) if entity_ctx else None
                    
                    def find_ctx_key(clean_hdr, ctx):
                        if not ctx:
                            return None
                        for k in ctx.get('nota1', {}).keys():
                            if k == clean_hdr or k.startswith(clean_hdr) or clean_hdr.startswith(k):
                                return k
                        for k in ctx.get('nota2', {}).keys():
                            if k == clean_hdr or k.startswith(clean_hdr) or clean_hdr.startswith(k):
                                return k
                        return None

                    # Rellenar cada columna que tenga una clase de intangibles
                    for c in range(1, ws.max_column + 1):
                        hdr = ws.cell(row=r_hdr, column=c).value
                        if not hdr or 'total' in str(hdr).lower() or 'detalle' in str(hdr).lower() or 'm$' in str(hdr).lower():
                            continue
                        clean_hdr = clean_label(hdr)
                        if len(clean_hdr) <= 2:
                            continue
                            
                        # Buscar la clave correspondiente en los contextos
                        key_act = find_ctx_key(clean_hdr, ctx_act)
                        key_cmp = find_ctx_key(clean_hdr, ctx_cmp)
                        
                        key_cmp = key_cmp or key_act
                        key_act = key_act or key_cmp
                        
                        if not key_act:
                            continue
                            
                        # Valores Iniciales (Diciembre 2025)
                        gross_ini = get_filtered_value(ctx_cmp, key_cmp, 'nota1', 'val', r_ini or r_hdr, ws) if ctx_cmp else 0.0
                        depr_ini = get_filtered_value(ctx_cmp, key_cmp, 'nota2', 'val', r_ini or r_hdr, ws) if ctx_cmp else 0.0
                        net_ini = round((gross_ini + depr_ini) / scale_factor)
                        
                        # Valores Finales (Mayo 2026 / Junio 2026)
                        gross_fin = get_filtered_value(ctx_act, key_act, 'nota1', 'val', r_fin or r_hdr, ws) if ctx_act else 0.0
                        depr_fin = get_filtered_value(ctx_act, key_act, 'nota2', 'val', r_fin or r_hdr, ws) if ctx_act else 0.0
                        net_fin = round((gross_fin + depr_fin) / scale_factor)
                        
                        # Adiciones: max(0, gross_fin - gross_ini)
                        adiciones = round(max(0.0, (gross_fin - gross_ini) / scale_factor))
                        
                        # Amortización del ejercicio: net_fin - net_ini - adiciones
                        amortizacion = round(net_fin - net_ini - adiciones)
                        
                        # Escribir en las celdas
                        if r_ini is not None:
                            cell = ws.cell(row=r_ini, column=c)
                            if not (isinstance(cell.value, str) and cell.value.startswith('=')):
                                cell.value = net_ini
                                cell.number_format = r'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
                        if r_adi is not None:
                            cell = ws.cell(row=r_adi, column=c)
                            if not (isinstance(cell.value, str) and cell.value.startswith('=')):
                                cell.value = adiciones
                                cell.number_format = r'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
                        if r_amo is not None:
                            cell = ws.cell(row=r_amo, column=c)
                            if not (isinstance(cell.value, str) and cell.value.startswith('=')):
                                cell.value = amortizacion
                                cell.number_format = r'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
                        if r_fin is not None:
                            cell = ws.cell(row=r_fin, column=c)
                            if not (isinstance(cell.value, str) and cell.value.startswith('=')):
                                cell.value = net_fin
                                cell.number_format = r'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'

            # 4. Filtrar columnas si es reporte individual
            from openpyxl.utils import get_column_letter
            col_letter_map = {get_column_letter(c): get_column_letter(c) for c in range(1, ws.max_column + 1)}
            
            if not is_consolidated:
                def _is_same_entity(e1, e2):
                    if not e1 or not e2:
                        return True
                    return normalize_company_name(e1) == normalize_company_name(e2)

                # Solo filtrar si hay columnas que pertenecen a otra empresa que no sea la activa ni Consolidado
                has_other_entities = any(
                    info['entity'] is not None and not _is_same_entity(info['entity'], active_entity_name) and info['entity'] != 'Consolidado'
                    for info in col_mapping.values()
                )
                
                if has_other_entities:
                    columns_to_keep = {1, 2}
                    for c in range(3, ws.max_column + 1):
                        info = col_mapping.get(c)
                        # Mantener si pertenece a la empresa activa, o si no se tiene mapeado de entidad
                        if info and (_is_same_entity(info['entity'], active_entity_name) or info['entity'] is None):
                            columns_to_keep.add(c)
                            
                    # Construir col_letter_map con los índices nuevos resultantes
                    col_letter_map = {}
                    for new_idx, orig_idx in enumerate(sorted(list(columns_to_keep)), 1):
                        orig_letter = get_column_letter(orig_idx)
                        new_letter = get_column_letter(new_idx)
                        col_letter_map[orig_letter] = new_letter
                        
                    for c in range(ws.max_column, 1, -1):
                        if c not in columns_to_keep:
                            ws.delete_cols(c)

            # Duplicar cuadros marcados con [COMPARATIVO] para el período comparativo únicamente si evaluate_formulas es True.
            # Si evaluate_formulas es False, la duplicación se realiza en ExternalNotesManager para inyectar datos externos con ext_cmp_ws/ext_act_ws.
            if evaluate_formulas:
                process_comparative_tables_in_ws(ws, periodo_actual_str=periodo_actual_str, periodo_comp_str=periodo_comp_str)

            # Corregir letras de columna en las fórmulas después de reubicar/filtrar columnas
            fix_formulas_column_letters(ws, col_letter_map)

            if str(target_lang).lower() == 'en':
                from src.core.ifrs_glossary import translate_ifrs_term
                for r in range(1, ws.max_row + 1):
                    for c in range(1, ws.max_column + 1):
                        cell_val = ws.cell(row=r, column=c).value
                        if cell_val and isinstance(cell_val, str) and not cell_val.startswith('='):
                            ws.cell(row=r, column=c).value = translate_ifrs_term(cell_val.strip(), target_lang='en')

        # 5. Guardar libro resultante
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        if evaluate_formulas:
            try:
                from src.ui_pages.informes_y_notas import evaluate_formulas_in_workbook
                output = evaluate_formulas_in_workbook(output)
            except Exception:
                pass
            
        return output
