"""
Motor Central de Redondeo Jerárquico por Mayor Residuo y Reconciliación EEFF -> Notas (M$).

Reglas de Negocio Auditables:
1. Nivel 1 (Estados Financieros mandan siempre):
   - Al convertir desde Pesos ($) a Miles de Pesos (M$, scale_factor == 1000.0), se aplica el
     Método del Mayor Residuo (Largest Remainder / Hamilton) preservando la suma exacta,
     sin modificar jamás líneas en cero ni invertir signos contables.
   - Tope estricto de seguridad: si el descuadre base en pesos supera $4.000 (> 4 M$),
     NO se aplica cuadratura forzada para que el módulo de Validaciones detecte el error real.

2. Nivel 2 y 3 (Notas y Sub-cuadros se ajustan al Estado Financiero):
   - Cuando la diferencia entre el Total de un cuadro de Nota (#N04.1 .. #N26.2) y su rubro
     oficial en el Balance / Estado de Resultados está entre ±1 y ±4 M$ (1 <= |diff| <= 4),
     el ajuste se aplica sobre la celda de entrada elegible de MAYOR valor absoluto (|valor|)
     que NO esté vinculada a otra nota/sub-cuadro, preservando intactas las fórmulas =SUM(...)
     del archivo Excel descargable.
   - Si |diff| > 4 M$, el cuadro no se toca para exponer el descuadre real en Validaciones.
"""

from __future__ import annotations

import copy
import math
import os
import re
import unicodedata
from io import BytesIO
from typing import Any, Dict, List, Optional, Set, Tuple

import openpyxl
import openpyxl.utils
import pandas as pd


MAX_TOLERANCE_PESOS: float = 4000.0
MAX_TOLERANCE_M: int = 4


def _norm_label(val: Any) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val).strip().lower()
    if s in ("nan", "nat", "none"):
        return ""
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    s = re.sub(r"\(\d+\)$", "", s).strip()
    s = s.rstrip("*.,:; ")
    return re.sub(r"\s+", " ", s)


def is_pl_classification_label(label: str) -> bool:
    """Identifica si una etiqueta corresponde a una línea de Estado de Resultados (P&L)."""
    norm = _norm_label(label)
    if not norm:
        return False
    pl_keywords = (
        "ingreso", "gasto", "costo", "egreso", "depreciacion", "amortizacion",
        "diferencia de cambio", "diferencias de cambio", "unidad de reajuste",
        "unidades de reajuste", "ganancia", "perdida", "resultado por",
    )
    bs_exceptions = (
        "diferido", "diferidos", "anticipado", "anticipados", "por pagar",
        "por cobrar", "acumulado", "acumulados", "acumulada", "acumuladas",
        "capital", "reserva", "activo", "pasivo", "plusvalia", "inventario",
    )
    if any(k in norm for k in pl_keywords) and not any(b in norm for b in bs_exceptions):
        return True
    return False


def largest_remainder_round(
    raw_pesos: Dict[str, float],
    target_total_m: int,
    scale_factor: float = 1000.0,
    max_tol_pesos: float = MAX_TOLERANCE_PESOS,
    locked_keys: Optional[Set[str]] = None,
) -> Tuple[Dict[str, int], bool]:
    """
    Convierte un diccionario {clave: monto_en_pesos} a enteros en M$ (scale_factor=1000.0)
    de modo que sum(resultado.values()) == target_total_m usando el Método del Mayor Residuo.

    Garantías:
    - Si |sum(raw_pesos) - target_total_m * scale_factor| > max_tol_pesos: NO ajusta (retorna
      redondeo estándar y ok=False).
    - Las claves con valor 0.0 en pesos permanecen estrictamente en 0.
    - Las claves en `locked_keys` conservan su `int(round(val / scale_factor))` fijo (útil para
      inmovilizar Efectivo, Plusvalía, Capital o Reservas sin movimiento).
    - Ninguna línea invierte su signo contable.
    """
    if not raw_pesos:
        return {}, (target_total_m == 0)

    if abs(float(scale_factor) - 1000.0) >= 0.01:
        simple = {k: int(round(float(v or 0.0) / scale_factor)) for k, v in raw_pesos.items()}
        return simple, True

    clean_pesos: Dict[str, float] = {
        k: (0.0 if (v is None or (isinstance(v, float) and pd.isna(v))) else float(v))
        for k, v in raw_pesos.items()
    }
    raw_sum = sum(clean_pesos.values())
    target_pesos = float(target_total_m) * scale_factor

    # Guardarraíl estricto: > $4.000 no se fuerza cuadratura automática
    if abs(raw_sum - target_pesos) > max_tol_pesos + 1e-6:
        fallback = {k: int(round(v / scale_factor)) for k, v in clean_pesos.items()}
        return fallback, False

    locked = locked_keys or set()
    result: Dict[str, int] = {}
    adjustable_keys: List[str] = []

    for k, v in clean_pesos.items():
        if abs(v) < 1e-9:
            result[k] = 0
        elif k in locked:
            result[k] = int(round(v / scale_factor))
        else:
            adjustable_keys.append(k)

    locked_sum = sum(result.values())
    remaining_target = int(target_total_m - locked_sum)

    if not adjustable_keys:
        return result, (remaining_target == 0)

    floored: Dict[str, int] = {}
    remainders: Dict[str, float] = {}
    for k in adjustable_keys:
        scaled = clean_pesos[k] / scale_factor
        fl = math.floor(scaled)
        floored[k] = fl
        remainders[k] = scaled - fl

    sum_floored = sum(floored.values())
    needed = int(remaining_target - sum_floored)

    if 0 <= needed <= len(adjustable_keys):
        # Caso estándar de Hamilton / Mayor Residuo: cada línea varía < 1 M$ y jamás invierte signo
        ranked = sorted(
            adjustable_keys,
            key=lambda k: (remainders[k], abs(clean_pesos[k]), k),
            reverse=True,
        )
        plus_one = set(ranked[:needed])
        for k in adjustable_keys:
            result[k] = floored[k] + (1 if k in plus_one else 0)
        return result, True

    # Caso borde dentro del umbral <= $4.000 (ej. residuo base de $1.500 o varias claves bloqueadas):
    # redondear normalmente y llevar el delta residual (<= 4 M$) a la partida ajustable de mayor magnitud
    for k in adjustable_keys:
        result[k] = int(round(clean_pesos[k] / scale_factor))

    residual_m = int(target_total_m - sum(result.values()))
    if residual_m == 0:
        return result, True

    if abs(residual_m) <= MAX_TOLERANCE_M:
        candidates = [
            k for k in adjustable_keys
            if result[k] != 0
            and ((result[k] > 0 and result[k] + residual_m > 0) or (result[k] < 0 and result[k] + residual_m < 0))
        ]
        if candidates:
            best_k = max(candidates, key=lambda k: (abs(result[k]), abs(clean_pesos[k])))
            result[best_k] += residual_m
            return result, True

    return result, (sum(result.values()) == target_total_m)


# =============================================================================
# EXTRACCIÓN DE METAS EEFF Y RECONCILIACIÓN DE CUADROS DE NOTAS EN OPENPYXL
# =============================================================================

def _parse_int_m(val: Any) -> Tuple[int, bool]:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return 0, False
    if isinstance(val, bool):
        return 0, False
    if isinstance(val, (int, float)):
        return int(round(float(val))), True
    s = str(val).strip()
    if not s or s in ("-", "—", "–", "nan", "None", "null"):
        return 0, True
    is_neg = False
    if s.startswith("(") and s.endswith(")"):
        is_neg = True
        s = s[1:-1].strip()
    elif s.startswith("-"):
        is_neg = True
        s = s[1:].strip()
    s = s.replace("$", "").replace("M", "").replace(" ", "")
    if not s:
        return 0, False
    if "." in s and "," in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "." in s:
        parts = s.split(".")
        if len(parts) > 2 or (len(parts) == 2 and len(parts[1]) == 3):
            s = s.replace(".", "")
    elif "," in s:
        parts = s.split(",")
        if len(parts) > 2 or (len(parts) == 2 and len(parts[1]) == 3):
            s = s.replace(",", "")
        else:
            s = s.replace(",", ".")
    try:
        num = int(round(float(s)))
        return (-num if is_neg else num), True
    except (ValueError, TypeError):
        return 0, False


def _extract_stmt_target(df_stmt: Optional[pd.DataFrame], periodo: str, aliases: List[str]) -> Tuple[int, bool]:
    if df_stmt is None or df_stmt.empty:
        return 0, False
    label_col = df_stmt.columns[0]
    p_norm = _norm_label(periodo)
    year = periodo.split("-")[0] if "-" in str(periodo) else str(periodo)[:4]
    data_cols = [
        c for c in df_stmt.columns[1:]
        if _norm_label(c) not in ("nota", "notas", "note", "notes", "id_nota", "codigo")
    ]
    val_col = None
    for c in data_cols:
        if p_norm and p_norm in _norm_label(c):
            val_col = c
            break
    if val_col is None and year:
        for c in data_cols:
            if year in str(c):
                val_col = c
                break
    if val_col is None and data_cols:
        val_col = data_cols[0]
    if val_col is None:
        return 0, False

    norm_aliases = [_norm_label(a) for a in aliases]
    for a_norm in norm_aliases:
        for _, row in df_stmt.iterrows():
            if _norm_label(row[label_col]) == a_norm:
                v, ok = _parse_int_m(row[val_col])
                return (v if ok else 0), True
    for a_norm in norm_aliases:
        for _, row in df_stmt.iterrows():
            r_lbl = _norm_label(row[label_col])
            if r_lbl and a_norm in r_lbl and "total" not in r_lbl:
                v, ok = _parse_int_m(row[val_col])
                return (v if ok else 0), True
    return 0, False


def _period_tokens(periodo: str) -> List[str]:
    if not periodo or "-" not in str(periodo):
        return [str(periodo)]
    parts = str(periodo).split("-")
    year, month = parts[0], parts[1].zfill(2)
    tokens = [f".{month}.{year}", f"/{month}/{year}", f"{year}-{month}", year]
    m_names = {
        "01": "enero", "02": "febrero", "03": "marzo", "04": "abril",
        "05": "mayo", "06": "junio", "07": "julio", "08": "agosto",
        "09": "septiembre", "10": "octubre", "11": "noviembre", "12": "diciembre",
    }
    if month in m_names:
        tokens.insert(2, f"{m_names[month]} de {year}")
        tokens.insert(3, f"{m_names[month]} {year}")
    return tokens


def _find_subcode_blocks_in_ws(ws: openpyxl.worksheet.worksheet.Worksheet, sub_code: str) -> List[Tuple[int, int]]:
    """
    Localiza los rangos de filas (start_row, end_row) en una hoja de openpyxl pertenecientes a `sub_code`
    (ej. '#N04.1', '#N21.2'). Soporta tanto códigos explícitos en celdas como índice secuencial.
    """
    code_pat = re.compile(r"#(N\d+(?:\.\d+)?)", re.IGNORECASE)
    markers: List[Tuple[int, str]] = []
    for r in range(1, ws.max_row + 1):
        for c in range(1, min(ws.max_column + 1, 8)):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and "#" in v:
                m = code_pat.search(v.strip())
                if m:
                    markers.append((r, f"#{m.group(1).upper()}"))
                    break

    target_upper = sub_code.upper()
    blocks: List[Tuple[int, int]] = []
    if markers:
        for idx, (r_start, m_code) in enumerate(markers):
            r_end = markers[idx + 1][0] - 1 if idx + 1 < len(markers) else ws.max_row
            if m_code == target_upper:
                blocks.append((r_start, r_end))
    if not blocks:
        blocks.append((1, ws.max_row))
    return blocks


def _locate_total_cell_in_ws(
    ws_eval: openpyxl.worksheet.worksheet.Worksheet,
    sub_code: str,
    periodo: str,
    row_labels: List[str],
    col_subheader: Optional[str] = None,
    after_section_label: Optional[str] = None,
    allow_blank_total_label: bool = False,
) -> Optional[Tuple[int, int, int, int]]:
    """
    Localiza la celda de Total `(tot_row, tot_col, block_start, block_end)` dentro de `ws_eval`
    para el cuadro `sub_code` y el período actual `periodo`.
    """
    blocks = _find_subcode_blocks_in_ws(ws_eval, sub_code)
    if not blocks:
        return None

    r_start, r_end = blocks[0]
    tokens = _period_tokens(periodo)
    max_col = ws_eval.max_column

    # 1. Detectar filas de cabecera (las primeras filas hasta la primera fila con datos o hasta 8 filas)
    hdr_end = min(r_start + 9, r_end)

    # Propagar etiquetas de período hacia la derecha sobre celdas vacías de cabecera (tablas multicolumna)
    period_cols: List[int] = []
    for c in range(1, max_col + 1):
        col_texts: List[str] = []
        for r in range(r_start, hdr_end + 1):
            val = ws_eval.cell(row=r, column=c).value
            if val is None or str(val).strip() == "":
                # Buscar hacia la izquierda en la misma fila si es cabecera de fecha agrupada
                for c_left in range(c - 1, 0, -1):
                    v_left = ws_eval.cell(row=r, column=c_left).value
                    if v_left is not None and str(v_left).strip() != "":
                        s_l = _norm_label(v_left)
                        if any(_norm_label(tk) in s_l for tk in tokens) or re.search(r"202\d", s_l):
                            val = v_left
                        break
            if val is not None:
                col_texts.append(_norm_label(val))
        joined = " | ".join(t for t in col_texts if t)
        if any(_norm_label(tk) in joined for tk in tokens[:3]):
            period_cols.append(c)

    if not period_cols:
        # Probar por año (ej. 2026)
        yr = tokens[-1]
        for c in range(2, max_col + 1):
            joined = " | ".join(
                _norm_label(ws_eval.cell(row=r, column=c).value)
                for r in range(r_start, hdr_end + 1)
                if ws_eval.cell(row=r, column=c).value is not None
            )
            if yr in joined:
                period_cols.append(c)

    # Filtrar por sub-cabecera de columna (ej. 'Neto', 'Activo fijo neto', 'Corriente', 'No corriente', 'Activos', 'Pasivos')
    target_col: Optional[int] = None
    if col_subheader:
        sub_norm = _norm_label(col_subheader)
        search_pool = period_cols if period_cols else list(range(2, max_col + 1))
        if sub_norm == "total_segmento":
            for c in search_pool:
                hdr_j = " | ".join(
                    _norm_label(ws_eval.cell(row=r, column=c).value)
                    for r in range(r_start, hdr_end + 1)
                )
                if "total" in hdr_j and "elimin" not in hdr_j:
                    target_col = c
                    break
        else:
            for c in search_pool:
                hdr_j = " | ".join(
                    _norm_label(ws_eval.cell(row=r, column=c).value)
                    for r in range(r_start, hdr_end + 1)
                )
                if "no corrien" in sub_norm:
                    if "no corrien" in hdr_j or "no-corrien" in hdr_j:
                        target_col = c
                        break
                elif "corrien" in sub_norm:
                    if "corrien" in hdr_j and "no corrien" not in hdr_j and "no-corrien" not in hdr_j:
                        target_col = c
                        break
                elif sub_norm in hdr_j:
                    target_col = c
                    break

    if target_col is None:
        if period_cols:
            target_col = period_cols[0]
        else:
            # Primera columna numérica a la derecha de las etiquetas
            for c in range(2, max_col + 1):
                if any(isinstance(ws_eval.cell(row=r, column=c).value, (int, float)) for r in range(r_start, r_end + 1)):
                    target_col = c
                    break

    if target_col is None:
        return None

    # 2. Localizar fila de inicio de búsqueda si hay sub-sección interna (ej. 'Cuentas por pagar' en #N14.1)
    search_start = r_start
    if after_section_label:
        sec_norm = _norm_label(after_section_label)
        for r in range(r_start, r_end + 1):
            row_txt = " ".join(
                _norm_label(ws_eval.cell(row=r, column=c).value)
                for c in range(1, min(target_col, 4) + 1)
                if ws_eval.cell(row=r, column=c).value is not None
            )
            if sec_norm in row_txt:
                search_start = r + 1
                break

    # 3. Buscar la fila de Total correspondiente
    norm_targets = [_norm_label(lbl) for lbl in row_labels]
    tot_row: Optional[int] = None

    for t_lbl in norm_targets:
        for r in range(search_start, r_end + 1):
            lbl_cells = [
                _norm_label(ws_eval.cell(row=r, column=c).value)
                for c in range(1, target_col)
                if ws_eval.cell(row=r, column=c).value is not None
            ]
            lbl_cells = [x for x in lbl_cells if x]
            if any(x == t_lbl for x in lbl_cells):
                tot_row = r
                break
        if tot_row is not None:
            break

    if tot_row is None:
        for t_lbl in norm_targets:
            for r in range(search_start, r_end + 1):
                joined_lbl = " ".join(
                    _norm_label(ws_eval.cell(row=r, column=c).value)
                    for c in range(1, target_col)
                    if ws_eval.cell(row=r, column=c).value is not None
                ).strip()
                if joined_lbl and t_lbl in joined_lbl:
                    tot_row = r
                    break
            if tot_row is not None:
                break

    if tot_row is None and allow_blank_total_label:
        # Ej: #N22.2 cuya última fila tiene '=SUM(B35:B37)' y celda de texto vacía
        for r in range(r_end, search_start - 1, -1):
            v_num = ws_eval.cell(row=r, column=target_col).value
            if isinstance(v_num, (int, float)) and not isinstance(v_num, bool):
                lbl_txt = "".join(
                    _norm_label(ws_eval.cell(row=r, column=c).value)
                    for c in range(1, target_col)
                    if ws_eval.cell(row=r, column=c).value is not None
                )
                if lbl_txt == "":
                    tot_row = r
                    break

    if tot_row is None:
        return None

    return (tot_row, target_col, r_start, r_end)


_CELL_REF_RE = re.compile(r"\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?", re.IGNORECASE)


def _collect_upstream_leaf_cells(
    ws_raw: openpyxl.worksheet.worksheet.Worksheet,
    row: int,
    col: int,
    visited: Optional[Set[Tuple[int, int]]] = None,
) -> List[Tuple[int, int, float]]:
    """
    Expande recursivamente una fórmula =SUM(...) o =+B9+C9 en `ws_raw` hasta encontrar
    todas las celdas hoja con valores numéricos directos (no fórmulas).
    """
    if visited is None:
        visited = set()
    if (row, col) in visited:
        return []
    visited.add((row, col))

    val = ws_raw.cell(row=row, column=col).value
    if val is None or isinstance(val, bool):
        return []

    if isinstance(val, (int, float)):
        return [(row, col, float(val))]

    if isinstance(val, str) and val.strip().startswith("="):
        formula = val.strip()[1:]
        # Evitar cruzar referencias a otras hojas ('OtraHoja'!A1)
        if "!" in formula:
            return []
        leaves: List[Tuple[int, int, float]] = []
        for m in _CELL_REF_RE.finditer(formula):
            c1_str, r1_str, c2_str, r2_str = m.groups()
            c1 = openpyxl.utils.column_index_from_string(c1_str.upper())
            r1 = int(r1_str)
            if c2_str and r2_str:
                c2 = openpyxl.utils.column_index_from_string(c2_str.upper())
                r2 = int(r2_str)
                for rr in range(min(r1, r2), max(r1, r2) + 1):
                    for cc in range(min(c1, c2), max(c1, c2) + 1):
                        leaves.extend(_collect_upstream_leaf_cells(ws_raw, rr, cc, visited))
            else:
                leaves.extend(_collect_upstream_leaf_cells(ws_raw, r1, c1, visited))
        return leaves

    return []


def _get_row_label_in_ws(ws: openpyxl.worksheet.worksheet.Worksheet, row: int, max_label_col: int) -> str:
    parts = []
    for c in range(1, max(2, max_label_col)):
        v = ws.cell(row=row, column=c).value
        if v is not None and not isinstance(v, (int, float)):
            t = _norm_label(v)
            if t:
                parts.append(t)
    return " ".join(parts)


def _evaluate_ws_inplace(ws: openpyxl.worksheet.worksheet.Worksheet) -> None:
    """Evalúa todas las fórmulas de una hoja openpyxl in-place."""
    from src.ui_pages.informes_y_notas import evaluate_openpyxl_formula

    for r_idx in range(1, ws.max_row + 1):
        for c_idx in range(1, ws.max_column + 1):
            c_obj = ws.cell(row=r_idx, column=c_idx)
            c_val = c_obj.value
            if isinstance(c_val, str) and c_val.startswith("="):
                ev = evaluate_openpyxl_formula(c_val, ws, c_idx, r_idx, set())
                if ev is not None:
                    c_obj.value = ev


class RoundingReconciler:
    """
    Orquestador de cuadratura por redondeo (<= $4.000 / <= 4 M$) entre EEFF y Notas.
    """

    _eeff_targets_cache: Dict[Tuple[str, str], Tuple[Optional[pd.DataFrame], Optional[pd.DataFrame]]] = {}

    @classmethod
    def clear_cache(cls) -> None:
        cls._eeff_targets_cache.clear()

    @classmethod
    def get_authoritative_eeff(
        cls,
        empresa: str,
        periodo_actual: str,
    ) -> Tuple[Optional[pd.DataFrame], Optional[pd.DataFrame]]:
        """
        Obtiene los DataFrames autoritativos en M$ (#BALANCE y #ER) para `empresa` y `periodo_actual`.
        Usa caché en memoria para evitar recalcular cuando se consultan varias notas seguidas.
        """
        emp_clean = str(empresa).strip()
        per_clean = str(periodo_actual).strip()
        cache_key = (emp_clean, per_clean)
        if cache_key in cls._eeff_targets_cache:
            return cls._eeff_targets_cache[cache_key]

        from src.core.snapshot_vault_manager import SnapshotVaultManager

        if SnapshotVaultManager.has_active_snapshot(emp_clean, per_clean):
            b_snap = SnapshotVaultManager.get_statement(emp_clean, per_clean, "#BALANCE", scale_factor=1000.0)
            e_snap = SnapshotVaultManager.get_statement(emp_clean, per_clean, "#ER", scale_factor=1000.0)
            df_bal = b_snap[0] if b_snap else None
            df_er = e_snap[0] if e_snap else None
            cls._eeff_targets_cache[cache_key] = (df_bal, df_er)
            return df_bal, df_er

        is_consolidated = emp_clean.startswith("[GRUPO]")
        parts = per_clean.split("-")
        prev_year = str(int(parts[0]) - 1) if len(parts) >= 2 and parts[0].isdigit() else "2025"
        comp_bal = f"{prev_year}-12"
        comp_pl = f"{prev_year}-{parts[1]}" if len(parts) >= 2 else comp_bal

        from src.core.word_template_engine import WordTemplateEngine, load_mappings_for_entity
        from src.models.pl_cubo_db import PlCuboDB
        from src.models.trial_balance_db import TrialBalanceDB
        from src.reporting.balance_generator import BalanceGenerator
        from src.reporting.er_generator import ERGenerator

        map_balance_df, map_pl_df = load_mappings_for_entity(emp_clean)
        wte = WordTemplateEngine(None)

        empresa_path = os.path.join("data", "empresas", emp_clean)
        if not os.path.exists(empresa_path):
            empresa_path = os.path.join("data", "empresas", emp_clean.replace("[GRUPO] ", "").strip())

        df_bal = None
        df_er = None

        try:
            if is_consolidated:
                from src.core.consolidacion_engine import generar_hoja_trabajo
                from src.models.consolidacion import ConsolidationGroup
                from src.models.database import SessionLocal

                g_name = emp_clean.replace("[GRUPO] ", "").strip()
                db = SessionLocal()
                g_obj = db.query(ConsolidationGroup).filter_by(nombre_grupo=g_name).first()
                g_id = g_obj.id if g_obj else None
                db.close()

                if g_id:
                    df_hoja_act, _ = generar_hoja_trabajo(g_id, per_clean)
                    df_hoja_comp, _ = generar_hoja_trabajo(g_id, comp_bal)
                    if df_hoja_act is not None and not df_hoja_act.empty:
                        bs_subtotals = {
                            "activos corrientes", "activos corrientes totales",
                            "activos no corrientes", "activos no corrientes totales",
                            "total activos", "patrimonio y pasivos",
                            "pasivos corrientes", "pasivo corrientes totales",
                            "pasivos no corrientes", "pasivo no corrientes totales",
                            "total pasivos", "patrimonio", "patrimonio total",
                            "total patrimonio y pasivos", "estado de resultados",
                        }
                        idx_er = df_hoja_act[df_hoja_act["Balance clasificado"] == "Estado de Resultados"].index
                        sec_bal = df_hoja_act.loc[: idx_er[0] - 1] if not idx_er.empty else df_hoja_act
                        sec_bal = sec_bal[sec_bal["Balance clasificado"].notna() & (sec_bal["Balance clasificado"].astype(str).str.strip() != "")]
                        sec_bal = sec_bal[~sec_bal["Balance clasificado"].apply(_norm_label).isin(bs_subtotals)]
                        sec_bal = sec_bal[~sec_bal["Balance clasificado"].apply(is_pl_classification_label)]
                        tb_df = pd.DataFrame({"cuenta_id": sec_bal["Balance clasificado"], "saldo_final": sec_bal["CONSOLIDADO"]})

                        tb_comp = None
                        if df_hoja_comp is not None and not df_hoja_comp.empty:
                            idx_er_c = df_hoja_comp[df_hoja_comp["Balance clasificado"] == "Estado de Resultados"].index
                            sec_c = df_hoja_comp.loc[: idx_er_c[0] - 1] if not idx_er_c.empty else df_hoja_comp
                            sec_c = sec_c[sec_c["Balance clasificado"].notna() & (sec_c["Balance clasificado"].astype(str).str.strip() != "")]
                            sec_c = sec_c[~sec_c["Balance clasificado"].apply(_norm_label).isin(bs_subtotals)]
                            sec_c = sec_c[~sec_c["Balance clasificado"].apply(is_pl_classification_label)]
                            tb_comp = pd.DataFrame({"cuenta_id": sec_c["Balance clasificado"], "saldo_final": sec_c["CONSOLIDADO"]})

                        t_bal = os.path.join(empresa_path, "Balance clasificado.xlsx")
                        if not os.path.exists(t_bal):
                            t_bal = os.path.join("templates", "Balance clasificado.xlsx")
                        dummy_map = pd.DataFrame({"N° de Cuenta": tb_df["cuenta_id"], "Clasificación balance": tb_df["cuenta_id"]})
                        res_b = BalanceGenerator(t_bal).generate(
                            tb_df, dummy_map, scale_factor=1000.0, tb_df_comp=tb_comp,
                            periodo_actual_str=per_clean, periodo_comp_str=comp_bal,
                        )
                        df_bal = wte._extract_balance_preview(res_b, per_clean, comp_bal)

                        # ER Consolidado
                        pl_subs = {
                            "ganancia bruta", "resultado antes de impuestos",
                            "ganancias (perdida) del ejercicio", "ganancia (perdida) del ejercicio",
                            "estado de resultados", "otros rubros no clasificados",
                        }
                        sec_er = df_hoja_act.loc[idx_er[0] + 1 :] if not idx_er.empty else df_hoja_act
                        sec_er = sec_er[sec_er["Balance clasificado"].notna() & (sec_er["Balance clasificado"].astype(str).str.strip() != "")]
                        sec_er = sec_er[~sec_er["Balance clasificado"].apply(_norm_label).isin(pl_subs)]
                        pl_wide = pd.DataFrame({r["Balance clasificado"]: [r["CONSOLIDADO"]] for _, r in sec_er.iterrows()})
                        t_er = os.path.join(empresa_path, "Estado de Resultados Clasificados.xlsx")
                        if not os.path.exists(t_er):
                            t_er = os.path.join("templates", "Estado de Resultados Clasificados.xlsx")
                        _, df_er = ERGenerator(t_er).generate(
                            pl_wide, scale_factor=1000.0, pl_df_comp=None,
                            periodo_actual_str=per_clean, periodo_comp_str=comp_pl, map_pl_df=map_pl_df,
                        )
            else:
                tb_df = TrialBalanceDB.get_trial_balance(emp_clean, per_clean)
                tb_comp = TrialBalanceDB.get_trial_balance(emp_clean, comp_bal)
                if tb_df is not None and not tb_df.empty:
                    t_bal = os.path.join(empresa_path, "Balance clasificado.xlsx")
                    if not os.path.exists(t_bal):
                        t_bal = os.path.join("templates", "Balance clasificado.xlsx")
                    res_b = BalanceGenerator(t_bal).generate(
                        tb_df, map_balance_df, scale_factor=1000.0, tb_df_comp=tb_comp,
                        periodo_actual_str=per_clean, periodo_comp_str=comp_bal, empresa=emp_clean,
                    )
                    df_bal = wte._extract_balance_preview(res_b, per_clean, comp_bal)
                    df_bal = SnapshotVaultManager.apply_closed_comp_to_statement(
                        df_bal, emp_clean, "#BALANCE", per_clean, comp_bal, scale_factor=1000.0
                    )

                pl_df = PlCuboDB.get_pl_cubo(emp_clean, per_clean)
                if (pl_df is None or pl_df.empty) and tb_df is not None and not tb_df.empty and map_pl_df is not None:
                    from src.core.sabana_builder import build_pl_sabana
                    pl_df = build_pl_sabana(None, map_pl_df, tb_df)
                if pl_df is not None and not pl_df.empty:
                    t_er = os.path.join(empresa_path, "Estado de Resultados Clasificados.xlsx")
                    if not os.path.exists(t_er):
                        t_er = os.path.join("templates", "Estado de Resultados Clasificados.xlsx")
                    res_e, _ = ERGenerator(t_er).generate(
                        pl_df, scale_factor=1000.0, pl_df_comp=None,
                        periodo_actual_str=per_clean, periodo_comp_str=comp_pl, map_pl_df=map_pl_df,
                    )
                    df_er = wte._extract_er_preview(res_e, per_clean, comp_pl)
                    df_er = SnapshotVaultManager.apply_closed_comp_to_statement(
                        df_er, emp_clean, "#ER", per_clean, comp_pl, scale_factor=1000.0
                    )
        except Exception as exc:
            print(f"Advertencia obteniendo EEFF autoritativos ({emp_clean} {per_clean}): {exc}")

        cls._eeff_targets_cache[cache_key] = (df_bal, df_er)
        return df_bal, df_er

    @staticmethod
    def _apply_single_table_reconciliation(
        ws_raw: openpyxl.worksheet.worksheet.Worksheet,
        ws_eval: openpyxl.worksheet.worksheet.Worksheet,
        sub_code: str,
        periodo: str,
        target_eeff_m: int,
        sign_factor: int,
        row_labels: List[str],
        col_subheader: Optional[str] = None,
        after_section_label: Optional[str] = None,
        allow_blank_total_label: bool = False,
        protected_row_labels: Optional[Set[str]] = None,
        prefer_leaf_col_keyword: Optional[str] = None,
        max_tol_m: int = MAX_TOLERANCE_M,
    ) -> bool:
        """
        Compara el Total de un cuadro de nota en `ws_eval` contra `target_eeff_m * sign_factor`.
        Si `1 <= |delta| <= max_tol_m`, ajusta la celda hoja de entrada elegible de MAYOR |valor|
        en `ws_raw` y re-evalúa `ws_eval` in-place.
        """
        loc = _locate_total_cell_in_ws(
            ws_eval=ws_eval,
            sub_code=sub_code,
            periodo=periodo,
            row_labels=row_labels,
            col_subheader=col_subheader,
            after_section_label=after_section_label,
            allow_blank_total_label=allow_blank_total_label,
        )
        if loc is None:
            return False

        tot_row, tot_col, block_start, block_end = loc
        tot_val_raw = ws_eval.cell(row=tot_row, column=tot_col).value
        current_note_m, ok = _parse_int_m(tot_val_raw)
        if not ok:
            return False

        # Nunca inventar saldos en una nota completamente vacía (current_note_m == 0)
        if current_note_m == 0:
            return False

        expected_note_m = int(target_eeff_m) * int(sign_factor)
        delta_note = int(expected_note_m - current_note_m)

        if delta_note == 0 or abs(delta_note) > max_tol_m:
            return False

        protected_norms = {_norm_label(p) for p in (protected_row_labels or set())}
        raw_tot_cell_val = ws_raw.cell(row=tot_row, column=tot_col).value
        is_tot_formula = isinstance(raw_tot_cell_val, str) and raw_tot_cell_val.strip().startswith("=")

        if is_tot_formula:
            leaves = _collect_upstream_leaf_cells(ws_raw, tot_row, tot_col)
        else:
            # Si un anexo externo inyectó valores fijos y el total sin fórmula, recolectar celdas numéricas de la columna
            leaves = []
            for rr in range(block_start + 1, tot_row):
                v_c = ws_raw.cell(row=rr, column=tot_col).value
                if isinstance(v_c, (int, float)) and not isinstance(v_c, bool):
                    lbl_r = _get_row_label_in_ws(ws_raw, rr, tot_col)
                    if "sub total" not in lbl_r and "subtotal" not in lbl_r and not lbl_r.startswith("total"):
                        leaves.append((rr, tot_col, float(v_c)))

        if not leaves:
            return False

        # Deduplicar celdas hoja preservando orden
        seen_rc: Set[Tuple[int, int]] = set()
        unique_leaves: List[Tuple[int, int, float]] = []
        for r_l, c_l, v_l in leaves:
            if (r_l, c_l) not in seen_rc:
                seen_rc.add((r_l, c_l))
                unique_leaves.append((r_l, c_l, v_l))

        # Si es tabla 2D (Bruto + Depreciación/Amortización = Neto), preferir la columna de Depreciación/Amortización
        if prefer_leaf_col_keyword:
            kw_norm = _norm_label(prefer_leaf_col_keyword)
            hdr_end = min(block_start + 9, tot_row)
            matching_cols: Set[int] = set()
            for _, c_l, _ in unique_leaves:
                hdr_txt = " ".join(
                    _norm_label(ws_eval.cell(row=rr, column=c_l).value)
                    for rr in range(block_start, hdr_end + 1)
                    if ws_eval.cell(row=rr, column=c_l).value is not None
                )
                if kw_norm in hdr_txt:
                    matching_cols.add(c_l)
            if matching_cols:
                filtered_2d = [item for item in unique_leaves if item[1] in matching_cols]
                if filtered_2d:
                    unique_leaves = filtered_2d

        def _filter_candidates(exclude_protected: bool) -> List[Tuple[int, int, int]]:
            cands: List[Tuple[int, int, int]] = []
            for r_l, c_l, v_l in unique_leaves:
                v_int = int(round(v_l))
                if v_int == 0:
                    continue
                # Proteger signo contable: el ajuste nunca puede invertir el signo ni dejar la línea en cero
                new_v = v_int + delta_note
                if (v_int > 0 and new_v <= 0) or (v_int < 0 and new_v >= 0):
                    continue
                if exclude_protected and protected_norms:
                    lbl_r = _get_row_label_in_ws(ws_raw, r_l, c_l)
                    if any(p in lbl_r for p in protected_norms):
                        continue
                cands.append((r_l, c_l, v_int))
            return cands

        candidates = _filter_candidates(exclude_protected=True)
        if not candidates and protected_norms:
            candidates = _filter_candidates(exclude_protected=False)

        if not candidates:
            return False

        # Regla: seleccionar la línea elegible de MAYOR valor absoluto (|valor|)
        r_best, c_best, v_best = max(candidates, key=lambda item: (abs(item[2]), -item[0]))
        ws_raw.cell(row=r_best, column=c_best).value = int(v_best + delta_note)

        if not is_tot_formula:
            ws_raw.cell(row=tot_row, column=tot_col).value = int(current_note_m + delta_note)
            ws_eval.cell(row=r_best, column=c_best).value = int(v_best + delta_note)
            ws_eval.cell(row=tot_row, column=tot_col).value = int(current_note_m + delta_note)
        else:
            ws_eval.cell(row=r_best, column=c_best).value = int(v_best + delta_note)
            # Re-restaurar fórmulas de la hoja desde ws_raw y re-evaluar hoja para propagar subtotales y cruces (#N08.2, #N09.2, #N17.22)
            for rr in range(1, ws_raw.max_row + 1):
                for cc in range(1, ws_raw.max_column + 1):
                    rv = ws_raw.cell(row=rr, column=cc).value
                    if isinstance(rv, str) and rv.startswith("="):
                        ws_eval.cell(row=rr, column=cc).value = rv
            _evaluate_ws_inplace(ws_eval)

        return True

    @classmethod
    def reconcile_note_workbooks_inplace(
        cls,
        wb_raw: openpyxl.Workbook,
        wb_eval: openpyxl.Workbook,
        df_bal: Optional[pd.DataFrame],
        df_er: Optional[pd.DataFrame],
        periodo_actual: str,
        is_consolidated: bool = False,
        max_tol_m: int = MAX_TOLERANCE_M,
    ) -> int:
        """
        Ejecuta todas las reglas de reconciliación (Nivel 2: EEFF -> Notas y Nivel 3: Sub-cuadros internos)
        sobre las hojas presentes en `wb_raw` / `wb_eval`. Retorna la cantidad de cuadros ajustados.
        """
        if df_bal is None and df_er is None:
            return 0

        sheet_set = set(wb_raw.sheetnames)
        adjusted_count = 0

        def _run_rule(
            sheet_name: str,
            sub_code: str,
            stmt_df: Optional[pd.DataFrame],
            stmt_aliases: List[str],
            sign_factor: int,
            row_labels: List[str],
            col_subheader: Optional[str] = None,
            after_section_label: Optional[str] = None,
            allow_blank_total_label: bool = False,
            protected_row_labels: Optional[Set[str]] = None,
            prefer_leaf_col_keyword: Optional[str] = None,
        ) -> None:
            nonlocal adjusted_count
            if sheet_name not in sheet_set or stmt_df is None or stmt_df.empty:
                return
            target_m, exists = _extract_stmt_target(stmt_df, periodo_actual, stmt_aliases)
            if not exists:
                return
            changed = cls._apply_single_table_reconciliation(
                ws_raw=wb_raw[sheet_name],
                ws_eval=wb_eval[sheet_name],
                sub_code=sub_code,
                periodo=periodo_actual,
                target_eeff_m=target_m,
                sign_factor=sign_factor,
                row_labels=row_labels,
                col_subheader=col_subheader,
                after_section_label=after_section_label,
                allow_blank_total_label=allow_blank_total_label,
                protected_row_labels=protected_row_labels,
                prefer_leaf_col_keyword=prefer_leaf_col_keyword,
                max_tol_m=max_tol_m,
            )
            if changed:
                adjusted_count += 1

        # ---------------------------------------------------------------------
        # NIVEL 2A: BALANCE CLASIFICADO (#BALANCE) -> NOTAS PRINCIPALES
        # ---------------------------------------------------------------------
        _run_rule(
            "Efectivo", "#N04.1", df_bal, ["Efectivo y efectivo equivalente"], 1,
            ["Total", "Totales"], protected_row_labels={"fondos mutuos"},
        )
        _run_rule(
            "Otros activos no financieros, c", "#N05.1", df_bal,
            ["Otros activos no financieros, corrientes"], 1, ["Total", "Totales"],
        )
        _run_rule(
            "Deudores", "#N06.1", df_bal,
            ["Deudores comerciales y otras cuentas por cobrar, corrientes"], 1, ["Total", "Totales"],
        )
        _run_rule(
            "Inventarios", "#N07.1", df_bal, ["Inventarios"], 1, ["Total", "Totales"],
        )
        _run_rule(
            "Intangibles", "#N08.1", df_bal,
            ["Activos intangibles distinto a la plusvalía", "Activos intangibles distintos de la plusvalía"], 1,
            ["Total", "Totales"], col_subheader="Neto", prefer_leaf_col_keyword="amortizacion",
        )
        _run_rule(
            "Activo Fijo", "#N09.1", df_bal,
            ["Propiedades, plantas y equipos", "Propiedades, planta y equipo"], 1,
            ["Totales", "Total"], col_subheader="Activo fijo neto", prefer_leaf_col_keyword="depreciacion",
        )
        _run_rule(
            "Activo por derechos de uso", "#N10.1", df_bal,
            ["Activo por derechos de uso", "Activos por derechos de uso"], 1,
            ["Totales", "Total"], col_subheader="Activo fijo neto", prefer_leaf_col_keyword="depreciacion",
        )
        _run_rule(
            "Plusvalia", "#N11.1", df_bal, ["Plusvalia", "Plusvalía"], 1, ["Total", "Totales"],
        )
        _run_rule(
            "Impuestos corrientes", "#N12.1", df_bal,
            ["Activo por impuestos, corrientes", "Activos por impuestos, corrientes"], 1,
            ["Total activos", "Total", "Totales"],
        )
        _run_rule(
            "Impuestos corrientes", "#N12.2", df_bal,
            ["Pasivo por impuestos, corrientes", "Pasivos por impuestos, corrientes"], 1,
            ["Total pasivos", "Total", "Totales"],
        )
        _run_rule(
            "Impuestos Diferidos", "#N13.1", df_bal,
            ["Activo por impuestos diferidos, no corrientes", "Activos por impuestos diferidos"], 1,
            ["Totales", "Total"], col_subheader="Activos",
        )
        _run_rule(
            "Impuestos Diferidos", "#N13.1", df_bal,
            ["Pasivos por impuestos diferidos, no corrientes", "Pasivo por impuestos diferidos, no corrientes"], 1,
            ["Totales", "Total"], col_subheader="Pasivos",
        )
        _run_rule(
            "Empresas relacionadas", "#N14.1", df_bal,
            ["Cuentas por cobrar a entidades relacionadas, corrientes"], 1,
            ["Total", "Totales"], col_subheader="Corriente",
        )
        _run_rule(
            "Empresas relacionadas", "#N14.1", df_bal,
            ["Cuentas por cobrar a entidades relacionadas, no corrientes"], 1,
            ["Total", "Totales"], col_subheader="No corriente",
        )
        _run_rule(
            "Empresas relacionadas", "#N14.1", df_bal,
            ["Cuentas por pagar entidades relacionadas, corrientes", "Cuentas por pagar a entidades relacionadas, corrientes"], 1,
            ["Total", "Totales"], col_subheader="Corriente", after_section_label="Cuentas por pagar",
        )
        _run_rule(
            "Empresas relacionadas", "#N14.1", df_bal,
            ["Cuentas por pagar entidades relacionadas, no corrientes", "Cuentas por pagar a entidades relacionadas, no corrientes"], 1,
            ["Total", "Totales"], col_subheader="No corriente", after_section_label="Cuentas por pagar",
        )
        _run_rule(
            "Inversion en relacionadas", "#N14.3", df_bal,
            ["Inversion en empresas relacionadas", "Inversión en empresas relacionadas"], 1,
            ["Total", "Totales"],
        )
        _run_rule(
            "Pasivos financieros ", "#N15.1", df_bal,
            ["Activos financieros, Corrientes", "Otros activos financieros, corrientes"], 1,
            ["Total otros activos financieros", "Total", "Totales"], col_subheader="Corrientes",
        )
        _run_rule(
            "Pasivos financieros ", "#N15.1", df_bal,
            ["Otros activos financieros, no corrientes"], 1,
            ["Total otros activos financieros", "Total", "Totales"], col_subheader="No corrientes",
        )
        _run_rule(
            "Pasivos financieros ", "#N15.2", df_bal,
            ["Otros pasivos financieros corrientes", "Otros pasivos financieros, corrientes"], 1,
            ["Total otros pasivos financieros", "Total", "Totales"], col_subheader="Corrientes",
        )
        _run_rule(
            "Pasivos financieros ", "#N15.2", df_bal,
            ["Otros pasivos financieros, no corriente", "Otros pasivos financieros, no corrientes"], 1,
            ["Total otros pasivos financieros", "Total", "Totales"], col_subheader="No corrientes",
        )
        _run_rule(
            "Pasivos derechos de  uso", "#N16.1", df_bal,
            ["Pasivos por derechos de uso, corrientes"], 1,
            ["Total", "Totales"], col_subheader="Corriente",
        )
        _run_rule(
            "Pasivos derechos de  uso", "#N16.1", df_bal,
            ["Pasivos por derechos de uso, no corrientes"], 1,
            ["Total", "Totales"], col_subheader="No corriente",
        )
        _run_rule(
            "Cuentas por pagar", "#N17.1", df_bal,
            ["Cuentas comerciales y otras cuentas por pagar, corrientes"], 1, ["Total", "Totales"],
        )
        _run_rule(
            "Cuentas por pagar", "#N17.2", df_bal,
            ["Cuentas comerciales y otras cuentas por pagar, no corrientes"], 1, ["Total", "Totales"],
        )
        _run_rule(
            "Provisiones", "#N18.1", df_bal,
            ["Provisiones por beneficios a los empleados"], 1, ["Total", "Totales"],
        )
        _run_rule(
            "Provisiones", "#N18.2", df_bal,
            ["Otras provisiones, no corrientes"], 1, ["Total", "Totales"],
        )
        _run_rule(
            "Patrimonio", "#N20.4", df_bal, ["Otras reservas"], 1,
            ["Otras reservas", "Saldo final", "Total"],
        )

        # ---------------------------------------------------------------------
        # NIVEL 2B: ESTADO DE RESULTADOS (#ER) -> NOTAS DE RESULTADOS
        # ---------------------------------------------------------------------
        _run_rule(
            "Ingresos Ctos operacion", "#N21.1", df_er,
            ["Ingresos de actividades ordinarias"], 1, ["Total ingresos", "Total", "Totales"],
        )
        _run_rule(
            "Ingresos Ctos operacion", "#N21.2", df_er,
            ["Costo de ventas"], -1, ["Total costos", "Total costo de ventas"],
        )
        _run_rule(
            "Ingresos Ctos operacion", "#N21.2", df_er,
            ["Depreciacion operacional", "Depreciación operacional", "Depreciación y amortización operacional"], -1,
            ["Total Depreciacion", "Total depreciación"],
        )
        _run_rule(
            "Gtos Adm", "#N22.1", df_er,
            ["Gastos de administración", "Gastos de administracion"], -1,
            ["Total gastos de administración", "Total", "Totales"],
        )
        _run_rule(
            "Gtos Adm", "#N22.2", df_er,
            ["Depreciación y amortizaciones", "Depreciacion y amortizaciones"], -1,
            ["Total", "Totales"], allow_blank_total_label=True,
        )
        _run_rule(
            "DC y Reajustes", "#N23.1", df_er,
            ["Diferencias de cambio", "Diferencia de cambio"], -1,
            ["Total diferencia de cambio", "Total", "Totales"],
        )
        _run_rule(
            "DC y Reajustes", "#N23.2", df_er,
            ["Resultado por unidad de reajuste", "Resultados por unidades de reajuste"], -1,
            ["Total resultado por unidad de reajuste", "Total", "Totales"],
        )
        _run_rule(
            "Costos e ingresos Financieros", "#N24.1", df_er,
            ["Costos financieros"], -1,
            ["Total, costos financieros", "Total costos financieros", "Total", "Totales"],
        )
        _run_rule(
            "Costos e ingresos Financieros", "#N24.2", df_er,
            ["Ingresos financieros"], 1,
            ["Total", "Totales", "Total ingresos financieros"],
        )
        for sh_n25_1 in ("Otros ingresos y egresos", "Otros gastos por funcion"):
            _run_rule(
                sh_n25_1, "#N25.1", df_er,
                ["Otros egresos por función", "Otros Egresos por función", "Otros egresos por funcion"], -1,
                ["Total", "Totales"],
            )
        for sh_n25_2 in ("Otros ingresos y egresos", "Otros ingresos por funcion"):
            _run_rule(
                sh_n25_2, "#N25.2", df_er,
                ["Otros ingresos por función", "Otros ingresos por funcion"], 1,
                ["Total", "Totales"],
            )
        _run_rule(
            "Impuestos Diferidos", "#N13.4", df_er,
            ["Ganancia (Perdida) por impuesto a las ganancias", "Resultado por impuestos a las ganancias"], 1,
            ["(Gasto) ingreso por impuesto a las ganancias", "Total"],
        )

        # ---------------------------------------------------------------------
        # NIVEL 2C: SOLO CONSOLIDADO — #N26 SEGMENTOS DE OPERACIÓN
        # ---------------------------------------------------------------------
        if is_consolidated and "Segmentos" in sheet_set:
            _run_rule(
                "Segmentos", "#N26.1", df_bal, ["Activos corrientes totales"], 1,
                ["Total activos corrientes por segmento"], col_subheader="total_segmento",
            )
            _run_rule(
                "Segmentos", "#N26.1", df_bal, ["Total activos no corrientes"], 1,
                ["Total activos no corrientes por segmento"], col_subheader="total_segmento",
            )
            _run_rule(
                "Segmentos", "#N26.1", df_bal, ["Pasivos corrientes totales"], 1,
                ["Total pasivos corrientes por segmento"], col_subheader="total_segmento",
            )
            _run_rule(
                "Segmentos", "#N26.1", df_bal, ["Total pasivos no corrientes"], 1,
                ["Total pasivos no corrientes por segmento"], col_subheader="total_segmento",
            )
            _run_rule(
                "Segmentos", "#N26.2", df_er, ["Ingresos de actividades ordinarias"], 1,
                ["Ingresos de actividades ordinarias"], col_subheader="total_segmento",
            )
            _run_rule(
                "Segmentos", "#N26.2", df_er, ["Ganancia bruta", "Ganancia Bruta"], 1,
                ["Margen bruto"], col_subheader="total_segmento",
            )
            _run_rule(
                "Segmentos", "#N26.2", df_er, ["Resultado antes de impuestos", "Ganancia antes de impuesto"], 1,
                ["Ganancia (Pérdida) antes de impuestos"], col_subheader="total_segmento",
            )
            _run_rule(
                "Segmentos", "#N26.2", df_er, ["(Pérdida) procedente de operaciones continuadas", "Ganancias (Pérdida) del Ejercicio"], 1,
                ["Ganancia (Pérdida) neta del ejercicio"], col_subheader="total_segmento",
            )

        # ---------------------------------------------------------------------
        # NIVEL 3: SUB-CUADROS INTERNOS EN CASCADA (#N04.2, #N04.3, #N17.3)
        # ---------------------------------------------------------------------
        if "Efectivo" in sheet_set:
            ws_ef_eval = wb_eval["Efectivo"]
            ws_ef_raw = wb_raw["Efectivo"]
            loc_n04_1 = _locate_total_cell_in_ws(ws_ef_eval, "#N04.1", periodo_actual, ["Total", "Totales"])
            if loc_n04_1 is not None:
                v_n04_1, ok_n04_1 = _parse_int_m(ws_ef_eval.cell(row=loc_n04_1[0], column=loc_n04_1[1]).value)
                if ok_n04_1:
                    if cls._apply_single_table_reconciliation(
                        ws_raw=ws_ef_raw, ws_eval=ws_ef_eval, sub_code="#N04.2",
                        periodo=periodo_actual, target_eeff_m=v_n04_1, sign_factor=1,
                        row_labels=["Total", "Totales"], max_tol_m=max_tol_m,
                    ):
                        adjusted_count += 1

            loc_fm = _locate_total_cell_in_ws(ws_ef_eval, "#N04.1", periodo_actual, ["Fondos mutuos"])
            if loc_fm is not None:
                v_fm, ok_fm = _parse_int_m(ws_ef_eval.cell(row=loc_fm[0], column=loc_fm[1]).value)
                if ok_fm and v_fm != 0:
                    if cls._apply_single_table_reconciliation(
                        ws_raw=ws_ef_raw, ws_eval=ws_ef_eval, sub_code="#N04.3",
                        periodo=periodo_actual, target_eeff_m=v_fm, sign_factor=1,
                        row_labels=["Total", "Totales"], max_tol_m=max_tol_m,
                    ):
                        adjusted_count += 1

        if "Cuentas por pagar" in sheet_set:
            ws_cxp_eval = wb_eval["Cuentas por pagar"]
            ws_cxp_raw = wb_raw["Cuentas por pagar"]
            loc_17_1 = _locate_total_cell_in_ws(ws_cxp_eval, "#N17.1", periodo_actual, ["Total", "Totales"])
            loc_17_2 = _locate_total_cell_in_ws(ws_cxp_eval, "#N17.2", periodo_actual, ["Total", "Totales"])
            v17_1 = _parse_int_m(ws_cxp_eval.cell(row=loc_17_1[0], column=loc_17_1[1]).value)[0] if loc_17_1 else 0
            v17_2 = _parse_int_m(ws_cxp_eval.cell(row=loc_17_2[0], column=loc_17_2[1]).value)[0] if loc_17_2 else 0
            target_cxp_tot = v17_1 + v17_2
            if target_cxp_tot != 0:
                if cls._apply_single_table_reconciliation(
                    ws_raw=ws_cxp_raw, ws_eval=ws_cxp_eval, sub_code="#N17.3",
                    periodo=periodo_actual, target_eeff_m=target_cxp_tot, sign_factor=1,
                    row_labels=["Totales", "Total"], max_tol_m=max_tol_m,
                ):
                    adjusted_count += 1

        return adjusted_count

    @classmethod
    def reconcile_and_evaluate_note_workbook(
        cls,
        wb_raw_bytes: BytesIO,
        empresa: str,
        periodo_actual: str,
        scale_factor: float = 1000.0,
        df_bal: Optional[pd.DataFrame] = None,
        df_er: Optional[pd.DataFrame] = None,
        target_sheets: Optional[List[str]] = None,
        return_raw: bool = True,
    ) -> Tuple[Optional[BytesIO], BytesIO]:
        """
        Punto de entrada unificado para los pipelines de Notas:
        1. Evalúa las fórmulas del libro de notas (`wb_eval`).
        2. Si `scale_factor == 1000.0` (M$), aplica la reconciliación de redondeo (<= 4 M$)
           contra los EEFF autoritativos sobre `wb_raw` (preservando fórmulas =SUM(...) para Excel)
           y sobre `wb_eval` (valores evaluados para Vista Web, Word, Snapshot y Validaciones).
        3. Retorna `(wb_raw_reconciled_io, wb_eval_reconciled_io)`.
        """
        from src.ui_pages.informes_y_notas import evaluate_formulas_in_workbook

        if abs(float(scale_factor) - 1000.0) >= 0.01:
            wb_raw_bytes.seek(0)
            raw_copy = BytesIO(wb_raw_bytes.getvalue())
            eval_out = evaluate_formulas_in_workbook(BytesIO(wb_raw_bytes.getvalue()))
            return raw_copy, eval_out

        wb_raw_bytes.seek(0)
        raw_buf = wb_raw_bytes.getvalue()
        wb_raw = openpyxl.load_workbook(BytesIO(raw_buf), data_only=False)
        wb_eval = openpyxl.load_workbook(BytesIO(raw_buf), data_only=False)

        active_set = set(target_sheets) if target_sheets else set(wb_eval.sheetnames)
        for sh_name in wb_eval.sheetnames:
            if sh_name in active_set:
                _evaluate_ws_inplace(wb_eval[sh_name])

        if df_bal is None and df_er is None:
            df_bal, df_er = cls.get_authoritative_eeff(empresa, periodo_actual)

        is_cons = str(empresa).strip().startswith("[GRUPO]")
        cls.reconcile_note_workbooks_inplace(
            wb_raw=wb_raw,
            wb_eval=wb_eval,
            df_bal=df_bal,
            df_er=df_er,
            periodo_actual=periodo_actual,
            is_consolidated=is_cons,
            max_tol_m=MAX_TOLERANCE_M,
        )

        out_raw = None
        if return_raw:
            out_raw = BytesIO()
            wb_raw.save(out_raw)
            out_raw.seek(0)
        wb_raw.close()

        out_eval = BytesIO()
        wb_eval.save(out_eval)
        wb_eval.close()
        out_eval.seek(0)
        return out_raw, out_eval
