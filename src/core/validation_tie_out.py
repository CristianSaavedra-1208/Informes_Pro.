import io
import os
import re
import time
from io import BytesIO
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from src.core.snapshot_vault_manager import SnapshotVaultManager
from src.models.consolidacion import ConsolidationGroup
from src.models.database import SessionLocal
from src.models.pl_cubo_db import PlCuboDB
from src.models.trial_balance_db import TrialBalanceDB
from src.reporting.notes import get_full_note_registry


def _norm_txt(val: Any) -> str:
    """Normaliza etiquetas para comparación determinística sin tildes ni puntuación final."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val).strip().lower()
    if s in ("nan", "nat", "none"):
        return ""
    replacements = {
        "á": "a", "é": "e", "í": "i", "ó": "o", "ú": "u", "ü": "u",
    }
    for k, v in replacements.items():
        s = s.replace(k, v)
    s = re.sub(r"\s+", " ", s).strip()
    s = s.rstrip(":.")
    return s


def _parse_int_m(val: Any) -> Tuple[int, bool]:
    """
    Convierte una celda expresada en M$ a entero redondeado (int(round(x))).
    Retorna (valor_entero, tiene_dato_explicito).
    """
    if val is None:
        return 0, False
    if isinstance(val, bool):
        return 0, False
    if isinstance(val, (int, float)):
        if pd.isna(val):
            return 0, False
        return int(round(float(val))), True

    s = str(val).strip()
    if not s or s.lower() in ("nan", "nat", "none", "-", "—", ""):
        return 0, False

    # Soporte de paréntesis contables: (1.234) -> -1234
    is_neg = False
    if s.startswith("(") and s.endswith(")"):
        is_neg = True
        s = s[1:-1].strip()

    s_clean = s.replace("$", "").replace("M", "").replace(" ", "")
    # Distinguir formato miles con comas ("391,778,762") vs decimal con punto ("66198760.0")
    if "," in s_clean and "." in s_clean:
        if s_clean.rfind(",") > s_clean.rfind("."):
            s_clean = s_clean.replace(".", "").replace(",", ".")
        else:
            s_clean = s_clean.replace(",", "")
    elif s_clean.count(",") >= 2:
        s_clean = s_clean.replace(",", "")
    elif s_clean.count(".") >= 2:
        s_clean = s_clean.replace(".", "")
    elif "," in s_clean and "." not in s_clean:
        parts = s_clean.split(",")
        if len(parts) == 2 and len(parts[1]) == 3 and len(parts[0]) <= 3:
            s_clean = s_clean.replace(",", "")
        else:
            s_clean = s_clean.replace(",", ".")

    try:
        f_val = float(s_clean)
        if is_neg:
            f_val = -abs(f_val)
        return int(round(f_val)), True
    except (ValueError, TypeError):
        return 0, False


def _get_stmt_period_col(df: pd.DataFrame, periodo: str) -> Optional[Any]:
    """
    Obtiene el nombre exacto de la columna del período actual en un DataFrame de EEFF,
    saltando siempre columnas 'Clasificación' y 'Nota' / 'Notas'.
    """
    if df is None or df.empty or len(df.columns) < 2:
        return None
    p_clean = str(periodo).strip()
    for col in df.columns[1:]:
        if str(col).strip() == p_clean or str(col).strip().startswith(p_clean):
            return col
    # Buscar primera columna que no sea 'Nota'
    for col in df.columns[1:]:
        c_norm = _norm_txt(col)
        if c_norm not in ("nota", "notas", "note", "notes"):
            return col
    return df.columns[1]


def _extract_stmt_line(df: pd.DataFrame, periodo: str, aliases: List[str]) -> Tuple[str, int, bool]:
    """
    Extrae (etiqueta_original, valor_M$, existe_linea) desde #BALANCE o #ER.
    """
    if df is None or df.empty:
        return (aliases[0] if aliases else ""), 0, False
    col_target = _get_stmt_period_col(df, periodo)
    if col_target is None:
        return (aliases[0] if aliases else ""), 0, False

    norm_aliases = [_norm_txt(a) for a in aliases]
    # 1. Coincidencia exacta normalizada en orden de prioridad de alias
    for target_norm in norm_aliases:
        for _, row in df.iterrows():
            raw_lbl = row.iloc[0]
            lbl_norm = _norm_txt(raw_lbl)
            if lbl_norm == target_norm:
                val_int, _ = _parse_int_m(row[col_target])
                return str(raw_lbl).strip(), val_int, True

    return (aliases[0] if aliases else ""), 0, False


def _extract_patrimonio_values(df_pat: pd.DataFrame) -> Dict[str, Any]:
    """
    Extrae los saldos del bloque del período actual en #PATRIMONIO:
    Capital emitido, Otras reservas, Ganancias (pérdidas) acumuladas, Patrimonio Total y Ganancia (pérdida).
    """
    out = {
        "exists": False,
        "capital": 0,
        "reservas": 0,
        "acumulados": 0,
        "total": 0,
        "resultado": 0,
    }
    if df_pat is None or df_pat.empty or len(df_pat.columns) < 5:
        return out

    # Identificar columnas por nombre
    col_cap = col_acum = col_res = col_tot = None
    for c in df_pat.columns[1:]:
        cn = _norm_txt(c)
        if "capital" in cn and col_cap is None:
            col_cap = c
        elif ("acumulad" in cn or "resultados" in cn) and col_acum is None:
            col_acum = c
        elif "reserva" in cn and col_res is None:
            col_res = c
        elif "total" in cn and col_tot is None:
            col_tot = c

    if col_cap is None:
        col_cap = df_pat.columns[1]
    if col_acum is None:
        col_acum = df_pat.columns[2]
    if col_res is None:
        col_res = df_pat.columns[3]
    if col_tot is None:
        col_tot = df_pat.columns[4]

    # Solo recorrer el primer bloque (período actual), deteniéndose si aparece la cabecera del comparativo
    saldo_rows = []
    for r_i in range(len(df_pat)):
        lbl_norm = _norm_txt(df_pat.iloc[r_i, 0])
        c1_norm = _norm_txt(df_pat.iloc[r_i][col_cap])
        if r_i > 2 and "capital emitido" in c1_norm:
            break
        if lbl_norm in ("ganancia (perdida)", "ganancia (perdida) del ejercicio", "resultado del ejercicio"):
            val_r, _ = _parse_int_m(df_pat.iloc[r_i][col_tot])
            out["resultado"] = val_r
        if lbl_norm.startswith("saldo al"):
            saldo_rows.append(r_i)

    if saldo_rows:
        end_idx = saldo_rows[-1]  # Último 'Saldo al ...' del primer bloque = saldo final del período actual
        out["capital"], _ = _parse_int_m(df_pat.iloc[end_idx][col_cap])
        out["acumulados"], _ = _parse_int_m(df_pat.iloc[end_idx][col_acum])
        out["reservas"], _ = _parse_int_m(df_pat.iloc[end_idx][col_res])
        out["total"], _ = _parse_int_m(df_pat.iloc[end_idx][col_tot])
        out["exists"] = True

    return out


def _extract_efe_ending_cash(df_efe: pd.DataFrame, periodo: str) -> Tuple[int, bool]:
    """Extrae el 'Saldo final de efectivo y equivalentes al efectivo' desde #EFE."""
    if df_efe is None or df_efe.empty:
        return 0, False
    col_target = _get_stmt_period_col(df_efe, periodo)
    if col_target is None:
        return 0, False
    for r_i in range(len(df_efe) - 1, -1, -1):
        lbl = _norm_txt(df_efe.iloc[r_i, 0])
        if "saldo final" in lbl and "efectivo" in lbl:
            val, _ = _parse_int_m(df_efe.iloc[r_i][col_target])
            return val, True
    return 0, False


def _is_period_date_match(cell_text: str, periodo: str) -> bool:
    """Verifica si un texto de encabezado corresponde a la fecha de cierre de `periodo` (ej. '2026-10' -> '.10.2026')."""
    s = _norm_txt(cell_text)
    if not s or "-" not in str(periodo):
        return False
    yr, mo = str(periodo).strip().split("-")[:2]
    if f".{mo}.{yr}" in s or f"/{mo}/{yr}" in s or f"{yr}-{mo}" in s:
        return True
    return False


def _is_any_date_token(cell_text: str) -> bool:
    """Detecta si una celda de encabezado contiene una fecha de corte (ej. 31.10.2026, 31.12.2025, 2026-06-30)."""
    s = str(cell_text or "").strip()
    if not s or s.lower() in ("nan", "nat", "none"):
        return False
    if re.search(r"\b\d{2}[./]\d{2}[./]20\d{2}\b", s):
        return True
    if re.search(r"\b20\d{2}-\d{2}", s):
        return True
    return False


class NoteTableExtractor:
    """Extractor determinístico de celdas de totales y subtotales desde cuadros evaluados de Notas (#N04..#N26)."""

    @staticmethod
    def get_tables(note_elements_map: Dict[str, list], note_code: str) -> List[pd.DataFrame]:
        els = note_elements_map.get(note_code, []) or []
        tables = []
        for item in els:
            if isinstance(item, (list, tuple)) and len(item) >= 2 and item[0] == "table":
                df_t = item[1]
                if isinstance(df_t, pd.DataFrame) and not df_t.empty:
                    # Descartar placeholder 1x1 de nota inexistente en plantilla
                    if df_t.shape == (1, 1) and "no se encuentra en la plantilla" in _norm_txt(df_t.iloc[0, 0]):
                        continue
                    tables.append(df_t)
        return tables

    @staticmethod
    def get_subtable(note_elements_map: Dict[str, list], note_code: str, sub_idx: int) -> Optional[pd.DataFrame]:
        tables = NoteTableExtractor.get_tables(note_elements_map, note_code)
        if 1 <= sub_idx <= len(tables):
            return tables[sub_idx - 1]
        return None

    @staticmethod
    def find_target_col(
        df: pd.DataFrame,
        periodo: str,
        col_subheader: Optional[str] = None,
        header_rows: int = 5,
    ) -> Optional[int]:
        """
        Identifica el índice de columna (0-based) del período actual, soportando:
        - Tablas simples con columna de fecha (ej. 31.10.2026 | 31.12.2025)
        - Tablas bidimensionales con fecha combinada arriba + Corriente / No corriente o Activos / Pasivos abajo
        - Tablas de Activo Fijo / Intangibles / Derechos de Uso con subcolumna 'Neto' / 'Activo fijo neto'
        """
        if df is None or df.empty or df.shape[1] < 2:
            return None

        n_rows = min(header_rows, len(df))
        n_cols = df.shape[1]

        # Caso especial: columna 'Neto' / 'Activo fijo neto' del primer bloque (período actual)
        if col_subheader and _norm_txt(col_subheader) in ("neto", "activo fijo neto"):
            for r_i in range(n_rows):
                for c_i in range(1, n_cols):
                    cell_n = _norm_txt(df.iloc[r_i, c_i])
                    if cell_n in ("neto", "activo fijo neto"):
                        return c_i

        # Caso especial: Segmentos (#N26.1 / #N26.2) -> primera columna 'Total' corresponde al período actual
        if col_subheader and _norm_txt(col_subheader) == "total_segmento":
            for r_i in range(n_rows):
                for c_i in range(1, n_cols):
                    if _norm_txt(df.iloc[r_i, c_i]) == "total":
                        return c_i

        # Construir mapa por columna de fecha propagada (ffill sobre encabezados combinados) y subencabezados
        col_dates: List[str] = [""] * n_cols
        for r_i in range(n_rows):
            row_dates = [_norm_txt(df.iloc[r_i, c_i]) if _is_any_date_token(df.iloc[r_i, c_i]) else "" for c_i in range(n_cols)]
            if any(row_dates):
                active_dt = ""
                for c_i in range(1, n_cols):
                    if row_dates[c_i]:
                        active_dt = row_dates[c_i]
                    if active_dt and not col_dates[c_i]:
                        col_dates[c_i] = active_dt

        # Propagar también encabezados de clasificación (ej. 'Corriente' / 'No corriente' en fila 0 de #N14.1)
        col_tags: List[List[str]] = [[] for _ in range(n_cols)]
        sub_keywords = ("corriente", "corrientes", "no corriente", "no corrientes", "activos", "pasivos", "activo", "pasivo")
        for r_i in range(n_rows):
            row_vals = [_norm_txt(df.iloc[r_i, c_i]) for c_i in range(n_cols)]
            has_kw = any(v in sub_keywords for v in row_vals[1:])
            active_kw = ""
            for c_i in range(1, n_cols):
                v = row_vals[c_i]
                if has_kw:
                    if v in sub_keywords:
                        active_kw = v
                    if active_kw:
                        col_tags[c_i].append(active_kw)
                elif v and not _is_any_date_token(v) and v != "m$":
                    col_tags[c_i].append(v)

        target_sub = _norm_txt(col_subheader) if col_subheader else None

        def _matches_sub(tags: List[str], target: str) -> bool:
            if target in ("corriente", "corrientes"):
                return any(t in ("corriente", "corrientes") for t in tags)
            if target in ("no corriente", "no corrientes"):
                return any(t in ("no corriente", "no corrientes") for t in tags)
            if target in ("activo", "activos"):
                return any(t in ("activo", "activos") for t in tags)
            if target in ("pasivo", "pasivos"):
                return any(t in ("pasivo", "pasivos") for t in tags)
            return any(target == t or target in t for t in tags)

        # 1. Buscar columna que coincida con fecha del período actual + subencabezado (si se pidió)
        for c_i in range(1, n_cols):
            if _is_period_date_match(col_dates[c_i], periodo):
                if target_sub is None or _matches_sub(col_tags[c_i], target_sub):
                    return c_i

        # 2. Fallback si la plantilla no puso fecha legible pero sí subencabezado (primer bloque = período actual)
        if target_sub is not None:
            for c_i in range(1, n_cols):
                if _matches_sub(col_tags[c_i], target_sub):
                    return c_i

        # 3. Fallback tablas simples (ej. #N21.3 o #N18.3): primera columna con 'M$' o columna 1
        for r_i in range(n_rows):
            for c_i in range(1, n_cols):
                if _norm_txt(df.iloc[r_i, c_i]) == "m$":
                    return c_i
        return 1

    @staticmethod
    def extract_value(
        note_elements_map: Dict[str, list],
        note_code: str,
        sub_idx: int,
        periodo: str,
        row_labels: List[str],
        col_subheader: Optional[str] = None,
        occurrence: str = "first",
        after_section_label: Optional[str] = None,
        allow_last_row_blank_label: bool = False,
    ) -> Tuple[int, bool]:
        """
        Extrae (valor_entero_M$, encontrado) de un sub-cuadro específico (#Nxx.sub_idx).
        """
        df = NoteTableExtractor.get_subtable(note_elements_map, note_code, sub_idx)
        if df is None or df.empty:
            return 0, False

        # Caso especial #N20.4 ('Otras reservas' trae el monto en r0 col 1: ['Otras reservas', '108556', ...])
        if note_code == "#N20" and sub_idx == 4:
            for r_i in range(len(df)):
                if _norm_txt(df.iloc[r_i, 0]) == "otras reservas":
                    v_h, has_h = _parse_int_m(df.iloc[r_i, 1] if df.shape[1] > 1 else None)
                    if has_h:
                        return v_h, True

        col_idx = NoteTableExtractor.find_target_col(df, periodo, col_subheader=col_subheader)
        if col_idx is None or col_idx >= df.shape[1]:
            return 0, False

        start_row = 0
        if after_section_label:
            sec_norm = _norm_txt(after_section_label)
            found_sec = False
            for r_i in range(len(df)):
                row_text = " ".join(_norm_txt(x) for x in df.iloc[r_i].tolist()[:3])
                if sec_norm in row_text:
                    start_row = r_i + 1
                    found_sec = True
                    break
            if not found_sec:
                return 0, False

        norm_targets = [_norm_txt(lbl) for lbl in row_labels]
        matched_rows: List[int] = []
        for r_i in range(start_row, len(df)):
            # Buscar etiqueta en columna 0 (o columna 1 si col 0 está vacía como en #N14.3 ['', 'Total', '', '-'])
            c0 = _norm_txt(df.iloc[r_i, 0])
            c1 = _norm_txt(df.iloc[r_i, 1]) if df.shape[1] > 1 and not c0 else ""
            cand_lbl = c0 if c0 else c1
            if cand_lbl in norm_targets:
                matched_rows.append(r_i)

        if matched_rows:
            target_r = matched_rows[0] if occurrence == "first" else matched_rows[-1]
            val_int, _ = _parse_int_m(df.iloc[target_r, col_idx])
            return val_int, True

        # Fallback controlado para cuadros cuya fila final de suma no tiene texto en col 0 (ej. #N22.2 r5: ['', '3230342', '0'])
        if allow_last_row_blank_label and len(df) >= 2:
            last_r = len(df) - 1
            if _norm_txt(df.iloc[last_r, 0]) == "":
                val_int, has_v = _parse_int_m(df.iloc[last_r, col_idx])
                if has_v:
                    return val_int, True

        return 0, False


class ValidationTieOutEngine:
    """
    Motor Simplificado de Validación y Cuadratura Directa (EEFF vs. Notas y Cruces Internos) en M$.
    Tolerancia estricta CERO (Diferencia == 0 M$).
    """

    @staticmethod
    def get_available_validation_periods(empresa_seleccionada: str) -> List[str]:
        """Retorna la lista unificada de períodos disponibles (cerrados en Snapshot + abiertos en BD)."""
        is_cons = str(empresa_seleccionada).startswith("[GRUPO] ")
        clean_name = str(empresa_seleccionada).replace("[GRUPO] ", "").strip()
        periods_set = set()

        # 1. Períodos cerrados en ClosingSnapshotRecord
        try:
            from src.models.historical_data import ClosingSnapshotRecord, HistoricalDataRecord
            db = SessionLocal()
            try:
                snap_recs = db.query(ClosingSnapshotRecord.periodo).filter_by(
                    empresa=empresa_seleccionada, is_active=True
                ).all()
                for (p,) in snap_recs:
                    if p:
                        periods_set.add(str(p).strip())

                if is_cons:
                    hist_recs = db.query(HistoricalDataRecord.periodo).distinct().all()
                    for (p,) in hist_recs:
                        if p:
                            periods_set.add(str(p).strip())
                    # También incluir períodos activos de las filiales del grupo
                    grp = db.query(ConsolidationGroup).filter_by(nombre_grupo=clean_name).first()
                    if grp and hasattr(grp, "miembros"):
                        for m in grp.miembros:
                            for p_sub in TrialBalanceDB.get_available_periods(m.empresa_nombre):
                                periods_set.add(str(p_sub).strip())
            finally:
                db.close()
        except Exception:
            pass

        # 2. Períodos abiertos en TrialBalanceDB / PlCuboDB
        if not is_cons:
            for p in TrialBalanceDB.get_available_periods(clean_name):
                periods_set.add(str(p).strip())
        else:
            # Revisar sociedades estándar por si no estaban en relación ORM
            for emp_dir in ("Pacifico Cable SpA", "Db Terra Chile Holdco SpA", "Db Terra Chile Parent SpA"):
                for p in TrialBalanceDB.get_available_periods(emp_dir):
                    periods_set.add(str(p).strip())

        valid_periods = sorted(
            [p for p in periods_set if re.match(r"^20\d{2}-\d{2}$", p)],
            reverse=True,
        )
        return valid_periods if valid_periods else ["2026-10", "2025-12"]

    @staticmethod
    def _load_bundle(empresa_full: str, periodo: str, is_consolidated: bool) -> Dict[str, Any]:
        """
        Carga los Estados Financieros (#BALANCE, #ER, #PATRIMONIO, #EFE) y todos los cuadros de Notas (#N04..#N26)
        en escala M$ (scale_factor = 1000.0):
        - Si el período tiene Snapshot activo en Bóveda -> lectura instantánea desde SnapshotVaultManager.
        - Si el período está abierto -> generación en memoria idéntica al proceso de cierre (con Notas Externas y fórmulas).
        """
        base_scale = 1000.0
        clean_name = empresa_full.replace("[GRUPO] ", "").strip()

        if SnapshotVaultManager.has_active_snapshot(empresa_full, periodo):
            bal_res = SnapshotVaultManager.get_statement(empresa_full, periodo, "#BALANCE", scale_factor=base_scale)
            er_res = SnapshotVaultManager.get_statement(empresa_full, periodo, "#ER", scale_factor=base_scale)
            pat_res = SnapshotVaultManager.get_statement(empresa_full, periodo, "#PATRIMONIO", scale_factor=base_scale)
            efe_res = SnapshotVaultManager.get_statement(empresa_full, periodo, "#EFE", scale_factor=base_scale)

            notes_map: Dict[str, list] = {}
            full_reg = get_full_note_registry()
            for code in full_reg.keys():
                els = SnapshotVaultManager.get_note_elements(empresa_full, periodo, code, scale_factor=base_scale)
                if els:
                    notes_map[code] = els

            return {
                "source_mode": "snapshot",
                "source_label": f"🔒 Snapshot Congelado en Bóveda ({periodo})",
                "balance": bal_res[0] if bal_res else None,
                "er": er_res[0] if er_res else None,
                "patrimonio": pat_res[0] if pat_res else None,
                "efe": efe_res[0] if efe_res else None,
                "notes": notes_map,
            }

        # =========================================================================
        # MODO PERÍODO ABIERTO (PRE-CIERRE EN VIVO EN M$)
        # =========================================================================
        yr_str, mo_str = periodo.split("-")[:2]
        prev_yr = int(yr_str) - 1
        comp_period_bal = f"{prev_yr}-12"
        comp_period_pl = f"{prev_yr}-{mo_str}"

        from src.core.word_template_engine import WordTemplateEngine, load_mappings_for_entity
        map_balance_df, map_pl_df = load_mappings_for_entity(empresa_full)
        wte = WordTemplateEngine(None)

        empresa_path = os.path.join("data", "empresas", empresa_full)
        if not os.path.exists(empresa_path):
            empresa_path = os.path.join("data", "empresas", clean_name)

        bal_preview = None
        er_preview = None
        pat_preview = None
        efe_preview = None
        df_hoja_act = None
        df_hoja_comp = None

        # 1. #BALANCE y #ER en vivo
        try:
            if is_consolidated:
                db = SessionLocal()
                try:
                    grp = db.query(ConsolidationGroup).filter_by(nombre_grupo=clean_name).first()
                    grupo_id = grp.id if grp else None
                finally:
                    db.close()
                if grupo_id:
                    from src.core.consolidacion_engine import generar_hoja_trabajo
                    df_hoja_act, _ = generar_hoja_trabajo(grupo_id, periodo)
                    df_hoja_comp, _ = generar_hoja_trabajo(grupo_id, comp_period_bal)
                    bal_preview, _, _, _ = wte._get_consolidated_bal_and_pl(
                        grupo_id=grupo_id,
                        periodo_actual=periodo,
                        comp_period=comp_period_bal,
                        scale_factor=base_scale,
                        empresa_path=empresa_path,
                        grupo_name=clean_name,
                    )
                    _, er_preview, _, _ = wte._get_consolidated_bal_and_pl(
                        grupo_id=grupo_id,
                        periodo_actual=periodo,
                        comp_period=comp_period_pl,
                        scale_factor=base_scale,
                        empresa_path=empresa_path,
                        grupo_name=clean_name,
                    )
            else:
                from src.reporting.balance_generator import BalanceGenerator
                from src.reporting.er_generator import ERGenerator

                tb_df = TrialBalanceDB.get_trial_balance(clean_name, periodo)
                tb_comp = TrialBalanceDB.get_trial_balance(clean_name, comp_period_bal)
                if tb_df is not None and not tb_df.empty:
                    t_bal = os.path.join(empresa_path, "Balance clasificado.xlsx")
                    if not os.path.exists(t_bal):
                        t_bal = os.path.join("templates", "Balance clasificado.xlsx")
                    res_bal = BalanceGenerator(t_bal).generate(
                        tb_df, map_balance_df, scale_factor=base_scale,
                        tb_df_comp=tb_comp, periodo_actual_str=periodo, periodo_comp_str=comp_period_bal
                    )
                    bal_preview = wte._extract_balance_preview(res_bal, periodo, comp_period_bal)
                    bal_preview = SnapshotVaultManager.apply_closed_comp_to_statement(
                        bal_preview, empresa_full, "#BALANCE", periodo, comp_period_bal, scale_factor=base_scale
                    )

                pl_wide = PlCuboDB.get_pl_cubo(clean_name, periodo)
                pl_comp_wide = PlCuboDB.get_pl_cubo(clean_name, comp_period_pl)
                if pl_wide is not None and not pl_wide.empty:
                    t_er = os.path.join(empresa_path, "Estado de Resultados Clasificados.xlsx")
                    if not os.path.exists(t_er):
                        t_er = os.path.join("templates", "Estado de Resultados Clasificados.xlsx")
                    res_er, _ = ERGenerator(t_er).generate(
                        pl_wide, scale_factor=base_scale, pl_df_comp=pl_comp_wide,
                        periodo_actual_str=periodo, periodo_comp_str=comp_period_pl
                    )
                    er_preview = wte._extract_er_preview(res_er, periodo, comp_period_pl)
                    er_preview = SnapshotVaultManager.apply_closed_comp_to_statement(
                        er_preview, empresa_full, "#ER", periodo, comp_period_pl, scale_factor=base_scale
                    )
        except Exception as e_st:
            print(f"Advertencia generando Balance/ER en vivo para validación ({empresa_full} {periodo}): {e_st}")

        # 2. #PATRIMONIO en vivo
        try:
            t_pat = os.path.join(empresa_path, "Estado de Cambios en el Patrimonio.xlsx")
            if not os.path.exists(t_pat):
                t_pat = os.path.join("templates", "Estado de Cambios en el Patrimonio.xlsx")
            if os.path.exists(t_pat) and bal_preview is not None:
                from src.core.excel_utils import clean_patrimonio_preview_df, detect_patrimonio_skiprows
                from src.reporting.patrimonio_generator import PatrimonioGenerator

                pl_for_pat = er_preview
                res_pat = PatrimonioGenerator(t_pat).generate(
                    bal_preview_df=bal_preview,
                    pl_preview_df=pl_for_pat,
                    periodo_actual_str=periodo,
                    periodo_comp_str=comp_period_bal,
                    empresa=empresa_full,
                )
                res_pat.seek(0)
                pat_skip = detect_patrimonio_skiprows(res_pat)
                res_pat.seek(0)
                pat_preview = clean_patrimonio_preview_df(pd.read_excel(res_pat, skiprows=pat_skip))
        except Exception as e_pat:
            print(f"Advertencia generando Patrimonio en vivo ({empresa_full} {periodo}): {e_pat}")

        # 3. #EFE en vivo
        try:
            t_cf = os.path.join(empresa_path, "Estado de Flujos de Efectivo.xlsx")
            if not os.path.exists(t_cf):
                t_cf = os.path.join("templates", "Estado de Flujos de Efectivo.xlsx")
            if os.path.exists(t_cf) and bal_preview is not None:
                from src.reporting.cash_flow_generator import CashFlowGenerator
                res_cf, _ = CashFlowGenerator(t_cf).generate(
                    empresa=empresa_full,
                    periodo_actual_str=periodo,
                    periodo_comp_str=comp_period_bal,
                    map_balance_df=map_balance_df,
                    map_pl_df=map_pl_df,
                    method="Directo",
                    is_consolidado=is_consolidated,
                    consolidated_hoja_trabajo_df=df_hoja_act,
                    consolidated_hoja_trabajo_comp_df=df_hoja_comp,
                    scale_factor=base_scale,
                )
                efe_preview = wte._extract_cf_preview(res_cf, periodo, comp_period_bal)
        except Exception as e_cf:
            print(f"Advertencia generando EFE en vivo ({empresa_full} {periodo}): {e_cf}")

        # 4. Notas (#N04..#N26) en vivo por lotes con Notas Externas + Evaluación de Fórmulas
        notes_map: Dict[str, list] = {}
        template_nota = "Plantilla de notas_v1.xlsx"
        if os.path.exists(template_nota) and (bal_preview is not None or er_preview is not None):
            try:
                from src.core.excel_utils import extract_named_ranges_from_excel
                from src.core.external_notes_manager import ExternalNotesManager
                from src.reporting.note_generator import NoteGenerator
                from src.ui_pages.informes_y_notas import (
                    evaluate_formulas_in_workbook,
                    load_all_entity_contexts,
                    split_sheet_into_elements,
                )

                engine_notes = NoteGenerator(template_nota)
                full_registry = get_full_note_registry()

                # Para validación del período actual basta 1 sola evaluación unificada de todas las hojas de notas
                all_target_sheets: List[str] = []
                for code, info in full_registry.items():
                    t_sheets = info.get("sheets", [])
                    if not t_sheets:
                        continue
                    if info.get("consolidated_only", False) and not is_consolidated:
                        continue
                    for s_n in t_sheets:
                        if s_n not in all_target_sheets:
                            all_target_sheets.append(s_n)

                if all_target_sheets:
                    entity_contexts = load_all_entity_contexts(
                        active_entity=empresa_full,
                        periodo_actual=periodo,
                        periodo_comp=comp_period_bal,
                        map_balance_df=map_balance_df,
                        map_pl_df=map_pl_df,
                    )
                    wb_out = engine_notes.generate(
                        sheet_names=all_target_sheets,
                        entity_contexts=entity_contexts,
                        active_entity_name=empresa_full,
                        is_consolidated=is_consolidated,
                        scale_factor=base_scale,
                        periodo_actual_str=periodo,
                        periodo_comp_str=comp_period_bal,
                        map_balance_df=map_balance_df,
                        map_pl_df=map_pl_df,
                        evaluate_formulas=False,
                    )
                    wb_out = ExternalNotesManager.inject_external_data(
                        wb_out,
                        empresa=empresa_full,
                        periodo_actual=periodo,
                        periodo_comp=comp_period_bal,
                    )
                    # Evaluar fórmulas únicamente en las hojas de notas activas y reconciliar redondeos (<= 4 M$)
                    from src.ui_pages.informes_y_notas import evaluate_openpyxl_formula
                    from src.core.rounding_reconciler import RoundingReconciler
                    wb_out.seek(0)
                    raw_note_buf = wb_out.getvalue()
                    wb_raw_obj = openpyxl.load_workbook(io.BytesIO(raw_note_buf), data_only=False)
                    wb_eval_obj = openpyxl.load_workbook(io.BytesIO(raw_note_buf), data_only=False)
                    target_set = set(all_target_sheets)
                    for sh_name in wb_eval_obj.sheetnames:
                        if sh_name not in target_set:
                            continue
                        ws_n = wb_eval_obj[sh_name]
                        for r_idx in range(1, ws_n.max_row + 1):
                            for c_idx in range(1, ws_n.max_column + 1):
                                c_obj = ws_n.cell(row=r_idx, column=c_idx)
                                c_val = c_obj.value
                                if isinstance(c_val, str) and c_val.startswith("="):
                                    ev = evaluate_openpyxl_formula(c_val, ws_n, c_idx, r_idx, set())
                                    if ev is not None:
                                        c_obj.value = ev

                    RoundingReconciler.reconcile_note_workbooks_inplace(
                        wb_raw=wb_raw_obj,
                        wb_eval=wb_eval_obj,
                        df_bal=bal_preview,
                        df_er=er_preview,
                        periodo_actual=periodo,
                        is_consolidated=is_consolidated,
                    )

                    for code, info in full_registry.items():
                        t_sheets = info.get("sheets", [])
                        if not t_sheets or (info.get("consolidated_only", False) and not is_consolidated):
                            continue
                        named_ranges = extract_named_ranges_from_excel(wb_eval_obj, code, wb=wb_eval_obj)
                        raw_els = []
                        if named_ranges:
                            for _, df_chunk in named_ranges:
                                raw_els.append(("table", df_chunk, t_sheets[0], None))
                        else:
                            for s_name in t_sheets:
                                if s_name in wb_eval_obj.sheetnames:
                                    preview_df = pd.DataFrame(wb_eval_obj[s_name].values)
                                    for s_item in split_sheet_into_elements(preview_df, sheet_name=s_name):
                                        raw_els.append((s_item[0], s_item[1], s_name, None))
                        if raw_els:
                            notes_map[code] = raw_els
            except Exception as e_notes:
                print(f"Advertencia evaluando notas en vivo ({empresa_full} {periodo}): {e_notes}")

        return {
            "source_mode": "live",
            "source_label": f"⚡ Período Abierto — Cálculo Vivo en M$ ({periodo})",
            "balance": bal_preview,
            "er": er_preview,
            "patrimonio": pat_preview,
            "efe": efe_preview,
            "notes": notes_map,
        }

    @staticmethod
    def obtener_matriz_tie_out(empresa_o_grupo: str, periodo: str, is_consolidated: bool = False) -> Tuple[pd.DataFrame, dict]:
        """
        Ejecuta la validación directa EEFF vs. Notas y Cruces Internos en M$ con Tolerancia Estricta CERO (diff == 0).
        """
        t0 = time.perf_counter()
        clean_name = str(empresa_o_grupo).replace("[GRUPO] ", "").strip()
        empresa_full = f"[GRUPO] {clean_name}" if is_consolidated else clean_name

        bundle = ValidationTieOutEngine._load_bundle(empresa_full, periodo, is_consolidated)
        df_bal = bundle["balance"]
        df_er = bundle["er"]
        df_pat = bundle["patrimonio"]
        df_efe = bundle["efe"]
        notes_map = bundle["notes"]

        pat_vals = _extract_patrimonio_values(df_pat)
        efe_cash, has_efe_cash = _extract_efe_ending_cash(df_efe, periodo)

        # Semáforos estructurales superiores
        _, total_activos, has_tot_act = _extract_stmt_line(df_bal, periodo, ["Total activos"])
        _, total_pas_pat, has_tot_pp = _extract_stmt_line(df_bal, periodo, ["Total patrimonio y pasivos"])
        diff_eq = total_activos - total_pas_pat

        _, res_er, has_res_er = _extract_stmt_line(
            df_er,
            periodo,
            [
                "Ganancias (Pérdida) del Ejercicio",
                "Ganancia (Pérdida) del Ejercicio",
                "(Pérdida) procedente de operaciones continuadas",
                "Ganancia (pérdida) procedente de operaciones continuadas",
                "Propietarios de la controladora",
            ],
        )
        res_pat = pat_vals["resultado"]
        diff_res = res_er - res_pat if (has_res_er and pat_vals["exists"]) else 0

        _, bal_efectivo, has_bal_ef = _extract_stmt_line(df_bal, periodo, ["Efectivo y efectivo equivalente"])
        diff_efe = bal_efectivo - efe_cash if (has_bal_ef and has_efe_cash) else 0

        health_checks = {
            "is_valid": True,
            "total_reglas": 0,
            "total_cuadrados": 0,
            "total_descuadres": 0,
            "source_mode": bundle["source_mode"],
            "source_label": bundle["source_label"],
            "ecuacion_patrimonial": {
                "activos": total_activos,
                "pasivos_patrimonio": total_pas_pat,
                "diff": diff_eq,
                "ok": bool(has_tot_act and has_tot_pp and diff_eq == 0),
            },
            "cruce_resultado": {
                "er": res_er,
                "patrimonio": res_pat,
                "diff": diff_res,
                "ok": bool(has_res_er and pat_vals["exists"] and diff_res == 0),
            },
            "cruce_efectivo": {
                "balance": bal_efectivo,
                "efe": efe_cash,
                "diff": diff_efe,
                "ok": bool(has_bal_ef and has_efe_cash and diff_efe == 0),
            },
            "elapsed_sec": 0.0,
        }

        if (df_bal is None or df_bal.empty) and (df_er is None or df_er.empty):
            health_checks["is_valid"] = False
            health_checks["elapsed_sec"] = round(time.perf_counter() - t0, 3)
            return pd.DataFrame(), health_checks

        rows: List[Dict[str, Any]] = []

        def _add_check(
            sort_order: float,
            code_badge: str,
            nivel: str,
            reporte: str,
            origen_lbl: str,
            val_origen: int,
            destino_lbl: str,
            val_destino: int,
            has_destino: bool = True,
        ):
            diff = int(val_origen) - int(val_destino)
            if not has_destino and val_origen != 0:
                estado = "❌ Descuadre"
            elif diff == 0:
                estado = "✅ Cuadrado"
            else:
                estado = "❌ Descuadre"

            rows.append({
                "_seq": len(rows),
                "Código": code_badge,
                "Rubro EEFF": origen_lbl,
                "Nota": destino_lbl,
                "Saldo EEFF (M$)": int(val_origen),
                "Saldo Nota (M$)": int(val_destino),
                "Diferencia (M$)": int(diff),
                "Estado": estado,
                "Reporte": reporte,
                "Nivel": nivel,
            })

        # =========================================================================
        # A. BALANCE CLASIFICADO (#BALANCE) vs. NOTAS (#N04 - #N20)
        # =========================================================================
        if df_bal is not None and not df_bal.empty:
            # #N04 Efectivo y efectivo equivalente
            lbl, v_eeff, ex = _extract_stmt_line(df_bal, periodo, ["Efectivo y efectivo equivalente"])
            if ex:
                v_n04_1, ok_n04_1 = NoteTableExtractor.extract_value(notes_map, "#N04", 1, periodo, ["Total", "Totales"])
                _add_check(4.1, "#N04.1", "EEFF vs. Nota", "Balance Clasificado", lbl, v_eeff, "#N04.1 — Efectivo (Composición del rubro)", v_n04_1, ok_n04_1)

                # Cruces internos de Nota 4 (#N04.2 Monedas y #N04.3 Detalle Fondos Mutuos)
                v_n04_2, ok_n04_2 = NoteTableExtractor.extract_value(notes_map, "#N04", 2, periodo, ["Total", "Totales"])
                if ok_n04_2 or v_eeff != 0:
                    _add_check(4.2, "#N04.2", "Cuadratura Interna Nota", "Balance Clasificado", "#N04.1 — Total Efectivo", v_n04_1 if ok_n04_1 else v_eeff, "#N04.2 — Efectivo (Composición por monedas)", v_n04_2, ok_n04_2)

                v_fm_linea, ok_fm_linea = NoteTableExtractor.extract_value(notes_map, "#N04", 1, periodo, ["Fondos mutuos"])
                v_n04_3, ok_n04_3 = NoteTableExtractor.extract_value(notes_map, "#N04", 3, periodo, ["Total", "Totales"])
                if ok_fm_linea and (v_fm_linea != 0 or v_n04_3 != 0):
                    _add_check(4.3, "#N04.3", "Cuadratura Interna Nota", "Balance Clasificado", "#N04.1 — Línea 'Fondos mutuos'", v_fm_linea, "#N04.3 — Detalle de Fondos Mutuos (Total)", v_n04_3, ok_n04_3)

            # #N05 Otros activos no financieros, corrientes
            lbl, v_eeff, ex = _extract_stmt_line(df_bal, periodo, ["Otros activos no financieros, corrientes"])
            if ex:
                v_n, ok_n = NoteTableExtractor.extract_value(notes_map, "#N05", 1, periodo, ["Total", "Totales"])
                _add_check(5.1, "#N05.1", "EEFF vs. Nota", "Balance Clasificado", lbl, v_eeff, "#N05.1 — Otros activos no financieros, corrientes", v_n, ok_n)

            # #N06 Deudores comerciales y otras cuentas por cobrar, corrientes
            lbl, v_eeff, ex = _extract_stmt_line(df_bal, periodo, ["Deudores comerciales y otras cuentas por cobrar, corrientes"])
            if ex:
                v_n, ok_n = NoteTableExtractor.extract_value(notes_map, "#N06", 1, periodo, ["Total", "Totales"])
                _add_check(6.1, "#N06.1", "EEFF vs. Nota", "Balance Clasificado", lbl, v_eeff, "#N06.1 — Deudores comerciales y otras cuentas por cobrar", v_n, ok_n)

            # #N07 Inventarios
            lbl, v_eeff, ex = _extract_stmt_line(df_bal, periodo, ["Inventarios"])
            if ex:
                v_n, ok_n = NoteTableExtractor.extract_value(notes_map, "#N07", 1, periodo, ["Total", "Totales"])
                _add_check(7.1, "#N07.1", "EEFF vs. Nota", "Balance Clasificado", lbl, v_eeff, "#N07.1 — Inventarios (Composición)", v_n, ok_n)

            # #N08 Activos intangibles distinto a la plusvalía
            lbl, v_eeff, ex = _extract_stmt_line(df_bal, periodo, ["Activos intangibles distinto a la plusvalía", "Activos intangibles distintos de la plusvalía"])
            if ex:
                v_n, ok_n = NoteTableExtractor.extract_value(notes_map, "#N08", 1, periodo, ["Total", "Totales"], col_subheader="Neto", occurrence="first")
                _add_check(8.1, "#N08.1", "EEFF vs. Nota", "Balance Clasificado", lbl, v_eeff, "#N08.1 — Intangibles (Col: Neto)", v_n, ok_n)

            # #N09 Propiedades, plantas y equipos
            lbl, v_eeff, ex = _extract_stmt_line(df_bal, periodo, ["Propiedades, plantas y equipos", "Propiedades, planta y equipo"])
            if ex:
                v_n, ok_n = NoteTableExtractor.extract_value(notes_map, "#N09", 1, periodo, ["Totales", "Total"], col_subheader="Activo fijo neto")
                _add_check(9.1, "#N09.1", "EEFF vs. Nota", "Balance Clasificado", lbl, v_eeff, "#N09.1 — Propiedades, planta y equipo (Col: Activo fijo neto)", v_n, ok_n)

            # #N10 Activo por derechos de uso
            lbl, v_eeff, ex = _extract_stmt_line(df_bal, periodo, ["Activo por derechos de uso", "Activos por derechos de uso"])
            if ex:
                v_n, ok_n = NoteTableExtractor.extract_value(notes_map, "#N10", 1, periodo, ["Totales", "Total"], col_subheader="Activo fijo neto")
                _add_check(10.1, "#N10.1", "EEFF vs. Nota", "Balance Clasificado", lbl, v_eeff, "#N10.1 — Activos por derechos de uso (Col: Activo fijo neto)", v_n, ok_n)

            # #N11 Plusvalía
            lbl, v_eeff, ex = _extract_stmt_line(df_bal, periodo, ["Plusvalia", "Plusvalía"])
            if ex:
                v_n, ok_n = NoteTableExtractor.extract_value(notes_map, "#N11", 1, periodo, ["Total", "Totales"])
                _add_check(11.1, "#N11.1", "EEFF vs. Nota", "Balance Clasificado", lbl, v_eeff, "#N11.1 — Plusvalía", v_n, ok_n)

            # #N12 Impuestos corrientes (Activo #N12.1 y Pasivo #N12.2 por separado)
            lbl, v_eeff, ex = _extract_stmt_line(df_bal, periodo, ["Activo por impuestos, corrientes", "Activos por impuestos, corrientes"])
            if ex:
                v_n, ok_n = NoteTableExtractor.extract_value(notes_map, "#N12", 1, periodo, ["Total activos", "Total", "Totales"])
                _add_check(12.1, "#N12.1", "EEFF vs. Nota", "Balance Clasificado", lbl, v_eeff, "#N12.1 — Activos por impuestos corrientes", v_n, ok_n)

            lbl, v_eeff, ex = _extract_stmt_line(df_bal, periodo, ["Pasivo por impuestos, corrientes", "Pasivos por impuestos, corrientes"])
            if ex:
                v_n, ok_n = NoteTableExtractor.extract_value(notes_map, "#N12", 2, periodo, ["Total pasivos", "Total", "Totales"])
                _add_check(12.2, "#N12.2", "EEFF vs. Nota", "Balance Clasificado", lbl, v_eeff, "#N12.2 — Pasivos por impuestos corrientes", v_n, ok_n)

            # #N13 Impuestos diferidos (Col: Activos y Col: Pasivos en #N13.1)
            lbl, v_eeff, ex = _extract_stmt_line(df_bal, periodo, ["Activo por impuestos diferidos, no corrientes", "Activos por impuestos diferidos"])
            if ex:
                v_n, ok_n = NoteTableExtractor.extract_value(notes_map, "#N13", 1, periodo, ["Totales", "Total"], col_subheader="Activos")
                _add_check(13.1, "#N13.1", "EEFF vs. Nota", "Balance Clasificado", lbl, v_eeff, "#N13.1 — Impuestos diferidos (Col: Activos)", v_n, ok_n)

            lbl, v_eeff, ex = _extract_stmt_line(df_bal, periodo, ["Pasivos por impuestos diferidos, no corrientes", "Pasivo por impuestos diferidos, no corrientes"])
            if ex:
                v_n, ok_n = NoteTableExtractor.extract_value(notes_map, "#N13", 1, periodo, ["Totales", "Total"], col_subheader="Pasivos")
                _add_check(13.2, "#N13.1", "EEFF vs. Nota", "Balance Clasificado", lbl, v_eeff, "#N13.1 — Impuestos diferidos (Col: Pasivos)", v_n, ok_n)

            # #N14 Entidades relacionadas (CxC Corriente / No Corriente, CxP Corriente / No Corriente, e Inversión #N14.3)
            lbl, v_eeff, ex = _extract_stmt_line(df_bal, periodo, ["Cuentas por cobrar a entidades relacionadas, corrientes"])
            if ex:
                v_n, ok_n = NoteTableExtractor.extract_value(notes_map, "#N14", 1, periodo, ["Total", "Totales"], col_subheader="Corriente", occurrence="first")
                _add_check(14.1, "#N14.1", "EEFF vs. Nota", "Balance Clasificado", lbl, v_eeff, "#N14.1 — CxC Relacionadas (Col: Corriente)", v_n, ok_n)

            lbl, v_eeff, ex = _extract_stmt_line(df_bal, periodo, ["Cuentas por cobrar a entidades relacionadas, no corrientes"])
            if ex:
                v_n, ok_n = NoteTableExtractor.extract_value(notes_map, "#N14", 1, periodo, ["Total", "Totales"], col_subheader="No corriente", occurrence="first")
                _add_check(14.2, "#N14.1", "EEFF vs. Nota", "Balance Clasificado", lbl, v_eeff, "#N14.1 — CxC Relacionadas (Col: No corriente)", v_n, ok_n)

            lbl, v_eeff, ex = _extract_stmt_line(df_bal, periodo, ["Cuentas por pagar entidades relacionadas, corrientes", "Cuentas por pagar a entidades relacionadas, corrientes"])
            if ex:
                v_n, ok_n = NoteTableExtractor.extract_value(notes_map, "#N14", 1, periodo, ["Total", "Totales"], col_subheader="Corriente", after_section_label="Cuentas por pagar")
                _add_check(14.3, "#N14.1", "EEFF vs. Nota", "Balance Clasificado", lbl, v_eeff, "#N14.1 — CxP Relacionadas (Col: Corriente)", v_n, ok_n)

            lbl, v_eeff, ex = _extract_stmt_line(df_bal, periodo, ["Cuentas por pagar entidades relacionadas, no corrientes", "Cuentas por pagar a entidades relacionadas, no corrientes"])
            if ex:
                v_n, ok_n = NoteTableExtractor.extract_value(notes_map, "#N14", 1, periodo, ["Total", "Totales"], col_subheader="No corriente", after_section_label="Cuentas por pagar")
                _add_check(14.4, "#N14.1", "EEFF vs. Nota", "Balance Clasificado", lbl, v_eeff, "#N14.1 — CxP Relacionadas (Col: No corriente)", v_n, ok_n)

            lbl, v_eeff, ex = _extract_stmt_line(df_bal, periodo, ["Inversion en empresas relacionadas", "Inversión en empresas relacionadas"])
            if ex:
                v_n, ok_n = NoteTableExtractor.extract_value(notes_map, "#N14", 3, periodo, ["Total", "Totales"])
                _add_check(14.5, "#N14.3", "EEFF vs. Nota", "Balance Clasificado", lbl, v_eeff, "#N14.3 — Inversión en empresas relacionadas", v_n, ok_n)

            # #N15 Otros activos financieros (#N15.1) y Otros pasivos financieros (#N15.2) separados Corriente / No Corriente
            lbl, v_eeff, ex = _extract_stmt_line(df_bal, periodo, ["Activos financieros, Corrientes", "Otros activos financieros, corrientes"])
            if ex:
                v_n, ok_n = NoteTableExtractor.extract_value(notes_map, "#N15", 1, periodo, ["Total otros activos financieros", "Total", "Totales"], col_subheader="Corrientes")
                _add_check(15.1, "#N15.1", "EEFF vs. Nota", "Balance Clasificado", lbl, v_eeff, "#N15.1 — Otros activos financieros (Col: Corrientes)", v_n, ok_n)

            lbl, v_eeff, ex = _extract_stmt_line(df_bal, periodo, ["Otros activos financieros, no corrientes"])
            if ex:
                v_n, ok_n = NoteTableExtractor.extract_value(notes_map, "#N15", 1, periodo, ["Total otros activos financieros", "Total", "Totales"], col_subheader="No corrientes")
                _add_check(15.2, "#N15.1", "EEFF vs. Nota", "Balance Clasificado", lbl, v_eeff, "#N15.1 — Otros activos financieros (Col: No corrientes)", v_n, ok_n)

            lbl, v_eeff, ex = _extract_stmt_line(df_bal, periodo, ["Otros pasivos financieros corrientes", "Otros pasivos financieros, corrientes"])
            if ex:
                v_n, ok_n = NoteTableExtractor.extract_value(notes_map, "#N15", 2, periodo, ["Total otros pasivos financieros", "Total", "Totales"], col_subheader="Corrientes")
                _add_check(15.3, "#N15.2", "EEFF vs. Nota", "Balance Clasificado", lbl, v_eeff, "#N15.2 — Otros pasivos financieros (Col: Corrientes)", v_n, ok_n)

            lbl, v_eeff, ex = _extract_stmt_line(df_bal, periodo, ["Otros pasivos financieros, no corriente", "Otros pasivos financieros, no corrientes"])
            if ex:
                v_n, ok_n = NoteTableExtractor.extract_value(notes_map, "#N15", 2, periodo, ["Total otros pasivos financieros", "Total", "Totales"], col_subheader="No corrientes")
                _add_check(15.4, "#N15.2", "EEFF vs. Nota", "Balance Clasificado", lbl, v_eeff, "#N15.2 — Otros pasivos financieros (Col: No corrientes)", v_n, ok_n)

            # #N16 Pasivos por derechos de uso (Corriente y No corriente en #N16.1)
            lbl, v_eeff, ex = _extract_stmt_line(df_bal, periodo, ["Pasivos por derechos de uso, corrientes"])
            if ex:
                v_n, ok_n = NoteTableExtractor.extract_value(notes_map, "#N16", 1, periodo, ["Total", "Totales"], col_subheader="Corriente")
                _add_check(16.1, "#N16.1", "EEFF vs. Nota", "Balance Clasificado", lbl, v_eeff, "#N16.1 — Pasivos por derechos de uso (Col: Corriente)", v_n, ok_n)

            lbl, v_eeff, ex = _extract_stmt_line(df_bal, periodo, ["Pasivos por derechos de uso, no corrientes"])
            if ex:
                v_n, ok_n = NoteTableExtractor.extract_value(notes_map, "#N16", 1, periodo, ["Total", "Totales"], col_subheader="No corriente")
                _add_check(16.2, "#N16.1", "EEFF vs. Nota", "Balance Clasificado", lbl, v_eeff, "#N16.1 — Pasivos por derechos de uso (Col: No corriente)", v_n, ok_n)

            # #N17 Cuentas comerciales y otras cuentas por pagar (#N17.1 Corriente, #N17.2 No Corriente, #N17.3 Monedas)
            lbl_c, v_cxp_c, ex_c = _extract_stmt_line(df_bal, periodo, ["Cuentas comerciales y otras cuentas por pagar, corrientes"])
            v_n17_1, ok_n17_1 = NoteTableExtractor.extract_value(notes_map, "#N17", 1, periodo, ["Total", "Totales"])
            if ex_c:
                _add_check(17.1, "#N17.1", "EEFF vs. Nota", "Balance Clasificado", lbl_c, v_cxp_c, "#N17.1 — Cuentas comerciales por pagar (Corriente)", v_n17_1, ok_n17_1)

            lbl_nc, v_cxp_nc, ex_nc = _extract_stmt_line(df_bal, periodo, ["Cuentas comerciales y otras cuentas por pagar, no corrientes"])
            v_n17_2, ok_n17_2 = NoteTableExtractor.extract_value(notes_map, "#N17", 2, periodo, ["Total", "Totales"])
            if ex_nc:
                _add_check(17.2, "#N17.2", "EEFF vs. Nota", "Balance Clasificado", lbl_nc, v_cxp_nc, "#N17.2 — Cuentas comerciales por pagar (No corriente)", v_n17_2, ok_n17_2)

            v_n17_3, ok_n17_3 = NoteTableExtractor.extract_value(notes_map, "#N17", 3, periodo, ["Totales", "Total"])
            if ok_n17_3 and (ex_c or ex_nc):
                suma_cxp = (v_n17_1 if ok_n17_1 else v_cxp_c) + (v_n17_2 if ok_n17_2 else v_cxp_nc)
                _add_check(17.3, "#N17.3", "Cuadratura Interna Nota", "Balance Clasificado", "#N17.1 + #N17.2 — Total CxP (Corr. + No Corr.)", suma_cxp, "#N17.3 — Cuentas comerciales por pagar (Por moneda)", v_n17_3, ok_n17_3)

            # #N18 Provisiones (#N18.1 Beneficios empleados y #N18.2 Otras provisiones no corrientes)
            lbl, v_eeff, ex = _extract_stmt_line(df_bal, periodo, ["Provisiones por beneficios a los empleados"])
            if ex:
                v_n, ok_n = NoteTableExtractor.extract_value(notes_map, "#N18", 1, periodo, ["Total", "Totales"])
                _add_check(18.1, "#N18.1", "EEFF vs. Nota", "Balance Clasificado", lbl, v_eeff, "#N18.1 — Provisiones por beneficios a los empleados", v_n, ok_n)

            lbl, v_eeff, ex = _extract_stmt_line(df_bal, periodo, ["Otras provisiones, no corrientes"])
            if ex:
                v_n, ok_n = NoteTableExtractor.extract_value(notes_map, "#N18", 2, periodo, ["Total", "Totales"])
                _add_check(18.2, "#N18.2", "EEFF vs. Nota", "Balance Clasificado", lbl, v_eeff, "#N18.2 — Otras provisiones, no corrientes", v_n, ok_n)

            # #N20 / #PATRIMONIO — Rubros de Patrimonio
            lbl, v_cap, ex_cap = _extract_stmt_line(df_bal, periodo, ["Capital emitido"])
            if ex_cap and pat_vals["exists"]:
                _add_check(20.1, "#PATRIMONIO", "Cruce entre EEFF", "Balance Clasificado", lbl, v_cap, "#PATRIMONIO — Saldo final 'Capital emitido'", pat_vals["capital"], True)

            lbl, v_res, ex_res = _extract_stmt_line(df_bal, periodo, ["Otras reservas"])
            if ex_res:
                if pat_vals["exists"]:
                    _add_check(20.2, "#PATRIMONIO", "Cruce entre EEFF", "Balance Clasificado", lbl, v_res, "#PATRIMONIO — Saldo final 'Otras reservas'", pat_vals["reservas"], True)
                v_n20_4, ok_n20_4 = NoteTableExtractor.extract_value(notes_map, "#N20", 4, periodo, ["Otras reservas", "Saldo final", "Total"])
                _add_check(20.3, "#N20.4", "EEFF vs. Nota", "Balance Clasificado", lbl, v_res, "#N20.4 — Patrimonio (Otras reservas)", v_n20_4, ok_n20_4)

            lbl, v_acum, ex_acum = _extract_stmt_line(df_bal, periodo, ["Resultados acumulados", "Ganancias (pérdidas) acumuladas"])
            if ex_acum and pat_vals["exists"]:
                _add_check(20.4, "#PATRIMONIO", "Cruce entre EEFF", "Balance Clasificado", lbl, v_acum, "#PATRIMONIO — Saldo final 'Ganancias (pérdidas) acumuladas'", pat_vals["acumulados"], True)

            if has_bal_ef and has_efe_cash:
                _add_check(20.5, "#EFE", "Cruce entre EEFF", "Balance Clasificado", "Efectivo y efectivo equivalente", bal_efectivo, "#EFE — Saldo final de efectivo y equivalentes", efe_cash, True)

        # =========================================================================
        # B. ESTADO DE RESULTADOS (#ER) vs. NOTAS (#N13.4, #N21 - #N25)
        # =========================================================================
        if df_er is not None and not df_er.empty:
            # #N21.1 Ingresos de actividades ordinarias (+1)
            lbl, v_eeff, ex = _extract_stmt_line(df_er, periodo, ["Ingresos de actividades ordinarias"])
            if ex:
                v_n, ok_n = NoteTableExtractor.extract_value(notes_map, "#N21", 1, periodo, ["Total ingresos", "Total", "Totales"])
                _add_check(21.1, "#N21.1", "EEFF vs. Nota", "Estado de Resultados", lbl, v_eeff, "#N21.1 — Ingresos de la operación (Total ingresos)", v_n, ok_n)

            # #N21.2 Costo de ventas (en #ER es negativo, en #N21.2 es positivo -> sign_factor = -1)
            lbl, v_eeff, ex = _extract_stmt_line(df_er, periodo, ["Costo de ventas"])
            if ex:
                v_n, ok_n = NoteTableExtractor.extract_value(notes_map, "#N21", 2, periodo, ["Total costos", "Total costo de ventas"])
                _add_check(21.2, "#N21.2", "EEFF vs. Nota", "Estado de Resultados", lbl, v_eeff, "#N21.2 — Costos de la operación (Total costos × -1)", -v_n, ok_n)

            # #N21.2 Depreciación operacional (-1)
            lbl, v_eeff, ex = _extract_stmt_line(df_er, periodo, ["Depreciacion operacional", "Depreciación operacional", "Depreciación y amortización operacional"])
            if ex:
                v_n, ok_n = NoteTableExtractor.extract_value(notes_map, "#N21", 2, periodo, ["Total Depreciacion", "Total depreciación"])
                _add_check(21.3, "#N21.2", "EEFF vs. Nota", "Estado de Resultados", lbl, v_eeff, "#N21.2 — Depreciación operacional (Total Depreciación × -1)", -v_n, ok_n)

            # #N21.3 Arriendo Fibra Óptica IC (Costo en Pacífico / Ingreso en Holdco)
            lbl_fib_cost, v_fib_cost, ex_fib_cost = _extract_stmt_line(df_er, periodo, ["Acceso a infraestructura fibra óptica", "Acceso a infraestructura fibra optica"])
            if ex_fib_cost and v_fib_cost != 0:
                v_n, ok_n = NoteTableExtractor.extract_value(notes_map, "#N21", 3, periodo, ["Costos uso fibra optica", "Costos arriendos IC"])
                _add_check(21.4, "#N21.3", "EEFF vs. Nota", "Estado de Resultados", lbl_fib_cost, v_fib_cost, "#N21.3 — Arriendo de fibra IC (Costos × -1)", -v_n, ok_n)

            lbl_fib_ing, v_fib_ing, ex_fib_ing = _extract_stmt_line(df_er, periodo, ["Ingresos de arriendo fibra optica", "Ingresos de arriendo fibra óptica"])
            if ex_fib_ing and v_fib_ing != 0:
                v_n, ok_n = NoteTableExtractor.extract_value(notes_map, "#N21", 3, periodo, ["Ingresos de arriendo fibra optica", "Ingresos de arriendo fibra óptica"])
                _add_check(21.5, "#N21.3", "EEFF vs. Nota", "Estado de Resultados", lbl_fib_ing, v_fib_ing, "#N21.3 — Arriendo de fibra IC (Ingresos)", v_n, ok_n)

            # #N22.1 Gastos de administración (-1)
            lbl, v_eeff, ex = _extract_stmt_line(df_er, periodo, ["Gastos de administración", "Gastos de administracion"])
            if ex:
                v_n, ok_n = NoteTableExtractor.extract_value(notes_map, "#N22", 1, periodo, ["Total gastos de administración", "Total", "Totales"])
                _add_check(22.1, "#N22.1", "EEFF vs. Nota", "Estado de Resultados", lbl, v_eeff, "#N22.1 — Gastos de administración (Total × -1)", -v_n, ok_n)

            # #N22.2 Depreciación y amortizaciones en administración (-1)
            lbl, v_eeff, ex = _extract_stmt_line(df_er, periodo, ["Depreciación y amortizaciones", "Depreciacion y amortizaciones"])
            if ex:
                v_n, ok_n = NoteTableExtractor.extract_value(notes_map, "#N22", 2, periodo, ["Total", "Totales"], allow_last_row_blank_label=True)
                _add_check(22.2, "#N22.2", "EEFF vs. Nota", "Estado de Resultados", lbl, v_eeff, "#N22.2 — Depreciación y amortizaciones adm. (Total × -1)", -v_n, ok_n)

            # #N23.1 Diferencias de cambio (en #N23 viene con signo contable invertido respecto a #ER -> × -1)
            lbl, v_eeff, ex = _extract_stmt_line(df_er, periodo, ["Diferencias de cambio", "Diferencia de cambio"])
            if ex:
                v_n, ok_n = NoteTableExtractor.extract_value(notes_map, "#N23", 1, periodo, ["Total diferencia de cambio", "Total", "Totales"])
                _add_check(23.1, "#N23.1", "EEFF vs. Nota", "Estado de Resultados", lbl, v_eeff, "#N23.1 — Diferencia de cambios (Total × -1)", -v_n, ok_n)

            # #N23.2 Resultado por unidad de reajuste (signo contable invertido respecto a #ER -> × -1)
            lbl, v_eeff, ex = _extract_stmt_line(df_er, periodo, ["Resultado por unidad de reajuste", "Resultados por unidades de reajuste"])
            if ex:
                v_n, ok_n = NoteTableExtractor.extract_value(notes_map, "#N23", 2, periodo, ["Total resultado por unidad de reajuste", "Total", "Totales"])
                _add_check(23.2, "#N23.2", "EEFF vs. Nota", "Estado de Resultados", lbl, v_eeff, "#N23.2 — Resultado por unidad de reajuste (Total × -1)", -v_n, ok_n)

            # #N24.1 Costos financieros (-1) y #N24.2 Ingresos financieros (+1)
            lbl, v_eeff, ex = _extract_stmt_line(df_er, periodo, ["Costos financieros"])
            if ex:
                v_n, ok_n = NoteTableExtractor.extract_value(notes_map, "#N24", 1, periodo, ["Total, costos financieros", "Total costos financieros", "Total", "Totales"])
                _add_check(24.1, "#N24.1", "EEFF vs. Nota", "Estado de Resultados", lbl, v_eeff, "#N24.1 — Costos financieros (Total × -1)", -v_n, ok_n)

            lbl, v_eeff, ex = _extract_stmt_line(df_er, periodo, ["Ingresos financieros"])
            if ex:
                v_n, ok_n = NoteTableExtractor.extract_value(notes_map, "#N24", 2, periodo, ["Total", "Totales", "Total ingresos financieros"])
                _add_check(24.2, "#N24.2", "EEFF vs. Nota", "Estado de Resultados", lbl, v_eeff, "#N24.2 — Ingresos financieros (Total)", v_n, ok_n)

            # #N25.1 Otros egresos por función (-1) y #N25.2 Otros ingresos por función (+1)
            lbl, v_eeff, ex = _extract_stmt_line(df_er, periodo, ["Otros egresos por función", "Otros Egresos por función", "Otros egresos por funcion"])
            if ex:
                v_n, ok_n = NoteTableExtractor.extract_value(notes_map, "#N25", 1, periodo, ["Total", "Totales"])
                _add_check(25.1, "#N25.1", "EEFF vs. Nota", "Estado de Resultados", lbl, v_eeff, "#N25.1 — Otros egresos por función (Total × -1)", -v_n, ok_n)

            lbl, v_eeff, ex = _extract_stmt_line(df_er, periodo, ["Otros ingresos por función", "Otros ingresos por funcion"])
            if ex:
                v_n, ok_n = NoteTableExtractor.extract_value(notes_map, "#N25", 2, periodo, ["Total", "Totales"])
                _add_check(25.2, "#N25.2", "EEFF vs. Nota", "Estado de Resultados", lbl, v_eeff, "#N25.2 — Otros ingresos por función (Total)", v_n, ok_n)

            # #N13.4 Impuesto a las ganancias (+1)
            lbl, v_eeff, ex = _extract_stmt_line(df_er, periodo, ["Ganancia (Perdida) por impuesto a las ganancias", "Resultado por impuestos a las ganancias"])
            if ex:
                v_n, ok_n = NoteTableExtractor.extract_value(notes_map, "#N13", 4, periodo, ["(Gasto) ingreso por impuesto a las ganancias", "Total"])
                _add_check(13.4, "#N13.4", "EEFF vs. Nota", "Estado de Resultados", lbl, v_eeff, "#N13.4 — Gasto / ingreso por impuesto a las ganancias", v_n, ok_n)

            # Cruce Resultado #ER vs #PATRIMONIO
            if has_res_er and pat_vals["exists"]:
                _add_check(25.9, "#PATRIMONIO", "Cruce entre EEFF", "Estado de Resultados", "Resultado del Ejercicio (#ER)", res_er, "#PATRIMONIO — Ganancia (pérdida) del ejercicio", res_pat, True)

        # =========================================================================
        # C. SOLO CONSOLIDADO — #N26 SEGMENTOS DE OPERACIÓN vs. #BALANCE / #ER
        # =========================================================================
        if is_consolidated and "#N26" in notes_map:
            seg_bal_checks = [
                (26.1, ["Activos corrientes totales"], ["Total activos corrientes por segmento"], "#N26.1 — Segmentos: Total activos corrientes"),
                (26.2, ["Total activos no corrientes"], ["Total activos no corrientes por segmento"], "#N26.1 — Segmentos: Total activos no corrientes"),
                (26.3, ["Pasivos corrientes totales"], ["Total pasivos corrientes por segmento"], "#N26.1 — Segmentos: Total pasivos corrientes"),
                (26.4, ["Total pasivos no corrientes"], ["Total pasivos no corrientes por segmento"], "#N26.1 — Segmentos: Total pasivos no corrientes"),
            ]
            for s_ord, b_aliases, n_rows_lbl, desc in seg_bal_checks:
                lbl_b, v_b, ex_b = _extract_stmt_line(df_bal, periodo, b_aliases)
                if ex_b:
                    v_s, ok_s = NoteTableExtractor.extract_value(notes_map, "#N26", 1, periodo, n_rows_lbl, col_subheader="total_segmento")
                    _add_check(s_ord, "#N26.1", "EEFF vs. Nota", "Balance Clasificado", lbl_b, v_b, desc, v_s, ok_s)

            seg_er_checks = [
                (26.5, ["Ingresos de actividades ordinarias"], ["Ingresos de actividades ordinarias"], "#N26.2 — Segmentos: Ingresos ordinarios"),
                (26.6, ["Ganancia bruta", "Ganancia Bruta"], ["Margen bruto"], "#N26.2 — Segmentos: Margen bruto"),
                (26.7, ["Resultado antes de impuestos", "Ganancia antes de impuesto"], ["Ganancia (Pérdida) antes de impuestos"], "#N26.2 — Segmentos: Resultado antes de impuestos"),
                (26.8, ["(Pérdida) procedente de operaciones continuadas", "Ganancias (Pérdida) del Ejercicio"], ["Ganancia (Pérdida) neta del ejercicio"], "#N26.2 — Segmentos: Resultado neto del ejercicio"),
            ]
            for s_ord, e_aliases, n_rows_lbl, desc in seg_er_checks:
                lbl_e, v_e, ex_e = _extract_stmt_line(df_er, periodo, e_aliases)
                if ex_e:
                    v_s, ok_s = NoteTableExtractor.extract_value(notes_map, "#N26", 2, periodo, n_rows_lbl, col_subheader="total_segmento")
                    _add_check(s_ord, "#N26.2", "EEFF vs. Nota", "Estado de Resultados", lbl_e, v_e, desc, v_s, ok_s)

        def _note_sort_key(r: Dict[str, Any]) -> Tuple[int, int, int, str, int]:
            nota_txt = str(r.get("Nota", "")).strip()
            m = re.search(r"#N(\d+)(?:\.(\d+))?", nota_txt)
            if m:
                return (0, int(m.group(1)), int(m.group(2) or 0), "", int(r.get("_seq", 0)))
            return (1, 999, 0, str(r.get("Código", "")), int(r.get("_seq", 0)))

        rows.sort(key=_note_sort_key)
        res_df = pd.DataFrame(rows)
        if not res_df.empty:
            res_df = res_df.drop(columns=["_seq"]).reset_index(drop=True)
            total_reg = len(res_df)
            total_ok = int((res_df["Diferencia (M$)"] == 0).sum() - ((res_df["Diferencia (M$)"] == 0) & (~res_df["Estado"].str.startswith("✅"))).sum())
            total_err = int(total_reg - total_ok)
            health_checks["total_reglas"] = total_reg
            health_checks["total_cuadrados"] = total_ok
            health_checks["total_descuadres"] = total_err
            health_checks["is_valid"] = (total_err == 0)

        health_checks["elapsed_sec"] = round(time.perf_counter() - t0, 3)
        return res_df, health_checks

    @staticmethod
    def generar_excel_tie_out(df_matrix: pd.DataFrame, health_checks: dict, empresa_o_grupo: str, periodo: str) -> bytes:
        """
        Exporta la Matriz de Validación Directa (en M$, Tolerancia Cero) a un archivo Excel formateado.
        """
        output = io.BytesIO()
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Matriz Validacion M$"

        title_font = Font(name="Arial", size=13, bold=True, color="1F4E78")
        header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        ok_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        error_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
        center_align = Alignment(horizontal="center", vertical="center")
        left_align = Alignment(horizontal="left", vertical="center")
        right_align = Alignment(horizontal="right", vertical="center")
        thin_border = Border(
            left=Side(style="thin", color="D9D9D9"),
            right=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9"),
            bottom=Side(style="thin", color="D9D9D9"),
        )

        ws.cell(row=1, column=1, value="MATRIZ DE VALIDACIÓN Y CUADRATURA DIRECTA (EEFF vs. NOTAS EN M$ — TOLERANCIA CERO)").font = title_font
        ws.cell(
            row=2,
            column=1,
            value=f"Entidad: {empresa_o_grupo} | Período: {periodo} | Origen: {health_checks.get('source_label', '')}",
        ).font = Font(name="Arial", size=10, italic=True)

        status_global = (
            f"✅ 100% CUADRADO ({health_checks.get('total_cuadrados', 0)}/{health_checks.get('total_reglas', 0)} controles en 0 M$)"
            if health_checks.get("is_valid")
            else f"❌ {health_checks.get('total_descuadres', 0)} DESCUADRE(S) DETECTADO(S) DE {health_checks.get('total_reglas', 0)} CONTROLES (TOLERANCIA 0 M$)"
        )
        ws.cell(row=3, column=1, value=f"Estado Global: {status_global}").font = Font(
            name="Arial", size=10, bold=True, color="008000" if health_checks.get("is_valid") else "C00000"
        )

        start_row = 5
        headers = [
            "Código",
            "Rubro EEFF",
            "Nota",
            "Saldo EEFF (M$)",
            "Saldo Nota (M$)",
            "Diferencia (M$)",
            "Estado",
            "Reporte",
            "Nivel",
        ]
        for col_num, h in enumerate(headers, 1):
            cell = ws.cell(row=start_row, column=col_num, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center_align

        current_row = start_row + 1
        if df_matrix is not None and not df_matrix.empty:
            for _, r in df_matrix.iterrows():
                ws.cell(row=current_row, column=1, value=r.get("Código", "")).alignment = center_align
                ws.cell(row=current_row, column=2, value=r.get("Rubro EEFF", "")).alignment = left_align
                ws.cell(row=current_row, column=3, value=r.get("Nota", "")).alignment = left_align

                for c_idx, key_col in (
                    (4, "Saldo EEFF (M$)"),
                    (5, "Saldo Nota (M$)"),
                    (6, "Diferencia (M$)"),
                ):
                    c_num = ws.cell(row=current_row, column=c_idx, value=int(r.get(key_col, 0)))
                    c_num.number_format = "#,##0"
                    c_num.alignment = right_align

                est_str = str(r.get("Estado", ""))
                c7 = ws.cell(row=current_row, column=7, value=est_str)
                c7.alignment = center_align
                c7.fill = ok_fill if est_str.startswith("✅") else error_fill

                ws.cell(row=current_row, column=8, value=r.get("Reporte", "")).alignment = left_align
                ws.cell(row=current_row, column=9, value=r.get("Nivel", "")).alignment = left_align

                for c_i in range(1, 10):
                    ws.cell(row=current_row, column=c_i).border = thin_border
                current_row += 1

        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = min(max(max_len + 4, 14), 52)

        wb.save(output)
        return output.getvalue()

    @staticmethod
    def validar_balance_vs_pl(empresa: str, periodo: str, pl_df_override: pd.DataFrame = None, tb_df_override: pd.DataFrame = None) -> dict:
        """
        Valida que la suma de todas las cuentas del Balance clasificadas como 'Resultados acumulados'
        (cuentas de resultado del ejercicio) coincida exactamente con la suma total de los montos cargados en el Cubo P&L.
        """
        tb_df = tb_df_override if tb_df_override is not None else TrialBalanceDB.get_trial_balance(empresa, periodo)
        has_tb = tb_df is not None and not tb_df.empty

        empresa_dir = os.path.join("data", "empresas", empresa)
        if not os.path.exists(empresa_dir):
            empresa_dir = os.path.join("data", "empresas", f"[GRUPO] {empresa}")
        map_bal_path = os.path.join(empresa_dir, "map_balance.xlsx")
        if not os.path.exists(map_bal_path):
            map_bal_path = "map_balance.xlsx"

        saldo_balance_res_acum = 0.0
        cuentas_balance_res_count = 0
        if has_tb:
            try:
                map_res_accs = set()
                if os.path.exists(map_bal_path):
                    map_bal_df = pd.read_excel(map_bal_path, dtype=str)
                    col_acc_map = map_bal_df.columns[0]
                    col_clas_map = next(
                        (c for c in map_bal_df.columns if "clasifica" in str(c).lower() or "balance" in str(c).lower()),
                        map_bal_df.columns[2] if len(map_bal_df.columns) > 2 else map_bal_df.columns[1],
                    )

                    from src.core.pl_cubo_processor import normalize_text
                    target_keys = [
                        "resultados acumulados",
                        "resultado acumulado",
                        "ganancias acumuladas",
                        "ganancias (perdidas) acumuladas",
                        "ganancias (pérdidas) acumuladas",
                    ]
                    map_res_accs = set(
                        map_bal_df[map_bal_df[col_clas_map].apply(normalize_text).isin(target_keys)][col_acc_map]
                        .dropna()
                        .astype(str)
                        .str.strip()
                    )

                col_acc_tb = next(
                    (c for c in tb_df.columns if "cuenta" in str(c).lower() and "nombre" not in str(c).lower()),
                    "cuenta_id",
                )
                if "saldo_final" in tb_df.columns:
                    col_monto_tb = "saldo_final"
                else:
                    col_monto_tb = next(
                        (c for c in tb_df.columns if "final" in str(c).lower() or "monto" in str(c).lower()),
                        tb_df.columns[-1],
                    )

                tb_df_copy = tb_df.copy()
                tb_df_copy[col_acc_tb] = tb_df_copy[col_acc_tb].astype(str).str.strip()

                mask_res = (
                    (tb_df_copy[col_acc_tb].isin(map_res_accs) & (~tb_df_copy[col_acc_tb].str.startswith(("1", "2"))))
                    | tb_df_copy[col_acc_tb].str.startswith(("3", "4", "5", "6", "7", "8"))
                )
                tb_match = tb_df_copy[mask_res]

                saldo_balance_res_acum = float(pd.to_numeric(tb_match[col_monto_tb], errors="coerce").fillna(0.0).sum())
                cuentas_balance_res_count = len(tb_match)
            except Exception:
                saldo_balance_res_acum = 0.0

        if pl_df_override is not None and not pl_df_override.empty:
            exclude_cols = ["n° de cuenta", "n de cuenta", "cuenta", "cuenta_id", "nombre de la cuenta", "nombre", "descripcion", "rubro"]
            c_acc_pl = next((c for c in pl_df_override.columns if "cuenta" in str(c).lower() and "nombre" not in str(c).lower()), None)
            pl_df_clean = pl_df_override.copy()
            if c_acc_pl:
                pl_df_clean = pl_df_clean[pl_df_clean[c_acc_pl].astype(str).str.strip().str.upper() != "TOTAL"]

            num_cols = [c for c in pl_df_clean.columns if not any(ex in str(c).lower() for ex in exclude_cols)]
            pl_total_sum = 0.0
            from src.ui_pages.cargas_de_datos import parse_numeric_cell
            for col in num_cols:
                pl_total_sum += pl_df_clean[col].apply(parse_numeric_cell).sum()
            has_pl = True
        else:
            pl_total_sum = PlCuboDB.get_pl_cubo_total_sum(empresa, periodo)
            has_pl = PlCuboDB.get_pl_cubo(empresa, periodo) is not None

        diff_direct = round(abs(saldo_balance_res_acum - pl_total_sum), 2)
        diff_opp = round(abs(saldo_balance_res_acum + pl_total_sum), 2)
        diff_abs = round(abs(abs(saldo_balance_res_acum) - abs(pl_total_sum)), 2)
        min_diff = min(diff_direct, diff_opp, diff_abs)

        is_cuadrado = min_diff < 1.0

        return {
            "has_tb": has_tb,
            "has_pl": has_pl,
            "saldo_balance": saldo_balance_res_acum,
            "cuentas_res_count": cuentas_balance_res_count,
            "saldo_pl": pl_total_sum,
            "diferencia": min_diff,
            "is_cuadrado": is_cuadrado,
            "status_str": "✅ 100% CUADRADO" if is_cuadrado else "❌ DESCUADRE DETECTADO",
        }
