# -*- coding: utf-8 -*-
import sys
import io
if sys.stdout.encoding and sys.stdout.encoding.lower() not in ('utf-8', 'utf_8'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

"""
validate_regression.py
======================
Compara el entity_context actual contra el baseline_snapshot.json.

Uso:
    python tests/perf/validate_regression.py
    python tests/perf/validate_regression.py --verbose      # muestra diferencias valor a valor
    python tests/perf/validate_regression.py --empresa "Pacifico SpA"
    python tests/perf/validate_regression.py --periodo 2026-06

Retorna:
    Exit code 0 -> sin regresiones
    Exit code 1 -> regresion detectada (numeros cambiaron)
"""

import os
import sys
import json
import argparse
from datetime import datetime

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.chdir(ROOT_DIR)
sys.path.insert(0, ROOT_DIR)

import pandas as pd
from src.ui_pages.informes_y_notas import load_all_entity_contexts

BASELINE_FILE = os.path.join("tests", "perf", "baseline_snapshot.json")
GLOBAL_EMPRESA_PATH = os.path.join("data", "empresas", "Pacifico SpA")

PERIODO_COMP_MAP = {
    "2025-03": "Ninguno",
    "2025-06": "2025-03",
    "2025-12": "2025-06",
    "2026-03": "2025-12",
    "2026-04": "2026-03",
    "2026-05": "2026-04",
    "2026-06": "2026-05",
}


def serialize_context(ctx: dict) -> dict:
    """Mismo serializador que snapshot_generator.py — debe mantenerse identico."""
    result = {}
    for section in ("nota1", "nota2"):
        section_data = ctx.get(section, {})
        result[section] = {}
        for label, data in section_data.items():
            if isinstance(data, dict):
                result[section][label] = {
                    "val":      round(float(data.get("val", 0.0)), 4),
                    "inicial":  round(float(data.get("inicial", 0.0)), 4),
                    "debitos":  round(float(data.get("debitos", 0.0)), 4),
                    "creditos": round(float(data.get("creditos", 0.0)), 4),
                }
    pl_data = ctx.get("pl", {})
    result["pl"] = {}
    for rubro, rubro_data in pl_data.items():
        if not isinstance(rubro_data, dict):
            continue
        result["pl"][rubro] = {}
        for label, data in rubro_data.items():
            if isinstance(data, dict):
                result["pl"][rubro][label] = {
                    "val":      round(float(data.get("val", 0.0)), 4),
                    "inicial":  round(float(data.get("inicial", 0.0)), 4),
                    "debitos":  round(float(data.get("debitos", 0.0)), 4),
                    "creditos": round(float(data.get("creditos", 0.0)), 4),
                }
    return result


def flat_values(ctx_serial: dict) -> dict:
    """
    Aplana el contexto serializado a un dict {path: value} para comparacion directa.
    path = "nota1/label/val", "pl/rubro/label/creditos", etc.
    """
    out = {}
    for section in ("nota1", "nota2"):
        for label, vals in ctx_serial.get(section, {}).items():
            for field, v in vals.items():
                out[f"{section}/{label}/{field}"] = v
    for rubro, rubro_data in ctx_serial.get("pl", {}).items():
        for label, vals in rubro_data.items():
            for field, v in vals.items():
                out[f"pl/{rubro}/{label}/{field}"] = v
    return out


def compare_contexts(base_ctx: dict, curr_ctx: dict, tolerance: float, verbose: bool) -> list:
    """Compara dos contextos serializados. Retorna lista de diferencias."""
    diffs = []
    base_flat = flat_values(base_ctx)
    curr_flat = flat_values(curr_ctx)

    # Claves que desaparecieron o son nuevas
    base_keys = set(base_flat.keys())
    curr_keys = set(curr_flat.keys())

    if verbose:
        for k in sorted(base_keys - curr_keys):
            diffs.append(f"  DESAPARECIO: {k} (era {base_flat[k]:,.4f})")
        for k in sorted(curr_keys - base_keys):
            diffs.append(f"  NUEVO: {k} = {curr_flat[k]:,.4f}")

    # Comparar valores en comun
    for k in sorted(base_keys & curr_keys):
        bv = base_flat[k]
        cv = curr_flat[k]
        if abs(cv - bv) > tolerance:
            diffs.append(
                f"  CAMBIO [{k}]: {bv:,.4f} -> {cv:,.4f} (delta={cv - bv:+,.4f})"
            )
    return diffs


def run_validation(filter_empresa=None, filter_periodo=None, verbose=False):
    print("=" * 65)
    print("  VALIDATE REGRESSION - Informes Pro")
    print(f"  Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 65)

    if not os.path.exists(BASELINE_FILE):
        print(f"[ERROR] No existe el baseline: {BASELINE_FILE}")
        print("  Ejecuta primero: python tests/perf/snapshot_generator.py")
        sys.exit(1)

    with open(BASELINE_FILE, "r", encoding="utf-8") as f:
        baseline = json.load(f)

    meta = baseline.get("_meta", {})
    tolerance = meta.get("tolerance", 0.01)
    print(f"  Baseline generado: {meta.get('generated_at', '?')}")
    print(f"  Tolerancia:        +/-{tolerance}")
    print(f"  Filtro empresa:    {filter_empresa or 'Todas'}")
    print(f"  Filtro periodo:    {filter_periodo or 'Todos'}")
    print()

    # Cargar mapeos
    map_balance_df = None
    map_pl_df = None
    try:
        map_balance_df = pd.read_excel(
            os.path.join(GLOBAL_EMPRESA_PATH, "map_balance.xlsx"), dtype=str
        )
    except Exception as e:
        print(f"  [WARN] No se pudo cargar Mapeo Balance: {e}")
    try:
        map_pl_df = pd.read_excel(
            os.path.join(GLOBAL_EMPRESA_PATH, "map_pl.xlsx"), dtype=str
        )
    except Exception as e:
        print(f"  [WARN] No se pudo cargar Mapeo PL: {e}")

    baseline_contexts = baseline.get("contexts", {})
    total_compared = 0
    total_diffs = []
    error_count = 0

    for empresa, periodos_data in baseline_contexts.items():
        if filter_empresa and filter_empresa.lower() not in empresa.lower():
            continue

        for periodo_actual, base_period_data in periodos_data.items():
            if filter_periodo and filter_periodo != periodo_actual:
                continue
            if "error" in base_period_data:
                continue

            periodo_comp = PERIODO_COMP_MAP.get(periodo_actual, "Ninguno")

            try:
                entity_contexts = load_all_entity_contexts(
                    active_entity=empresa,
                    periodo_actual=periodo_actual,
                    periodo_comp=periodo_comp,
                    map_balance_df=map_balance_df,
                    map_pl_df=map_pl_df,
                )
            except Exception as e:
                error_count += 1
                print(f"  [ERROR] Cargando contexto {empresa}/{periodo_actual}: {e}")
                continue

            period_has_diff = False
            for entity_name, base_pair in base_period_data.items():
                if entity_name not in entity_contexts:
                    if verbose:
                        total_diffs.append((empresa, periodo_actual, entity_name, [f"  Entidad '{entity_name}' desaparecio del contexto"]))
                    continue

                curr_pair = entity_contexts[entity_name]

                for slot in ("actual", "comp"):
                    base_ctx = base_pair.get(slot, {})
                    curr_raw = curr_pair.get(slot, {})
                    curr_ctx = serialize_context(curr_raw)

                    n_vals = len(flat_values(base_ctx))
                    total_compared += n_vals

                    diffs = compare_contexts(base_ctx, curr_ctx, tolerance, verbose)
                    if diffs:
                        period_has_diff = True
                        total_diffs.append((empresa, periodo_actual, f"{entity_name}/{slot}", diffs))

            status = "[ERROR]" if period_has_diff else "[OK]"
            n_entities = len(base_period_data)
            print(f"  {status}  {empresa[:35]:<35} | {periodo_actual} | {n_entities} entidades")

    # ─── Resumen final ───────────────────────────────────────────────────────
    print()
    print("=" * 65)
    if not total_diffs and error_count == 0:
        print(f"  [PASS] VALIDACION OK -- {total_compared:,} valores comparados, 0 diferencias")
        print("=" * 65)
        sys.exit(0)
    else:
        print(f"  [FAIL] REGRESION DETECTADA -- {len(total_diffs)} contexto(s) con diferencias")
        print(f"  Total valores comparados: {total_compared:,}")
        print(f"  Errores de ejecucion:     {error_count}")
        print()
        for empresa, periodo, ctx_label, diffs in total_diffs:
            print(f"  ── {empresa} | {periodo} | {ctx_label} ──")
            for d in diffs[:30]:
                print(d)
            if len(diffs) > 30:
                print(f"  ... y {len(diffs) - 30} diferencias mas")
            print()
        print("=" * 65)
        print("  Para revertir: python tests/perf/autocorrect.py")
        print("=" * 65)
        sys.exit(1)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Valida regresion numerica vs baseline")
    parser.add_argument("--verbose",  action="store_true", help="Mostrar diferencias valor a valor")
    parser.add_argument("--empresa",  type=str, default=None, help="Filtrar por empresa")
    parser.add_argument("--periodo",  type=str, default=None, help="Filtrar por periodo, ej: 2026-06")
    args = parser.parse_args()
    run_validation(filter_empresa=args.empresa, filter_periodo=args.periodo, verbose=args.verbose)
