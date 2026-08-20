# -*- coding: utf-8 -*-
import sys
import io
if sys.stdout.encoding and sys.stdout.encoding.lower() not in ('utf-8', 'utf_8'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

"""
snapshot_generator.py
=====================
Etapa 0 - Genera el snapshot numerico de referencia (baseline).

Estrategia rapida: en lugar de generar el Excel completo (costoso),
capturamos directamente el entity_context (el diccionario de saldos agrupados
que alimenta TODAS las notas). Esto es 50x mas rapido y captura exactamente
los mismos numeros que aparecen en las notas.

Ejecutar UNA SOLA VEZ antes de cualquier optimizacion:
    python tests/perf/snapshot_generator.py

Guarda el resultado en tests/perf/baseline_snapshot.json
"""

import os
import json
import traceback
from datetime import datetime

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.chdir(ROOT_DIR)
sys.path.insert(0, ROOT_DIR)

import pandas as pd

from src.ui_pages.informes_y_notas import load_all_entity_contexts

# ─────────────────────────────────────────────
#  SCOPE
# ─────────────────────────────────────────────
EMPRESAS_INDIVIDUALES = [
    "Pacifico SpA",
    "DB Holdco Terra SpA",
]

GRUPOS_CONSOLIDADOS = [
    "[GRUPO] Consolidado DB Terra Holdco",
]

PERIODOS_DISPONIBLES = [
    "2025-03", "2025-06", "2025-12",
    "2026-03", "2026-04", "2026-05", "2026-06",
]

PERIODO_COMP_MAP = {
    "2025-03": "Ninguno",
    "2025-06": "2025-03",
    "2025-12": "2025-06",
    "2026-03": "2025-12",
    "2026-04": "2026-03",
    "2026-05": "2026-04",
    "2026-06": "2026-05",
}

GLOBAL_TEMPLATE = "Plantilla de notas_v1.xlsx"
OUTPUT_FILE = os.path.join("tests", "perf", "baseline_snapshot.json")
TOLERANCE = 0.01


def serialize_context(ctx: dict) -> dict:
    """
    Serializa el entity_context a un dict JSON-serializable.
    Captura los valores numericos 'val', 'inicial', 'debitos', 'creditos'
    para cada clave en nota1, nota2, y pl (por rubro).
    """
    result = {}

    # nota1 y nota2: {label: {val, inicial, debitos, creditos}}
    for section in ("nota1", "nota2"):
        section_data = ctx.get(section, {})
        result[section] = {}
        for label, data in section_data.items():
            if isinstance(data, dict):
                result[section][label] = {
                    "val": round(float(data.get("val", 0.0)), 4),
                    "inicial": round(float(data.get("inicial", 0.0)), 4),
                    "debitos": round(float(data.get("debitos", 0.0)), 4),
                    "creditos": round(float(data.get("creditos", 0.0)), 4),
                }

    # pl: {rubro: {label: {val, inicial, debitos, creditos}}}
    pl_data = ctx.get("pl", {})
    result["pl"] = {}
    for rubro, rubro_data in pl_data.items():
        if not isinstance(rubro_data, dict):
            continue
        result["pl"][rubro] = {}
        for label, data in rubro_data.items():
            if isinstance(data, dict):
                result["pl"][rubro][label] = {
                    "val": round(float(data.get("val", 0.0)), 4),
                    "inicial": round(float(data.get("inicial", 0.0)), 4),
                    "debitos": round(float(data.get("debitos", 0.0)), 4),
                    "creditos": round(float(data.get("creditos", 0.0)), 4),
                }

    return result


def compute_context_checksum(ctx_serialized: dict) -> float:
    """Suma de todos los valores 'val' en el contexto — checksum rapido."""
    total = 0.0
    for section in ("nota1", "nota2"):
        for data in ctx_serialized.get(section, {}).values():
            total += data.get("val", 0.0)
    for rubro_data in ctx_serialized.get("pl", {}).values():
        for data in rubro_data.values():
            total += data.get("val", 0.0)
    return round(total, 4)


def run_snapshot():
    print("=" * 60)
    print("  SNAPSHOT GENERATOR - Informes Pro")
    print(f"  Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)

    if not os.path.exists(GLOBAL_TEMPLATE):
        print(f"[ERROR] No se encontro la plantilla: {GLOBAL_TEMPLATE}")
        sys.exit(1)

    # Cargar mapeos desde la empresa "referencia" (Pacifico SpA tiene los mapeos maestros)
    GLOBAL_EMPRESA_PATH = os.path.join("data", "empresas", "Pacifico SpA")
    map_balance_path = os.path.join(GLOBAL_EMPRESA_PATH, "map_balance.xlsx")
    map_pl_path      = os.path.join(GLOBAL_EMPRESA_PATH, "map_pl.xlsx")

    try:
        map_balance_df = pd.read_excel(map_balance_path, dtype=str)
        print(f"  [OK] Mapeo Balance cargado: {len(map_balance_df)} filas")
    except Exception as e:
        print(f"  [WARN] No se pudo cargar Mapeo Balance: {e}")
        map_balance_df = None

    try:
        map_pl_df = pd.read_excel(map_pl_path, dtype=str)
        print(f"  [OK] Mapeo PL cargado: {len(map_pl_df)} filas")
    except Exception as e:
        print(f"  [WARN] No se pudo cargar Mapeo PL: {e}")
        map_pl_df = None

    snapshot = {
        "_meta": {
            "generated_at": datetime.now().isoformat(),
            "template": GLOBAL_TEMPLATE,
            "tolerance": TOLERANCE,
            "scope_empresas": EMPRESAS_INDIVIDUALES,
            "scope_grupos": GRUPOS_CONSOLIDADOS,
            "periodos": PERIODOS_DISPONIBLES,
            "strategy": "entity_context"
        },
        "contexts": {}
    }

    all_entities = EMPRESAS_INDIVIDUALES + GRUPOS_CONSOLIDADOS
    total_ok = 0
    total_errors = 0

    for empresa in all_entities:
        snapshot["contexts"][empresa] = {}
        print(f"\n>> Empresa: {empresa}")

        for periodo_actual in PERIODOS_DISPONIBLES:
            periodo_comp = PERIODO_COMP_MAP.get(periodo_actual, "Ninguno")
            print(f"   Periodo: {periodo_actual} | Comp: {periodo_comp} ", end="", flush=True)

            try:
                entity_contexts = load_all_entity_contexts(
                    active_entity=empresa,
                    periodo_actual=periodo_actual,
                    periodo_comp=periodo_comp,
                    map_balance_df=map_balance_df,
                    map_pl_df=map_pl_df,
                )

                period_snapshot = {}

                # Capturar contexto de cada entidad encontrada
                for entity_name, ctx_pair in entity_contexts.items():
                    ctx_actual = serialize_context(ctx_pair.get("actual", {}))
                    ctx_comp = serialize_context(ctx_pair.get("comp", {}))
                    checksum_actual = compute_context_checksum(ctx_actual)
                    checksum_comp = compute_context_checksum(ctx_comp)
                    period_snapshot[entity_name] = {
                        "actual": ctx_actual,
                        "comp": ctx_comp,
                        "checksum_actual": checksum_actual,
                        "checksum_comp": checksum_comp,
                    }

                snapshot["contexts"][empresa][periodo_actual] = period_snapshot
                total_entries = sum(
                    len(v.get("nota1", {})) + len(v.get("nota2", {})) +
                    sum(len(rv) for rv in v.get("pl", {}).values())
                    for pair in period_snapshot.values()
                    for v in [pair["actual"], pair["comp"]]
                )
                print(f"-> [OK] {total_entries} entradas capturadas")
                total_ok += 1

            except Exception as e:
                print(f"-> [ERROR] {e}")
                snapshot["contexts"][empresa][periodo_actual] = {"error": str(e)}
                total_errors += 1

    # Guardar
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(snapshot, f, ensure_ascii=False, indent=2)

    print("\n" + "=" * 60)
    print("  SNAPSHOT COMPLETADO")
    print(f"  Periodos procesados OK: {total_ok}")
    print(f"  Errores: {total_errors}")
    print(f"  Guardado en: {OUTPUT_FILE}")
    print(f"  Tamano del archivo: {os.path.getsize(OUTPUT_FILE) / 1024:.1f} KB")
    print("=" * 60)


if __name__ == "__main__":
    run_snapshot()
