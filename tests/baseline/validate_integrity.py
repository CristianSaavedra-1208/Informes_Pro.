import os
import sys
import json
import hashlib
import sqlite3

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

DB_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "data", "informes_pro.db")
BASELINE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "baseline_hashes.json")

from capture_baseline import hash_table, capture_aggregates, CRITICAL_TABLES

def validate():
    if not os.path.exists(BASELINE_PATH):
        print(f"[ERROR] Baseline file not found at {BASELINE_PATH}")
        return False

    with open(BASELINE_PATH, "r", encoding="utf-8") as f:
        baseline = json.load(f)

    conn = sqlite3.connect(DB_PATH)
    discrepancies = []
    try:
        # 1. Compare table hashes
        for tbl in CRITICAL_TABLES:
            expected = baseline["tables"].get(tbl)
            current = hash_table(conn, tbl)
            if expected is None and current is None:
                continue
            if expected is None or current is None:
                discrepancies.append(f"Table existence mismatch for {tbl}: expected={expected is not None}, current={current is not None}")
                continue

            if expected["count"] != current["count"]:
                discrepancies.append(f"Row count mismatch in {tbl}: expected={expected['count']}, current={current['count']}")
            if expected["sha256"] != current["sha256"]:
                discrepancies.append(f"SHA-256 hash mismatch in {tbl}: expected={expected['sha256']}, current={current['sha256']}")

        # 2. Compare aggregates
        curr_aggregates = capture_aggregates(conn)
        for tbl, expected_agg in baseline.get("aggregates", {}).items():
            curr_agg = curr_aggregates.get(tbl, [])
            if expected_agg != curr_agg:
                discrepancies.append(f"Aggregates mismatch in {tbl}:\n  Expected: {expected_agg}\n  Current:  {curr_agg}")

    finally:
        conn.close()

    if discrepancies:
        print("[FAIL] Integridad de datos violada! Se encontraron discrepancias:")
        for d in discrepancies:
            print(f"  ❌ {d}")
        return False
    else:
        print("[SUCCESS] Integridad de datos 100% preservada. Todas las tablas críticas coinciden con el baseline.")
        return True

if __name__ == "__main__":
    success = validate()
    sys.exit(0 if success else 1)
