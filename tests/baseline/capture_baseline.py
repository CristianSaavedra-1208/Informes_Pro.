import os
import sys
import json
import hashlib
import sqlite3

# Ensure Informes_Pro is on sys.path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

DB_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "data", "informes_pro.db")
OUTPUT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "baseline_hashes.json")

CRITICAL_TABLES = [
    "pl_records_dim",
    "trial_balance_records",
    "historical_data_records",
    "historical_detail_records",
    "taxonomy_master"
]

def hash_table(conn, table_name):
    cursor = conn.cursor()
    # Check if table exists
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (table_name,))
    if not cursor.fetchone():
        return None

    # Get column info ordered
    cursor.execute(f"PRAGMA table_info({table_name})")
    cols = [col[1] for col in cursor.fetchall()]
    order_cols = ", ".join([f'"{c}"' for c in cols])

    cursor.execute(f"SELECT * FROM {table_name} ORDER BY {order_cols}")
    rows = cursor.fetchall()

    h = hashlib.sha256()
    row_count = len(rows)
    for r in rows:
        row_str = "|".join("" if v is None else str(v) for v in r)
        h.update(row_str.encode("utf-8"))

    return {
        "count": row_count,
        "sha256": h.hexdigest(),
        "columns": cols
    }

def capture_aggregates(conn):
    cursor = conn.cursor()
    aggregates = {}

    # pl_records_dim
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='pl_records_dim'")
    if cursor.fetchone():
        cursor.execute("""
            SELECT empresa, periodo, COUNT(*), ROUND(COALESCE(SUM(monto), 0), 4)
            FROM pl_records_dim
            GROUP BY empresa, periodo
            ORDER BY empresa, periodo
        """)
        aggregates["pl_records_dim"] = [
            {"empresa": r[0], "periodo": r[1], "count": r[2], "sum_monto": r[3]}
            for r in cursor.fetchall()
        ]

    # trial_balance_records
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='trial_balance_records'")
    if cursor.fetchone():
        cursor.execute("""
            SELECT empresa, periodo, COUNT(*), ROUND(COALESCE(SUM(saldo_final), 0), 4)
            FROM trial_balance_records
            GROUP BY empresa, periodo
            ORDER BY empresa, periodo
        """)
        aggregates["trial_balance_records"] = [
            {"empresa": r[0], "periodo": r[1], "count": r[2], "sum_saldo_final": r[3]}
            for r in cursor.fetchall()
        ]

    # historical_data_records
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='historical_data_records'")
    if cursor.fetchone():
        cursor.execute("""
            SELECT empresa, periodo, COUNT(*), ROUND(COALESCE(SUM(monto), 0), 4)
            FROM historical_data_records
            GROUP BY empresa, periodo
            ORDER BY empresa, periodo
        """)
        aggregates["historical_data_records"] = [
            {"empresa": r[0], "periodo": r[1], "count": r[2], "sum_monto": r[3]}
            for r in cursor.fetchall()
        ]

    # historical_detail_records
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='historical_detail_records'")
    if cursor.fetchone():
        cursor.execute("""
            SELECT empresa, periodo, COUNT(*), ROUND(COALESCE(SUM(saldo_final), 0), 4)
            FROM historical_detail_records
            GROUP BY empresa, periodo
            ORDER BY empresa, periodo
        """)
        aggregates["historical_detail_records"] = [
            {"empresa": r[0], "periodo": r[1], "count": r[2], "sum_saldo_final": r[3]}
            for r in cursor.fetchall()
        ]

    # taxonomy_master
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='taxonomy_master'")
    if cursor.fetchone():
        cursor.execute("""
            SELECT empresa, id_reporte, COUNT(*)
            FROM taxonomy_master
            GROUP BY empresa, id_reporte
            ORDER BY empresa, id_reporte
        """)
        aggregates["taxonomy_master"] = [
            {"empresa": r[0], "id_reporte": r[1], "count": r[2]}
            for r in cursor.fetchall()
        ]

    return aggregates

def main():
    if not os.path.exists(DB_PATH):
        print(f"Error: Database not found at {DB_PATH}")
        sys.exit(1)

    conn = sqlite3.connect(DB_PATH)
    try:
        tables_snapshot = {}
        for tbl in CRITICAL_TABLES:
            tables_snapshot[tbl] = hash_table(conn, tbl)

        aggregates = capture_aggregates(conn)

        baseline_data = {
            "tables": tables_snapshot,
            "aggregates": aggregates
        }

        os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
        with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
            json.dump(baseline_data, f, indent=2, ensure_ascii=False)

        print(f"[OK] Baseline snapshot successfully saved to {OUTPUT_PATH}")
        for tbl, data in tables_snapshot.items():
            if data:
                print(f"  - {tbl}: {data['count']} rows, SHA256={data['sha256'][:12]}...")
            else:
                print(f"  - {tbl}: (table not found)")
    finally:
        conn.close()

if __name__ == "__main__":
    main()
