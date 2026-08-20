import os
import sys

# Append dir for absolute imports
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.models.database import engine, Base, init_db
from sqlalchemy import text

def migrate():
    # Asegurar que pl_records_dim se crea si no existe
    init_db()
    
    with engine.begin() as conn:
        # Check if old table exists
        result = conn.execute(text("SELECT name FROM sqlite_master WHERE type='table' AND name='pl_records'")).fetchall()
        if not result:
            print("Old table 'pl_records' does not exist. Nothing to migrate.")
            return

        print("Migrating from pl_records to pl_records_dim...")
        # Clear target just in case
        conn.execute(text("DELETE FROM pl_records_dim"))
        
        # Read old
        old_records = conn.execute(text("SELECT * FROM pl_records")).fetchall()
        
        columns = [
            ("ingresos_ordinarios", "Ingresos de actividades ordinarias"),
            ("costo_ventas", "Costo de ventas"),
            ("depreciacion_operacional", "Depreciación operacional"),
            ("gastos_administracion", "Gastos de administración"),
            ("depreciacion_amortizacion", "Depreciación y amortizaciones"),
            ("otros_ingresos", "Otros ingresos por función"),
            ("otros_egresos", "Otros egresos por función"),
            ("ingresos_financieros", "Ingresos financieros"),
            ("costos_financieros", "Costos financieros"),
            ("diferencias_cambio", "Diferencias de cambio"),
            ("resultado_unidades_reajuste", "Resultados por unidades de reajuste"),
            ("impuestos_ganancias", "Resultado por impuestos a las ganancias")
        ]
        
        insert_query = text("""
            INSERT INTO pl_records_dim (empresa, periodo, cuenta_id, descripcion, rubro, monto)
            VALUES (:emp, :per, :cid, :desc, :rub, :mon)
        """)
        
        count = 0
        for r in old_records:
            emp = r[1]
            per = r[2]
            cid = r[3]
            desc = r[4]
            
            # r[5] corresponds to ingresos_ordinarios onwards if id is 0
            # Let's dynamically fetch by column name using dictionary indexing if possible, 
            # but sqlite fetchall returns tuples. We map index manually.
            # id=0, empresa=1, periodo=2, cuenta_id=3, descripcion=4
            
            val_dict = {
                "ingresos_ordinarios": r[5],
                "costo_ventas": r[6],
                "depreciacion_operacional": r[7],
                "gastos_administracion": r[8],
                "depreciacion_amortizacion": r[9],
                "otros_ingresos": r[10],
                "otros_egresos": r[11],
                "ingresos_financieros": r[12],
                "costos_financieros": r[13],
                "diferencias_cambio": r[14],
                "resultado_unidades_reajuste": r[15],
                "impuestos_ganancias": r[16]
            }
            
            for col_sql, rubro_name in columns:
                val = val_dict[col_sql]
                if val != 0.0:
                    conn.execute(insert_query, {
                        'emp': emp, 'per': per, 'cid': cid, 'desc': desc, 'rub': rubro_name, 'mon': val
                    })
                    count += 1
                    
        print(f"Migrated {count} dimensional values!")
        
        try:
            conn.execute(text("DROP TABLE pl_records"))
            print("Dropped old pl_records table.")
        except Exception as e:
            print("Could not drop old table:", e)

if __name__ == "__main__":
    migrate()
