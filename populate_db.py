import os
import pandas as pd
from src.models.database import SessionLocal, init_db
from src.models.taxonomy_master import TaxonomyMasterRecord

init_db()
db = SessionLocal()

empresa_name = "Pacifico SpA"
empresa_path = r"c:\Users\cfsaa\OneDrive\Desktop\Informes_Pro\data\empresas\Pacifico SpA"
tax_path = os.path.join(empresa_path, "taxonomy_master_auto.xlsx")

if os.path.exists(tax_path):
    df = pd.read_excel(tax_path)
    count = 0
    for _, row in df.iterrows():
        id_rep = str(row.get("ID_Reporte")).strip()
        if not id_rep or str(id_rep) == "nan": continue
        
        existing = db.query(TaxonomyMasterRecord).filter_by(empresa=empresa_name, id_reporte=id_rep).first()
        if not existing:
            existing = TaxonomyMasterRecord(empresa=empresa_name, id_reporte=id_rep)
            db.add(existing)
            
        existing.reporte_destino = str(row.get("Reporte_Destino", ""))
        existing.nombre_linea_es = str(row.get("Nombre_Linea_ES", ""))
        
        n_asoc = str(row.get("ID_Nota_Asociada", ""))
        if n_asoc != "nan": existing.id_nota_asociada = n_asoc
        
        d_nota = str(row.get("Desglose_Nota_ES", ""))
        if d_nota != "nan": existing.desglose_nota_es = d_nota
        
        db.commit()
        count += 1
        
    print(f"Injected {count} records into SQL DB for {empresa_name}")
db.close()
