import sys
import os
sys.path.append(os.getcwd())
from src.models.database import SessionLocal
from src.models.pl_record import PlRecord

db = SessionLocal()
records = db.query(PlRecord.empresa, PlRecord.periodo).distinct().all()
print("All PlRecords distinct by (empresa, periodo):", records)
db.close()
