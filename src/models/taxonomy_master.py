from sqlalchemy import Column, Integer, String, DateTime
from src.models.database import Base
from datetime import datetime

class TaxonomyMasterRecord(Base):
    __tablename__ = 'taxonomy_master'
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    empresa = Column(String, nullable=False, index=True)
    id_reporte = Column(String, nullable=False, index=True)
    reporte_destino = Column(String, nullable=False)
    nombre_linea_es = Column(String, nullable=False)
    id_nota_asociada = Column(String, nullable=True)
    desglose_nota_es = Column(String, nullable=True)
    nombre_idioma_1 = Column(String, nullable=True)
    nombre_idioma_2 = Column(String, nullable=True)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
