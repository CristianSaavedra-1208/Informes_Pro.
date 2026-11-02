from sqlalchemy import Column, Integer, String, Float, DateTime
from src.models.database import Base
from datetime import datetime

class AuditAdjustmentRecord(Base):
    __tablename__ = 'audit_adjustments'

    id = Column(Integer, primary_key=True, index=True)
    empresa = Column(String(200), index=True, nullable=False)
    periodo = Column(String(20), index=True, nullable=False)
    asiento_codigo = Column(String(50), index=True, nullable=False)
    num_linea = Column(Integer, default=1)
    cuenta_codigo = Column(String(100), index=True, nullable=False)
    cuenta_id = Column(String(100), index=True, nullable=True)
    nombre_cuenta = Column(String(255), nullable=True)
    linea_item = Column(String(255), nullable=True)
    linea_nota = Column(String(100), nullable=True)
    debe = Column(Float, nullable=False, default=0.0)
    haber = Column(Float, nullable=False, default=0.0)
    monto = Column(Float, nullable=False, default=0.0)
    glosa = Column(String(500), nullable=False)
    descripcion = Column(String(255), nullable=True)
    fecha_creacion = Column(DateTime, default=datetime.utcnow)

