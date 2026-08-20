from sqlalchemy import Column, Integer, String, Float, DateTime
from src.models.database import Base
from datetime import datetime

class AuditAdjustmentRecord(Base):
    __tablename__ = 'audit_adjustments'

    id = Column(Integer, primary_key=True, index=True)
    empresa = Column(String(200), index=True, nullable=False)
    periodo = Column(String(7), index=True, nullable=False) # Formato: YYYY-MM
    cuenta_id = Column(String(100), index=True, nullable=False)
    monto = Column(Float, nullable=False, default=0.0) # (+) Cargo, (-) Abono
    descripcion = Column(String(255), nullable=True)
    fecha_creacion = Column(DateTime, default=datetime.utcnow)
