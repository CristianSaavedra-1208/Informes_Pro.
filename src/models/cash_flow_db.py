from sqlalchemy import Column, Integer, String, Float, Boolean
from src.models.database import Base

class CashFlowAdjustment(Base):
    __tablename__ = 'cash_flow_adjustments'

    id = Column(Integer, primary_key=True, index=True)
    empresa = Column(String(200), index=True, nullable=False)
    periodo = Column(String(7), index=True, nullable=False) # YYYY-MM
    glosa = Column(String(255), nullable=False)
    linea_item = Column(String(255), nullable=False) # Cuenta ID o nombre del rubro
    ingreso_caja = Column(Float, nullable=False, default=0.0)
    egreso_caja = Column(Float, nullable=False, default=0.0)
    es_consolidado = Column(Boolean, nullable=False, default=False)
