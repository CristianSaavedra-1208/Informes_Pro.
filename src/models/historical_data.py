from sqlalchemy import Column, Integer, String, Float, UniqueConstraint, DateTime
from src.models.database import Base
from datetime import datetime

class HistoricalDataRecord(Base):
    """
    Guarda los valores duros finales y congelados tras un Cierre de Año
    o carga histórica (valores pegados).
    """
    __tablename__ = 'historical_data_records'

    id = Column(Integer, primary_key=True, index=True)
    empresa = Column(String(200), index=True, nullable=False)
    periodo = Column(String(7), index=True, nullable=False) # Formato: YYYY-MM
    reporte = Column(String(100), index=True, nullable=False) # ej: 'Balance Clasificado' o 'Nota_Efectivo'
    linea_item = Column(String(255), nullable=False) # ej: 'Efectivo y equivalentes al efectivo'
    monto = Column(Float, nullable=False, default=0.0)
    fecha_congelamiento = Column(DateTime, default=datetime.utcnow)
    
    __table_args__ = (
        UniqueConstraint('empresa', 'periodo', 'reporte', 'linea_item', name='uq_emp_per_rep_lin'),
    )

class HistoricalDetailRecord(Base):
    """
    Guarda el papel de trabajo detallado (cuenta por cuenta) 
    para un año cerrado (Trial Balance Mapeado Congelado).
    """
    __tablename__ = 'historical_detail_records'

    id = Column(Integer, primary_key=True, index=True)
    empresa = Column(String(200), index=True, nullable=False)
    periodo = Column(String(7), index=True, nullable=False)
    cuenta_id = Column(String(100), nullable=False)
    descripcion = Column(String(255), nullable=False)
    saldo_inicial = Column(Float, nullable=False, default=0.0)
    debitos = Column(Float, nullable=False, default=0.0)
    creditos = Column(Float, nullable=False, default=0.0)
    saldo_final = Column(Float, nullable=False, default=0.0)
    clasificacion_balance = Column(String(255), nullable=True)
    clasificacion_pl = Column(String(255), nullable=True)
    id_nota_asociada = Column(String(255), nullable=True)
    fecha_congelamiento = Column(DateTime, default=datetime.utcnow)

