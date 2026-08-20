from sqlalchemy import Column, Integer, String, Float, Index
from src.models.database import Base

class PlRecordDim(Base):
    __tablename__ = "pl_records_dim"

    id = Column(Integer, primary_key=True, index=True)
    empresa = Column(String, index=True, nullable=False)
    periodo = Column(String(7), index=True, nullable=False)  # Formato YYYY-MM
    cuenta_id = Column(String, index=True, nullable=False)
    descripcion = Column(String, nullable=True)
    
    rubro = Column(String, index=True, nullable=False)
    monto = Column(Float, default=0.0)

    # Indice para agrupación rápida
    __table_args__ = (
        Index('idx_pldim_empresa_periodo', 'empresa', 'periodo'),
        Index('idx_pldim_rubro', 'rubro')
    )
