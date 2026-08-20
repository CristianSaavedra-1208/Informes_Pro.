from sqlalchemy import Column, Integer, String, Float, UniqueConstraint, Date
from src.models.database import Base

class TrialBalanceRecord(Base):
    __tablename__ = 'trial_balance_records'

    id = Column(Integer, primary_key=True, index=True)
    empresa = Column(String(200), index=True, nullable=False)
    periodo = Column(String(7), index=True, nullable=False) # Format: YYYY-MM
    cuenta_id = Column(String(100), index=True, nullable=False)
    descripcion = Column(String(255), nullable=True)
    saldo_inicial = Column(Float, nullable=False, default=0.0)
    debitos = Column(Float, nullable=False, default=0.0)
    creditos = Column(Float, nullable=False, default=0.0)
    saldo_final = Column(Float, nullable=False, default=0.0)

    __table_args__ = (
        UniqueConstraint('empresa', 'periodo', 'cuenta_id', name='uq_emp_per_cta'),
    )
