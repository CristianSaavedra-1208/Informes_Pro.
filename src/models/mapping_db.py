from sqlalchemy import Column, Integer, String, ForeignKey, UniqueConstraint
from src.models.database import Base

class PlanDeCuentas(Base):
    """(Req 2) Almacena el sistema de mapping universal, reutilizable."""
    __tablename__ = "plan_de_cuentas"

    id = Column(Integer, primary_key=True, index=True)
    cliente_id = Column(Integer, ForeignKey("clientes.id"), nullable=True) # Si es genérico, Null
    cuenta_origen = Column(String, index=True, nullable=False) # Ej: '1101'
    nombre_cuenta_origen = Column(String)  # Ej: 'Caja Moneda Nacional'
    categoria_financiera = Column(String, index=True, nullable=False) # Ej: 'Efectivo y Equivalentes'
    clasificacion = Column(String) # Ej: 'Activo Corriente'
    
    __table_args__ = (UniqueConstraint('cliente_id', 'cuenta_origen', name='_cliente_cuenta_uc'),)
