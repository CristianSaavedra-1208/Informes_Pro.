from sqlalchemy import Column, Integer, String, DateTime, ForeignKey
from sqlalchemy.sql import func
from src.models.database import Base

class Cliente(Base):
    __tablename__ = "clientes"

    id = Column(Integer, primary_key=True, index=True)
    nombre = Column(String, unique=True, index=True, nullable=False)
    rut = Column(String, unique=True, nullable=True)
    created_at = Column(DateTime(timezone=True), server_default=func.now())

class ProyectoFinanciero(Base):
    """Representa los EEFF de un cliente en un año específico."""
    __tablename__ = "proyectos"

    id = Column(Integer, primary_key=True, index=True)
    cliente_id = Column(Integer, ForeignKey("clientes.id"), nullable=False)
    ano = Column(Integer, nullable=False)
    mes = Column(Integer, nullable=False) # 1 a 12
    estado = Column(String, default="Borrador") # Borrador, Revisado, Auditado
    created_at = Column(DateTime(timezone=True), server_default=func.now())

class ReporteVersion(Base):
    """Historial de versiones de informes generados (Req 11)."""
    __tablename__ = "versiones_reporte"

    id = Column(Integer, primary_key=True, index=True)
    proyecto_id = Column(Integer, ForeignKey("proyectos.id"), nullable=False)
    version_num = Column(Integer, nullable=False)
    ruta_archivo = Column(String, nullable=False) # ej: /data/exports/ClienteA_2024_v1.docx
    created_at = Column(DateTime(timezone=True), server_default=func.now())
