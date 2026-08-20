from sqlalchemy import Column, Integer, String, Boolean, ForeignKey
from src.models.database import Base

class CompanyEntity(Base):
    __tablename__ = "company_entities"

    id = Column(Integer, primary_key=True, index=True)
    nombre = Column(String, unique=True, index=True, nullable=False)
    rut = Column(String, nullable=True)
    es_consolidado = Column(Boolean, default=False)
    id_matriz = Column(Integer, ForeignKey("company_entities.id"), nullable=True)
    moneda_funcional = Column(String(3), default="CLP")
    activa = Column(Boolean, default=True)
