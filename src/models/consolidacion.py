from sqlalchemy import Column, Integer, String, Float, Boolean, DateTime, ForeignKey
from src.models.database import Base
from datetime import datetime

class ConsolidationGroup(Base):
    """
    Define un perímetro de consolidación relacionando una matriz y una filial.
    """
    __tablename__ = 'consolidation_groups'
    
    id = Column(Integer, primary_key=True, index=True)
    nombre_grupo = Column(String(100), unique=True, index=True, nullable=False)
    empresa_matriz = Column(String(200), nullable=False)
    empresa_filial = Column(String(200), nullable=False)
    filial_is_group = Column(Boolean, default=False)
    created_at = Column(DateTime, default=datetime.utcnow)

class ConsolidationJournalEntry(Base):
    """
    Guarda los asientos formales de ajuste para la hoja de trabajo de consolidación.
    """
    __tablename__ = 'consolidation_journal_entries'
    
    id = Column(Integer, primary_key=True, index=True)
    grupo_id = Column(Integer, ForeignKey('consolidation_groups.id'), nullable=False)
    periodo = Column(String(7), index=True, nullable=False) # ej YYYY-MM
    fecha = Column(DateTime, default=datetime.utcnow)
    glosa = Column(String(255), nullable=False)
    columna_ajuste = Column(String(100), nullable=False) # ej: "Elim inversion", "PPA"
    linea_item = Column(String(255), nullable=False) # Rubro a afectar
    linea_nota = Column(String(255), nullable=True) # Sub-ítem o desglose de Nota a los EEFF afectada
    debe = Column(Float, default=0.0)
    haber = Column(Float, default=0.0)
    es_recurrente = Column(Boolean, default=False)
    elimina_saldo_total = Column(Boolean, default=False)
    asiento_codigo = Column(String(50), index=True, nullable=True) # ej AST-202605-001
    num_linea = Column(Integer, default=1)
    created_by = Column(String(100), nullable=True) # Auditoría: usuario creador
    updated_by = Column(String(100), nullable=True) # Auditoría: usuario modificador
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    created_at = Column(DateTime, default=datetime.utcnow)
