import datetime
from sqlalchemy import Column, Integer, String, Boolean, DateTime, Text
from src.models.database import Base

class UserModel(Base):
    __tablename__ = 'usuarios'

    usuario = Column(String(50), primary_key=True, index=True)
    password_hash = Column(String(255), nullable=False)
    nombre_completo = Column(String(100), nullable=True)
    email = Column(String(100), nullable=True)
    rol = Column(String(50), default='Analista Contable')
    activo = Column(Boolean, default=True)
    created_at = Column(DateTime, default=datetime.datetime.utcnow)
    last_login = Column(DateTime, nullable=True)

    def to_dict(self):
        return {
            "usuario": self.usuario,
            "nombre_completo": self.nombre_completo or self.usuario,
            "email": self.email or "",
            "rol": self.rol or "Analista Contable",
            "activo": self.activo,
            "created_at": self.created_at.strftime("%Y-%m-%d %H:%M:%S") if self.created_at else "",
            "last_login": self.last_login.strftime("%Y-%m-%d %H:%M:%S") if self.last_login else "Nunca"
        }

class AuditLogRecord(Base):
    __tablename__ = 'bitacora_auditoria'

    id = Column(Integer, primary_key=True, autoincrement=True)
    fecha_hora = Column(DateTime, default=datetime.datetime.utcnow)
    usuario = Column(String(50), index=True)
    accion = Column(String(100), index=True)
    entidad_id = Column(String(100), nullable=True)
    detalles = Column(Text, nullable=True)

    def to_dict(self):
        return {
            "id": self.id,
            "fecha_hora": self.fecha_hora.strftime("%Y-%m-%d %H:%M:%S") if self.fecha_hora else "",
            "usuario": self.usuario or "Sistema",
            "accion": self.accion or "",
            "entidad_id": self.entidad_id or "",
            "detalles": self.detalles or ""
        }
