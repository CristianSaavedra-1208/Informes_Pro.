import hashlib
import datetime
from sqlalchemy.orm import Session
from src.models.database import SessionLocal
from src.models.security import UserModel, AuditLogRecord

def hash_password(password: str) -> str:
    if not password:
        return ""
    return hashlib.sha256(password.encode('utf-8')).hexdigest()

def verify_password(password: str, stored_hash: str) -> bool:
    if not password or not stored_hash:
        return False
    return hash_password(password) == stored_hash

def register_audit_log(usuario: str, accion: str, entidad_id: str = None, detalles: str = None):
    """Registra una acción en la bitácora de auditoría."""
    db: Session = SessionLocal()
    try:
        log_entry = AuditLogRecord(
            fecha_hora=datetime.datetime.utcnow(),
            usuario=usuario or "Sistema",
            accion=accion,
            entidad_id=entidad_id or "",
            detalles=detalles or ""
        )
        db.add(log_entry)
        db.commit()
    except Exception as e:
        db.rollback()
        print(f"Error al registrar log de auditoría: {e}")
    finally:
        db.close()

def authenticate_user(username: str, password: str):
    """Autentica a un usuario y retorna su diccionario de perfil si es válido y activo."""
    if not username or not password:
        return None
        
    db: Session = SessionLocal()
    try:
        user_obj = db.query(UserModel).filter(UserModel.usuario == username.strip()).first()
        if not user_obj:
            return None
            
        if not user_obj.activo:
            return "INACTIVO"
            
        if verify_password(password, user_obj.password_hash):
            user_obj.last_login = datetime.datetime.utcnow()
            db.commit()
            
            user_data = user_obj.to_dict()
            register_audit_log(
                usuario=username,
                accion="LOGIN",
                entidad_id=username,
                detalles=f"Inicio de sesión exitoso como rol {user_obj.rol}"
            )
            return user_data
    except Exception as e:
        db.rollback()
        print(f"Error en autenticación: {e}")
    finally:
        db.close()
    return None

def create_user(username: str, password: str, nombre_completo: str = None, email: str = None, rol: str = "Analista Contable", created_by: str = "admin"):
    """Crea un nuevo usuario en la base de datos."""
    if not username or not password:
        return False, "Nombre de usuario y contraseña son obligatorios."
        
    db: Session = SessionLocal()
    try:
        u_clean = username.strip()
        existing = db.query(UserModel).filter(UserModel.usuario == u_clean).first()
        if existing:
            return False, f"El usuario '{u_clean}' ya se encuentra registrado."
            
        new_user = UserModel(
            usuario=u_clean,
            password_hash=hash_password(password),
            nombre_completo=nombre_completo or u_clean,
            email=email or "",
            rol=rol or "Analista Contable",
            activo=True,
            created_at=datetime.datetime.utcnow()
        )
        db.add(new_user)
        db.commit()
        
        register_audit_log(
            usuario=created_by,
            accion="CREACION_USUARIO",
            entidad_id=u_clean,
            detalles=f"Usuario creado con rol '{rol}'."
        )
        return True, f"Usuario '{u_clean}' creado exitosamente."
    except Exception as e:
        db.rollback()
        return False, f"Error al crear usuario: {str(e)}"
    finally:
        db.close()

def update_user_role(username: str, nuevo_rol: str, admin_username: str = "admin"):
    """Actualiza el rol de un usuario existente."""
    if username == "admin" and nuevo_rol != "Administrador":
        return False, "La cuenta principal 'admin' no puede cambiar su rol de Administrador."

    db: Session = SessionLocal()
    try:
        user_obj = db.query(UserModel).filter(UserModel.usuario == username).first()
        if not user_obj:
            return False, f"El usuario '{username}' no existe."
            
        rol_anterior = user_obj.rol
        user_obj.rol = nuevo_rol
        db.commit()
        
        register_audit_log(
            usuario=admin_username,
            accion="CAMBIO_ROL",
            entidad_id=username,
            detalles=f"Rol cambiado de '{rol_anterior}' a '{nuevo_rol}'."
        )
        return True, f"Rol de '{username}' actualizado a '{nuevo_rol}'."
    except Exception as e:
        db.rollback()
        return False, f"Error al actualizar rol: {str(e)}"
    finally:
        db.close()

def update_user_status(username: str, activo: bool, admin_username: str = "admin"):
    """Habilita o deshabilita el acceso de un usuario."""
    if username == "admin" and not activo:
        return False, "La cuenta principal 'admin' no se puede deshabilitar."

    db: Session = SessionLocal()
    try:
        user_obj = db.query(UserModel).filter(UserModel.usuario == username).first()
        if not user_obj:
            return False, f"El usuario '{username}' no existe."
            
        user_obj.activo = activo
        db.commit()
        
        estado_str = "Habilitado" if activo else "Deshabilitado"
        register_audit_log(
            usuario=admin_username,
            accion="CAMBIO_ESTADO_USUARIO",
            entidad_id=username,
            detalles=f"Estado del usuario cambiado a {estado_str}."
        )
        return True, f"Usuario '{username}' {estado_str} correctamente."
    except Exception as e:
        db.rollback()
        return False, f"Error al actualizar estado: {str(e)}"
    finally:
        db.close()

def change_user_password(username: str, nueva_clave: str, actor_username: str = "admin"):
    """Cambia la clave de un usuario."""
    if not nueva_clave:
        return False, "La nueva contraseña no puede estar vacía."
        
    db: Session = SessionLocal()
    try:
        user_obj = db.query(UserModel).filter(UserModel.usuario == username).first()
        if not user_obj:
            return False, f"El usuario '{username}' no existe."
            
        user_obj.password_hash = hash_password(nueva_clave)
        db.commit()
        
        register_audit_log(
            usuario=actor_username,
            accion="CAMBIO_CLAVE",
            entidad_id=username,
            detalles="Contraseña actualizada exitosamente."
        )
        return True, f"Contraseña de '{username}' actualizada con éxito."
    except Exception as e:
        db.rollback()
        return False, f"Error al cambiar clave: {str(e)}"
    finally:
        db.close()

def delete_user(username: str, admin_username: str = "admin"):
    """Elimina un usuario de la base de datos."""
    if username == "admin":
        return False, "No se puede eliminar la cuenta principal de Administrador ('admin')."
        
    db: Session = SessionLocal()
    try:
        user_obj = db.query(UserModel).filter(UserModel.usuario == username).first()
        if not user_obj:
            return False, f"El usuario '{username}' no existe."
            
        db.delete(user_obj)
        db.commit()
        
        register_audit_log(
            usuario=admin_username,
            accion="ELIMINACION_USUARIO",
            entidad_id=username,
            detalles=f"Usuario '{username}' eliminado del sistema."
        )
        return True, f"Usuario '{username}' eliminado correctamente."
    except Exception as e:
        db.rollback()
        return False, f"Error al eliminar usuario: {str(e)}"
    finally:
        db.close()

def get_all_users():
    """Retorna la lista de todos los usuarios registrados."""
    db: Session = SessionLocal()
    try:
        users = db.query(UserModel).order_by(UserModel.usuario).all()
        return [u.to_dict() for u in users]
    finally:
        db.close()

def get_audit_logs(filter_user: str = None, filter_action: str = None, limit: int = 200):
    """Obtiene los registros de la bitácora de auditoría."""
    db: Session = SessionLocal()
    try:
        query = db.query(AuditLogRecord)
        if filter_user:
            query = query.filter(AuditLogRecord.usuario == filter_user)
        if filter_action:
            query = query.filter(AuditLogRecord.accion == filter_action)
            
        logs = query.order_by(AuditLogRecord.id.desc()).limit(limit).all()
        return [l.to_dict() for l in logs]
    finally:
        db.close()
