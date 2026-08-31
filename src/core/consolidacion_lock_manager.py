import os
from datetime import datetime
from sqlalchemy import or_
from src.models.database import SessionLocal, Base, engine
from src.models.consolidacion import ConsolidationGroup, ConsolidationJournalEntry, ConsolidationPeriodLock
from src.models.trial_balance_db import TrialBalanceDB
from src.models.pl_cubo_db import PlCuboDB
from src.core.company_resolver import CompanyResolver

# Asegurar creación de la tabla
Base.metadata.create_all(bind=engine)

class ConsolidationLockManager:
    """
    Gestiona el candado de cierre y protección contra escritura para períodos consolidados.
    """

    @staticmethod
    def is_period_locked(grupo_id: int, periodo: str) -> bool:
        """Verifica si un período específico está cerrado con candado."""
        if not grupo_id or not periodo:
            return False
        db = SessionLocal()
        try:
            lock = db.query(ConsolidationPeriodLock).filter_by(
                grupo_id=grupo_id,
                periodo=str(periodo).strip(),
                is_locked=True
            ).first()
            return lock is not None
        finally:
            db.close()

    @staticmethod
    def lock_period(grupo_id: int, periodo: str, user: str = "Administrador") -> tuple:
        """Cierra y bloquea un período de consolidación."""
        if not grupo_id or not periodo:
            return False, "Grupo o período inválido."
            
        periodo_clean = str(periodo).strip()
        db = SessionLocal()
        try:
            lock = db.query(ConsolidationPeriodLock).filter_by(
                grupo_id=grupo_id,
                periodo=periodo_clean
            ).first()
            
            if lock:
                lock.is_locked = True
                lock.locked_at = datetime.now()
                lock.locked_by = user
            else:
                new_lock = ConsolidationPeriodLock(
                    grupo_id=grupo_id,
                    periodo=periodo_clean,
                    is_locked=True,
                    locked_at=datetime.now(),
                    locked_by=user
                )
                db.add(new_lock)
            db.commit()
            return True, f"Período {periodo_clean} cerrado y bloqueado exitosamente."
        except Exception as e:
            db.rollback()
            return False, f"Error al cerrar período: {e}"
        finally:
            db.close()

    @staticmethod
    def unlock_period(grupo_id: int, periodo: str, user: str = "Administrador") -> tuple:
        """Reabre un período de consolidación para permitir ajustes."""
        if not grupo_id or not periodo:
            return False, "Grupo o período inválido."
            
        periodo_clean = str(periodo).strip()
        db = SessionLocal()
        try:
            lock = db.query(ConsolidationPeriodLock).filter_by(
                grupo_id=grupo_id,
                periodo=periodo_clean
            ).first()
            
            if lock:
                lock.is_locked = False
                lock.unlocked_at = datetime.now()
                lock.unlocked_by = user
                db.commit()
                return True, f"Período {periodo_clean} reabierto exitosamente para ajustes."
            return False, f"El período {periodo_clean} no estaba cerrado."
        except Exception as e:
            db.rollback()
            return False, f"Error al reabrir período: {e}"
        finally:
            db.close()

    @staticmethod
    def get_locked_periods(grupo_id: int) -> list:
        """Retorna la lista de períodos bloqueados para un grupo."""
        db = SessionLocal()
        try:
            locks = db.query(ConsolidationPeriodLock).filter_by(
                grupo_id=grupo_id,
                is_locked=True
            ).order_by(ConsolidationPeriodLock.periodo.desc()).all()
            
            return [{
                "id": l.id,
                "periodo": l.periodo,
                "locked_at": l.locked_at,
                "locked_by": l.locked_by
            } for l in locks]
        finally:
            db.close()

    @staticmethod
    def get_available_periods_for_group(grupo_id: int) -> list:
        """
        Retorna la lista de todos los períodos disponibles para un grupo consolidado,
        consultando los períodos cargados en la matriz, filiales y asientos registrados.
        """
        db = SessionLocal()
        try:
            grupo = db.query(ConsolidationGroup).filter_by(id=grupo_id).first()
            if not grupo:
                return []
                
            matriz_name = CompanyResolver.get_display_name(grupo.empresa_matriz)
            filial_name = CompanyResolver.get_display_name(grupo.empresa_filial)
            
            # Periodos en base de datos de matriz y filial
            TrialBalanceDB.initialize()
            tb_matriz = TrialBalanceDB.get_available_periods(matriz_name) or []
            tb_filial = TrialBalanceDB.get_available_periods(filial_name) or []
            pl_matriz = PlCuboDB.get_available_periods(matriz_name) or []
            pl_filial = PlCuboDB.get_available_periods(filial_name) or []
            
            # Periodos en asientos de consolidacion
            asientos_pers = [r[0] for r in db.query(ConsolidationJournalEntry.periodo).filter_by(grupo_id=grupo_id).distinct().all()]
            
            all_pers = set(tb_matriz + tb_filial + pl_matriz + pl_filial + asientos_pers)
            return sorted(list(all_pers), reverse=True)
        finally:
            db.close()
