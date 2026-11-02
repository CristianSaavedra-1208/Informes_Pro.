import pandas as pd
from typing import List, Optional
from sqlalchemy.orm import Session
from sqlalchemy import func
from src.models.database import SessionLocal, engine, init_db
from src.models.trial_balance import TrialBalanceRecord

class TrialBalanceDB:
    @staticmethod
    def initialize():
        init_db()

    @staticmethod
    def save_trial_balance(empresa: str, periodo: str, df: pd.DataFrame):
        """
        Guarda o sobreescribe un Trial Balance para un periodo específico.
        `df` debe tener las columnas ['cuenta_id', 'descripcion', 'saldo_final']
        """
        db = SessionLocal()
        try:
            # Borrar data previa para este periodo y empresa
            db.query(TrialBalanceRecord).filter(
                TrialBalanceRecord.empresa == empresa,
                TrialBalanceRecord.periodo == periodo
            ).delete()
            
            # Prepare objects
            records = []
            for _, row in df.iterrows():
                cuenta_id = str(row.get('cuenta_id', '')).strip()
                if not cuenta_id or cuenta_id.lower() in ('nan', 'none', ''):
                    continue
                
                # Handle possible NaN in descripcion
                desc = row.get('descripcion', '')
                if pd.isna(desc):
                    desc = ''
                
                si = row.get('saldo_inicial', 0.0)
                if pd.isna(si): si = 0.0
                deb = row.get('debitos', 0.0)
                if pd.isna(deb): deb = 0.0
                cred = row.get('creditos', 0.0)
                if pd.isna(cred): cred = 0.0
                sf = row.get('saldo_final', 0.0)
                if pd.isna(sf):
                    sf = 0.0

                records.append(TrialBalanceRecord(
                    empresa=empresa,
                    periodo=periodo,
                    cuenta_id=cuenta_id,
                    descripcion=str(desc).strip(),
                    saldo_inicial=float(si),
                    debitos=float(deb),
                    creditos=float(cred),
                    saldo_final=float(sf)
                ))
            
            db.bulk_save_objects(records)
            db.commit()
            return len(records)
        except Exception as e:
            db.rollback()
            raise e
        finally:
            db.close()

    @staticmethod
    def get_trial_balance(empresa: str, periodo: str) -> Optional[pd.DataFrame]:
        """
        Retorna el Trial Balance de un mes específico como DataFrame.
        Retorna None si no hay datos.
        """
        db = SessionLocal()
        try:
            records = db.query(
                TrialBalanceRecord.cuenta_id,
                func.max(TrialBalanceRecord.descripcion).label('descripcion'),
                func.sum(TrialBalanceRecord.saldo_inicial).label('saldo_inicial'),
                func.sum(TrialBalanceRecord.debitos).label('debitos'),
                func.sum(TrialBalanceRecord.creditos).label('creditos'),
                func.sum(TrialBalanceRecord.saldo_final).label('saldo_final')
            ).filter(
                TrialBalanceRecord.empresa == empresa,
                TrialBalanceRecord.periodo == periodo
            ).group_by(
                TrialBalanceRecord.cuenta_id
            ).all()

            is_active_tb = bool(records)

            if not records:
                from src.models.historical_data import HistoricalDetailRecord
                records = db.query(
                    HistoricalDetailRecord.cuenta_id,
                    func.max(HistoricalDetailRecord.descripcion).label('descripcion'),
                    func.sum(HistoricalDetailRecord.saldo_inicial).label('saldo_inicial'),
                    func.sum(HistoricalDetailRecord.debitos).label('debitos'),
                    func.sum(HistoricalDetailRecord.creditos).label('creditos'),
                    func.sum(HistoricalDetailRecord.saldo_final).label('saldo_final')
                ).filter(
                    HistoricalDetailRecord.empresa == empresa,
                    HistoricalDetailRecord.periodo == periodo
                ).group_by(
                    HistoricalDetailRecord.cuenta_id
                ).all()

            if not records:
                return None

            data = [
                {
                    'cuenta_id': str(r.cuenta_id).strip(),
                    'descripcion': str(r.descripcion or '').strip(),
                    'saldo_inicial': float(r.saldo_inicial or 0.0),
                    'debitos': float(r.debitos or 0.0),
                    'creditos': float(r.creditos or 0.0),
                    'saldo_final': float(r.saldo_final or 0.0)
                } for r in records
            ]
            df_tb = pd.DataFrame(data)

            # Inyectar ajustes de auditoría individual si es empresa individual (activo o histórico cerrado)
            if not str(empresa).startswith("[GRUPO]"):
                try:
                    from src.core.audit_adjustments_engine import obtener_ajustes_tb_df
                    df_adj = obtener_ajustes_tb_df(empresa, periodo, db)
                    if df_adj is not None and not df_adj.empty:
                        # Si provenía de HistoricalDetailRecord, remover cuentas 9999xxx previas para evitar duplicación
                        df_tb = df_tb[~df_tb['cuenta_id'].astype(str).str.strip().str.startswith('9999')].copy()
                        df_tb = pd.concat([df_tb, df_adj], ignore_index=True)
                        df_tb = df_tb.groupby('cuenta_id', as_index=False).agg({
                            'descripcion': 'first',
                            'saldo_inicial': 'sum',
                            'debitos': 'sum',
                            'creditos': 'sum',
                            'saldo_final': 'sum'
                        })
                except Exception as ex:
                    print(f"Advertencia al inyectar ajustes de auditoría en TB: {ex}")

            return df_tb
        finally:
            db.close()

    @staticmethod
    def get_available_periods(empresa: str) -> List[str]:
        """
        Retorna una lista ordenada unificada de los periodos disponibles para una empresa,
        leyendo tanto de la memoria activa (TrialBalance) como de la bóveda histórica.
        """
        db = SessionLocal()
        try:
            if empresa and str(empresa).startswith("[GRUPO]"):
                grupo_name = str(empresa).replace("[GRUPO] ", "").strip()
                from src.models.consolidacion import ConsolidationGroup
                grupo = db.query(ConsolidationGroup).filter_by(nombre_grupo=grupo_name).first()
                if grupo and grupo.empresa_matriz:
                    db.close()
                    return TrialBalanceDB.get_available_periods(grupo.empresa_matriz)

            # Obtener periodos de Memoria Activa
            periods_tb = db.query(TrialBalanceRecord.periodo).filter(
                TrialBalanceRecord.empresa == empresa
            ).distinct().all()
            
            # Obtener periodos de Memoria Histórica (valores pegados)
            from src.models.historical_data import HistoricalDataRecord
            periods_hist = db.query(HistoricalDataRecord.periodo).filter(
                HistoricalDataRecord.empresa == empresa
            ).distinct().all()
            
            res = list(set([p[0] for p in periods_tb] + [p[0] for p in periods_hist]))
            res.sort(reverse=True)
            return res
        finally:
            try:
                db.close()
            except:
                pass
