import pandas as pd
from sqlalchemy import func
from typing import Optional, List
from src.models.database import SessionLocal
from src.models.pl_record import PlRecordDim

class PlCuboDB:
    
    @staticmethod
    def save_pl_cubo(empresa: str, periodo: str, df: pd.DataFrame) -> int:
        """
        Recibe un DataFrame con el cubo transaccional de P&L, formatea y guarda 
        dimensionalmente para permitir infinidad de columnas sin hardcoding.
        """
        db = SessionLocal()
        try:
            # Borrar data previa
            db.query(PlRecordDim).filter(
                PlRecordDim.empresa == empresa,
                PlRecordDim.periodo == periodo
            ).delete()
            
            ignorar = ["n° de cuenta", "nombre de la cuenta", "cuenta", "nombre", "unnamed: 0"]
            
            # Identificar dinámicamente columnas de llaves (Cuenta y Nombre)
            cuenta_col = next((c for c in df.columns if "cuenta" in str(c).lower() and "nombre" not in str(c).lower()), None)
            desc_col = next((c for c in df.columns if "nombre" in str(c).lower()), None)
            
            if not cuenta_col:
                raise ValueError("El archivo subido no tiene una columna llamada 'Cuenta'")
                
            records = []
            
            for _, row in df.iterrows():
                cuenta_id = str(row.get(cuenta_col, "")).strip()
                if not cuenta_id or cuenta_id.lower() == 'nan':
                    continue
                    
                desc = str(row.get(desc_col, "")).strip() if desc_col else ""
                
                for col in df.columns:
                    col_str = str(col).strip()
                    if col_str.lower() in ignorar or col_str == str(cuenta_col) or col_str == str(desc_col):
                        continue
                    
                    v = row.get(col, 0.0)
                    try:
                        val = float(v) if pd.notna(v) and str(v).strip() != '' else 0.0
                    except:
                        val = 0.0
                        
                    if val != 0.0:
                        records.append(PlRecordDim(
                            empresa=empresa,
                            periodo=periodo,
                            cuenta_id=cuenta_id,
                            descripcion=desc,
                            rubro=col_str,
                            monto=val
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
    def get_pl_cubo(empresa: str, periodo: str) -> Optional[pd.DataFrame]:
        """
        Retorna el Cubo P&L agrupado por cuenta de un mes específico como DataFrame.
        Usa pivoteo con pandas para reconstruir las columnas dinámicamente.
        """
        db = SessionLocal()
        try:
            records = db.query(
                PlRecordDim.cuenta_id,
                func.max(PlRecordDim.descripcion).label('descripcion'),
                PlRecordDim.rubro,
                func.sum(PlRecordDim.monto).label('monto')
            ).filter(
                PlRecordDim.empresa == empresa,
                PlRecordDim.periodo == periodo
            ).group_by(
                PlRecordDim.cuenta_id,
                PlRecordDim.rubro
            ).all()

            # Si no hay registros de cubo cargados en absoluto para este periodo, ver históricos
            if not records:
                from src.models.historical_data import HistoricalDataRecord
                hist_records = db.query(HistoricalDataRecord).filter(
                    HistoricalDataRecord.empresa == empresa,
                    HistoricalDataRecord.periodo == periodo,
                    HistoricalDataRecord.reporte == 'P&L'
                ).all()
                if hist_records:
                    row = {
                        'N° de cuenta': 'dummy',
                        'Nombre de la cuenta': 'dummy'
                    }
                    for r in hist_records:
                        row[r.linea_item] = r.monto
                    return pd.DataFrame([row])
                return None

            # Construir DataFrame crudo (Melted)
            raw_data = [
                {
                    'N° de cuenta': r.cuenta_id,
                    'Nombre de la cuenta': r.descripcion,
                    'rubro': r.rubro,
                    'monto': r.monto
                } for r in records
            ]
            df_raw = pd.DataFrame(raw_data)
            
            # Pivotear DataFrame
            df_pivot = df_raw.pivot_table(
                index=['N° de cuenta', 'Nombre de la cuenta'],
                columns='rubro',
                values='monto',
                aggfunc='sum',
                fill_value=0.0
            ).reset_index()
            
            # Verificar si existen registros consolidados/congelados en la memoria histórica para escalar
            from src.models.historical_data import HistoricalDataRecord
            hist_records = db.query(HistoricalDataRecord).filter(
                HistoricalDataRecord.empresa == empresa,
                HistoricalDataRecord.periodo == periodo,
                HistoricalDataRecord.reporte == 'P&L'
            ).all()
            
            if hist_records:
                hist_dict = {r.linea_item: r.monto for r in hist_records}
                
                # Escalar proporcionalmente cada columna del P&L Cubo pivoteado
                for col in df_pivot.columns:
                    if col in ['N° de cuenta', 'Nombre de la cuenta']:
                        continue
                    
                    hist_monto = hist_dict.get(col)
                    if hist_monto is not None:
                        raw_sum = df_pivot[col].sum()
                        if abs(raw_sum) > 0.001:
                            scale = hist_monto / raw_sum
                            df_pivot[col] = df_pivot[col] * scale
                        else:
                            # Si era 0 pero en el histórico no lo es, inyectamos en la primera cuenta
                            df_pivot.loc[df_pivot.index[0], col] = hist_monto
                    else:
                        df_pivot[col] = 0.0
                        
                # Agregar columnas históricas que no existían en el cubo crudo
                for h_col, h_monto in hist_dict.items():
                    if h_col not in df_pivot.columns:
                        df_pivot[h_col] = 0.0
                        df_pivot.loc[df_pivot.index[0], h_col] = h_monto
            
            return df_pivot
        finally:
            db.close()

    @staticmethod
    def get_pl_cubo_total_sum(empresa: str, periodo: str) -> float:
        """
        Retorna la suma total de todos los montos cargados en el Cubo P&L
        para una empresa y período determinados.
        """
        db = SessionLocal()
        try:
            # Primero, verificar registros históricos
            from src.models.historical_data import HistoricalDataRecord
            hist_records = db.query(HistoricalDataRecord).filter(
                HistoricalDataRecord.empresa == empresa,
                HistoricalDataRecord.periodo == periodo,
                HistoricalDataRecord.reporte == 'P&L'
            ).all()
            if hist_records:
                return sum(r.monto for r in hist_records)

            total = db.query(func.sum(PlRecordDim.monto)).filter(
                PlRecordDim.empresa == empresa,
                PlRecordDim.periodo == periodo
            ).scalar()
            return float(total) if total is not None else 0.0
        except Exception:
            return 0.0
        finally:
            db.close()

    @staticmethod
    def get_available_periods(empresa: str) -> List[str]:
        db = SessionLocal()
        try:
            periods = db.query(PlRecordDim.periodo).filter(
                PlRecordDim.empresa == empresa
            ).distinct().all()
            
            # Combinar con los periodos que tienen registros de P&L históricos
            from src.models.historical_data import HistoricalDataRecord
            hist_periods = db.query(HistoricalDataRecord.periodo).filter(
                HistoricalDataRecord.empresa == empresa,
                HistoricalDataRecord.reporte == 'P&L'
            ).distinct().all()
            
            all_pers = list(set([p[0] for p in periods] + [p[0] for p in hist_periods]))
            return sorted(all_pers) if all_pers else []
        finally:
            db.close()
