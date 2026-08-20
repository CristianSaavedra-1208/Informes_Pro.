import os
import pandas as pd
from datetime import datetime
from src.models.database import SessionLocal
from src.models.trial_balance import TrialBalanceRecord
from src.models.audit_adjustment import AuditAdjustmentRecord
from src.models.historical_data import HistoricalDataRecord, HistoricalDetailRecord
from src.models.trial_balance_db import TrialBalanceDB

def ejecutar_cierre_periodo(empresa: str, periodo: str, empresa_path: str):
    """
    Ejecuta el cierre de un periodo:
    1. Lee el Trial Balance activo.
    2. Cruza con los mapeos actuales de Balance de forma vectorizada.
    3. Guarda los registros en HistoricalDetailRecord (Papel de trabajo).
    4. Guarda los agregados en HistoricalDataRecord.
    5. Elimina el periodo de la memoria activa (TrialBalanceRecord, AuditAdjustmentRecord).
    """
    db = SessionLocal()
    try:
        # 1. Obtener el TB activo
        tb_df = TrialBalanceDB.get_trial_balance(empresa, periodo)
        if tb_df is None or tb_df.empty:
            return False, f"No hay datos activos para el periodo {periodo}."

        # 2. Obtener mapeos de forma vectorizada
        map_bal_path = os.path.join(empresa_path, "map_balance.xlsx")
        
        df_bal = pd.DataFrame(columns=['cuenta_id', 'clasificacion_balance', 'id_nota_asociada'])
        if os.path.exists(map_bal_path):
            try:
                df_b_raw = pd.read_excel(map_bal_path, dtype=str)
                df_b_raw.columns = [str(c).strip() for c in df_b_raw.columns]
                cuenta_col = next((c for c in df_b_raw.columns if "cuenta" in c.lower()), None)
                clasif_col = next((c for c in df_b_raw.columns if "clasificaci" in c.lower() and "balance" in c.lower()), None)
                nota_col = next((c for c in df_b_raw.columns if "id_nota" in c.lower()), None)
                
                if cuenta_col and clasif_col:
                    df_bal = pd.DataFrame({
                        'cuenta_id': df_b_raw[cuenta_col].astype(str).str.strip(),
                        'clasificacion_balance': df_b_raw[clasif_col].astype(str).str.strip(),
                        'id_nota_asociada': df_b_raw[nota_col].astype(str).str.strip() if nota_col else None
                    })
            except Exception:
                pass

        # 2b. Leer P&L desde el Cubo Transaccional
        from sqlalchemy import func
        from src.models.pl_record import PlRecordDim
        
        pl_records = db.query(
            PlRecordDim.rubro,
            func.sum(PlRecordDim.monto).label('monto')
        ).filter(
            PlRecordDim.empresa == empresa,
            PlRecordDim.periodo == periodo
        ).group_by(PlRecordDim.rubro).all()

        # 3. Borrar históricos previos si existieran (re-cierre)
        db.query(HistoricalDetailRecord).filter_by(empresa=empresa, periodo=periodo).delete()
        db.query(HistoricalDataRecord).filter_by(empresa=empresa, periodo=periodo).delete()

        # 4. Cruzar TB y Mapeos de forma vectorizada
        tb_df = tb_df.copy()
        tb_df['cuenta_id'] = tb_df['cuenta_id'].astype(str).str.strip()
        merged = pd.merge(tb_df, df_bal, on='cuenta_id', how='left')
        merged['saldo_inicial'] = pd.to_numeric(merged.get('saldo_inicial', 0.0), errors='coerce').fillna(0.0)
        merged['debitos'] = pd.to_numeric(merged.get('debitos', 0.0), errors='coerce').fillna(0.0)
        merged['creditos'] = pd.to_numeric(merged.get('creditos', 0.0), errors='coerce').fillna(0.0)
        merged['saldo_final'] = pd.to_numeric(merged['saldo_final'], errors='coerce').fillna(0.0)

        # Rellenar lista de registros históricos de detalle mediante list comprehension rápido
        details = [
            HistoricalDetailRecord(
                empresa=empresa,
                periodo=periodo,
                cuenta_id=str(r['cuenta_id']),
                descripcion=str(r.get('descripcion', '')),
                saldo_inicial=float(r['saldo_inicial']),
                debitos=float(r['debitos']),
                creditos=float(r['creditos']),
                saldo_final=float(r['saldo_final']),
                clasificacion_balance=str(r['clasificacion_balance']) if pd.notna(r['clasificacion_balance']) and str(r['clasificacion_balance']) != 'nan' and str(r['clasificacion_balance']).strip() != "" else None,
                clasificacion_pl=None,
                id_nota_asociada=str(r['id_nota_asociada']) if pd.notna(r['id_nota_asociada']) and str(r['id_nota_asociada']) != 'nan' and str(r['id_nota_asociada']).strip() != "" else None
            )
            for r in merged.to_dict(orient='records')
        ]

        # 5. Agrupar y preparar summaries usando vectorización
        summaries = {}
        
        # Filtrar registros mapeados para balance
        mapped_bal = merged[
            merged['clasificacion_balance'].notna() & 
            (merged['clasificacion_balance'] != '') & 
            (merged['clasificacion_balance'] != 'nan')
        ]
        
        if not mapped_bal.empty:
            bal_grouped = mapped_bal.groupby('clasificacion_balance')['saldo_final'].sum().to_dict()
            for cls_b, sum_val in bal_grouped.items():
                summaries[('Balance', cls_b)] = float(sum_val)

        # Calcular residuo de P&L (cuentas no clasificadas en Balance van a Resultados del ejercicio)
        unmapped_bal = merged[
            merged['clasificacion_balance'].isna() | 
            (merged['clasificacion_balance'] == '') | 
            (merged['clasificacion_balance'] == 'nan')
        ]
        pl_residual = float(unmapped_bal['saldo_final'].sum())

        # Inyectar el resultado del ejercicio a Resultados acumulados
        found_key = ('Balance', 'Resultados acumulados')
        for k in list(summaries.keys()):
            if k[0] == 'Balance' and k[1].lower() in ['resultados acumulados', 'ganancias (pérdidas) acumuladas']:
                found_key = k
                break
        summaries[found_key] = summaries.get(found_key, 0.0) + pl_residual
                
        # Agregar sumatorias del Cubo P&L
        for r in pl_records:
            if pd.notna(r.rubro) and str(r.rubro).strip():
                key = ('P&L', str(r.rubro).strip())
                summaries[key] = summaries.get(key, 0.0) + float(r.monto)
                
        # 6. Guardar registros de detalle en bloque
        db.bulk_save_objects(details)
        
        # 7. Guardar summaries en bloque
        sum_records = []
        for (reporte, linea), monto in summaries.items():
            sum_records.append(HistoricalDataRecord(
                empresa=empresa,
                periodo=periodo,
                reporte=reporte,
                linea_item=linea,
                monto=monto
            ))
        db.bulk_save_objects(sum_records)
        
        # 8. Borrar data viva
        db.query(TrialBalanceRecord).filter_by(empresa=empresa, periodo=periodo).delete()
        db.query(AuditAdjustmentRecord).filter_by(empresa=empresa, periodo=periodo).delete()
        
        db.commit()
        return True, f"Cierre ejecutado correctamente para {periodo}. Se guardaron {len(details)} registros en memoria histórica."
    except Exception as e:
        db.rollback()
        return False, str(e)
    finally:
        db.close()


def reversar_cierre_periodo(empresa: str, periodo: str):
    """
    Deshace el cierre de un periodo:
    1. Toma los registros detallados de HistoricalDetailRecord.
    2. Los reinserta en TrialBalanceRecord (memoria activa).
    3. Elimina los registros históricos (detalles y resúmenes) de ese periodo.
    """
    db = SessionLocal()
    try:
        # Obtener detalle histórico
        historicos = db.query(HistoricalDetailRecord).filter_by(empresa=empresa, periodo=periodo).all()
        if not historicos:
            return False, f"No se encontró data histórica para el periodo {periodo}."

        # Preparar para memoria activa
        tb_records = []
        for h in historicos:
            tb_records.append(TrialBalanceRecord(
                empresa=empresa,
                periodo=periodo,
                cuenta_id=h.cuenta_id,
                descripcion=h.descripcion,
                saldo_final=h.saldo_final
            ))
            
        # Limpiar memoria activa actual por si acaso
        db.query(TrialBalanceRecord).filter_by(empresa=empresa, periodo=periodo).delete()
        
        # Insertar a memoria activa
        db.bulk_save_objects(tb_records)
        
        # Borrar de memoria histórica
        db.query(HistoricalDetailRecord).filter_by(empresa=empresa, periodo=periodo).delete()
        db.query(HistoricalDataRecord).filter_by(empresa=empresa, periodo=periodo).delete()
        
        db.commit()
        return True, f"El periodo {periodo} ha sido restaurado exitosamente a la memoria activa."
    except Exception as e:
        db.rollback()
        return False, str(e)
    finally:
        db.close()

def es_periodo_cerrado(empresa: str, periodo: str) -> bool:
    """
    Verifica si existe registro de cierre histórico para la empresa y periodo especificados.
    """
    db = SessionLocal()
    try:
        from src.models.historical_data import HistoricalDetailRecord
        exists = db.query(HistoricalDetailRecord).filter_by(empresa=empresa, periodo=periodo).first() is not None
        return exists
    finally:
        db.close()


