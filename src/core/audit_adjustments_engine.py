import re
from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple
import pandas as pd
from sqlalchemy import func
from src.models.database import SessionLocal
from src.models.audit_adjustment import AuditAdjustmentRecord

def generar_siguiente_codigo_ajuste(empresa: str, periodo: str, db=None) -> str:
    """
    Genera el siguiente correlativo para asientos de auditoría individual.
    Ejemplo: AJU-202512-001, AJU-202512-002, etc.
    """
    is_local = False
    if db is None:
        db = SessionLocal()
        is_local = True
    try:
        entries = db.query(AuditAdjustmentRecord).filter(
            AuditAdjustmentRecord.empresa == empresa,
            AuditAdjustmentRecord.periodo == periodo
        ).all()

        max_seq = 0
        for a in entries:
            code = getattr(a, 'asiento_codigo', None)
            if code:
                match = re.search(r'AJU-\d{6}-(\d+)', code)
                if match:
                    seq = int(match.group(1))
                    if seq > max_seq:
                        max_seq = seq

        next_seq = max_seq + 1
        periodo_clean = periodo.replace('-', '').strip()
        return f"AJU-{periodo_clean}-{next_seq:03d}"
    finally:
        if is_local:
            db.close()

def guardar_asiento_ajuste(
    empresa: str, 
    periodo: str, 
    glosa: str, 
    lineas: List[Dict[str, Any]], 
    asiento_codigo: Optional[str] = None,
    db=None
) -> Tuple[bool, str, Optional[str]]:
    """
    Valida y guarda un asiento de ajuste de auditoría completo con partida doble.
    Si se proporciona `asiento_codigo`, actualiza el asiento existente reemplazando sus líneas.
    """
    if not lineas:
        return False, "El asiento debe tener al menos una línea.", None

    glosa_clean = (glosa or "").strip()
    if not glosa_clean:
        return False, "La glosa explicativa es obligatoria.", None

    tot_debe = sum(float(l.get("debe", 0.0)) for l in lineas)
    tot_haber = sum(float(l.get("haber", 0.0)) for l in lineas)
    diff = round(tot_debe - tot_haber, 2)

    if abs(diff) > 0.01:
        return False, f"El asiento no está cuadrado. Debe: ${tot_debe:,.0f} | Haber: ${tot_haber:,.0f} | Diferencia: ${diff:,.0f}", None

    if tot_debe == 0.0 and tot_haber == 0.0:
        return False, "El asiento no puede tener montos en cero.", None

    is_local = False
    if db is None:
        db = SessionLocal()
        is_local = True
    try:
        now = datetime.utcnow()
        if asiento_codigo and str(asiento_codigo).strip():
            codigo_asiento = str(asiento_codigo).strip()
            # Eliminar las líneas previas del asiento a actualizar
            db.query(AuditAdjustmentRecord).filter(
                AuditAdjustmentRecord.empresa == empresa,
                AuditAdjustmentRecord.periodo == periodo,
                AuditAdjustmentRecord.asiento_codigo == codigo_asiento
            ).delete()
        else:
            codigo_asiento = generar_siguiente_codigo_ajuste(empresa, periodo, db)


        records = []
        for idx, l in enumerate(lineas, start=1):
            debe_val = float(l.get("debe", 0.0))
            haber_val = float(l.get("haber", 0.0))
            monto_neto = debe_val - haber_val
            cta_cod = str(l.get("cuenta_codigo", "")).strip()
            nom_cta = str(l.get("nombre_cuenta", "")).strip()
            rubro = str(l.get("linea_item", "")).strip()
            nota = str(l.get("linea_nota", "")).strip() if l.get("linea_nota") else None

            rec = AuditAdjustmentRecord(
                empresa=empresa,
                periodo=periodo,
                asiento_codigo=codigo_asiento,
                num_linea=idx,
                cuenta_codigo=cta_cod,
                cuenta_id=cta_cod,
                nombre_cuenta=nom_cta,
                linea_item=rubro,
                linea_nota=nota,
                debe=debe_val,
                haber=haber_val,
                monto=monto_neto,
                glosa=glosa_clean,
                descripcion=glosa_clean,
                fecha_creacion=now
            )
            records.append(rec)

        db.bulk_save_objects(records)
        db.commit()
        return True, f"Asiento {codigo_asiento} guardado exitosamente.", codigo_asiento
    except Exception as e:
        db.rollback()
        return False, f"Error al guardar asiento: {str(e)}", None
    finally:
        if is_local:
            db.close()

def obtener_asientos_ajuste(empresa: str, periodo: str, db=None) -> List[Dict[str, Any]]:
    """
    Retorna la lista de comprobantes agrupados para el periodo y empresa dados.
    """
    is_local = False
    if db is None:
        db = SessionLocal()
        is_local = True
    try:
        records = db.query(AuditAdjustmentRecord).filter(
            AuditAdjustmentRecord.empresa == empresa,
            AuditAdjustmentRecord.periodo == periodo
        ).order_by(
            AuditAdjustmentRecord.asiento_codigo.desc(),
            AuditAdjustmentRecord.num_linea.asc()
        ).all()

        vouchers: Dict[str, Dict[str, Any]] = {}
        for r in records:
            cod = r.asiento_codigo
            if cod not in vouchers:
                vouchers[cod] = {
                    "asiento_codigo": cod,
                    "glosa": r.glosa,
                    "fecha_creacion": r.fecha_creacion,
                    "total_debe": 0.0,
                    "total_haber": 0.0,
                    "lineas": []
                }
            vouchers[cod]["total_debe"] += r.debe
            vouchers[cod]["total_haber"] += r.haber
            vouchers[cod]["lineas"].append({
                "num_linea": r.num_linea,
                "cuenta_codigo": r.cuenta_codigo,
                "nombre_cuenta": r.nombre_cuenta,
                "linea_item": r.linea_item,
                "linea_nota": r.linea_nota,
                "debe": r.debe,
                "haber": r.haber,
                "monto_neto": r.monto
            })

        return list(vouchers.values())
    finally:
        if is_local:
            db.close()

def eliminar_asiento_ajuste(empresa: str, periodo: str, asiento_codigo: str, db=None) -> Tuple[bool, str]:
    """
    Elimina un comprobante completo de ajuste de auditoría.
    """
    is_local = False
    if db is None:
        db = SessionLocal()
        is_local = True
    try:
        deleted = db.query(AuditAdjustmentRecord).filter(
            AuditAdjustmentRecord.empresa == empresa,
            AuditAdjustmentRecord.periodo == periodo,
            AuditAdjustmentRecord.asiento_codigo == asiento_codigo
        ).delete()
        db.commit()
        if deleted > 0:
            return True, f"Comprobante {asiento_codigo} eliminado correctamente."
        return False, f"No se encontró el comprobante {asiento_codigo}."
    except Exception as e:
        db.rollback()
        return False, f"Error al eliminar asiento: {str(e)}"
    finally:
        if is_local:
            db.close()

def obtener_ajustes_tb_df(empresa: str, periodo: str, db=None) -> Optional[pd.DataFrame]:
    """
    Obtiene los ajustes agregados por cuenta_codigo listos para inyectarse al Trial Balance.
    Retorna un DataFrame con ['cuenta_id', 'descripcion', 'saldo_inicial', 'debitos', 'creditos', 'saldo_final'].
    """
    is_local = False
    if db is None:
        db = SessionLocal()
        is_local = True
    try:
        rows = db.query(
            AuditAdjustmentRecord.cuenta_codigo,
            func.max(AuditAdjustmentRecord.nombre_cuenta).label('nombre_cuenta'),
            func.max(AuditAdjustmentRecord.linea_item).label('linea_item'),
            func.sum(AuditAdjustmentRecord.debe).label('total_debe'),
            func.sum(AuditAdjustmentRecord.haber).label('total_haber'),
            func.sum(AuditAdjustmentRecord.monto).label('total_monto')
        ).filter(
            AuditAdjustmentRecord.empresa == empresa,
            AuditAdjustmentRecord.periodo == periodo
        ).group_by(
            AuditAdjustmentRecord.cuenta_codigo
        ).all()

        if not rows:
            return None

        data = []
        for r in rows:
            desc = r.nombre_cuenta or r.linea_item or f"Ajuste Auditoría {r.cuenta_codigo}"
            data.append({
                'cuenta_id': str(r.cuenta_codigo).strip(),
                'descripcion': str(desc).strip(),
                'saldo_inicial': 0.0,
                'debitos': float(r.total_debe or 0.0),
                'creditos': float(r.total_haber or 0.0),
                'saldo_final': float(r.total_monto or 0.0) # Debe - Haber
            })
        return pd.DataFrame(data)
    finally:
        if is_local:
            db.close()
