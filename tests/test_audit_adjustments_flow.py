import pandas as pd
from src.models.database import SessionLocal, init_db
from src.models.audit_adjustment import AuditAdjustmentRecord
from src.core.audit_adjustments_engine import (
    guardar_asiento_ajuste,
    obtener_asientos_ajuste,
    eliminar_asiento_ajuste,
    obtener_ajustes_tb_df
)
from src.models.trial_balance_db import TrialBalanceDB
from src.models.trial_balance import TrialBalanceRecord

def test_audit_adjustments_full_flow():
    init_db()
    test_empresa = "Pacifico Cable SpA"
    test_periodo = "2026-05"
    
    # Limpiar asientos previos de prueba si hubiesen
    db = SessionLocal()
    db.query(AuditAdjustmentRecord).filter_by(empresa=test_empresa, periodo=test_periodo).delete()
    db.commit()
    db.close()
    
    # 1. Validar que un asiento descuadrado sea rechazado
    lineas_descuadradas = [
        {
            "cuenta_codigo": "9999010",
            "nombre_cuenta": "Ajuste Efectivo",
            "linea_item": "Efectivo y efectivo equivalente",
            "linea_nota": None,
            "debe": 1000000.0,
            "haber": 0.0
        },
        {
            "cuenta_codigo": "9999300",
            "nombre_cuenta": "Ajuste Costos",
            "linea_item": "Costo de ventas",
            "linea_nota": None,
            "debe": 0.0,
            "haber": 800000.0
        }
    ]
    ok, msg, cod = guardar_asiento_ajuste(test_empresa, test_periodo, "Test Descuadrado", lineas_descuadradas)
    assert not ok
    assert "no está cuadrado" in msg
    assert cod is None

    # 2. Guardar asiento cuadrado (Debe = Haber = 1,000,000)
    lineas_cuadradas = [
        {
            "cuenta_codigo": "9999010",
            "nombre_cuenta": "Ajuste Efectivo",
            "linea_item": "Efectivo y efectivo equivalente",
            "linea_nota": None,
            "debe": 1000000.0,
            "haber": 0.0
        },
        {
            "cuenta_codigo": "9999200",
            "nombre_cuenta": "Ajuste Proveedores",
            "linea_item": "Cuentas por pagar comerciales",
            "linea_nota": None,
            "debe": 0.0,
            "haber": 1000000.0
        }
    ]
    ok, msg, cod = guardar_asiento_ajuste(test_empresa, test_periodo, "Ajuste Auditoria Test Cuadrado", lineas_cuadradas)
    assert ok
    assert cod is not None
    assert cod.startswith(f"AJU-{test_periodo.replace('-', '')}")

    # 3. Consultar asientos guardados
    asientos = obtener_asientos_ajuste(test_empresa, test_periodo)
    assert len(asientos) == 1
    assert asientos[0]["asiento_codigo"] == cod
    assert asientos[0]["total_debe"] == 1000000.0
    assert asientos[0]["total_haber"] == 1000000.0
    assert len(asientos[0]["lineas"]) == 2

    # 4. Validar inyección en TrialBalanceDB
    # Crear un TB dummy activo si no existe
    db = SessionLocal()
    has_tb = db.query(TrialBalanceRecord).filter_by(empresa=test_empresa, periodo=test_periodo).first()
    if not has_tb:
        db.add(TrialBalanceRecord(
            empresa=test_empresa,
            periodo=test_periodo,
            cuenta_id="110101",
            descripcion="Caja chica",
            saldo_inicial=0.0,
            debitos=500000.0,
            creditos=0.0,
            saldo_final=500000.0
        ))
        db.commit()
    db.close()

    tb_df = TrialBalanceDB.get_trial_balance(test_empresa, test_periodo)
    assert tb_df is not None
    # Debe contener la cuenta 9999010 inyectada
    row_efectivo = tb_df[tb_df['cuenta_id'] == '9999010']
    assert not row_efectivo.empty
    assert float(row_efectivo['debitos'].values[0]) == 1000000.0
    assert float(row_efectivo['saldo_final'].values[0]) == 1000000.0

    # Debe contener la cuenta 9999200 inyectada
    row_prov = tb_df[tb_df['cuenta_id'] == '9999200']
    assert not row_prov.empty
    assert float(row_prov['creditos'].values[0]) == 1000000.0
    assert float(row_prov['saldo_final'].values[0]) == -1000000.0

    # 5. Modificar / Actualizar el asiento existente (Edit workflow)
    lineas_modificadas = [
        {
            "cuenta_codigo": "9999010",
            "nombre_cuenta": "Ajuste Efectivo Modificado",
            "linea_item": "Efectivo y efectivo equivalente",
            "linea_nota": None,
            "debe": 1500000.0,
            "haber": 0.0
        },
        {
            "cuenta_codigo": "9999200",
            "nombre_cuenta": "Ajuste Proveedores Modificado",
            "linea_item": "Cuentas por pagar comerciales",
            "linea_nota": None,
            "debe": 0.0,
            "haber": 1500000.0
        }
    ]
    ok_edit, msg_edit, cod_edit = guardar_asiento_ajuste(
        test_empresa, test_periodo, "Glosa Modificada", lineas_modificadas, asiento_codigo=cod
    )
    assert ok_edit
    assert cod_edit == cod  # Mismo folio conservado

    asientos_mod = obtener_asientos_ajuste(test_empresa, test_periodo)
    assert len(asientos_mod) == 1
    assert asientos_mod[0]["glosa"] == "Glosa Modificada"
    assert asientos_mod[0]["total_debe"] == 1500000.0
    assert asientos_mod[0]["total_haber"] == 1500000.0

    # 6. Eliminar el asiento de prueba y comprobar limpieza
    ok_del, msg_del = eliminar_asiento_ajuste(test_empresa, test_periodo, cod)
    assert ok_del
    asientos_post = obtener_asientos_ajuste(test_empresa, test_periodo)
    assert len(asientos_post) == 0

    print("[OK] Todos los tests del flujo de Ajustes de Auditoria (incluyendo Edicion) pasaron exitosamente.")


def test_audit_adjustments_historical_closed_period():
    from src.models.pl_cubo_db import PlCuboDB
    from src.models.historical_data import HistoricalDataRecord, HistoricalDetailRecord
    from src.models.pl_record import PlRecordDim

    test_empresa = "Pacifico Cable SpA"
    test_periodo = "2099-12"

    db = SessionLocal()
    # Limpieza previa
    db.query(AuditAdjustmentRecord).filter_by(empresa=test_empresa, periodo=test_periodo).delete()
    db.query(HistoricalDataRecord).filter_by(empresa=test_empresa, periodo=test_periodo).delete()
    db.query(HistoricalDetailRecord).filter_by(empresa=test_empresa, periodo=test_periodo).delete()
    db.query(PlRecordDim).filter_by(empresa=test_empresa, periodo=test_periodo).delete()

    # Simular periodo cerrado con histórico en HistoricalDataRecord y detalle
    db.add(HistoricalDataRecord(
        empresa=test_empresa,
        periodo=test_periodo,
        reporte='P&L',
        linea_item='Ingresos de actividades ordinarias',
        monto=-5000000.0
    ))
    db.add(HistoricalDetailRecord(
        empresa=test_empresa,
        periodo=test_periodo,
        cuenta_id='110101',
        descripcion='Caja',
        saldo_inicial=0.0,
        debitos=100000.0,
        creditos=0.0,
        saldo_final=100000.0
    ))
    # Y un registro transaccional en PlRecordDim
    db.add(PlRecordDim(
        empresa=test_empresa,
        periodo=test_periodo,
        cuenta_id='410101',
        descripcion='Venta',
        rubro='Ingresos de actividades ordinarias',
        monto=-5000000.0
    ))
    db.commit()
    db.close()

    # Guardar ajuste de auditoría para P&L y Balance en periodo cerrado
    lineas = [
        {
            "cuenta_codigo": "9999340",
            "nombre_cuenta": "Ajuste Otros Egresos",
            "linea_item": "Otros egresos por función",
            "linea_nota": None,
            "debe": 250000.0,
            "haber": 0.0
        },
        {
            "cuenta_codigo": "9999010",
            "nombre_cuenta": "Ajuste Efectivo",
            "linea_item": "Efectivo y efectivo equivalente",
            "linea_nota": None,
            "debe": 0.0,
            "haber": 250000.0
        }
    ]
    ok, msg, cod = guardar_asiento_ajuste(test_empresa, test_periodo, "Ajuste Auditoria Periodo Cerrado", lineas)
    assert ok, f"Error guardando ajuste: {msg}"

    # 1. Validar que PlCuboDB.get_pl_cubo NO anula el ajuste a 0.0 por causa de HistoricalDataRecord
    df_pl = PlCuboDB.get_pl_cubo(test_empresa, test_periodo)
    assert df_pl is not None
    assert "Otros egresos por función" in df_pl.columns
    assert float(df_pl["Otros egresos por función"].sum()) == 250000.0

    # 2. Validar que TrialBalanceDB.get_trial_balance inyecta el ajuste aun si es periodo histórico cerrado
    df_tb = TrialBalanceDB.get_trial_balance(test_empresa, test_periodo)
    assert df_tb is not None
    row_9999 = df_tb[df_tb['cuenta_id'] == '9999010']
    assert not row_9999.empty
    assert float(row_9999['creditos'].values[0]) == 250000.0

    # Limpieza final
    db = SessionLocal()
    db.query(AuditAdjustmentRecord).filter_by(empresa=test_empresa, periodo=test_periodo).delete()
    db.query(HistoricalDataRecord).filter_by(empresa=test_empresa, periodo=test_periodo).delete()
    db.query(HistoricalDetailRecord).filter_by(empresa=test_empresa, periodo=test_periodo).delete()
    db.query(PlRecordDim).filter_by(empresa=test_empresa, periodo=test_periodo).delete()
    db.commit()
    db.close()

    print("[OK] Test de Ajustes de Auditoría en períodos cerrados/históricos pasó exitosamente.")


if __name__ == "__main__":
    test_audit_adjustments_full_flow()
    test_audit_adjustments_historical_closed_period()

