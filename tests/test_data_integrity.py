import unittest
import pandas as pd
from sqlalchemy import func
from src.models.database import SessionLocal
from src.models.pl_record import PlRecordDim
from src.models.trial_balance_db import TrialBalanceDB, TrialBalanceRecord
from src.models.historical_data import HistoricalDataRecord
from src.models.pl_cubo_db import PlCuboDB

class TestDataIntegrity(unittest.TestCase):

    def setUp(self):
        self.db = SessionLocal()

    def tearDown(self):
        self.db.close()

    def test_existing_data_snapshot_immutability(self):
        """
        Verifica que los números y registros existentes en la base de datos
        antes de cualquier operación de carga o edición para un nuevo periodo
        permanezcan 100% inalterados.
        """
        # 1. Snapshot inicial de PlRecordDim
        snapshot_pl_before = self.db.query(
            PlRecordDim.empresa,
            PlRecordDim.periodo,
            func.count(PlRecordDim.id).label('cant'),
            func.sum(PlRecordDim.monto).label('suma')
        ).group_by(PlRecordDim.empresa, PlRecordDim.periodo).all()

        # Convertir a diccionario para comparación exacta
        dict_pl_before = {
            (r.empresa, r.periodo): (r.cant, float(r.suma) if r.suma is not None else 0.0)
            for r in snapshot_pl_before
        }

        # 2. Snapshot inicial de TrialBalanceRecord
        snapshot_tb_before = self.db.query(
            TrialBalanceRecord.empresa,
            TrialBalanceRecord.periodo,
            func.count(TrialBalanceRecord.id).label('cant'),
            func.sum(TrialBalanceRecord.saldo_final).label('suma')
        ).group_by(TrialBalanceRecord.empresa, TrialBalanceRecord.periodo).all()

        dict_tb_before = {
            (r.empresa, r.periodo): (r.cant, float(r.suma) if r.suma is not None else 0.0)
            for r in snapshot_tb_before
        }

        # 3. Snapshot inicial de HistoricalDataRecord
        snapshot_hist_before = self.db.query(
            HistoricalDataRecord.empresa,
            HistoricalDataRecord.periodo,
            func.count(HistoricalDataRecord.id).label('cant'),
            func.sum(HistoricalDataRecord.monto).label('suma')
        ).group_by(HistoricalDataRecord.empresa, HistoricalDataRecord.periodo).all()

        dict_hist_before = {
            (r.empresa, r.periodo): (r.cant, float(r.suma) if r.suma is not None else 0.0)
            for r in snapshot_hist_before
        }

        # 4. Simular la carga y guardado de un P&L de prueba para un período nuevo (2026-06)
        empresa_test = "Pacifico SpA"
        periodo_test = "2026-06-TEST"

        df_test_pl = pd.DataFrame({
            'N° de cuenta': ['410101', '510101'],
            'Nombre de la cuenta': ['Ventas Test', 'Costos Test'],
            'Ingresos de actividades ordinarias': [1500000.0, 0.0],
            'Costo de ventas': [0.0, -900000.0]
        })

        # Guardar en DB
        PlCuboDB.save_pl_cubo(empresa_test, periodo_test, df_test_pl)

        try:
            # 5. Snapshot POST-guardado de los períodos preexistentes
            snapshot_pl_after = self.db.query(
                PlRecordDim.empresa,
                PlRecordDim.periodo,
                func.count(PlRecordDim.id).label('cant'),
                func.sum(PlRecordDim.monto).label('suma')
            ).filter(PlRecordDim.periodo != periodo_test).group_by(PlRecordDim.empresa, PlRecordDim.periodo).all()

            dict_pl_after = {
                (r.empresa, r.periodo): (r.cant, float(r.suma) if r.suma is not None else 0.0)
                for r in snapshot_pl_after
            }

            snapshot_tb_after = self.db.query(
                TrialBalanceRecord.empresa,
                TrialBalanceRecord.periodo,
                func.count(TrialBalanceRecord.id).label('cant'),
                func.sum(TrialBalanceRecord.saldo_final).label('suma')
            ).group_by(TrialBalanceRecord.empresa, TrialBalanceRecord.periodo).all()

            dict_tb_after = {
                (r.empresa, r.periodo): (r.cant, float(r.suma) if r.suma is not None else 0.0)
                for r in snapshot_tb_after
            }

            snapshot_hist_after = self.db.query(
                HistoricalDataRecord.empresa,
                HistoricalDataRecord.periodo,
                func.count(HistoricalDataRecord.id).label('cant'),
                func.sum(HistoricalDataRecord.monto).label('suma')
            ).group_by(HistoricalDataRecord.empresa, HistoricalDataRecord.periodo).all()

            dict_hist_after = {
                (r.empresa, r.periodo): (r.cant, float(r.suma) if r.suma is not None else 0.0)
                for r in snapshot_hist_after
            }

            # VERIFICACIONES RIGUROSAS
            self.assertEqual(dict_pl_before, dict_pl_after, "Los registros previos de P&L sufrieron modificaciones inesperadas.")
            self.assertEqual(dict_tb_before, dict_tb_after, "Los registros previos de Trial Balance sufrieron modificaciones inesperadas.")
            self.assertEqual(dict_hist_before, dict_hist_after, "Los registros históricos previos sufrieron modificaciones inesperadas.")

            # Verificar precisión de los datos guardados en el nuevo período
            df_retrieved = PlCuboDB.get_pl_cubo(empresa_test, periodo_test)
            self.assertIsNotNone(df_retrieved, "No se pudieron recuperar los datos del período de prueba.")
            self.assertFalse(df_retrieved.empty, "El DataFrame recuperado está vacío.")
            
            sum_retrieved = PlCuboDB.get_pl_cubo_total_sum(empresa_test, periodo_test)
            self.assertAlmostEqual(sum_retrieved, 600000.0, places=2, msg="La suma total recuperada no coincide con los montos guardados.")

        finally:
            # Limpiar datos de prueba creados
            self.db.query(PlRecordDim).filter(
                PlRecordDim.empresa == empresa_test,
                PlRecordDim.periodo == periodo_test
            ).delete()
            self.db.commit()

if __name__ == '__main__':
    unittest.main()
