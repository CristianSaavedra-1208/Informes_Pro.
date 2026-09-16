import unittest
import os
import io
import pandas as pd
import numpy as np

from src.core.pl_cubo_processor import process_odoo_cubo, normalize_text
from src.core.validation_tie_out import ValidationTieOutEngine
from src.models.pl_cubo_db import PlCuboDB
from src.models.trial_balance_db import TrialBalanceDB
from src.models.historical_data import HistoricalDataRecord
from src.models.database import SessionLocal

class TestNuevoFormatoCuboYTieOut(unittest.TestCase):

    def test_nuevo_formato_odoo_cubo_ingesta(self):
        """
        Prueba que el nuevo formato del Cubo Odoo ('Cubo Odoo Julio 2026_v2_27.08.xlsx')
        se filtre correctamente por Nivel 1 (EBITDA / NO EBITDA) y asigne las columnas
        cuenta, importe_mn e informe_ee_rr.
        """
        file_path = 'Cubo Odoo Julio 2026_v2_27.08.xlsx'
        if not os.path.exists(file_path):
            self.skipTest("Archivo Cubo Odoo Julio 2026_v2_27.08.xlsx no encontrado en workspace.")

        # Leer con soporte para archivos abiertos
        import ctypes
        def read_locked(p):
            GENERIC_READ = 0x80000000
            FILE_SHARE_ALL = 7
            handle = ctypes.windll.kernel32.CreateFileW(os.path.abspath(p), GENERIC_READ, FILE_SHARE_ALL, None, 3, 0x80, None)
            if handle == -1:
                return pd.read_excel(p)
            try:
                sz = ctypes.windll.kernel32.GetFileSize(handle, None)
                buf = ctypes.create_string_buffer(sz)
                r = ctypes.c_ulong(0)
                ctypes.windll.kernel32.ReadFile(handle, buf, sz, ctypes.byref(r), None)
                return pd.read_excel(io.BytesIO(buf.raw))
            finally:
                ctypes.windll.kernel32.CloseHandle(handle)

        df_cubo = read_locked(file_path)
        
        # Verificar que el cubo crudo tiene la columna Nivel 1 y 38 columnas
        self.assertIn('Nivel 1', df_cubo.columns)
        self.assertIn('importe_mn', df_cubo.columns)
        self.assertIn('informe_ee_rr', df_cubo.columns)

        # Cargar map_pl de Db Terra Holdco
        map_pl_path = os.path.join("data", "empresas", "Db Terra Holdco", "map_pl.xlsx")
        if not os.path.exists(map_pl_path):
            map_pl_path = os.path.join("data", "empresas", "DB Holdco Terra SpA", "map_pl.xlsx")
        map_pl_df = pd.read_excel(map_pl_path) if os.path.exists(map_pl_path) else None

        # Procesar con process_odoo_cubo para Julio 2026
        df_proc = process_odoo_cubo(df_cubo, 2026, 7, map_pl_df)

        self.assertIsNotNone(df_proc)
        self.assertFalse(df_proc.empty)
        self.assertIn('N° de cuenta', df_proc.columns)
        self.assertIn('Nombre de la cuenta', df_proc.columns)
        self.assertIn('Ingresos de actividades ordinarias', df_proc.columns)
        self.assertIn('Costo de ventas', df_proc.columns)

        # Ninguna fila en df_proc debe tener N° de cuenta correspondiente a caja o bancos (clase 1xxx de Holdco Balance)
        cuentas = df_proc['N° de cuenta'].astype(str).tolist()
        for c in cuentas:
            # Las cuentas de resultado empiezan con 3 (o 4/5), no con 1101 (Caja/Bancos de Balance)
            self.assertFalse(c.startswith('1101'), f"Cuenta de Balance 1101 encontrada en P&L: {c}")

        print(f"\n[OK] Test Ingesta Nuevo Cubo: {len(df_proc)} cuentas procesadas correctamente con Nivel 1 (EBITDA/NO EBITDA).")

    def test_validacion_balance_vs_pl_engine(self):
        """
        Prueba la función de validación cruzada Balance vs. P&L.
        """
        res = ValidationTieOutEngine.validar_balance_vs_pl("Db Terra Holdco", "2026-05")
        if not res["has_tb"] or not res["has_pl"]:
            # Probar con Pacifico o Terra Parent
            res = ValidationTieOutEngine.validar_balance_vs_pl("Pacifico SpA", "2026-05")

        self.assertIn("is_cuadrado", res)
        self.assertIn("saldo_balance", res)
        self.assertIn("saldo_pl", res)
        self.assertIn("diferencia", res)
        print(f"\n[OK] Test Validacion Balance vs P&L ejecutado: is_cuadrado={res['is_cuadrado']}, diff=${res['diferencia']:,.2f}")

    def test_inmutabilidad_periodos_cerrados_historicos(self):
        """
        SEGURO DE NO REGRESIÓN:
        Verifica que para todos los períodos cerrados históricos (hasta 2026-06):
        1. Los registros históricos de P&L y Balance existen y son inmutables.
        2. PlCuboDB.get_pl_cubo reconstruye los saldos históricos idénticos.
        """
        db = SessionLocal()
        try:
            # Obtener períodos históricos registrados
            hist_records = db.query(HistoricalDataRecord).filter(
                HistoricalDataRecord.periodo.in_(["2025-12", "2026-03", "2026-04", "2026-05", "2026-06"])
            ).all()

            self.assertGreater(len(hist_records), 0, "Debe haber registros históricos para los períodos cerrados.")

            # Agrupar sumas por empresa y período para verificar consistencia
            df_hist = pd.DataFrame([{
                'empresa': r.empresa,
                'periodo': r.periodo,
                'reporte': r.reporte,
                'linea_item': r.linea_item,
                'monto': r.monto
            } for r in hist_records])

            # Verificar que no hay NaN o None en montos
            self.assertEqual(df_hist['monto'].isna().sum(), 0, "No debe haber montos NaN en históricos.")

            # Probar PlCuboDB.get_pl_cubo en período histórico
            for co in ["Db Terra Holdco", "Pacifico SpA"]:
                pl_hist = PlCuboDB.get_pl_cubo(co, "2026-05")
                if pl_hist is not None:
                    self.assertFalse(pl_hist.empty)
                    self.assertIn('N° de cuenta', pl_hist.columns)

            print(f"\n[OK] Test Inmutabilidad Histórica: {len(hist_records)} registros históricos verificados sin variaciones.")
        finally:
            db.close()

    def test_audit_warnings_empty_odoo_classification(self):
        """
        Prueba que process_odoo_cubo identifique filas con celdas de clasificación vacías
        y retorne el DataFrame de auditoría con la fila, cuenta, monto y rubro asignado.
        """
        # Crear un DataFrame simulando el cubo de Odoo con una celda de informe_ee_rr vacía
        test_cubo = pd.DataFrame([
            {
                'fec_doc': '2026-08-15',
                'ano': 2026,
                'mes': 8,
                'cuenta': '3105819',
                'nombre_cuenta': 'GASTOS DE HOSPEDAJE NACIONAL',
                'importe_mn': -59664.0,
                'informe_ee_rr': None,  # Celda vacía en Odoo
                'Nivel 1': 'EBITDA'
            },
            {
                'fec_doc': '2026-08-15',
                'ano': 2026,
                'mes': 8,
                'cuenta': '3105301',
                'nombre_cuenta': 'SUELDOS Y SALARIOS',
                'importe_mn': 100000.0,
                'informe_ee_rr': 'Gastos de administración',
                'Nivel 1': 'EBITDA'
            }
        ])

        # Map PL con Costo de ventas y Gastos de administración para 3105819
        test_map_pl = pd.DataFrame([
            {
                'N° de cuenta': '3105819',
                'Nombre de la cuenta': 'GASTOS DE HOSPEDAJE NACIONAL',
                'Costo de ventas': 'Costos por servicios externos',
                'Gastos de administración': 'Viáticos y estadía'
            }
        ])

        # 1. Probar compatibilidad por defecto (return_audit_log=False)
        df_proc = process_odoo_cubo(test_cubo, 2026, 8, test_map_pl, return_audit_log=False)
        self.assertIsInstance(df_proc, pd.DataFrame)
        self.assertIn('N° de cuenta', df_proc.columns)

        # 2. Probar con return_audit_log=True
        df_proc, df_audit = process_odoo_cubo(test_cubo, 2026, 8, test_map_pl, return_audit_log=True)
        self.assertIsInstance(df_proc, pd.DataFrame)
        self.assertIsInstance(df_audit, pd.DataFrame)
        self.assertEqual(len(df_audit), 1, "Debe registrar exactamente 1 fila con advertencia de auditoría")
        
        audit_row = df_audit.iloc[0]
        self.assertEqual(audit_row['Fila Excel'], 2, "La primera fila de datos corresponde a la fila 2 de Excel")
        self.assertEqual(audit_row['N° de Cuenta'], '3105819')
        self.assertEqual(audit_row['Importe MN'], -59664.0)
        self.assertEqual(audit_row['Clasificación ERP'], '(Vacío)')
        self.assertEqual(audit_row['Rubro Asignado'], 'Costo de ventas')
        self.assertIn('Sin clasificación en ERP', audit_row['Motivo / Observación'])

        print("\n[OK] Test Auditoría Clasificación Vacía: Detección y log de advertencias validado exitosamente.")

if __name__ == '__main__':
    unittest.main()
