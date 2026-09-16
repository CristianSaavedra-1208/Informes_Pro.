import os
import sys
import unittest
import pandas as pd
import numpy as np
from pandas.testing import assert_frame_equal

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.core.pl_cubo_processor import process_odoo_cubo, normalize_text

class TestPlProcessorEquivalence(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.excel_cubo = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "Cubo Odoo Julio 2026_v2_27.08.xlsx")
        cls.map_pl_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data", "empresas", "Pacifico Cable SpA", "map_pl.xlsx")
        
        if os.path.exists(cls.excel_cubo) and os.path.exists(cls.map_pl_path):
            cls.cubo_df = pd.read_excel(cls.excel_cubo)
            cls.map_pl_df = pd.read_excel(cls.map_pl_path)
        else:
            cls.cubo_df = None
            cls.map_pl_df = None

    def test_real_data_processing_and_totals(self):
        """Verifica que el procesamiento con el Cubo Odoo real sea matemáticamente idéntico y consistente."""
        if self.cubo_df is None or self.map_pl_df is None:
            self.skipTest("Archivos de prueba reales no encontrados")

        pivot_df, audit_df = process_odoo_cubo(
            self.cubo_df,
            year=2026,
            month=7,
            map_pl_df=self.map_pl_df,
            return_audit_log=True
        )

        # 1. Verificar columnas esperadas
        self.assertIn("N° de cuenta", pivot_df.columns)
        self.assertIn("Nombre de la cuenta", pivot_df.columns)
        self.assertIn("Ingresos de actividades ordinarias", pivot_df.columns)
        self.assertIn("Costo de ventas", pivot_df.columns)

        # 2. Verificar suma total de montos
        total_sum = pivot_df.iloc[:, 2:].sum().sum()
        # Debe coincidir exactamente con el total de 11,135,647,802.00
        self.assertAlmostEqual(total_sum, 11135647802.0, delta=1.0)
        self.assertEqual(len(pivot_df), 135)

    def test_synthetic_edge_cases(self):
        """Verifica casos borde sintéticos: overrides, cuentas sin mapear, Odoo categorías y auditoría."""
        synthetic_cubo = pd.DataFrame({
            'N° de cuenta': ['3105301', '3105302', '999999', '888888', '3105703'],
            'Nombre de la cuenta': ['Admin Override', 'Costo Override', 'Sin Map Odoo OK', 'Sin Map Sin Odoo', 'Deprec Dual'],
            'informe_ee_rr': ['Ventas', 'Costo', 'Ingresos de actividades ordinarias', '', 'Costo de ventas'],
            'Importe en moneda del informe': [100.0, 200.0, 300.0, 400.0, 500.0],
            'Año': [2026, 2026, 2026, 2026, 2026],
            'Mes': [7, 7, 7, 7, 7]
        })

        synthetic_map = pd.DataFrame({
            'Cuenta': ['3105703'],
            'Depreciación operacional': ['X'],
            'Depreciación y amortizaciones': ['X']
        })

        pivot_df, audit_df = process_odoo_cubo(
            synthetic_cubo,
            year=2026,
            month=7,
            map_pl_df=synthetic_map,
            return_audit_log=True
        )

        # Verificar overrides
        row_3105301 = pivot_df[pivot_df['N° de cuenta'] == '3105301']
        self.assertEqual(row_3105301['Gastos de administración'].values[0], 100.0)

        row_3105302 = pivot_df[pivot_df['N° de cuenta'] == '3105302']
        self.assertEqual(row_3105302['Costo de ventas'].values[0], 200.0)

        # Verificar mapeo directo desde Odoo sin map_pl
        row_999999 = pivot_df[pivot_df['N° de cuenta'] == '999999']
        self.assertEqual(row_999999['Ingresos de actividades ordinarias'].values[0], 300.0)

        # Verificar fallback sin Odoo ni map_pl
        row_888888 = pivot_df[pivot_df['N° de cuenta'] == '888888']
        self.assertEqual(row_888888['Otros egresos por función'].values[0], 400.0)

        # Verificar desempate de depreciación operacional cuando Odoo dice Costo de ventas
        row_deprec = pivot_df[pivot_df['N° de cuenta'] == '3105703']
        self.assertEqual(row_deprec['Depreciación operacional'].values[0], 500.0)

        # Verificar auditoría para la fila 888888 (vacía/no reconocida en Odoo)
        cuentas_audit = audit_df['N° de Cuenta'].tolist()
        self.assertIn('888888', cuentas_audit)

if __name__ == "__main__":
    unittest.main()
