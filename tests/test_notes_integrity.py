import unittest
import os
import openpyxl
import pandas as pd
from src.models.trial_balance_db import TrialBalanceDB
from src.reporting.note_generator import NoteGenerator
from src.reporting.notes import NOTE_REGISTRY
from src.ui_pages.informes_y_notas import load_all_entity_contexts

class TestNotesIntegrity(unittest.TestCase):
    """
    Suite de pruebas automatizadas para verificar que:
    1. Existe la Plantilla Maestra Global 'Plantilla de notas_v1.xlsx' en la raíz.
    2. Contiene las pestañas requeridas para el renderizado de notas contables.
    3. Carga los mapeos (map_balance / map_pl) y procesa todas las notas para Pacífico SpA.
    4. Cada nota genera un binario Excel válido y completo sin excepciones.
    """

    @classmethod
    def setUpClass(cls):
        cls.global_template = "Plantilla de notas_v1.xlsx"
        cls.empresa_test = "Pacifico SpA"
        cls.periodo_actual = "2026-03"
        cls.periodo_comp = "2025-12"

    def test_01_global_template_exists_in_root(self):
        """Verifica que la plantilla de notas global exista en la raíz del proyecto."""
        self.assertTrue(
            os.path.exists(self.global_template),
            f"❌ No se encontró la plantilla maestra global '{self.global_template}' en la raíz."
        )

    def test_02_note_sheets_exist_in_global_template(self):
        """Verifica que las pestañas clave de notas existan en la plantilla global."""
        wb = openpyxl.load_workbook(self.global_template, read_only=True)
        sheets = wb.sheetnames
        wb.close()
        
        expected_sheets = ["Efectivo", "Deudores", "Inventarios", "Intangibles", "Activo Fijo", "Gtos Adm", "Segmentos"]
        for sheet in expected_sheets:
            self.assertIn(sheet, sheets, f"❌ Falta la pestaña '{sheet}' en la plantilla global de notas.")

    def test_03_generate_pacifico_notes_integrity(self):
        """
        Ejecuta la generación de notas con la plantilla global para Pacífico SpA 
        y verifica la validez del libro Excel generado para cada nota.
        """
        # Cargar mapeos de Pacífico SpA / Global
        empresa_path = os.path.join("data", "empresas", self.empresa_test)
        map_bal_path = os.path.join(empresa_path, "map_balance.xlsx")
        map_pl_path = os.path.join(empresa_path, "map_pl.xlsx")

        if not os.path.exists(map_bal_path):
            map_bal_path = "map_balance.xlsx"
        if not os.path.exists(map_pl_path):
            map_pl_path = "map_pl.xlsx"

        map_balance_df = pd.read_excel(map_bal_path) if os.path.exists(map_bal_path) else None
        map_pl_df = pd.read_excel(map_pl_path) if os.path.exists(map_pl_path) else None

        self.assertIsNotNone(map_balance_df, "❌ No se pudo cargar el mapa de balance para el test.")

        # Cargar contextos de datos para Pacífico SpA
        entity_contexts = load_all_entity_contexts(
            active_entity=self.empresa_test,
            periodo_actual=self.periodo_actual,
            periodo_comp=self.periodo_comp,
            map_balance_df=map_balance_df,
            map_pl_df=map_pl_df
        )

        self.assertIn(self.empresa_test, entity_contexts, f"❌ No se pudo cargar el contexto para {self.empresa_test}")

        engine = NoteGenerator(self.global_template)

        # Probar notas clave configuradas
        notas_probadas = 0
        for code, info in NOTE_REGISTRY.items():
            sheets = info.get('sheets', [])
            if not sheets:
                continue

            try:
                res_bytes = engine.generate(
                    sheet_names=sheets,
                    entity_contexts=entity_contexts,
                    active_entity_name=self.empresa_test,
                    is_consolidated=False,
                    scale_factor=1.0,
                    periodo_actual_str=self.periodo_actual,
                    periodo_comp_str=self.periodo_comp,
                    map_balance_df=map_balance_df,
                    map_pl_df=map_pl_df
                )
                output_bytes = res_bytes.getvalue() if hasattr(res_bytes, 'getvalue') else res_bytes
                self.assertIsNotNone(output_bytes, f"❌ Falló la generación de bytes para la nota {code}")
                self.assertGreater(len(output_bytes), 1000, f"❌ El archivo generado para la nota {code} es demasiado pequeño o inválido")
                notas_probadas += 1
            except Exception as e:
                self.fail(f"❌ Error al generar la nota {code} ({info.get('label')}): {str(e)}")

        self.assertGreater(notas_probadas, 0, "❌ No se probó ninguna nota.")

if __name__ == '__main__':
    unittest.main()
