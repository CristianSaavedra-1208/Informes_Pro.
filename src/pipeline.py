import os
import pprint
from sys import path

path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.ingestion.trial_balance import TrialBalanceIngestor
from src.core.mapping import MappingEngine
from src.core.rules import AccountingRulesEngine
from src.core.validation import ValidationEngine
from src.reporting.notes import NotesOrchestrator
from src.core.tie_out import TieOutEngine

def run_pipeline():
    excel_path = r"c:\Users\cfsaa\OneDrive\Desktop\Informes_Pro\data\mock_trial_balance.xlsx"
    
    print("\n================================================")
    print("      INICIANDO PIPELINE DE INFORMES PRO        ")
    print("================================================\n")
    
    print("--- 1. INGESTIÓN DE DATOS ---")
    ingestor = TrialBalanceIngestor(excel_path)
    tb_df = ingestor.load_and_standardize()
    print(f"[OK] Trial Balance Excel cargado: detectadas {len(tb_df)} cuentas.")
    
    print("\n--- 2. MAPEO DE CUENTAS (IFRS) ---")
    mapping_dict = {
        '110101': 'Efectivo',
        '110201': 'Efectivo',
        '120101': 'Cuentas_x_Cobrar',
        '130101': 'Inventarios',
        '210101': 'Cuentas_x_Pagar',
        '210201': 'Impuestos_x_Pagar',
        '310101': 'Capital'
    }
    mapper = MappingEngine(mapping_dict)
    mapped_df = mapper.apply_mapping(tb_df)
    unmapped = mapper.detect_unmapped(mapped_df)
    print(f"[OK] Cuentas mapeadas correctamente. {len(unmapped)} sin mapear (NaN).")
    
    print("\n--- 3. MOTOR DE REGLAS (ESTADOS FINANCIEROS) ---")
    rules_config = {
        'Activo Corriente: Efectivo': ['Efectivo'],
        'Activo Corriente: Cuentas x Cobrar': ['Cuentas_x_Cobrar'],
        'Activo Corriente: Inventarios': ['Inventarios'],
        'Total Activos': ['Efectivo', 'Cuentas_x_Cobrar', 'Inventarios'],
        'Total Pasivos': ['Cuentas_x_Pagar', 'Impuestos_x_Pagar'],
        'Total Patrimonio': ['Capital']
    }
    rules_engine = AccountingRulesEngine(rules_config)
    statements = rules_engine.generate_statement(mapped_df)
    for k, v in statements.items():
        print(f"  > {k}: {v:,.2f}")
    
    print("\n--- 4. VALIDACIÓN DE ECUACIÓN CONTABLE (A = P + Pt) ---")
    val_report = ValidationEngine.validate_accounting_equation(mapped_df, statements)
    print(f"[OK] Estatus de Cuadratura: {'VÁLIDO (Cuadrado)' if val_report['is_valid'] else 'DESCUADRADO'}")
    if not val_report['is_valid']:
        print("  [X] Errores:", val_report['errors'])
        
    print("\n--- 5. GENERADOR ORQUESTADOR DE NOTAS ---")
    notes_mapping = {
        'Nota_3_Efectivo': ['Efectivo'],
        'Nota_4_Inventarios': ['Inventarios']
    }
    notes_orch = NotesOrchestrator(notes_mapping)
    notes_data = notes_orch.generate_note_tables(mapped_df)
    for k, v in notes_data.items():
        print(f"  > {k}: Total -> {v['total']:,.2f} ({len(v['detalle'])} cuentas agrupadas)")
    
    print("\n--- 6. MOTOR TIE-OUT (AUDITORÍA AUTOMÁTICA) ---")
    eeff_to_notas_mapping = {
        'Activo Corriente: Efectivo': 'Nota_3_Efectivo',
        'Activo Corriente: Inventarios': 'Nota_4_Inventarios'
    }
    tieout_report = TieOutEngine.verify_tie_out(statements, notes_data, eeff_to_notas_mapping)
    print(f"[OK] Resultado Tie-Out: {'CUADRATURA PERFECTA' if tieout_report['is_tied_out'] else 'INCONSISTENCIA'}")
    for item, metrics in tieout_report['report'].items():
        if metrics['diferencia'] == 0:
            print(f"  > {item}: OK MATCH ({metrics['eeff_total']:,.2f})")
        else:
            print(f"  > {item}: ERROR (Diferencias encontradas)")
            
    print("\n================================================")
    print("           COMPLETADO EXITOSAMENTE              ")
    print("================================================\n")

if __name__ == "__main__":
    run_pipeline()
