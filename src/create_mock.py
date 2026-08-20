import pandas as pd
import os

def generate_mock_trial_balance(output_path):
    data = [
        {"N° de Cuenta": "110101", "Nombre de la cuenta": "Caja Fuerte Central", "Saldo DR/CR": 10000},
        {"N° de Cuenta": "110201", "Nombre de la cuenta": "Banco de Chile CTA 123", "Saldo DR/CR": 45000},
        {"N° de Cuenta": "120101", "Nombre de la cuenta": "Facturas por Cobrar Clientes", "Saldo DR/CR": 20000},
        {"N° de Cuenta": "130101", "Nombre de la cuenta": "Inventario Materia Prima", "Saldo DR/CR": 15000},
        {"N° de Cuenta": "210101", "Nombre de la cuenta": "Proveedores Nacionales", "Saldo DR/CR": -30000},
        {"N° de Cuenta": "210201", "Nombre de la cuenta": "Impuestos por Pagar (IVA)", "Saldo DR/CR": -10000},
        {"N° de Cuenta": "310101", "Nombre de la cuenta": "Capital Social Emitido", "Saldo DR/CR": -50000},
    ]
    df = pd.DataFrame(data)
    
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    df.to_excel(output_path, index=False, engine='openpyxl')
    print(f"[OK] Excel 'Trial Balance' generado en: {output_path}")

if __name__ == "__main__":
    generate_mock_trial_balance(r"c:\Users\cfsaa\OneDrive\Desktop\Informes_Pro\data\mock_trial_balance.xlsx")
