import pandas as pd
import os

def create_mock_plan():
    # Creamos un plan de cuentas maestro simulado
    # Omite a propósito alguna cuenta del mock_trial_balance para forzar la alerta de auditoría si fuera necesario,
    # O coloquemos todas para que pase, y que el usuario suba otro.
    # Colocaré todas menos "210201" para probar la alerta de Auditoría (Cuenta huérfana).
    data = {
        "Cuenta": ["110101", "110201", "120101", "130101", "210101", "310101", "410101", "510101"],
        "Name": ["Caja", "Banco", "CXC", "Inventarios", "Proveedores", "Capital", "Ingresos", "Gastos"]
    }
            
    df = pd.DataFrame(data)
    
    # Root dir of execution is usually Software_IFRS16_Pro or Informes_Pro
    # We will use absolute/relative safely.
    out_path = r"c:\Users\cfsaa\OneDrive\Desktop\Informes_Pro\data\mock_plan_cuentas.xlsx"
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    df.to_excel(out_path, index=False)
    print(f"Mock Plan de Cuentas creado en {out_path}")
    
if __name__ == "__main__":
    create_mock_plan()
