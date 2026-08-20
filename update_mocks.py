import pandas as pd
import os

data_dir = r"c:\Users\cfsaa\OneDrive\Desktop\Informes_Pro\data"
os.makedirs(data_dir, exist_ok=True)

# 4. Mapeo Balance (New: "nota 1", "nota 2")
pd.DataFrame({
    'N° de Cuenta': ['110101', '210101'], 
    'Nombre cuenta': ['Caja', 'Proveedores'], 
    'Clasificación balance': ['Activo Corriente', 'Pasivo Corriente'],
    'nota 1': ['', ''],
    'nota 2': ['', '']
}).to_excel(os.path.join(data_dir, "mock_map_balance.xlsx"), index=False)

# 5. Mapeo P&L (New columns for detailed EERR)
pd.DataFrame({
    'N° de Cuenta': ['410101', '510101'], 
    'Nombre de la cuenta': ['Ventas', 'Costo Ventas'], 
    'Clasificacion estado de resultados': ['Ingreso Actividades Ordinarias', 'Costo Actividades Ordinarias'], 
    'Nota Ingresos de la operación': ['X', ''],
    'Nota Costo Ventas': ['', 'X'],
    'Nota Gtos Administracion': ['', ''],
    'Nota Gtos Financieros': ['', '']
}).to_excel(os.path.join(data_dir, "mock_map_pl.xlsx"), index=False)

print(f"Archivos mock de mapeo actualizados exitosamente en: {data_dir}")

# 6. Cubo P&L (New columns for analytical P&L)
pd.DataFrame({
    'Periodo': ['2025-12', '2025-12'],
    'N° de Cuenta': ['410101', '510101'], 
    'Nombre de la cuenta': ['Ventas Corporativas', 'Costo Insumos'], 
    'Centro de Costo': ['CC01-Ventas', 'CC02-Produccion'],
    'Unidad de Negocio': ['B2B', 'B2B'],
    'Monto': [15000000, -8500000]
}).to_excel(os.path.join(data_dir, "mock_carga_pl_cubo.xlsx"), index=False)

print(f"Plantillas de carga P&L Cubo generadas.")
