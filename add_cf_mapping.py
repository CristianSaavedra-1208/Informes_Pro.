import pandas as pd
import os

import pandas as pd
import os

bal_map = {
    'Efectivo y efectivo equivalente': 'Saldo inicial de efectivo y equivalentes al efectivo',
    'Otros activos no financieros, corrientes': 'Flujo Operativo - Otros cobros y pagos (Ajuste)',
    'Deudores comerciales y otras cuentas por cobrar, corrientes': 'Cobros procedentes de las ventas de bienes y prestación de servicios',
    'Cuentas comerciales y otras cuentas por pagar, corrientes': 'Pagos a proveedores por el suministro de bienes y servicios',
    'Cuentas por pagar entidades relacionadas, corrientes': 'Otras entradas y (salidas) de dinero', 
    'Inventarios': 'Pagos a proveedores por el suministro de bienes y servicios',
    'Pasivo por impuestos, corrientes': 'Impuestos a las ganancias reembolsados (pagados)',
    'Activo por impuestos, corrientes': 'Impuestos a las ganancias reembolsados (pagados)',
    'Inversion en empresas relacionadas': 'Compra de intangibles', # Fallback
    'Activos intangibles distinto a la plusvalía': 'Compra de intangibles',
    'Plusvalia': 'ELIMINACION - Variación No Monetaria (Plusvalía)',
    'Activo por derechos de uso': 'ELIMINACION - Variación No Monetaria (Derechos Uso)',
    'Propiedades, plantas y equipos': 'Compra de Propiedades, planta y equipo',
    'Activo por impuestos diferidos, no corrientes': 'ELIMINACION - Variación No Monetaria (Impuestos Diferidos)',
    'Otros pasivos financieros corrientes': 'Pagos de préstamos',
    'Otras provisiones, no corrientes': 'ELIMINACION - Variación No Monetaria (Provisiones)',
    'Otros pasivos financieros, no corriente': 'Importes procedentes de préstamos de largo plazo', 
    'Cuentas por pagar entidades relacionadas, no corrientes': 'Importes procedentes de préstamos de largo plazo',
    'Resultados acumulados': 'Otras entradas y (salidas) de dinero',
    'Provisiones por beneficios a los empleados': 'Pagos a y por cuenta de los empleados',
    'Pasivos por derechos de uso, corrientes': 'Pagos de pasivos por arrendamientos financieros',
    'Pasivos por derechos de uso, no corrientes': 'Pagos de pasivos por arrendamientos financieros',
    'Capital emitido': 'Otras entradas y (salidas) de dinero',
    'Aporte por enterar': 'Otras entradas y (salidas) de dinero',
    'Otras reservas': 'Otras entradas y (salidas) de dinero',
    'Gastos de administración': 'Flujo Operativo - Otros cobros y pagos (Ajuste)',
    'Cuentas comerciales y otras cuentas por pagar, no corrientes': 'Pagos a proveedores por el suministro de bienes y servicios',
}

def get_pl_cf_mapping(row, pl_cols):
    for col in pl_cols:
        val = row.get(col)
        if pd.notna(val) and str(val).strip() != '' and str(val).strip() != '0':
            cat = col.lower()
            if 'ingresos de actividades' in cat:
                return 'Cobros procedentes de las ventas de bienes y prestación de servicios'
            elif 'costo' in cat and 'ventas' in cat:
                return 'Pagos a proveedores por el suministro de bienes y servicios'
            elif 'depreciaci' in cat or 'amortizaci' in cat:
                return 'ELIMINACION - Depreciación y Amortización'
            elif 'administraci' in cat:
                return 'Pagos a y por cuenta de los empleados' # Simple proxy, usually separated manual
            elif 'financiero' in cat and 'costo' in cat:
                return 'Intereses pagados'
            elif 'financiero' in cat and 'ingreso' in cat:
                return 'Intereses recibidos'
            elif 'diferencias de cambio' in cat:
                return 'Efectos de la variación en la tasa de cambio sobre el efectivo y equivalentes al efectivo'
            elif 'unidades de reajuste' in cat:
                return 'ELIMINACION - Unidades de Reajuste'
            elif 'impuestos' in cat:
                return 'Impuestos a las ganancias reembolsados (pagados)'
            elif 'otros' in cat:
                return 'Otras entradas y (salidas) de dinero'
    return 'Pendiente Clasificar (EFE)'

def main():
    base_dir = r"data/empresas/Pacifico SpA"
    
    bal_path = os.path.join(base_dir, "map_balance.xlsx")
    if os.path.exists(bal_path):
        df_b = pd.read_excel(bal_path, dtype=str, engine='openpyxl')
        clasif_col = [c for c in df_b.columns if 'balance' in c.lower()][0]
        
        def fallback_match(val):
            val_str = str(val)
            if val_str in bal_map:
                return bal_map[val_str]
            for k, v in bal_map.items():
                if k[:10] in val_str:
                    return v
            return 'Pendiente Clasificar (EFE)'
        
        if 'Clasificación Flujo Efectivo' in df_b.columns:
            df_b.drop(columns=['Clasificación Flujo Efectivo'], inplace=True)
            
        df_b['Clasificación Flujo Efectivo'] = df_b[clasif_col].apply(fallback_match)
        df_b.to_excel(bal_path, index=False)
        print("Balance procesado.")

    pl_path = os.path.join(base_dir, "map_pl.xlsx")
    if os.path.exists(pl_path):
        df_p = pd.read_excel(pl_path, dtype=str, engine='openpyxl')
        pl_check_cols = [c for c in df_p.columns if c not in ['Cuenta', 'Detalle', 'N de Cuenta', 'N° de Cuenta', 'Nombre cuenta'] and 'unnamed' not in c.lower()]
        
        if 'Clasificación Flujo Efectivo' in df_p.columns:
            df_p.drop(columns=['Clasificación Flujo Efectivo'], inplace=True)
            
        df_p['Clasificación Flujo Efectivo'] = df_p.apply(lambda row: get_pl_cf_mapping(row, pl_check_cols), axis=1)
        df_p.to_excel(pl_path, index=False)
        print("PL procesado.")

if __name__ == '__main__':
    main()
