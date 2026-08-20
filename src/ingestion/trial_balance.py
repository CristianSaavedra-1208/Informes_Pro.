import pandas as pd

class TrialBalanceIngestor:
    def __init__(self, file_path):
        self.file_path = file_path
        self.df = None

    def load_and_standardize(self):
        """Importa balances de comprobación en formato flexible y validado (Req 1)."""
        try:
            # Leer excel (openpyxl)
            raw_df = pd.read_excel(self.file_path, engine='openpyxl')
            
            # Limpiar cabeceras
            raw_df.columns = [str(c).strip() for c in raw_df.columns]
            cols_lower = [c.lower() for c in raw_df.columns]
            
            # Búsqueda flexible de columnas críticas
            cuenta_col = next((raw_df.columns[i] for i, c in enumerate(cols_lower) if "cuenta" in c), None)
            desc_col = next((raw_df.columns[i] for i, c in enumerate(cols_lower) if "nombre" in c or "descrip" in c or "detalle" in c), None)
            
            # Columnas de saldo final / saldo de cierre
            saldo_final_col = next((raw_df.columns[i] for i, c in enumerate(cols_lower) if "final" in c or "cierre" in c or "dr/cr" in c or "saldo_final" in c), None)
            if not saldo_final_col:
                # Fallback a buscar simplemente "saldo"
                saldo_final_col = next((raw_df.columns[i] for i, c in enumerate(cols_lower) if "saldo" in c), None)
                
            if not cuenta_col or not desc_col or not saldo_final_col:
                raise ValueError("No se encontraron columnas obligatorias para el Balance de Comprobación (debe contener columnas para Cuenta, Descripción y Saldo).")
                
            # Columnas opcionales de flujos
            saldo_ini_col = next((raw_df.columns[i] for i, c in enumerate(cols_lower) if "inicial" in c or "apertura" in c or "comienzo" in c), None)
            debitos_col = next((raw_df.columns[i] for i, c in enumerate(cols_lower) if "debit" in c or "debe" in c or "cargo" in c or "adicion" in c), None)
            creditos_col = next((raw_df.columns[i] for i, c in enumerate(cols_lower) if "credit" in c or "haber" in c or "abono" in c or "retiro" in c), None)
            
            # Validación de cuadratura (la suma del saldo final debe ser 0)
            saldo_final_series = pd.to_numeric(raw_df[saldo_final_col], errors='coerce').fillna(0.0)
            saldo_sum = saldo_final_series.sum()
            if abs(saldo_sum) > 0.05: # Tolerancia flotante
                raise ValueError(f"🚨 El balance de comprobación cargado no está cuadrado. La columna de saldo suma {saldo_sum:,.2f} en vez de 0.00")
            
            # Crear DataFrame estandarizado
            self.df = pd.DataFrame()
            self.df['cuenta_id'] = raw_df[cuenta_col].astype(str).str.strip()
            self.df['descripcion'] = raw_df[desc_col].fillna("").astype(str).str.strip()
            
            self.df['saldo_inicial'] = pd.to_numeric(raw_df[saldo_ini_col], errors='coerce').fillna(0.0) if saldo_ini_col else 0.0
            self.df['debitos'] = pd.to_numeric(raw_df[debitos_col], errors='coerce').fillna(0.0) if debitos_col else 0.0
            self.df['creditos'] = pd.to_numeric(raw_df[creditos_col], errors='coerce').fillna(0.0) if creditos_col else 0.0
            self.df['saldo_final'] = saldo_final_series
            
            # Limpieza de filas vacías
            self.df = self.df[self.df['cuenta_id'].notna() & (self.df['cuenta_id'] != "") & (self.df['cuenta_id'] != "nan")].reset_index(drop=True)
            
            return self.df
            
        except Exception as e:
            raise ValueError(str(e))
