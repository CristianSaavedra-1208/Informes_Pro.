import pandas as pd

class MappingEngine:
    def __init__(self, mapping_dict):
        """
        Req 2: Motor de mapping de cuentas hacia categoría financiera universal.
        mapping_dict format: {'1101': 'Efectivo', '1201': 'Cuentas x Cobrar'}
        """
        self.mapping = mapping_dict
        
    def apply_mapping(self, trial_balance_df):
        """Aplica el mapeo al dataframe ingresado por TrialBalanceIngestor."""
        
        # Copiamos para no mutar el original
        mapped_df = trial_balance_df.copy()
        
        # Mapeo Vectorizado mediante Pandas O(1) tiempo
        mapped_df['categoria_financiera'] = mapped_df['cuenta_id'].astype(str).map(self.mapping)
        
        return mapped_df
        
    def detect_unmapped(self, mapped_df):
        """Devuelve un dataframe de las cuentas sin mapear."""
        unmapped = mapped_df[mapped_df['categoria_financiera'].isna()]
        return unmapped
