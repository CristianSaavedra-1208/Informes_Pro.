import pandas as pd
import json
import os

class ErpAdapterLogger:
    @staticmethod
    def is_configured(empresa_path):
        settings_path = os.path.join(empresa_path, "erp_settings.json")
        if os.path.exists(settings_path):
            with open(settings_path, 'r') as f:
                data = json.load(f)
                return data.get("configured", False)
        return False

    @staticmethod
    def fetch_trial_balance(empresa_path, ano, mes):
        """
        Stub / Dummy interface para extraer Trial Balance de la API del ERP.
        En producción, usará los datos de erp_settings.json para hacer un requests.get() al endpoint.
        """
        settings_path = os.path.join(empresa_path, "erp_settings.json")
        erp_name = "tu ERP"
        if os.path.exists(settings_path):
            with open(settings_path, 'r') as f:
                data = json.load(f)
                erp_name = data.get("erp", "tu ERP")
                
        # Simulando payload de la API
        # Este payload dummy ya tiene la estructura que el app espera (id_cuenta, descripcion, etc.)
        dummy_payload = {
            "cuenta_id": ["110101", "210101", "310101", "410101", "510101"],
            "cuenta_desc": ["Caja General API", "Cuentas por Pagar API", "Capital Emitido API", "Ingresos por Ventas API", "Costo de Ventas API"],
            "saldo_final": [1500000, -800000, -700000, -1000000, 1000000]
        }
        df = pd.DataFrame(dummy_payload)
        return df, erp_name

    @staticmethod
    def fetch_pl_cubo(empresa_path, ano, mes):
        """
        Stub para cubos P&L analíticos desde el ERP.
        """
        return pd.DataFrame(), "tu ERP"
