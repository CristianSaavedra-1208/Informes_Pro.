class AccountingRulesEngine:
    def __init__(self, rules_config: dict):
        """
        Req 3: Motor de reglas contables paramétrico.
        rules_config = {
            'Activos Corrientes': ['Efectivo y Equivalentes', 'Cuentas por Cobrar', 'Inventarios'],
            'Pasivos Corrientes': ['Cuentas por Pagar', 'Impuestos por Pagar']
        }
        """
        self.rules = rules_config

    def generate_statement(self, mapped_tb_df):
        """
        Req 4: Generación automática de EEFF.
        Calcula las sumas basándose en el DataFrame que ya pasó por el MappingEngine.
        """
        # Sumarizamos los saldos finales por categoría
        balances = mapped_tb_df.groupby('categoria_financiera')['saldo_final'].sum().to_dict()
        
        financial_statements = {}
        for rule_name, categories in self.rules.items():
            # Cruza las categorías requeridas por la regla con los saldos existentes
            total = sum(balances.get(cat, 0) for cat in categories)
            financial_statements[rule_name] = total
            
        return financial_statements
