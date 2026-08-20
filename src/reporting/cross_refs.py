class CrossReferenceGenerator:
    def __init__(self, eeff_to_note_mapping):
        """
        Req 10: Genera el índice de Notas dinámico para el Word y las Tablas EEFF.
        eeff_to_note_mapping: {'Inventarios': '10', 'Efectivo': '3'}
        """
        self.mapping = eeff_to_note_mapping
        
    def attach_references(self, financial_statements):
        """
        Añade la etiqueta dinámica (Ej: 'ver Nota 9') a los rubros para la vista o exportación.
        """
        statements_with_refs = []
        for item, total in financial_statements.items():
            ref_num = self.mapping.get(item, "")
            label = f"(Nota {ref_num})" if ref_num else ""
            
            statements_with_refs.append({
                "rubro": item,
                "saldo": total,
                "referencia": label,
                "display_name": f"{item} {label}".strip()
            })
            
        return statements_with_refs
