class TieOutEngine:
    @staticmethod
    def verify_tie_out(financial_statements, notes_data, eeff_nota_mapping, tolerance=0.1):
        """
        Req 7: Compara las cifras de EEFF con los totales de las Notas correspondientes.
        eeff_nota_mapping vincula nombre rubro EEFF -> nombre llave Nota.
        """
        inconsistencies = []
        tie_out_report = {}
        
        for eeff_item, eeff_total in financial_statements.items():
            note_key = eeff_nota_mapping.get(eeff_item)
            
            if note_key and note_key in notes_data:
                note_total = notes_data[note_key]['total']
                diff = abs(abs(eeff_total) - abs(note_total))
                
                status = "OK" if diff <= tolerance else "ERROR"
                tie_out_report[eeff_item] = {
                    "eeff_total": eeff_total,
                    "nota_total": note_total,
                    "diferencia": diff,
                    "status": status
                }
                
                if status == "ERROR":
                    inconsistencies.append(f"Tie-Out Fallido en '{eeff_item}'. EEFF: {eeff_total} != Nota '{note_key}': {note_total} (Diferencia Absoluta: {diff})")
                    
        return {
            "is_tied_out": len(inconsistencies) == 0,
            "report": tie_out_report,
            "errors": inconsistencies
        }
