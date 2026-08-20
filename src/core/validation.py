class ValidationEngine:
    @staticmethod
    def validate_accounting_equation(mapped_tb_df, statements_dict=None, tolerance=0.1):
        """
        Req 6: Valida reglas contables básicas como A = P + Pt y Trial Balance Sum = 0.
        Retorna reporte de inconsistencias.
        """
        errors = []
        
        # 1. Validación de Trial Balance (La suma de débitos y créditos debe ser 0)
        total_balance = mapped_tb_df['saldo_final'].sum()
        if abs(total_balance) > tolerance:
            errors.append(f"El Balance de Comprobación está descuadrado por {total_balance:,.2f}")
            
        # 2. Validación en Estados Financieros (Si se proporcionan)
        if statements_dict:
            # Busca claves universales si el usuario las configuró
            activos = statements_dict.get('Total Activos', 0)
            pasivos = statements_dict.get('Total Pasivos', 0)
            patrimonio = statements_dict.get('Total Patrimonio', 0)
            
            # Solo validad si existen realmente las tres patas
            if 'Total Activos' in statements_dict and 'Total Pasivos' in statements_dict and 'Total Patrimonio' in statements_dict:
                diff = abs(activos - (pasivos + patrimonio))
                if diff > tolerance:
                    errors.append(f"Ecuación Contable descuadrada. Activos: {activos}, Pasivos+Patrimonio: {pasivos+patrimonio}. Diferencia: {diff}")
        
        return {
            "is_valid": len(errors) == 0,
            "errors": errors
        }

    @staticmethod
    def validate_excel_template(file_path, expected_sheets=None, required_columns=None):
        """
        Realiza un Pre-Flight Check de la estructura de un archivo Excel de plantilla o mapeo.
        Verifica la existencia de hojas clave y de columnas críticas para evitar fallas silenciosas.
        """
        import openpyxl
        import os
        
        errors = []
        if not os.path.exists(file_path):
            return {
                "is_valid": False,
                "errors": [f"El archivo '{os.path.basename(file_path)}' no existe en la ruta especificada."]
            }
            
        try:
            wb = openpyxl.load_workbook(file_path, read_only=True)
            sheet_names = wb.sheetnames
            wb.close()
        except Exception as e:
            return {
                "is_valid": False,
                "errors": [f"Error al abrir el archivo Excel: {e}"]
            }
            
        # 1. Validar hojas esperadas
        if expected_sheets:
            for sh in expected_sheets:
                if sh not in sheet_names:
                    errors.append(f"Falta la pestaña obligatoria '{sh}' en el archivo Excel.")
                    
        # 2. Validar columnas requeridas (si se especifica y existe al menos una hoja)
        if required_columns and sheet_names:
            try:
                import pandas as pd
                # Leer solo la primera fila de las hojas especificadas para checkear columnas
                sheets_to_check = expected_sheets if expected_sheets else [sheet_names[0]]
                for sh in sheets_to_check:
                    if sh in sheet_names:
                        df_first = pd.read_excel(file_path, sheet_name=sh, nrows=0)
                        cols = [str(c).strip().lower() for c in df_first.columns]
                        for rc in required_columns:
                            # Búsqueda flexible de la columna
                            if not any(rc.lower() in c for c in cols):
                                errors.append(f"Falta la columna '{rc}' (o una similar) en la pestaña '{sh}' del Excel.")
            except Exception as e:
                errors.append(f"Error al validar columnas del Excel: {e}")
                
        return {
            "is_valid": len(errors) == 0,
            "errors": errors
        }

