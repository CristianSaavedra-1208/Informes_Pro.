import re

FILE_PATH = r"c:\Users\cfsaa\OneDrive\Desktop\Informes_Pro\src\ui_pages\organizacion_de_cuentas.py"

with open(FILE_PATH, "r", encoding="utf-8") as f:
    orig = f.read()

# REPLACE 1: Account dynamic list (Lines 276-289)
replacement_1 = """        # 1. Cuentas dinámicas del programa (Plan Maestro + Trial Balance)
        cuentas_disponibles_set = set()
        if 'plan_cuentas_df' in st.session_state:
            cuentas_disponibles_set.update([str(x).strip() for x in st.session_state['plan_cuentas_df']['Cuenta'].dropna().unique()])
            
        if 'tb_df' in st.session_state:
            cuentas_disponibles_set.update([str(x).strip() for x in st.session_state['tb_df']['cuenta_id'].unique()])
        else:
            try:
                from src.models.trial_balance_db import TrialBalanceDB
                # get_available_periods and get_trial_balance
                per = TrialBalanceDB.get_available_periods(empresa_seleccionada)
                if per:
                    db_tb = TrialBalanceDB.get_trial_balance(empresa_seleccionada, per[-1])
                    if db_tb is not None:
                        cuentas_disponibles_set.update([str(r['cuenta_id']).strip() for _, r in db_tb.iterrows()])
            except Exception:
                pass
                
        cuentas_disponibles = sorted(list(cuentas_disponibles_set))
        if not cuentas_disponibles:
            cuentas_disponibles = ["- No hay cuentas cargadas -"]

        cuenta_ejemplo = col1.selectbox("N° de Cuenta:", cuentas_disponibles, key="c_ej")
"""

pattern_1 = re.compile(
    r"        # 1\. Cuentas dinámicas del programa.*?cuenta_ejemplo = col1\.selectbox\([^)]*\)\s",
    re.DOTALL
)

new_content = pattern_1.sub(replacement_1, orig)


# REPLACE 2: Dynamic Selectboxes for Notes (Lines 297-348)
replacement_2 = """        # 2. Construir Selectores de Notas independientes para cada columna
        nota_ejemplo = ""
        nota_config_dict = {}
        
        if tipo_mapeo == "Balance" and 'map_balance_df' in st.session_state:
            df_b = st.session_state['map_balance_df']
            cls_col = next((c for c in df_b.columns if "clasificaci" in c.lower() and "balance" in c.lower()), "Clasificación balance")
            
            sub_df = df_b[df_b[cls_col] == clasificacion] if cls_col in df_b.columns else df_b
            nota_cols = [c for c in df_b.columns if "nota" in c.lower() and "id" not in c.lower()]
            
            for n_idx, nc in enumerate(nota_cols):
                # Extraemos opciones: primero las específicas a la clasificación, luego las globales para no limitar al usuario
                opciones_especificas = sorted([str(x).strip() for x in sub_df[nc].dropna().unique() if str(x).strip()])
                opciones_globales = sorted([str(x).strip() for x in df_b[nc].dropna().unique() if str(x).strip()])
                
                # Mantenemos orden y removemos duplicados
                todas_opc = list(dict.fromkeys(opciones_especificas + opciones_globales))
                
                # Permitir opción vacía para que no se obligue a poner nota si no es necesario (ej. subnota vacía)
                todas_opc.insert(0, "")
                
                sel_val = col3.selectbox(f"{nc}:", todas_opc, key=f"n_sel_{n_idx}")
                nota_config_dict[nc] = sel_val

            # Si no hay nota_cols, la UI no mostrará nada, pero mantenemos el dict por defecto
            
            # Recuperar el ID de Nota Asociada genérico según la clasificación
            id_nota_c = next((c for c in df_b.columns if "id_nota" in c.lower()), "ID_Nota_Asociada")
            id_rep_c = next((c for c in df_b.columns if "id_reporte" in c.lower()), "ID_Reporte")
            
            if id_nota_c in df_b.columns: 
                v = sub_df[id_nota_c].dropna().values
                nota_config_dict[id_nota_c] = str(v[0]) if len(v) > 0 else ""
            if id_rep_c in df_b.columns: 
                v = sub_df[id_rep_c].dropna().values
                nota_config_dict[id_rep_c] = str(v[0]) if len(v) > 0 else ""
                
            import json
            nota_ejemplo = json.dumps(nota_config_dict)
            
        elif tipo_mapeo != "Balance":
            dict_notas = {}
            try:
                from src.models.database import SessionLocal
                from src.models.taxonomy_master import TaxonomyMasterRecord
                import json
                db = SessionLocal()
                tax_notas = db.query(TaxonomyMasterRecord.id_nota_asociada, TaxonomyMasterRecord.desglose_nota_es).filter_by(
                    empresa=empresa_seleccionada, 
                    nombre_linea_es=clasificacion
                ).filter(TaxonomyMasterRecord.id_nota_asociada.isnot(None), TaxonomyMasterRecord.id_nota_asociada != "").all()
                db.close()
                for r in tax_notas:
                    cod = str(r[0]).strip()
                    desc = str(r[1]).strip() if r[1] else cod
                    if cod: dict_notas[json.dumps({"ID_Nota_Asociada": cod})] = desc
            except Exception:
                pass
                
            if not dict_notas:
                dict_notas = {"": "- Sin Notas o Sub-Notas Encontradas -"}
                
            llaves_notas = sorted(list(dict_notas.keys()), key=lambda k: dict_notas[k])
            
            nota_ejemplo = col3.selectbox(
                "Nota Asociada Obligatoria:", 
                llaves_notas, 
                format_func=lambda x: dict_notas[x],
                key="n_sel_pl"
            )
        
        if st.button("Guardar Clasificación", type="primary"):
"""

pattern_2 = re.compile(
    r"        # 2\. Construir Diccionario de Notas dinámicas.*?if st\.button\(\"Guardar Clasificación\", type=\"primary\"\):",
    re.DOTALL
)

new_content = pattern_2.sub(replacement_2, new_content)

with open(FILE_PATH, "w", encoding="utf-8") as f:
    f.write(new_content)

print("Actualizacion lista")
