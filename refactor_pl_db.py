import re

# 1. ACTUALIZAR main.py
MAIN_FILE = r"c:\Users\cfsaa\OneDrive\Desktop\Informes_Pro\src\main.py"
with open(MAIN_FILE, "r", encoding="utf-8") as f:
    orig = f.read()

replacement_main = """    if 'pl_df' not in st.session_state:
        from src.models.pl_cubo_db import PlCuboDB
        try:
            per_pl = PlCuboDB.get_available_periods(empresa_seleccionada)
            if per_pl:
                st.session_state['pl_df'] = PlCuboDB.get_pl_cubo(empresa_seleccionada, per_pl[-1])
        except Exception:
            pass"""

pattern_main = re.compile(
    r"    if 'pl_df' not in st\.session_state:.*?except:\n\s+pass",
    re.DOTALL
)

new_main = pattern_main.sub(replacement_main, orig)
with open(MAIN_FILE, "w", encoding="utf-8") as f:
    f.write(new_main)

# 2. ACTUALIZAR cargas_de_datos.py
CARGAS_FILE = r"c:\Users\cfsaa\OneDrive\Desktop\Informes_Pro\src\ui_pages\cargas_de_datos.py"
with open(CARGAS_FILE, "r", encoding="utf-8") as f:
    orig_cargas = f.read()

replacement_cargas = """                if cuenta_col and has_pl_cols:
                    from src.models.pl_cubo_db import PlCuboDB
                    PlCuboDB.save_pl_cubo(empresa_seleccionada, periodo_str_pl, df_cubo)
                    
                    st.success(f"✅ Archivo procesado con éxito e histórico respaldado bajo {periodo_str_pl}. Se ingestaron {len(df_cubo)} líneas en Base de Datos de forma robusta.")
                    st.session_state['pl_df'] = PlCuboDB.get_pl_cubo(empresa_seleccionada, periodo_str_pl)
                    
                    # Conservamos el volcado excel para legacy tools if needed
                    df_cubo.to_excel(os.path.join(empresa_path, "pl_cubo.xlsx"), index=False)
                    df_cubo.to_excel(os.path.join(empresa_path, f"pl_cubo_{periodo_str_pl}.xlsx"), index=False)
                else:"""

pattern_cargas = re.compile(
    r"                if cuenta_col and has_pl_cols:\n                    st\.success.*?else:",
    re.DOTALL
)

new_cargas = pattern_cargas.sub(replacement_cargas, orig_cargas)
with open(CARGAS_FILE, "w", encoding="utf-8") as f:
    f.write(new_cargas)

# 3. ACTUALIZAR configuraciones.py (La aniquilacion de P&L db)
CONF_FILE = r"c:\Users\cfsaa\OneDrive\Desktop\Informes_Pro\src\ui_pages\configuraciones.py"
with open(CONF_FILE, "r", encoding="utf-8") as f:
    orig_conf = f.read()

new_conf = orig_conf.replace(
    "db.query(TrialBalanceRecord).filter(TrialBalanceRecord.empresa == empresa_seleccionada, TrialBalanceRecord.periodo == per_to_del).delete()",
    "db.query(TrialBalanceRecord).filter(TrialBalanceRecord.empresa == empresa_seleccionada, TrialBalanceRecord.periodo == per_to_del).delete()\n                        from src.models.pl_record import PlRecord\n                        db.query(PlRecord).filter(PlRecord.empresa == empresa_seleccionada, PlRecord.periodo == per_to_del).delete()"
)
with open(CONF_FILE, "w", encoding="utf-8") as f:
    f.write(new_conf)

print("Actualización de arquitectura SQL PL Finalizada")
