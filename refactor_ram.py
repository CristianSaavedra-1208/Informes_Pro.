import re

# 1. Update main.py
MAIN_FILE = r"c:\Users\cfsaa\OneDrive\Desktop\Informes_Pro\src\main.py"
with open(MAIN_FILE, "r", encoding="utf-8") as f:
    orig_main = f.read()

replacement_main = """    if 'tb_df' not in st.session_state:
        from src.models.trial_balance_db import TrialBalanceDB
        try:
            periodos = TrialBalanceDB.get_available_periods(empresa_seleccionada)
            if periodos:
                st.session_state['tb_df'] = TrialBalanceDB.get_trial_balance(empresa_seleccionada, periodos[-1])
        except Exception:
            pass"""

# The match starts from `    if 'tb_df' not in st.session_state:` down to the end of `except: pass`
pattern_main = re.compile(
    r"    if 'tb_df' not in st\.session_state:.*?except:\n\s+pass",
    re.DOTALL
)

new_main = pattern_main.sub(replacement_main, orig_main)

with open(MAIN_FILE, "w", encoding="utf-8") as f:
    f.write(new_main)

# 2. Update cargas_de_datos.py
CARGAS_FILE = r"c:\Users\cfsaa\OneDrive\Desktop\Informes_Pro\src\ui_pages\cargas_de_datos.py"
with open(CARGAS_FILE, "r", encoding="utf-8") as f:
    orig_cargas = f.read()

new_cargas = orig_cargas.replace(
    "st.session_state['tb_df'] = df_erp",
    "st.session_state['tb_df'] = TrialBalanceDB.get_trial_balance(empresa_seleccionada, periodo_str)"
).replace(
    "st.session_state['tb_df'] = df",
    "st.session_state['tb_df'] = TrialBalanceDB.get_trial_balance(empresa_seleccionada, periodo_str)"
)

with open(CARGAS_FILE, "w", encoding="utf-8") as f:
    f.write(new_cargas)

print("Phase 4 RAM optimization applied successfully.")
