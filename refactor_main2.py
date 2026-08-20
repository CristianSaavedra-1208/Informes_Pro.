import re

MAIN_FILE = r"c:\Users\cfsaa\OneDrive\Desktop\Informes_Pro\src\main.py"

with open(MAIN_FILE, "r", encoding="utf-8") as f:
    orig_content = f.read()

# We want to replace everything from `menu = st.sidebar.radio` down to the end of the file.
blocks = re.split(r'(    menu = st\.sidebar\.radio\("Navegación", \[)', orig_content)
header_only = blocks[0]

import_pages = """
import src.ui_pages.inicio as pg_inicio
import src.ui_pages.1_cargas_de_datos as pg_cargas
import src.ui_pages.2_organizacion_de_cuentas as pg_map
import src.ui_pages.3_ajustes_manuales as pg_ajustes
import src.ui_pages.4_consolidacin as pg_cons
import src.ui_pages.5_estados_financieros as pg_estados
import src.ui_pages.6_informes_y_notas as pg_informes
import src.ui_pages.7_cierre_y_memoria_histrica as pg_cierre
import src.ui_pages.8_auditora_diccionario as pg_audit
import src.ui_pages.parmetros_globales as pg_param
import src.ui_pages.configuraciones as pg_conf

    # Definir wrappers para st.navigation
    def run_inicio(): pg_inicio.render(empresa_seleccionada, empresa_path)
    def run_cargas(): pg_cargas.render(empresa_seleccionada, empresa_path)
    def run_map(): pg_map.render(empresa_seleccionada, empresa_path)
    def run_ajustes(): pg_ajustes.render(empresa_seleccionada, empresa_path)
    def run_cons(): pg_cons.render(empresa_seleccionada, empresa_path)
    def run_estados(): pg_estados.render(empresa_seleccionada, empresa_path)
    def run_informes(): pg_informes.render(empresa_seleccionada, empresa_path)
    def run_cierre(): pg_cierre.render(empresa_seleccionada, empresa_path)
    def run_audit(): pg_audit.render(empresa_seleccionada, empresa_path)
    def run_param(): pg_param.render(empresa_seleccionada, empresa_path)
    def run_conf(): pg_conf.render(empresa_seleccionada, empresa_path)

    pages = {
        "Operaciones": [
            st.Page(run_inicio, title="Inicio", icon="🏠"),
            st.Page(run_cargas, title="Carga de Datos", icon="📥"),
            st.Page(run_map, title="Orquestación Cuentas", icon="🧬"),
            st.Page(run_ajustes, title="Comprobantes", icon="✍️"),
            st.Page(run_cons, title="Consolidación", icon="🏢")
        ],
        "Reportes": [
            st.Page(run_estados, title="Estados Financieros", icon="📊"),
            st.Page(run_informes, title="Informes & Notas", icon="📄"),
            st.Page(run_audit, title="Auditoría IFRS", icon="🔎")
        ],
        "Administración": [
            st.Page(run_cierre, title="Históricos", icon="💾"),
            st.Page(run_param, title="Parámetros", icon="⚙️"),
            st.Page(run_conf, title="Roles & Settings", icon="🛡️")
        ]
    }
    
    pg = st.navigation(pages)
    pg.run()

if __name__ == "__main__":
    main()
"""

new_content = header_only + import_pages

with open(MAIN_FILE, "w", encoding="utf-8") as f:
    f.write(new_content)

print("main.py refactored for st.navigation!")
