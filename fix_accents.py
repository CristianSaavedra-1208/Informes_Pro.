import os

UI_DIR = r"c:\Users\cfsaa\OneDrive\Desktop\Informes_Pro\src\ui_pages"
MAIN_FILE = r"c:\Users\cfsaa\OneDrive\Desktop\Informes_Pro\src\main.py"

# Mapping de archivos viejos (con tildes) a nuevos (limpios)
renames = {
    "auditoría_diccionario.py": "auditoria_diccionario.py",
    "cierre_y_memoria_histórica.py": "cierre_y_memoria_historica.py",
    "consolidación.py": "consolidacion.py",
    "parámetros_globales.py": "parametros_globales.py"
}

# 1. Renombrar archivos
for old_name, new_name in renames.items():
    old_path = os.path.join(UI_DIR, old_name)
    new_path = os.path.join(UI_DIR, new_name)
    if os.path.exists(old_path):
        os.rename(old_path, new_path)
        print(f"Renombrado: {old_name} -> {new_name}")

# 2. Reemplazar en main.py
with open(MAIN_FILE, "r", encoding="utf-8") as f:
    content = f.read()

content = content.replace("import src.ui_pages.consolidación as pg_cons", "import src.ui_pages.consolidacion as pg_cons")
content = content.replace("import src.ui_pages.cierre_y_memoria_histórica as pg_cierre", "import src.ui_pages.cierre_y_memoria_historica as pg_cierre")
content = content.replace("import src.ui_pages.auditoría_diccionario as pg_audit", "import src.ui_pages.auditoria_diccionario as pg_audit")
content = content.replace("import src.ui_pages.parámetros_globales as pg_param", "import src.ui_pages.parametros_globales as pg_param")

with open(MAIN_FILE, "w", encoding="utf-8") as f:
    f.write(content)

print("Acento purgado completado en script e imports.")
