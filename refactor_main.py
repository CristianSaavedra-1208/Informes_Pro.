import re
import os

MAIN_FILE = r"c:\Users\cfsaa\OneDrive\Desktop\Informes_Pro\src\main.py"
UI_DIR = r"c:\Users\cfsaa\OneDrive\Desktop\Informes_Pro\src\ui_pages"

os.makedirs(UI_DIR, exist_ok=True)
with open(os.path.join(UI_DIR, "__init__.py"), "w") as f:
    pass

with open(MAIN_FILE, "r", encoding="utf-8") as f:
    content = f.read()

blocks = re.split(r'(    (?:if|elif) menu == "[^"]+":\n)', content)
header = blocks[0]

routers = []
for i in range(1, len(blocks), 2):
    match_str = blocks[i]
    body_str = blocks[i+1]
    
    menu_name = re.search(r'== "([^"]+)"', match_str).group(1)
    safe_name = re.sub(r'[^\w\s-]', '', menu_name).strip().replace(' ', '_').lower()
    
    file_path = os.path.join(UI_DIR, f"{safe_name}.py")
    with open(file_path, "w", encoding="utf-8") as f:
        f.write("import streamlit as st\n")
        f.write("import pandas as pd\n")
        f.write("import os\n")
        f.write("from src.core.excel_utils import df_to_excel_bytes\n\n")
        f.write("def render(empresa_seleccionada, empresa_path):\n")
        
        # We need to un-indent the body_str by 4 spaces.
        # body_str currently has 8 spaces. If we keep 8, it belongs to the def perfectly.
        # But wait, `if menu` was at 4 spaces, the content inside `if` is 8 spaces.
        # Inside `def render:` it requires 4 spaces. So 8 spaces is exactly 4 spaces (inside function).
        # Actually, `def render` is at 0 spaces, so 4 spaces inside is correct.
        lines = body_str.split("\n")
        for line in lines:
            if line.startswith("        "):
                f.write(line[4:] + "\n")
            elif line.startswith("    "):
                f.write(line[4:] + "\n")
            else:
                f.write(line + "\n")
        
    routers.append((match_str, safe_name))
    print(f"Created {file_path}")

print("Done generating pages.")
