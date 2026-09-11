import os
import re
import sqlite3
import json
from datetime import datetime

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DATA_DIR = os.path.join(ROOT_DIR, "data")
EMPRESAS_DIR = os.path.join(DATA_DIR, "empresas")
DB_PATH = os.path.join(DATA_DIR, "informes_pro.db")

class CompanyResolver:
    """
    Motor central de resolución y desacoplamiento de Códigos Únicos vs Carátulas visibles.
    Garantiza que toda la base de datos y motores operen con códigos inmutables (EMP_001, GRP_001)
    mientras la interfaz y reportes muestran la carátula visible editable.
    """
    
    @staticmethod
    def _get_connection():
        os.makedirs(DATA_DIR, exist_ok=True)
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        return conn

    @classmethod
    def ensure_catalog_table(cls):
        """Crea y asegura la estructura de la tabla de catálogo maestro de empresas."""
        conn = cls._get_connection()
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS company_catalog (
                codigo_id VARCHAR(50) PRIMARY KEY,
                nombre_caratula VARCHAR(255) NOT NULL,
                rut VARCHAR(50),
                es_consolidado BOOLEAN DEFAULT 0,
                slug_dir VARCHAR(255) NOT NULL,
                activa BOOLEAN DEFAULT 1,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        conn.commit()
        conn.close()
        
        # Sincronizar catálogo con las empresas existentes si está vacío o incompleto
        cls.sync_catalog()

    @classmethod
    def sync_catalog(cls):
        """Sincroniza y migra las empresas y grupos existentes al catálogo maestro."""
        conn = cls._get_connection()
        cursor = conn.cursor()
        
        # 1. Obtener empresas registradas en catálogo
        cursor.execute("SELECT codigo_id, nombre_caratula, slug_dir FROM company_catalog")
        existing_catalog = {row["codigo_id"]: dict(row) for row in cursor.fetchall()}
        existing_names = {row["nombre_caratula"]: row["codigo_id"] for row in existing_catalog.values()}
        existing_slugs = {row["slug_dir"]: row["codigo_id"] for row in existing_catalog.values()}
        
        # 2. Escanear carpetas físicas en data/empresas
        if os.path.exists(EMPRESAS_DIR):
            disk_folders = sorted([d for d in os.listdir(EMPRESAS_DIR) if os.path.isdir(os.path.join(EMPRESAS_DIR, d))])
            for folder in disk_folders:
                if folder not in existing_slugs and folder not in existing_names:
                    es_grupo = folder.startswith("[GRUPO]") or "consolidado" in folder.lower()
                    next_code = cls._generate_next_code(cursor, es_grupo)
                    cursor.execute("""
                        INSERT INTO company_catalog (codigo_id, nombre_caratula, rut, es_consolidado, slug_dir, activa)
                        VALUES (?, ?, ?, ?, ?, 1)
                    """, (next_code, folder, None, 1 if es_grupo else 0, folder))
                    print(f"[CompanyResolver] Registrada empresa existente '{folder}' con código '{next_code}'.")
                    
        conn.commit()
        conn.close()

    @staticmethod
    def _generate_next_code(cursor, es_grupo: bool) -> str:
        prefix = "GRP" if es_grupo else "EMP"
        cursor.execute("SELECT codigo_id FROM company_catalog WHERE codigo_id LIKE ?", (f"{prefix}_%",))
        rows = cursor.fetchall()
        max_num = 0
        for r in rows:
            code = r[0] if isinstance(r, (tuple, list)) else r["codigo_id"]
            match = re.search(r'(\d+)$', str(code))
            if match:
                max_num = max(max_num, int(match.group(1)))
        return f"{prefix}_{max_num + 1:03d}"

    @classmethod
    def get_code(cls, name_or_code: str) -> str:
        """Obtiene el codigo_id inmutable a partir de un nombre de carátula o código."""
        if not name_or_code:
            return ""
        s = str(name_or_code).strip()
        
        conn = cls._get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT codigo_id FROM company_catalog WHERE codigo_id = ? OR nombre_caratula = ? OR slug_dir = ?", (s, s, s))
        row = cursor.fetchone()
        conn.close()
        
        if row:
            return row["codigo_id"]
        return s

    @classmethod
    def get_display_name(cls, code_or_name: str) -> str:
        """Obtiene el nombre de carátula visible actual a partir de un código o nombre."""
        if not code_or_name:
            return ""
        s = str(code_or_name).strip()
        
        conn = cls._get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT nombre_caratula FROM company_catalog WHERE codigo_id = ? OR nombre_caratula = ? OR slug_dir = ?", (s, s, s))
        row = cursor.fetchone()
        conn.close()
        
        if row:
            return row["nombre_caratula"]
        return s

    @classmethod
    def get_folder_path(cls, code_or_name: str) -> str:
        """Obtiene la ruta física absoluta de la carpeta de la empresa."""
        if not code_or_name:
            return EMPRESAS_DIR
        s = str(code_or_name).strip()
        
        conn = cls._get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT slug_dir FROM company_catalog WHERE codigo_id = ? OR nombre_caratula = ? OR slug_dir = ?", (s, s, s))
        row = cursor.fetchone()
        conn.close()
        
        folder_name = row["slug_dir"] if row else s
        return os.path.join(EMPRESAS_DIR, folder_name)

    @classmethod
    def get_all_companies(cls, include_groups: bool = True, only_active: bool = True) -> list:
        """Retorna la lista de todas las empresas registradas con su código y carátula."""
        cls.ensure_catalog_table()
        conn = cls._get_connection()
        cursor = conn.cursor()
        
        query = "SELECT codigo_id, nombre_caratula, rut, es_consolidado, slug_dir, activa FROM company_catalog WHERE 1=1"
        params = []
        if not include_groups:
            query += " AND es_consolidado = 0"
        if only_active:
            query += " AND activa = 1"
        query += " ORDER BY es_consolidado ASC, nombre_caratula ASC"
        
        cursor.execute(query, params)
        res = [dict(r) for r in cursor.fetchall()]
        conn.close()
        return res

    @classmethod
    def register_company(cls, nombre_caratula: str, es_consolidado: bool = False, rut: str = None) -> tuple:
        """Registra una nueva empresa o grupo, generando automáticamente su código único."""
        cls.ensure_catalog_table()
        name_clean = str(nombre_caratula).strip()
        if not name_clean:
            return False, "El nombre de la empresa no puede estar vacío.", None
            
        conn = cls._get_connection()
        cursor = conn.cursor()
        
        # Verificar si ya existe una empresa con ese nombre
        cursor.execute("SELECT codigo_id FROM company_catalog WHERE nombre_caratula = ?", (name_clean,))
        if cursor.fetchone():
            conn.close()
            return False, f"Ya existe una empresa registrada con el nombre '{name_clean}'.", None
            
        code = cls._generate_next_code(cursor, es_consolidado)
        slug_dir = name_clean # Conserva el nombre como carpeta o usa el código
        
        cursor.execute("""
            INSERT INTO company_catalog (codigo_id, nombre_caratula, rut, es_consolidado, slug_dir, activa)
            VALUES (?, ?, ?, ?, ?, 1)
        """, (code, name_clean, rut, 1 if es_consolidado else 0, slug_dir))
        
        conn.commit()
        conn.close()
        
        # Crear carpeta física y copiar plantillas default
        target_dir = os.path.join(EMPRESAS_DIR, slug_dir)
        os.makedirs(target_dir, exist_ok=True)
        
        import shutil
        templates_dir = os.path.join(ROOT_DIR, "templates")
        if os.path.exists(templates_dir):
            for t_file in os.listdir(templates_dir):
                if t_file.endswith(".xlsx"):
                    shutil.copy2(os.path.join(templates_dir, t_file), os.path.join(target_dir, t_file))
                    
        global_master_dir = os.path.join(EMPRESAS_DIR, "Pacifico Cable SpA")
        if not os.path.exists(global_master_dir):
            global_master_dir = os.path.join(EMPRESAS_DIR, "Pacifico SpA")
        if os.path.exists(global_master_dir):
            for master_f in ["plan_cuentas.xlsx", "map_balance.xlsx", "map_pl.xlsx"]:
                src_master = os.path.join(global_master_dir, master_f)
                if os.path.exists(src_master):
                    shutil.copy2(src_master, os.path.join(target_dir, master_f))
                    
        return True, f"Empresa '{name_clean}' creada exitosamente con código único '{code}'.", code

    @classmethod
    def update_display_name(cls, code_or_name: str, new_display_name: str) -> tuple:
        """
        Modifica la carátula visible de una empresa en 1 segundo sin tocar datos contables.
        """
        cls.ensure_catalog_table()
        new_name_clean = str(new_display_name).strip()
        if not new_name_clean:
            return False, "El nuevo nombre no puede estar vacío."
            
        code = cls.get_code(code_or_name)
        if not code:
            return False, f"No se encontró la empresa '{code_or_name}' en el catálogo."
            
        conn = cls._get_connection()
        cursor = conn.cursor()
        
        # Verificar que no colisione con otra carátula
        cursor.execute("SELECT codigo_id FROM company_catalog WHERE nombre_caratula = ? AND codigo_id != ?", (new_name_clean, code))
        if cursor.fetchone():
            conn.close()
            return False, f"Ya existe otra empresa con el nombre '{new_name_clean}'."
            
        cursor.execute("UPDATE company_catalog SET nombre_caratula = ? WHERE codigo_id = ?", (new_name_clean, code))
        conn.commit()
        conn.close()
        
        return True, f"Carátula actualizada exitosamente: '{new_name_clean}' (Código inmutable: {code})."
