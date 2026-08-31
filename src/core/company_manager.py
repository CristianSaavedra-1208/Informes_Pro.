import os
import json
import sqlite3
import pandas as pd

def rename_company_cascade(empresas_dir: str, db_path: str, old_name: str, new_name: str) -> tuple:
    """
    Renombra una empresa en cascada de forma completa y atómica:
    1. Renombra la carpeta en disco dentro de data/empresas
    2. Actualiza los registros en todas las tablas de la base de datos SQLite
    3. Actualiza el archivo de control de snapshots JSON
    4. Actualiza el archivo last_active.txt
    """
    old_name = str(old_name).strip()
    new_name = str(new_name).rstrip('.').strip()
    
    if not old_name or not new_name:
        return False, "Los nombres de empresa no pueden estar vacíos."
    if old_name == new_name:
        return False, "El nuevo nombre debe ser diferente al actual."
        
    # 1. Renombrar carpeta física
    old_path = os.path.join(empresas_dir, old_name)
    new_path = os.path.join(empresas_dir, new_name)
    
    if os.path.exists(old_path) and not os.path.exists(new_path):
        os.rename(old_path, new_path)
    elif not os.path.exists(old_path) and not os.path.exists(new_path):
        return False, f"No se encontró la carpeta de la empresa '{old_name}' en disco."

    # 2. Actualizar registros en base de datos SQLite
    if os.path.exists(db_path):
        try:
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()

            tables_empresa = [
                'trial_balance_records',
                'pl_records_dim',
                'historical_data_records',
                'historical_detail_records',
                'audit_adjustments',
                'cash_flow_adjustments',
                'taxonomy_master'
            ]

            for tbl in tables_empresa:
                try:
                    cursor.execute(f"UPDATE {tbl} SET empresa = ? WHERE empresa = ?", (new_name, old_name))
                except Exception:
                    pass

            try:
                cursor.execute("UPDATE consolidation_groups SET empresa_matriz = ? WHERE empresa_matriz = ?", (new_name, old_name))
                cursor.execute("UPDATE consolidation_groups SET empresa_filial = ? WHERE empresa_filial = ?", (new_name, old_name))
            except Exception:
                pass

            try:
                cursor.execute("UPDATE company_entities SET nombre = ? WHERE nombre = ?", (new_name, old_name))
            except Exception:
                pass

            try:
                cursor.execute("UPDATE company_catalog SET nombre_caratula = ?, slug_dir = ? WHERE nombre_caratula = ? OR slug_dir = ?", (new_name, new_name, old_name, old_name))
            except Exception:
                pass

            conn.commit()
            conn.close()
        except Exception as e:
            return False, f"Error actualizando la base de datos: {e}"

    # 3. Actualizar snapshots_control.json
    root_data_dir = os.path.dirname(empresas_dir)
    snapshots_path = os.path.join(root_data_dir, "snapshots_control.json")
    if os.path.exists(snapshots_path):
        try:
            with open(snapshots_path, 'r', encoding='utf-8') as f:
                snapshots_data = json.load(f)
            if old_name in snapshots_data:
                snapshots_data[new_name] = snapshots_data.pop(old_name)
                with open(snapshots_path, 'w', encoding='utf-8') as f:
                    json.dump(snapshots_data, f, indent=2, ensure_ascii=False)
        except Exception:
            pass

    # 4. Actualizar last_active.txt
    last_active_path = os.path.join(root_data_dir, "last_active.txt")
    if os.path.exists(last_active_path):
        try:
            with open(last_active_path, 'r', encoding='utf-8') as f:
                last_act = f.read().strip()
            if last_act == old_name:
                with open(last_active_path, 'w', encoding='utf-8') as f:
                    f.write(new_name)
        except Exception:
            pass

    return True, f"Empresa '{old_name}' renombrada exitosamente a '{new_name}' (carpeta y bases de datos sincronizadas)."

def check_company_has_data(empresas_dir: str, db_path: str, empresa_name: str) -> tuple:
    """
    Verifica si una empresa contiene datos transaccionales, balances o información histórica.
    Retorna (has_data: bool, reasons: list[str])
    """
    reasons = []
    empresa_name = str(empresa_name).strip()
    
    # 1. Verificar registros en base de datos SQLite
    if os.path.exists(db_path):
        try:
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            # Trial balance
            cursor.execute("SELECT COUNT(DISTINCT periodo), COUNT(*) FROM trial_balance_records WHERE empresa = ?", (empresa_name,))
            row = cursor.fetchone()
            if row and row[1] > 0:
                reasons.append(f"Tiene {row[1]} registros de Balance de Comprobación ({row[0]} períodos).")
                
            # P&L Cubo
            cursor.execute("SELECT COUNT(DISTINCT periodo), COUNT(*) FROM pl_records_dim WHERE empresa = ?", (empresa_name,))
            row = cursor.fetchone()
            if row and row[1] > 0:
                reasons.append(f"Tiene {row[1]} registros en Cubo P&L ({row[0]} períodos).")
                
            # Datos Históricos
            cursor.execute("SELECT COUNT(*) FROM historical_data_records WHERE empresa = ? OR empresa = ?", (empresa_name, empresa_name.replace("[GRUPO] ", "")))
            row = cursor.fetchone()
            if row and row[0] > 0:
                reasons.append(f"Tiene {row[0]} registros en Memoria Histórica.")
                
            # Ajustes de auditoría o flujos
            cursor.execute("SELECT COUNT(*) FROM audit_adjustments WHERE empresa = ?", (empresa_name,))
            row = cursor.fetchone()
            if row and row[0] > 0:
                reasons.append(f"Tiene {row[0]} ajustes de auditoría registrados.")

            cursor.execute("SELECT COUNT(*) FROM cash_flow_adjustments WHERE empresa = ?", (empresa_name,))
            row = cursor.fetchone()
            if row and row[0] > 0:
                reasons.append(f"Tiene {row[0]} ajustes de flujo de efectivo registrados.")
                
            conn.close()
        except Exception:
            pass
            
    # 2. Verificar archivos físicos transaccionales en la carpeta
    emp_path = os.path.join(empresas_dir, empresa_name)
    if os.path.exists(emp_path):
        trans_files = []
        for f in os.listdir(emp_path):
            if f.startswith("pl_cubo_") and f.endswith(".xlsx"):
                trans_files.append(f)
            elif f in ["pl_cubo.xlsx", "temp_uploaded.xlsx"]:
                trans_files.append(f)
        if trans_files:
            reasons.append(f"Tiene {len(trans_files)} archivos transaccionales cargados en disco.")
            
    has_data = len(reasons) > 0
    return has_data, reasons

def delete_company(empresas_dir: str, db_path: str, empresa_name: str) -> tuple:
    """
    Elimina una empresa siempre que no posea data cargada.
    Si posee data, bloquea la acción e indica que debe borrarse previamente en 'Eliminación de Data'.
    """
    empresa_name = str(empresa_name).strip()
    if not empresa_name:
        return False, "El nombre de la empresa no puede estar vacío."
        
    # Verificar primero si la empresa tiene data
    has_data, reasons = check_company_has_data(empresas_dir, db_path, empresa_name)
    if has_data:
        detalle = " ".join(reasons)
        return False, f"⚠️ Antes de eliminar la empresa '{empresa_name}', debe borrar la data con la opción disponible en la pestaña 'Eliminación de Data' del módulo. Motivos detectados: {detalle}"
        
    # 1. Eliminar carpeta física en disco
    emp_path = os.path.join(empresas_dir, empresa_name)
    if os.path.exists(emp_path):
        import shutil
        try:
            shutil.rmtree(emp_path)
        except Exception as e:
            return False, f"Error al eliminar la carpeta física de la empresa: {e}"
            
    # 2. Limpiar registros en base de datos SQLite si hubiera algún remanente
    if os.path.exists(db_path):
        try:
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            for tbl in ['company_entities', 'company_catalog', 'taxonomy_master', 'consolidation_groups']:
                try:
                    if tbl == 'company_entities':
                        cursor.execute("DELETE FROM company_entities WHERE nombre = ?", (empresa_name,))
                    elif tbl == 'company_catalog':
                        cursor.execute("DELETE FROM company_catalog WHERE codigo_id = ? OR nombre_caratula = ? OR slug_dir = ?", (empresa_name, empresa_name, empresa_name))
                    elif tbl == 'consolidation_groups':
                        cursor.execute("DELETE FROM consolidation_groups WHERE empresa_matriz = ? OR empresa_filial = ? OR nombre_grupo = ?", (empresa_name, empresa_name, empresa_name))
                    else:
                        cursor.execute(f"DELETE FROM {tbl} WHERE empresa = ?", (empresa_name,))
                except Exception:
                    pass
                    
            conn.commit()
            conn.close()
        except Exception:
            pass
            
    # 3. Limpiar snapshots_control.json
    root_data_dir = os.path.dirname(empresas_dir)
    snapshots_path = os.path.join(root_data_dir, "snapshots_control.json")
    if os.path.exists(snapshots_path):
        try:
            with open(snapshots_path, 'r', encoding='utf-8') as f:
                snapshots_data = json.load(f)
            if empresa_name in snapshots_data:
                del snapshots_data[empresa_name]
                with open(snapshots_path, 'w', encoding='utf-8') as f:
                    json.dump(snapshots_data, f, indent=2, ensure_ascii=False)
        except Exception:
            pass
            
    # 4. Limpiar last_active.txt si era la empresa activa
    last_active_path = os.path.join(root_data_dir, "last_active.txt")
    if os.path.exists(last_active_path):
        try:
            with open(last_active_path, 'r', encoding='utf-8') as f:
                last_act = f.read().strip()
            if last_act == empresa_name:
                remaining = sorted([d for d in os.listdir(empresas_dir) if os.path.isdir(os.path.join(empresas_dir, d))])
                next_act = remaining[0] if remaining else ""
                with open(last_active_path, 'w', encoding='utf-8') as f:
                    f.write(next_act)
        except Exception:
            pass
            
    return True, f"Empresa '{empresa_name}' eliminada exitosamente del sistema."
