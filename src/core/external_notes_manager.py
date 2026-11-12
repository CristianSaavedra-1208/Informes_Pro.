import os
import shutil
import openpyxl
from openpyxl.cell.cell import MergedCell
from io import BytesIO
import datetime
import re

ROOT_PROJECT_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

class ExternalNotesManager:
    MASTER_EXTERNAL_TEMPLATE = os.path.join(ROOT_PROJECT_DIR, "Plantilla EXTERNA_de_notas_v1.xlsx")

    @classmethod
    def get_master_template_path(cls) -> str:
        if os.path.exists(cls.MASTER_EXTERNAL_TEMPLATE):
            return cls.MASTER_EXTERNAL_TEMPLATE
        fallback = os.path.join(ROOT_PROJECT_DIR, "Plantilla de notas_v1.xlsx")
        return fallback

    @classmethod
    def save_master_template(cls, file_bytes: bytes) -> str:
        path = cls.MASTER_EXTERNAL_TEMPLATE
        with open(path, "wb") as f:
            f.write(file_bytes)
        return path

    @classmethod
    def get_master_template_info(cls) -> dict:
        path = cls.get_master_template_path()
        if os.path.exists(path) and os.path.getsize(path) > 0:
            mtime = os.path.getmtime(path)
            dt_str = datetime.datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S")
            size_kb = os.path.getsize(path) / 1024.0
            sheets = []
            detected_codes = []
            try:
                from src.reporting.notes import is_technical_sheet
                wb = openpyxl.load_workbook(path, read_only=True)
                sheets = [s for s in wb.sheetnames if not is_technical_sheet(s)]
                codes_set = set()
                for sh in sheets:
                    ws = wb[sh]
                    for row in ws.iter_rows(values_only=True):
                        for val in row:
                            if val and isinstance(val, str):
                                for m in re.findall(r'(#N\d+\.\d+)', val):
                                    codes_set.add(m)
                detected_codes = sorted(codes_set, key=lambda x: [int(n) for n in re.findall(r'\d+', x)])
                wb.close()
            except Exception:
                pass
            return {
                "exists": True,
                "path": path,
                "filename": os.path.basename(path),
                "last_modified": dt_str,
                "size_kb": f"{size_kb:.1f} KB",
                "mtime": int(mtime),
                "sheets": sheets,
                "detected_codes": detected_codes
            }
        return {
            "exists": False,
            "path": path,
            "filename": os.path.basename(path),
            "last_modified": "Sin registros",
            "size_kb": "0 KB",
            "mtime": 0,
            "sheets": [],
            "detected_codes": []
        }

    @classmethod
    def resolve_canonical_company_name(cls, empresa: str) -> str:
        if not empresa:
            return empresa
        empresas_dir = os.path.join("data", "empresas")
        if not os.path.exists(empresas_dir):
            return empresa
            
        all_dirs = [d for d in os.listdir(empresas_dir) if os.path.isdir(os.path.join(empresas_dir, d))]
        existing_dirs = []
        for d in all_dirs:
            p = os.path.join(empresas_dir, d)
            root_files = [f for f in os.listdir(p) if os.path.isfile(os.path.join(p, f))]
            if root_files or d.startswith("[GRUPO]"):
                existing_dirs.append(d)
        if not existing_dirs:
            existing_dirs = all_dirs

        if empresa in existing_dirs:
            return empresa
            
        import re
        def _norm(s):
            s_clean = str(s).strip().lower()
            is_grupo = s_clean.startswith("[grupo]") or s_clean.startswith("_grupo_") or "consolidado" in s_clean
            s_clean = re.sub(r'\[grupo\]\s*|_grupo_\s*|consolidado\s*', '', s_clean)
            s_clean = re.sub(r'\b(spa|s\.a\.|sa|ltd|inc)\b', '', s_clean)
            s_clean = re.sub(r'[^a-z0-9]', '', s_clean)
            return ('grupo_' + s_clean) if is_grupo else s_clean

        norm_target = _norm(empresa)
        for d in existing_dirs:
            if _norm(d) == norm_target:
                return d
        for d in existing_dirs:
            norm_d = _norm(d)
            if norm_target and norm_d and (norm_target in norm_d or norm_d in norm_target):
                return d
        return empresa

    @classmethod
    def normalize_period_str(cls, periodo: str) -> str:
        if not periodo:
            return periodo
        s = str(periodo).strip()
        month_map = {
            "ene": "01", "jan": "01", "feb": "02", "mar": "03", "abr": "04", "apr": "04",
            "may": "05", "jun": "06", "jul": "07", "ago": "08", "aug": "08",
            "sep": "09", "set": "09", "oct": "10", "nov": "11", "dic": "12", "dec": "12"
        }
        m = re.match(r'^(20\d{2})[-_\s/]([a-zA-Z]{3,})$', s)
        if m:
            yr = m.group(1)
            mon_txt = m.group(2).lower()[:3]
            if mon_txt in month_map:
                return f"{yr}-{month_map[mon_txt]}"
        return s

    @classmethod
    def get_company_anexos_dir(cls, empresa: str, periodo: str) -> str:
        canonical = cls.resolve_canonical_company_name(empresa)
        safe_empresa = "".join([c if c.isalnum() or c in (" ", "_", "-", "[", "]") else "_" for c in canonical]).strip()
        norm_per = cls.normalize_period_str(periodo)
        base_anexos = os.path.join("data", "empresas", safe_empresa, "anexos")
        p_norm = os.path.join(base_anexos, norm_per)
        p_raw = os.path.join(base_anexos, periodo)
        
        if os.path.exists(p_norm):
            return p_norm
        elif os.path.exists(p_raw):
            return p_raw
        else:
            os.makedirs(p_norm, exist_ok=True)
            return p_norm

    @classmethod
    def get_anexo_path(cls, empresa: str, periodo: str) -> str:
        dir_path = cls.get_company_anexos_dir(empresa, periodo)
        return os.path.join(dir_path, "Plantilla_EXTERNA_notas.xlsx")

    EXCLUDED_FROM_AGGREGATION = {
        "impuestos diferidos",
        "empresas relacionadas",
        "inversion en relacionadas",
        "inversiones en relacionadas",
        "relacionadas",
        "ajuste manual noviembre 2023",
        "ajustes manuales",
        "bce final 2022",
        "calculo rli",
        "ppa v3_04 abril",
        "ajuste cartera clientes 2022",
        "db_data"
    }

    @classmethod
    def is_consolidated_group(cls, empresa: str) -> bool:
        if not empresa:
            return False
        s = str(empresa).strip().lower()
        return s.startswith("[grupo]") or s.startswith("_grupo_") or "consolidado" in s

    @classmethod
    def get_group_companies(cls, grupo_name: str) -> list:
        """
        Retorna la lista de empresas filiales que componen el grupo consolidado.
        """
        clean_group = str(grupo_name).strip()
        for prefix in ["[GRUPO] ", "[GRUPO]", "_GRUPO_ ", "_GRUPO_"]:
            if clean_group.startswith(prefix):
                clean_group = clean_group[len(prefix):].strip()
                break

        from src.models.database import SessionLocal
        from src.models.consolidacion import ConsolidationGroup

        companies = []
        db = SessionLocal()
        try:
            grupo_obj = db.query(ConsolidationGroup).filter(
                (ConsolidationGroup.nombre_grupo == clean_group) |
                (ConsolidationGroup.nombre_grupo == grupo_name)
            ).first()

            if grupo_obj:
                if grupo_obj.empresa_matriz:
                    companies.append(grupo_obj.empresa_matriz)

                def _collect(filial_val, is_subgroup):
                    if is_subgroup:
                        try:
                            sub_id = int(filial_val)
                            sub_g = db.query(ConsolidationGroup).filter_by(id=sub_id).first()
                            if sub_g:
                                if sub_g.empresa_matriz:
                                    companies.append(sub_g.empresa_matriz)
                                _collect(sub_g.empresa_filial, sub_g.filial_is_group)
                        except (ValueError, TypeError):
                            pass
                    else:
                        companies.append(filial_val)

                _collect(grupo_obj.empresa_filial, grupo_obj.filial_is_group)
        except Exception as e:
            print(f"Error resolviendo empresas del grupo {grupo_name}: {e}")
        finally:
            db.close()

        seen = set()
        result = []
        for c in companies:
            c_clean = str(c).strip()
            if c_clean and c_clean not in seen:
                seen.add(c_clean)
                result.append(c_clean)
        return result

    @classmethod
    def has_explicit_anexo(cls, empresa: str, periodo: str) -> bool:
        path = cls.get_anexo_path(empresa, periodo)
        return os.path.exists(path) and os.path.getsize(path) > 0

    @classmethod
    def has_anexo(cls, empresa: str, periodo: str) -> bool:
        if cls.has_explicit_anexo(empresa, periodo):
            return True
        if cls.is_consolidated_group(empresa):
            subs = cls.get_group_companies(empresa)
            return any(cls.has_explicit_anexo(sub, periodo) for sub in subs)
        return False

    _AGG_ANEXO_CACHE = {}

    @classmethod
    def build_aggregated_group_anexo(cls, grupo_name: str, periodo: str) -> BytesIO | None:
        """
        Suma automáticamente las planillas externas de las filiales del grupo,
        con excepción explícita de Impuestos Diferidos y Saldos Intercompañía/Relacionadas.
        """
        subs = cls.get_group_companies(grupo_name)
        active_subs = [s for s in subs if cls.has_explicit_anexo(s, periodo)]
        if not active_subs:
            return None

        master_tpl_path = cls.get_master_template_path()
        master_mtime = os.path.getmtime(master_tpl_path) if os.path.exists(master_tpl_path) else 0.0
        sub_mtimes = tuple(
            (s, os.path.getmtime(cls.get_anexo_path(s, periodo)))
            for s in active_subs
        )
        cache_key = (str(grupo_name).strip(), cls.normalize_period_str(periodo), master_mtime, sub_mtimes)
        if cache_key in cls._AGG_ANEXO_CACHE:
            return BytesIO(cls._AGG_ANEXO_CACHE[cache_key])

        sheet_cell_sums = {}

        for sub in active_subs:
            sub_path = cls.get_anexo_path(sub, periodo)
            try:
                sub_wb = openpyxl.load_workbook(sub_path, data_only=True)
                for sh in sub_wb.sheetnames:
                    sh_norm = sh.strip().lower()
                    if any(ex in sh_norm for ex in ["impuestos diferidos", "relacionadas", "ajuste", "bce", "calculo", "ppa", "db_data"]):
                        continue
                    ws = sub_wb[sh]
                    for r, row_vals in enumerate(ws.iter_rows(values_only=True), start=1):
                        for c, val in enumerate(row_vals, start=1):
                            if val is not None and isinstance(val, (int, float)) and not isinstance(val, bool):
                                if val != 0:
                                    key = (sh, r, c)
                                    sheet_cell_sums[key] = sheet_cell_sums.get(key, 0.0) + float(val)
                sub_wb.close()
            except Exception as e:
                print(f"Error leyendo anexo de filial {sub}: {e}")

        try:
            agg_wb = openpyxl.load_workbook(master_tpl_path, data_only=False)
        except Exception:
            return None

        for (sh, r, c), total_val in sheet_cell_sums.items():
            if sh in agg_wb.sheetnames:
                target_ws = agg_wb[sh]
                target_cell = target_ws.cell(r, c)
                if isinstance(target_cell, MergedCell) or target_cell.__class__.__name__ == 'MergedCell':
                    continue
                dest_val = target_cell.value
                if isinstance(dest_val, str) and dest_val.startswith('='):
                    continue
                if total_val == int(total_val):
                    target_cell.value = int(total_val)
                else:
                    target_cell.value = round(total_val, 4)

        out_stream = BytesIO()
        agg_wb.save(out_stream)
        agg_wb.close()
        raw_agg_bytes = out_stream.getvalue()
        cls._AGG_ANEXO_CACHE[cache_key] = raw_agg_bytes
        return BytesIO(raw_agg_bytes)

    @classmethod
    def get_anexo_workbook(cls, empresa: str, periodo: str, data_only: bool = True):
        """
        Retorna el libro openpyxl del anexo para la empresa/período.
        Si es un grupo consolidado sin anexo explícito propio, genera y retorna el libro auto-agregado de sus filiales.
        """
        if cls.has_explicit_anexo(empresa, periodo):
            path = cls.get_anexo_path(empresa, periodo)
            return openpyxl.load_workbook(path, data_only=data_only)
        if cls.is_consolidated_group(empresa):
            agg_stream = cls.build_aggregated_group_anexo(empresa, periodo)
            if agg_stream:
                return openpyxl.load_workbook(agg_stream, data_only=data_only)
        return None

    @classmethod
    def save_anexo(cls, empresa: str, periodo: str, file_bytes: bytes) -> str:
        path = cls.get_anexo_path(empresa, periodo)
        with open(path, "wb") as f:
            f.write(file_bytes)
        return path

    @classmethod
    def get_anexo_info(cls, empresa: str, periodo: str) -> dict:
        if cls.has_explicit_anexo(empresa, periodo):
            path = cls.get_anexo_path(empresa, periodo)
            mtime = os.path.getmtime(path)
            dt_str = datetime.datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S")
            size_kb = os.path.getsize(path) / 1024.0
            return {
                "exists": True,
                "is_auto_aggregated": False,
                "path": path,
                "last_modified": dt_str,
                "size_kb": f"{size_kb:.1f} KB",
                "subsidiaries": []
            }

        if cls.is_consolidated_group(empresa):
            subs = cls.get_group_companies(empresa)
            active_subs = [s for s in subs if cls.has_explicit_anexo(s, periodo)]
            if active_subs:
                return {
                    "exists": True,
                    "is_auto_aggregated": True,
                    "path": None,
                    "last_modified": "Suma en vivo",
                    "size_kb": "Auto-generado",
                    "subsidiaries": active_subs
                }

        return {
            "exists": False,
            "is_auto_aggregated": False,
            "path": cls.get_anexo_path(empresa, periodo),
            "last_modified": "Sin registros",
            "size_kb": "0 KB",
            "subsidiaries": []
        }

    @classmethod
    def is_activo_fijo_injected(cls, ext_ws) -> bool:
        """
        Verifica si la hoja 'Activo Fijo' (#N09.1 o #N09.2) en la plantilla externa tiene datos ingresados.
        Si las celdas numéricas tienen valores > 0 o no vacíos, se considera inyectada.
        """
        if ext_ws is None:
            return False
        for r in range(6, 45):
            for c in range(3, min(ext_ws.max_column + 1, 18)):
                val = ext_ws.cell(r, c).value
                if val is not None and isinstance(val, (int, float)) and abs(val) > 0.001:
                    return True
        return False

    EXTERNAL_TABLE_REGISTRY = {
        "#N04.2": {
            "desc_cols": [1],
            "ext_data_cols": [2],
            "base_act_data_cols": [2],
            "base_cmp_data_cols": [3],
        },
        "#N04.3": {
            "desc_cols": [1, 2, 3],  # Detalle, Banco, Moneda
            "ext_data_cols": [4],
            "base_act_data_cols": [4],
            "base_cmp_data_cols": [5],
        },
        "#N06.2": {
            "desc_cols": [2],
            "ext_data_cols": [3],
            "base_act_data_cols": [3],
            "base_cmp_data_cols": [4],
        },
        "#N06.3": {
            "desc_cols": [2],
            "ext_data_cols": [3],
            "base_act_data_cols": [3],
            "base_cmp_data_cols": [4],
        },
        "#N07.1": {
            "desc_cols": [2],
            "ext_data_cols": [3],
            "base_act_data_cols": [3],
            "base_cmp_data_cols": [4],
        },
        "#N07.2": {
            "desc_cols": [2],
            "ext_data_cols": [3],
            "base_act_data_cols": [3],
            "base_cmp_data_cols": [4],
        },
        "#N07.3": {
            "desc_cols": [2],
            "ext_data_cols": [3],
            "base_act_data_cols": [3],
            "base_cmp_data_cols": [4],
        },
        "#N07.4": {
            "desc_cols": [2],
            "ext_data_cols": [3],
            "base_act_data_cols": [3],
            "base_cmp_data_cols": [4],
        },
        "#N08.1": {
            "desc_cols": [1],
            "ext_data_cols": [2, 3],
            "base_act_data_cols": [2, 3],
            "base_cmp_data_cols": [2, 3],
        },
        "#N08.2": {
            "desc_cols": [1],
            "ext_data_cols": [2, 3, 4, 5],
            "base_act_data_cols": [2, 3, 4, 5],
            "base_cmp_data_cols": [2, 3, 4, 5],
        },
        "#N09.1": {
            "desc_cols": [2],
            "ext_data_cols": [3, 4, 5],
            "base_act_data_cols": [3, 4, 5],
            "base_cmp_data_cols": [6, 7, 8],
        },
        "#N10.2": {
            "desc_cols": [2],
            "ext_data_cols": [3],
            "base_act_data_cols": [3],
            "base_cmp_data_cols": [4],
        },
        "#N10.3": {
            "desc_cols": [2],
            "ext_data_cols": [3],
            "base_act_data_cols": [3],
            "base_cmp_data_cols": [4],
        },
        "#N10.4": {
            "desc_cols": [2],
            "ext_data_cols": [3, 4, 5],
            "base_act_data_cols": [3, 4, 5],
            "base_cmp_data_cols": [3, 4, 5],
        },
        "#N16.1": {
            "desc_cols": [2],
            "ext_data_cols": [3, 4],
            "base_act_data_cols": [3, 4],
            "base_cmp_data_cols": [5, 6],
        },
        "#N15.5": {
            "desc_cols": [2, 3, 4, 5, 6],
            "ext_data_cols": [7, 8],
            "base_act_data_cols": [7, 8],
            "base_cmp_data_cols": [9, 10],
        },
        "#N15.6": {
            "desc_cols": [2, 3, 4, 5, 6],
            "ext_data_cols": [8, 9],
            "base_act_data_cols": [8, 9],
            "base_cmp_data_cols": [10, 11],
        },
        "#N15.7": {
            "desc_cols": [2, 3, 4, 5, 6, 7, 8, 9],
            "ext_data_cols": [10, 11],
            "base_act_data_cols": [10, 11],
            "base_cmp_data_cols": [10, 11],
        },
        "#N15.8": {
            "desc_cols": [2, 3, 4, 5, 6, 7, 8, 9, 10],
            "ext_data_cols": [11, 12],
            "base_act_data_cols": [11, 12],
            "base_cmp_data_cols": [13, 14],
        },
        "#N17.4": {
            "desc_cols": [2],
            "ext_data_cols": [3],
            "base_act_data_cols": [3],
            "base_cmp_data_cols": [4],
        },
        "#N18.3": {
            "desc_cols": [1],
            "ext_data_cols": [2],
            "base_act_data_cols": [2],
            "base_cmp_data_cols": [3],
        },
        "#N13.1": {
            "desc_cols": [2],
            "ext_data_cols": [3, 4],
            "base_act_data_cols": [3, 4],
            "base_cmp_data_cols": [5, 6],
        },
        "#N13.2": {
            "desc_cols": [2],
            "ext_data_cols": [3],
            "base_act_data_cols": [3],
            "base_cmp_data_cols": [4],
        },
        "#N13.3": {
            "desc_cols": [2],
            "ext_data_cols": [3],
            "base_act_data_cols": [3],
            "base_cmp_data_cols": [4],
        },
        "#N13.4": {
            "desc_cols": [2],
            "ext_data_cols": [3],
            "base_act_data_cols": [3],
            "base_cmp_data_cols": [4],
        },
        "#N13.5": {
            "desc_cols": [2],
            "ext_data_cols": [3],
            "base_act_data_cols": [3],
            "base_cmp_data_cols": [4],
        },
        "#N14.1": {
            "desc_cols": [3, 4, 5],
            "ext_data_cols": [6, 7],
            "base_act_data_cols": [6, 8],
            "base_cmp_data_cols": [7, 9],
        },
        "#N14.3": {
            "desc_cols": [2, 3, 4],
            "ext_data_cols": [5],
            "base_act_data_cols": [5],
            "base_cmp_data_cols": [6],
        },
        "#N20.1": {
            "desc_cols": [3],
            "ext_data_cols": [4, 5],
            "base_act_data_cols": [4, 5],
            "base_cmp_data_cols": [6, 7],
        },
        "#N20.2": {
            "desc_cols": [3],
            "ext_data_cols": [4],
            "base_act_data_cols": [4],
            "base_cmp_data_cols": [5],
        },
        "#N20.3": {
            "desc_cols": [3],
            "ext_data_cols": [4],
            "base_act_data_cols": [4],
            "base_cmp_data_cols": [5],
        },
        "#N20.4": {
            "desc_cols": [3],
            "ext_data_cols": [4],
            "base_act_data_cols": [4],
            "base_cmp_data_cols": [5],
        },
        "#N25.1": {
            "desc_cols": [2],
            "ext_data_cols": [3],
            "base_act_data_cols": [3],
            "base_cmp_data_cols": [4],
        },
        "#N25.2": {
            "desc_cols": [2],
            "ext_data_cols": [3],
            "base_act_data_cols": [3],
            "base_cmp_data_cols": [4],
        },
        "#N26.1": {
            "desc_cols": [1],
            "ext_data_cols": [2, 3, 4],
            "base_act_data_cols": [2, 3, 4],
            "base_cmp_data_cols": [6, 7, 8],
        },
        "#N26.2": {
            "desc_cols": [1],
            "ext_data_cols": [2, 3, 4],
            "base_act_data_cols": [2, 3, 4],
            "base_cmp_data_cols": [6, 7, 8],
        },
        "#N26.3": {
            "desc_cols": [1],
            "ext_data_cols": [2, 3, 4],
            "base_act_data_cols": [2, 3, 4],
            "base_cmp_data_cols": [6, 7, 8],
        },
        "#N27.1": {
            "desc_cols": [5],
            "ext_data_cols": [6, 7, 8],
            "base_act_data_cols": [6, 7, 8],
            "base_cmp_data_cols": [6, 7, 8],
        },
        "#N27.2": {
            "desc_cols": [5],
            "ext_data_cols": [6, 7, 8],
            "base_act_data_cols": [6, 7, 8],
            "base_cmp_data_cols": [6, 7, 8],
        },
        "#N27.3": {
            "desc_cols": [5, 6, 9, 10],
            "ext_data_cols": [7, 8],
            "base_act_data_cols": [7, 8],
            "base_cmp_data_cols": [7, 8],
        },
    }

    @classmethod
    def get_code_locations(cls, ws):
        locs = {}
        import re
        for r in range(1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                val = ws.cell(r, c).value
                if val and isinstance(val, str):
                    m = re.search(r'(#N\d+\.\d+)', val)
                    if m:
                        code = m.group(1)
                        if code not in locs:
                            locs[code] = {"row": r, "col": c, "val": val.strip()}
        return locs

    @classmethod
    def _infer_dynamic_table_config(cls, ext_ws, ext_r0: int, ext_c0: int, table_height: int, target_ws, target_r0: int, target_c0: int, target_r_end: int, is_comparative: bool):
        """
        Infiere dinámicamente las columnas descriptivas y numéricas para cualquier cuadro #Nxx.y
        que no esté en EXTERNAL_TABLE_REGISTRY o cuya estructura haya sido personalizada.
        """
        ext_m_cols = []
        for r in range(ext_r0 + 1, min(ext_r0 + 7, ext_r0 + table_height + 1, ext_ws.max_row + 1)):
            for c in range(ext_c0, min(ext_c0 + 14, ext_ws.max_column + 1)):
                v = ext_ws.cell(r, c).value
                if isinstance(v, str) and "m$" in v.strip().lower():
                    if c not in ext_m_cols:
                        ext_m_cols.append(c)
        if not ext_m_cols:
            for r in range(ext_r0 + 1, min(ext_r0 + table_height + 1, ext_ws.max_row + 1)):
                for c in range(ext_c0 + 1, min(ext_c0 + 14, ext_ws.max_column + 1)):
                    v = ext_ws.cell(r, c).value
                    if isinstance(v, (int, float)) or (isinstance(v, str) and v.startswith("=")):
                        if c not in ext_m_cols:
                            ext_m_cols.append(c)
        ext_m_cols.sort()
        first_ext_data = ext_m_cols[0] if ext_m_cols else (ext_c0 + 1)
        ext_desc_cols = list(range(ext_c0, max(ext_c0 + 1, first_ext_data)))
        ext_data_cols = ext_m_cols if ext_m_cols else [ext_c0 + 1]

        tgt_m_cols = []
        for r in range(target_r0 + 1, min(target_r0 + 7, target_r_end + 1, target_ws.max_row + 1)):
            for c in range(target_c0, min(target_c0 + 16, target_ws.max_column + 1)):
                v = target_ws.cell(r, c).value
                if isinstance(v, str) and "m$" in v.strip().lower():
                    if c not in tgt_m_cols:
                        tgt_m_cols.append(c)
        if not tgt_m_cols:
            col_delta = target_c0 - ext_c0
            tgt_m_cols = [c + col_delta for c in ext_data_cols if c + col_delta >= 1]
        tgt_m_cols.sort()

        first_tgt_data = tgt_m_cols[0] if tgt_m_cols else (target_c0 + len(ext_desc_cols))
        target_desc_cols = list(range(target_c0, max(target_c0 + 1, first_tgt_data)))

        n_ext = len(ext_data_cols)
        if len(tgt_m_cols) >= 2 * n_ext and n_ext > 0:
            target_data_cols = tgt_m_cols[n_ext:2 * n_ext] if is_comparative else tgt_m_cols[:n_ext]
        else:
            col_delta = target_c0 - ext_c0
            if is_comparative and len(tgt_m_cols) == n_ext:
                target_data_cols = [c + n_ext for c in tgt_m_cols]
            else:
                target_data_cols = tgt_m_cols[:n_ext] if tgt_m_cols else [c + col_delta for c in ext_data_cols if c + col_delta >= 1]

        return ext_desc_cols, target_desc_cols, ext_data_cols, target_data_cols

    @classmethod
    def _inject_table_data(cls, ext_ws, target_ws, is_comparative: bool = False, target_lang: str = "es", structure_only: bool = False):
        """
        Inyecta los datos y la estructura dinámica de ext_ws en target_ws alineando inteligentemente
        por código de cuadro (#Nxx.y), respetando si se inyecta como Período Actual o Período Comparativo
        y compensando desplazamientos de columnas y filas entre la plantilla maestra y el anexo externo.
        """
        from src.core.ifrs_glossary import translate_ifrs_term
        from src.reporting.note_generator import shift_formula_rows
        from openpyxl.utils import get_column_letter
        import re
        import datetime
        import unicodedata
        is_en = (str(target_lang).lower() == "en")

        def _norm_concept(val) -> str:
            if val is None:
                return ""
            s = str(val).strip().lower()
            s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
            s = re.sub(r"\s+", " ", s).rstrip(":.")
            return s

        def _is_header_value(val, allow_dates: bool = False) -> bool:
            if val is None:
                return False
            if isinstance(val, (datetime.datetime, datetime.date)):
                return not allow_dates
            if isinstance(val, str):
                v_low = val.strip().lower()
                if v_low in ("m$", "ms", "activos", "pasivos", "corriente", "no corriente", "concepto", "detalle"):
                    return True
                if not allow_dates and any(h in v_low for h in ["m$", "31-", "30-", "01-", "31.", "30.", "01.", "202"]):
                    return True
                if allow_dates and "m$" in v_low:
                    return True
            return False

        def _is_footer_or_summary_label(norm_str: str) -> bool:
            if not norm_str:
                return False
            return any(
                norm_str.startswith(k) or k in norm_str
                for k in (
                    "total", "sub total", "subtotal", "sub-total", "sub- total",
                    "saldo final", "tasa efectiva", "gasto por impuesto utilizando la tasa efectiva",
                    "intereses diferidos", "intereses devengados", "gastos de credito",
                    "subtotal bonds", "accrued interest", "total bonds"
                )
            )

        ext_codes = cls.get_code_locations(ext_ws)
        target_codes = cls.get_code_locations(target_ws)

        # Si el anexo externo contiene cuadros nuevos (#Nxx.y) que aún no existen en target_ws, agregarlos dinámicamente
        unmatched_ext_codes = [c for c in ext_codes if c not in target_codes]
        if unmatched_ext_codes and not is_comparative:
            for new_code in unmatched_ext_codes:
                e_info = ext_codes[new_code]
                e_r0 = e_info["row"]
                t_height = 0
                empty_cnt = 0
                for off in range(1, 45):
                    rc = e_r0 + off
                    if rc > ext_ws.max_row:
                        break
                    if any(isinstance(ext_ws.cell(rc, cc).value, str) and re.search(r'#N\d+\.\d+', ext_ws.cell(rc, cc).value) for cc in range(1, min(ext_ws.max_column + 1, 15))):
                        break
                    if any(ext_ws.cell(rc, cc).value is not None for cc in range(1, min(ext_ws.max_column + 1, 15))):
                        t_height = off
                        empty_cnt = 0
                    else:
                        empty_cnt += 1
                        if empty_cnt >= 3:
                            break
                start_append_r = target_ws.max_row + 2
                row_delta = start_append_r - e_r0
                for off in range(0, t_height + 1):
                    r_src = e_r0 + off
                    r_dst = start_append_r + off
                    for c_src in range(1, min(ext_ws.max_column + 1, 16)):
                        v_src = ext_ws.cell(r_src, c_src).value
                        if v_src is not None:
                            if isinstance(v_src, str) and v_src.startswith("="):
                                target_ws.cell(r_dst, c_src).value = shift_formula_rows(v_src, row_delta, min_row=e_r0)
                            else:
                                target_ws.cell(r_dst, c_src).value = v_src
            target_codes = cls.get_code_locations(target_ws)

        matched_codes = [c for c in ext_codes if c in target_codes]
        if not matched_codes:
            if not is_comparative and not structure_only:
                cls._copy_sheet_data(ext_ws, target_ws, target_col_offset=0, target_lang=target_lang)
            return

        for code in matched_codes:
            ext_codes = cls.get_code_locations(ext_ws)
            target_codes = cls.get_code_locations(target_ws)
            if code not in ext_codes or code not in target_codes:
                continue
            ext_info = ext_codes[code]
            target_info = target_codes[code]
            if "[COMPARATIVO]" in str(target_info.get("val", "")).upper() or "[COMPARATIVO]" in str(ext_info.get("val", "")).upper():
                continue
            ext_r0 = ext_info["row"]
            target_r0 = target_info["row"]

            # Determinar altura de la tabla en ext_ws (permitiendo 1 o 2 filas vacías previas a Totales o sub-bloques)
            table_height = 0
            consecutive_empty = 0
            for off in range(1, 50):
                r_curr = ext_r0 + off
                if r_curr > ext_ws.max_row:
                    break
                hit_next_ext_code = False
                for c_chk in range(1, min(ext_ws.max_column + 1, 15)):
                    v_check = ext_ws.cell(r_curr, c_chk).value
                    if v_check and isinstance(v_check, str) and re.search(r'#N\d+\.\d+', v_check):
                        hit_next_ext_code = True
                        break
                if hit_next_ext_code:
                    break
                has_content = any(ext_ws.cell(r_curr, c).value is not None for c in range(1, min(ext_ws.max_column + 1, 15)))
                if has_content:
                    table_height = off
                    consecutive_empty = 0
                else:
                    consecutive_empty += 1
                    if consecutive_empty >= 3:
                        break

            if table_height == 0:
                continue

            # Determinar frontera inferior de esta tabla en target_ws
            target_r_end = min(target_ws.max_row, target_r0 + max(table_height + 6, 35))
            for r_scan in range(target_r0 + 1, target_r_end + 1):
                hit_boundary = False
                for c_chk in range(1, min(target_ws.max_column + 1, 15)):
                    vt_chk = target_ws.cell(r_scan, c_chk).value
                    if vt_chk and isinstance(vt_chk, str):
                        vt_str = vt_chk.strip()
                        if re.search(r'#N\d+\.\d+', vt_str):
                            hit_boundary = True
                            break
                        if c_chk == target_info["col"] and re.match(r'^(?:\d+|[a-z])[\)\.\-]\s+', vt_str.lower()):
                            hit_boundary = True
                            break
                if hit_boundary:
                    target_r_end = r_scan - 1
                    break

            cfg = cls.EXTERNAL_TABLE_REGISTRY.get(code)
            if cfg is not None:
                base_desc_cols = cfg.get("desc_cols", [])
                base_ext_data_cols = cfg.get("ext_data_cols", [])
                base_target_data_cols = cfg.get("base_cmp_data_cols", []) if is_comparative else cfg.get("base_act_data_cols", [])
                if base_desc_cols:
                    base_ref_col = base_desc_cols[0]
                    ext_delta = ext_info["col"] - base_ref_col
                    target_delta = target_info["col"] - base_ref_col
                    ext_desc_cols = [c + ext_delta for c in base_desc_cols if c + ext_delta >= 1]
                    target_desc_cols = [c + target_delta for c in base_desc_cols if c + target_delta >= 1]
                    ext_data_cols = [c + ext_delta for c in base_ext_data_cols if c + ext_delta >= 1]
                    target_data_cols = [c + target_delta for c in base_target_data_cols if c + target_delta >= 1]
                else:
                    ext_desc_cols, target_desc_cols, ext_data_cols, target_data_cols = cls._infer_dynamic_table_config(
                        ext_ws, ext_r0, ext_info["col"], table_height, target_ws, target_r0, target_info["col"], target_r_end, is_comparative
                    )
            else:
                if structure_only:
                    continue
                ext_desc_cols, target_desc_cols, ext_data_cols, target_data_cols = cls._infer_dynamic_table_config(
                    ext_ws, ext_r0, ext_info["col"], table_height, target_ws, target_r0, target_info["col"], target_r_end, is_comparative
                )

            # Si la tabla en ext_ws tiene múltiples bloques con "Total" (ej. #N14.1 Cuentas por cobrar + Cuentas por pagar)
            # o más altura que el espacio disponible en target_ws antes de la siguiente nota, expandir target_ws
            ext_total_offs = []
            for off in range(1, table_height + 1):
                ve_lbl = _norm_concept(ext_ws.cell(ext_r0 + off, ext_info["col"]).value)
                if ve_lbl.startswith("total"):
                    ext_total_offs.append(off)

            tgt_total_rows = []
            for r_t in range(target_r0 + 1, target_r_end + 1):
                vt_lbl = _norm_concept(target_ws.cell(r_t, target_info["col"]).value)
                if vt_lbl.startswith("total"):
                    tgt_total_rows.append(r_t)

            if len(ext_total_offs) > len(tgt_total_rows) and target_r0 + table_height > target_r_end:
                extra_needed = (target_r0 + table_height) - target_r_end
                insert_pos = target_r_end + 1
                target_ws.insert_rows(insert_pos, amount=extra_needed)
                for r_below in range(insert_pos + extra_needed, target_ws.max_row + 1):
                    for c_below in range(1, target_ws.max_column + 1):
                        vb = target_ws.cell(r_below, c_below).value
                        if isinstance(vb, str) and vb.startswith("="):
                            target_ws.cell(r_below, c_below).value = shift_formula_rows(vb, extra_needed, min_row=insert_pos)
                target_r_end += extra_needed

            # Construir mapeo inteligente de filas (r_ext -> r_target)
            row_mapping = {}
            used_target_rows = set()

            if len(target_desc_cols) > 1 and len(ext_total_offs) <= 1:
                # Tabla de detalle multicolumna (ej. #N04.3, #N14.1 simple, #N14.3, #N15.5, #N15.6, #N15.7, #N15.8, #N27.3)
                # 1. Clasificar filas en ext_ws y target_ws en: Cabeceras, Detalle y Pie (Totales/Subtotales/Intereses)
                ext_detail_rows = []
                ext_footer_rows = []
                seen_ext_hdr = False
                for off in range(1, table_height + 1):
                    r_e = ext_r0 + off
                    is_h = any(c_e <= ext_ws.max_column and _is_header_value(ext_ws.cell(r_e, c_e).value) for c_e in ext_data_cols)
                    if not is_h:
                        v_first_e = _norm_concept(ext_ws.cell(r_e, ext_info["col"]).value)
                        if v_first_e in ("institucion", "banco", "detalle", "nombre entidad", "rut", "r.u.t", "instrument type", "entidad"):
                            is_h = True
                    if is_h and not ext_detail_rows and not ext_footer_rows:
                        seen_ext_hdr = True
                        continue
                    # Verificar si alguna celda descriptiva o de datos tiene etiqueta de pie/subtotal/total (ej. 'Total' en col 7 de #N27.3)
                    f_lbl = ""
                    for c_ed in ext_desc_cols + ext_data_cols:
                        if c_ed <= ext_ws.max_column:
                            cand = _norm_concept(ext_ws.cell(r_e, c_ed).value)
                            if _is_footer_or_summary_label(cand):
                                f_lbl = cand
                                break
                    has_sum_f = any(
                        c_e <= ext_ws.max_column and isinstance(ext_ws.cell(r_e, c_e).value, str) and "SUM" in str(ext_ws.cell(r_e, c_e).value).upper()
                        for c_e in ext_data_cols
                    )
                    if f_lbl or has_sum_f:
                        ext_footer_rows.append((r_e, f_lbl or "__subtotal_sum__"))
                    elif seen_ext_hdr or off >= 3:
                        ext_detail_rows.append(r_e)

                tgt_detail_rows = []
                tgt_footer_rows = []
                seen_tgt_hdr = False
                # Detectar primera fila cubierta por =SUM(...) en target_ws para evitar filas vacías huérfanas tras M$ (ej. R66 en #N15.6)
                min_sum_start_row = None
                for r_t in range(target_r0 + 1, target_r_end + 1):
                    for c_t in target_data_cols:
                        if c_t <= target_ws.max_column:
                            vt_f = target_ws.cell(r_t, c_t).value
                            if isinstance(vt_f, str) and vt_f.upper().startswith("=SUM("):
                                m_rng = re.search(r'[A-Za-z]+(\d+):[A-Za-z]+(\d+)', vt_f)
                                if m_rng:
                                    r_start_cand = int(m_rng.group(1))
                                    if target_r0 < r_start_cand <= r_t:
                                        if min_sum_start_row is None or r_start_cand < min_sum_start_row:
                                            min_sum_start_row = r_start_cand

                for r_t in range(target_r0 + 1, target_r_end + 1):
                    is_h = any(c_t <= target_ws.max_column and _is_header_value(target_ws.cell(r_t, c_t).value) for c_t in target_data_cols)
                    if not is_h:
                        v_first_t = _norm_concept(target_ws.cell(r_t, target_info["col"]).value)
                        if v_first_t in ("institucion", "banco", "detalle", "nombre entidad", "rut", "r.u.t", "instrument type", "entidad"):
                            is_h = True
                    if is_h and not tgt_detail_rows and not tgt_footer_rows:
                        seen_tgt_hdr = True
                        continue
                    f_lbl_t = ""
                    for c_td in target_desc_cols + target_data_cols:
                        if c_td <= target_ws.max_column:
                            cand_t = _norm_concept(target_ws.cell(r_t, c_td).value)
                            if _is_footer_or_summary_label(cand_t):
                                f_lbl_t = cand_t
                                break
                    has_sum_t = any(
                        c_t <= target_ws.max_column and isinstance(target_ws.cell(r_t, c_t).value, str) and "SUM" in str(target_ws.cell(r_t, c_t).value).upper()
                        for c_t in target_data_cols
                    )
                    if f_lbl_t or has_sum_t:
                        tgt_footer_rows.append((r_t, f_lbl_t or "__subtotal_sum__"))
                    elif (seen_tgt_hdr or (r_t - target_r0) >= 3) and not tgt_footer_rows:
                        if min_sum_start_row is not None and r_t < min_sum_start_row:
                            has_any_t = any(target_ws.cell(r_t, c).value is not None for c in target_desc_cols + target_data_cols if c <= target_ws.max_column)
                            if not has_any_t:
                                continue
                        tgt_detail_rows.append(r_t)

                # Si ext_ws tiene más filas de detalle con contenido que la capacidad de detalle de target_ws, expandir antes del primer pie
                active_ext_details = [
                    r_e for r_e in ext_detail_rows
                    if any(
                        (c <= ext_ws.max_column and ext_ws.cell(r_e, c).value is not None and str(ext_ws.cell(r_e, c).value).strip() not in ("", "-", "0", "0.0"))
                        for c in ext_desc_cols + ext_data_cols
                    )
                ]
                if len(active_ext_details) > len(tgt_detail_rows) and tgt_detail_rows:
                    extra_det = len(active_ext_details) - len(tgt_detail_rows)
                    ins_at = tgt_detail_rows[-1] + 1
                    target_ws.insert_rows(ins_at, amount=extra_det)
                    for r_below in range(ins_at + extra_det, target_ws.max_row + 1):
                        for c_below in range(1, target_ws.max_column + 1):
                            vb = target_ws.cell(r_below, c_below).value
                            if isinstance(vb, str) and vb.startswith("="):
                                target_ws.cell(r_below, c_below).value = shift_formula_rows(vb, extra_det, min_row=ins_at)
                    # Expandir el rango =SUM(...) en la fila de subtotal/total para abarcar las nuevas filas insertadas
                    old_last_det = tgt_detail_rows[-1]
                    new_last_det = old_last_det + extra_det
                    for r_f_old, _ in tgt_footer_rows:
                        r_f_new = r_f_old + extra_det
                        for c_col in range(1, target_ws.max_column + 1):
                            vf = target_ws.cell(r_f_new, c_col).value
                            if isinstance(vf, str) and vf.upper().startswith("=SUM("):
                                target_ws.cell(r_f_new, c_col).value = re.sub(
                                    rf'(:[A-Za-z]+){old_last_det}\b',
                                    rf'\g<1>{new_last_det}',
                                    vf
                                )
                    tgt_detail_rows.extend(range(ins_at, ins_at + extra_det))
                    tgt_footer_rows = [(r_f + extra_det, lbl_f) for (r_f, lbl_f) in tgt_footer_rows]
                    target_r_end += extra_det

                # Mapear filas de detalle posicionalmente
                for idx_det, r_e in enumerate(ext_detail_rows):
                    if idx_det < len(tgt_detail_rows):
                        row_mapping[r_e] = tgt_detail_rows[idx_det]
                        used_target_rows.add(tgt_detail_rows[idx_det])

                # Mapear filas de pie/resumen por concepto
                for r_e_f, lbl_e_f in ext_footer_rows:
                    for r_t_f, lbl_t_f in tgt_footer_rows:
                        if r_t_f not in used_target_rows and (lbl_e_f == lbl_t_f or (lbl_e_f[:15] and lbl_e_f[:15] == lbl_t_f[:15])):
                            row_mapping[r_e_f] = r_t_f
                            used_target_rows.add(r_t_f)
                            break

                # Limpiar filas de detalle sobrantes en target_ws que en ext_ws están vacías (cuando no es comparativo)
                if not is_comparative and not structure_only:
                    for idx_det, r_t_det in enumerate(tgt_detail_rows):
                        r_e_det = ext_detail_rows[idx_det] if idx_det < len(ext_detail_rows) else None
                        ext_has_desc = False
                        if r_e_det is not None:
                            ext_has_desc = any(
                                c_ed <= ext_ws.max_column and ext_ws.cell(r_e_det, c_ed).value is not None and str(ext_ws.cell(r_e_det, c_ed).value).strip() != ""
                                for c_ed in ext_desc_cols
                            )
                        if not ext_has_desc:
                            for c_td in target_data_cols:
                                if c_td <= target_ws.max_column:
                                    tc = target_ws.cell(r_t_det, c_td)
                                    if not (isinstance(tc, MergedCell) or tc.__class__.__name__ == 'MergedCell'):
                                        tc.value = None
                            # Si todas las columnas numéricas de la fila están en 0 o vacías, limpiar también los textos placeholder
                            any_other_num = False
                            for c_chk_n in range(max(target_desc_cols) + 1, min(target_ws.max_column + 1, max(target_desc_cols) + 8)):
                                v_n = target_ws.cell(r_t_det, c_chk_n).value
                                if isinstance(v_n, (int, float)) and abs(v_n) > 0.001:
                                    any_other_num = True
                                    break
                            if not any_other_num:
                                for c_td in target_desc_cols:
                                    if c_td <= target_ws.max_column:
                                        tc = target_ws.cell(r_t_det, c_td)
                                        if not (isinstance(tc, MergedCell) or tc.__class__.__name__ == 'MergedCell'):
                                            tc.value = None

            # Si es una tabla de una sola columna de concepto (ej. #N13.1..#N13.5, #N25.1, #N25.2), indexar las filas de destino por concepto
            target_concept_rows = []
            if len(target_desc_cols) == 1:
                c_t_desc = target_desc_cols[0]
                for r_t in range(target_r0 + 1, target_r_end + 1):
                    v_t = target_ws.cell(r_t, c_t_desc).value
                    if isinstance(v_t, str) and v_t.strip() and not v_t.strip().startswith("="):
                        norm_t = _norm_concept(v_t)
                        if norm_t and norm_t not in ("concepto", "detalle", "m$"):
                            target_concept_rows.append((r_t, norm_t))

            for off in range(1, table_height + 1):
                r_ext = ext_r0 + off
                if r_ext in row_mapping:
                    r_target = row_mapping[r_ext]
                elif len(target_desc_cols) > 1 and len(ext_total_offs) <= 1:
                    continue
                else:
                    r_target = target_r0 + off
                    matched_by_concept = False
                    if len(ext_desc_cols) == 1 and target_concept_rows:
                        v_e_desc = ext_ws.cell(r_ext, ext_desc_cols[0]).value
                        norm_e = _norm_concept(v_e_desc)
                        if norm_e:
                            candidates = [
                                r_t for (r_t, norm_t) in target_concept_rows
                                if r_t not in used_target_rows and (
                                    norm_t == norm_e or
                                    (len(norm_e) >= 20 and len(norm_t) >= 20 and (norm_t.startswith(norm_e[:25]) or norm_e.startswith(norm_t[:25])))
                                )
                            ]
                            if candidates:
                                best_r_t = min(candidates, key=lambda rt: abs((rt - target_r0) - off))
                                if abs((best_r_t - target_r0) - off) <= 4:
                                    r_target = best_r_t
                                    used_target_rows.add(best_r_t)
                                    matched_by_concept = True
                    if not matched_by_concept and len(ext_desc_cols) == 1:
                        # No permitir que una fila de control no emparejada (ej. 'reclasificación' en #N13.1) pise una fila ya usada o de Total
                        if r_target in used_target_rows:
                            continue
                        vt_existing = _norm_concept(target_ws.cell(r_target, target_desc_cols[0]).value) if r_target <= target_ws.max_row else ""
                        ve_current = _norm_concept(ext_ws.cell(r_ext, ext_desc_cols[0]).value)
                        if _is_footer_or_summary_label(vt_existing) and not _is_footer_or_summary_label(ve_current):
                            continue
                        used_target_rows.add(r_target)

                if r_target > target_r_end or r_target > target_ws.max_row:
                    continue

                # Para tablas multi-sección expandidas (ej. #N14.1 con Cuentas por cobrar y Cuentas por pagar),
                # copiar cabeceras de sub-sección y fórmulas =SUM nuevas en las filas expandidas
                is_hdr_row = False
                for c_e in ext_data_cols:
                    if c_e <= ext_ws.max_column and _is_header_value(ext_ws.cell(r_ext, c_e).value):
                        is_hdr_row = True
                        break
                if not is_hdr_row:
                    for c_t in target_data_cols:
                        if c_t <= target_ws.max_column and _is_header_value(target_ws.cell(r_target, c_t).value):
                            is_hdr_row = True
                            break
                if is_hdr_row:
                    if len(ext_total_offs) > 1 and not is_comparative:
                        for idx_d, c_ext_desc in enumerate(ext_desc_cols):
                            if idx_d < len(target_desc_cols):
                                ve_h = ext_ws.cell(r_ext, c_ext_desc).value
                                if ve_h is not None and target_ws.cell(r_target, target_desc_cols[idx_d]).value is None:
                                    target_ws.cell(r_target, target_desc_cols[idx_d]).value = ve_h
                        for idx, c_ext in enumerate(ext_data_cols):
                            if idx < len(target_data_cols):
                                ve_h = ext_ws.cell(r_ext, c_ext).value
                                if ve_h is not None and target_ws.cell(r_target, target_data_cols[idx]).value is None:
                                    target_ws.cell(r_target, target_data_cols[idx]).value = ve_h
                    continue

                if ext_desc_cols or target_desc_cols:
                    has_ext_desc = any(
                        c_ed <= ext_ws.max_column and ext_ws.cell(r_ext, c_ed).value is not None and str(ext_ws.cell(r_ext, c_ed).value).strip() != ""
                        for c_ed in ext_desc_cols + ext_data_cols
                    )
                    has_tgt_desc = any(
                        c_td <= target_ws.max_column and target_ws.cell(r_target, c_td).value is not None and str(target_ws.cell(r_target, c_td).value).strip() != ""
                        for c_td in target_desc_cols + target_data_cols
                    )
                    if not has_ext_desc and not has_tgt_desc:
                        continue

                # 1. Columnas descriptivas (bancos, monedas, fechas de vencimiento/inicio/término, conceptos)
                allow_desc_dates = len(ext_desc_cols) > 1
                for idx_d, c_ext_desc in enumerate(ext_desc_cols):
                    if idx_d >= len(target_desc_cols):
                        break
                    c_target_desc = target_desc_cols[idx_d]
                    if c_ext_desc <= ext_ws.max_column and c_target_desc <= target_ws.max_column:
                        v_ext = ext_ws.cell(r_ext, c_ext_desc).value
                        if v_ext is not None and not (isinstance(v_ext, str) and v_ext.startswith("=")) and not _is_header_value(v_ext, allow_dates=allow_desc_dates):
                            target_cell = target_ws.cell(r_target, c_target_desc)
                            if isinstance(target_cell, MergedCell) or target_cell.__class__.__name__ == 'MergedCell':
                                continue
                            curr_t_val = target_cell.value
                            should_write_desc = (
                                not is_comparative or
                                curr_t_val is None or
                                str(curr_t_val).strip() == "" or
                                curr_t_val == 0 or
                                (isinstance(curr_t_val, str) and curr_t_val.strip().lower() == "fondos mutuos" and c_target_desc > 1)
                            )
                            if should_write_desc:
                                if is_en and isinstance(v_ext, str) and v_ext.strip():
                                    target_cell.value = translate_ifrs_term(v_ext.strip(), target_lang="en")
                                else:
                                    target_cell.value = v_ext

                if structure_only:
                    continue

                # 2. Columnas numéricas / datos
                row_label_norm = " ".join(
                    _norm_concept(target_ws.cell(r_target, c_lbl).value)
                    for c_lbl in (target_desc_cols + target_data_cols)
                    if c_lbl <= target_ws.max_column and isinstance(target_ws.cell(r_target, c_lbl).value, str) and not str(target_ws.cell(r_target, c_lbl).value).startswith("=")
                )
                is_total_or_rate_row = any(
                    kw in row_label_norm
                    for kw in ("total", "sub total", "subtotal", "sub-total", "sub- total", "saldo final", "tasa efectiva", "gasto por impuesto utilizando la tasa efectiva", "subtotal bonds", "total bonds")
                )

                for idx, c_ext in enumerate(ext_data_cols):
                    if idx < len(target_data_cols):
                        c_target = target_data_cols[idx]
                        if c_ext <= ext_ws.max_column and c_target <= target_ws.max_column:
                            val_ext = ext_ws.cell(r_ext, c_ext).value
                            target_cell = target_ws.cell(r_target, c_target)
                            if isinstance(target_cell, MergedCell) or target_cell.__class__.__name__ == 'MergedCell':
                                continue
                            dest_val = target_cell.value

                            if _is_header_value(dest_val) or _is_header_value(val_ext):
                                continue
                            if _is_footer_or_summary_label(_norm_concept(dest_val)) or _is_footer_or_summary_label(_norm_concept(val_ext)):
                                continue

                            # Preservar fórmulas de totales/subtotales/tasas/sensibilidad en destino (=SUM..., =+E9*-0.1),
                            # pero permitir sobrescribir fórmulas de arrastre de saldo inicial (ej. =+D42 en #N13.2) si el anexo trae el dato
                            if isinstance(dest_val, str) and dest_val.startswith("="):
                                dest_up = dest_val.upper()
                                if "SUM" in dest_up or "*" in dest_up or is_total_or_rate_row:
                                    continue

                            if isinstance(val_ext, str) and val_ext.startswith("="):
                                continue

                            if isinstance(val_ext, str):
                                val_stripped = val_ext.strip()
                                if val_stripped in ("-", "—", "–"):
                                    val_ext = 0

                            if val_ext is not None and str(val_ext).strip() != "":
                                try:
                                    clean_v = str(val_ext).strip().replace(",", "") if isinstance(val_ext, str) and re.match(r'^-?\d{1,3}(?:,\d{3})+(?:\.\d+)?$', str(val_ext).strip()) else val_ext
                                    target_cell.value = float(clean_v) if isinstance(clean_v, (int, float)) else float(str(clean_v).strip())
                                except Exception:
                                    target_cell.value = val_ext
                            else:
                                if not is_total_or_rate_row:
                                    target_cell.value = None

    @classmethod
    def inject_external_data(cls, excel_bytes_in: BytesIO, empresa: str, periodo_actual: str, periodo_comp: str = None, target_lang: str = "es", **kwargs) -> BytesIO:
        """
        Inyecta los datos de los anexos externos guardados por empresa para el período actual (y comparativo)
        en el libro evaluado de notas.
        Para grupos consolidados sin anexo explícito propio, inyecta automáticamente la suma de sus filiales
        (con excepción de Impuestos Diferidos y Saldos Intercompañía).
        Si aún no se ha cargado un anexo para el período actual, sincroniza dinámicamente la estructura
        (nuevos cuadros, filas o glosas) desde la Plantilla Externa Máster vigente.
        """
        ext_act_wb = cls.get_anexo_workbook(empresa, periodo_actual, data_only=True)
        ext_cmp_wb = cls.get_anexo_workbook(empresa, periodo_comp, data_only=True) if periodo_comp else None

        master_struct_wb = None
        if ext_act_wb is None and os.path.exists(cls.MASTER_EXTERNAL_TEMPLATE):
            try:
                master_struct_wb = openpyxl.load_workbook(cls.MASTER_EXTERNAL_TEMPLATE, data_only=True)
            except Exception:
                master_struct_wb = None

        excel_bytes_in.seek(0)
        out_wb = openpyxl.load_workbook(excel_bytes_in, data_only=False)

        # Si el anexo externo o la plantilla máster tiene hojas nuevas que no estaban en out_wb, incorporarlas
        ref_wb_for_sheets = ext_act_wb or master_struct_wb
        if ref_wb_for_sheets:
            from src.reporting.notes import is_technical_sheet
            for ext_sh in ref_wb_for_sheets.sheetnames:
                if ext_sh not in out_wb.sheetnames and not is_technical_sheet(ext_sh):
                    new_ws = out_wb.create_sheet(title=ext_sh)
                    cls._copy_sheet_data(ref_wb_for_sheets[ext_sh], new_ws, target_col_offset=0, target_lang=target_lang)

        for sheet_name in out_wb.sheetnames:
            target_ws = out_wb[sheet_name]
            act_ws = ext_act_wb[sheet_name] if (ext_act_wb and sheet_name in ext_act_wb.sheetnames) else None
            cmp_ws = ext_cmp_wb[sheet_name] if (ext_cmp_wb and sheet_name in ext_cmp_wb.sheetnames) else None
            struct_ws = master_struct_wb[sheet_name] if (master_struct_wb and sheet_name in master_struct_wb.sheetnames) else None

            # Regla especial para Activo Fijo (#N09.1)
            if sheet_name == "Activo Fijo":
                if act_ws and not cls.is_activo_fijo_injected(act_ws):
                    act_ws = None
                if cmp_ws and not cls.is_activo_fijo_injected(cmp_ws):
                    cmp_ws = None
                struct_ws = None

            # Regla especial para Segmentos (#N26.1, #N26.2, #N26.3):
            # Se calcula de forma 100% automática desde los EEFF consolidados en NoteGenerator,
            # por lo que no debe ser sobrescrita por plantillas externas vacías.
            if sheet_name.strip().lower() == "segmentos":
                act_ws = None
                cmp_ws = None
                struct_ws = None

            # 1. Inyectar datos del período comparativo (si existe)
            if cmp_ws:
                cls._inject_table_data(cmp_ws, target_ws, is_comparative=True, target_lang=target_lang)

            # 2. Inyectar datos del período actual (si existe) o sincronizar estructura desde la plantilla máster
            if act_ws:
                cls._inject_table_data(act_ws, target_ws, is_comparative=False, target_lang=target_lang)
            elif struct_ws:
                cls._inject_table_data(struct_ws, target_ws, is_comparative=False, target_lang=target_lang, structure_only=True)

            # 3. Procesar duplicación de cuadros comparativos para tablas con [COMPARATIVO]
            from src.reporting.note_generator import process_comparative_tables_in_ws
            process_comparative_tables_in_ws(target_ws, periodo_actual_str=periodo_actual, periodo_comp_str=periodo_comp, ext_cmp_ws=cmp_ws, ext_act_ws=act_ws)

        if ext_act_wb:
            ext_act_wb.close()
        if ext_cmp_wb:
            ext_cmp_wb.close()
        if master_struct_wb:
            master_struct_wb.close()

        out_stream = BytesIO()
        out_wb.save(out_stream)
        out_wb.close()
        out_stream.seek(0)
        return out_stream

    @classmethod
    def _copy_sheet_data(cls, source_ws, target_ws, target_col_offset=0, target_lang: str = "es"):
        from src.core.ifrs_glossary import translate_ifrs_term
        import re
        is_en = (str(target_lang).lower() == "en")
        for r in range(1, source_ws.max_row + 1):
            for c in range(1, source_ws.max_column + 1):
                val = source_ws.cell(r, c).value
                if val is not None:
                    if isinstance(val, str) and val.startswith("="):
                        continue
                    target_c = c + target_col_offset
                    if target_c <= target_ws.max_column:
                        dest_cell = target_ws.cell(r, target_c)
                        if isinstance(dest_cell, MergedCell) or dest_cell.__class__.__name__ == 'MergedCell':
                            continue
                        dest_val = dest_cell.value
                        if not (isinstance(dest_val, str) and dest_val.startswith("=")):
                            if isinstance(val, str) and "[COMPARATIVO]" in val.upper():
                                val = re.sub(r'\[COMPARATIVO\]', '', val, flags=re.IGNORECASE).strip()
                            if is_en and isinstance(val, str) and val.strip():
                                dest_cell.value = translate_ifrs_term(val.strip(), target_lang="en")
                            else:
                                dest_cell.value = val
