import os
import shutil
import openpyxl
from io import BytesIO
import datetime

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
            try:
                from src.reporting.notes import is_technical_sheet
                wb = openpyxl.load_workbook(path, read_only=True)
                sheets = [s for s in wb.sheetnames if not is_technical_sheet(s)]
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
                "sheets": sheets
            }
        return {
            "exists": False,
            "path": path,
            "filename": os.path.basename(path),
            "last_modified": "Sin registros",
            "size_kb": "0 KB",
            "mtime": 0,
            "sheets": []
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
    def get_company_anexos_dir(cls, empresa: str, periodo: str) -> str:
        canonical = cls.resolve_canonical_company_name(empresa)
        safe_empresa = "".join([c if c.isalnum() or c in (" ", "_", "-", "[", "]") else "_" for c in canonical]).strip()
        dir_path = os.path.join("data", "empresas", safe_empresa, "anexos", periodo)
        os.makedirs(dir_path, exist_ok=True)
        return dir_path

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
                    for r in range(1, ws.max_row + 1):
                        for c in range(1, ws.max_column + 1):
                            val = ws.cell(r, c).value
                            if val is not None and isinstance(val, (int, float)) and not isinstance(val, bool):
                                if val != 0:
                                    key = (sh, r, c)
                                    sheet_cell_sums[key] = sheet_cell_sums.get(key, 0.0) + float(val)
                sub_wb.close()
            except Exception as e:
                print(f"Error leyendo anexo de filial {sub}: {e}")

        try:
            agg_wb = openpyxl.load_workbook(cls.get_master_template_path(), data_only=False)
        except Exception:
            return None

        for (sh, r, c), total_val in sheet_cell_sums.items():
            if sh in agg_wb.sheetnames:
                target_ws = agg_wb[sh]
                target_cell = target_ws.cell(r, c)
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
        out_stream.seek(0)
        return out_stream

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
        Verifica si la hoja 'Activo Fijo' (#N09.1) en la plantilla externa tiene datos ingresados.
        Si las celdas C y D (bruto/depreciación) o E tienen valores > 0 o no vacíos, se considera inyectada.
        """
        if ext_ws is None:
            return False
        for r in range(6, 25):
            val_c = ext_ws.cell(r, 3).value
            val_d = ext_ws.cell(r, 4).value
            val_e = ext_ws.cell(r, 5).value
            for v in [val_c, val_d, val_e]:
                if v is not None and isinstance(v, (int, float)) and abs(v) > 0.001:
                    return True
        return False

    EXTERNAL_TABLE_REGISTRY = {
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
            "ext_data_cols": [7],
            "base_act_data_cols": [7],
            "base_cmp_data_cols": [8],
        },
        "#N15.6": {
            "desc_cols": [2, 3, 4, 5, 6],
            "ext_data_cols": [8],
            "base_act_data_cols": [8],
            "base_cmp_data_cols": [9],
        },
        "#N17.4": {
            "desc_cols": [2],
            "ext_data_cols": [3],
            "base_act_data_cols": [3],
            "base_cmp_data_cols": [4],
        },
        "#N13.1": {
            "desc_cols": [3],
            "ext_data_cols": [4, 5],
            "base_act_data_cols": [4, 5],
            "base_cmp_data_cols": [6, 7],
        },
        "#N13.2": {
            "desc_cols": [3],
            "ext_data_cols": [4],
            "base_act_data_cols": [4],
            "base_cmp_data_cols": [5],
        },
        "#N13.3": {
            "desc_cols": [3],
            "ext_data_cols": [4],
            "base_act_data_cols": [4],
            "base_cmp_data_cols": [5],
        },
        "#N13.4": {
            "desc_cols": [3],
            "ext_data_cols": [4],
            "base_act_data_cols": [4],
            "base_cmp_data_cols": [5],
        },
        "#N13.5": {
            "desc_cols": [3],
            "ext_data_cols": [4],
            "base_act_data_cols": [4],
            "base_cmp_data_cols": [5],
        },
        "#N14.1": {
            "desc_cols": [3, 4, 5],
            "ext_data_cols": [6],
            "base_act_data_cols": [6],
            "base_cmp_data_cols": [7],
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
    def _inject_table_data(cls, ext_ws, target_ws, is_comparative: bool = False, target_lang: str = "es"):
        """
        Inyecta los datos de ext_ws en target_ws alineando inteligentemente por código de cuadro (#Nxx.y),
        respetando si se inyecta como Período Actual o Período Comparativo.
        """
        from src.core.ifrs_glossary import translate_ifrs_term
        import re
        is_en = (str(target_lang).lower() == "en")

        ext_codes = cls.get_code_locations(ext_ws)
        target_codes = cls.get_code_locations(target_ws)

        matched_codes = [c for c in ext_codes if c in target_codes]
        if not matched_codes:
            if not is_comparative:
                cls._copy_sheet_data(ext_ws, target_ws, target_col_offset=0, target_lang=target_lang)
            return

        for code in matched_codes:
            ext_info = ext_codes[code]
            target_info = target_codes[code]
            ext_r0 = ext_info["row"]
            target_r0 = target_info["row"]

            cfg = cls.EXTERNAL_TABLE_REGISTRY.get(code)

            # Determinar altura de la tabla en ext_ws
            table_height = 0
            for off in range(1, 40):
                r_curr = ext_r0 + off
                if r_curr > ext_ws.max_row:
                    break
                v_check = ext_ws.cell(r_curr, ext_info["col"]).value
                if v_check and isinstance(v_check, str) and re.search(r'#N\d+\.\d+', v_check):
                    break
                has_content = any(ext_ws.cell(r_curr, c).value is not None for c in range(1, min(ext_ws.max_column + 1, 15)))
                if has_content:
                    table_height = off
                else:
                    if off > 15:
                        break

            if table_height == 0:
                continue

            desc_cols = cfg.get("desc_cols", []) if cfg else []
            ext_data_cols = cfg.get("ext_data_cols", []) if cfg else []
            target_data_cols = (cfg.get("base_cmp_data_cols", []) if is_comparative else cfg.get("base_act_data_cols", [])) if cfg else []

            if not target_data_cols and cfg is None:
                target_data_cols = ext_data_cols

            for off in range(1, table_height + 1):
                r_ext = ext_r0 + off
                r_target = target_r0 + off
                if r_target > target_ws.max_row:
                    break

                # 1. Columnas descriptivas (bancos, monedas, conceptos)
                for c_desc in desc_cols:
                    if c_desc <= ext_ws.max_column and c_desc <= target_ws.max_column:
                        v_ext = ext_ws.cell(r_ext, c_desc).value
                        if v_ext is not None and not (isinstance(v_ext, str) and v_ext.startswith("=")):
                            target_cell = target_ws.cell(r_target, c_desc)
                            curr_t_val = target_cell.value
                            # Rellenar si está vacío, es 0, o es placeholder genérico repetitivo ("fondos mutuos" en c_desc > 1)
                            if (curr_t_val is None or str(curr_t_val).strip() == "" or curr_t_val == 0 or
                                (isinstance(curr_t_val, str) and curr_t_val.strip().lower() == "fondos mutuos" and c_desc > 1)):
                                if is_en and isinstance(v_ext, str) and v_ext.strip():
                                    target_cell.value = translate_ifrs_term(v_ext.strip(), target_lang="en")
                                else:
                                    target_cell.value = v_ext

                # 2. Columnas numéricas / datos
                for idx, c_ext in enumerate(ext_data_cols):
                    if idx < len(target_data_cols):
                        c_target = target_data_cols[idx]
                        if c_ext <= ext_ws.max_column and c_target <= target_ws.max_column:
                            val_ext = ext_ws.cell(r_ext, c_ext).value
                            target_cell = target_ws.cell(r_target, c_target)
                            dest_val = target_cell.value

                            # Preservar fórmulas de totales en destino (=SUM...)
                            if isinstance(dest_val, str) and dest_val.startswith("="):
                                continue

                            # No copiar fórmulas del anexo externo directamente
                            if isinstance(val_ext, str) and val_ext.startswith("="):
                                continue

                            # Evitar sobreescribir con textos de cabecera (M$, fechas)
                            if isinstance(val_ext, str):
                                val_lower = val_ext.strip().lower()
                                if any(h in val_lower for h in ["m$", "31-", "30-", "01-", "202"]):
                                    continue

                            if val_ext is not None:
                                try:
                                    target_cell.value = float(val_ext) if isinstance(val_ext, (int, float)) else val_ext
                                except Exception:
                                    target_cell.value = val_ext

    @classmethod
    def inject_external_data(cls, excel_bytes_in: BytesIO, empresa: str, periodo_actual: str, periodo_comp: str = None, target_lang: str = "es", **kwargs) -> BytesIO:
        """
        Inyecta los datos de los anexos externos guardados por empresa para el período actual (y comparativo)
        en el libro evaluado de notas.
        Para grupos consolidados sin anexo explícito propio, inyecta automáticamente la suma de sus filiales
        (con excepción de Impuestos Diferidos y Saldos Intercompañía).
        """
        ext_act_wb = cls.get_anexo_workbook(empresa, periodo_actual, data_only=True)
        ext_cmp_wb = cls.get_anexo_workbook(empresa, periodo_comp, data_only=True) if periodo_comp else None

        if not ext_act_wb and not ext_cmp_wb:
            excel_bytes_in.seek(0)
            return excel_bytes_in

        excel_bytes_in.seek(0)
        out_wb = openpyxl.load_workbook(excel_bytes_in, data_only=False)

        for sheet_name in out_wb.sheetnames:
            target_ws = out_wb[sheet_name]
            act_ws = ext_act_wb[sheet_name] if (ext_act_wb and sheet_name in ext_act_wb.sheetnames) else None
            cmp_ws = ext_cmp_wb[sheet_name] if (ext_cmp_wb and sheet_name in ext_cmp_wb.sheetnames) else None

            # Regla especial para Activo Fijo (#N09.1)
            if sheet_name == "Activo Fijo":
                if act_ws and not cls.is_activo_fijo_injected(act_ws):
                    act_ws = None
                if cmp_ws and not cls.is_activo_fijo_injected(cmp_ws):
                    cmp_ws = None

            # 1. Inyectar datos del período comparativo (si existe)
            if cmp_ws:
                cls._inject_table_data(cmp_ws, target_ws, is_comparative=True, target_lang=target_lang)

            # 2. Inyectar datos del período actual (si existe)
            if act_ws:
                cls._inject_table_data(act_ws, target_ws, is_comparative=False, target_lang=target_lang)

            # 3. Procesar duplicación de cuadros comparativos para tablas con [COMPARATIVO]
            from src.reporting.note_generator import process_comparative_tables_in_ws
            process_comparative_tables_in_ws(target_ws, periodo_actual_str=periodo_actual, periodo_comp_str=periodo_comp, ext_cmp_ws=cmp_ws)

        if ext_act_wb:
            ext_act_wb.close()
        if ext_cmp_wb:
            ext_cmp_wb.close()

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
                        dest_val = dest_cell.value
                        if not (isinstance(dest_val, str) and dest_val.startswith("=")):
                            if isinstance(val, str) and "[COMPARATIVO]" in val.upper():
                                val = re.sub(r'\[COMPARATIVO\]', '', val, flags=re.IGNORECASE).strip()
                            if is_en and isinstance(val, str) and val.strip():
                                dest_cell.value = translate_ifrs_term(val.strip(), target_lang="en")
                            else:
                                dest_cell.value = val
