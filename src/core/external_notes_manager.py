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
        cls._AGG_ANEXO_CACHE.clear()
        try:
            from src.core.snapshot_vault_manager import _HYBRID_COMP_CACHE
            _HYBRID_COMP_CACHE.clear()
        except Exception:
            pass
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

    @classmethod
    def get_group_parent_company(cls, grupo_name: str) -> str | None:
        """Retorna el nombre de la empresa matriz (controladora) del grupo consolidado."""
        clean_group = str(grupo_name).strip()
        for prefix in ["[GRUPO] ", "[GRUPO]", "_GRUPO_ ", "_GRUPO_"]:
            if clean_group.startswith(prefix):
                clean_group = clean_group[len(prefix):].strip()
                break

        from src.models.database import SessionLocal
        from src.models.consolidacion import ConsolidationGroup

        db = SessionLocal()
        try:
            grupo_obj = db.query(ConsolidationGroup).filter(
                (ConsolidationGroup.nombre_grupo == clean_group) |
                (ConsolidationGroup.nombre_grupo == grupo_name)
            ).first()
            if grupo_obj and grupo_obj.empresa_matriz:
                return str(grupo_obj.empresa_matriz).strip()
        except Exception:
            pass
        finally:
            db.close()
        return None

    @classmethod
    def _is_excluded_agg_sheet(cls, sh_name: str) -> bool:
        if not sh_name:
            return True
        sh_norm = str(sh_name).strip().lower()
        if sh_norm in cls.EXCLUDED_FROM_AGGREGATION or sh_norm == "segmentos":
            return True
        return any(
            ex in sh_norm
            for ex in ("impuestos diferidos", "relacionadas", "ajuste", "bce", "calculo", "ppa", "db_data")
        )

    @staticmethod
    def _norm_concept_static(val) -> str:
        import unicodedata
        if val is None:
            return ""
        s = str(val).strip().lower()
        s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
        s = re.sub(r"\s+", " ", s).rstrip(":.")
        return s

    @staticmethod
    def _is_header_cell_static(val, allow_dates: bool = False) -> bool:
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

    @staticmethod
    def _is_footer_label_static(norm_str: str) -> bool:
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

    @classmethod
    def _build_desc_row_key(cls, desc_vals: list) -> tuple:
        """
        Construye la llave normalizada de una fila a partir de todas sus columnas descriptivas:
        - Textos: sin tildes/diacríticos, minúsculas y espacios limpios (sin alterar palabras/nombres).
        - Fechas: YYYY-MM-DD.
        - Números: representación canónica (7500000 == 7500000.0).
        - Celdas vacías o guiones: ''.
        """
        parts = []
        for v in desc_vals:
            if v is None:
                parts.append("")
            elif isinstance(v, bool):
                parts.append(str(v).lower())
            elif isinstance(v, (datetime.datetime, datetime.date)):
                parts.append(v.strftime("%Y-%m-%d"))
            elif isinstance(v, (int, float)):
                fv = float(v)
                if abs(fv - round(fv)) < 1e-7:
                    parts.append(str(int(round(fv))))
                else:
                    parts.append(f"{fv:.6f}".rstrip("0").rstrip("."))
            elif isinstance(v, str):
                s = v.strip()
                if not s or s in ("-", "—", "–"):
                    parts.append("")
                else:
                    ok_n, fv = cls._parse_numeric_cell(s)
                    if ok_n:
                        if abs(fv - round(fv)) < 1e-7:
                            parts.append(str(int(round(fv))))
                        else:
                            parts.append(f"{fv:.6f}".rstrip("0").rstrip("."))
                    else:
                        parts.append(cls._norm_concept_static(s))
            else:
                parts.append(cls._norm_concept_static(str(v)))
        return tuple(parts)

    @staticmethod
    def _parse_numeric_cell(val) -> tuple[bool, float]:
        if val is None or isinstance(val, bool) or isinstance(val, (datetime.datetime, datetime.date)):
            return False, 0.0
        if isinstance(val, (int, float)):
            return True, float(val)
        if isinstance(val, str):
            s = val.strip()
            if not s or s.startswith("="):
                return False, 0.0
            if s in ("-", "—", "–"):
                return True, 0.0
            try:
                clean_s = s.replace(",", "") if re.match(r'^-?\d{1,3}(?:,\d{3})+(?:\.\d+)?$', s) else s
                return True, float(clean_s)
            except Exception:
                return False, 0.0
        return False, 0.0

    @classmethod
    def _get_ws_table_height(cls, ws, r0: int) -> int:
        table_height = 0
        consecutive_empty = 0
        for off in range(1, 55):
            r_curr = r0 + off
            if r_curr > ws.max_row:
                break
            hit_next_code = False
            for c_chk in range(1, min(ws.max_column + 1, 15)):
                v_check = ws.cell(r_curr, c_chk).value
                if v_check and isinstance(v_check, str) and re.search(r'#N\d+\.\d+', v_check):
                    hit_next_code = True
                    break
            if hit_next_code:
                break
            has_content = any(ws.cell(r_curr, c).value is not None for c in range(1, min(ws.max_column + 1, 16)))
            if has_content:
                table_height = off
                consecutive_empty = 0
            else:
                consecutive_empty += 1
                if consecutive_empty >= 3:
                    break
        return table_height

    @classmethod
    def _resolve_ws_table_columns(cls, ws, code: str, r0: int, c0: int, table_height: int):
        cfg = cls.EXTERNAL_TABLE_REGISTRY.get(code)
        if cfg is not None and cfg.get("desc_cols"):
            base_desc_cols = cfg["desc_cols"]
            base_ext_data_cols = cfg.get("ext_data_cols", [])
            base_ref_col = base_desc_cols[0]
            delta = c0 - base_ref_col
            desc_cols = [c + delta for c in base_desc_cols if c + delta >= 1]
            data_cols = [c + delta for c in base_ext_data_cols if c + delta >= 1]

            # Refinar columnas de datos si existen subcabeceras exactas 'M$' desplazadas (ej. columna espaciadora en #N15.6)
            actual_m_cols = []
            for r_m in range(r0 + 1, min(r0 + 6, r0 + table_height + 1, ws.max_row + 1)):
                for c_m in range(c0 + 1, min(c0 + 15, ws.max_column + 1)):
                    vm = ws.cell(r_m, c_m).value
                    if isinstance(vm, str) and vm.strip().lower() in ("m$", "ms"):
                        if c_m not in actual_m_cols:
                            actual_m_cols.append(c_m)
            actual_m_cols.sort()
            if base_ext_data_cols and len(actual_m_cols) == len(base_ext_data_cols):
                data_cols = actual_m_cols
            return desc_cols, data_cols

        desc_cols, _, data_cols, _ = cls._infer_dynamic_table_config(
            ws, r0, c0, table_height, ws, r0, c0, min(ws.max_row, r0 + max(table_height, 1)), False
        )
        return desc_cols, data_cols

    @classmethod
    def _classify_detail_and_footer_rows(cls, ws, r0: int, c0: int, r_end: int, desc_cols: list, data_cols: list):
        detail_rows = []
        footer_rows = []
        seen_hdr = False
        last_hdr_r = None
        for r_curr in range(r0 + 1, min(r_end + 1, ws.max_row + 1)):
            is_h = any(c_d <= ws.max_column and cls._is_header_cell_static(ws.cell(r_curr, c_d).value) for c_d in data_cols)
            if not is_h:
                v_first = cls._norm_concept_static(ws.cell(r_curr, c0).value)
                if v_first in ("institucion", "banco", "detalle", "nombre entidad", "rut", "r.u.t", "instrument type", "entidad"):
                    is_h = True
            if is_h and not detail_rows and not footer_rows:
                seen_hdr = True
                last_hdr_r = r_curr
                continue
            f_lbl = ""
            for c_any in desc_cols + data_cols:
                if c_any <= ws.max_column:
                    cand = cls._norm_concept_static(ws.cell(r_curr, c_any).value)
                    if cls._is_footer_label_static(cand):
                        f_lbl = cand
                        break
            has_sum_f = any(
                c_d <= ws.max_column and isinstance(ws.cell(r_curr, c_d).value, str) and "SUM" in str(ws.cell(r_curr, c_d).value).upper()
                for c_d in data_cols
            )
            if f_lbl or has_sum_f:
                footer_rows.append((r_curr, f_lbl or "__subtotal_sum__"))
            elif (seen_hdr or (r_curr - r0) >= 3) and not footer_rows:
                if not detail_rows and last_hdr_r is not None and r_curr == last_hdr_r + 1:
                    is_blank_spacer = not any(
                        c_b <= ws.max_column and ws.cell(r_curr, c_b).value is not None and str(ws.cell(r_curr, c_b).value).strip() != ""
                        for c_b in desc_cols + data_cols
                    )
                    if is_blank_spacer and r_curr + 1 <= min(r_end, ws.max_row):
                        detail_rows_skip = True
                        if detail_rows_skip:
                            continue
                detail_rows.append(r_curr)
        return detail_rows, footer_rows

    _AGG_ANEXO_CACHE = {}

    @classmethod
    def build_aggregated_group_anexo(cls, grupo_name: str, periodo: str) -> BytesIO | None:
        """
        Suma automáticamente las planillas externas de las filiales del grupo alineando por código
        de cuadro (#Nxx.y), concepto de fila y desplazamiento relativo de columnas, con excepción
        explícita de Impuestos Diferidos y Saldos Intercompañía/Relacionadas.
        """
        from openpyxl.utils import get_column_letter
        from src.reporting.note_generator import (
            shift_formula_rows,
            _extract_movement_table_rows,
            _build_comp_column_mapping,
        )
        from src.reporting.notes import is_technical_sheet

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
        cache_key = ("v4_comp_hdr_map", str(grupo_name).strip(), cls.normalize_period_str(periodo), master_mtime, sub_mtimes)
        if cache_key in cls._AGG_ANEXO_CACHE:
            return BytesIO(cls._AGG_ANEXO_CACHE[cache_key])

        try:
            agg_wb = openpyxl.load_workbook(master_tpl_path, data_only=False)
        except Exception:
            return None

        matriz_name = cls.get_group_parent_company(grupo_name)
        matriz_canon = cls.resolve_canonical_company_name(matriz_name) if matriz_name else None

        sub_wbs = []
        for sub in active_subs:
            sub_path = cls.get_anexo_path(sub, periodo)
            try:
                swb = openpyxl.load_workbook(sub_path, data_only=True)
                sub_wbs.append((sub, cls.resolve_canonical_company_name(sub), swb))
            except Exception as e:
                print(f"Error leyendo anexo de filial {sub}: {e}")

        if not sub_wbs:
            agg_wb.close()
            return None

        try:
            # 1. Sincronizar hojas o cuadros #Nxx.y presentes en filiales pero ausentes en la plantilla externa máster (ej. #N16.1)
            for _, _, swb in sub_wbs:
                for sh in swb.sheetnames:
                    if is_technical_sheet(sh) or cls._is_excluded_agg_sheet(sh):
                        continue
                    sws = swb[sh]
                    if sh not in agg_wb.sheetnames:
                        new_ws = agg_wb.create_sheet(title=sh)
                        cls._copy_sheet_data(sws, new_ws, target_col_offset=0, target_lang="es")
                    else:
                        aws = agg_wb[sh]
                        agg_codes = cls.get_code_locations(aws)
                        sub_codes = cls.get_code_locations(sws)
                        for new_code, s_info in sub_codes.items():
                            if new_code not in agg_codes:
                                s_r0 = s_info["row"]
                                s_c0 = s_info["col"]
                                t_height = cls._get_ws_table_height(sws, s_r0)
                                if t_height <= 0:
                                    continue
                                start_append_r = aws.max_row + 2
                                col_shift = 1 - s_c0
                                s_desc_cols, s_data_cols = cls._resolve_ws_table_columns(sws, new_code, s_r0, s_c0, t_height)
                                first_data_dst_r = None
                                for off in range(0, t_height + 1):
                                    r_src = s_r0 + off
                                    r_dst = start_append_r + off
                                    lbl_norm = cls._norm_concept_static(sws.cell(r_src, s_c0).value)
                                    is_tot_line = lbl_norm.startswith("total")
                                    for c_src in range(s_c0, min(sws.max_column + 1, s_c0 + 14)):
                                        c_dst = c_src + col_shift
                                        if c_dst < 1:
                                            continue
                                        v_src = sws.cell(r_src, c_src).value
                                        if v_src is None:
                                            continue
                                        if c_src in s_data_cols and off >= 1 and not cls._is_header_cell_static(v_src):
                                            if is_tot_line and first_data_dst_r is not None and r_dst > first_data_dst_r:
                                                col_let = get_column_letter(c_dst)
                                                aws.cell(r_dst, c_dst).value = f"=SUM({col_let}{first_data_dst_r}:{col_let}{r_dst - 1})"
                                            else:
                                                if first_data_dst_r is None and not is_tot_line:
                                                    first_data_dst_r = r_dst
                                                aws.cell(r_dst, c_dst).value = None
                                        else:
                                            aws.cell(r_dst, c_dst).value = v_src
                                            if off >= 1 and c_src in s_desc_cols and not cls._is_header_cell_static(v_src) and not is_tot_line:
                                                if first_data_dst_r is None:
                                                    first_data_dst_r = r_dst
                                    if is_tot_line and first_data_dst_r is not None and r_dst > first_data_dst_r:
                                        for c_src in s_data_cols:
                                            c_dst = c_src + col_shift
                                            if c_dst >= 1 and aws.cell(r_dst, c_dst).value is None:
                                                col_let = get_column_letter(c_dst)
                                                aws.cell(r_dst, c_dst).value = f"=SUM({col_let}{first_data_dst_r}:{col_let}{r_dst - 1})"

            # 2. Limpiar valores de muestra en pestañas excluidas (Impuestos Diferidos / Relacionadas) para descarga limpia
            for sh in agg_wb.sheetnames:
                if cls._is_excluded_agg_sheet(sh) and not is_technical_sheet(sh):
                    ews = agg_wb[sh]
                    e_codes = cls.get_code_locations(ews)
                    for e_code, e_info in e_codes.items():
                        e_r0, e_c0 = e_info["row"], e_info["col"]
                        e_h = cls._get_ws_table_height(ews, e_r0)
                        if "[COMPARATIVO]" in str(e_info.get("val", "")).upper():
                            for r_cl in range(e_r0 + 1, min(ews.max_row + 1, e_r0 + e_h + 1)):
                                if any(cls._is_header_cell_static(ews.cell(r_cl, c_cl).value) for c_cl in range(e_c0, min(ews.max_column + 1, e_c0 + 12))):
                                    continue
                                if cls._is_footer_label_static(cls._norm_concept_static(ews.cell(r_cl, e_c0).value)):
                                    continue
                                for c_cl in range(e_c0, min(ews.max_column + 1, e_c0 + 12)):
                                    cell_cl = ews.cell(r_cl, c_cl)
                                    if not isinstance(cell_cl, MergedCell) and not (isinstance(cell_cl.value, str) and cell_cl.value.startswith("=")):
                                        cell_cl.value = None
                        else:
                            e_desc, e_data = cls._resolve_ws_table_columns(ews, e_code, e_r0, e_c0, e_h)
                            cols_to_clean = (e_desc + e_data) if len(e_desc) > 1 else e_data
                            for r_cl in range(e_r0 + 1, min(ews.max_row + 1, e_r0 + e_h + 1)):
                                if any(c_cl <= ews.max_column and cls._is_header_cell_static(ews.cell(r_cl, c_cl).value) for c_cl in e_data):
                                    continue
                                v_first = cls._norm_concept_static(ews.cell(r_cl, e_c0).value)
                                if v_first in ("rut", "r.u.t", "nombre entidad", "concepto", "detalle") or cls._is_footer_label_static(v_first):
                                    continue
                                for c_cl in cols_to_clean:
                                    if c_cl <= ews.max_column:
                                        cell_cl = ews.cell(r_cl, c_cl)
                                        if not isinstance(cell_cl, MergedCell) and not (isinstance(cell_cl.value, str) and cell_cl.value.startswith("=")):
                                            cell_cl.value = None

            # 3. Agregar cada hoja y cada cuadro #Nxx.y alineado por código y concepto
            for sh in agg_wb.sheetnames:
                if is_technical_sheet(sh) or cls._is_excluded_agg_sheet(sh):
                    continue
                agg_ws = agg_wb[sh]
                matching_subs = [(sn, sc, swb[sh]) for (sn, sc, swb) in sub_wbs if sh in swb.sheetnames]

                initial_codes = cls.get_code_locations(agg_ws)
                ordered_codes = [c for c, _ in sorted(initial_codes.items(), key=lambda kv: kv[1]["row"])]

                for code in ordered_codes:
                    agg_codes = cls.get_code_locations(agg_ws)
                    if code not in agg_codes:
                        continue
                    agg_info = agg_codes[code]
                    agg_r0 = agg_info["row"]
                    agg_c0 = agg_info["col"]

                    # Regla NIIF 10 para Patrimonio (#N20.1, #N20.2, #N20.3):
                    # El número de acciones y estructura accionaria consolidada corresponden exclusivamente a la Matriz.
                    if code in ("#N20.1", "#N20.2", "#N20.3"):
                        code_subs = []
                        for (sn, sc, sws) in matching_subs:
                            if not (matriz_canon and sc == matriz_canon):
                                continue
                            scodes = cls.get_code_locations(sws)
                            if code not in scodes:
                                continue
                            sr0, sc0 = scodes[code]["row"], scodes[code]["col"]
                            sh_tmp = cls._get_ws_table_height(sws, sr0)
                            _, sdata_tmp = cls._resolve_ws_table_columns(sws, code, sr0, sc0, sh_tmp)
                            has_nonzero_shares = False
                            for roff in range(1, sh_tmp + 1):
                                for cd_t in sdata_tmp:
                                    if cd_t <= sws.max_column:
                                        ok_sh, num_sh = cls._parse_numeric_cell(sws.cell(sr0 + roff, cd_t).value)
                                        if ok_sh and abs(num_sh) > 1.0:
                                            has_nonzero_shares = True
                                            break
                                if has_nonzero_shares:
                                    break
                            if has_nonzero_shares:
                                code_subs.append((sn, sc, sws))
                        if not code_subs:
                            # Conservar estructura accionaria de la matriz en la plantilla si la matriz no tiene acciones informadas en este subcuadro
                            continue
                    else:
                        code_subs = matching_subs

                    agg_height = cls._get_ws_table_height(agg_ws, agg_r0)
                    if agg_height <= 0:
                        continue
                    agg_r_end = agg_r0 + agg_height
                    is_comp_marker = "[COMPARATIVO]" in str(agg_info.get("val", "")).upper()

                    # --- CASO A: Tablas de movimiento/vencimiento con tag [COMPARATIVO] ---
                    if is_comp_marker:
                        max_agg_col = agg_c0
                        for r_sc in range(agg_r0, min(agg_ws.max_row + 1, agg_r_end + 1)):
                            for c_sc in range(agg_c0 + 1, agg_ws.max_column + 1):
                                if agg_ws.cell(r_sc, c_sc).value is not None and c_sc > max_agg_col:
                                    max_agg_col = c_sc

                        agg_concept_map = _extract_movement_table_rows(agg_ws, agg_r0, label_col=agg_c0)
                        for r_a in agg_concept_map.values():
                            for c_a in range(agg_c0 + 1, max_agg_col + 1):
                                cell_a = agg_ws.cell(r_a, c_a)
                                if isinstance(cell_a, MergedCell) or cell_a.__class__.__name__ == "MergedCell":
                                    continue
                                if isinstance(cell_a.value, str) and cell_a.value.startswith("="):
                                    continue
                                if cls._is_header_cell_static(cell_a.value):
                                    continue
                                cell_a.value = 0

                        for _, _, sub_ws in code_subs:
                            sub_codes = cls.get_code_locations(sub_ws)
                            if code not in sub_codes:
                                continue
                            s_info = sub_codes[code]
                            s_r0, s_c0 = s_info["row"], s_info["col"]
                            sub_concept_map = _extract_movement_table_rows(sub_ws, s_r0, label_col=s_c0)
                            sub_col_map = _build_comp_column_mapping(agg_ws, agg_r0, agg_c0, max_agg_col, sub_ws, s_r0, s_c0)
                            for ckey, r_a in agg_concept_map.items():
                                if ckey not in sub_concept_map:
                                    continue
                                r_s = sub_concept_map[ckey]
                                for c_a in range(agg_c0 + 1, max_agg_col + 1):
                                    cell_a = agg_ws.cell(r_a, c_a)
                                    if isinstance(cell_a, MergedCell) or cell_a.__class__.__name__ == "MergedCell":
                                        continue
                                    if isinstance(cell_a.value, str) and cell_a.value.startswith("="):
                                        continue
                                    if cls._is_header_cell_static(cell_a.value):
                                        continue
                                    c_s = sub_col_map.get(c_a)
                                    if c_s is not None and 1 <= c_s <= sub_ws.max_column:
                                        ok, num = cls._parse_numeric_cell(sub_ws.cell(r_s, c_s).value)
                                        if ok:
                                            prev = float(cell_a.value) if isinstance(cell_a.value, (int, float)) and not isinstance(cell_a.value, bool) else 0.0
                                            tot = prev + num
                                            cell_a.value = int(tot) if tot == int(tot) else round(tot, 4)
                        continue

                    agg_desc_cols, agg_data_cols = cls._resolve_ws_table_columns(agg_ws, code, agg_r0, agg_c0, agg_height)
                    agg_total_offs = [
                        off for off in range(1, agg_height + 1)
                        if cls._norm_concept_static(agg_ws.cell(agg_r0 + off, agg_c0).value).startswith("total")
                    ]

                    # --- CASO B: Tablas de detalle multicolumna (#N04.3, #N15.5, #N15.6, #N15.7, #N15.8, #N27.3) ---
                    if len(agg_desc_cols) > 1 and len(agg_total_offs) <= 1:
                        agg_detail_rows, agg_footer_rows = cls._classify_detail_and_footer_rows(
                            agg_ws, agg_r0, agg_c0, agg_r_end, agg_desc_cols, agg_data_cols
                        )
                        # Limpiar filas de muestra en agg_ws antes de poblar las operaciones reales de las filiales
                        for r_det in agg_detail_rows:
                            for c_cl in agg_desc_cols + agg_data_cols:
                                if c_cl <= agg_ws.max_column:
                                    tc = agg_ws.cell(r_det, c_cl)
                                    if not (isinstance(tc, MergedCell) or tc.__class__.__name__ == "MergedCell"):
                                        tc.value = None
                        for r_foot, _ in agg_footer_rows:
                            for c_fd in agg_data_cols:
                                if c_fd <= agg_ws.max_column:
                                    tc = agg_ws.cell(r_foot, c_fd)
                                    if not (isinstance(tc, MergedCell) or tc.__class__.__name__ == "MergedCell"):
                                        if not (isinstance(tc.value, str) and tc.value.startswith("=")):
                                            tc.value = 0

                        per_sub_rows = []
                        max_intra_count = {}
                        footer_sums = {}
                        for _, _, sub_ws in code_subs:
                            sub_codes = cls.get_code_locations(sub_ws)
                            if code not in sub_codes:
                                continue
                            s_info = sub_codes[code]
                            s_r0, s_c0 = s_info["row"], s_info["col"]
                            s_height = cls._get_ws_table_height(sub_ws, s_r0)
                            if s_height <= 0:
                                continue
                            s_desc_cols, s_data_cols = cls._resolve_ws_table_columns(sub_ws, code, s_r0, s_c0, s_height)
                            s_detail_rows, s_footer_rows = cls._classify_detail_and_footer_rows(
                                sub_ws, s_r0, s_c0, s_r0 + s_height, s_desc_cols, s_data_cols
                            )
                            sub_rows = []
                            sub_counts = {}
                            for r_s in s_detail_rows:
                                d_vals = [sub_ws.cell(r_s, cd).value if cd <= sub_ws.max_column else None for cd in s_desc_cols]
                                parsed_n = [
                                    cls._parse_numeric_cell(sub_ws.cell(r_s, cn).value if cn <= sub_ws.max_column else None)
                                    for cn in s_data_cols
                                ]
                                has_nonzero_num = any(ok and abs(num) > 1e-6 for ok, num in parsed_n)
                                if has_nonzero_num:
                                    n_vals = [
                                        (int(num) if num == int(num) else round(num, 4)) if ok else 0.0
                                        for ok, num in parsed_n
                                    ]
                                    desc_key = cls._build_desc_row_key(d_vals)
                                    sub_counts[desc_key] = sub_counts.get(desc_key, 0) + 1
                                    sub_rows.append((d_vals, n_vals, desc_key))
                            for dk, cnt in sub_counts.items():
                                if cnt > max_intra_count.get(dk, 0):
                                    max_intra_count[dk] = cnt
                            per_sub_rows.append(sub_rows)

                            used_f = set()
                            for r_sf, lbl_sf in s_footer_rows:
                                for r_af, lbl_af in agg_footer_rows:
                                    if r_af not in used_f and (lbl_sf == lbl_af or (lbl_sf[:15] and lbl_sf[:15] == lbl_af[:15])):
                                        used_f.add(r_af)
                                        for idx_c, cn in enumerate(s_data_cols):
                                            if idx_c < len(agg_data_cols) and cn <= sub_ws.max_column:
                                                ok, num = cls._parse_numeric_cell(sub_ws.cell(r_sf, cn).value)
                                                if ok:
                                                    footer_sums[(r_af, idx_c)] = footer_sums.get((r_af, idx_c), 0.0) + num
                                        break

                        # Unificar: sumar entre compañías solo si la llave es única dentro de cada compañía (max_intra_count == 1);
                        # si una misma compañía informó múltiples filas iguales (max_intra_count > 1), preservarlas en líneas separadas.
                        collected_details = []
                        merged_idx_by_key = {}
                        for sub_rows in per_sub_rows:
                            for d_vals, n_vals, desc_key in sub_rows:
                                is_nonempty_key = any(part != "" for part in desc_key)
                                if is_nonempty_key and max_intra_count.get(desc_key, 0) == 1:
                                    if desc_key in merged_idx_by_key:
                                        existing_d, existing_n = collected_details[merged_idx_by_key[desc_key]]
                                        for idx_d, dv in enumerate(d_vals):
                                            if idx_d < len(existing_d) and (existing_d[idx_d] is None or str(existing_d[idx_d]).strip() == ""):
                                                existing_d[idx_d] = dv
                                        for idx_n, nv in enumerate(n_vals):
                                            if idx_n < len(existing_n):
                                                tot_n = float(existing_n[idx_n]) + float(nv)
                                                existing_n[idx_n] = int(tot_n) if tot_n == int(tot_n) else round(tot_n, 4)
                                    else:
                                        fmt_n = [int(v) if v == int(v) else round(float(v), 4) for v in n_vals]
                                        merged_idx_by_key[desc_key] = len(collected_details)
                                        collected_details.append([list(d_vals), fmt_n])
                                else:
                                    fmt_n = [int(v) if v == int(v) else round(float(v), 4) for v in n_vals]
                                    collected_details.append([list(d_vals), fmt_n])

                        if len(collected_details) > len(agg_detail_rows) and agg_detail_rows:
                            extra_det = len(collected_details) - len(agg_detail_rows)
                            ins_at = agg_detail_rows[-1] + 1
                            agg_ws.insert_rows(ins_at, amount=extra_det)
                            for r_below in range(ins_at + extra_det, agg_ws.max_row + 1):
                                for c_below in range(1, agg_ws.max_column + 1):
                                    vb = agg_ws.cell(r_below, c_below).value
                                    if isinstance(vb, str) and vb.startswith("="):
                                        agg_ws.cell(r_below, c_below).value = shift_formula_rows(vb, extra_det, min_row=ins_at)
                            shifted_footer_sums = {}
                            for (r_af_old, idx_c), val_sum in footer_sums.items():
                                r_af_new = r_af_old + extra_det if r_af_old >= ins_at else r_af_old
                                shifted_footer_sums[(r_af_new, idx_c)] = val_sum
                            footer_sums = shifted_footer_sums
                            agg_detail_rows.extend(range(ins_at, ins_at + extra_det))
                            agg_footer_rows = [(rf + extra_det if rf >= ins_at else rf, lf) for (rf, lf) in agg_footer_rows]

                        for idx_det, (d_vals, n_vals) in enumerate(collected_details):
                            if idx_det >= len(agg_detail_rows):
                                break
                            r_dst = agg_detail_rows[idx_det]
                            for idx_d, cd in enumerate(agg_desc_cols):
                                if idx_d < len(d_vals):
                                    tc = agg_ws.cell(r_dst, cd)
                                    if not (isinstance(tc, MergedCell) or tc.__class__.__name__ == "MergedCell"):
                                        tc.value = d_vals[idx_d]
                            for idx_n, cn in enumerate(agg_data_cols):
                                if idx_n < len(n_vals):
                                    tc = agg_ws.cell(r_dst, cn)
                                    if not (isinstance(tc, MergedCell) or tc.__class__.__name__ == "MergedCell"):
                                        tc.value = n_vals[idx_n]

                        for (r_af, idx_c), val_sum in footer_sums.items():
                            c_af = agg_data_cols[idx_c]
                            tc = agg_ws.cell(r_af, c_af)
                            if not (isinstance(tc, MergedCell) or tc.__class__.__name__ == "MergedCell"):
                                if not (isinstance(tc.value, str) and tc.value.startswith("=")):
                                    tc.value = int(val_sum) if val_sum == int(val_sum) else round(val_sum, 4)

                        # Asegurar rango =SUM(...) coherente en el pie principal de detalle
                        if agg_detail_rows and agg_footer_rows:
                            first_det_r = agg_detail_rows[0]
                            last_det_r = agg_detail_rows[-1]
                            r_first_foot = agg_footer_rows[0][0]
                            for cn in agg_data_cols:
                                tc = agg_ws.cell(r_first_foot, cn)
                                if isinstance(tc.value, str) and tc.value.upper().startswith("=SUM("):
                                    col_let = get_column_letter(cn)
                                    tc.value = f"=SUM({col_let}{first_det_r}:{col_let}{last_det_r})"
                        continue

                    # --- CASO C: Tablas estándar de una columna de concepto (#N06.2, #N07.1..4, #N09.1, #N10.2..4, #N16.1, #N17.4, #N25.x, #N27.1..2, etc.) ---
                    agg_concept_rows = []
                    if agg_desc_cols:
                        c_a_desc = agg_desc_cols[0]
                        for r_a in range(agg_r0 + 1, min(agg_ws.max_row + 1, agg_r_end + 1)):
                            v_a = agg_ws.cell(r_a, c_a_desc).value
                            if isinstance(v_a, str) and v_a.strip() and not v_a.strip().startswith("="):
                                norm_a = cls._norm_concept_static(v_a)
                                if norm_a and norm_a not in ("concepto", "detalle", "m$"):
                                    agg_concept_rows.append((r_a, norm_a))

                    for r_a in range(agg_r0 + 1, min(agg_ws.max_row + 1, agg_r_end + 1)):
                        if any(ca <= agg_ws.max_column and cls._is_header_cell_static(agg_ws.cell(r_a, ca).value) for ca in agg_desc_cols + agg_data_cols):
                            continue
                        for ca in agg_data_cols:
                            cell_a = agg_ws.cell(r_a, ca)
                            if isinstance(cell_a, MergedCell) or cell_a.__class__.__name__ == "MergedCell":
                                continue
                            if isinstance(cell_a.value, str):
                                s_val = cell_a.value.strip()
                                if s_val.startswith("="):
                                    continue
                                ok_str, _ = cls._parse_numeric_cell(s_val)
                                if not ok_str and s_val != "":
                                    continue
                            if cls._is_header_cell_static(cell_a.value):
                                continue
                            cell_a.value = None

                    for _, _, sub_ws in code_subs:
                        sub_codes = cls.get_code_locations(sub_ws)
                        if code not in sub_codes:
                            continue
                        s_info = sub_codes[code]
                        s_r0, s_c0 = s_info["row"], s_info["col"]
                        s_height = cls._get_ws_table_height(sub_ws, s_r0)
                        if s_height <= 0:
                            continue
                        s_desc_cols, s_data_cols = cls._resolve_ws_table_columns(sub_ws, code, s_r0, s_c0, s_height)
                        used_agg_rows = set()

                        for off in range(1, s_height + 1):
                            r_s = s_r0 + off
                            if any(cs <= sub_ws.max_column and cls._is_header_cell_static(sub_ws.cell(r_s, cs).value) for cs in s_data_cols):
                                continue
                            v_s_desc = sub_ws.cell(r_s, s_desc_cols[0]).value if (s_desc_cols and s_desc_cols[0] <= sub_ws.max_column) else None
                            norm_s = cls._norm_concept_static(v_s_desc)
                            if not norm_s or norm_s in ("concepto", "detalle", "m$"):
                                continue

                            r_target = None
                            if agg_concept_rows:
                                candidates = [
                                    ra for (ra, norm_a) in agg_concept_rows
                                    if ra not in used_agg_rows and (
                                        norm_a == norm_s or
                                        (len(norm_s) >= 20 and len(norm_a) >= 20 and (norm_a.startswith(norm_s[:25]) or norm_s.startswith(norm_a[:25])))
                                    )
                                ]
                                if candidates:
                                    best_ra = min(candidates, key=lambda ra: abs((ra - agg_r0) - off))
                                    if norm_s in [na for (rc, na) in agg_concept_rows if rc == best_ra] or abs((best_ra - agg_r0) - off) <= max(8, agg_r_end - agg_r0):
                                        r_target = best_ra

                            if r_target is None and code in ("#N20.1", "#N20.2"):
                                cand_pos = agg_r0 + off
                                if cand_pos <= agg_r_end and cand_pos not in used_agg_rows:
                                    r_target = cand_pos
                                    if agg_desc_cols and v_s_desc is not None:
                                        tc_desc = agg_ws.cell(r_target, agg_desc_cols[0])
                                        if not (isinstance(tc_desc, MergedCell) or tc_desc.__class__.__name__ == "MergedCell"):
                                            tc_desc.value = v_s_desc

                            # Si una sociedad informa un concepto nuevo con saldo != 0 que no existía en la plantilla base,
                            # insertarlo dinámicamente antes de la fila de Total / Saldo final y expandir =SUM(...)
                            if r_target is None and code not in ("#N20.1", "#N20.2", "#N20.3") and not cls._is_footer_label_static(norm_s):
                                has_sub_nonzero = False
                                for cs_chk in s_data_cols:
                                    if cs_chk <= sub_ws.max_column:
                                        ok_chk, num_chk = cls._parse_numeric_cell(sub_ws.cell(r_s, cs_chk).value)
                                        if ok_chk and abs(num_chk) > 1e-6:
                                            has_sub_nonzero = True
                                            break
                                if has_sub_nonzero and agg_desc_cols:
                                    tot_r = None
                                    for r_chk in range(min(agg_ws.max_row, agg_r_end), agg_r0, -1):
                                        lbl_chk = cls._norm_concept_static(agg_ws.cell(r_chk, agg_desc_cols[0]).value)
                                        has_sum_chk = any(
                                            ca_chk <= agg_ws.max_column and isinstance(agg_ws.cell(r_chk, ca_chk).value, str) and "SUM" in str(agg_ws.cell(r_chk, ca_chk).value).upper()
                                            for ca_chk in agg_data_cols
                                        )
                                        if cls._is_footer_label_static(lbl_chk) or has_sum_chk:
                                            tot_r = r_chk
                                            break
                                    ins_r = tot_r if tot_r is not None else (agg_r_end + 1)
                                    old_prev_r = ins_r - 1
                                    agg_ws.insert_rows(ins_r, amount=1)
                                    for r_below in range(ins_r + 1, agg_ws.max_row + 1):
                                        for c_below in range(1, agg_ws.max_column + 1):
                                            vb = agg_ws.cell(r_below, c_below).value
                                            if isinstance(vb, str) and vb.startswith("="):
                                                agg_ws.cell(r_below, c_below).value = shift_formula_rows(vb, 1, min_row=ins_r)
                                    if tot_r is not None and old_prev_r > agg_r0:
                                        tot_r_new = tot_r + 1
                                        for c_col in range(1, agg_ws.max_column + 1):
                                            vf = agg_ws.cell(tot_r_new, c_col).value
                                            if isinstance(vf, str) and vf.upper().startswith("=SUM("):
                                                agg_ws.cell(tot_r_new, c_col).value = re.sub(
                                                    rf'(:[A-Za-z]+){old_prev_r}\b',
                                                    rf'\g<1>{ins_r}',
                                                    vf
                                                )
                                    # Copiar fórmulas horizontales de fila si la fila anterior las tenía (ej. Activo Fijo Neto =SUM(B23:C23))
                                    if old_prev_r > agg_r0:
                                        for ca_f in agg_data_cols:
                                            if ca_f <= agg_ws.max_column:
                                                v_prev_f = agg_ws.cell(old_prev_r, ca_f).value
                                                if isinstance(v_prev_f, str) and v_prev_f.startswith("="):
                                                    agg_ws.cell(ins_r, ca_f).value = shift_formula_rows(v_prev_f, 1, min_row=old_prev_r)
                                    agg_ws.cell(ins_r, agg_desc_cols[0]).value = v_s_desc
                                    agg_r_end += 1
                                    agg_concept_rows = [(ra + 1 if ra >= ins_r else ra, na) for (ra, na) in agg_concept_rows]
                                    used_agg_rows = {ra + 1 if ra >= ins_r else ra for ra in used_agg_rows}
                                    agg_concept_rows.append((ins_r, norm_s))
                                    r_target = ins_r

                            if r_target is None or r_target > agg_ws.max_row:
                                continue
                            used_agg_rows.add(r_target)

                            is_tot_row = (
                                cls._is_footer_label_static(norm_s) or
                                (agg_desc_cols and cls._is_footer_label_static(cls._norm_concept_static(agg_ws.cell(r_target, agg_desc_cols[0]).value)))
                            )
                            if is_tot_row and code not in ("#N20.1", "#N20.2"):
                                continue

                            for idx_c, cs in enumerate(s_data_cols):
                                if idx_c >= len(agg_data_cols):
                                    break
                                ca = agg_data_cols[idx_c]
                                cell_a = agg_ws.cell(r_target, ca)
                                if isinstance(cell_a, MergedCell) or cell_a.__class__.__name__ == "MergedCell":
                                    continue
                                if isinstance(cell_a.value, str) and cell_a.value.startswith("="):
                                    continue
                                if cls._is_header_cell_static(cell_a.value):
                                    continue
                                ok, num = cls._parse_numeric_cell(sub_ws.cell(r_s, cs).value if cs <= sub_ws.max_column else None)
                                if ok:
                                    prev = float(cell_a.value) if isinstance(cell_a.value, (int, float)) and not isinstance(cell_a.value, bool) else 0.0
                                    tot = prev + num
                                    cell_a.value = int(tot) if tot == int(tot) else round(tot, 4)
        finally:
            for _, _, swb in sub_wbs:
                try:
                    swb.close()
                except Exception:
                    pass

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
        cls._AGG_ANEXO_CACHE.clear()
        try:
            from src.core.snapshot_vault_manager import _HYBRID_COMP_CACHE
            _HYBRID_COMP_CACHE.clear()
        except Exception:
            pass
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
            for c in range(2, min(ext_ws.max_column + 1, 18)):
                val = ext_ws.cell(r, c).value
                if val is not None and isinstance(val, (int, float)) and not isinstance(val, bool) and abs(val) > 0.001:
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
    def _inject_table_data(cls, ext_ws, target_ws, is_comparative: bool = False, target_lang: str = "es", structure_only: bool = False, skip_codes: set = None, clean_unused_when_comp_only: bool = False):
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
            if skip_codes and code in skip_codes:
                continue
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
                    actual_ext_m_cols = []
                    for r_m in range(ext_r0 + 1, min(ext_r0 + 6, ext_r0 + table_height + 1, ext_ws.max_row + 1)):
                        for c_m in range(ext_info["col"] + 1, min(ext_info["col"] + 15, ext_ws.max_column + 1)):
                            vm = ext_ws.cell(r_m, c_m).value
                            if isinstance(vm, str) and vm.strip().lower() in ("m$", "ms"):
                                if c_m not in actual_ext_m_cols:
                                    actual_ext_m_cols.append(c_m)
                    actual_ext_m_cols.sort()
                    if base_ext_data_cols and len(actual_ext_m_cols) == len(base_ext_data_cols):
                        ext_data_cols = actual_ext_m_cols
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

                # Si ext_ws tiene filas vacías espaciadoras justo debajo del encabezado antes del primer dato real (ej. R90 en #N15.8), omitirlas
                active_ext_details = [
                    r_e for r_e in ext_detail_rows
                    if any(
                        (c <= ext_ws.max_column and ext_ws.cell(r_e, c).value is not None and str(ext_ws.cell(r_e, c).value).strip() not in ("", "-", "0", "0.0"))
                        for c in ext_desc_cols + ext_data_cols
                    )
                ]
                if active_ext_details:
                    first_act_idx = ext_detail_rows.index(active_ext_details[0])
                    if first_act_idx > 0:
                        ext_detail_rows = ext_detail_rows[first_act_idx:]
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

                # Limpiar filas de detalle sobrantes en target_ws que en ext_ws están vacías
                if (not is_comparative or clean_unused_when_comp_only) and not structure_only:
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
                                if norm_e in [nt for (rc, nt) in target_concept_rows if rc == best_r_t] or abs((best_r_t - target_r0) - off) <= max(4, target_r_end - target_r0):
                                    r_target = best_r_t
                                    used_target_rows.add(best_r_t)
                                    matched_by_concept = True
                    if not matched_by_concept and len(ext_desc_cols) == 1:
                        ve_current = _norm_concept(ext_ws.cell(r_ext, ext_desc_cols[0]).value)
                        has_ext_nonzero = False
                        for ce_chk in ext_data_cols:
                            if ce_chk <= ext_ws.max_column:
                                ok_e, num_e = cls._parse_numeric_cell(ext_ws.cell(r_ext, ce_chk).value)
                                if ok_e and abs(num_e) > 1e-6:
                                    has_ext_nonzero = True
                                    break
                        # Si es un concepto nuevo con saldo != 0 ausente en target_ws, insertarlo antes del Total y expandir =SUM(...)
                        if (
                            ve_current
                            and ve_current not in ("concepto", "detalle", "m$")
                            and not _is_header_value(ext_ws.cell(r_ext, ext_desc_cols[0]).value)
                            and not _is_footer_or_summary_label(ve_current)
                            and has_ext_nonzero
                            and code not in ("#N20.1", "#N20.2", "#N20.3")
                            and not structure_only
                            and target_desc_cols
                        ):
                            tot_t = None
                            for rt_chk in range(min(target_ws.max_row, target_r_end), target_r0, -1):
                                lbl_t_chk = _norm_concept(target_ws.cell(rt_chk, target_desc_cols[0]).value)
                                has_sum_t_chk = any(
                                    ct_chk <= target_ws.max_column and isinstance(target_ws.cell(rt_chk, ct_chk).value, str) and "SUM" in str(target_ws.cell(rt_chk, ct_chk).value).upper()
                                    for ct_chk in target_data_cols
                                )
                                if _is_footer_or_summary_label(lbl_t_chk) or has_sum_t_chk:
                                    tot_t = rt_chk
                                    break
                            ins_t = tot_t if tot_t is not None else (target_r_end + 1)
                            old_prev_t = ins_t - 1
                            target_ws.insert_rows(ins_t, amount=1)
                            for r_below in range(ins_t + 1, target_ws.max_row + 1):
                                for c_below in range(1, target_ws.max_column + 1):
                                    vb = target_ws.cell(r_below, c_below).value
                                    if isinstance(vb, str) and vb.startswith("="):
                                        target_ws.cell(r_below, c_below).value = shift_formula_rows(vb, 1, min_row=ins_t)
                            if tot_t is not None and old_prev_t > target_r0:
                                tot_t_new = tot_t + 1
                                for c_col in range(1, target_ws.max_column + 1):
                                    vf = target_ws.cell(tot_t_new, c_col).value
                                    if isinstance(vf, str) and vf.upper().startswith("=SUM("):
                                        target_ws.cell(tot_t_new, c_col).value = re.sub(
                                            rf'(:[A-Za-z]+){old_prev_t}\b',
                                            rf'\g<1>{ins_t}',
                                            vf
                                        )
                            if old_prev_t > target_r0:
                                for c_f in range(max(target_desc_cols) + 1, target_ws.max_column + 1):
                                    v_pf = target_ws.cell(old_prev_t, c_f).value
                                    if isinstance(v_pf, str) and v_pf.startswith("="):
                                        target_ws.cell(ins_t, c_f).value = shift_formula_rows(v_pf, 1, min_row=old_prev_t)
                            v_raw_desc = ext_ws.cell(r_ext, ext_desc_cols[0]).value
                            target_ws.cell(ins_t, target_desc_cols[0]).value = (
                                translate_ifrs_term(str(v_raw_desc).strip(), target_lang="en")
                                if (is_en and isinstance(v_raw_desc, str) and v_raw_desc.strip())
                                else v_raw_desc
                            )
                            target_r_end += 1
                            target_concept_rows = [(rt + 1 if rt >= ins_t else rt, nt) for (rt, nt) in target_concept_rows]
                            used_target_rows = {rt + 1 if rt >= ins_t else rt for rt in used_target_rows}
                            target_concept_rows.append((ins_t, ve_current))
                            used_target_rows.add(ins_t)
                            r_target = ins_t
                        else:
                            # No permitir que una fila de control no emparejada (ej. 'reclasificación' en #N13.1) pise una fila ya usada o de Total
                            if r_target in used_target_rows:
                                continue
                            vt_existing = _norm_concept(target_ws.cell(r_target, target_desc_cols[0]).value) if r_target <= target_ws.max_row else ""
                            if _is_footer_or_summary_label(vt_existing) and not _is_footer_or_summary_label(ve_current):
                                continue
                            if vt_existing and ve_current and vt_existing != ve_current and code not in ("#N20.1", "#N20.2"):
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

            # 3. Asegurar unidad (M$), fórmulas =SUM(...) y bordes de sumatoria en columnas comparativas dinámicas o filas Saldo final / Total sin fórmula
            if not structure_only and target_data_cols:
                from copy import copy as _copy_st
                from openpyxl.utils import get_column_letter
                from openpyxl.styles import Border, Side, Font

                max_desc_c = max(target_desc_cols) if target_desc_cols else target_info["col"]
                n_td = len(target_data_cols)
                r_lim = min(target_ws.max_row, target_r_end)

                for c_target in target_data_cols:
                    tgt_let = get_column_letter(c_target)
                    c_ref = (c_target - n_td) if is_comparative else None
                    if c_ref is not None and c_ref > max_desc_c:
                        ref_let = get_column_letter(c_ref)
                        for r_chk in range(target_r0 + 1, r_lim + 1):
                            ref_c = target_ws.cell(r_chk, c_ref)
                            tgt_c = target_ws.cell(r_chk, c_target)
                            if isinstance(tgt_c, MergedCell) or tgt_c.__class__.__name__ == 'MergedCell':
                                continue
                            rv = ref_c.value
                            if isinstance(rv, str) and rv.strip().lower() in ("m$", "ms", "thch$") and tgt_c.value is None:
                                tgt_c.value = rv
                                if ref_c.has_style:
                                    tgt_c.font = _copy_st(ref_c.font)
                                    tgt_c.border = _copy_st(ref_c.border)
                                    tgt_c.fill = _copy_st(ref_c.fill)
                                    tgt_c.alignment = _copy_st(ref_c.alignment)
                            elif isinstance(rv, str) and rv.strip().upper().startswith("=SUM("):
                                cur_tv = tgt_c.value
                                if not (isinstance(cur_tv, str) and cur_tv.strip().upper().startswith("=SUM(")):
                                    tgt_c.value = re.sub(rf'\b{ref_let}(\d+)\b', rf'{tgt_let}\1', rv, flags=re.IGNORECASE)
                                if ref_c.has_style:
                                    tgt_c.font = _copy_st(ref_c.font)
                                    tgt_c.border = _copy_st(ref_c.border)
                                    tgt_c.fill = _copy_st(ref_c.fill)
                                    tgt_c.alignment = _copy_st(ref_c.alignment)
                                    tgt_c.number_format = ref_c.number_format
                            elif isinstance(tgt_c.value, (int, float)) and ref_c.has_style:
                                tgt_c.number_format = ref_c.number_format
                                tgt_c.font = _copy_st(ref_c.font)

                    # Respaldo general: si la fila final 'Total' / 'Saldo final' quedó vacía en c_target habiendo datos numéricos arriba, generar =SUM(...)
                    if target_desc_cols:
                        c_d0 = target_desc_cols[0]
                        first_num_r = None
                        last_num_r = None
                        for r_chk in range(target_r0 + 1, r_lim + 1):
                            lbl_c = _norm_concept(target_ws.cell(r_chk, c_d0).value)
                            tc_chk = target_ws.cell(r_chk, c_target)
                            if isinstance(tc_chk, MergedCell) or tc_chk.__class__.__name__ == 'MergedCell':
                                continue
                            if _is_footer_or_summary_label(lbl_c):
                                if tc_chk.value is None and first_num_r is not None and last_num_r is not None:
                                    tc_chk.value = f"=SUM({tgt_let}{first_num_r}:{tgt_let}{last_num_r})"
                                    tc_chk.font = Font(bold=True)
                                    tc_chk.border = Border(
                                        top=Side(style="thin", color="000000"),
                                        bottom=Side(style="thin", color="000000"),
                                    )
                                    tc_chk.number_format = r'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
                                first_num_r = None
                                last_num_r = None
                            elif isinstance(tc_chk.value, (int, float)) and not _is_header_value(tc_chk.value):
                                if first_num_r is None:
                                    first_num_r = r_chk
                                last_num_r = r_chk

    @classmethod
    def _reconcile_multicol_act_and_cmp_in_ws(cls, act_ws, cmp_ws, target_ws, target_lang: str = "es") -> set:
        """
        Alinea conjuntamente Período Actual y Período Comparativo por llave descriptiva completa
        (_build_desc_row_key + índice de ocurrencia intra-período) en tablas multicolumna con
        columnas comparativas lado a lado (ej. #N04.3, #N14.3, #N15.5, #N15.6, #N15.8).
        """
        from src.core.ifrs_glossary import translate_ifrs_term
        from src.reporting.note_generator import shift_formula_rows

        handled = set()
        if act_ws is None or cmp_ws is None or target_ws is None:
            return handled

        is_en = str(target_lang).lower() == "en"
        act_codes = cls.get_code_locations(act_ws)
        cmp_codes = cls.get_code_locations(cmp_ws)
        init_tgt_codes = cls.get_code_locations(target_ws)
        ordered_codes = [c for c, _ in sorted(init_tgt_codes.items(), key=lambda kv: kv[1]["row"])]

        for code in ordered_codes:
            if code not in act_codes or code not in cmp_codes:
                continue
            cfg = cls.EXTERNAL_TABLE_REGISTRY.get(code)
            if not cfg:
                continue
            base_desc_cols = cfg.get("desc_cols", [])
            base_act_cols = cfg.get("base_act_data_cols", [])
            base_cmp_cols = cfg.get("base_cmp_data_cols", [])
            if len(base_desc_cols) <= 1 or not base_act_cols or not base_cmp_cols or base_act_cols == base_cmp_cols:
                continue

            tgt_codes = cls.get_code_locations(target_ws)
            if code not in tgt_codes:
                continue
            t_info = tgt_codes[code]
            if "[COMPARATIVO]" in str(t_info.get("val", "")).upper():
                continue
            t_r0, t_c0 = t_info["row"], t_info["col"]

            a_info = act_codes[code]
            c_info = cmp_codes[code]
            a_r0, a_c0 = a_info["row"], a_info["col"]
            c_r0, c_c0 = c_info["row"], c_info["col"]
            a_h = cls._get_ws_table_height(act_ws, a_r0)
            c_h = cls._get_ws_table_height(cmp_ws, c_r0)
            if a_h <= 0 and c_h <= 0:
                continue

            t_r_end = min(target_ws.max_row, t_r0 + max(a_h, c_h, 10) + 6)
            for r_scan in range(t_r0 + 1, t_r_end + 1):
                hit_b = False
                for c_chk in range(1, min(target_ws.max_column + 1, 15)):
                    vt_chk = target_ws.cell(r_scan, c_chk).value
                    if vt_chk and isinstance(vt_chk, str):
                        vt_s = vt_chk.strip()
                        if re.search(r'#N\d+\.\d+', vt_s):
                            hit_b = True
                            break
                        if c_chk == t_c0 and re.match(r'^(?:\d+|[a-z])[\)\.\-]\s+', vt_s.lower()):
                            hit_b = True
                            break
                if hit_b:
                    t_r_end = r_scan - 1
                    break

            tgt_total_rows = [
                r_t for r_t in range(t_r0 + 1, t_r_end + 1)
                if cls._norm_concept_static(target_ws.cell(r_t, t_c0).value).startswith("total")
            ]
            if len(tgt_total_rows) > 1:
                continue

            base_ref_col = base_desc_cols[0]
            t_delta = t_c0 - base_ref_col
            target_desc_cols = [c + t_delta for c in base_desc_cols if c + t_delta >= 1]
            target_act_cols = [c + t_delta for c in base_act_cols if c + t_delta >= 1]
            target_cmp_cols = [c + t_delta for c in base_cmp_cols if c + t_delta >= 1]

            a_desc_cols, a_data_cols = cls._resolve_ws_table_columns(act_ws, code, a_r0, a_c0, max(a_h, 1))
            c_desc_cols, c_data_cols = cls._resolve_ws_table_columns(cmp_ws, code, c_r0, c_c0, max(c_h, 1))

            # Clasificar filas de target_ws respetando min_sum_start_row
            min_sum_start = None
            for r_t in range(t_r0 + 1, t_r_end + 1):
                for c_t in target_act_cols + target_cmp_cols:
                    if c_t <= target_ws.max_column:
                        vt_f = target_ws.cell(r_t, c_t).value
                        if isinstance(vt_f, str) and vt_f.upper().startswith("=SUM("):
                            m_rng = re.search(r'[A-Za-z]+(\d+):[A-Za-z]+(\d+)', vt_f)
                            if m_rng:
                                r_sc = int(m_rng.group(1))
                                if t_r0 < r_sc <= r_t and (min_sum_start is None or r_sc < min_sum_start):
                                    min_sum_start = r_sc

            raw_tgt_det, tgt_footer_rows = cls._classify_detail_and_footer_rows(
                target_ws, t_r0, t_c0, t_r_end, target_desc_cols, target_act_cols + target_cmp_cols
            )
            tgt_detail_rows = []
            for r_t in raw_tgt_det:
                if min_sum_start is not None and r_t < min_sum_start:
                    has_any_t = any(
                        target_ws.cell(r_t, c).value is not None
                        for c in target_desc_cols + target_act_cols + target_cmp_cols
                        if c <= target_ws.max_column
                    )
                    if not has_any_t:
                        continue
                tgt_detail_rows.append(r_t)
            if not tgt_detail_rows:
                continue

            def _extract_period_items(ws_p, r0_p, c0_p, h_p, desc_c_p, data_c_p):
                if h_p <= 0:
                    return [], []
                det_r_p, foot_r_p = cls._classify_detail_and_footer_rows(
                    ws_p, r0_p, c0_p, r0_p + h_p, desc_c_p, data_c_p
                )
                items = []
                occ_tracker = {}
                for rp in det_r_p:
                    d_vals = [ws_p.cell(rp, cd).value if cd <= ws_p.max_column else None for cd in desc_c_p]
                    parsed_n = [
                        cls._parse_numeric_cell(ws_p.cell(rp, cn).value if cn <= ws_p.max_column else None)
                        for cn in data_c_p
                    ]
                    if any(ok and abs(num) > 1e-6 for ok, num in parsed_n):
                        n_vals = [
                            (int(num) if num == int(num) else round(num, 4)) if ok else 0
                            for ok, num in parsed_n
                        ]
                        dk = cls._build_desc_row_key(d_vals)
                        occ = occ_tracker.get(dk, 0)
                        occ_tracker[dk] = occ + 1
                        items.append((d_vals, n_vals, (dk, occ)))
                return items, foot_r_p

            act_items, act_footers = _extract_period_items(act_ws, a_r0, a_c0, a_h, a_desc_cols, a_data_cols)
            cmp_items, cmp_footers = _extract_period_items(cmp_ws, c_r0, c_c0, c_h, c_desc_cols, c_data_cols)

            cmp_lookup = {k_occ: (idx_c, d_c, n_c) for idx_c, (d_c, n_c, k_occ) in enumerate(cmp_items)}
            matched_cmp_indices = set()
            unified_rows = []

            for d_a, n_a, k_occ in act_items:
                if k_occ in cmp_lookup:
                    idx_c, d_c, n_c = cmp_lookup[k_occ]
                    matched_cmp_indices.add(idx_c)
                    merged_d = [
                        d_a[i] if (i < len(d_a) and d_a[i] is not None and str(d_a[i]).strip() != "") else (d_c[i] if i < len(d_c) else None)
                        for i in range(max(len(d_a), len(d_c)))
                    ]
                    unified_rows.append((merged_d, n_a, n_c))
                else:
                    unified_rows.append((d_a, n_a, None))

            for idx_c, (d_c, n_c, _) in enumerate(cmp_items):
                if idx_c not in matched_cmp_indices:
                    unified_rows.append((d_c, None, n_c))

            if len(unified_rows) > len(tgt_detail_rows):
                extra_det = len(unified_rows) - len(tgt_detail_rows)
                ins_at = tgt_detail_rows[-1] + 1
                old_last_det = tgt_detail_rows[-1]
                new_last_det = old_last_det + extra_det
                target_ws.insert_rows(ins_at, amount=extra_det)
                for r_below in range(ins_at + extra_det, target_ws.max_row + 1):
                    for c_below in range(1, target_ws.max_column + 1):
                        vb = target_ws.cell(r_below, c_below).value
                        if isinstance(vb, str) and vb.startswith("="):
                            target_ws.cell(r_below, c_below).value = shift_formula_rows(vb, extra_det, min_row=ins_at)
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
                tgt_footer_rows = [(rf + extra_det, lf) for (rf, lf) in tgt_footer_rows]

            # Limpiar todas las celdas de detalle en target_ws antes de escribir las filas reconciliadas
            for r_det in tgt_detail_rows:
                for c_cl in target_desc_cols + target_act_cols + target_cmp_cols:
                    if c_cl <= target_ws.max_column:
                        tc = target_ws.cell(r_det, c_cl)
                        if not (isinstance(tc, MergedCell) or tc.__class__.__name__ == "MergedCell"):
                            if not (isinstance(tc.value, str) and tc.value.startswith("=")):
                                tc.value = None

            for idx_u, (d_vals, n_act, n_cmp) in enumerate(unified_rows):
                if idx_u >= len(tgt_detail_rows):
                    break
                r_dst = tgt_detail_rows[idx_u]
                for idx_d, cd_t in enumerate(target_desc_cols):
                    if idx_d < len(d_vals) and cd_t <= target_ws.max_column:
                        val_d = d_vals[idx_d]
                        if val_d is not None and not cls._is_header_cell_static(val_d, allow_dates=True):
                            tc = target_ws.cell(r_dst, cd_t)
                            if not (isinstance(tc, MergedCell) or tc.__class__.__name__ == "MergedCell"):
                                if is_en and isinstance(val_d, str) and val_d.strip():
                                    tc.value = translate_ifrs_term(val_d.strip(), target_lang="en")
                                else:
                                    tc.value = val_d
                if n_act is not None:
                    for idx_n, ca_t in enumerate(target_act_cols):
                        if idx_n < len(n_act) and ca_t <= target_ws.max_column:
                            tc = target_ws.cell(r_dst, ca_t)
                            if not (isinstance(tc, MergedCell) or tc.__class__.__name__ == "MergedCell"):
                                if not (isinstance(tc.value, str) and tc.value.startswith("=")):
                                    tc.value = float(n_act[idx_n])
                if n_cmp is not None:
                    for idx_n, cc_t in enumerate(target_cmp_cols):
                        if idx_n < len(n_cmp) and cc_t <= target_ws.max_column:
                            tc = target_ws.cell(r_dst, cc_t)
                            if not (isinstance(tc, MergedCell) or tc.__class__.__name__ == "MergedCell"):
                                if not (isinstance(tc.value, str) and tc.value.startswith("=")):
                                    tc.value = float(n_cmp[idx_n])

            # Inyectar valores de pie no-fórmula (ej. intereses diferidos / devengados) para Actual y Comparativo
            for src_ws, src_footers, src_data_cols, dst_data_cols in (
                (act_ws, act_footers, a_data_cols, target_act_cols),
                (cmp_ws, cmp_footers, c_data_cols, target_cmp_cols),
            ):
                used_tf = set()
                for r_sf, lbl_sf in src_footers:
                    for r_tf, lbl_tf in tgt_footer_rows:
                        if r_tf not in used_tf and (lbl_sf == lbl_tf or (lbl_sf[:15] and lbl_sf[:15] == lbl_tf[:15])):
                            used_tf.add(r_tf)
                            for idx_n, c_src in enumerate(src_data_cols):
                                if idx_n < len(dst_data_cols) and c_src <= src_ws.max_column and dst_data_cols[idx_n] <= target_ws.max_column:
                                    tc = target_ws.cell(r_tf, dst_data_cols[idx_n])
                                    if isinstance(tc, MergedCell) or tc.__class__.__name__ == "MergedCell":
                                        continue
                                    if isinstance(tc.value, str) and tc.value.startswith("="):
                                        continue
                                    ok_f, num_f = cls._parse_numeric_cell(src_ws.cell(r_sf, c_src).value)
                                    if ok_f:
                                        tc.value = float(num_f)
                            break

            handled.add(code)
        return handled

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

        is_auto_agg_act = cls.is_consolidated_group(empresa) and not cls.has_explicit_anexo(empresa, periodo_actual)
        is_auto_agg_cmp = bool(periodo_comp) and cls.is_consolidated_group(empresa) and not cls.has_explicit_anexo(empresa, periodo_comp)

        for sheet_name in out_wb.sheetnames:
            target_ws = out_wb[sheet_name]
            act_ws = ext_act_wb[sheet_name] if (ext_act_wb and sheet_name in ext_act_wb.sheetnames) else None
            cmp_ws = ext_cmp_wb[sheet_name] if (ext_cmp_wb and sheet_name in ext_cmp_wb.sheetnames) else None
            struct_ws = master_struct_wb[sheet_name] if (master_struct_wb and sheet_name in master_struct_wb.sheetnames) else None

            # Exclusión explícita para grupos consolidados auto-agregados (Impuestos Diferidos y Relacionadas)
            if cls._is_excluded_agg_sheet(sheet_name):
                if is_auto_agg_act:
                    act_ws = None
                if is_auto_agg_cmp:
                    cmp_ws = None

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

            # Preservar líneas contables PPA de consolidación ('Asignacion de compra' y 'Amortizacion PPA')
            # en #N09.1 cuando la nota Activo Fijo de un grupo consolidado se auto-agrega desde anexos de filiales.
            saved_af_ppa_cells = {}
            if sheet_name == "Activo Fijo" and str(empresa or "").strip().startswith("[GRUPO]"):
                from src.core.consolidacion_account_mapper import normalize_text
                for r_af in range(1, min(35, target_ws.max_row + 1)):
                    lbl_af = normalize_text(target_ws.cell(row=r_af, column=1).value)
                    if lbl_af in ("asignacion de compra", "amortizacion ppa"):
                        for c_af in (2, 3, 5, 6):
                            v_af = target_ws.cell(row=r_af, column=c_af).value
                            if isinstance(v_af, (int, float)) and not isinstance(v_af, bool) and abs(float(v_af)) > 0.001:
                                saved_af_ppa_cells[(r_af, c_af)] = v_af

            handled_multicol_codes = set()
            if act_ws and cmp_ws:
                handled_multicol_codes = cls._reconcile_multicol_act_and_cmp_in_ws(
                    act_ws, cmp_ws, target_ws, target_lang=target_lang
                )

            # 1. Inyectar datos del período comparativo (si existe)
            if cmp_ws:
                cls._inject_table_data(
                    cmp_ws,
                    target_ws,
                    is_comparative=True,
                    target_lang=target_lang,
                    skip_codes=handled_multicol_codes,
                    clean_unused_when_comp_only=(act_ws is None),
                )

            # 2. Inyectar datos del período actual (si existe) o sincronizar estructura desde la plantilla máster
            if act_ws:
                cls._inject_table_data(
                    act_ws,
                    target_ws,
                    is_comparative=False,
                    target_lang=target_lang,
                    skip_codes=handled_multicol_codes,
                )
            elif struct_ws and not cmp_ws:
                cls._inject_table_data(struct_ws, target_ws, is_comparative=False, target_lang=target_lang, structure_only=True)

            if saved_af_ppa_cells:
                num_fmt_af = r'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
                for (r_af, c_af), v_af in saved_af_ppa_cells.items():
                    if (c_af in (2, 3) and act_ws is not None and is_auto_agg_act) or (c_af in (5, 6) and cmp_ws is not None and is_auto_agg_cmp):
                        cur_cell = target_ws.cell(row=r_af, column=c_af)
                        if cur_cell.value in (None, 0, 0.0, ""):
                            cur_cell.value = v_af
                            cur_cell.number_format = num_fmt_af

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
