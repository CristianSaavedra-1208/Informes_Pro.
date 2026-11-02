import os
import shutil
import hashlib
import json
import pickle
import datetime
from typing import Dict, List, Tuple, Any, Optional
import pandas as pd
from io import BytesIO

from src.models.database import SessionLocal
from src.models.historical_data import ClosingSnapshotRecord, HistoricalDetailRecord
from src.models.consolidacion import ConsolidationGroup, ConsolidationPeriodLock
from src.models.trial_balance_db import TrialBalanceDB
from src.models.pl_cubo_db import PlCuboDB
from src.reporting.notes import get_full_note_registry
from src.core.excel_utils import format_periodo

# Caché en memoria para comprobación ultra rápida (< 0.1 ms)
_SNAPSHOT_STATUS_CACHE = {}

class SnapshotVaultManager:
    """
    Motor de Bóveda y Snapshots de Cierre Histórico.
    Garantiza inmutabilidad, auditoría y lectura ultra-rápida (< 1 seg)
    de Estados Financieros y Notas explicativas.
    """

    @staticmethod
    def get_snapshot_dir(empresa: str, periodo: str) -> str:
        """Ruta física donde se guardan los archivos congelados del snapshot."""
        clean_empresa = empresa.replace("[GRUPO] ", "").strip() if empresa.startswith("[GRUPO]") else empresa.strip()
        # Conservar el prefijo de carpeta si existe en disco
        target_dir = os.path.join("data", "empresas", empresa, "snapshots", periodo)
        if not os.path.exists(os.path.join("data", "empresas", empresa)):
            alt_dir = os.path.join("data", "empresas", clean_empresa, "snapshots", periodo)
            return alt_dir
        return target_dir

    @staticmethod
    def has_active_snapshot(empresa: str, periodo: str) -> bool:
        """Verifica si existe un snapshot oficial activo para la entidad y período."""
        if not empresa or not periodo:
            return False
            
        cache_key = (str(empresa).strip(), str(periodo).strip())
        if cache_key in _SNAPSHOT_STATUS_CACHE:
            return _SNAPSHOT_STATUS_CACHE[cache_key]

        snap_dir = SnapshotVaultManager.get_snapshot_dir(empresa, periodo)
        meta_file = os.path.join(snap_dir, "metadata.json")
        
        has_meta = os.path.exists(meta_file)
        if not has_meta:
            _SNAPSHOT_STATUS_CACHE[cache_key] = False
            return False

        # Verificar en base de datos
        db = SessionLocal()
        try:
            rec = db.query(ClosingSnapshotRecord).filter_by(
                empresa=str(empresa).strip(),
                periodo=str(periodo).strip(),
                is_active=True
            ).first()
            active = (rec is not None)
            _SNAPSHOT_STATUS_CACHE[cache_key] = active
            return active
        except Exception:
            _SNAPSHOT_STATUS_CACHE[cache_key] = has_meta
            return has_meta
        finally:
            db.close()

    @staticmethod
    def _compute_checksum(folder_path: str) -> str:
        """Calcula el hash SHA-256 acumulado de todos los archivos del snapshot."""
        sha = hashlib.sha256()
        for root, _, files in os.walk(folder_path):
            for f in sorted(files):
                if f == "metadata.json":
                    continue
                file_path = os.path.join(root, f)
                try:
                    with open(file_path, "rb") as fp:
                        while chunk := fp.read(65536):
                            sha.update(chunk)
                except Exception:
                    pass
        return sha.hexdigest()

    @staticmethod
    def _scale_dataframe_cells(df: pd.DataFrame, scale_factor: float) -> pd.DataFrame:
        """Escala las columnas numéricas de un DataFrame al factor deseado."""
        if df is None or df.empty or scale_factor == 1.0:
            return df.copy() if df is not None else None
            
        scaled = df.copy()
        for col in scaled.columns:
            scaled[col] = scaled[col].apply(
                lambda x: (x / scale_factor) if isinstance(x, (int, float)) and not isinstance(x, bool) and not pd.isna(x) else x
            )
        return scaled

    @staticmethod
    def create_snapshot(empresa: str, periodo: str, user: str = "Administrador", progress_callback=None) -> Tuple[bool, str]:
        """
        Calcula y congela en la Bóveda todos los Estados Financieros y Notas para una entidad y período.
        Valores almacenados en unidades base ($) exactas.
        """
        empresa_clean = str(empresa).strip()
        periodo_clean = str(periodo).strip()
        is_consolidated = empresa_clean.startswith("[GRUPO]")
        
        # Calcular periodo comparativo por regla IFRS
        parts = periodo_clean.split('-')
        prev_year = str(int(parts[0]) - 1) if len(parts) >= 2 else "2024"
        comp_period_bal = f"{prev_year}-12"
        comp_period_pl = f"{prev_year}-{parts[1]}" if len(parts) >= 2 else f"{prev_year}-12"

        snap_dir = SnapshotVaultManager.get_snapshot_dir(empresa_clean, periodo_clean)
        stmt_dir = os.path.join(snap_dir, "statements")
        notes_dir = os.path.join(snap_dir, "notes")
        os.makedirs(stmt_dir, exist_ok=True)
        os.makedirs(notes_dir, exist_ok=True)

        from src.core.word_template_engine import load_mappings_for_entity, WordTemplateEngine
        map_balance_df, map_pl_df = load_mappings_for_entity(empresa_clean)
        wte = WordTemplateEngine(None)
        
        empresa_path = os.path.join("data", "empresas", empresa_clean)
        if not os.path.exists(empresa_path):
            empresa_path = os.path.join("data", "empresas", empresa_clean.replace("[GRUPO] ", "").strip())

        statements_saved = {}
        notes_saved = []

        total_steps = 5 + len(get_full_note_registry())
        current_step = 0

        # =========================================================================
        # 1. GENERACIÓN DE ESTADOS FINANCIEROS (SCALE_FACTOR = 1.0)
        # =========================================================================
        
        # --- #BALANCE ---
        current_step += 1
        if progress_callback:
            progress_callback(current_step, total_steps, "Calculando y congelando Balance...")
        try:
            if is_consolidated:
                grupo_name = empresa_clean.replace("[GRUPO] ", "").strip()
                db = SessionLocal()
                grupo_obj = db.query(ConsolidationGroup).filter_by(nombre_grupo=grupo_name).first()
                grupo_id = grupo_obj.id if grupo_obj else None
                db.close()
                if grupo_id:
                    from src.core.consolidacion_engine import generar_hoja_trabajo
                    df_hoja_act, _ = generar_hoja_trabajo(grupo_id, periodo_clean)
                    df_hoja_comp, _ = generar_hoja_trabajo(grupo_id, comp_period_bal)
                    
                    def clean_str(s):
                        if pd.isna(s): return ""
                        return str(s).strip().lower().replace('á','a').replace('é','e').replace('í','i').replace('ó','o').replace('ú','u')

                    bs_subtotals = {
                        "activos corrientes", "activos corrientes totales",
                        "activos no corrientes", "activos no corrientes totales",
                        "total activos", "patrimonio y pasivos",
                        "pasivos corrientes", "pasivo corrientes totales",
                        "pasivos no corrientes", "pasivo no corrientes totales",
                        "total pasivos", "patrimonio", "patrimonio total",
                        "total patrimonio y pasivos", "estado de resultados"
                    }
                    bs_subtotals_normalized = {clean_str(name) for name in bs_subtotals}

                    idx_er = df_hoja_act[df_hoja_act['Balance clasificado'] == "Estado de Resultados"].index
                    df_hoja_act_sec = df_hoja_act.loc[:idx_er[0]-1] if not idx_er.empty else df_hoja_act
                    df_hoja_act_clean = df_hoja_act_sec[df_hoja_act_sec['Balance clasificado'].notna() & (df_hoja_act_sec['Balance clasificado'].str.strip() != "")]
                    df_hoja_act_clean = df_hoja_act_clean[~df_hoja_act_clean['Balance clasificado'].apply(clean_str).isin(bs_subtotals_normalized)]
                    tb_df = pd.DataFrame({'cuenta_id': df_hoja_act_clean['Balance clasificado'], 'saldo_final': df_hoja_act_clean['CONSOLIDADO']})
                    
                    tb_df_comp = None
                    if df_hoja_comp is not None:
                        idx_er_comp = df_hoja_comp[df_hoja_comp['Balance clasificado'] == "Estado de Resultados"].index
                        df_hoja_comp_sec = df_hoja_comp.loc[:idx_er_comp[0]-1] if not idx_er_comp.empty else df_hoja_comp
                        df_hoja_comp_clean = df_hoja_comp_sec[df_hoja_comp_sec['Balance clasificado'].notna() & (df_hoja_comp_sec['Balance clasificado'].str.strip() != "")]
                        df_hoja_comp_clean = df_hoja_comp_clean[~df_hoja_comp_clean['Balance clasificado'].apply(clean_str).isin(bs_subtotals_normalized)]
                        tb_df_comp = pd.DataFrame({'cuenta_id': df_hoja_comp_clean['Balance clasificado'], 'saldo_final': df_hoja_comp_clean['CONSOLIDADO']})

                    template_bal_path = os.path.join(empresa_path, "Balance clasificado.xlsx")
                    if not os.path.exists(template_bal_path):
                        template_bal_path = os.path.join("templates", "Balance clasificado.xlsx")
                    
                    dummy_map = pd.DataFrame({'N° de Cuenta': tb_df['cuenta_id'], 'Clasificación balance': tb_df['cuenta_id']})
                    from src.reporting.balance_generator import BalanceGenerator
                    gen = BalanceGenerator(template_bal_path)
                    res_bytes = gen.generate(tb_df, dummy_map, scale_factor=1.0, tb_df_comp=tb_df_comp, periodo_actual_str=periodo_clean, periodo_comp_str=comp_period_bal)
                    bal_preview = wte._extract_balance_preview(res_bytes, periodo_clean, comp_period_bal)
            else:
                tb_df = TrialBalanceDB.get_trial_balance(empresa_clean, periodo_clean)
                tb_comp = TrialBalanceDB.get_trial_balance(empresa_clean, comp_period_bal)
                template_bal_path = os.path.join(empresa_path, "Balance clasificado.xlsx")
                if not os.path.exists(template_bal_path):
                    template_bal_path = os.path.join("templates", "Balance clasificado.xlsx")
                from src.reporting.balance_generator import BalanceGenerator
                gen = BalanceGenerator(template_bal_path)
                res_bytes = gen.generate(tb_df, map_balance_df, scale_factor=1.0, tb_df_comp=tb_comp, periodo_actual_str=periodo_clean, periodo_comp_str=comp_period_bal)
                bal_preview = wte._extract_balance_preview(res_bytes, periodo_clean, comp_period_bal)

            rep_name = "ESTADO DE SITUACIÓN FINANCIERA CONSOLIDADO" if is_consolidated else "ESTADO DE SITUACIÓN FINANCIERA CLASIFICADO"
            with open(os.path.join(stmt_dir, "balance.pkl"), "wb") as f:
                pickle.dump((bal_preview, rep_name), f)
            statements_saved["#BALANCE"] = "balance.pkl"
            statements_saved["#ESF"] = "balance.pkl"
        except Exception as e:
            print(f"Error generando snapshot Balance para {empresa_clean} {periodo_clean}: {e}")

        # --- #ER (Estado de Resultados) ---
        current_step += 1
        if progress_callback:
            progress_callback(current_step, total_steps, "Calculando y congelando Estado de Resultados...")
        try:
            if is_consolidated:
                grupo_name = empresa_clean.replace("[GRUPO] ", "").strip()
                db = SessionLocal()
                grupo_obj = db.query(ConsolidationGroup).filter_by(nombre_grupo=grupo_name).first()
                grupo_id = grupo_obj.id if grupo_obj else None
                db.close()
                df_hoja_act, _ = generar_hoja_trabajo(grupo_id, periodo_clean)
                df_hoja_comp, _ = generar_hoja_trabajo(grupo_id, comp_period_pl)
                
                def clean_str(s):
                    if pd.isna(s): return ""
                    return str(s).strip().lower().replace('á','a').replace('é','e').replace('í','i').replace('ó','o').replace('ú','u')

                pl_subtotals = {
                    "ganancia bruta", "resultado antes de impuestos",
                    "ganancias (perdida) del ejercicio", "ganancia (perdida) del ejercicio",
                    "estado de resultados", "otros rubros no clasificados"
                }
                pl_subtotals_normalized = {clean_str(name) for name in pl_subtotals}

                idx_er = df_hoja_act[df_hoja_act['Balance clasificado'] == "Estado de Resultados"].index
                df_hoja_act_sec = df_hoja_act.loc[idx_er[0]+1:] if not idx_er.empty else df_hoja_act
                df_hoja_act_clean = df_hoja_act_sec[df_hoja_act_sec['Balance clasificado'].notna() & (df_hoja_act_sec['Balance clasificado'].str.strip() != "")]
                df_hoja_act_clean = df_hoja_act_clean[~df_hoja_act_clean['Balance clasificado'].apply(clean_str).isin(pl_subtotals_normalized)]
                pl_dict_act = {row['Balance clasificado']: [row['CONSOLIDADO']] for _, row in df_hoja_act_clean.iterrows()}
                pl_df_wide = pd.DataFrame(pl_dict_act)

                pl_df_comp_wide = None
                if df_hoja_comp is not None:
                    idx_er_comp = df_hoja_comp[df_hoja_comp['Balance clasificado'] == "Estado de Resultados"].index
                    df_hoja_comp_sec = df_hoja_comp.loc[idx_er_comp[0]+1:] if not idx_er_comp.empty else df_hoja_comp
                    df_hoja_comp_clean = df_hoja_comp_sec[df_hoja_comp_sec['Balance clasificado'].notna() & (df_hoja_comp_sec['Balance clasificado'].str.strip() != "")]
                    df_hoja_comp_clean = df_hoja_comp_clean[~df_hoja_comp_clean['Balance clasificado'].apply(clean_str).isin(pl_subtotals_normalized)]
                    pl_dict_comp = {row['Balance clasificado']: [row['CONSOLIDADO']] for _, row in df_hoja_comp_clean.iterrows()}
                    pl_df_comp_wide = pd.DataFrame(pl_dict_comp)

                template_er_path = os.path.join(empresa_path, "Estado de Resultados Clasificados.xlsx")
                if not os.path.exists(template_er_path):
                    template_er_path = os.path.join("templates", "Estado de Resultados Clasificados.xlsx")
                from src.reporting.er_generator import ERGenerator
                gen_er = ERGenerator(template_er_path)
                res_bytes_er, preview_er = gen_er.generate(pl_df_wide, scale_factor=1.0, pl_df_comp=pl_df_comp_wide, periodo_actual_str=periodo_clean, periodo_comp_str=comp_period_pl)
                er_preview = wte._extract_er_preview(res_bytes_er, periodo_clean, comp_period_pl)
            else:
                pl_df_wide = PlCuboDB.get_pl_cubo(empresa_clean, periodo_clean)
                pl_df_comp_wide = PlCuboDB.get_pl_cubo(empresa_clean, comp_period_pl)
                template_er_path = os.path.join(empresa_path, "Estado de Resultados Clasificados.xlsx")
                if not os.path.exists(template_er_path):
                    template_er_path = os.path.join("templates", "Estado de Resultados Clasificados.xlsx")
                from src.reporting.er_generator import ERGenerator
                gen_er = ERGenerator(template_er_path)
                res_bytes_er, preview_er = gen_er.generate(pl_df_wide, scale_factor=1.0, pl_df_comp=pl_df_comp_wide, periodo_actual_str=periodo_clean, periodo_comp_str=comp_period_pl)
                er_preview = wte._extract_er_preview(res_bytes_er, periodo_clean, comp_period_pl)

            rep_name_er = "ESTADO DE RESULTADOS CONSOLIDADO" if is_consolidated else "ESTADO DE RESULTADOS CLASIFICADOS"
            with open(os.path.join(stmt_dir, "er.pkl"), "wb") as f:
                pickle.dump((er_preview, rep_name_er), f)
            statements_saved["#ER"] = "er.pkl"
        except Exception as e:
            print(f"Error generando snapshot ER para {empresa_clean} {periodo_clean}: {e}")

        # --- #ORI (Resultados Integrales) ---
        current_step += 1
        if progress_callback:
            progress_callback(current_step, total_steps, "Calculando y congelando Resultados Integrales...")
        try:
            from src.reporting.ori_generator import OriGenerator
            template_ori_path = os.path.join(empresa_path, "Estado de Resultados Integrales.xlsx")
            if not os.path.exists(template_ori_path):
                template_ori_path = os.path.join("templates", "Estado de Resultados Integrales.xlsx")
            if os.path.exists(template_ori_path) and 'preview_er' in locals():
                gen_ori = OriGenerator(template_ori_path)
                res_bytes_ori = gen_ori.generate(preview_er, periodo_actual_str=periodo_clean, periodo_comp_str=comp_period_pl, empresa=empresa_clean)
                ori_preview = wte._extract_ori_preview(res_bytes_ori, periodo_clean, comp_period_pl)
                rep_name_ori = "ESTADO DE RESULTADOS INTEGRALES CONSOLIDADO" if is_consolidated else "ESTADO DE RESULTADOS INTEGRALES"
                with open(os.path.join(stmt_dir, "ori.pkl"), "wb") as f:
                    pickle.dump((ori_preview, rep_name_ori), f)
                statements_saved["#ORI"] = "ori.pkl"
        except Exception as e:
            print(f"Error generando snapshot ORI para {empresa_clean} {periodo_clean}: {e}")

        # --- #PATRIMONIO ---
        current_step += 1
        if progress_callback:
            progress_callback(current_step, total_steps, "Calculando y congelando Patrimonio...")
        try:
            template_pat_path = os.path.join(empresa_path, "Estado de Cambios en el Patrimonio.xlsx")
            if not os.path.exists(template_pat_path):
                template_pat_path = os.path.join("templates", "Estado de Cambios en el Patrimonio.xlsx")
            if os.path.exists(template_pat_path):
                from src.reporting.patrimonio_generator import PatrimonioGenerator
                pat_engine = PatrimonioGenerator(template_pat_path)
                bal_df_pat = bal_preview if 'bal_preview' in locals() else None
                pl_df_pat = er_preview if 'er_preview' in locals() else None
                res_bytes_pat = pat_engine.generate(bal_preview_df=bal_df_pat, pl_preview_df=pl_df_pat, periodo_actual_str=periodo_clean, periodo_comp_str=comp_period_bal, empresa=empresa_clean)
                res_bytes_pat.seek(0)
                from src.core.excel_utils import detect_patrimonio_skiprows, clean_preview_dataframe
                pat_skip = detect_patrimonio_skiprows(res_bytes_pat)
                res_bytes_pat.seek(0)
                pat_preview = pd.read_excel(res_bytes_pat, skiprows=pat_skip)
                pat_preview = clean_preview_dataframe(pat_preview)
                pat_preview = pat_preview.dropna(how='all', subset=[pat_preview.columns[0]]).reset_index(drop=True)
                rep_name_pat = "ESTADO DE CAMBIOS EN EL PATRIMONIO CONSOLIDADO" if is_consolidated else "ESTADO DE CAMBIOS EN EL PATRIMONIO"
                with open(os.path.join(stmt_dir, "patrimonio.pkl"), "wb") as f:
                    pickle.dump((pat_preview, rep_name_pat), f)
                statements_saved["#PATRIMONIO"] = "patrimonio.pkl"
        except Exception as e:
            print(f"Error generando snapshot Patrimonio para {empresa_clean} {periodo_clean}: {e}")

        # --- #EFE (Flujo de Efectivo) ---
        current_step += 1
        if progress_callback:
            progress_callback(current_step, total_steps, "Calculando y congelando Flujo de Efectivo...")
        try:
            template_cf_path = os.path.join(empresa_path, "Estado de Flujos de Efectivo.xlsx")
            if not os.path.exists(template_cf_path):
                template_cf_path = os.path.join("templates", "Estado de Flujos de Efectivo.xlsx")
            if os.path.exists(template_cf_path):
                from src.reporting.cash_flow_generator import CashFlowGenerator
                gen_cf = CashFlowGenerator(template_cf_path)
                if is_consolidated:
                    res_bytes_cf, _ = gen_cf.generate(
                        empresa=empresa_clean,
                        periodo_actual_str=periodo_clean,
                        periodo_comp_str=comp_period_bal,
                        map_balance_df=map_balance_df,
                        map_pl_df=map_pl_df,
                        method="Directo",
                        is_consolidado=True,
                        consolidated_hoja_trabajo_df=df_hoja_act if 'df_hoja_act' in locals() else None,
                        consolidated_hoja_trabajo_comp_df=df_hoja_comp if 'df_hoja_comp' in locals() else None,
                        scale_factor=1.0
                    )
                else:
                    res_bytes_cf, _ = gen_cf.generate(
                        empresa=empresa_clean,
                        periodo_actual_str=periodo_clean,
                        periodo_comp_str=comp_period_bal,
                        map_balance_df=map_balance_df,
                        map_pl_df=map_pl_df,
                        method="Directo",
                        is_consolidado=False,
                        scale_factor=1.0
                    )
                efe_preview = wte._extract_cf_preview(res_bytes_cf, periodo_clean, comp_period_bal)
                rep_name_efe = "ESTADO DE FLUJOS DE EFECTIVO CONSOLIDADO" if is_consolidated else "ESTADO DE FLUJOS DE EFECTIVO"
                with open(os.path.join(stmt_dir, "efe.pkl"), "wb") as f:
                    pickle.dump((efe_preview, rep_name_efe), f)
                statements_saved["#EFE"] = "efe.pkl"
        except Exception as e:
            print(f"Error generando snapshot EFE para {empresa_clean} {periodo_clean}: {e}")

        # =========================================================================
        # 2. GENERACIÓN DE NOTAS EXPLICATIVAS (#N01 - #N30) (SCALE_FACTOR = 1.0)
        # =========================================================================
        template_nota = "Plantilla de notas_v1.xlsx"
        if os.path.exists(template_nota):
            from src.reporting.note_generator import NoteGenerator
            from src.core.external_notes_manager import ExternalNotesManager
            from src.ui_pages.informes_y_notas import load_all_entity_contexts, evaluate_formulas_in_workbook, split_sheet_into_elements
            from src.core.excel_utils import extract_named_ranges_from_excel

            engine_notes = NoteGenerator(template_nota)
            full_registry = get_full_note_registry()

            contexts_cache = {}
            def get_contexts_for_comp(c_per):
                if c_per not in contexts_cache:
                    contexts_cache[c_per] = load_all_entity_contexts(
                        active_entity=empresa_clean,
                        periodo_actual=periodo_clean,
                        periodo_comp=c_per,
                        map_balance_df=map_balance_df,
                        map_pl_df=map_pl_df
                    )
                return contexts_cache[c_per]

            for code, info in full_registry.items():
                current_step += 1
                if progress_callback:
                    progress_callback(current_step, total_steps, f"Congelando {code}: {info.get('title', '')}...")
                
                target_sheets = info.get('sheets', [])
                if not target_sheets:
                    continue
                
                # Excluir notas solo consolidadas si es empresa individual
                if info.get('consolidated_only', False) and not is_consolidated:
                    continue

                category = info.get('category', 'resultados')
                is_balance_note = category in ["activos_corrientes", "activos_no_corrientes", "pasivos_corrientes", "pasivos_no_corrientes", "patrimonio"]
                note_comp = comp_period_bal if is_balance_note else comp_period_pl

                try:
                    entity_contexts = get_contexts_for_comp(note_comp)
                    excel_nota_out = engine_notes.generate(
                        sheet_names=target_sheets,
                        entity_contexts=entity_contexts,
                        active_entity_name=empresa_clean,
                        is_consolidated=is_consolidated,
                        scale_factor=1.0, # Pesos base exactos
                        periodo_actual_str=periodo_clean,
                        periodo_comp_str=note_comp,
                        map_balance_df=map_balance_df,
                        map_pl_df=map_pl_df,
                        evaluate_formulas=False
                    )
                    excel_nota_out = ExternalNotesManager.inject_external_data(
                        excel_nota_out,
                        empresa=empresa_clean,
                        periodo_actual=periodo_clean,
                        periodo_comp=note_comp if note_comp != "Ninguno" else None
                    )
                    excel_nota_out = evaluate_formulas_in_workbook(excel_nota_out)

                    named_ranges = extract_named_ranges_from_excel(excel_nota_out, code)
                    if named_ranges:
                        raw_els = []
                        for r_name, df_chunk in named_ranges:
                            raw_els.append(("table", df_chunk))
                    else:
                        excel_nota_out.seek(0)
                        preview_nota_df = pd.read_excel(excel_nota_out, sheet_name=target_sheets[0], header=None)
                        raw_els = split_sheet_into_elements(preview_nota_df)

                    if raw_els:
                        clean_code = code.replace("#", "").strip()
                        note_file = f"{clean_code}.pkl"
                        with open(os.path.join(notes_dir, note_file), "wb") as f:
                            pickle.dump(raw_els, f)
                        notes_saved.append(code)
                except Exception as ex_n:
                    print(f"Advertencia al congelar nota {code} para {empresa_clean} {periodo_clean}: {ex_n}")

        # =========================================================================
        # 3. METADATOS Y CONTROL RELACIONAL
        # =========================================================================
        checksum = SnapshotVaultManager._compute_checksum(snap_dir)
        meta = {
            "empresa": empresa_clean,
            "periodo": periodo_clean,
            "comp_period_bal": comp_period_bal,
            "comp_period_pl": comp_period_pl,
            "is_consolidated": is_consolidated,
            "fecha_cierre": datetime.datetime.now().isoformat(),
            "cerrado_por": user,
            "statements_count": len(statements_saved),
            "notes_count": len(notes_saved),
            "statements": list(statements_saved.keys()),
            "notes": notes_saved,
            "checksum": checksum,
            "is_active": True
        }
        with open(os.path.join(snap_dir, "metadata.json"), "w", encoding="utf-8") as f:
            json.dump(meta, f, indent=2, ensure_ascii=False)

        # Actualizar base de datos
        db = SessionLocal()
        try:
            rec = db.query(ClosingSnapshotRecord).filter_by(
                empresa=empresa_clean,
                periodo=periodo_clean
            ).first()
            if rec:
                rec.is_active = True
                rec.fecha_cierre = datetime.datetime.utcnow()
                rec.cerrado_por = user
                rec.checksum = checksum
                rec.storage_path = snap_dir
                rec.statements_count = len(statements_saved)
                rec.notes_count = len(notes_saved)
                rec.es_consolidado = is_consolidated
                rec.fecha_reapertura = None
                rec.reabierto_por = None
            else:
                new_rec = ClosingSnapshotRecord(
                    empresa=empresa_clean,
                    periodo=periodo_clean,
                    es_consolidado=is_consolidated,
                    fecha_cierre=datetime.datetime.now(),
                    cerrado_por=user,
                    is_active=True,
                    checksum=checksum,
                    storage_path=snap_dir,
                    statements_count=len(statements_saved),
                    notes_count=len(notes_saved)
                )
                db.add(new_rec)
            db.commit()
            _SNAPSHOT_STATUS_CACHE[(empresa_clean, periodo_clean)] = True
            return True, f"Snapshot congelado con éxito en Bóveda ({len(statements_saved)} estados financieros y {len(notes_saved)} notas explicativas)."
        except Exception as e_db:
            db.rollback()
            return False, f"Error al registrar snapshot en base de datos: {e_db}"
        finally:
            db.close()

    @staticmethod
    def get_statement(empresa: str, periodo_actual: str, code: str, comp_period: Optional[str] = None, scale_factor: float = 1.0) -> Optional[Tuple[pd.DataFrame, str]]:
        """
        Recupera un Estado Financiero congelado desde la Bóveda en milisegundos,
        escalando dinámicamente sus valores numéricos.
        """
        if not SnapshotVaultManager.has_active_snapshot(empresa, periodo_actual):
            return None

        snap_dir = SnapshotVaultManager.get_snapshot_dir(empresa, periodo_actual)
        stmt_dir = os.path.join(snap_dir, "statements")
        
        norm_code = code.upper().strip()
        filename_map = {
            "#BALANCE": "balance.pkl",
            "#ESF": "balance.pkl",
            "#ER": "er.pkl",
            "#ORI": "ori.pkl",
            "#PATRIMONIO": "patrimonio.pkl",
            "#EFE": "efe.pkl"
        }
        
        fname = filename_map.get(norm_code)
        if not fname:
            return None
            
        file_path = os.path.join(stmt_dir, fname)
        if not os.path.exists(file_path):
            return None

        try:
            with open(file_path, "rb") as f:
                df, rep_name = pickle.load(f)
                
            if df is not None:
                # Escalar los valores numéricos
                df_scaled = SnapshotVaultManager._scale_dataframe_cells(df, scale_factor)
                return df_scaled, rep_name
        except Exception as e:
            print(f"Error cargando statement {code} desde snapshot: {e}")
            return None

        return None

    @staticmethod
    def get_note_elements(empresa: str, periodo_actual: str, note_code: str, scale_factor: float = 1.0) -> Optional[List[Tuple[str, Any]]]:
        """
        Recupera las tablas y textos de una Nota explicativa congelada desde la Bóveda en milisegundos,
        escalando sus valores numéricos.
        """
        if not SnapshotVaultManager.has_active_snapshot(empresa, periodo_actual):
            return None

        snap_dir = SnapshotVaultManager.get_snapshot_dir(empresa, periodo_actual)
        clean_code = note_code.replace("#", "").strip()
        file_path = os.path.join(snap_dir, "notes", f"{clean_code}.pkl")
        
        if not os.path.exists(file_path):
            return None

        try:
            with open(file_path, "rb") as f:
                raw_els = pickle.load(f)
                
            scaled_els = []
            for item in raw_els:
                el_type = item[0]
                el_val = item[1]
                if el_type == "table" and isinstance(el_val, pd.DataFrame):
                    df_sc = SnapshotVaultManager._scale_dataframe_cells(el_val, scale_factor)
                    if len(item) == 4:
                        scaled_els.append((el_type, df_sc, item[2], item[3]))
                    elif len(item) == 3:
                        scaled_els.append((el_type, df_sc, item[2]))
                    else:
                        scaled_els.append((el_type, df_sc))
                else:
                    scaled_els.append(item)
            return scaled_els
        except Exception as e:
            print(f"Error cargando nota {note_code} desde snapshot: {e}")
            return None

    @staticmethod
    def archive_snapshot(empresa: str, periodo: str, user: str = "Administrador") -> Tuple[bool, str]:
        """
        Invalida y archiva el snapshot congelado al reabrir el período a memoria viva.
        """
        empresa_clean = str(empresa).strip()
        periodo_clean = str(periodo).strip()
        
        snap_dir = SnapshotVaultManager.get_snapshot_dir(empresa_clean, periodo_clean)
        if os.path.exists(snap_dir):
            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            parent_dir = os.path.dirname(os.path.dirname(snap_dir)) # data/empresas/{empresa}
            archive_dir = os.path.join(parent_dir, "snapshots_archive", f"{periodo_clean}_{ts}")
            os.makedirs(os.path.dirname(archive_dir), exist_ok=True)
            try:
                shutil.move(snap_dir, archive_dir)
            except Exception as e_mv:
                print(f"Advertencia al mover directorio de snapshot a archivo: {e_mv}")

        db = SessionLocal()
        try:
            rec = db.query(ClosingSnapshotRecord).filter_by(
                empresa=empresa_clean,
                periodo=periodo_clean,
                is_active=True
            ).first()
            if rec:
                rec.is_active = False
                rec.fecha_reapertura = datetime.datetime.now()
                rec.reabierto_por = user
                db.commit()
            _SNAPSHOT_STATUS_CACHE[(empresa_clean, periodo_clean)] = False
            return True, f"Snapshot de {periodo_clean} archivado con éxito. El período ha vuelto a memoria activa viva."
        except Exception as e:
            db.rollback()
            return False, f"Error al actualizar estado de snapshot: {e}"
        finally:
            db.close()

    @staticmethod
    def sync_all_existing_historical_periods(progress_callback=None) -> Tuple[int, int, List[str]]:
        """
        Escanea todos los períodos que ya se encuentran cerrados en la base de datos
        (tanto individuales en HistoricalDetailRecord como grupos consolidados con candado)
        y genera automáticamente sus Snapshots en la Bóveda para que queden 100% operativos.
        """
        db = SessionLocal()
        try:
            # 1. Períodos individuales cerrados
            ind_closed = db.query(
                HistoricalDetailRecord.empresa, 
                HistoricalDetailRecord.periodo
            ).distinct().all()
            
            # 2. Períodos consolidados cerrados
            grp_locks = db.query(
                ConsolidationGroup.nombre_grupo, 
                ConsolidationPeriodLock.periodo
            ).join(
                ConsolidationPeriodLock, 
                ConsolidationGroup.id == ConsolidationPeriodLock.grupo_id
            ).filter(
                ConsolidationPeriodLock.is_locked == True
            ).distinct().all()
            
            all_targets = []
            for emp, per in ind_closed:
                if emp and per and not emp.startswith("[GRUPO]"):
                    all_targets.append((emp, per))
            for grp_nom, per in grp_locks:
                full_grp = f"[GRUPO] {grp_nom}" if not grp_nom.startswith("[GRUPO]") else grp_nom
                if full_grp and per:
                    all_targets.append((full_grp, per))
        finally:
            db.close()

        total = len(all_targets)
        success_count = 0
        error_count = 0
        logs = []

        for idx, (emp, per) in enumerate(all_targets):
            if progress_callback:
                progress_callback(idx + 1, total, f"Sincronizando Bóveda: {emp} ({per})")
            
            # Si ya tiene snapshot activo, omitir
            if SnapshotVaultManager.has_active_snapshot(emp, per):
                success_count += 1
                logs.append(f"✓ {emp} - {per}: Ya en Bóveda")
                continue

            ok, msg = SnapshotVaultManager.create_snapshot(emp, per, user="Sincronizador Automático")
            if ok:
                success_count += 1
                logs.append(f"✓ {emp} - {per}: Snapshot generado con éxito")
            else:
                error_count += 1
                logs.append(f"✗ {emp} - {per}: Error {msg}")

        return success_count, error_count, logs
