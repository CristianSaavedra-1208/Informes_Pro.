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
_HYBRID_COMP_CACHE: Dict[Tuple, Tuple[float, Any]] = {}
_HYBRID_COMP_TTL_SEC = 15.0

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
    def _get_snapshot_meta(empresa: str, periodo: str) -> dict:
        """Lee el archivo metadata.json del snapshot si existe."""
        snap_dir = SnapshotVaultManager.get_snapshot_dir(empresa, periodo)
        meta_file = os.path.join(snap_dir, "metadata.json")
        if os.path.exists(meta_file):
            try:
                with open(meta_file, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return {}

    @staticmethod
    def _scale_dataframe_cells(df: pd.DataFrame, scale_factor: float, base_scale: float = 1000.0) -> pd.DataFrame:
        """Escala las columnas numéricas de un DataFrame desde base_scale hacia scale_factor."""
        if df is None or df.empty:
            return df.copy() if df is not None else None
        if not scale_factor or scale_factor <= 0:
            scale_factor = 1000.0
        if abs(float(scale_factor) - float(base_scale)) < 0.01:
            return df.copy()
            
        ratio = float(base_scale) / float(scale_factor)
        scaled = df.copy()
        for col in scaled.columns:
            scaled[col] = scaled[col].apply(
                lambda x: (x * ratio) if isinstance(x, (int, float)) and not isinstance(x, bool) and not pd.isna(x) and not (isinstance(x, int) and 1990 <= x <= 2050) else x
            )
        return scaled

    @staticmethod
    def get_frozen_mappings(empresa: str, periodo: str) -> Tuple[Optional[pd.DataFrame], Optional[pd.DataFrame]]:
        """Recupera los DataFrames de mapeo (map_balance_df, map_pl_df) congelados al cierre del período."""
        if not SnapshotVaultManager.has_active_snapshot(empresa, periodo):
            return None, None
        snap_dir = SnapshotVaultManager.get_snapshot_dir(empresa, periodo)
        ctx_file = os.path.join(snap_dir, "context.pkl")
        if os.path.exists(ctx_file):
            try:
                with open(ctx_file, "rb") as f:
                    data = pickle.load(f)
                mb = data.get("map_balance_df")
                mp = data.get("map_pl_df")
                return (mb.copy() if mb is not None else None), (mp.copy() if mp is not None else None)
            except Exception:
                pass
        return None, None

    @staticmethod
    def get_frozen_entity_context(empresa: str, periodo: str, target_company: Optional[str] = None) -> Optional[Dict[str, Any]]:
        """Recupera el contexto contable congelado ('nota1', 'nota2', 'pl') de un período cerrado."""
        if not periodo or periodo == "Ninguno":
            return None
        # Intentar primero en el snapshot de la propia entidad solicitada
        lookup_entities = []
        if target_company:
            lookup_entities.append(target_company)
        if empresa and empresa not in lookup_entities:
            lookup_entities.append(empresa)
            
        for ent in lookup_entities:
            if not SnapshotVaultManager.has_active_snapshot(ent, periodo):
                continue
            snap_dir = SnapshotVaultManager.get_snapshot_dir(ent, periodo)
            ctx_file = os.path.join(snap_dir, "context.pkl")
            if os.path.exists(ctx_file):
                try:
                    with open(ctx_file, "rb") as f:
                        data = pickle.load(f)
                    ctxs = data.get("contexts_actual", {})
                    key_co = target_company or ent
                    if key_co in ctxs:
                        import copy
                        return copy.deepcopy(ctxs[key_co])
                    if ent in ctxs:
                        import copy
                        return copy.deepcopy(ctxs[ent])
                except Exception:
                    pass
        return None

    @staticmethod
    def _build_and_save_frozen_sabanas(
        empresa_clean: str,
        periodo_clean: str,
        snap_dir: str,
        map_balance_df: Optional[pd.DataFrame] = None,
        map_pl_df: Optional[pd.DataFrame] = None
    ) -> Tuple[Optional[pd.DataFrame], Optional[pd.DataFrame]]:
        """
        Construye y congela en el Snapshot las Sábanas de Auditoría de Balance y P&L
        (en formato .pkl y .xlsx) usando exactamente los mapeos vigentes al cierre.
        """
        from src.core.sabana_builder import (
            build_balance_sabana,
            build_pl_sabana,
            build_consolidated_balance_sabana,
            build_consolidated_pl_sabana,
            get_group_info,
            _compute_monthly_adjustments,
        )
        from src.core.excel_utils import df_to_excel_bytes

        os.makedirs(snap_dir, exist_ok=True)
        is_consolidated = empresa_clean.startswith("[GRUPO]")

        if map_balance_df is None or map_pl_df is None:
            f_mb, f_mp = SnapshotVaultManager.get_frozen_mappings(empresa_clean, periodo_clean)
            if map_balance_df is None:
                map_balance_df = f_mb
            if map_pl_df is None:
                map_pl_df = f_mp
            if map_balance_df is None or map_pl_df is None:
                from src.core.word_template_engine import load_mappings_for_entity
                l_mb, l_mp = load_mappings_for_entity(empresa_clean)
                if map_balance_df is None:
                    map_balance_df = l_mb
                if map_pl_df is None:
                    map_pl_df = l_mp

        sab_bal_df = None
        sab_pl_df = None

        precomp_bal = None
        precomp_pl = None
        if is_consolidated:
            try:
                g_id, g_companies, emp_mat, emp_fil = get_group_info(empresa_clean)
                if g_companies:
                    emp_ref = emp_fil or emp_mat or g_companies[0]
                    ytd_m, m_cols, adj_b, adj_p, acc_m = _compute_monthly_adjustments(
                        g_id, periodo_clean, emp_ref, report_type="BOTH"
                    )
                    precomp_bal = (ytd_m, m_cols, adj_b, acc_m)
                    precomp_pl = (ytd_m, m_cols, adj_p, acc_m)
            except Exception as e_pre:
                print(f"Advertencia pre-calculando ajustes YTD para {empresa_clean} {periodo_clean}: {e_pre}")

        try:
            if is_consolidated:
                sab_bal_df = build_consolidated_balance_sabana(
                    empresa_clean, periodo_clean, map_balance_df, precomputed_adjustments=precomp_bal
                )
            else:
                tb_df = TrialBalanceDB.get_trial_balance(empresa_clean, periodo_clean)
                sab_bal_df = build_balance_sabana(tb_df, map_balance_df)
            if sab_bal_df is not None and not sab_bal_df.empty:
                with open(os.path.join(snap_dir, "sabana_balance.pkl"), "wb") as f_sb:
                    pickle.dump(sab_bal_df, f_sb)
                excel_bal = df_to_excel_bytes(sab_bal_df, sheet_name="Matriz_Auditoria")
                if excel_bal is not None:
                    raw_b = excel_bal.getvalue() if hasattr(excel_bal, "getvalue") else excel_bal
                    with open(os.path.join(snap_dir, "sabana_balance.xlsx"), "wb") as f_sbx:
                        f_sbx.write(raw_b)
        except Exception as e_b:
            print(f"Advertencia congelando sabana_balance para {empresa_clean} {periodo_clean}: {e_b}")

        try:
            if is_consolidated:
                sab_pl_df = build_consolidated_pl_sabana(
                    empresa_clean, periodo_clean, map_pl_df, precomputed_adjustments=precomp_pl
                )
            else:
                pl_df = PlCuboDB.get_pl_cubo(empresa_clean, periodo_clean)
                tb_df = TrialBalanceDB.get_trial_balance(empresa_clean, periodo_clean)
                sab_pl_df = build_pl_sabana(pl_df, map_pl_df, tb_df)
            if sab_pl_df is not None and not sab_pl_df.empty:
                with open(os.path.join(snap_dir, "sabana_pl.pkl"), "wb") as f_sp:
                    pickle.dump(sab_pl_df, f_sp)
                excel_pl = df_to_excel_bytes(sab_pl_df, sheet_name="Matriz_Auditoria")
                if excel_pl is not None:
                    raw_p = excel_pl.getvalue() if hasattr(excel_pl, "getvalue") else excel_pl
                    with open(os.path.join(snap_dir, "sabana_pl.xlsx"), "wb") as f_spx:
                        f_spx.write(raw_p)
        except Exception as e_p:
            print(f"Advertencia congelando sabana_pl para {empresa_clean} {periodo_clean}: {e_p}")

        return sab_bal_df, sab_pl_df

    @staticmethod
    def get_frozen_sabana(empresa: str, periodo: str, sabana_type: str = "BALANCE") -> Optional[pd.DataFrame]:
        """
        Recupera la Sábana de Auditoría congelada (Balance o P&L) desde el Snapshot del período cerrado.
        Si el snapshot existía antes de incorporar sabana_*.pkl, la genera automáticamente usando los
        mapeos congelados de context.pkl y la persiste en el snapshot.
        """
        empresa_clean = str(empresa).strip()
        periodo_clean = str(periodo).strip()
        if not SnapshotVaultManager.has_active_snapshot(empresa_clean, periodo_clean):
            return None

        snap_dir = SnapshotVaultManager.get_snapshot_dir(empresa_clean, periodo_clean)
        is_bal = "BAL" in str(sabana_type).upper()
        pkl_name = "sabana_balance.pkl" if is_bal else "sabana_pl.pkl"
        pkl_path = os.path.join(snap_dir, pkl_name)

        if os.path.exists(pkl_path):
            try:
                with open(pkl_path, "rb") as f:
                    df = pickle.load(f)
                if df is not None and not df.empty:
                    return df.copy()
            except Exception:
                pass

        # Si el snapshot no tenía aún el archivo pkl de sábana, construirlo con los mapeos congelados de context.pkl
        f_mb, f_mp = SnapshotVaultManager.get_frozen_mappings(empresa_clean, periodo_clean)
        sab_bal_df, sab_pl_df = SnapshotVaultManager._build_and_save_frozen_sabanas(
            empresa_clean=empresa_clean,
            periodo_clean=periodo_clean,
            snap_dir=snap_dir,
            map_balance_df=f_mb,
            map_pl_df=f_mp
        )
        target_df = sab_bal_df if is_bal else sab_pl_df
        return target_df.copy() if (target_df is not None and not target_df.empty) else None

    @staticmethod
    def get_frozen_sabana_excel(empresa: str, periodo: str, sabana_type: str = "BALANCE") -> Optional[bytes]:
        """Recupera los bytes del archivo Excel (.xlsx) de la Sábana de Auditoría congelada en el Snapshot."""
        empresa_clean = str(empresa).strip()
        periodo_clean = str(periodo).strip()
        if not SnapshotVaultManager.has_active_snapshot(empresa_clean, periodo_clean):
            return None

        snap_dir = SnapshotVaultManager.get_snapshot_dir(empresa_clean, periodo_clean)
        is_bal = "BAL" in str(sabana_type).upper()
        xlsx_name = "sabana_balance.xlsx" if is_bal else "sabana_pl.xlsx"
        xlsx_path = os.path.join(snap_dir, xlsx_name)

        if not os.path.exists(xlsx_path):
            SnapshotVaultManager.get_frozen_sabana(empresa_clean, periodo_clean, sabana_type=sabana_type)

        if os.path.exists(xlsx_path):
            try:
                with open(xlsx_path, "rb") as f:
                    return f.read()
            except Exception:
                pass
        return None

    @staticmethod
    def create_snapshot(empresa: str, periodo: str, user: str = "Administrador", progress_callback=None) -> Tuple[bool, str]:
        """
        Calcula y congela en la Bóveda todos los Estados Financieros y Notas para una entidad y período.
        Valores almacenados en unidad estándar M$ (base_scale = 1000.0) idéntica a plantillas y anexos.
        """
        base_scale = 1000.0
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

        # Guardar mapeos y contexto contable congelado del período (context.pkl)
        try:
            from src.ui_pages.informes_y_notas import load_all_entity_contexts
            from src.core.sabana_manager import SabanaManager
            SabanaManager.clear_sabana_cache()
            raw_ctx_snap = load_all_entity_contexts(
                active_entity=empresa_clean,
                periodo_actual=periodo_clean,
                periodo_comp="Ninguno",
                map_balance_df=map_balance_df,
                map_pl_df=map_pl_df
            )
            frozen_ctx_data = {
                "map_balance_df": map_balance_df.copy() if map_balance_df is not None else None,
                "map_pl_df": map_pl_df.copy() if map_pl_df is not None else None,
                "contexts_actual": {k: v.get("actual", {}) for k, v in (raw_ctx_snap or {}).items() if isinstance(v, dict)}
            }
            with open(os.path.join(snap_dir, "context.pkl"), "wb") as f_ctx:
                pickle.dump(frozen_ctx_data, f_ctx)
        except Exception as e_ctx:
            print(f"Advertencia guardando context.pkl para {empresa_clean} {periodo_clean}: {e_ctx}")

        # Guardar Sábanas de Auditoría congeladas (Balance y P&L) en el Snapshot
        try:
            SnapshotVaultManager._build_and_save_frozen_sabanas(
                empresa_clean=empresa_clean,
                periodo_clean=periodo_clean,
                snap_dir=snap_dir,
                map_balance_df=map_balance_df,
                map_pl_df=map_pl_df
            )
        except Exception as e_sab:
            print(f"Advertencia guardando sábanas de auditoría para {empresa_clean} {periodo_clean}: {e_sab}")
        
        empresa_path = os.path.join("data", "empresas", empresa_clean)
        if not os.path.exists(empresa_path):
            empresa_path = os.path.join("data", "empresas", empresa_clean.replace("[GRUPO] ", "").strip())

        statements_saved = {}
        notes_saved = []

        total_steps = 5 + len(get_full_note_registry())
        current_step = 0

        # =========================================================================
        # 1. GENERACIÓN DE ESTADOS FINANCIEROS (BASE_SCALE = 1000.0 M$)
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

                    def is_pl_line_local(li):
                        norm = str(li).lower().strip().replace('á','a').replace('é','e').replace('í','i').replace('ó','o').replace('ú','u')
                        if any(x in norm for x in ["ingreso", "gasto", "costo", "egreso", "depreciacion", "amortizacion", "diferencia de cambio", "unidad de reajuste", "ganancia", "perdida"]):
                            if not any(x in norm for x in ["diferido", "diferidos", "anticipado", "anticipados", "por pagar", "por cobrar", "acumulado", "acumulados", "acumulada", "acumuladas", "capital", "reserva", "ejercicio", "periodo"]):
                                return True
                        return False

                    tb_df = tb_df[~tb_df['cuenta_id'].apply(is_pl_line_local)]
                    if tb_df_comp is not None:
                        tb_df_comp = tb_df_comp[~tb_df_comp['cuenta_id'].apply(is_pl_line_local)]

                    template_bal_path = os.path.join(empresa_path, "Balance clasificado.xlsx")
                    if not os.path.exists(template_bal_path):
                        template_bal_path = os.path.join("templates", "Balance clasificado.xlsx")
                    
                    dummy_map = pd.DataFrame({'N° de Cuenta': tb_df['cuenta_id'], 'Clasificación balance': tb_df['cuenta_id']})
                    from src.reporting.balance_generator import BalanceGenerator
                    gen = BalanceGenerator(template_bal_path)
                    res_bytes = gen.generate(tb_df, dummy_map, scale_factor=base_scale, tb_df_comp=tb_df_comp, periodo_actual_str=periodo_clean, periodo_comp_str=comp_period_bal)
                    bal_preview = wte._extract_balance_preview(res_bytes, periodo_clean, comp_period_bal)
            else:
                tb_df = TrialBalanceDB.get_trial_balance(empresa_clean, periodo_clean)
                tb_comp = TrialBalanceDB.get_trial_balance(empresa_clean, comp_period_bal)
                template_bal_path = os.path.join(empresa_path, "Balance clasificado.xlsx")
                if not os.path.exists(template_bal_path):
                    template_bal_path = os.path.join("templates", "Balance clasificado.xlsx")
                from src.reporting.balance_generator import BalanceGenerator
                gen = BalanceGenerator(template_bal_path)
                res_bytes = gen.generate(tb_df, map_balance_df, scale_factor=base_scale, tb_df_comp=tb_comp, periodo_actual_str=periodo_clean, periodo_comp_str=comp_period_bal)
                bal_preview = wte._extract_balance_preview(res_bytes, periodo_clean, comp_period_bal)

            rep_name = "ESTADO DE SITUACIÓN FINANCIERA CONSOLIDADO" if is_consolidated else "ESTADO DE SITUACIÓN FINANCIERA CLASIFICADO"
            with open(os.path.join(stmt_dir, "balance.pkl"), "wb") as f:
                pickle.dump((bal_preview, rep_name), f)
            if 'res_bytes' in locals() and res_bytes is not None:
                with open(os.path.join(stmt_dir, "balance.xlsx"), "wb") as f_ex:
                    f_ex.write(res_bytes.getvalue())
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
                from src.core.consolidacion_engine import generar_hoja_trabajo
                if 'df_hoja_act' not in locals() or df_hoja_act is None:
                    df_hoja_act, _ = generar_hoja_trabajo(grupo_id, periodo_clean)
                if comp_period_pl == comp_period_bal and 'df_hoja_comp' in locals() and df_hoja_comp is not None:
                    df_hoja_comp_er = df_hoja_comp
                else:
                    df_hoja_comp_er, _ = generar_hoja_trabajo(grupo_id, comp_period_pl)
                
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
                if df_hoja_comp_er is not None:
                    idx_er_comp = df_hoja_comp_er[df_hoja_comp_er['Balance clasificado'] == "Estado de Resultados"].index
                    df_hoja_comp_sec = df_hoja_comp_er.loc[idx_er_comp[0]+1:] if not idx_er_comp.empty else df_hoja_comp_er
                    df_hoja_comp_clean = df_hoja_comp_sec[df_hoja_comp_sec['Balance clasificado'].notna() & (df_hoja_comp_sec['Balance clasificado'].str.strip() != "")]
                    df_hoja_comp_clean = df_hoja_comp_clean[~df_hoja_comp_clean['Balance clasificado'].apply(clean_str).isin(pl_subtotals_normalized)]
                    pl_dict_comp = {row['Balance clasificado']: [row['CONSOLIDADO']] for _, row in df_hoja_comp_clean.iterrows()}
                    pl_df_comp_wide = pd.DataFrame(pl_dict_comp)

                template_er_path = os.path.join(empresa_path, "Estado de Resultados Clasificados.xlsx")
                if not os.path.exists(template_er_path):
                    template_er_path = os.path.join("templates", "Estado de Resultados Clasificados.xlsx")
                from src.reporting.er_generator import ERGenerator
                gen_er = ERGenerator(template_er_path)
                res_bytes_er, preview_er = gen_er.generate(pl_df_wide, scale_factor=base_scale, pl_df_comp=pl_df_comp_wide, periodo_actual_str=periodo_clean, periodo_comp_str=comp_period_pl)
                er_preview = wte._extract_er_preview(res_bytes_er, periodo_clean, comp_period_pl)
            else:
                pl_df_wide = PlCuboDB.get_pl_cubo(empresa_clean, periodo_clean)
                pl_df_comp_wide = PlCuboDB.get_pl_cubo(empresa_clean, comp_period_pl)
                template_er_path = os.path.join(empresa_path, "Estado de Resultados Clasificados.xlsx")
                if not os.path.exists(template_er_path):
                    template_er_path = os.path.join("templates", "Estado de Resultados Clasificados.xlsx")
                from src.reporting.er_generator import ERGenerator
                gen_er = ERGenerator(template_er_path)
                res_bytes_er, preview_er = gen_er.generate(pl_df_wide, scale_factor=base_scale, pl_df_comp=pl_df_comp_wide, periodo_actual_str=periodo_clean, periodo_comp_str=comp_period_pl)
                er_preview = wte._extract_er_preview(res_bytes_er, periodo_clean, comp_period_pl)

            rep_name_er = "ESTADO DE RESULTADOS CONSOLIDADO" if is_consolidated else "ESTADO DE RESULTADOS CLASIFICADOS"
            with open(os.path.join(stmt_dir, "er.pkl"), "wb") as f:
                pickle.dump((er_preview, rep_name_er), f)
            if 'res_bytes_er' in locals() and res_bytes_er is not None:
                with open(os.path.join(stmt_dir, "er.xlsx"), "wb") as f_ex:
                    f_ex.write(res_bytes_er.getvalue())
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
                if res_bytes_ori is not None:
                    with open(os.path.join(stmt_dir, "ori.xlsx"), "wb") as f_ex:
                        f_ex.write(res_bytes_ori.getvalue())
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
                pl_df_pat = None
                if is_consolidated:
                    grupo_name = empresa_clean.replace("[GRUPO] ", "").strip()
                    db = SessionLocal()
                    try:
                        grupo_obj = db.query(ConsolidationGroup).filter_by(nombre_grupo=grupo_name).first()
                        grupo_id = grupo_obj.id if grupo_obj else None
                    finally:
                        db.close()
                    if grupo_id:
                        bal_df_pat, pl_df_pat, _, _ = wte._get_consolidated_bal_and_pl(
                            grupo_id=grupo_id,
                            periodo_actual=periodo_clean,
                            comp_period=comp_period_bal,
                            scale_factor=base_scale,
                            empresa_path=empresa_path,
                            grupo_name=grupo_name
                        )
                else:
                    pl_df_wide_p = PlCuboDB.get_pl_cubo(empresa_clean, periodo_clean)
                    pl_df_comp_wide_p = PlCuboDB.get_pl_cubo(empresa_clean, comp_period_bal)
                    template_er_p = os.path.join(empresa_path, "Estado de Resultados Clasificados.xlsx")
                    if not os.path.exists(template_er_p):
                        template_er_p = os.path.join("templates", "Estado de Resultados Clasificados.xlsx")
                    from src.reporting.er_generator import ERGenerator
                    gen_er_p = ERGenerator(template_er_p)
                    _, pl_df_pat = gen_er_p.generate(pl_df_wide_p, scale_factor=base_scale, pl_df_comp=pl_df_comp_wide_p, periodo_actual_str=periodo_clean, periodo_comp_str=comp_period_bal)

                res_bytes_pat = pat_engine.generate(bal_preview_df=bal_df_pat, pl_preview_df=pl_df_pat, periodo_actual_str=periodo_clean, periodo_comp_str=comp_period_bal, empresa=empresa_clean)
                res_bytes_pat.seek(0)
                from src.core.excel_utils import detect_patrimonio_skiprows, clean_patrimonio_preview_df
                pat_skip = detect_patrimonio_skiprows(res_bytes_pat)
                res_bytes_pat.seek(0)
                pat_preview = pd.read_excel(res_bytes_pat, skiprows=pat_skip)
                pat_preview = clean_patrimonio_preview_df(pat_preview)
                rep_name_pat = "ESTADO DE CAMBIOS EN EL PATRIMONIO CONSOLIDADO" if is_consolidated else "ESTADO DE CAMBIOS EN EL PATRIMONIO"
                with open(os.path.join(stmt_dir, "patrimonio.pkl"), "wb") as f:
                    pickle.dump((pat_preview, rep_name_pat), f)
                if res_bytes_pat is not None:
                    with open(os.path.join(stmt_dir, "patrimonio.xlsx"), "wb") as f_ex:
                        f_ex.write(res_bytes_pat.getvalue())
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
                        scale_factor=base_scale
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
                        scale_factor=base_scale
                    )
                efe_preview = wte._extract_cf_preview(res_bytes_cf, periodo_clean, comp_period_bal)
                rep_name_efe = "ESTADO DE FLUJOS DE EFECTIVO CONSOLIDADO" if is_consolidated else "ESTADO DE FLUJOS DE EFECTIVO"
                with open(os.path.join(stmt_dir, "efe.pkl"), "wb") as f:
                    pickle.dump((efe_preview, rep_name_efe), f)
                if res_bytes_cf is not None:
                    with open(os.path.join(stmt_dir, "efe.xlsx"), "wb") as f_ex:
                        f_ex.write(res_bytes_cf.getvalue())
                statements_saved["#EFE"] = "efe.pkl"
        except Exception as e:
            print(f"Error generando snapshot EFE para {empresa_clean} {periodo_clean}: {e}")

        # =========================================================================
        # 2. GENERACIÓN DE NOTAS EXPLICATIVAS (#N01 - #N30) (BASE_SCALE = 1000.0 M$)
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

            # Agrupar todas las hojas requeridas por período comparativo (comp_period_bal vs comp_period_pl)
            # para evaluar el libro maestro 1 o 2 veces en total (en lugar de 22 veces), acelerando >10x.
            sheets_by_comp = {}
            for code, info in full_registry.items():
                t_sheets = info.get('sheets', [])
                if not t_sheets:
                    continue
                if info.get('consolidated_only', False) and not is_consolidated:
                    continue
                cat = info.get('category', 'resultados')
                is_bal_n = cat in ["activos_corrientes", "activos_no_corrientes", "pasivos_corrientes", "pasivos_no_corrientes", "patrimonio"]
                n_comp = comp_period_bal if is_bal_n else comp_period_pl
                if n_comp not in sheets_by_comp:
                    sheets_by_comp[n_comp] = []
                for s_n in t_sheets:
                    if s_n not in sheets_by_comp[n_comp]:
                        sheets_by_comp[n_comp].append(s_n)

            evaluated_wb_cache = {}
            for n_comp, all_comp_sheets in sheets_by_comp.items():
                try:
                    entity_contexts = get_contexts_for_comp(n_comp)
                    wb_out = engine_notes.generate(
                        sheet_names=all_comp_sheets,
                        entity_contexts=entity_contexts,
                        active_entity_name=empresa_clean,
                        is_consolidated=is_consolidated,
                        scale_factor=base_scale,
                        periodo_actual_str=periodo_clean,
                        periodo_comp_str=n_comp,
                        map_balance_df=map_balance_df,
                        map_pl_df=map_pl_df,
                        evaluate_formulas=False
                    )
                    wb_out = ExternalNotesManager.inject_external_data(
                        wb_out,
                        empresa=empresa_clean,
                        periodo_actual=periodo_clean,
                        periodo_comp=n_comp if n_comp != "Ninguno" else None
                    )
                    from src.core.rounding_reconciler import RoundingReconciler
                    _, wb_eval_rec = RoundingReconciler.reconcile_and_evaluate_note_workbook(
                        wb_out,
                        empresa=empresa_clean,
                        periodo_actual=periodo_clean,
                        scale_factor=base_scale,
                        df_bal=bal_preview if 'bal_preview' in locals() else None,
                        df_er=er_preview if 'er_preview' in locals() else None,
                        target_sheets=all_comp_sheets,
                        return_raw=False,
                    )
                    wb_bytes = wb_eval_rec.getvalue()
                    import openpyxl
                    wb_openpyxl_eval = openpyxl.load_workbook(BytesIO(wb_bytes), data_only=True)
                    evaluated_wb_cache[n_comp] = (wb_bytes, pd.ExcelFile(BytesIO(wb_bytes)), wb_openpyxl_eval)
                except Exception as ex_wb:
                    print(f"Advertencia pre-evaluando libro de notas para {empresa_clean} {periodo_clean} (comp {n_comp}): {ex_wb}")

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
                    if note_comp not in evaluated_wb_cache:
                        continue
                    wb_bytes, xl_nota_f, wb_openpyxl_eval = evaluated_wb_cache[note_comp]

                    named_ranges = extract_named_ranges_from_excel(None, code, wb=wb_openpyxl_eval)
                    if named_ranges:
                        raw_els = []
                        for r_name, df_chunk in named_ranges:
                            raw_els.append(("table", df_chunk, target_sheets[0], None))
                    else:
                        raw_els = []
                        for s_name in target_sheets:
                            if s_name in xl_nota_f.sheet_names:
                                preview_nota_df = xl_nota_f.parse(sheet_name=s_name, header=None)
                                sheet_els = split_sheet_into_elements(preview_nota_df, sheet_name=s_name)
                                for s_item in sheet_els:
                                    el_t = s_item[0]
                                    el_v = s_item[1]
                                    sh_n = s_item[2] if len(s_item) > 2 else s_name
                                    exp_c = s_item[3] if len(s_item) > 3 else None
                                    raw_els.append((el_t, el_v, sh_n, exp_c))

                    if raw_els:
                        clean_code = code.replace("#", "").strip()
                        note_file = f"{clean_code}.pkl"
                        with open(os.path.join(notes_dir, note_file), "wb") as f:
                            pickle.dump(raw_els, f)
                        with open(os.path.join(notes_dir, f"{clean_code}.xlsx"), "wb") as f_ex:
                            f_ex.write(wb_bytes)
                        notes_saved.append(code)
                except Exception as ex_n:
                    print(f"Advertencia al congelar nota {code} para {empresa_clean} {periodo_clean}: {ex_n}")

            for _, xl_cached, wb_cached in evaluated_wb_cache.values():
                try:
                    xl_cached.close()
                except Exception:
                    pass
                try:
                    wb_cached.close()
                except Exception:
                    pass

        # =========================================================================
        # 3. METADATOS Y CONTROL RELACIONAL
        # =========================================================================
        checksum = SnapshotVaultManager._compute_checksum(snap_dir)
        meta = {
            "empresa": empresa_clean,
            "periodo": periodo_clean,
            "comp_period_bal": comp_period_bal,
            "comp_period_pl": comp_period_pl,
            "base_scale": base_scale,
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
            _HYBRID_COMP_CACHE.clear()
            return True, f"Snapshot congelado con éxito en Bóveda ({len(statements_saved)} estados financieros y {len(notes_saved)} notas explicativas)."
        except Exception as e_db:
            db.rollback()
            return False, f"Error al registrar snapshot en base de datos: {e_db}"
        finally:
            db.close()

    @staticmethod
    def _norm_stmt_code(code: str) -> str:
        norm_code = str(code or "").upper().strip()
        if not norm_code.startswith("#"):
            norm_code = "#" + norm_code
        return norm_code

    @staticmethod
    def _is_no_comp(comp_period: Optional[str]) -> bool:
        return comp_period is not None and str(comp_period).strip().lower() in ["ninguno", "none", ""]

    @staticmethod
    def _get_stmt_native_comp(meta: Dict[str, Any], norm_code: str) -> Optional[str]:
        if norm_code in ["#BALANCE", "#ESF", "#PATRIMONIO", "#EFE"]:
            return meta.get("comp_period_bal")
        return meta.get("comp_period_pl")

    @staticmethod
    def _get_note_native_comp(meta: Dict[str, Any], note_code: str) -> Optional[str]:
        norm_note = "#" + str(note_code or "").replace("#", "").strip().upper()
        full_reg = get_full_note_registry()
        info = full_reg.get(norm_note, {})
        cat = info.get("category", "resultados")
        is_bal_note = cat in [
            "activos_corrientes",
            "activos_no_corrientes",
            "pasivos_corrientes",
            "pasivos_no_corrientes",
            "patrimonio",
        ]
        return meta.get("comp_period_bal") if is_bal_note else meta.get("comp_period_pl")

    @staticmethod
    def _splice_comp_into_statement_df(
        df_actual_frozen: pd.DataFrame,
        df_comp_source: Optional[pd.DataFrame],
        comp_period: Optional[str],
    ) -> pd.DataFrame:
        """
        Preserva el 100% de la columna Actual congelada en df_actual_frozen e inyecta en la columna
        Comparativa los saldos provenientes de df_comp_source (Snapshot cerrado o cálculo al vuelo)
        o vacía la columna si comp_period es 'Ninguno'.
        """
        if df_actual_frozen is None or df_actual_frozen.empty or len(df_actual_frozen.columns) < 2:
            return df_actual_frozen

        import numpy as np
        import re as _re
        import unicodedata as _ud

        out_df = df_actual_frozen.copy()
        live_date_cols = [
            c for c in out_df.columns[1:]
            if _re.search(r'20\d{2}', str(c)) or str(c).strip().lower() not in ["nota", "notas", "note", "notes"]
        ]
        if len(live_date_cols) < 2:
            return out_df

        act_col = live_date_cols[0]
        comp_col = live_date_cols[1]

        if SnapshotVaultManager._is_no_comp(comp_period):
            out_df[comp_col] = np.nan
            if str(comp_col) != "Comp":
                out_df = out_df.rename(columns={comp_col: "Comp"})
            return out_df

        def _norm_lbl(v: Any) -> str:
            if pd.isna(v):
                return ""
            s = str(v).strip().lower()
            s = "".join(ch for ch in _ud.normalize("NFD", s) if _ud.category(ch) != "Mn")
            return _re.sub(r"\s+", " ", s)

        comp_str = str(comp_period).strip()
        target_col_name = f"{comp_str} (Comp)" if comp_str == str(act_col) else comp_str

        if df_comp_source is not None and not df_comp_source.empty and len(df_comp_source.columns) >= 2:
            src_date_cols = [
                c for c in df_comp_source.columns[1:]
                if _re.search(r'20\d{2}', str(c)) or str(c).strip().lower() not in ["nota", "notas", "note", "notes"]
            ]
            if src_date_cols:
                src_act_col = src_date_cols[0]
                if len(out_df) == len(df_comp_source):
                    out_df[comp_col] = df_comp_source[src_act_col].values
                else:
                    src_occ_map: Dict[Tuple[str, int], Any] = {}
                    src_first_map: Dict[str, Any] = {}
                    occ_counts_src: Dict[str, int] = {}
                    for _, r_s in df_comp_source.iterrows():
                        k = _norm_lbl(r_s.iloc[0])
                        if k:
                            occ = occ_counts_src.get(k, 0)
                            occ_counts_src[k] = occ + 1
                            src_occ_map[(k, occ)] = r_s[src_act_col]
                            if k not in src_first_map:
                                src_first_map[k] = r_s[src_act_col]

                    occ_counts_dst: Dict[str, int] = {}
                    new_vals = []
                    for idx_l, r_l in out_df.iterrows():
                        k_l = _norm_lbl(r_l.iloc[0])
                        if k_l:
                            occ_d = occ_counts_dst.get(k_l, 0)
                            occ_counts_dst[k_l] = occ_d + 1
                            if (k_l, occ_d) in src_occ_map:
                                new_vals.append(src_occ_map[(k_l, occ_d)])
                            elif k_l in src_first_map:
                                new_vals.append(src_first_map[k_l])
                            else:
                                new_vals.append(np.nan)
                        else:
                            new_vals.append(np.nan)
                    out_df[comp_col] = new_vals
        else:
            # Período abierto sin datos cargados en BD: limpiar columna (preservando solo N° Acciones si aplica)
            cleaned_vals = []
            for _, r_l in out_df.iterrows():
                k_l = _norm_lbl(r_l.iloc[0])
                if "acciones" in k_l:
                    cleaned_vals.append(r_l[act_col])
                else:
                    cleaned_vals.append(np.nan)
            out_df[comp_col] = cleaned_vals

        if str(comp_col) != target_col_name:
            out_df = out_df.rename(columns={comp_col: target_col_name})
        return out_df

    @staticmethod
    def _compute_live_statement_for_period(
        empresa: str,
        periodo: str,
        norm_code: str,
        scale_factor: float = 1000.0,
        target_lang: str = "es",
    ) -> Optional[pd.DataFrame]:
        """
        Calcula al vuelo (en memoria viva) el Estado Financiero de un período abierto `periodo`
        para extraer su columna principal como comparativo híbrido.
        """
        import time as _time
        empresa_clean = str(empresa or "").strip()
        periodo_clean = str(periodo or "").strip()
        norm_c = SnapshotVaultManager._norm_stmt_code(norm_code)
        if norm_c == "#ESF":
            norm_c = "#BALANCE"

        cache_key = ("stmt_live", empresa_clean, periodo_clean, norm_c, float(scale_factor), str(target_lang))
        cached = _HYBRID_COMP_CACHE.get(cache_key)
        if cached and (_time.time() - cached[0]) < _HYBRID_COMP_TTL_SEC:
            return cached[1].copy() if cached[1] is not None else None

        is_consolidated = empresa_clean.startswith("[GRUPO]")
        clean_folder_name = empresa_clean.replace("[GRUPO] ", "").strip() if is_consolidated else empresa_clean
        empresa_path = os.path.join("data", "empresas", clean_folder_name)
        if not os.path.exists(empresa_path) and os.path.exists(os.path.join("data", "empresas", empresa_clean)):
            empresa_path = os.path.join("data", "empresas", empresa_clean)

        from src.core.word_template_engine import WordTemplateEngine, load_mappings_for_entity
        map_balance_df, map_pl_df = load_mappings_for_entity(empresa_clean)
        wte = WordTemplateEngine(None)

        def _clean_str(s: Any) -> str:
            if pd.isna(s):
                return ""
            return (
                str(s).strip().lower()
                .replace("á", "a").replace("é", "e").replace("í", "i")
                .replace("ó", "o").replace("ú", "u")
            )

        out_df: Optional[pd.DataFrame] = None
        try:
            grupo_id = None
            df_hoja_act = None
            if is_consolidated:
                db = SessionLocal()
                try:
                    grupo_obj = db.query(ConsolidationGroup).filter_by(nombre_grupo=clean_folder_name).first()
                    grupo_id = grupo_obj.id if grupo_obj else None
                finally:
                    db.close()
                if not grupo_id:
                    _HYBRID_COMP_CACHE[cache_key] = (_time.time(), None)
                    return None
                from src.core.consolidacion_engine import generar_hoja_trabajo
                df_hoja_act, _ = generar_hoja_trabajo(grupo_id, periodo_clean)
                if df_hoja_act is None or df_hoja_act.empty:
                    _HYBRID_COMP_CACHE[cache_key] = (_time.time(), None)
                    return None

            if norm_c == "#BALANCE":
                template_bal_path = os.path.join(empresa_path, "Balance clasificado.xlsx")
                if not os.path.exists(template_bal_path):
                    template_bal_path = os.path.join("templates", "Balance clasificado.xlsx")
                if not os.path.exists(template_bal_path):
                    return None

                from src.reporting.balance_generator import BalanceGenerator
                gen_bal = BalanceGenerator(template_bal_path)
                if is_consolidated:
                    bs_subtotals = {
                        "activos corrientes", "activos corrientes totales",
                        "activos no corrientes", "activos no corrientes totales",
                        "total activos", "patrimonio y pasivos",
                        "pasivos corrientes", "pasivo corrientes totales",
                        "pasivos no corrientes", "pasivo no corrientes totales",
                        "total pasivos", "patrimonio", "patrimonio total",
                        "total patrimonio y pasivos", "estado de resultados",
                    }
                    bs_sub_norm = {_clean_str(x) for x in bs_subtotals}
                    idx_er = df_hoja_act[df_hoja_act["Balance clasificado"] == "Estado de Resultados"].index
                    df_sec = df_hoja_act.loc[:idx_er[0] - 1] if not idx_er.empty else df_hoja_act
                    df_clean = df_sec[df_sec["Balance clasificado"].notna() & (df_sec["Balance clasificado"].astype(str).str.strip() != "")]
                    df_clean = df_clean[~df_clean["Balance clasificado"].apply(_clean_str).isin(bs_sub_norm)]
                    tb_df = pd.DataFrame({"cuenta_id": df_clean["Balance clasificado"], "saldo_final": df_clean["CONSOLIDADO"]})

                    def _is_pl_line(li: Any) -> bool:
                        norm = _clean_str(li)
                        if any(x in norm for x in ["ingreso", "gasto", "costo", "egreso", "depreciacion", "amortizacion", "diferencia de cambio", "unidad de reajuste", "ganancia", "perdida"]):
                            if not any(x in norm for x in ["diferido", "diferidos", "anticipado", "anticipados", "por pagar", "por cobrar", "acumulado", "acumulados", "acumulada", "acumuladas", "capital", "reserva", "ejercicio", "periodo"]):
                                return True
                        return False

                    tb_df = tb_df[~tb_df["cuenta_id"].apply(_is_pl_line)]
                    if tb_df.empty or pd.to_numeric(tb_df["saldo_final"], errors="coerce").fillna(0.0).abs().sum() == 0.0:
                        _HYBRID_COMP_CACHE[cache_key] = (_time.time(), None)
                        return None
                    dummy_map = pd.DataFrame({"N° de Cuenta": tb_df["cuenta_id"], "Clasificación balance": tb_df["cuenta_id"]})
                    res_bytes = gen_bal.generate(
                        tb_df, dummy_map, scale_factor=scale_factor, tb_df_comp=None,
                        periodo_actual_str=periodo_clean, periodo_comp_str=None, target_lang=target_lang
                    )
                else:
                    tb_df = TrialBalanceDB.get_trial_balance(empresa_clean, periodo_clean)
                    if tb_df is None or tb_df.empty:
                        _HYBRID_COMP_CACHE[cache_key] = (_time.time(), None)
                        return None
                    res_bytes = gen_bal.generate(
                        tb_df, map_balance_df, scale_factor=scale_factor, tb_df_comp=None,
                        periodo_actual_str=periodo_clean, periodo_comp_str=None, target_lang=target_lang
                    )
                out_df = wte._extract_balance_preview(res_bytes, periodo_clean, "Comp")

            elif norm_c in ["#ER", "#ORI"]:
                template_er_path = os.path.join(empresa_path, "Estado de Resultados Clasificados.xlsx")
                if not os.path.exists(template_er_path):
                    template_er_path = os.path.join("templates", "Estado de Resultados Clasificados.xlsx")
                if not os.path.exists(template_er_path):
                    return None

                if is_consolidated:
                    pl_subtotals = {
                        "ganancia bruta", "resultado antes de impuestos",
                        "ganancias (perdida) del ejercicio", "ganancia (perdida) del ejercicio",
                        "estado de resultados", "otros rubros no clasificados",
                    }
                    pl_sub_norm = {_clean_str(x) for x in pl_subtotals}
                    idx_er = df_hoja_act[df_hoja_act["Balance clasificado"] == "Estado de Resultados"].index
                    df_sec = df_hoja_act.loc[idx_er[0] + 1:] if not idx_er.empty else df_hoja_act
                    df_clean = df_sec[df_sec["Balance clasificado"].notna() & (df_sec["Balance clasificado"].astype(str).str.strip() != "")]
                    df_clean = df_clean[~df_clean["Balance clasificado"].apply(_clean_str).isin(pl_sub_norm)]
                    pl_dict = {row["Balance clasificado"]: [row["CONSOLIDADO"]] for _, row in df_clean.iterrows()}
                    pl_df_wide = pd.DataFrame(pl_dict)
                    if pl_df_wide.empty or sum(abs(float(v[0] or 0.0)) for v in pl_dict.values() if pd.notna(v[0])) == 0.0:
                        _HYBRID_COMP_CACHE[cache_key] = (_time.time(), None)
                        return None
                else:
                    pl_df_wide = PlCuboDB.get_pl_cubo(empresa_clean, periodo_clean)
                    if pl_df_wide is None or pl_df_wide.empty:
                        tb_act_fallback = TrialBalanceDB.get_trial_balance(empresa_clean, periodo_clean)
                        if tb_act_fallback is not None and not tb_act_fallback.empty and map_pl_df is not None:
                            from src.core.sabana_builder import build_pl_sabana
                            pl_df_wide = build_pl_sabana(None, map_pl_df, tb_act_fallback)
                    if pl_df_wide is None or pl_df_wide.empty:
                        _HYBRID_COMP_CACHE[cache_key] = (_time.time(), None)
                        return None

                from src.reporting.er_generator import ERGenerator
                gen_er = ERGenerator(template_er_path)
                res_bytes_er, _ = gen_er.generate(
                    pl_df_wide, scale_factor=scale_factor, pl_df_comp=None,
                    periodo_actual_str=periodo_clean, periodo_comp_str=None, target_lang=target_lang
                )
                er_preview = wte._extract_er_preview(res_bytes_er, periodo_clean, "Comp")
                if norm_c == "#ER":
                    out_df = er_preview
                else:
                    template_ori_path = os.path.join(empresa_path, "Estado de Resultados Integrales.xlsx")
                    if not os.path.exists(template_ori_path):
                        template_ori_path = os.path.join("templates", "Estado de Resultados Integrales.xlsx")
                    if os.path.exists(template_ori_path):
                        from src.reporting.ori_generator import OriGenerator
                        bal_for_ori = SnapshotVaultManager._compute_live_statement_for_period(
                            empresa_clean, periodo_clean, "#BALANCE", scale_factor=scale_factor, target_lang=target_lang
                        )
                        gen_ori = OriGenerator(template_ori_path)
                        res_bytes_ori = gen_ori.generate(
                            pl_preview_df=er_preview,
                            periodo_actual_str=periodo_clean,
                            periodo_comp_str=None,
                            bal_preview_df=bal_for_ori,
                            empresa=empresa_clean,
                            target_lang=target_lang,
                        )
                        out_df = wte._extract_ori_preview(res_bytes_ori, periodo_clean, "Comp")

            elif norm_c == "#EFE":
                template_cf_path = os.path.join(empresa_path, "Estado de Flujos de Efectivo.xlsx")
                if not os.path.exists(template_cf_path):
                    template_cf_path = os.path.join("templates", "Estado de Flujos de Efectivo.xlsx")
                if not os.path.exists(template_cf_path):
                    return None

                try:
                    prior_dec_cf = f"{int(periodo_clean[:4]) - 1}-12"
                except Exception:
                    prior_dec_cf = None

                from src.reporting.cash_flow_generator import CashFlowGenerator
                gen_cf = CashFlowGenerator(template_cf_path)
                if is_consolidated:
                    from src.core.consolidacion_engine import generar_hoja_trabajo
                    df_hoja_prior = None
                    if prior_dec_cf:
                        df_hoja_prior, _ = generar_hoja_trabajo(grupo_id, prior_dec_cf)
                    res_bytes_cf, _ = gen_cf.generate(
                        empresa=empresa_clean,
                        periodo_actual_str=periodo_clean,
                        periodo_comp_str=prior_dec_cf,
                        map_balance_df=map_balance_df,
                        map_pl_df=map_pl_df,
                        method="Directo",
                        is_consolidado=True,
                        consolidated_hoja_trabajo_df=df_hoja_act,
                        consolidated_hoja_trabajo_comp_df=df_hoja_prior,
                        scale_factor=scale_factor,
                        target_lang=target_lang,
                    )
                else:
                    tb_check = TrialBalanceDB.get_trial_balance(empresa_clean, periodo_clean)
                    if tb_check is None or tb_check.empty:
                        _HYBRID_COMP_CACHE[cache_key] = (_time.time(), None)
                        return None
                    res_bytes_cf, _ = gen_cf.generate(
                        empresa=empresa_clean,
                        periodo_actual_str=periodo_clean,
                        periodo_comp_str=prior_dec_cf,
                        map_balance_df=map_balance_df,
                        map_pl_df=map_pl_df,
                        method="Directo",
                        is_consolidado=False,
                        scale_factor=scale_factor,
                        target_lang=target_lang,
                    )
                out_df = wte._extract_cf_preview(res_bytes_cf, periodo_clean, prior_dec_cf or "Comp")
        except Exception as e:
            print(f"Advertencia calculando comparativo al vuelo ({norm_c}) para {empresa_clean} {periodo_clean}: {e}")
            out_df = None

        _HYBRID_COMP_CACHE[cache_key] = (_time.time(), out_df.copy() if out_df is not None else None)
        return out_df

    @staticmethod
    def _build_hybrid_patrimonio(
        empresa: str,
        periodo_actual: str,
        comp_period: Optional[str],
        scale_factor: float = 1000.0,
        target_lang: str = "es",
    ) -> Optional[Tuple[pd.DataFrame, bytes]]:
        """
        Reconstruye #PATRIMONIO preservando el Bloque 1 congelado de periodo_actual
        y combinándolo con el Bloque 2 de comp_period (Snapshot cerrado o al vuelo).
        """
        import time as _time
        empresa_clean = str(empresa or "").strip()
        periodo_act_clean = str(periodo_actual or "").strip()
        comp_key = "Ninguno" if SnapshotVaultManager._is_no_comp(comp_period) else str(comp_period or "").strip()
        cache_key = ("pat_hybrid", empresa_clean, periodo_act_clean, comp_key, float(scale_factor), str(target_lang))
        cached = _HYBRID_COMP_CACHE.get(cache_key)
        if cached and (_time.time() - cached[0]) < _HYBRID_COMP_TTL_SEC:
            val = cached[1]
            return (val[0].copy(), val[1]) if val is not None else None

        is_consolidated = empresa_clean.startswith("[GRUPO]")
        clean_folder_name = empresa_clean.replace("[GRUPO] ", "").strip() if is_consolidated else empresa_clean
        empresa_path = os.path.join("data", "empresas", clean_folder_name)
        if not os.path.exists(empresa_path) and os.path.exists(os.path.join("data", "empresas", empresa_clean)):
            empresa_path = os.path.join("data", "empresas", empresa_clean)

        template_pat_path = os.path.join(empresa_path, "Estado de Cambios en el Patrimonio.xlsx")
        if not os.path.exists(template_pat_path):
            template_pat_path = os.path.join("templates", "Estado de Cambios en el Patrimonio.xlsx")
        if not os.path.exists(template_pat_path):
            return None

        try:
            is_no_comp = SnapshotVaultManager._is_no_comp(comp_period)
            if is_no_comp:
                try:
                    pat_comp_bal = f"{int(periodo_act_clean[:4]) - 1}-12"
                except Exception:
                    pat_comp_bal = "Ninguno"
            else:
                pat_comp_bal = comp_key

            bal_res = SnapshotVaultManager.get_statement(
                empresa_clean, periodo_act_clean, "#BALANCE", comp_period=pat_comp_bal, scale_factor=scale_factor
            )
            er_res = SnapshotVaultManager.get_statement(
                empresa_clean, periodo_act_clean, "#ER", comp_period=("Ninguno" if is_no_comp else comp_key), scale_factor=scale_factor
            )
            bal_df_pat = bal_res[0] if bal_res else None
            pl_df_pat = er_res[0] if er_res else None

            from src.reporting.patrimonio_generator import PatrimonioGenerator
            from src.core.excel_utils import detect_patrimonio_skiprows, clean_patrimonio_preview_df

            pat_engine = PatrimonioGenerator(template_pat_path)
            res_bytes_pat = pat_engine.generate(
                bal_preview_df=bal_df_pat,
                pl_preview_df=pl_df_pat,
                periodo_actual_str=periodo_act_clean,
                periodo_comp_str=None if is_no_comp else comp_key,
                empresa=empresa_clean,
                target_lang=target_lang,
            )
            raw_bytes = res_bytes_pat.getvalue()
            buf = BytesIO(raw_bytes)
            pat_skip = detect_patrimonio_skiprows(buf)
            buf.seek(0)
            pat_preview = pd.read_excel(buf, skiprows=pat_skip)
            pat_preview = clean_patrimonio_preview_df(pat_preview)
            _HYBRID_COMP_CACHE[cache_key] = (_time.time(), (pat_preview.copy(), raw_bytes))
            return pat_preview, raw_bytes
        except Exception as e:
            print(f"Advertencia reconstruyendo Patrimonio híbrido para {empresa_clean} {periodo_act_clean} vs {comp_key}: {e}")
            return None

    @staticmethod
    def _patch_frozen_statement_excel(
        raw_xlsx_bytes: bytes,
        norm_code: str,
        df_override: pd.DataFrame,
        periodo_actual: str,
        comp_period: Optional[str],
        target_lang: str = "es",
    ) -> bytes:
        """
        Actualiza en memoria únicamente el encabezado y la columna comparativa del Excel congelado (.xlsx),
        preservando intactos al 100% los estilos corporativos, bordes, negritas y los saldos de periodo_actual.
        """
        if not raw_xlsx_bytes or df_override is None or df_override.empty:
            return raw_xlsx_bytes

        try:
            import openpyxl
            import re as _re
            import unicodedata as _ud
            from src.reporting.er_generator import format_period_date

            wb = openpyxl.load_workbook(BytesIO(raw_xlsx_bytes))
            ws = wb.active

            def _norm_lbl(v: Any) -> str:
                if v is None or pd.isna(v):
                    return ""
                s = str(v).strip().lower()
                s = "".join(ch for ch in _ud.normalize("NFD", s) if _ud.category(ch) != "Mn")
                return _re.sub(r"\s+", " ", s)

            df_labels_set = {_norm_lbl(v) for v in df_override.iloc[:, 0] if _norm_lbl(v)}

            # 1. Detectar columna de etiquetas y columnas de fecha/valores
            name_col_idx = 1
            best_matches = -1
            for c_cand in range(1, min(5, ws.max_column + 1)):
                m_cnt = sum(
                    1 for r_i in range(1, ws.max_row + 1)
                    if _norm_lbl(ws.cell(row=r_i, column=c_cand).value) in df_labels_set
                )
                if m_cnt > best_matches:
                    best_matches = m_cnt
                    name_col_idx = c_cand

            first_data_row = ws.max_row + 1
            for r_i in range(1, ws.max_row + 1):
                if _norm_lbl(ws.cell(row=r_i, column=name_col_idx).value) in df_labels_set:
                    first_data_row = r_i
                    break
            header_max_row = max(1, first_data_row - 1)

            if norm_code in ["#BALANCE", "#ESF"]:
                from src.core.excel_utils import detect_balance_columns, read_template_config
                cfg = read_template_config(wb)
                if cfg:
                    act_col_idx = int(cfg.get("val_actual_col", 4))
                    comp_col_idx = int(cfg.get("val_comp_col", 5))
                else:
                    _, _, act_col_idx, comp_col_idx = detect_balance_columns(ws, wb)
            else:
                nota_col_idx = None
                for c_i in range(name_col_idx + 1, ws.max_column + 1):
                    for r_i in range(1, header_max_row + 1):
                        v = ws.cell(row=r_i, column=c_i).value
                        if v is not None and str(v).strip().lower() in ["nota", "notas", "note", "notes"]:
                            nota_col_idx = c_i
                            break
                    if nota_col_idx is not None:
                        break

                cand_val_cols = []
                for c_i in range(name_col_idx + 1, ws.max_column + 1):
                    if c_i == nota_col_idx:
                        continue
                    has_hdr_or_num = False
                    for r_i in range(1, ws.max_row + 1):
                        v = ws.cell(row=r_i, column=c_i).value
                        if v is not None and str(v).strip() != "":
                            has_hdr_or_num = True
                            break
                    if has_hdr_or_num:
                        cand_val_cols.append(c_i)

                if len(cand_val_cols) < 2:
                    return raw_xlsx_bytes
                act_col_idx, comp_col_idx = cand_val_cols[0], cand_val_cols[1]

            # 2. Actualizar encabezado de fecha comparativa (y corregir fecha actual si era datetime de plantilla)
            is_no_comp = SnapshotVaultManager._is_no_comp(comp_period)
            fmt_act = format_period_date(periodo_actual, lang=target_lang) or str(periodo_actual)
            fmt_comp = "" if is_no_comp else (format_period_date(str(comp_period).strip(), lang=target_lang) or str(comp_period).strip())

            updated_comp_hdr = False
            for r_h in range(1, header_max_row + 1):
                c_act = ws.cell(row=r_h, column=act_col_idx)
                c_cmp = ws.cell(row=r_h, column=comp_col_idx)
                v_act = c_act.value
                v_cmp = c_cmp.value

                if isinstance(v_act, (datetime.datetime, datetime.date)):
                    c_act.value = fmt_act

                is_date_hdr = (
                    isinstance(v_cmp, (datetime.datetime, datetime.date))
                    or (isinstance(v_cmp, str) and (_re.search(r"20\d{2}", v_cmp) or any(w in v_cmp.lower() for w in ["anterior", "comparat", "comp"])))
                    or (isinstance(v_cmp, (int, float)) and 2000 <= v_cmp <= 2100)
                )
                if is_date_hdr:
                    c_cmp.value = fmt_comp
                    updated_comp_hdr = True

            if not updated_comp_hdr:
                for r_h in range(1, header_max_row + 1):
                    v_act = ws.cell(row=r_h, column=act_col_idx).value
                    if v_act is not None and _re.search(r"20\d{2}", str(v_act)):
                        ws.cell(row=r_h, column=comp_col_idx).value = fmt_comp
                        break

            # 3. Extraer valores de la columna comparativa de df_override por (etiqueta_norm, ocurrencia)
            df_date_cols = [
                c for c in df_override.columns[1:]
                if _re.search(r"20\d{2}", str(c)) or str(c).strip().lower() not in ["nota", "notas", "note", "notes"]
            ]
            if len(df_date_cols) < 2:
                return raw_xlsx_bytes
            df_comp_col = df_date_cols[1]

            occ_src: Dict[str, int] = {}
            val_by_occ: Dict[Tuple[str, int], Any] = {}
            val_first: Dict[str, Any] = {}
            for _, row_df in df_override.iterrows():
                k = _norm_lbl(row_df.iloc[0])
                if k:
                    occ_i = occ_src.get(k, 0)
                    occ_src[k] = occ_i + 1
                    val_by_occ[(k, occ_i)] = row_df[df_comp_col]
                    if k not in val_first:
                        val_first[k] = row_df[df_comp_col]

            occ_ws: Dict[str, int] = {}
            for r_w in range(first_data_row, ws.max_row + 1):
                lbl_w = _norm_lbl(ws.cell(row=r_w, column=name_col_idx).value)
                if not lbl_w:
                    continue
                occ_i = occ_ws.get(lbl_w, 0)
                occ_ws[lbl_w] = occ_i + 1
                matched = False
                raw_v = None
                if (lbl_w, occ_i) in val_by_occ:
                    raw_v = val_by_occ[(lbl_w, occ_i)]
                    matched = True
                elif lbl_w in val_first:
                    raw_v = val_first[lbl_w]
                    matched = True

                if matched:
                    cell_cmp = ws.cell(row=r_w, column=comp_col_idx)
                    if is_no_comp or raw_v is None or pd.isna(raw_v) or str(raw_v).strip() == "":
                        cell_cmp.value = None
                    else:
                        try:
                            fv = float(raw_v)
                            cell_cmp.value = int(round(fv)) if abs(fv - round(fv)) < 1e-6 else fv
                        except Exception:
                            cell_cmp.value = raw_v

            out_buf = BytesIO()
            wb.save(out_buf)
            return out_buf.getvalue()
        except Exception as e:
            print(f"Advertencia parcheando Excel congelado {norm_code}: {e}")
            return raw_xlsx_bytes

    @staticmethod
    def _build_hybrid_note_for_comp(
        empresa: str,
        periodo_actual: str,
        note_code: str,
        comp_period: Optional[str],
        scale_factor: float = 1000.0,
        target_lang: str = "es",
    ) -> Optional[Tuple[List[Tuple[str, Any]], bytes]]:
        """
        Genera una Nota explicativa híbrida cuando periodo_actual está Cerrado pero se solicita un
        comp_period distinto al nativo del Snapshot (ya sea Cerrado en Bóveda o Abierto al vuelo).
        - periodo_actual usa su contexto/mapeo congelado e inmutable de Bóveda.
        - comp_period usa su Snapshot congelado (si está cerrado) o se calcula al vuelo (si está abierto).
        - Reconcilia cuadraturas contra los EEFF híbridos y preserva exactitud de celdas actuales congeladas.
        """
        import time as _time
        empresa_clean = str(empresa or "").strip()
        periodo_act_clean = str(periodo_actual or "").strip()
        norm_note_code = "#" + str(note_code or "").replace("#", "").strip().upper()
        comp_clean = "Ninguno" if SnapshotVaultManager._is_no_comp(comp_period) else str(comp_period or "").strip()

        cache_key = ("note_hybrid", empresa_clean, periodo_act_clean, norm_note_code, comp_clean, float(scale_factor), str(target_lang))
        cached = _HYBRID_COMP_CACHE.get(cache_key)
        if cached and (_time.time() - cached[0]) < _HYBRID_COMP_TTL_SEC:
            c_els, c_bytes = cached[1]
            copied_els = [
                (it[0], it[1].copy() if isinstance(it[1], pd.DataFrame) else it[1], *it[2:])
                for it in c_els
            ]
            return copied_els, c_bytes

        template_nota = "Plantilla de notas_v1.xlsx"
        if not os.path.exists(template_nota):
            return None

        full_reg = get_full_note_registry()
        info = full_reg.get(norm_note_code)
        if not info or not info.get("sheets"):
            return None
        target_sheets = info["sheets"]
        is_consolidated = empresa_clean.startswith("[GRUPO]")

        try:
            from src.core.word_template_engine import load_mappings_for_entity
            from src.core.sabana_manager import SabanaManager
            from src.reporting.note_generator import NoteGenerator
            from src.core.external_notes_manager import ExternalNotesManager
            from src.core.rounding_reconciler import RoundingReconciler
            from src.core.excel_utils import extract_named_ranges_from_excel
            from src.ui_pages.informes_y_notas import split_sheet_into_elements
            from src.reporting.formatting import ensure_subtable_separator_rows

            map_balance_df, map_pl_df = load_mappings_for_entity(empresa_clean)
            bundle = SabanaManager.build_sabana_bundle(
                active_entity=empresa_clean,
                periodo_actual=periodo_act_clean,
                periodo_comp=comp_clean,
                map_balance_df=map_balance_df,
                map_pl_df=map_pl_df,
            )
            entity_contexts = bundle["contexts"]

            engine_notes = NoteGenerator(template_nota)
            wb_out = engine_notes.generate(
                sheet_names=target_sheets,
                entity_contexts=entity_contexts,
                active_entity_name=empresa_clean,
                is_consolidated=is_consolidated,
                scale_factor=scale_factor,
                periodo_actual_str=periodo_act_clean,
                periodo_comp_str=comp_clean,
                map_balance_df=map_balance_df,
                map_pl_df=map_pl_df,
                target_lang=target_lang,
                evaluate_formulas=False,
            )
            wb_out = ExternalNotesManager.inject_external_data(
                wb_out,
                empresa=empresa_clean,
                periodo_actual=periodo_act_clean,
                periodo_comp=comp_clean if comp_clean != "Ninguno" else None,
                target_lang=target_lang,
            )

            bal_stmt = SnapshotVaultManager.get_statement(
                empresa_clean, periodo_act_clean, "#BALANCE", comp_period=comp_clean, scale_factor=scale_factor
            )
            er_stmt = SnapshotVaultManager.get_statement(
                empresa_clean, periodo_act_clean, "#ER", comp_period=comp_clean, scale_factor=scale_factor
            )
            df_bal_hyb = bal_stmt[0] if bal_stmt else None
            df_er_hyb = er_stmt[0] if er_stmt else None

            _, wb_eval_rec = RoundingReconciler.reconcile_and_evaluate_note_workbook(
                wb_out,
                empresa=empresa_clean,
                periodo_actual=periodo_act_clean,
                scale_factor=scale_factor,
                df_bal=df_bal_hyb,
                df_er=df_er_hyb,
                target_sheets=target_sheets,
                return_raw=False,
            )
            wb_bytes = wb_eval_rec.getvalue()

            import openpyxl
            wb_openpyxl_eval = openpyxl.load_workbook(BytesIO(wb_bytes), data_only=True)
            named_ranges = extract_named_ranges_from_excel(None, norm_note_code, wb=wb_openpyxl_eval)
            raw_els: List[Tuple] = []
            if named_ranges:
                for _, df_chunk in named_ranges:
                    raw_els.append(("table", ensure_subtable_separator_rows(df_chunk), target_sheets[0], None))
            else:
                xl_nota_f = pd.ExcelFile(BytesIO(wb_bytes))
                try:
                    for s_name in target_sheets:
                        if s_name in xl_nota_f.sheet_names:
                            preview_nota_df = xl_nota_f.parse(sheet_name=s_name, header=None)
                            sheet_els = split_sheet_into_elements(preview_nota_df, sheet_name=s_name)
                            for s_item in sheet_els:
                                el_t = s_item[0]
                                el_v = ensure_subtable_separator_rows(s_item[1]) if (el_t == "table" and isinstance(s_item[1], pd.DataFrame)) else s_item[1]
                                sh_n = s_item[2] if len(s_item) > 2 else s_name
                                exp_c = s_item[3] if len(s_item) > 3 else None
                                raw_els.append((el_t, el_v, sh_n, exp_c))
                finally:
                    xl_nota_f.close()
            wb_openpyxl_eval.close()

            # Garantizar inmutabilidad exacta de la columna Actual desde el Snapshot congelado de periodo_actual
            # en tablas estándar de 3 columnas (y de comp_period si comp_period también está cerrado)
            frozen_act_els = SnapshotVaultManager.get_note_elements(
                empresa_clean, periodo_act_clean, norm_note_code, scale_factor=scale_factor, comp_period=None
            )
            if frozen_act_els:
                import re as _re
                frozen_act_tables = [el[1] for el in frozen_act_els if el[0] == "table" and isinstance(el[1], pd.DataFrame)]
                tbl_i = 0
                for idx_el, item in enumerate(raw_els):
                    if item[0] == "table" and isinstance(item[1], pd.DataFrame):
                        df_h = item[1].copy()
                        if tbl_i < len(frozen_act_tables):
                            df_fa = frozen_act_tables[tbl_i]
                            if len(df_h.columns) == 3 and len(df_fa.columns) == 3 and len(df_h) == len(df_fa):
                                has_dates = any(
                                    _re.search(r"20\d{2}", str(df_fa.iloc[r_k, 1] or ""))
                                    for r_k in range(min(3, len(df_fa)))
                                )
                                if has_dates:
                                    df_h.iloc[:, 1] = df_fa.iloc[:, 1].values
                                    raw_els[idx_el] = (item[0], df_h, *item[2:])
                        tbl_i += 1

            if comp_clean != "Ninguno" and SnapshotVaultManager.has_active_snapshot(empresa_clean, comp_clean):
                raw_els = SnapshotVaultManager.apply_closed_comp_to_note_elements(
                    raw_els, empresa_clean, norm_note_code, periodo_act_clean, comp_clean, scale_factor=scale_factor
                )

            _HYBRID_COMP_CACHE[cache_key] = (_time.time(), (raw_els, wb_bytes))
            return [
                (it[0], it[1].copy() if isinstance(it[1], pd.DataFrame) else it[1], *it[2:])
                for it in raw_els
            ], wb_bytes
        except Exception as e:
            print(f"Advertencia generando nota híbrida {norm_note_code} ({empresa_clean} {periodo_act_clean} vs {comp_clean}): {e}")
            return None

    @staticmethod
    def get_statement_excel(
        empresa: str,
        periodo_actual: str,
        code: str,
        comp_period: Optional[str] = None,
        scale_factor: float = 1000.0,
        df_override: Optional[pd.DataFrame] = None,
        target_lang: str = "es",
    ) -> Optional[bytes]:
        """
        Recupera el archivo Excel (.xlsx) congelado de un Estado Financiero desde la Bóveda.
        Si se especifica un `comp_period` distinto al nativo del Snapshot (o cerrado con posterioridad),
        parchea únicamente la columna comparativa en memoria preservando intactos los estilos y saldos actuales.
        """
        if not SnapshotVaultManager.has_active_snapshot(empresa, periodo_actual):
            return None
        snap_dir = SnapshotVaultManager.get_snapshot_dir(empresa, periodo_actual)
        stmt_dir = os.path.join(snap_dir, "statements")
        norm_code = SnapshotVaultManager._norm_stmt_code(code)
        filename_map = {
            "#BALANCE": "balance.xlsx",
            "#ESF": "balance.xlsx",
            "#ER": "er.xlsx",
            "#ORI": "ori.xlsx",
            "#PATRIMONIO": "patrimonio.xlsx",
            "#EFE": "efe.xlsx"
        }
        fname = filename_map.get(norm_code)
        if not fname:
            return None
        file_path = os.path.join(stmt_dir, fname)
        raw_bytes = None
        if os.path.exists(file_path):
            try:
                with open(file_path, "rb") as f:
                    raw_bytes = f.read()
            except Exception:
                raw_bytes = None

        if comp_period is None or abs(float(scale_factor) - 1000.0) >= 0.01:
            return raw_bytes

        meta = SnapshotVaultManager._get_snapshot_meta(empresa, periodo_actual)
        native_comp = SnapshotVaultManager._get_stmt_native_comp(meta, norm_code)
        is_no_comp = SnapshotVaultManager._is_no_comp(comp_period)
        comp_str = str(comp_period).strip()
        comp_differs = is_no_comp or (not native_comp) or (comp_str != str(native_comp).strip())
        comp_is_closed = (not is_no_comp) and SnapshotVaultManager.has_active_snapshot(empresa, comp_str)

        if comp_differs or comp_is_closed:
            if norm_code == "#PATRIMONIO":
                pat_hyb = SnapshotVaultManager._build_hybrid_patrimonio(
                    empresa, periodo_actual, comp_period, scale_factor=scale_factor, target_lang=target_lang
                )
                if pat_hyb is not None:
                    return pat_hyb[1]
            elif raw_bytes is not None:
                df_use = df_override
                if df_use is None:
                    stmt_res = SnapshotVaultManager.get_statement(
                        empresa, periodo_actual, norm_code, comp_period=comp_period, scale_factor=scale_factor
                    )
                    df_use = stmt_res[0] if stmt_res else None
                if df_use is not None:
                    return SnapshotVaultManager._patch_frozen_statement_excel(
                        raw_bytes, norm_code, df_use, periodo_actual, comp_period, target_lang=target_lang
                    )

        return raw_bytes

    @staticmethod
    def get_note_excel(
        empresa: str,
        periodo_actual: str,
        note_code: str,
        comp_period: Optional[str] = None,
        scale_factor: float = 1000.0,
        target_lang: str = "es",
    ) -> Optional[bytes]:
        """
        Recupera el archivo Excel (.xlsx) evaluado de una Nota explicativa desde la Bóveda,
        o construye el libro híbrido si `comp_period` difiere del comparativo nativo congelado.
        """
        if not SnapshotVaultManager.has_active_snapshot(empresa, periodo_actual):
            return None

        if comp_period is not None:
            meta = SnapshotVaultManager._get_snapshot_meta(empresa, periodo_actual)
            note_native_comp = SnapshotVaultManager._get_note_native_comp(meta, note_code)
            is_no_comp = SnapshotVaultManager._is_no_comp(comp_period)
            comp_str = str(comp_period).strip()
            comp_differs = is_no_comp or (not note_native_comp) or (comp_str != str(note_native_comp).strip())
            if comp_differs or abs(float(scale_factor) - 1000.0) >= 0.01:
                hyb = SnapshotVaultManager._build_hybrid_note_for_comp(
                    empresa, periodo_actual, note_code, comp_period, scale_factor=scale_factor, target_lang=target_lang
                )
                if hyb is not None:
                    return hyb[1]

        snap_dir = SnapshotVaultManager.get_snapshot_dir(empresa, periodo_actual)
        clean_code = note_code.replace("#", "").strip()
        file_path = os.path.join(snap_dir, "notes", f"{clean_code}.xlsx")
        if os.path.exists(file_path):
            try:
                with open(file_path, "rb") as f:
                    return f.read()
            except Exception:
                pass
        return None

    @staticmethod
    def apply_closed_comp_to_statement(
        df_live: pd.DataFrame,
        empresa: str,
        code: str,
        periodo_actual: str,
        periodo_comp: Optional[str],
        scale_factor: float = 1000.0
    ) -> pd.DataFrame:
        """
        Cuando el período comparativo está cerrado en Bóveda, sobreescribe la columna comparativa
        del Estado Financiero con los valores congelados del Snapshot de periodo_comp.
        """
        if df_live is None or df_live.empty or not periodo_comp or SnapshotVaultManager._is_no_comp(periodo_comp):
            return df_live
        if not SnapshotVaultManager.has_active_snapshot(empresa, periodo_comp):
            return df_live

        norm_code = SnapshotVaultManager._norm_stmt_code(code)
        if norm_code not in ["#BALANCE", "#ESF", "#ER", "#ORI", "#EFE"]:
            return df_live

        snap_comp_res = SnapshotVaultManager.get_statement(
            empresa=empresa,
            periodo_actual=periodo_comp,
            code=norm_code,
            comp_period=None,
            scale_factor=scale_factor
        )
        if not snap_comp_res or snap_comp_res[0] is None:
            return df_live

        return SnapshotVaultManager._splice_comp_into_statement_df(
            df_actual_frozen=df_live,
            df_comp_source=snap_comp_res[0],
            comp_period=periodo_comp,
        )

    @staticmethod
    def apply_closed_comp_to_note_elements(
        elements_live: List[Tuple[str, Any]],
        empresa: str,
        note_code: str,
        periodo_actual: str,
        periodo_comp: Optional[str],
        scale_factor: float = 1000.0
    ) -> List[Tuple[str, Any]]:
        """
        Cuando una Nota fue recuperada de un Snapshot cuyo período comparativo nativo difiere del
        periodo_comp solicitado (y periodo_comp tiene su propio Snapshot activo), actualiza las
        columnas comparativas con los valores congelados del Snapshot de periodo_comp.
        """
        if not elements_live or not periodo_comp or str(periodo_comp).strip().lower() in ["ninguno", "none", ""]:
            return elements_live
        if not SnapshotVaultManager.has_active_snapshot(empresa, periodo_comp):
            return elements_live

        snap_comp_els = SnapshotVaultManager.get_note_elements(
            empresa=empresa,
            periodo_actual=periodo_comp,
            note_code=note_code,
            scale_factor=scale_factor
        )
        if not snap_comp_els:
            return elements_live

        from src.reporting.note_generator import format_date_dots
        comp_dot = format_date_dots(periodo_comp)
        comp_tables = [el[1] for el in snap_comp_els if el[0] == "table" and isinstance(el[1], pd.DataFrame)]

        updated_els = []
        tbl_idx = 0
        import re as _re
        for item in elements_live:
            el_type = item[0]
            el_val = item[1]
            if el_type == "table" and isinstance(el_val, pd.DataFrame):
                df_cur = el_val.copy()
                if tbl_idx < len(comp_tables):
                    df_cmp = comp_tables[tbl_idx]
                    # Tablas estándar de 3 columnas (Concepto, Fecha Actual, Fecha Comparativa) con igual número de filas
                    if len(df_cur.columns) == 3 and len(df_cmp.columns) == 3 and len(df_cur) == len(df_cmp):
                        has_two_dates = False
                        for r_i in range(min(3, len(df_cur))):
                            v1 = str(df_cur.iloc[r_i, 1] or "")
                            v2 = str(df_cur.iloc[r_i, 2] or "")
                            if _re.search(r'20\d{2}', v1) and _re.search(r'20\d{2}', v2):
                                has_two_dates = True
                                break
                        if has_two_dates:
                            for r_i in range(len(df_cur)):
                                v_cmp_act = df_cmp.iloc[r_i, 1]
                                if isinstance(v_cmp_act, str) and _re.search(r'\d{2}\.\d{2}\.20\d{2}', v_cmp_act):
                                    df_cur.iloc[r_i, 2] = comp_dot
                                else:
                                    df_cur.iloc[r_i, 2] = v_cmp_act
                tbl_idx += 1
                if len(item) == 4:
                    updated_els.append((el_type, df_cur, item[2], item[3]))
                elif len(item) == 3:
                    updated_els.append((el_type, df_cur, item[2]))
                else:
                    updated_els.append((el_type, df_cur))
            else:
                updated_els.append(item)
        return updated_els

    @staticmethod
    def get_statement(empresa: str, periodo_actual: str, code: str, comp_period: Optional[str] = None, scale_factor: float = 1000.0) -> Optional[Tuple[pd.DataFrame, str]]:
        """
        Recupera un Estado Financiero congelado desde la Bóveda en milisegundos,
        escalando dinámicamente sus valores numéricos según base_scale del snapshot.
        """
        if not SnapshotVaultManager.has_active_snapshot(empresa, periodo_actual):
            return None

        snap_dir = SnapshotVaultManager.get_snapshot_dir(empresa, periodo_actual)
        stmt_dir = os.path.join(snap_dir, "statements")
        meta = SnapshotVaultManager._get_snapshot_meta(empresa, periodo_actual)
        base_scale = float(meta.get("base_scale", 1.0))
        
        norm_code = code.upper().strip()
        if not norm_code.startswith("#"):
            norm_code = "#" + norm_code
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
                if norm_code == "#PATRIMONIO":
                    from src.core.excel_utils import clean_patrimonio_preview_df
                    df = clean_patrimonio_preview_df(df)
                # Escalar los valores numéricos desde base_scale a scale_factor
                df_scaled = SnapshotVaultManager._scale_dataframe_cells(df, scale_factor, base_scale=base_scale)
                
                # Resolver columna comparativa cuando comp_period fue especificado explícitamente:
                # 1) comp_period == 'Ninguno' -> vaciar columna comparativa
                # 2) comp_period Cerrado (Snapshot activo) -> inyectar desde su propio Snapshot certificado
                # 3) comp_period Abierto y distinto al nativo congelado -> calcular al vuelo e inyectar
                if comp_period is not None:
                    native_comp = SnapshotVaultManager._get_stmt_native_comp(meta, norm_code)
                    is_no_comp = SnapshotVaultManager._is_no_comp(comp_period)
                    comp_str = str(comp_period).strip()
                    comp_differs = is_no_comp or (not native_comp) or (comp_str != str(native_comp).strip())
                    comp_is_closed = (not is_no_comp) and SnapshotVaultManager.has_active_snapshot(empresa, comp_str)

                    if norm_code == "#PATRIMONIO":
                        if comp_differs or comp_is_closed:
                            pat_hyb = SnapshotVaultManager._build_hybrid_patrimonio(
                                empresa, periodo_actual, comp_period, scale_factor=scale_factor
                            )
                            if pat_hyb is not None:
                                df_scaled = pat_hyb[0]
                    elif norm_code in ["#BALANCE", "#ESF", "#ER", "#ORI", "#EFE"]:
                        if is_no_comp:
                            df_scaled = SnapshotVaultManager._splice_comp_into_statement_df(
                                df_scaled, None, "Ninguno"
                            )
                        elif comp_is_closed:
                            df_scaled = SnapshotVaultManager.apply_closed_comp_to_statement(
                                df_scaled, empresa, norm_code, periodo_actual, comp_str, scale_factor=scale_factor
                            )
                        elif comp_differs:
                            df_comp_live = SnapshotVaultManager._compute_live_statement_for_period(
                                empresa, comp_str, norm_code, scale_factor=scale_factor
                            )
                            df_scaled = SnapshotVaultManager._splice_comp_into_statement_df(
                                df_scaled, df_comp_live, comp_str
                            )
                return df_scaled, rep_name
        except Exception as e:
            print(f"Error cargando statement {code} desde snapshot: {e}")
            return None

        return None

    @staticmethod
    def get_note_elements(empresa: str, periodo_actual: str, note_code: str, scale_factor: float = 1000.0, comp_period: Optional[str] = None) -> Optional[List[Tuple[str, Any]]]:
        """
        Recupera las tablas y textos de una Nota explicativa congelada desde la Bóveda en milisegundos,
        escalando sus valores numéricos según base_scale del snapshot.
        """
        if not SnapshotVaultManager.has_active_snapshot(empresa, periodo_actual):
            return None

        snap_dir = SnapshotVaultManager.get_snapshot_dir(empresa, periodo_actual)
        meta = SnapshotVaultManager._get_snapshot_meta(empresa, periodo_actual)
        base_scale = float(meta.get("base_scale", 1.0))

        clean_code = note_code.replace("#", "").strip()
        file_path = os.path.join(snap_dir, "notes", f"{clean_code}.pkl")
        
        if not os.path.exists(file_path):
            return None

        try:
            with open(file_path, "rb") as f:
                raw_els = pickle.load(f)
                
            from src.reporting.formatting import ensure_subtable_separator_rows
            scaled_els = []
            for item in raw_els:
                el_type = item[0]
                el_val = item[1]
                if el_type == "table" and isinstance(el_val, pd.DataFrame):
                    df_sc = ensure_subtable_separator_rows(SnapshotVaultManager._scale_dataframe_cells(el_val, scale_factor, base_scale=base_scale))
                    if len(item) == 4:
                        scaled_els.append((el_type, df_sc, item[2], item[3]))
                    elif len(item) == 3:
                        scaled_els.append((el_type, df_sc, item[2]))
                    else:
                        scaled_els.append((el_type, df_sc))
                else:
                    scaled_els.append(item)

            if comp_period is not None:
                note_native_comp = SnapshotVaultManager._get_note_native_comp(meta, note_code)
                is_no_comp = SnapshotVaultManager._is_no_comp(comp_period)
                comp_str = str(comp_period).strip()
                comp_differs = is_no_comp or (not note_native_comp) or (comp_str != str(note_native_comp).strip())
                comp_is_closed = (not is_no_comp) and SnapshotVaultManager.has_active_snapshot(empresa, comp_str)

                if comp_differs:
                    hyb_note = SnapshotVaultManager._build_hybrid_note_for_comp(
                        empresa, periodo_actual, note_code, comp_period, scale_factor=scale_factor
                    )
                    if hyb_note is not None and hyb_note[0]:
                        return hyb_note[0]
                elif comp_is_closed:
                    scaled_els = SnapshotVaultManager.apply_closed_comp_to_note_elements(
                        scaled_els, empresa, note_code, periodo_actual, comp_str, scale_factor=scale_factor
                    )
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
            _HYBRID_COMP_CACHE.clear()
            return True, f"Snapshot de {periodo_clean} archivado con éxito. El período ha vuelto a memoria activa viva."
        except Exception as e:
            db.rollback()
            return False, f"Error al actualizar estado de snapshot: {e}"
        finally:
            db.close()

    @staticmethod
    def sync_all_existing_historical_periods(progress_callback=None, force_rebuild: bool = False) -> Tuple[int, int, List[str]]:
        """
        Escanea todos los períodos que ya se encuentran cerrados en la base de datos
        (tanto individuales en HistoricalDetailRecord como grupos consolidados con candado)
        y genera automáticamente sus Snapshots en la Bóveda para que queden 100% operativos.
        Si force_rebuild es True, regenera también aquellos que ya tenían snapshot activo.
        """
        db = SessionLocal()
        try:
            if force_rebuild:
                active_snaps = db.query(
                    ClosingSnapshotRecord.empresa,
                    ClosingSnapshotRecord.periodo
                ).filter_by(is_active=True).distinct().all()
                all_targets = [(emp, per) for emp, per in active_snaps if emp and per]
            else:
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
            
            # Si ya tiene snapshot activo y no se solicitó reconstrucción forzada, omitir
            if (not force_rebuild) and SnapshotVaultManager.has_active_snapshot(emp, per):
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
