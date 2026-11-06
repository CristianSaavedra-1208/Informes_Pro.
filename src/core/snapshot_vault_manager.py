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

        try:
            if is_consolidated:
                sab_bal_df = build_consolidated_balance_sabana(empresa_clean, periodo_clean, map_balance_df)
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
                sab_pl_df = build_consolidated_pl_sabana(empresa_clean, periodo_clean, map_pl_df)
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
                df_hoja_act, _ = generar_hoja_trabajo(grupo_id, periodo_clean)
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
                    wb_out = evaluate_formulas_in_workbook(wb_out)
                    wb_bytes = wb_out.getvalue()
                    evaluated_wb_cache[n_comp] = (wb_bytes, pd.ExcelFile(BytesIO(wb_bytes)))
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
                    wb_bytes, xl_nota_f = evaluated_wb_cache[note_comp]
                    excel_nota_out = BytesIO(wb_bytes)

                    named_ranges = extract_named_ranges_from_excel(excel_nota_out, code)
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
            return True, f"Snapshot congelado con éxito en Bóveda ({len(statements_saved)} estados financieros y {len(notes_saved)} notas explicativas)."
        except Exception as e_db:
            db.rollback()
            return False, f"Error al registrar snapshot en base de datos: {e_db}"
        finally:
            db.close()

    @staticmethod
    def get_statement_excel(empresa: str, periodo_actual: str, code: str) -> Optional[bytes]:
        """Recupera el archivo Excel (.xlsx) congelado de un Estado Financiero si existe en la Bóveda."""
        if not SnapshotVaultManager.has_active_snapshot(empresa, periodo_actual):
            return None
        snap_dir = SnapshotVaultManager.get_snapshot_dir(empresa, periodo_actual)
        stmt_dir = os.path.join(snap_dir, "statements")
        norm_code = code.upper().strip()
        if not norm_code.startswith("#"):
            norm_code = "#" + norm_code
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
        if os.path.exists(file_path):
            try:
                with open(file_path, "rb") as f:
                    return f.read()
            except Exception:
                pass
        return None

    @staticmethod
    def get_note_excel(empresa: str, periodo_actual: str, note_code: str) -> Optional[bytes]:
        """Recupera el archivo Excel (.xlsx) evaluado y congelado de una Nota explicativa."""
        if not SnapshotVaultManager.has_active_snapshot(empresa, periodo_actual):
            return None
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
        if df_live is None or df_live.empty or not periodo_comp or str(periodo_comp).strip().lower() in ["ninguno", "none", ""]:
            return df_live
        if not SnapshotVaultManager.has_active_snapshot(empresa, periodo_comp):
            return df_live

        norm_code = code.upper().strip()
        if not norm_code.startswith("#"):
            norm_code = "#" + norm_code
        if norm_code not in ["#BALANCE", "#ESF", "#ER", "#ORI"]:
            return df_live

        snap_comp_res = SnapshotVaultManager.get_statement(
            empresa=empresa,
            periodo_actual=periodo_comp,
            code=norm_code,
            scale_factor=scale_factor
        )
        if not snap_comp_res or snap_comp_res[0] is None:
            return df_live

        df_snap_comp = snap_comp_res[0]
        if len(df_live.columns) < 2 or len(df_snap_comp.columns) < 2:
            return df_live

        out_df = df_live.copy()
        # Identificar la columna Actual de df_snap_comp y la columna Comparativa de out_df
        import re as _re
        snap_date_cols = [
            c for c in df_snap_comp.columns[1:]
            if _re.search(r'20\d{2}', str(c)) or str(c).strip().lower() not in ["nota", "notas"]
        ]
        live_date_cols = [
            c for c in out_df.columns[1:]
            if _re.search(r'20\d{2}', str(c)) or str(c).strip().lower() not in ["nota", "notas"]
        ]
        if not snap_date_cols or len(live_date_cols) < 2:
            return out_df

        snap_act_col = snap_date_cols[0]
        live_comp_col = live_date_cols[1]

        if len(out_df) == len(df_snap_comp):
            out_df[live_comp_col] = df_snap_comp[snap_act_col].values
        else:
            # Emparejar por etiqueta de fila normalizada
            def _norm_lbl(v):
                return "" if pd.isna(v) else str(v).strip().lower()
            snap_map = {}
            for _, r_s in df_snap_comp.iterrows():
                k = _norm_lbl(r_s.iloc[0])
                if k and k not in snap_map:
                    snap_map[k] = r_s[snap_act_col]
            for idx_l, r_l in out_df.iterrows():
                k_l = _norm_lbl(r_l.iloc[0])
                if k_l in snap_map:
                    out_df.at[idx_l, live_comp_col] = snap_map[k_l]

        # Si el nombre de la columna comparativa era distinto, renombrar a periodo_comp si procede
        if str(live_comp_col) != str(periodo_comp) and _re.match(r'^20\d{2}-\d{2}$', str(live_comp_col)):
            out_df = out_df.rename(columns={live_comp_col: str(periodo_comp)})

        return out_df

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
                
                # Si se especificó un comp_period distinto al nativo del snapshot y comp_period está cerrado, inyectarlo
                if comp_period and str(comp_period).strip().lower() not in ["ninguno", "none", ""]:
                    native_comp = meta.get("comp_period_bal") if norm_code in ["#BALANCE", "#ESF", "#PATRIMONIO", "#EFE"] else meta.get("comp_period_pl")
                    if native_comp and str(comp_period).strip() != str(native_comp).strip():
                        if SnapshotVaultManager.has_active_snapshot(empresa, comp_period):
                            df_scaled = SnapshotVaultManager.apply_closed_comp_to_statement(
                                df_scaled, empresa, norm_code, periodo_actual, comp_period, scale_factor=scale_factor
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

            if comp_period and str(comp_period).strip().lower() not in ["ninguno", "none", ""]:
                native_bal = meta.get("comp_period_bal")
                native_pl = meta.get("comp_period_pl")
                if str(comp_period).strip() not in [str(native_bal).strip(), str(native_pl).strip()]:
                    if SnapshotVaultManager.has_active_snapshot(empresa, comp_period):
                        scaled_els = SnapshotVaultManager.apply_closed_comp_to_note_elements(
                            scaled_els, empresa, note_code, periodo_actual, comp_period, scale_factor=scale_factor
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
