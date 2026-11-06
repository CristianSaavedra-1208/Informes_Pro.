import os
import pandas as pd
import streamlit as st

from src.core.sabana_builder import build_balance_sabana, build_pl_sabana
from src.models.trial_balance_db import TrialBalanceDB
from src.models.pl_cubo_db import PlCuboDB
from src.models.database import SessionLocal
from src.models.consolidacion import ConsolidationGroup, ConsolidationJournalEntry
from src.reporting.note_generator import build_entity_context

class SabanaManager:
    """
    Gestor centralizado que sirve la Sábana Maestro de Auditoría (Balance y P&L)
    como Fuente Única de Verdad en memoria RAM para reportes y notas.
    """

    @staticmethod
    def get_sabana_bundle(active_entity: str, periodo_actual: str, periodo_comp: str = "Ninguno", map_balance_df: pd.DataFrame = None, map_pl_df: pd.DataFrame = None):
        """
        Retorna la Sábana Maestro compilada (DataFrames y Contextos) usando caché de session_state.
        """
        cache_key = f"_sabana_maestro_bundle__{active_entity}__{periodo_actual}__{periodo_comp}"
        
        if hasattr(st, 'session_state') and cache_key in st.session_state:
            return st.session_state[cache_key]
            
        bundle = SabanaManager.build_sabana_bundle(active_entity, periodo_actual, periodo_comp, map_balance_df, map_pl_df)
        
        if hasattr(st, 'session_state'):
            st.session_state[cache_key] = bundle
            
        return bundle

    @staticmethod
    def build_sabana_bundle(active_entity: str, periodo_actual: str, periodo_comp: str = "Ninguno", map_balance_df: pd.DataFrame = None, map_pl_df: pd.DataFrame = None):
        """
        Construye el paquete maestro de sábanas sin usar caché.
        """
        is_consolidated = active_entity.startswith("[GRUPO]")
        clean_entity_name = active_entity.replace("[GRUPO] ", "").strip()
        empresa_path = os.path.join("data", "empresas", clean_entity_name)
        
        if map_balance_df is None:
            map_bal_local = os.path.join(empresa_path, "map_balance.xlsx")
            map_balance_df = pd.read_excel(map_bal_local, dtype=str) if os.path.exists(map_bal_local) else None
            
        if map_pl_df is None:
            map_pl_local = os.path.join(empresa_path, "map_pl.xlsx")
            map_pl_df = pd.read_excel(map_pl_local, dtype=str) if os.path.exists(map_pl_local) else None

        # 1. Determinar empresas involucradas
        companies = []
        if is_consolidated:
            db = SessionLocal()
            try:
                grupo_obj = db.query(ConsolidationGroup).filter_by(nombre_grupo=clean_entity_name).first()
                if grupo_obj:
                    companies.append(grupo_obj.empresa_matriz)
                    if grupo_obj.filial_is_group:
                        def get_sub_companies(sub_g_id):
                            sub_g = db.query(ConsolidationGroup).filter_by(id=sub_g_id).first()
                            if sub_g:
                                c = [sub_g.empresa_matriz]
                                if sub_g.filial_is_group:
                                    c.extend(get_sub_companies(int(sub_g.empresa_filial)))
                                else:
                                    c.append(sub_g.empresa_filial)
                                return c
                            return []
                        companies.extend(get_sub_companies(int(grupo_obj.empresa_filial)))
                    else:
                        companies.append(grupo_obj.empresa_filial)
            finally:
                db.close()
        else:
            companies = [active_entity]
            
        if not companies:
            companies = [active_entity]

        # 2. Cargar contextos y sábanas individuales
        contexts = {}
        sabanas_individuales = {}
        
        _tb_cache_act = {}
        _tb_cache_comp = {}

        from src.core.snapshot_vault_manager import SnapshotVaultManager

        for co in companies:
            tb_act = TrialBalanceDB.get_trial_balance(co, periodo_actual)
            tb_comp = TrialBalanceDB.get_trial_balance(co, periodo_comp) if periodo_comp and periodo_comp != "Ninguno" else None
            
            _tb_cache_act[co] = tb_act
            _tb_cache_comp[co] = tb_comp

            # Resolver mapeos/contexto de periodo_actual (Congelado si está Cerrado)
            mb_act, mp_act = map_balance_df, map_pl_df
            frozen_ctx_act = None
            if periodo_actual and (SnapshotVaultManager.has_active_snapshot(co, periodo_actual) or SnapshotVaultManager.has_active_snapshot(active_entity, periodo_actual)):
                f_mb, f_mp = SnapshotVaultManager.get_frozen_mappings(co, periodo_actual)
                if f_mb is None:
                    f_mb, f_mp = SnapshotVaultManager.get_frozen_mappings(active_entity, periodo_actual)
                if f_mb is not None:
                    mb_act = f_mb
                if f_mp is not None:
                    mp_act = f_mp
                frozen_ctx_act = SnapshotVaultManager.get_frozen_entity_context(co, periodo_actual, target_company=co)
                if frozen_ctx_act is None:
                    frozen_ctx_act = SnapshotVaultManager.get_frozen_entity_context(active_entity, periodo_actual, target_company=co)

            # Resolver mapeos/contexto de periodo_comp (Congelado si está Cerrado)
            mb_comp, mp_comp = map_balance_df, map_pl_df
            frozen_ctx_comp = None
            if periodo_comp and periodo_comp != "Ninguno" and (SnapshotVaultManager.has_active_snapshot(co, periodo_comp) or SnapshotVaultManager.has_active_snapshot(active_entity, periodo_comp)):
                f_mb_c, f_mp_c = SnapshotVaultManager.get_frozen_mappings(co, periodo_comp)
                if f_mb_c is None:
                    f_mb_c, f_mp_c = SnapshotVaultManager.get_frozen_mappings(active_entity, periodo_comp)
                if f_mb_c is not None:
                    mb_comp = f_mb_c
                if f_mp_c is not None:
                    mp_comp = f_mp_c
                frozen_ctx_comp = SnapshotVaultManager.get_frozen_entity_context(co, periodo_comp, target_company=co)
                if frozen_ctx_comp is None:
                    frozen_ctx_comp = SnapshotVaultManager.get_frozen_entity_context(active_entity, periodo_comp, target_company=co)

            sab_bal_act = build_balance_sabana(tb_act, mb_act)
            sab_pl_act = build_pl_sabana(None, mp_act, tb_act)

            sab_bal_comp = build_balance_sabana(tb_comp, mb_comp) if tb_comp is not None else None
            sab_pl_comp = build_pl_sabana(None, mp_comp, tb_comp) if tb_comp is not None else None

            ctx_act = frozen_ctx_act if frozen_ctx_act is not None else build_entity_context(tb_act, mb_act, mp_act, empresa_name=co, periodo_str=periodo_actual)
            if tb_comp is not None:
                ctx_comp = frozen_ctx_comp if frozen_ctx_comp is not None else build_entity_context(tb_comp, mb_comp, mp_comp, empresa_name=co, periodo_str=periodo_comp)
            else:
                ctx_comp = {'nota1': {}, 'nota2': {}, 'pl': {}}

            contexts[co] = {
                'actual': ctx_act,
                'comp': ctx_comp
            }
            
            sabanas_individuales[co] = {
                'bal_actual': sab_bal_act,
                'bal_comp': sab_bal_comp,
                'pl_actual': sab_pl_act,
                'pl_comp': sab_pl_comp
            }

        # 3. Si es Consolidado, compilar Sábana Consolidada Combinada
        if is_consolidated:
            tb_act_list = [df for df in _tb_cache_act.values() if df is not None and not df.empty]
            tb_comp_list = [df for df in _tb_cache_comp.values() if df is not None and not df.empty]

            tb_act_comb = pd.concat(tb_act_list, ignore_index=True) if tb_act_list else None
            tb_comp_comb = pd.concat(tb_comp_list, ignore_index=True) if tb_comp_list else None

            frozen_comb_act = SnapshotVaultManager.get_frozen_entity_context(active_entity, periodo_actual, target_company=active_entity) if SnapshotVaultManager.has_active_snapshot(active_entity, periodo_actual) else None
            frozen_comb_comp = SnapshotVaultManager.get_frozen_entity_context(active_entity, periodo_comp, target_company=active_entity) if (periodo_comp and periodo_comp != "Ninguno" and SnapshotVaultManager.has_active_snapshot(active_entity, periodo_comp)) else None

            mb_comb_act, mp_comb_act = map_balance_df, map_pl_df
            if SnapshotVaultManager.has_active_snapshot(active_entity, periodo_actual):
                f_mb, f_mp = SnapshotVaultManager.get_frozen_mappings(active_entity, periodo_actual)
                if f_mb is not None: mb_comb_act = f_mb
                if f_mp is not None: mp_comb_act = f_mp

            mb_comb_comp, mp_comb_comp = map_balance_df, map_pl_df
            if periodo_comp and periodo_comp != "Ninguno" and SnapshotVaultManager.has_active_snapshot(active_entity, periodo_comp):
                f_mb_c, f_mp_c = SnapshotVaultManager.get_frozen_mappings(active_entity, periodo_comp)
                if f_mb_c is not None: mb_comb_comp = f_mb_c
                if f_mp_c is not None: mp_comb_comp = f_mp_c

            ctx_act_comb = frozen_comb_act if frozen_comb_act is not None else build_entity_context(tb_act_comb, mb_comb_act, mp_comb_act, empresa_name=companies, periodo_str=periodo_actual)
            if tb_comp_comb is not None:
                ctx_comp_comb = frozen_comb_comp if frozen_comb_comp is not None else build_entity_context(tb_comp_comb, mb_comb_comp, mp_comb_comp, empresa_name=companies, periodo_str=periodo_comp)
            else:
                ctx_comp_comb = {'nota1': {}, 'nota2': {}, 'pl': {}}

            # Aplicar eliminaciones intercompañía
            def apply_intercompany_eliminations(ctx):
                if not ctx: return ctx
                for rubro, items in ctx.get('pl', {}).items():
                    for k in list(items.keys()):
                        if 'intercompany' in k.lower() or 'intercompa' in k.lower():
                            items[k]['val'] = 0.0
                return ctx

            ctx_act_comb = apply_intercompany_eliminations(ctx_act_comb)
            ctx_comp_comb = apply_intercompany_eliminations(ctx_comp_comb)

            contexts[active_entity] = {
                'actual': ctx_act_comb,
                'comp': ctx_comp_comb
            }
            contexts['Consolidado'] = contexts[active_entity]

        bundle = {
            'active_entity': active_entity,
            'is_consolidated': is_consolidated,
            'companies': companies,
            'periodo_actual': periodo_actual,
            'periodo_comp': periodo_comp,
            'contexts': contexts,
            'sabanas_individuales': sabanas_individuales,
            'map_balance_df': map_balance_df,
            'map_pl_df': map_pl_df
        }
        
        return bundle

    @staticmethod
    def clear_sabana_cache():
        """
        Limpia todas las claves de Sábana Maestro de session_state.
        """
        if hasattr(st, 'session_state'):
            for k in list(st.session_state.keys()):
                if k.startswith("_sabana_maestro_bundle__") or k.startswith("_entity_ctx__") or k.startswith("_note_result__"):
                    del st.session_state[k]
