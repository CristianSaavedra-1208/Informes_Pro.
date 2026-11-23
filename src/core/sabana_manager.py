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
        grupo_id_val = None
        if is_consolidated:
            db = SessionLocal()
            try:
                grupo_obj = db.query(ConsolidationGroup).filter_by(nombre_grupo=clean_entity_name).first()
                if grupo_obj:
                    grupo_id_val = grupo_obj.id
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
        from src.core.excel_utils import read_excel_cached

        def _load_company_mappings(co_name, fallback_mb, fallback_mp):
            co_mb, co_mp = fallback_mb, fallback_mp
            mb_path = os.path.join("data", "empresas", co_name, "map_balance.xlsx")
            mp_path = os.path.join("data", "empresas", co_name, "map_pl.xlsx")
            try:
                if os.path.exists(mb_path):
                    co_mb = read_excel_cached(mb_path, dtype=str)
            except Exception:
                pass
            try:
                if os.path.exists(mp_path):
                    co_mp = read_excel_cached(mp_path, dtype=str)
            except Exception:
                pass
            return co_mb, co_mp

        def _merge_entity_contexts(ctx_list):
            merged_n1 = {}
            merged_n2 = {}
            merged_pl = {}
            for ctx in ctx_list:
                if not ctx:
                    continue
                for n_key, target_dict in (('nota1', merged_n1), ('nota2', merged_n2)):
                    for ck, v in (ctx.get(n_key) or {}).items():
                        if ck not in target_dict:
                            target_dict[ck] = {
                                'val': 0.0, 'inicial': 0.0, 'debitos': 0.0, 'creditos': 0.0,
                                'accts': [], 'acct_names': [], 'acct_details': []
                            }
                        target_dict[ck]['val'] += float(v.get('val', 0.0) or 0.0)
                        target_dict[ck]['inicial'] += float(v.get('inicial', 0.0) or 0.0)
                        target_dict[ck]['debitos'] += float(v.get('debitos', 0.0) or 0.0)
                        target_dict[ck]['creditos'] += float(v.get('creditos', 0.0) or 0.0)
                        target_dict[ck]['accts'].extend(list(v.get('accts') or []))
                        target_dict[ck]['acct_names'].extend(list(v.get('acct_names') or []))
                        target_dict[ck]['acct_details'].extend([dict(d) for d in (v.get('acct_details') or [])])
                for rubro_col, rubro_items in (ctx.get('pl') or {}).items():
                    if rubro_col not in merged_pl:
                        merged_pl[rubro_col] = {}
                    for ck, v in (rubro_items or {}).items():
                        if ck not in merged_pl[rubro_col]:
                            merged_pl[rubro_col][ck] = {
                                'val': 0.0, 'inicial': 0.0, 'debitos': 0.0, 'creditos': 0.0,
                                'accts': [], 'acct_names': []
                            }
                        merged_pl[rubro_col][ck]['val'] += float(v.get('val', 0.0) or 0.0)
                        merged_pl[rubro_col][ck]['inicial'] += float(v.get('inicial', 0.0) or 0.0)
                        merged_pl[rubro_col][ck]['debitos'] += float(v.get('debitos', 0.0) or 0.0)
                        merged_pl[rubro_col][ck]['creditos'] += float(v.get('creditos', 0.0) or 0.0)
                        merged_pl[rubro_col][ck]['accts'].extend(list(v.get('accts') or []))
                        merged_pl[rubro_col][ck]['acct_names'].extend(list(v.get('acct_names') or []))
            return {
                'nota1': merged_n1,
                'nota2': merged_n2,
                'pl': merged_pl
            }

        for co in companies:
            tb_act = TrialBalanceDB.get_trial_balance(co, periodo_actual)
            tb_comp = TrialBalanceDB.get_trial_balance(co, periodo_comp) if periodo_comp and periodo_comp != "Ninguno" else None
            
            _tb_cache_act[co] = tb_act
            _tb_cache_comp[co] = tb_comp

            co_default_mb, co_default_mp = _load_company_mappings(co, map_balance_df, map_pl_df)

            # Resolver mapeos/contexto de periodo_actual (Congelado si la empresa o el grupo está Cerrado)
            mb_act, mp_act = co_default_mb, co_default_mp
            frozen_ctx_act = None
            if periodo_actual:
                if SnapshotVaultManager.has_active_snapshot(co, periodo_actual):
                    f_mb, f_mp = SnapshotVaultManager.get_frozen_mappings(co, periodo_actual)
                    if f_mb is not None:
                        mb_act = f_mb
                    if f_mp is not None:
                        mp_act = f_mp
                    frozen_ctx_act = SnapshotVaultManager.get_frozen_entity_context(co, periodo_actual, target_company=co)
                if frozen_ctx_act is None and SnapshotVaultManager.has_active_snapshot(active_entity, periodo_actual):
                    frozen_ctx_act = SnapshotVaultManager.get_frozen_entity_context(active_entity, periodo_actual, target_company=co)

            # Resolver mapeos/contexto de periodo_comp (Congelado si la empresa o el grupo está Cerrado)
            mb_comp, mp_comp = co_default_mb, co_default_mp
            frozen_ctx_comp = None
            if periodo_comp and periodo_comp != "Ninguno":
                if SnapshotVaultManager.has_active_snapshot(co, periodo_comp):
                    f_mb_c, f_mp_c = SnapshotVaultManager.get_frozen_mappings(co, periodo_comp)
                    if f_mb_c is not None:
                        mb_comp = f_mb_c
                    if f_mp_c is not None:
                        mp_comp = f_mp_c
                    frozen_ctx_comp = SnapshotVaultManager.get_frozen_entity_context(co, periodo_comp, target_company=co)
                if frozen_ctx_comp is None and SnapshotVaultManager.has_active_snapshot(active_entity, periodo_comp):
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

        # 3. Si es Consolidado, compilar Sábana Consolidada Combinada por agregación aditiva de contextos individuales
        if is_consolidated:
            tb_comp_list = [df for df in _tb_cache_comp.values() if df is not None and not df.empty]

            frozen_comb_act = SnapshotVaultManager.get_frozen_entity_context(active_entity, periodo_actual, target_company=active_entity) if SnapshotVaultManager.has_active_snapshot(active_entity, periodo_actual) else None
            frozen_comb_comp = SnapshotVaultManager.get_frozen_entity_context(active_entity, periodo_comp, target_company=active_entity) if (periodo_comp and periodo_comp != "Ninguno" and SnapshotVaultManager.has_active_snapshot(active_entity, periodo_comp)) else None

            import copy
            ctx_act_comb = copy.deepcopy(frozen_comb_act) if frozen_comb_act is not None else _merge_entity_contexts([contexts[co]['actual'] for co in companies if co in contexts])
            if frozen_comb_comp is not None:
                ctx_comp_comb = copy.deepcopy(frozen_comb_comp)
            elif tb_comp_list or any((contexts.get(co, {}).get('comp') or {}).get('nota1') or (contexts.get(co, {}).get('comp') or {}).get('pl') for co in companies):
                ctx_comp_comb = _merge_entity_contexts([contexts[co]['comp'] for co in companies if co in contexts])
            else:
                ctx_comp_comb = {'nota1': {}, 'nota2': {}, 'pl': {}}

            # Aplicar eliminaciones intercompañía
            def _norm_pl_key(s):
                if not s:
                    return ""
                from src.reporting.note_generator import clean_label
                return clean_label(s)

            def apply_intercompany_eliminations(ctx, per_str=None):
                if not ctx:
                    return ctx
                pl_dict = ctx.get('pl', {})

                # Sincronizar contra Hoja de Trabajo respetando los asientos de consolidación ya aplicados
                if grupo_id_val and per_str and per_str != "Ninguno":
                    try:
                        from src.core.consolidacion_engine import generar_hoja_trabajo
                        df_h, _ = generar_hoja_trabajo(grupo_id_val, per_str)
                        if df_h is not None and not df_h.empty:
                            idx_er_h = df_h[df_h['Balance clasificado'] == 'Estado de Resultados'].index
                            df_pl_h = df_h.loc[idx_er_h[0] + 1:] if not idx_er_h.empty else df_h
                            hoja_cons_map = {}
                            for _, r_h in df_pl_h.iterrows():
                                li_h = _norm_pl_key(r_h.get('Balance clasificado'))
                                val_c = r_h.get('CONSOLIDADO')
                                if li_h and isinstance(val_c, (int, float)) and not pd.isna(val_c):
                                    hoja_cons_map[li_h] = abs(float(val_c))

                            for rubro, items in pl_dict.items():
                                r_norm = _norm_pl_key(rubro)
                                if r_norm not in hoja_cons_map:
                                    continue
                                target_abs = hoja_cons_map[r_norm]
                                cur_sum = sum(float(v.get('val', 0.0) or 0.0) for v in items.values())
                                diff = cur_sum - target_abs
                                if target_abs == 0.0:
                                    for ik in list(items.keys()):
                                        items[ik]['val'] = 0.0
                                    continue
                                if abs(diff) <= 1.0:
                                    continue
                                ic_keys = [
                                    k for k in items.keys()
                                    if any(w in _norm_pl_key(k) for w in ('intercompa', 'intercompany', 'relacionad'))
                                    or _norm_pl_key(k).endswith(' ic')
                                    or ' ic ' in _norm_pl_key(k)
                                ]
                                if diff > 1.0 and ic_keys:
                                    for ik in ic_keys:
                                        ik_val = float(items[ik].get('val', 0.0) or 0.0)
                                        if abs(ik_val - diff) <= 10.0:
                                            items[ik]['val'] = 0.0
                                            break
                            return ctx
                    except Exception:
                        pass

                for rubro, items in pl_dict.items():
                    for k in list(items.keys()):
                        k_low = _norm_pl_key(k)
                        if (
                            'intercompany' in k_low
                            or 'intercompa' in k_low
                            or 'intereses prestamos empresas relacionadas' in k_low
                        ):
                            items[k]['val'] = 0.0
                return ctx

            def apply_consolidation_journal_entries(ctx, per_str=None):
                """
                Inyecta en el contexto consolidado de Notas los ajustes contables de consolidación
                (PPA de Intangibles y PPE, Plusvalía, reclasificaciones de Impuestos Corrientes
                y depreciación/amortización PPA en resultados), alineado exactamente con `generar_hoja_trabajo`.
                """
                if not ctx or not grupo_id_val or not per_str or per_str == "Ninguno":
                    return ctx

                from collections import defaultdict
                from sqlalchemy import or_
                from src.core.consolidacion_engine import resolver_montos_asiento, is_pl_account, obtener_saldos_base
                from src.core.consolidacion_account_mapper import load_consolidation_account_map, resolve_je_account_and_nota

                def _add_bal_item(sec_name, item_key, delta_val, acct_id, acct_name, clasif_bal):
                    if abs(delta_val) < 0.01:
                        return
                    sec = ctx.setdefault(sec_name, {})
                    if item_key not in sec:
                        sec[item_key] = {
                            'val': 0.0, 'inicial': 0.0, 'debitos': 0.0, 'creditos': 0.0,
                            'accts': [], 'acct_names': [], 'acct_details': []
                        }
                    sec[item_key]['val'] += float(delta_val)
                    if acct_id not in sec[item_key]['accts']:
                        sec[item_key]['accts'].append(acct_id)
                        sec[item_key]['acct_names'].append(acct_name)
                    sec[item_key]['acct_details'].append({
                        'cuenta_id': acct_id,
                        'nombre': acct_name,
                        'val': float(delta_val),
                        'inicial': 0.0,
                        'debitos': 0.0,
                        'creditos': 0.0,
                        'clasificacion_bal': clasif_bal,
                    })

                def _add_pl_item(rubro_key, sub_key, delta_val, acct_id, acct_name):
                    if abs(delta_val) < 0.01:
                        return
                    pl_sec = ctx.setdefault('pl', {})
                    r_dict = pl_sec.setdefault(rubro_key, {})
                    if sub_key not in r_dict:
                        for k_ex in list(r_dict.keys()):
                            if k_ex.startswith(sub_key + " ("):
                                sub_key = k_ex
                                break
                    if sub_key not in r_dict:
                        r_dict[sub_key] = {
                            'val': 0.0, 'inicial': 0.0, 'debitos': 0.0, 'creditos': 0.0,
                            'accts': [], 'acct_names': []
                        }
                    r_dict[sub_key]['val'] += float(delta_val)
                    if acct_id not in r_dict[sub_key]['accts']:
                        r_dict[sub_key]['accts'].append(acct_id)
                        r_dict[sub_key]['acct_names'].append(acct_name)

                def _apply_tax_netting(primary_key, fallback_keys, delta_val, acct_id, acct_name, clasif_bal):
                    if abs(delta_val) < 0.01:
                        return
                    sec = ctx.setdefault('nota1', {})
                    if delta_val > 0:
                        _add_bal_item('nota1', primary_key, delta_val, acct_id, acct_name, clasif_bal)
                        return
                    rem = float(delta_val)  # negative reduction
                    for k_cand in [primary_key] + list(fallback_keys):
                        if k_cand in sec and float(sec[k_cand].get('val', 0.0) or 0.0) > 0:
                            cur_v = float(sec[k_cand]['val'])
                            absorb = min(cur_v, -rem)
                            _add_bal_item('nota1', k_cand, -absorb, acct_id, acct_name, clasif_bal)
                            rem += absorb
                            if abs(rem) < 0.01:
                                break
                    if abs(rem) >= 0.01:
                        _add_bal_item('nota1', primary_key, rem, acct_id, acct_name, clasif_bal)

                db = SessionLocal()
                try:
                    entries_p = db.query(ConsolidationJournalEntry).filter(
                        ConsolidationJournalEntry.grupo_id == grupo_id_val,
                        or_(
                            ConsolidationJournalEntry.periodo == per_str,
                            (ConsolidationJournalEntry.periodo < per_str) & (ConsolidationJournalEntry.es_recurrente == True)
                        )
                    ).order_by(
                        ConsolidationJournalEntry.asiento_codigo.asc(),
                        ConsolidationJournalEntry.num_linea.asc(),
                        ConsolidationJournalEntry.id.asc()
                    ).all()

                    if not entries_p:
                        return ctx

                    acc_map = load_consolidation_account_map(companies[0] if companies else None)
                    base_data_p = obtener_saldos_base(grupo_id_val, per_str, db=db)
                    col_base_cache = {}
                    vouchers_dict = defaultdict(list)
                    for e in entries_p:
                        v_key = (
                            e.columna_ajuste,
                            e.glosa,
                            getattr(e, 'asiento_codigo', None) or f"AST-{per_str}-{e.id}"
                        )
                        vouchers_dict[v_key].append(e)

                    tax_asset_adj = 0.0
                    tax_liab_adj = 0.0

                    for (col_val, _, _), entries_list in vouchers_dict.items():
                        try:
                            resolved = resolver_montos_asiento(
                                grupo_id_val,
                                per_str,
                                entries_list,
                                db=db,
                                columna_destino=col_val,
                                preloaded_base_data=base_data_p,
                                preloaded_asientos_periodo=entries_p,
                                col_base_cache=col_base_cache,
                            )
                        except Exception:
                            resolved = [{
                                "debe_calculado": e.debe or 0.0,
                                "haber_calculado": e.haber or 0.0,
                            } for e in entries_list]

                        for idx, e in enumerate(entries_list):
                            res_line = resolved[idx] if idx < len(resolved) else {}
                            debe_val = float(res_line.get("debe_calculado", e.debe or 0.0))
                            haber_val = float(res_line.get("haber_calculado", e.haber or 0.0))
                            neto = debe_val - haber_val
                            if abs(neto) < 0.01:
                                continue

                            es_prior_year_static_pl = (
                                str(e.periodo)[:4] < str(per_str)[:4]
                                and (not bool(getattr(e, 'elimina_saldo_total', False)))
                                and is_pl_account(e.linea_item)
                            )
                            if es_prior_year_static_pl:
                                continue

                            cuenta_cod, _ = resolve_je_account_and_nota(
                                e.linea_item,
                                cuenta_codigo=getattr(e, 'cuenta_codigo', None),
                                linea_nota=getattr(e, 'linea_nota', None),
                                account_map=acc_map,
                            )
                            if not cuenta_cod:
                                continue

                            acc_info = acc_map.get(cuenta_cod, {})
                            acct_name = acc_info.get('nombre_cuenta') or f"Ajuste {cuenta_cod}"

                            if cuenta_cod == '9999050':
                                # #N08.1 Intangibles -> Cartera de clientes (Debe: Bruto nota1, Haber: Amort. Acum. nota2)
                                if neto > 0:
                                    _add_bal_item('nota1', 'cartera de clientes', neto, cuenta_cod, acct_name, 'Activos intangibles distinto a la plusvalía')
                                else:
                                    _add_bal_item('nota2', 'cartera de clientes', neto, cuenta_cod, acct_name, 'Activos intangibles distinto a la plusvalía')
                            elif cuenta_cod == '9999055':
                                # #N09.1 Activo Fijo -> Asignacion de compra (Bruto nota1) / Amortizacion PPA (Deprec. Acum. nota2)
                                if neto > 0:
                                    _add_bal_item('nota1', 'asignacion de compra', neto, cuenta_cod, acct_name, 'Propiedades, plantas y equipos')
                                else:
                                    _add_bal_item('nota2', 'amortizacion ppa', neto, cuenta_cod, acct_name, 'Propiedades, plantas y equipos')
                            elif cuenta_cod == '9999060':
                                if neto > 0:
                                    _add_bal_item('nota1', 'derecho de uso (1)', neto, cuenta_cod, acct_name, 'Activo por derechos de uso')
                                else:
                                    _add_bal_item('nota2', 'derecho de uso (1)', neto, cuenta_cod, acct_name, 'Activo por derechos de uso')
                            elif cuenta_cod == '9999080':
                                _add_bal_item('nota1', 'plusvalia por adquisiciones de empresas', neto, cuenta_cod, acct_name, 'Plusvalia')
                            elif cuenta_cod == '9999035':
                                tax_asset_adj += neto
                            elif cuenta_cod == '9999220':
                                # En nota1 los pasivos se expresan con signo positivo (crédito > 0, débito < 0)
                                tax_liab_adj += (-neto)
                            elif cuenta_cod == '9999300':
                                _add_pl_item('ingresos de arriendo fibra optica', 'uso fibra optica', -neto, cuenta_cod, acct_name)
                            elif cuenta_cod == '9999305':
                                # #N21.1 Ingresos intercompañía (eliminación IC en Ingresos de actividades ordinarias)
                                _add_pl_item('ingresos de actividades ordinarias', 'ingresos intercompania', -neto, cuenta_cod, acct_name)
                            elif cuenta_cod == '9999310':
                                # #N21.2 Costo de ventas (ej. Costos venta de materiales IC)
                                sub_c = _norm_pl_key(getattr(e, 'linea_nota', None)) or 'costos venta de materiales ic'
                                _add_pl_item('costo de ventas', sub_c, neto, cuenta_cod, acct_name)
                            elif cuenta_cod == '9999315':
                                _add_pl_item('acceso a infraestructura fibra optica', 'costos de uso fibra optica', neto, cuenta_cod, acct_name)
                            elif cuenta_cod == '9999320':
                                # #N21.2 Depreciación operacional PPA
                                _add_pl_item('depreciacion operacional', 'depreciacion activos fijos de operacion', neto, cuenta_cod, acct_name)
                            elif cuenta_cod == '9999325':
                                _add_pl_item('otros ingresos por funcion', 'otros', -neto, cuenta_cod, acct_name)
                            elif cuenta_cod == '9999330':
                                # #N22.2 Amortización cartera de clientes PPA
                                _add_pl_item('depreciacion y amortizaciones', 'amortizacion cartera clientes', neto, cuenta_cod, acct_name)
                            elif cuenta_cod == '9999335':
                                # #N22.1 Gastos de administración IC
                                sub_ga = _norm_pl_key(getattr(e, 'linea_nota', None)) or 'gastos de servicios intercompania'
                                _add_pl_item('gastos de administracion', sub_ga, neto, cuenta_cod, acct_name)
                            elif cuenta_cod == '9999340':
                                _add_pl_item('otros egresos por funcion', 'otros', neto, cuenta_cod, acct_name)
                            elif cuenta_cod in ('9999350', '9999355'):
                                li_norm = _norm_pl_key(getattr(e, 'linea_item', None))
                                target_r = 'ingresos financieros ic' if li_norm == 'ingresos financieros ic' else 'ingresos financieros'
                                _add_pl_item(target_r, 'intereses con empresas relacionadas', -neto, cuenta_cod, acct_name)
                            elif cuenta_cod == '9999360':
                                _add_pl_item('costos financieros', 'intereses prestamos empresas relacionadas', neto, cuenta_cod, acct_name)
                            elif cuenta_cod == '9999365':
                                _add_pl_item('diferencias de cambio', 'otros activos financieros', neto, cuenta_cod, acct_name)
                            elif cuenta_cod == '9999370':
                                _add_pl_item('resultados por unidades de reajuste', 'otros', neto, cuenta_cod, acct_name)
                            elif cuenta_cod == '9999375':
                                # #N13.4 Impuesto diferido PPA en resultados
                                _add_pl_item('resultado por impuestos a las ganancias', 'gasto por impuestos diferidos', neto, cuenta_cod, acct_name)

                    if abs(tax_asset_adj) >= 0.01:
                        _apply_tax_netting(
                            'iva credito fiscal',
                            ['creditos renta', 'pagos previsionales mensuales (ppm)'],
                            tax_asset_adj,
                            '9999035',
                            'Ajustes consolidacion_ACTIVO POR IMPUESTOS, CORRIENTES',
                            'Activo por impuestos, corrientes',
                        )
                    if abs(tax_liab_adj) >= 0.01:
                        _apply_tax_netting(
                            'iva debito fiscal',
                            ['impuestos mensuales por pagar (f29-f50)', 'impuesto retencion 10% art. 42', 'impuesto a la renta'],
                            tax_liab_adj,
                            '9999220',
                            'Ajustes consolidacion_PASIVO POR IMPUESTOS, CORRIENTES',
                            'Pasivo por impuestos, corrientes',
                        )
                except Exception as ex_je:
                    print(f"Advertencia aplicando asientos de consolidación a notas ({per_str}): {ex_je}")
                finally:
                    db.close()
                return ctx

            if frozen_comb_act is None:
                ctx_act_comb = apply_consolidation_journal_entries(ctx_act_comb, periodo_actual)
                ctx_act_comb = apply_intercompany_eliminations(ctx_act_comb, periodo_actual)
            if frozen_comb_comp is None:
                ctx_comp_comb = apply_consolidation_journal_entries(ctx_comp_comb, periodo_comp)
                ctx_comp_comb = apply_intercompany_eliminations(ctx_comp_comb, periodo_comp)

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
