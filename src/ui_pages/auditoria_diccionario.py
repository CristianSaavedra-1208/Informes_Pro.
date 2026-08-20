import streamlit as st
import pandas as pd
import os
from src.core.excel_utils import df_to_excel_bytes, read_excel_cached
from src.models.trial_balance_db import TrialBalanceDB
from src.models.pl_cubo_db import PlCuboDB
from src.models.database import SessionLocal
from src.models.taxonomy_master import TaxonomyMasterRecord
from src.models.historical_data import HistoricalDataRecord

def render(empresa_seleccionada, empresa_path):
    if empresa_seleccionada == "[GLOBAL] Configuración General" or "GLOBAL" in empresa_seleccionada:
        st.warning("Módulo de Sociedad Activa: Por favor, selecciona una empresa de trabajo específica (ej. Pacifico SpA) en la barra lateral izquierda para acceder a esta sección.")
        st.stop()
        
    st.title("Auditoría")
    st.write("Extrae directamente desde la Bóveda de SQL del Sistema tus reglas contables oficiales.")
    st.info("Descarga en Excel el Diccionario Maestro que el motor del software usa para organizar los reportes de esta empresa en tiempo real.")
    
    if st.button("Generar Diccionario Oficial F/S", type="primary"):
        db = SessionLocal()
        records = db.query(TaxonomyMasterRecord).filter_by(empresa=empresa_seleccionada).all()
        db.close()
        
        if records:
            accts_por_codigo = {}
            map_bal_path = os.path.join(empresa_path, "map_balance.xlsx")
            map_pl_path = os.path.join(empresa_path, "map_pl.xlsx")
            
            try:
                if os.path.exists(map_bal_path):
                    df_b = read_excel_cached(map_bal_path, dtype=str)
                    c_col = next((c for c in df_b.columns if "cuenta" in str(c).lower()), None)
                    id_rep_col = next((c for c in df_b.columns if "reporte" in str(c).lower() or "balance" in str(c).lower()), None)
                    id_nota_col = next((c for c in df_b.columns if "nota_asociada" in str(c).lower() or "nota" in str(c).lower()), None)
                    
                    if c_col:
                        for _, ro in df_b.iterrows():
                            acct = str(ro.get(c_col, "")).strip()
                            if acct and acct != "nan":
                                if id_rep_col:
                                    c_rep = str(ro.get(id_rep_col, "")).strip()
                                    if c_rep and c_rep != "nan":
                                        if c_rep not in accts_por_codigo: accts_por_codigo[c_rep] = []
                                        accts_por_codigo[c_rep].append(acct)
                                if id_nota_col:
                                    c_nota = str(ro.get(id_nota_col, "")).strip()
                                    if c_nota and c_nota != "nan":
                                        if c_nota not in accts_por_codigo: accts_por_codigo[c_nota] = []
                                        accts_por_codigo[c_nota].append(acct)
                                        
                if os.path.exists(map_pl_path):
                    df_p = read_excel_cached(map_pl_path, dtype=str)
                    c_col = next((c for c in df_p.columns if "cuenta" in str(c).lower()), None)
                    expected_pl_cols = [c for c in df_p.columns if "cuenta" not in str(c).lower() and "detalle" not in str(c).lower() and "flujo" not in str(c).lower() and "unnamed" not in str(c).lower()]
                    tax_name_dict = {str(r_obj.nombre_linea_es).strip().upper(): r_obj.id_reporte for r_obj in records if r_obj.nombre_linea_es}
                    tax_note_dict = {str(r_obj.desglose_nota_es).strip().upper(): r_obj.id_reporte for r_obj in records if r_obj.desglose_nota_es}
                    
                    if c_col:
                        for _, ro in df_p.iterrows():
                            acct = str(ro.get(c_col, "")).strip()
                            if not acct or acct == "nan": continue
                            for col in expected_pl_cols:
                                val = ro.get(col)
                                if not pd.isna(val) and str(val).strip() != "" and str(val).lower() != "nan":
                                    # Asociar al codigo P&L (header)
                                    code_pl = tax_name_dict.get(str(col).strip().upper())
                                    if code_pl:
                                        if code_pl not in accts_por_codigo: accts_por_codigo[code_pl] = []
                                        accts_por_codigo[code_pl].append(acct)
                                    # Asociar al codigo de Nota P&L (celda interna usando Desglose)
                                    code_nota = tax_note_dict.get(str(val).strip().upper())
                                    if code_nota:
                                        if code_nota not in accts_por_codigo: accts_por_codigo[code_nota] = []
                                        accts_por_codigo[code_nota].append(acct)
            except Exception as e:
                pass
                
            data = [{"ID_Reporte": r.id_reporte, "Reporte_Destino": r.reporte_destino, "Nombre_Linea_ES": r.nombre_linea_es, "ID_Nota_Asociada": r.id_nota_asociada, "Desglose_Nota_ES": r.desglose_nota_es, "Traduccion_Ingles": r.nombre_idioma_1, "Cuentas_ERP_Asociadas": ", ".join(sorted(list(set(accts_por_codigo.get(r.id_reporte, [])))))} for r in records]
            df_out = pd.DataFrame(data)
            excel_data = df_to_excel_bytes(df_out, 'Taxonomia')
            
            st.download_button(
                label="📥 Descargar Documento_Auditoria_Maestro.xlsx",
                data=excel_data,
                file_name="Documento_Auditoria_Maestro.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.success(f"✅ Documento generado con {len(records)} variables F/S.")
        else:
            st.warning("⚠️ Esta empresa no tiene Bóveda de Taxonomía registrada aún. Primero sube los mapeos en el Módulo 1 para activar el auto-descubrimiento, o sube manualmente la plantilla.")
            
    st.write("---")
    st.write("---")
    st.subheader("Sábanas de Cuadratura (Lead Sheets)")
    st.write("Traza cada línea del Diccionario Maestro hacia las cuentas de tu ERP que la componen (cruzando el Mapeo activo y el Trial Balance).")
    
    is_consolidated = empresa_seleccionada.startswith("[GRUPO]")
    import importlib
    import src.core.sabana_builder
    importlib.reload(src.core.sabana_builder)
    from src.core.sabana_builder import build_consolidated_balance_sabana, build_consolidated_pl_sabana
    
    if is_consolidated:
        db = SessionLocal()
        per_recs = db.query(HistoricalDataRecord.periodo).distinct().all()
        db.close()
        available_tb_periods = sorted([r[0] for r in per_recs], reverse=True)
        if not available_tb_periods: available_tb_periods = ["2026-03", "2025-12"]
        available_pl_periods = available_tb_periods
    else:
        available_tb_periods = TrialBalanceDB.get_available_periods(empresa_seleccionada)
        available_pl_periods = PlCuboDB.get_available_periods(empresa_seleccionada)
        if not available_tb_periods:
            db = SessionLocal()
            per_recs = db.query(HistoricalDataRecord.periodo).filter_by(empresa=empresa_seleccionada).distinct().all()
            db.close()
            available_tb_periods = sorted([r[0] for r in per_recs], reverse=True)
        if not available_pl_periods:
            available_pl_periods = available_tb_periods if available_tb_periods else ["2026-03", "2025-12"]
    
    col_per1, col_per2 = st.columns(2)
    with col_per1:
        if available_tb_periods:
            sel_periodo_bal = st.selectbox("📅 Período de Balance a Auditar:", available_tb_periods, key="sel_periodo_aud_bal")
        else:
            sel_periodo_bal = "2026-03"
            st.info("ℹ️ Usando período por defecto 2026-03.")

    with col_per2:
        if available_pl_periods:
            sel_periodo_pl = st.selectbox("📅 Período de P&L a Auditar:", available_pl_periods, key="sel_periodo_aud_pl")
        else:
            sel_periodo_pl = "2026-03"
            st.info("ℹ️ Usando período por defecto 2026-03.")

    col_btn_aud1, col_btn_aud2 = st.columns(2)
    
    with col_btn_aud1:
        if st.button("Generar Sábana Cuadratura (Balance)", type="secondary", use_container_width=True):
            try:
                # Cargar mapeo de balance (local o global)
                map_bal_path = os.path.join(empresa_path, "map_balance.xlsx")
                if not os.path.exists(map_bal_path): map_bal_path = "map_balance.xlsx"
                map_bal = read_excel_cached(map_bal_path, dtype=str) if os.path.exists(map_bal_path) else None
                
                global_tpl = "Plantilla de notas_v1.xlsx"
                if (map_bal is None or map_bal.empty) and os.path.exists(global_tpl):
                    try: map_bal = pd.read_excel(global_tpl, sheet_name="Mapeo Balance", dtype=str)
                    except: pass
                    
                if is_consolidated:
                    df_tie = build_consolidated_balance_sabana(empresa_seleccionada, sel_periodo_bal, map_bal)
                else:
                    tb_df = TrialBalanceDB.get_trial_balance(empresa_seleccionada, sel_periodo_bal)
                    tb_path = os.path.join(empresa_path, "temp_uploaded.xlsx")
                    if (tb_df is None or tb_df.empty) and os.path.exists(tb_path):
                        try: tb_df = read_excel_cached(tb_path, dtype=str)
                        except: pass
                        
                    if tb_df is None or tb_df.empty:
                        st.error("⚠️ No se encontró Trial Balance para esta empresa en el período seleccionado.")
                        df_tie = None
                    else:
                        from src.core.sabana_builder import build_balance_sabana
                        df_tie = build_balance_sabana(tb_df, map_bal)
                        
                if df_tie is not None and not df_tie.empty:
                    p_name = sel_periodo_bal if sel_periodo_bal else "Actual"
                    excel_data = df_to_excel_bytes(df_tie, 'Lead Sheet Balance')
                    
                    st.write("### 📊 Vista Previa: Sábana de Cuadratura (Balance)")
                    st.dataframe(df_tie, use_container_width=True)
                    
                    clean_co_name = empresa_seleccionada.replace("[GRUPO] ", "").replace(" ", "_")
                    st.download_button(
                        label=f"📥 Descargar Lead_Sheet_Balance_{clean_co_name}_{p_name}.xlsx",
                        data=excel_data,
                        file_name=f"Lead_Sheet_Balance_{clean_co_name}_{p_name}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                    st.success(f"✅ Lead Sheet / Sábana de Balance ({p_name}) generada con {len(df_tie)} filas.")
                else:
                    st.warning("⚠️ No se pudieron consolidar datos de balance para esta selección.")
            except Exception as e:
                st.error(f"Error al generar Sábana de Balance: {e}")

    with col_btn_aud2:
        if st.button("Generar Sábana Cuadratura (P&L Matrix)", type="secondary", use_container_width=True):
            try:
                # Cargar mapeo de P&L (local o global)
                map_pl_path = os.path.join(empresa_path, "map_pl.xlsx")
                if not os.path.exists(map_pl_path): map_pl_path = "map_pl.xlsx"
                map_pl = read_excel_cached(map_pl_path, dtype=str) if os.path.exists(map_pl_path) else None
                
                global_tpl = "Plantilla de notas_v1.xlsx"
                if (map_pl is None or map_pl.empty) and os.path.exists(global_tpl):
                    try: map_pl = pd.read_excel(global_tpl, sheet_name="Mapeo Ctas P&L Cubo", dtype=str)
                    except: pass
                    
                if is_consolidated:
                    df_pl_tie = build_consolidated_pl_sabana(empresa_seleccionada, sel_periodo_pl, map_pl)
                else:
                    pl_cubo = PlCuboDB.get_pl_cubo(empresa_seleccionada, sel_periodo_pl)
                    tb_df = TrialBalanceDB.get_trial_balance(empresa_seleccionada, sel_periodo_pl)
                    from src.core.sabana_builder import build_pl_sabana
                    df_pl_tie = build_pl_sabana(pl_cubo, map_pl, tb_df)
                    
                if df_pl_tie is not None and not df_pl_tie.empty:
                    p_name_pl = sel_periodo_pl if sel_periodo_pl else "Actual"
                    excel_data_pl = df_to_excel_bytes(df_pl_tie, 'Lead Sheet PL Matrix')
                    
                    st.write("### 📊 Vista Previa: Sábana de Cuadratura (P&L Matrix)")
                    st.dataframe(df_pl_tie, use_container_width=True)
                    
                    clean_co_name = empresa_seleccionada.replace("[GRUPO] ", "").replace(" ", "_")
                    st.download_button(
                        label=f"📥 Descargar Lead_Sheet_PL_{clean_co_name}_{p_name_pl}.xlsx",
                        data=excel_data_pl,
                        file_name=f"Lead_Sheet_PL_{clean_co_name}_{p_name_pl}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                    st.success(f"✅ Lead Sheet / Sábana de P&L Matrix ({p_name_pl}) generada con {len(df_pl_tie)} filas.")
                else:
                    st.warning("⚠️ No se pudieron consolidar datos de P&L para esta selección.")
            except Exception as e:
                st.error(f"Error al generar Sábana de P&L Matrix: {e}")



