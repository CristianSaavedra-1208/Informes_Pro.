import json
import os
import pandas as pd
from typing import Dict, Any, Tuple
from src.models.database import SessionLocal
from src.models.taxonomy_master import TaxonomyMasterRecord
from src.core.excel_utils import sort_accounts, heal_mapping_fields

def save_manual_mapping(
    cuenta_ejemplo: str,
    nombre_loc: str,
    clasificacion: str,
    nota_config_str: str,
    tipo_mapeo: str,
    empresa_seleccionada: str,
    empresa_path: str,
    session_state: Dict[str, Any]
) -> Tuple[bool, str]:
    """
    Guarda la clasificación manual tanto en Base de Datos (TaxonomyMasterRecord) 
    como en los DataFrames de sesión (map_balance_df, map_pl_df) y archivos Excel.
    """
    try:
        clasificacion = clasificacion.strip()
        cuenta_ejemplo = cuenta_ejemplo.strip()
        nombre_loc = nombre_loc.strip()
        nota_config = json.loads(nota_config_str)
        import random
        # 1. Asegurar en DB la taxonomía
        db = SessionLocal()
        tax_rec = db.query(TaxonomyMasterRecord).filter_by(
            empresa=empresa_seleccionada,
            nombre_linea_es=clasificacion,
            id_nota_asociada=nota_config_str
        ).first()
        
        if not tax_rec:
            new_rec = TaxonomyMasterRecord(
                empresa=empresa_seleccionada,
                id_reporte=f"MNL_{random.randint(1000,99999)}",
                reporte_destino="Balance" if tipo_mapeo == "Balance" else "P&L",
                nombre_linea_es=clasificacion,
                id_nota_asociada=nota_config_str.strip()
            )
            db.add(new_rec)

        # Registrar también las sub-notas/desgloses en la DB para que estén disponibles
        if tipo_mapeo != "Balance":
            p_code = nota_config.get("ID_Nota_Asociada", "")
            desglose = nota_config.get("Desglose_Nota_ES", "")
            if desglose:
                sub_rec = db.query(TaxonomyMasterRecord).filter_by(
                    empresa=empresa_seleccionada,
                    reporte_destino="Notas P&L",
                    nombre_linea_es=clasificacion,
                    desglose_nota_es=desglose
                ).first()
                if not sub_rec:
                    note_code = f"N_PL_{random.randint(1000,9999)}"
                    new_sub = TaxonomyMasterRecord(
                        empresa=empresa_seleccionada,
                        id_reporte=note_code,
                        reporte_destino="Notas P&L",
                        nombre_linea_es=clasificacion,
                        id_nota_asociada=p_code,
                        desglose_nota_es=desglose
                    )
                    db.add(new_sub)
        else:
            # Balance sub-notes
            parent_rec = db.query(TaxonomyMasterRecord.id_reporte).filter_by(
                empresa=empresa_seleccionada,
                nombre_linea_es=clasificacion,
                reporte_destino="Balance"
            ).first()
            p_code = parent_rec[0] if parent_rec else ""
            for k, val in nota_config.items():
                if k.startswith("nota") and val:
                    sub_rec = db.query(TaxonomyMasterRecord).filter_by(
                        empresa=empresa_seleccionada,
                        reporte_destino="Notas Balance",
                        nombre_linea_es=clasificacion,
                        desglose_nota_es=val
                    ).first()
                    if not sub_rec:
                        note_code = f"N_BAL_{random.randint(1000,9999)}"
                        new_sub = TaxonomyMasterRecord(
                            empresa=empresa_seleccionada,
                            id_reporte=note_code,
                            reporte_destino="Notas Balance",
                            nombre_linea_es=clasificacion,
                            id_nota_asociada=p_code,
                            desglose_nota_es=val
                        )
                        db.add(new_sub)

        db.commit()
        db.close()

        # 2. Guardar en Balance
        if tipo_mapeo == "Balance" and "map_balance_df" in session_state:
            df_bal = session_state["map_balance_df"]
            cls_col = next((c for c in df_bal.columns if "clasific" in c.lower()), None)
            cta_col = next((c for c in df_bal.columns if "cuenta" in c.lower() and "n°" in c.lower()), "N° de Cuenta")
            nom_col = next((c for c in df_bal.columns if "nombre" in c.lower() and "cuenta" in c.lower()), "Nombre cuenta")
            
            if cls_col and cta_col in df_bal.columns:
                idx = df_bal[df_bal[cta_col].astype(str) == cuenta_ejemplo].index
                if not idx.empty:
                    df_bal.loc[idx[0], cls_col] = clasificacion
                    for colk, valk in nota_config.items():
                        if colk in df_bal.columns: df_bal.loc[idx[0], colk] = valk
                else:
                    nr = {c: "" for c in df_bal.columns}
                    nr[cta_col] = cuenta_ejemplo
                    nr[nom_col] = nombre_loc
                    nr[cls_col] = clasificacion
                    for colk, valk in nota_config.items():
                        if colk in df_bal.columns: nr[colk] = valk
                    df_bal = pd.concat([df_bal, pd.DataFrame([nr])], ignore_index=True)
                    
                df_bal = sort_accounts(df_bal, cta_col)
                df_bal = heal_mapping_fields(df_bal)
                session_state["map_balance_df"] = df_bal
                df_bal.to_excel(os.path.join(empresa_path, "map_balance.xlsx"), index=False)
                import streamlit as st
                st.cache_data.clear()
        
        # 3. Guardar en P&L
        if tipo_mapeo != "Balance" and "map_pl_df" in session_state:
            df_pl = session_state["map_pl_df"]
            cta_col_pl = next((c for c in df_pl.columns if "cuenta" in c.lower() and "n°" in c.lower()), "N° de cuenta")
            nom_col_pl = next((c for c in df_pl.columns if "nombre" in c.lower() and "cuenta" in c.lower()), "Nombre de la cuenta")
            
            if cta_col_pl in df_pl.columns:
                idx = df_pl[df_pl[cta_col_pl].astype(str) == cuenta_ejemplo].index
                if not idx.empty:
                    for c in df_pl.columns:
                        if c != cta_col_pl and c != nom_col_pl:
                            df_pl.loc[idx[0], c] = None
                    if clasificacion in df_pl.columns:
                        df_pl.loc[idx[0], clasificacion] = "x"
                    n_pl_col = next((c for c in df_pl.columns if "nota" in c.lower()), None)
                    if n_pl_col and "ID_Nota_Asociada" in nota_config:
                        df_pl.loc[idx[0], n_pl_col] = nota_config["ID_Nota_Asociada"]
                else:
                    nr = {c: "" for c in df_pl.columns}
                    nr[cta_col_pl] = cuenta_ejemplo
                    nr[nom_col_pl] = nombre_loc
                    if clasificacion in df_pl.columns:
                        nr[clasificacion] = "x"
                    n_pl_col = next((c for c in df_pl.columns if "nota" in c.lower()), None)
                    if n_pl_col and "ID_Nota_Asociada" in nota_config:
                        nr[n_pl_col] = nota_config["ID_Nota_Asociada"]
                    df_pl = pd.concat([df_pl, pd.DataFrame([nr])], ignore_index=True)
                    
                df_pl = sort_accounts(df_pl, cta_col_pl)
                session_state["map_pl_df"] = df_pl
                df_pl.to_excel(os.path.join(empresa_path, "map_pl.xlsx"), index=False)
                
        # Replicar si es global
        global_opt = "🌐 [GLOBAL] Configuración General"
        if empresa_seleccionada == global_opt:
            import shutil
            empresas_dir = os.path.dirname(empresa_path) # data/empresas
            real_empresas = sorted([d for d in os.listdir(empresas_dir) if os.path.isdir(os.path.join(empresas_dir, d))])
            for co in real_empresas:
                if co == "Pacifico SpA":
                    continue
                dest_dir = os.path.join(empresas_dir, co)
                if tipo_mapeo == "Balance":
                    shutil.copy2(os.path.join(empresa_path, "map_balance.xlsx"), os.path.join(dest_dir, "map_balance.xlsx"))
                else:
                    shutil.copy2(os.path.join(empresa_path, "map_pl.xlsx"), os.path.join(dest_dir, "map_pl.xlsx"))

        return True, "Guardado exitoso"
    except Exception as e:
        return False, str(e)
