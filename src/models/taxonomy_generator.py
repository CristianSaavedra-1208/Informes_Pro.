import unicodedata
import pandas as pd
from functools import lru_cache
from src.models.database import SessionLocal
from src.models.taxonomy_master import TaxonomyMasterRecord

@lru_cache(maxsize=1024)
def clean_str(s):
    if pd.isna(s): return ""
    s = str(s).strip().upper()
    s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
    return s

@lru_cache(maxsize=1024)
def generate_prefix3(word):
    clean = ''.join(c for c in word if c.isalnum())
    if len(clean) == 0: return "XXX"
    return clean[:3]

def process_mapping_for_taxonomy(df, map_type, empresa_name):
    """
    Escanea un dataframe de mapeo (Balance o PL) y auto-genera/upserta
    los códigos de taxonomía faltantes en la base de datos SQL.
    Returns: amount of new codes generated.
    """
    db = SessionLocal()
    new_codes_count = 0
    
    try:
        # Load existing taxonomy from DB to memory for fast checking
        existing_records = db.query(TaxonomyMasterRecord).filter_by(empresa=empresa_name).all()
        # Dictionary of (Upper(Nombre_Linea_ES), Reporte_Destino) -> Record
        name_to_record = {(str(r.nombre_linea_es).strip().upper(), str(r.reporte_destino).strip()): r for r in existing_records}
        
        # We also need a fast way to get the latest counter for 100/200/300
        # For simplicity in auto-gen, we can just use a large base or query max.
        # But wait, actually, if a user uploads a new item, we just assign a large number or 999 
        # or find max of prefix.
        
        def assign_code(name, destino, specific_prefix=""):
            c_name = clean_str(name)
            p3 = generate_prefix3(c_name)
            
            # Find max counter for this prefix
            # E.g. finding matches for ACT_C or FC_OPE
            import re
            base_prefix = specific_prefix
            if not base_prefix:
                base_prefix = "GEN_C" # Fallback
                
            matches = [r.id_reporte for r in existing_records if r.id_reporte.startswith(base_prefix)]
            max_num = 0
            for code in matches:
                parts = code.split("_")
                try:
                    num = int(parts[-1])
                    if num > max_num: max_num = num
                except: pass
            
            new_num = max_num + 100
            if new_num < 100: new_num = 100
            
            return f"{base_prefix}_{p3}_{str(new_num).zfill(3)}"
 
        # 1. PROCESS FLUIDITY (Flujo de Efectivo) in ANY file
        flujo_col = next((c for c in df.columns if "flujo" in c.lower() and "efectivo" in c.lower()), None)
        if flujo_col:
            unique_flows = df[flujo_col].dropna().unique()
            for flow in unique_flows:
                flow_str = str(flow).strip()
                if flow_str == "": continue
                if (flow_str.upper(), "Flujo de Efectivo") not in name_to_record:
                    c_flow = clean_str(flow_str)
                    pref = "FC_OPE" 
                    if "INVERSI" in c_flow: pref = "FC_INV"
                    elif "FINANCIA" in c_flow or "PRESTAMO" in c_flow or "DIVIDENDO" in c_flow or "CAPITAL" in c_flow: 
                        pref = "FC_FIN"
                    
                    new_code = assign_code(flow_str, "Flujo de Efectivo", pref)
                    new_rec = TaxonomyMasterRecord(
                        empresa=empresa_name,
                        id_reporte=new_code,
                        reporte_destino="Flujo de Efectivo",
                        nombre_linea_es=flow_str
                    )
                    db.add(new_rec)
                    db.commit()
                    db.refresh(new_rec)
                    existing_records.append(new_rec)
                    name_to_record[(flow_str.upper(), "Flujo de Efectivo")] = new_rec
                    new_codes_count += 1
 
        # 2. PROCESS BALANCE
        if map_type == "Balance":
            clasif_col = next((c for c in df.columns if "Clasificaci" in c and "balance" in c.lower()), None)
            if clasif_col:
                unique_bals = df[clasif_col].dropna().unique()
                for rubro in unique_bals:
                    rubro_str = str(rubro).strip()
                    if rubro_str == "": continue
                    if (rubro_str.upper(), "Balance") not in name_to_record:
                        c_rubro = clean_str(rubro_str)
                        pref = "ACT_C" 
                        if "PASIVO" in c_rubro: pref = "PAS_C"
                        elif "PATRIMONIO" in c_rubro or "CAPITAL" in c_rubro: pref = "PAT_C"
                        
                        new_code = assign_code(rubro_str, "Balance", pref)
                        new_rec = TaxonomyMasterRecord(
                            empresa=empresa_name,
                            id_reporte=new_code,
                            reporte_destino="Balance",
                            nombre_linea_es=rubro_str
                        )
                        db.add(new_rec)
                        db.commit()
                        existing_records.append(new_rec)
                        name_to_record[(rubro_str.upper(), "Balance")] = new_rec
                        new_codes_count += 1
                        
            # Notes processing can be complex but works similarly, omitted for brevity right here 
            # to preserve system simplicity on runtime, focusing on Master Items and Flows.
            
        # 3. PROCESS P&L
        if map_type == "PL":
            expected_pl_cols = [c for c in df.columns if "cuenta" not in c.lower() and "detalle" not in c.lower() and "unnamed" not in c.lower() and "flujo" not in c.lower() and "reporte" not in c.lower() and "asociada" not in c.lower()]
            for col in expected_pl_cols:
                col_str = str(col).strip()
                if col_str == "": continue
                if (col_str.upper(), "P&L") not in name_to_record:
                    c_rubro = clean_str(col_str)
                    pref = "ER_OPE"
                    if "FINANCIER" in c_rubro or "IMPUESTO" in c_rubro or "CAMBIO" in c_rubro: pref = "ER_NOO"
                    
                    new_code = assign_code(col_str, "P&L", pref)
                    new_rec = TaxonomyMasterRecord(
                        empresa=empresa_name,
                        id_reporte=new_code,
                        reporte_destino="P&L",
                        nombre_linea_es=col_str
                    )
                    db.add(new_rec)
                    db.commit()
                    existing_records.append(new_rec)
                    name_to_record[(col_str.upper(), "P&L")] = new_rec
                    new_codes_count += 1
 
    except Exception as e:
        print(f"Error en auto-generador de taxonomía: {e}")
    finally:
        db.close()
        
    return new_codes_count
