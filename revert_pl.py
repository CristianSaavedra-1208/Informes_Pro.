import os
import pandas as pd

empresa_path = r"c:\Users\cfsaa\OneDrive\Desktop\Informes_Pro\data\empresas\Pacifico SpA"
map_pl_path = os.path.join(empresa_path, "map_pl.xlsx")

if os.path.exists(map_pl_path):
    df_pl = pd.read_excel(map_pl_path)
    # Remove the single-mapping columns we wrongly added, to preserve the Matrix X feature
    if "ID_Reporte" in df_pl.columns:
        df_pl.drop(columns=["ID_Reporte"], inplace=True)
    if "ID_Nota_Asociada" in df_pl.columns:
        df_pl.drop(columns=["ID_Nota_Asociada"], inplace=True)
        
    df_pl.to_excel(map_pl_path, index=False)
    print("Reverted P&L absolute columns to preserve Matrix design.")
    
