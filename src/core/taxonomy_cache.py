"""
Taxonomy Cache: Centralized cache for taxonomy columns and lines.
Eliminates redundant SQLite queries on every page render.
"""
from typing import List, Optional
from src.models.database import SessionLocal
from src.models.taxonomy_master import TaxonomyMasterRecord

_taxonomy_columns_cache = {}

PL_ORDER_LIST = [
    "ingresos de arriendo fibra optica",
    "ingresos de actividades ordinarias",
    "costo de ventas",
    "acceso a infraestructura fibra optica",
    "costos de uso fibra optica",
    "depreciacion operacional",
    "depreciacion y amortizacion operacional",
    "otros ingresos por funcion",
    "costos de distribucion",
    "gastos de administracion",
    "depreciacion y amortizaciones",
    "otros egresos por funcion",
    "resultado por inversion en empresas relacionadas",
    "ingresos financieros",
    "ingresos financieros con empresas relacionadas",
    "ingresos financieros ic",
    "costos financieros",
    "diferencias de cambio",
    "resultado por unidad de reajuste",
    "resultados por unidades de reajuste",
    "ganancia (perdida) por impuesto a las ganancias",
    "resultado por impuestos a las ganancias"
]

def get_pl_sort_key(name: str) -> int:
    if not name:
        return 9999
    norm = name.lower().strip().replace('á', 'a').replace('é', 'e').replace('í', 'i').replace('ó', 'o').replace('ú', 'u')
    if norm in PL_ORDER_LIST:
        return PL_ORDER_LIST.index(norm)
    for idx, item in enumerate(PL_ORDER_LIST):
        if item in norm or norm in item:
            return idx
    return 9999

def get_pl_taxonomy_columns(empresa: str) -> List[str]:
    """
    Retorna la lista ordenada de columnas de P&L definidas en taxonomy_master para la empresa.
    Usa caché en memoria de proceso para evitar consultas recurrentes en cada re-render.
    """
    if empresa in _taxonomy_columns_cache:
        return list(_taxonomy_columns_cache[empresa])

    db = SessionLocal()
    try:
        tax_recs = db.query(TaxonomyMasterRecord.nombre_linea_es).filter_by(
            empresa=empresa,
            reporte_destino="P&L"
        ).order_by(TaxonomyMasterRecord.id_reporte).all()

        db_cols = [r[0] for r in tax_recs if r[0]]
        db_cols = list(dict.fromkeys(db_cols))
        db_cols = sorted(db_cols, key=get_pl_sort_key)
        _taxonomy_columns_cache[empresa] = db_cols
        return list(db_cols)
    except Exception:
        return []
    finally:
        db.close()

def clear_taxonomy_cache(empresa: Optional[str] = None) -> None:
    """
    Invalida el caché de taxonomía para una empresa específica o globalmente si empresa es None.
    """
    global _taxonomy_columns_cache
    if empresa is not None:
        _taxonomy_columns_cache.pop(empresa, None)
    else:
        _taxonomy_columns_cache.clear()
