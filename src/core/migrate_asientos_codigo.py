import re
from collections import defaultdict
from sqlalchemy import func
from src.models.consolidacion import ConsolidationJournalEntry

def migrar_y_autocorregir_codigos_asiento(db):
    """
    Recorre todos los registros en `consolidation_journal_entries` sin `asiento_codigo`,
    los agrupa como vouchers contables por (grupo_id, periodo, columna_ajuste, glosa)
    y les asigna correlativos únicos mensuales `AST-YYYYMM-XXX`.
    También asigna `num_linea` consecutivo a cada línea del voucher.
    """
    entries = db.query(ConsolidationJournalEntry).all()
    if not entries:
        return {"migrados": 0, "ya_con_codigo": 0, "vouchers_creados": 0}

    # Mapa de correlativos existentes por (grupo_id, periodo)
    existing_max_seq = defaultdict(int)
    
    for a in entries:
        code_val = getattr(a, 'asiento_codigo', None)
        if code_val:
            # Intentar extraer el número de la secuencia si sigue la máscara AST-YYYYMM-XXX
            match = re.search(r'AST-\d{6}-(\d+)', code_val)
            if match:
                seq = int(match.group(1))
                key = (a.grupo_id, a.periodo)
                if seq > existing_max_seq[key]:
                    existing_max_seq[key] = seq

    # Agrupar entradas no migradas por (grupo_id, periodo, columna_ajuste, glosa)
    unassigned_groups = defaultdict(list)
    ya_con_codigo_count = 0

    for a in entries:
        if not getattr(a, 'asiento_codigo', None):
            unassigned_groups[(a.grupo_id, a.periodo, a.columna_ajuste, a.glosa)].append(a)
        else:
            ya_con_codigo_count += 1

    vouchers_creados = 0
    migrados_count = 0

    # Asignar códigos únicos por cada voucher contable
    for (grupo_id, periodo, col, glosa), lines in unassigned_groups.items():
        key = (grupo_id, periodo)
        existing_max_seq[key] += 1
        seq_num = existing_max_seq[key]
        
        # Formato YYYYMM
        periodo_clean = periodo.replace('-', '').strip()
        codigo_voucher = f"AST-{periodo_clean}-{seq_num:03d}"

        for idx, line in enumerate(lines, start=1):
            line.asiento_codigo = codigo_voucher
            line.num_linea = idx
            if not line.created_by:
                line.created_by = "system_migration"
            migrados_count += 1

        vouchers_creados += 1

    if migrados_count > 0:
        db.commit()

    return {
        "migrados": migrados_count,
        "ya_con_codigo": ya_con_codigo_count,
        "vouchers_creados": vouchers_creados
    }

if __name__ == "__main__":
    from src.models.database import SessionLocal, init_db
    init_db()
    db = SessionLocal()
    try:
        res = migrar_y_autocorregir_codigos_asiento(db)
        print(f"Resultado de la migración: {res}")
    finally:
        db.close()
