"""
autocorrect.py
==============
Revierte automáticamente los cambios de la última etapa de optimización
si validate_regression.py detectó una regresión numérica.

Uso:
    python tests/perf/autocorrect.py

Estrategia: usa backups de archivos guardados antes de cada etapa.
Los backups se crean automáticamente por stage_manager.py antes de cada cambio.
"""

import os
import sys
import shutil
import glob
from datetime import datetime

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.chdir(ROOT_DIR)
sys.path.insert(0, ROOT_DIR)

BACKUP_DIR = os.path.join("tests", "perf", "backups")

# Archivos que pueden ser modificados por las etapas de optimización
TRACKED_FILES = [
    "src/main.py",
    "src/ui_pages/informes_y_notas.py",
    "src/reporting/note_generator.py",
    "src/core/excel_utils.py",
]


def list_backups():
    """Lista todos los backups disponibles."""
    if not os.path.exists(BACKUP_DIR):
        return []
    backups = glob.glob(os.path.join(BACKUP_DIR, "etapa_*"))
    return sorted(backups, reverse=True)  # más reciente primero


def create_backup(etapa_label: str):
    """
    Guarda una copia de todos los archivos tracked ANTES de aplicar una etapa.
    Llamar a esto antes de cada cambio de código.
    """
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = os.path.join(BACKUP_DIR, f"etapa_{etapa_label}_{ts}")
    os.makedirs(backup_path, exist_ok=True)

    saved = []
    for rel_path in TRACKED_FILES:
        abs_path = os.path.join(ROOT_DIR, rel_path)
        if os.path.exists(abs_path):
            dest = os.path.join(backup_path, rel_path.replace("/", "_").replace("\\", "_"))
            shutil.copy2(abs_path, dest)
            saved.append(rel_path)

    # Guardar metadata del backup
    meta_path = os.path.join(backup_path, "_meta.txt")
    with open(meta_path, "w", encoding="utf-8") as f:
        f.write(f"etapa: {etapa_label}\n")
        f.write(f"timestamp: {ts}\n")
        f.write("archivos:\n")
        for s in saved:
            f.write(f"  - {s}\n")

    print(f"✅ Backup creado: {backup_path}")
    print(f"   Archivos guardados: {len(saved)}")
    return backup_path


def restore_backup(backup_path: str):
    """Restaura los archivos desde un backup específico."""
    if not os.path.exists(backup_path):
        print(f"❌ Backup no encontrado: {backup_path}")
        return False

    restored = []
    for rel_path in TRACKED_FILES:
        backup_file = os.path.join(backup_path, rel_path.replace("/", "_").replace("\\", "_"))
        if os.path.exists(backup_file):
            dest = os.path.join(ROOT_DIR, rel_path)
            shutil.copy2(backup_file, dest)
            restored.append(rel_path)

    if restored:
        print(f"✅ Restaurados {len(restored)} archivos desde: {backup_path}")
        for r in restored:
            print(f"   ↩️  {r}")
        return True
    else:
        print(f"⚠️  No se encontraron archivos en el backup: {backup_path}")
        return False


def run_autocorrect():
    print("=" * 60)
    print("  AUTOCORRECT — Revertir última etapa de optimización")
    print("=" * 60)

    backups = list_backups()
    if not backups:
        print("❌ No hay backups disponibles.")
        print("   Los backups se crean con: python tests/perf/autocorrect.py --backup etapaN")
        sys.exit(1)

    # Mostrar backups disponibles
    print("\nBackups disponibles:")
    for i, bp in enumerate(backups[:5]):  # mostrar los 5 más recientes
        meta_file = os.path.join(bp, "_meta.txt")
        meta = ""
        if os.path.exists(meta_file):
            with open(meta_file, "r", encoding="utf-8") as f:
                meta = f.read().strip()
        print(f"  [{i}] {os.path.basename(bp)}")
        if meta:
            for line in meta.split("\n")[:2]:
                print(f"      {line}")

    print("\n  → Restaurando el backup más reciente (el de la última etapa)...")
    latest = backups[0]
    success = restore_backup(latest)

    if success:
        print("\n✅ Rollback completado. Los archivos han sido restaurados al estado pre-etapa.")
        print("   Verifica con: python tests/perf/validate_regression.py")
    else:
        print("\n❌ El rollback falló. Revisar los backups manualmente en:")
        print(f"   {BACKUP_DIR}")

    print("=" * 60)


# ─── Modo CLI ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Autocorrección por regresión numérica")
    parser.add_argument(
        "--backup", type=str, default=None,
        help="Crear backup antes de una etapa: --backup etapa1"
    )
    parser.add_argument(
        "--restore", type=str, default=None,
        help="Restaurar un backup específico por nombre de directorio"
    )
    args = parser.parse_args()

    if args.backup:
        create_backup(args.backup)
    elif args.restore:
        restore_path = os.path.join(BACKUP_DIR, args.restore)
        restore_backup(restore_path)
    else:
        run_autocorrect()
