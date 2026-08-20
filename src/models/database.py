import os
from sqlalchemy import create_engine
from sqlalchemy.orm import declarative_base, sessionmaker

# Asegurar que el directorio data exista
BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DATA_DIR = os.path.join(BASE_DIR, 'data')
os.makedirs(DATA_DIR, exist_ok=True)

# URL de conexión SQLite local
DATABASE_URL = f"sqlite:///{os.path.join(DATA_DIR, 'informes_pro.db')}"

# Engine y Session factory
engine = create_engine(DATABASE_URL, echo=False, connect_args={"timeout": 30})
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)

# Base class para los modelos
Base = declarative_base()

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

def init_db():
    # Importar los modelos para que SQLAlchemy los detecte al hacer create_all
    from src.models.trial_balance import TrialBalanceRecord
    from src.models.audit_adjustment import AuditAdjustmentRecord
    from src.models.historical_data import HistoricalDataRecord, HistoricalDetailRecord
    from src.models.taxonomy_master import TaxonomyMasterRecord
    from src.models.company import CompanyEntity
    from src.models.pl_record import PlRecordDim
    from src.models.consolidacion import ConsolidationGroup, ConsolidationJournalEntry
    from src.models.cash_flow_db import CashFlowAdjustment
    from src.models.security import UserModel, AuditLogRecord
    
    # Crea todas las tablas mapeadas si no existen
    Base.metadata.create_all(bind=engine)
    
    # Migración ligera para agregar nuevas columnas
    try:
        from sqlalchemy import text
        with engine.begin() as conn:
            # 1. Check consolidation_journal_entries
            result = conn.execute(text("PRAGMA table_info(consolidation_journal_entries)"))
            columns = [row[1] for row in result.fetchall()]
            if "elimina_saldo_total" not in columns:
                conn.execute(text("ALTER TABLE consolidation_journal_entries ADD COLUMN elimina_saldo_total BOOLEAN DEFAULT 0"))
            if "linea_nota" not in columns:
                conn.execute(text("ALTER TABLE consolidation_journal_entries ADD COLUMN linea_nota VARCHAR(255)"))
            if "asiento_codigo" not in columns:
                conn.execute(text("ALTER TABLE consolidation_journal_entries ADD COLUMN asiento_codigo VARCHAR(50)"))
            if "num_linea" not in columns:
                conn.execute(text("ALTER TABLE consolidation_journal_entries ADD COLUMN num_linea INTEGER DEFAULT 1"))
            if "created_by" not in columns:
                conn.execute(text("ALTER TABLE consolidation_journal_entries ADD COLUMN created_by VARCHAR(100)"))
            if "updated_by" not in columns:
                conn.execute(text("ALTER TABLE consolidation_journal_entries ADD COLUMN updated_by VARCHAR(100)"))
            if "updated_at" not in columns:
                conn.execute(text("ALTER TABLE consolidation_journal_entries ADD COLUMN updated_at DATETIME"))
                
            # 2. Check trial_balance_records
            result_tb = conn.execute(text("PRAGMA table_info(trial_balance_records)"))
            columns_tb = [row[1] for row in result_tb.fetchall()]
            for col in ["saldo_inicial", "debitos", "creditos"]:
                if col not in columns_tb:
                    conn.execute(text(f"ALTER TABLE trial_balance_records ADD COLUMN {col} FLOAT DEFAULT 0.0"))
                    
            # 3. Check historical_detail_records
            result_hd = conn.execute(text("PRAGMA table_info(historical_detail_records)"))
            columns_hd = [row[1] for row in result_hd.fetchall()]
            for col in ["saldo_inicial", "debitos", "creditos"]:
                if col not in columns_hd:
                    conn.execute(text(f"ALTER TABLE historical_detail_records ADD COLUMN {col} FLOAT DEFAULT 0.0"))
    except Exception as e:
        print(f"Error running lightweight migration: {e}")

    # Sembrar usuarios por defecto para todos los roles si no existen
    db = SessionLocal()
    try:
        import hashlib
        default_users = [
            ("admin", "admin123", "Administrador del Sistema", "admin@informespro.cl", "Administrador"),
            ("analista_contable", "contable123", "Analista Contable", "contable@informespro.cl", "Analista Contable"),
            ("analista_reportes", "reportes123", "Analista de Reportes", "reportes@informespro.cl", "Analista de Reportes"),
            ("auditor_lector", "auditor123", "Auditor Lector", "auditor@informespro.cl", "Auditor Lector"),
        ]
        
        seeded_any = False
        for username, password, full_name, email, role in default_users:
            existing = db.query(UserModel).filter_by(usuario=username).first()
            if not existing:
                u = UserModel(
                    usuario=username,
                    password_hash=hashlib.sha256(password.encode()).hexdigest(),
                    nombre_completo=full_name,
                    email=email,
                    rol=role,
                    activo=True
                )
                db.add(u)
                seeded_any = True
                
        if seeded_any:
            init_log = AuditLogRecord(
                usuario="Sistema",
                accion="INICIALIZACION",
                entidad_id="sistema",
                detalles="Inicialización de usuarios por defecto para todos los roles."
            )
            db.add(init_log)
            db.commit()
    except Exception as e:
        db.rollback()
        print(f"Error seeding default users: {e}")
    finally:
        db.close()
