from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
from sqlalchemy.pool import QueuePool
import psycopg2
from psycopg2.extensions import ISOLATION_LEVEL_AUTOCOMMIT
from pgvector.psycopg2 import register_vector

from database.models import Base
from config import DB_CONFIG

def get_db_url():
    """Generate database URL from configuration."""
    return f"postgresql://{DB_CONFIG['user']}:{DB_CONFIG['password']}@{DB_CONFIG['host']}:{DB_CONFIG['port']}/{DB_CONFIG['database']}"

def create_database_if_not_exists():
    """Create the database if it doesn't exist."""
    # Connect to PostgreSQL server
    conn = psycopg2.connect(
        host=DB_CONFIG['host'],
        port=DB_CONFIG['port'],
        user=DB_CONFIG['user'],
        password=DB_CONFIG['password']
    )
    conn.set_isolation_level(ISOLATION_LEVEL_AUTOCOMMIT)
    
    # Check if database exists
    cursor = conn.cursor()
    cursor.execute(f"SELECT 1 FROM pg_catalog.pg_database WHERE datname = '{DB_CONFIG['database']}'")
    exists = cursor.fetchone()
    
    if not exists:
        print(f"Creating database: {DB_CONFIG['database']}")
        cursor.execute(f"CREATE DATABASE {DB_CONFIG['database']}")
    
    cursor.close()
    conn.close()

def initialize_database():
    """Initialize the database with tables and pgvector extension."""
    # Create database if it doesn't exist
    create_database_if_not_exists()
    
    # Connect to the database
    engine = create_engine(
        get_db_url(),
        poolclass=QueuePool,
        pool_size=5,
        max_overflow=10
    )
    
    # Create a connection
    with engine.connect() as conn:
        # Check if pgvector extension exists
        result = conn.execute("SELECT 1 FROM pg_extension WHERE extname = 'vector'")
        if not result.fetchone():
            print("Creating pgvector extension")
            conn.execute("CREATE EXTENSION IF NOT EXISTS vector")
    
    # Create all tables
    Base.metadata.create_all(engine)
    
    return engine

def get_db_session():
    """Get a database session."""
    engine = create_engine(get_db_url())
    Session = sessionmaker(bind=engine)
    return Session()

# Initialize the engine
engine = initialize_database()
Session = sessionmaker(bind=engine)
