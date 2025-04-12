import psycopg2
import json
from config import DB_HOST, DB_PORT, DB_NAME, DB_USER, DB_PASSWORD, TEXT_EMBEDDING_DIMENSION

class VectorDatabase:
    def __init__(self):
        self.conn = None
        self.cursor = None

    def connect(self):
        """Connect to the PostgreSQL database server"""
        try:
            # Connect to the PostgreSQL server
            self.conn = psycopg2.connect(
                host=DB_HOST,
                port=DB_PORT,
                database=DB_NAME,
                user=DB_USER,
                password=DB_PASSWORD
            )
            self.cursor = self.conn.cursor()
            
            # Create extension if it doesn't exist
            self.cursor.execute("CREATE EXTENSION IF NOT EXISTS vector;")
            
            # Create tables if they don't exist
            self.create_tables()
            
            print("Connected to the PostgreSQL database successfully.")
        except (Exception, psycopg2.DatabaseError) as error:
            print(f"Error connecting to PostgreSQL database: {error}")
    
    def create_tables(self):
        """Create necessary tables for storing EUDA data and embeddings"""
        try:
            # Create tables for EUDA metadata and embeddings
            self.cursor.execute(f"""
                CREATE TABLE IF NOT EXISTS euda_files (
                    id SERIAL PRIMARY KEY,
                    filename VARCHAR(255) NOT NULL,
                    file_path TEXT NOT NULL,
                    file_size_kb INTEGER,
                    sheet_count INTEGER,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    metadata JSONB
                );
            """)
            
            self.cursor.execute(f"""
                CREATE TABLE IF NOT EXISTS euda_embeddings (
                    id SERIAL PRIMARY KEY,
                    euda_id INTEGER REFERENCES euda_files(id) ON DELETE CASCADE,
                    content_type VARCHAR(50) NOT NULL,
                    content_text TEXT,
                    embedding vector({TEXT_EMBEDDING_DIMENSION}),
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );
            """)
            
            # Create an index for vector similarity search
            self.cursor.execute("""
                CREATE INDEX IF NOT EXISTS euda_embeddings_embedding_idx 
                ON euda_embeddings USING ivfflat (embedding vector_cosine_ops);
            """)
            
            self.conn.commit()
            print("Tables created successfully.")
        except (Exception, psycopg2.DatabaseError) as error:
            self.conn.rollback()
            print(f"Error creating tables: {error}")

    def store_euda_metadata(self, filename, file_path, file_size_kb, sheet_count, metadata_dict):
        """Store EUDA file metadata and return the id"""
        try:
            metadata_json = json.dumps(metadata_dict)
            
            self.cursor.execute("""
                INSERT INTO euda_files (filename, file_path, file_size_kb, sheet_count, metadata)
                VALUES (%s, %s, %s, %s, %s)
                RETURNING id;
            """, (filename, file_path, file_size_kb, sheet_count, metadata_json))
            
            euda_id = self.cursor.fetchone()[0]
            self.conn.commit()
            return euda_id
        except (Exception, psycopg2.DatabaseError) as error:
            self.conn.rollback()
            print(f"Error storing EUDA metadata: {error}")
            return None

    def store_embedding(self, euda_id, content_type, content_text, embedding):
        """Store a vector embedding for a specific EUDA content piece"""
        try:
            self.cursor.execute("""
                INSERT INTO euda_embeddings (euda_id, content_type, content_text, embedding)
                VALUES (%s, %s, %s, %s);
            """, (euda_id, content_type, content_text, embedding))
            
            self.conn.commit()
            print(f"Embedding for {content_type} stored successfully.")
        except (Exception, psycopg2.DatabaseError) as error:
            self.conn.rollback()
            print(f"Error storing embedding: {error}")

    def close(self):
        """Close the database connection"""
        if self.conn is not None:
            self.cursor.close()
            self.conn.close()
            print("Database connection closed.")
