import os
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# AWS Configuration
AWS_ACCESS_KEY_ID = os.getenv("AWS_ACCESS_KEY_ID")
AWS_SECRET_ACCESS_KEY = os.getenv("AWS_SECRET_ACCESS_KEY")
AWS_REGION = os.getenv("AWS_REGION", "us-east-1")

# Bedrock Model IDs
TEXT_EMBEDDER_MODEL_ID = "amazon.titan-embed-text-v2"
IMAGE_EMBEDDER_MODEL_ID = "amazon.titan-embed-image-v1"

# PostgreSQL Configuration
DB_HOST = os.getenv("DB_HOST", "localhost")
DB_PORT = int(os.getenv("DB_PORT", "5432"))
DB_NAME = os.getenv("DB_NAME", "euda_vectors")
DB_USER = os.getenv("DB_USER", "postgres")
DB_PASSWORD = os.getenv("DB_PASSWORD", "")

# Vector dimensions
TEXT_EMBEDDING_DIMENSION = 1536  # Titan text embedder dimension
