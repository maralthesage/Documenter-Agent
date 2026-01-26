import os
from dotenv import load_dotenv

load_dotenv()

# Ollama Configuration
OLLAMA_BASE_URL = os.getenv("OLLAMA_BASE_URL", "http://localhost:11434")
OLLAMA_MODEL = os.getenv("OLLAMA_MODEL", "mistral-small3.2:latest")

# Processing Configuration
MAX_SAMPLE_ROWS = 50
MAX_COLUMN_NAME_LENGTH = 100
CHUNK_SIZE = 4  # Number of columns to analyze at once

# Documentation Configuration
MAX_FORMULA_LENGTH = 200
INCLUDE_DATA_SAMPLES = True
SAMPLE_SIZE = 10

# Streamlit Configuration
PAGE_TITLE = "Excel Documentation AI Agent"
PAGE_ICON = "📊"
