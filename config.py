import os

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DATA_DIR = os.path.join(BASE_DIR, 'data')

if not os.path.exists(DATA_DIR):
    os.makedirs(DATA_DIR, exist_ok=True)

CHROMA_DIR = os.path.join(DATA_DIR, 'chroma')

class Config:
    SECRET_KEY = os.environ.get('SECRET_KEY', 'oes-super-secret-key-2026-v1')
    SQLALCHEMY_DATABASE_URI = os.environ.get(
        'DATABASE_URL', 
        f"sqlite:///{os.path.join(DATA_DIR, 'oes.db')}"
    )
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    
    # OpenRouter AI Configuration
    OPEN_ROUTER_API_KEY = os.environ.get(
        'OPEN_ROUTER_API_KEY', 
        ''
    )
    OPEN_ROUTER_BASE_URL = os.environ.get('OPEN_ROUTER_BASE_URL', 'https://openrouter.ai/api/v1')
    OPEN_ROUTER_MODEL = os.environ.get('OPEN_ROUTER_MODEL', 'google/gemini-2.5-flash')
    
    # Chroma Vector DB Path
    CHROMA_PERSIST_DIRECTORY = CHROMA_DIR
