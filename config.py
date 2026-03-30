import os
from dotenv import load_dotenv

load_dotenv()

try:
    import streamlit as st
except Exception:
    st = None


def _get_secret(name: str, default: str = "") -> str:
    value = os.getenv(name)
    if value:
        return value

    if st is not None:
        try:
            return st.secrets.get(name, default)
        except Exception:
            pass

    return default


# OpenRouter Configuration
OPENROUTER_API_KEY = _get_secret("OPENROUTER_API_KEY", "")
OPENROUTER_BASE_URL = "https://openrouter.ai/api/v1"
DEFAULT_MODEL = "anthropic/claude-sonnet-4"

# Available models on OpenRouter
AVAILABLE_MODELS = [
    "anthropic/claude-sonnet-4",
    "anthropic/claude-haiku-4",
    "openai/gpt-4o",
    "openai/gpt-4o-mini",
    "google/gemini-2.0-flash-001",
]

# Vision model for scanned PDF OCR (cheap, fast)
VISION_MODEL = "google/gemini-2.0-flash-001"

# Excel Configuration
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "templates", "tracker_template.xlsx")
HEADER_ROW_SEARCH_RANGE = 10  # Search first N rows to find the header row
MIN_HEADER_COLUMNS = 5  # Minimum non-empty columns to consider a row as header
