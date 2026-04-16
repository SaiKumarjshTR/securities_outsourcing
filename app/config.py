"""
Configuration for SGML Pipeline API.
All settings are read from environment variables with sensible defaults.
"""
import os

# ── TR AI Platform / Anthropic ────────────────────────────────────────────────
TR_AUTH_URL: str = os.getenv(
    "TR_AUTH_URL",
    "https://aiplatform.gcs.int.thomsonreuters.com/v1/anthropic/token",
)
WORKSPACE_ID: str = os.getenv("WORKSPACE_ID", "Saikumar3Y0Z")
ANTHROPIC_MODEL: str = os.getenv("ANTHROPIC_MODEL", "claude-sonnet-4-5-20250929")
OPUS_MODEL: str = os.getenv("OPUS_MODEL", "claude-opus-4-20250514")

# ── Pipeline behaviour ────────────────────────────────────────────────────────
USE_LLM: bool = os.getenv("USE_LLM", "true").lower() == "true"
LLM_BATCH_SIZE: int = int(os.getenv("LLM_BATCH_SIZE", "10"))
EXTRACT_TABLES: bool = os.getenv("EXTRACT_TABLES", "true").lower() == "true"
EXTRACT_IMAGES: bool = os.getenv("EXTRACT_IMAGES", "false").lower() == "true"
APPLY_INLINE_FORMATTING: bool = (
    os.getenv("APPLY_INLINE_FORMATTING", "true").lower() == "true"
)
MAX_TOKENS: int = int(os.getenv("MAX_TOKENS", "16000"))
IMAGE_DPI: int = int(os.getenv("IMAGE_DPI", "300"))

# ── RAG ───────────────────────────────────────────────────────────────────────
RAG_ENABLED: bool = os.getenv("RAG_ENABLED", "true").lower() == "true"
RAG_PERSIST_DIR: str = os.getenv("RAG_PERSIST_DIR", "/app/data/chroma_db")
RAG_N_RULES: int = int(os.getenv("RAG_N_RULES", "12"))
RAG_N_EXAMPLES: int = int(os.getenv("RAG_N_EXAMPLES", "25"))

# ── File paths inside container ───────────────────────────────────────────────
KEYING_RULES_PATH: str = os.getenv(
    "KEYING_RULES_PATH", "/app/data/COMPLETE_KEYING_RULES_UPDATED.txt"
)
VENDOR_SGMS_DIR: str = os.getenv("VENDOR_SGMS_DIR", "/app/data/vendor_sgms")

# ── Server ────────────────────────────────────────────────────────────────────
HOST: str = os.getenv("HOST", "0.0.0.0")
PORT: int = int(os.getenv("PORT", "8501"))
MAX_UPLOAD_SIZE_MB: int = int(os.getenv("MAX_UPLOAD_SIZE_MB", "50"))
TEMP_DIR: str = os.getenv("TEMP_DIR", "/tmp/sgml_pipeline")
