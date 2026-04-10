"""Pydantic models for request / response schemas."""
from pydantic import BaseModel
from typing import Optional


class ConvertResponse(BaseModel):
    status: str                  # "success" | "error"
    doc_name: str
    sgml: Optional[str] = None   # SGML content (on success)
    message: Optional[str] = None  # Error detail (on failure)
    score: Optional[float] = None  # Internal confidence score (0-100)


class HealthResponse(BaseModel):
    status: str       # "healthy" | "degraded"
    pipeline: bool    # Pipeline initialised
    llm: bool         # LLM client available
    rag: bool         # RAG manager available


class InfoResponse(BaseModel):
    name: str
    version: str
    description: str
    endpoints: dict
