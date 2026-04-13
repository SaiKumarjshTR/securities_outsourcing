"""
SGML Pipeline — FastAPI Application
====================================
Endpoints:
  GET  /          → API info
  GET  /health    → liveness + readiness probe
  POST /convert   → upload DOCX, receive SGML
"""
import os
import logging
from contextlib import asynccontextmanager

from fastapi import FastAPI, File, UploadFile, HTTPException, status
from fastapi.responses import JSONResponse, PlainTextResponse
from starlette.middleware.base import BaseHTTPMiddleware

from app import config
from app.models import ConvertResponse, HealthResponse, InfoResponse
from app import pipeline_runner

# Prefix that Plexus ingress passes through to the pod (does not strip it).
_PLEXUS_PREFIX = "/208321/securities-commission-conversion"

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("sgml_pipeline")


# ── Startup / shutdown lifecycle ────────────────────────────────────────────
@asynccontextmanager
async def lifespan(application: FastAPI):
    # Create temp directory
    os.makedirs(config.TEMP_DIR, exist_ok=True)
    log.info("SGML Pipeline API starting…")
    log.info("Temp dir: %s", config.TEMP_DIR)
    log.info("Pipeline available: %s", pipeline_runner.is_pipeline_available())
    yield
    log.info("SGML Pipeline API shutting down.")


app = FastAPI(
    title="SGML Pipeline API",
    description=(
        "Thomson Reuters Securities SGML Pipeline — converts legal/regulatory "
        "DOCX documents to SGML format for content publishing."
    ),
    version="0.0.3",
    lifespan=lifespan,
)


class _StripPrefixMiddleware(BaseHTTPMiddleware):
    """Strip the Plexus ingress prefix before FastAPI routing.

    Plexus forwards the full path (e.g. /208321/.../health) to the pod
    without stripping the prefix, so we handle it here.
    """
    async def dispatch(self, request, call_next):
        path = request.scope.get("path", "")
        if path.startswith(_PLEXUS_PREFIX):
            new_path = path[len(_PLEXUS_PREFIX):] or "/"
            request.scope["path"] = new_path
            request.scope["raw_path"] = new_path.encode("latin-1")
        return await call_next(request)


app.add_middleware(_StripPrefixMiddleware)


# ── Routes ───────────────────────────────────────────────────────────────────

@app.get("/", response_model=InfoResponse, tags=["Info"])
async def root() -> InfoResponse:
    """API information endpoint."""
    return InfoResponse(
        name="SGML Pipeline API",
        version="0.0.1",
        description="Converts legal/regulatory DOCX documents to SGML.",
        endpoints={
            "GET /health": "Liveness and readiness probe",
            "POST /convert": "Convert a DOCX file to SGML (multipart upload)",
        },
    )


@app.get("/health", response_model=HealthResponse, tags=["Health"])
@app.get("/healthz", response_model=HealthResponse, tags=["Health"], include_in_schema=False)
async def health() -> HealthResponse:
    """Health check endpoint — /health and /healthz both supported."""
    return HealthResponse(
        status="healthy",
        pipeline=pipeline_runner.is_pipeline_available(),
        llm=config.USE_LLM,
        rag=config.RAG_ENABLED,
    )


@app.post("/convert", response_model=ConvertResponse, tags=["Pipeline"])
async def convert_document(
    file: UploadFile = File(..., description="DOCX file to convert"),
    doc_name: str = None,
) -> ConvertResponse:
    """
    Upload a DOCX file and receive the converted SGML output.

    - **file**: The `.docx` file (multipart/form-data)
    - **doc_name**: Optional document identifier; defaults to the uploaded filename stem
    """
    # Validate file type
    if not file.filename.lower().endswith(".docx"):
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="Only .docx files are accepted.",
        )

    # Resolve document name
    stem = os.path.splitext(file.filename)[0] if file.filename else "document"
    resolved_name = doc_name or stem

    # Read file bytes with size guard
    max_bytes = config.MAX_UPLOAD_SIZE_MB * 1024 * 1024
    file_bytes = await file.read()
    if len(file_bytes) > max_bytes:
        raise HTTPException(
            status_code=status.HTTP_413_REQUEST_ENTITY_TOO_LARGE,
            detail=f"File exceeds maximum size of {config.MAX_UPLOAD_SIZE_MB} MB.",
        )
    if len(file_bytes) == 0:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="Uploaded file is empty.",
        )

    log.info("Received: %s (%d bytes)", resolved_name, len(file_bytes))

    if not pipeline_runner.is_pipeline_available():
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail=(
                "Pipeline script (batch_runner_standalone.py) not found in the "
                "container. Rebuild the image with the pipeline file included."
            ),
        )

    result = pipeline_runner.run_pipeline(file_bytes, resolved_name)

    if result["status"] == "error":
        log.error("Pipeline error for %s: %s", resolved_name, result.get("message"))
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=result.get("message", "Pipeline conversion failed."),
        )

    log.info("Conversion successful: %s (score=%.1f)", resolved_name, result.get("score") or 0)
    return ConvertResponse(
        status="success",
        doc_name=resolved_name,
        sgml=result.get("sgml"),
        score=result.get("score"),
    )


@app.get("/convert/{doc_name}/sgml", response_class=PlainTextResponse, tags=["Pipeline"])
async def get_sgml_plaintext(doc_name: str) -> str:
    """Placeholder — use POST /convert to submit a document."""
    return f"POST to /convert with your DOCX file to convert '{doc_name}'."
