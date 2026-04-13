# ─────────────────────────────────────────────────────────────────────────────
# SGML Pipeline — Dockerfile
# ─────────────────────────────────────────────────────────────────────────────
# Target runtime: TR AI Platform (AWS ECS / Fargate, Linux x86-64)
# UI port      : 8501  (Streamlit)
# Entry point  : streamlit run streamlit_app.py
# Health check : GET /_stcore/health  (Streamlit built-in)
# ─────────────────────────────────────────────────────────────────────────────

FROM python:3.12-slim

# Metadata
LABEL maintainer="Thomson Reuters — Securities SGML Team"
LABEL version="0.0.5"
LABEL description="SGML Pipeline UI — PDF/DOCX to SGML conversion (Streamlit)"  

# ── System dependencies ──────────────────────────────────────────────────────
RUN apt-get update && apt-get install -y --no-install-recommends \
        build-essential \
        libglib2.0-0 \
        libsm6 \
        libxext6 \
        libxrender-dev \
        libgomp1 \
        git \
    && rm -rf /var/lib/apt/lists/*

# ── Build arguments for JFrog (TR internal PyPI) ─────────────────────────────
ARG TR_JFROG_USERNAME
ARG TR_JFROG_TOKEN

# ── Working directory ────────────────────────────────────────────────────────
WORKDIR /app

# ── Python dependencies ──────────────────────────────────────────────────────
COPY requirements.txt .

# Install from TR JFrog (if credentials supplied) then public PyPI
RUN if [ -n "$TR_JFROG_USERNAME" ] && [ -n "$TR_JFROG_TOKEN" ]; then \
        pip3 install --no-cache-dir -r requirements.txt \
            --extra-index-url "https://${TR_JFROG_USERNAME}:${TR_JFROG_TOKEN}@tr1.jfrog.io/tr1/api/pypi/pypi-local/simple"; \
    else \
        pip3 install --no-cache-dir -r requirements.txt; \
    fi

# ── Copy application code ────────────────────────────────────────────────────
COPY app/             ./app/
COPY pipeline/        ./pipeline/
COPY streamlit_app.py ./streamlit_app.py
COPY .streamlit/      ./.streamlit/

# ── Data directory (keying rules, vendor SGMLs, ChromaDB) ───────────────────
RUN mkdir -p /app/data/vendor_sgms /tmp/sgml_pipeline

# ── Copy static data files (if present at build time) ───────────────────────
# These files are optional — mount via volume or env overrides in production
COPY data/ /app/data/

# ── Non-root user for security ────────────────────────────────────────────────
RUN useradd --create-home --shell /bin/bash appuser \
    && chown -R appuser:appuser /app /tmp/sgml_pipeline
USER appuser

# ── Runtime configuration ────────────────────────────────────────────────────
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PORT=8501 \
    HOST=0.0.0.0 \
    TEMP_DIR=/tmp/sgml_pipeline

# ── Health check ─────────────────────────────────────────────────────────────
# Streamlit exposes /_stcore/health  (returns 200 OK when ready).
# Update the Plexus liveness-probe path to: /_stcore/health
HEALTHCHECK --interval=30s --timeout=15s --start-period=90s --retries=3 \
    CMD python3 -c "import urllib.request; urllib.request.urlopen('http://localhost:8501/_stcore/health')" || exit 1

# ── Expose port ───────────────────────────────────────────────────────────────
EXPOSE 8501

# ── Entry point ───────────────────────────────────────────────────────────────
CMD ["streamlit", "run", "streamlit_app.py", \
     "--server.port=8501", \
     "--server.address=0.0.0.0", \
     "--server.headless=true", \
     "--server.baseUrlPath=/208321/securities-commission-conversion"]
