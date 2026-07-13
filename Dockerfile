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
LABEL version="0.0.22"
LABEL description="SGML Pipeline UI — PDF/DOCX to SGML conversion (Streamlit, ABBYY CLI direct)"

# ── System dependencies ──────────────────────────────────────────────────────
# NOTE: Per ABBYY official Docker documentation, the following libraries are
# required for FREngine12 on Linux (ubuntu:noble package names):
#   ca-certificates, libc6, libglib2.0-0, libgcc-s1, libstdc++6,
#   zlib1g, libx11-6, libfreetype6, libxext6 (libxext-dev), libice6,
#   libsm6, locales
# See: https://docs.abbyy.com/fine-reader/engine/distribution/distribution-linux/
#      running-abbyy-finereader-engine-12-inside-a-docker-container-2
#
# IMPORTANT — /dev/shm size: ABBYY requires shm_size ≥ 1GB (POSIX shared memory).
# In Kubernetes (Plexus), the default is 64MB which is NOT enough and will cause
# silent OCR failures. Add this to the Plexus pod spec:
#   volumes:
#     - name: dshm
#       emptyDir: {medium: Memory, sizeLimit: 1Gi}
#   volumeMounts:
#     - {mountPath: /dev/shm, name: dshm}
RUN apt-get update && apt-get install -y --no-install-recommends \
        build-essential \
        libglib2.0-0 \
        libsm6 \
        libxext6 \
        libxrender-dev \
        libgomp1 \
        libstdc++6 \
        libx11-6 \
        libfreetype6 \
        libice6 \
        locales \
        git \
        curl \
        procps \
        libusb-1.0-0 \
        iproute2 \
    && locale-gen en_US.UTF-8 \
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
        pip3 install --no-cache-dir --timeout 300 --retries 10 -r requirements.txt \
            --extra-index-url "https://${TR_JFROG_USERNAME}:${TR_JFROG_TOKEN}@tr1.jfrog.io/tr1/api/pypi/pypi-local/simple"; \
    else \
        pip3 install --no-cache-dir --timeout 300 --retries 10 -r requirements.txt; \
    fi

# ── Copy application code ────────────────────────────────────────────────────
COPY app/             ./app/
COPY pipeline/        ./pipeline/
COPY streamlit_app.py  ./streamlit_app.py
COPY abbyy_convert.py  ./abbyy_convert.py
COPY .streamlit/       ./.streamlit/

# ── Data directory (keying rules, vendor SGMLs, ChromaDB) ───────────────────
RUN mkdir -p /app/data/vendor_sgms /tmp/sgml_pipeline

# ── Copy static data files (if present at build time) ───────────────────────
COPY data/ /app/data/

# ── ABBYY FREngine 12 — bundled directly (no external bridge needed) ─────────
# abbyy_bundle/ is created by running: bash scripts/bundle_abbyy.sh
# It contains the full ABBYY installation from the Ubuntu host.
# All binaries are native Linux x86-64 ELF — run in any Linux container.

# TR / Zscaler CA certificates — required for ABBYY online license validation
# TR's Zscaler proxy intercepts SSL and signs with TR CA (not the real ABBYY cert).
# Without these certs the container rejects the SSL cert → "Failed to communicate
# with the online licensing service."
COPY abbyy_bundle/tr_certs/ /usr/local/share/ca-certificates/
RUN update-ca-certificates

# Install CodeMeter software license manager (from deb package)
# procps + libusb-1.0-0 are pre-installed above — CodeMeter needs them.
# Do NOT run 'apt-get install -f' here; it would remove CodeMeter due to
# missing X11/GL GUI deps that are irrelevant for a headless daemon.
COPY abbyy_bundle/codemeter.deb /tmp/codemeter.deb
RUN dpkg -i --force-depends /tmp/codemeter.deb \
    && rm /tmp/codemeter.deb

# Copy full FREngine 12 SDK (CLI, all .so libraries, language packs, etc.)
COPY abbyy_bundle/FREngine12 /opt/ABBYY/FREngine12
RUN chmod +x \
        /opt/ABBYY/FREngine12/Samples/CommandLineInterface/CommandLineInterface \
        /opt/ABBYY/FREngine12/Bin/FREngineProcessor \
        /opt/ABBYY/FREngine12/Bin/LicenseManager.Console \
    && find /opt/ABBYY/FREngine12 -name "*.sh" -exec chmod +x {} \; 2>/dev/null || true \
    && rm -f /opt/ABBYY/FREngine12/Bin/libProtection.Developer.so

# Copy ABBYY LicensingService daemon (binary + required lib dir per AdminGuide)
RUN mkdir -p /usr/local/bin/ABBYY/SDK/12/Licensing \
    && mkdir -p /usr/local/lib/ABBYY/SDK/12/Licensing
COPY abbyy_bundle/LicensingService /usr/local/bin/ABBYY/SDK/12/Licensing/LicensingService
COPY abbyy_bundle/usr_local_lib_abbyy/ /usr/local/lib/ABBYY/
RUN chmod +x /usr/local/bin/ABBYY/SDK/12/Licensing/LicensingService

# Copy full ABBYY license state (activation tokens + Protection.ccf activated state)
# Note: bundle script copies /var/lib/ABBYY/ → var_lib_abbyy/ABBYY/ (one extra level)
# so we COPY the ABBYY/ subdirectory directly into /var/lib/ABBYY/
RUN mkdir -p /var/lib/ABBYY
COPY abbyy_bundle/var_lib_abbyy/ABBYY/ /var/lib/ABBYY/
RUN chmod -R 777 /var/lib/ABBYY/

# Copy container startup script (CodeMeterLin → LicensingService → Streamlit)
COPY scripts/docker_start.sh /app/docker_start.sh
RUN chmod +x /app/docker_start.sh

# ── Runtime configuration ────────────────────────────────────────────────────
# ABBYY CLI called directly from Python (subprocess) — no HTTP bridge needed.
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PORT=8501 \
    HOST=0.0.0.0 \
    TEMP_DIR=/tmp/sgml_pipeline \
    ABBYY_CLI=/opt/ABBYY/FREngine12/Samples/CommandLineInterface/CommandLineInterface \
    ABBYY_LIB=/opt/ABBYY/FREngine12/Bin \
    LD_LIBRARY_PATH=/opt/ABBYY/FREngine12/Bin:/usr/local/lib/ABBYY/SDK/12/Licensing

# ── Health check ─────────────────────────────────────────────────────────────
# Longer start-period because ABBYY services need time to initialize.
HEALTHCHECK --interval=30s --timeout=15s --start-period=60s --retries=5 \
    CMD python3 -c "import urllib.request; urllib.request.urlopen('http://localhost:8501/_stcore/health')" || exit 1

# ── Expose port ───────────────────────────────────────────────────────────────
EXPOSE 8501

# ── Entry point ───────────────────────────────────────────────────────────────
# docker_start.sh: CodeMeterLin → LicensingService → Streamlit
CMD ["/app/docker_start.sh"]
