# ─────────────────────────────────────────────────────────────────────────────
# SGML Pipeline — Dockerfile
# ─────────────────────────────────────────────────────────────────────────────
# Target runtime: TR AI Platform (Plexus — AWS ECS / Fargate, Linux x86-64)
# UI port      : 8501  (Streamlit)
# Entry point  : docker_entrypoint.sh (CodeMeter + ABBYY LicensingService + Streamlit)
# Health check : GET /_stcore/health  (Streamlit built-in)
# PDF -> DOCX  : ABBYY FREngine 12, direct CLI invocation (see streamlit_app.py
#                _pdf_to_docx) — the FRS14 Windows-bridge architecture was
#                retired; ABBYY now runs in-process on the same host as the UI.
# ─────────────────────────────────────────────────────────────────────────────

FROM python:3.12-slim

# Metadata
LABEL maintainer="Thomson Reuters — Securities SGML Team"
LABEL version="1.0.4"
LABEL description="SGML Pipeline UI — PDF/DOCX to SGML conversion (Streamlit + ABBYY FREngine 12)"

# ── System dependencies (app + ABBYY FRE12 runtime requirements) ─────────────
RUN apt-get update && apt-get install -y --no-install-recommends \
        build-essential \
        libglib2.0-0 \
        libsm6 \
        libxext6 \
        libxrender-dev \
        libgomp1 \
        git \
        libstdc++6 \
        libx11-6 \
        libfreetype6 \
        libice6 \
        locales \
        procps \
        libusb-1.0-0 \
        iproute2 \
        curl \
        ca-certificates \
    && echo 'en_US.UTF-8 UTF-8' >> /etc/locale.gen \
    && locale-gen \
    && rm -rf /var/lib/apt/lists/*

# ── TR/Zscaler corporate CA certs ────────────────────────────────────────────
# Without these, outbound HTTPS (ABBYY online license activation to
# account.abbyy.com, TR AI Platform/Anthropic calls, etc.) fails with
# "SSL certificate problem: unable to get local issuer certificate" — the
# corporate proxy re-signs traffic with these certs, confirmed missing from
# the base image via a local test run.
COPY abbyy_bundle/tr_certs/ /usr/local/share/ca-certificates/
RUN update-ca-certificates

# ── Working directory ────────────────────────────────────────────────────────
WORKDIR /app

# ── Python dependencies ──────────────────────────────────────────────────────
# .wheel_cache/: complete local wheelhouse for requirements.txt, pre-downloaded
# on the host via `pip3 download` (see scripts/build_push.sh) — direct pip
# installs from files.pythonhosted.org proved unreliable on this corporate
# network even via JFrog, so the build installs fully offline instead.
COPY requirements.txt .
COPY .wheel_cache/ /tmp/wheel_cache/
RUN --mount=type=secret,id=jfrog_username,env=TR_JFROG_USERNAME \
    --mount=type=secret,id=jfrog_token,env=TR_JFROG_TOKEN \
    pip3 install --no-index --find-links /tmp/wheel_cache -r requirements.txt

# ── ABBYY FREngine 12 + CodeMeter + LicensingService ─────────────────────────
# abbyy_bundle/ is staged from an already-licensed, working FRE12 installation
# (see scripts/build_push.sh) — NOT committed to git (see .gitignore), and
# contains no license/token files, only the software itself. The actual
# license is activated at runtime via pages/2_License_Setup.py (Option B:
# Serial Number + Customer Project ID, no file/password needed for this
# license type — see ABBYY email), so no credential ever needs to be baked
# into this image or into source control.
COPY abbyy_bundle/codemeter.deb /tmp/codemeter.deb
RUN dpkg -i --force-depends /tmp/codemeter.deb || apt-get install -f -y --no-install-recommends; \
    rm -f /tmp/codemeter.deb

COPY abbyy_bundle/FREngine12/ /opt/ABBYY/FREngine12/
RUN chmod +x /opt/ABBYY/FREngine12/Samples/CommandLineInterface/CommandLineInterface \
             /opt/ABBYY/FREngine12/Bin/FREngineProcessor \
             /opt/ABBYY/FREngine12/Bin/LicenseManager.Console \
             /opt/ABBYY/FREngine12/activatefre.sh \
    && find /opt/ABBYY/FREngine12 -name "*.sh" -exec chmod +x {} \; \
    && rm -f /opt/ABBYY/FREngine12/Bin/libProtection.Developer.so

COPY abbyy_bundle/usr_local_bin_abbyy/ /usr/local/bin/ABBYY/
COPY abbyy_bundle/usr_local_lib_abbyy/ /usr/local/lib/ABBYY/
RUN chmod +x /usr/local/bin/ABBYY/SDK/12/Licensing/LicensingService \
    && mkdir -p /var/lib/ABBYY/SDK/12/Licenses /tmp/abbyy_licenses \
    && chmod 777 /var/lib/ABBYY/SDK/12/Licenses /tmp/abbyy_licenses

# ── Redirect ABBYY/CodeMeter mutable state onto /tmp ─────────────────────────
# Confirmed the entire Plexus root filesystem is read-only, not just the
# license folder (reproduced locally with `docker run --read-only --tmpfs
# /tmp`). Tracing activatefre.sh (run as root during activation) and
# CodeMeterLin end to end shows they write to every path below — all fail
# with "Read-only file system" otherwise, and CodeMeterLin in particular
# exits silently (rc=1, no output) when it can't write its state dir. Each
# path is tiny (<= 24K) vs. the 3.2GB FREngine12 tree, so the pre-baked
# content is preserved under /opt/abbyy-seed/ and the real path replaced
# with a symlink into /tmp; docker_entrypoint.sh restores each target from
# its seed on every container start, since /tmp is a fresh mount each time.
RUN mkdir -p /opt/abbyy-seed \
    && cp -a /var/lib/ABBYY      /opt/abbyy-seed/var_lib_ABBYY \
    && cp -a /var/lib/CodeMeter  /opt/abbyy-seed/var_lib_CodeMeter \
    && cp -a /var/log/CodeMeter  /opt/abbyy-seed/var_log_CodeMeter \
    && cp -a /etc/wibu/CodeMeter /opt/abbyy-seed/etc_wibu_CodeMeter \
    && cp -a /opt/ABBYY/FREngine12/Bin/LicensingSettings.xml \
             /opt/abbyy-seed/Bin_LicensingSettings.xml \
    && cp -a /opt/ABBYY/FREngine12/CommonBin/Licensing/LicensingSettings.xml \
             /opt/abbyy-seed/CommonBin_LicensingSettings.xml \
    && cp -a /opt/ABBYY/FREngine12/Samples/SamplesConfig.h \
             /opt/abbyy-seed/SamplesConfig.h \
    && cp -a /opt/ABBYY/FREngine12/Samples/Java/SamplesConfig.java \
             /opt/abbyy-seed/SamplesConfig.java \
    && rm -rf /var/lib/ABBYY /var/lib/CodeMeter /var/log/CodeMeter /etc/wibu/CodeMeter \
              /opt/ABBYY/FREngine12/Bin/LicensingSettings.xml \
              /opt/ABBYY/FREngine12/CommonBin/Licensing/LicensingSettings.xml \
              /opt/ABBYY/FREngine12/Samples/SamplesConfig.h \
              /opt/ABBYY/FREngine12/Samples/Java/SamplesConfig.java \
    && ln -s /tmp/abbyy_runtime/var_lib_ABBYY      /var/lib/ABBYY \
    && ln -s /tmp/abbyy_runtime/var_lib_CodeMeter  /var/lib/CodeMeter \
    && ln -s /tmp/abbyy_runtime/var_log_CodeMeter  /var/log/CodeMeter \
    && ln -s /tmp/abbyy_runtime/etc_wibu_CodeMeter /etc/wibu/CodeMeter \
    && ln -s /tmp/abbyy_runtime/Bin_LicensingSettings.xml \
             /opt/ABBYY/FREngine12/Bin/LicensingSettings.xml \
    && ln -s /tmp/abbyy_runtime/CommonBin_LicensingSettings.xml \
             /opt/ABBYY/FREngine12/CommonBin/Licensing/LicensingSettings.xml \
    && ln -s /tmp/abbyy_runtime/SamplesConfig.h \
             /opt/ABBYY/FREngine12/Samples/SamplesConfig.h \
    && ln -s /tmp/abbyy_runtime/SamplesConfig.java \
             /opt/ABBYY/FREngine12/Samples/Java/SamplesConfig.java

# ── Guard: no license values may be baked into the image ─────────────────────
# The ABBYY sample config/property files that were copied from a licensed dev machine
# can carry a project id, token path and PASSWORD in plain text, and anything in an
# image layer can be read by whoever can pull the image. The production license comes
# only from the deployment (ABBYY_SERIAL_NUMBER / ABBYY_CUSTOMER_PROJECT_ID).
RUN python3 - <<'PY'
import glob, re, sys
bad = []
for f in ["/opt/abbyy-seed/SamplesConfig.h", "/opt/abbyy-seed/SamplesConfig.java"] + \
        glob.glob("/opt/ABBYY/FREngine12/Samples/Java/*/*.properties"):
    t = open(f, errors="replace").read()
    for pat in (r'(?:FreProjectId|LicensePath|LicensePassword)\[\]\s*=\s*L"([^"]+)"',
                r'(?m)^FREngine\.(?:projectId|licensePath|licensePassword)[ \t]*=[ \t]*(\S+)',
                r'GetCustomerProjectId\(\)\s*\{\s*return\s*"([^"]+)"',
                r'GetLicense(?:Path|Password)\(\)\s*\{\s*return\s*"([^"]+)"'):
        if re.search(pat, t):
            bad.append(f)
if bad:
    sys.exit("BUILD STOPPED: license values found in " + ", ".join(sorted(set(bad))) +
             " - blank them in abbyy_bundle/FREngine12/Samples before building.")
print("license guard: no license values baked into the image")
PY

# ABBYY_LICENSES_DIR is a separate /tmp staging dir used only by the
# Streamlit upload flow (pages/2_License_Setup.py Option A) — activatefre.sh
# copies the uploaded file from here into /var/lib/ABBYY/SDK/12/Licenses
# itself (now transparently backed by /tmp via the symlink above), so this
# doesn't need to match that path, just be writable.
ENV LD_LIBRARY_PATH=/opt/ABBYY/FREngine12/Bin:/usr/local/lib/ABBYY/SDK/12/Licensing \
    ABBYY_CLI=/opt/ABBYY/FREngine12/Samples/CommandLineInterface/CommandLineInterface \
    ABBYY_LIB=/opt/ABBYY/FREngine12/Bin \
    ABBYY_LICENSES_DIR=/tmp/abbyy_licenses \
    LANG=en_US.UTF-8 \
    LC_ALL=en_US.UTF-8

# ── Copy application code ────────────────────────────────────────────────────
COPY app/             ./app/
COPY pipeline/        ./pipeline/
COPY pages/           ./pages/
COPY streamlit_app.py ./streamlit_app.py
COPY abbyy_convert.py ./abbyy_convert.py
COPY abbyy_license.py ./abbyy_license.py
COPY .streamlit/      ./.streamlit/
COPY docker_entrypoint.sh /entrypoint.sh
RUN chmod +x /entrypoint.sh

# ── Data directory (keying rules, vendor SGMLs, ChromaDB) ───────────────────
RUN mkdir -p /app/data/vendor_sgms /tmp/sgml_pipeline

# ── Copy static data files (if present at build time) ───────────────────────
# These files are optional — mount via volume or env overrides in production
COPY data/ /app/data/

# NOTE: this image runs as root (no USER switch). CodeMeter/LicensingService
# expect to install/run with root-level access to /usr/local, /var/lib, and
# (for hardware-dongle licenses, not used here) udev — matching ABBYY's own
# reference container examples, which also run as root. Revisit if a
# non-root setup is required later.

# ── Runtime configuration ────────────────────────────────────────────────────
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PORT=8501 \
    HOST=0.0.0.0 \
    TEMP_DIR=/tmp/sgml_pipeline

# ── Health check ─────────────────────────────────────────────────────────────
# Streamlit exposes /_stcore/health  (returns 200 OK when ready).
# NOTE: Do NOT set --server.baseUrlPath — it causes Plexus readiness probe (GET /)
# to receive a 307 redirect and mark the pod unhealthy.
# Plexus ingress handles the URL prefix at the load balancer level.
HEALTHCHECK --interval=30s --timeout=15s --start-period=120s --retries=5 \
    CMD python3 -c "import urllib.request; urllib.request.urlopen('http://localhost:8501/_stcore/health')" || exit 1

# ── Expose port ───────────────────────────────────────────────────────────────
EXPOSE 8501

# ── Entry point ───────────────────────────────────────────────────────────────
# Starts CodeMeter + ABBYY LicensingService in the background, then Streamlit
# in the foreground. No --server.baseUrlPath (see health check note above).
ENTRYPOINT ["/entrypoint.sh"]
