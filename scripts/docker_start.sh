#!/bin/bash
# ─────────────────────────────────────────────────────────────────────────────
# docker_start.sh — Container entrypoint
#
# Startup sequence:
#   1. CodeMeterLin      (ABBYY software license daemon) — background
#   2. LicensingService  (local) OR external LS config   — background/config
#   3. Streamlit         — exec (becomes PID 1)
#
# ENVIRONMENT VARIABLES:
#   ABBYY_LS_HOST  — If set, FREngine connects to this external LicensingService
#                    host instead of starting a local one. The host must have
#                    LicensingService running on port 3023 with internet access.
#                    Example: ABBYY_LS_HOST=10.0.1.50
#
# IMPORTANT: No set -e here — licensing services may fail in some environments
# and the container MUST still start Streamlit for the UI to be reachable.
# ─────────────────────────────────────────────────────────────────────────────

LICENSING_SVC="/usr/local/bin/ABBYY/SDK/12/Licensing/LicensingService"
LICENSING_LOG="/tmp/licensing_service.log"
CODEMETER_LOG="/tmp/codemeter.log"

# All LicensingSettings.xml copies that must point to the License Server.
# CRITICAL: the FREngine CLI reads its copy from the FREngine12 tree
# (Bin / CommonBin/Licensing), NOT from /usr/local/lib. Every copy must be
# patched or the engine keeps using ServerAddress="127.0.0.1" and fails.
LICENSING_SETTINGS_FILES=(
    "/opt/ABBYY/FREngine12/Bin/LicensingSettings.xml"
    "/opt/ABBYY/FREngine12/CommonBin/Licensing/LicensingSettings.xml"
    "/usr/local/lib/ABBYY/SDK/12/Licensing/LicensingSettings.xml"
)

echo "[startup] $(date -u) — Container starting (v0.0.22)"

# ── CRITICAL: /dev/shm size check ────────────────────────────────────────────
# ABBYY FREngine12 uses POSIX shared memory and requires /dev/shm ≥ 1GB.
# The Kubernetes default is 64MB which will cause silent OCR failures.
# Fix in Plexus pod spec:
#   volumes: [{name: dshm, emptyDir: {medium: Memory, sizeLimit: 1Gi}}]
#   volumeMounts: [{mountPath: /dev/shm, name: dshm}]
SHM_KB=$(df -B1024 /dev/shm 2>/dev/null | awk 'NR==2{print $2}')
if [ -z "$SHM_KB" ]; then
    echo "[startup] ⚠ WARNING: Cannot determine /dev/shm size"
elif [ "$SHM_KB" -lt 1048576 ]; then
    echo "[startup] ⚠ WARNING: /dev/shm = ${SHM_KB}KB — ABBYY requires ≥ 1GB (1048576 KB)"
    echo "[startup] → Plexus fix: add emptyDir(medium=Memory, sizeLimit=1Gi) volume at /dev/shm"
    echo "[startup] → Without this, OCR conversions may silently fail with shm errors"
else
    echo "[startup] ✓ /dev/shm = ${SHM_KB}KB (≥ 1GB required by ABBYY)"
fi

# ── Diagnostics: show what license files are present ─────────────────────────
echo "[startup] License files in /var/lib/ABBYY/SDK/12/Licenses/:"
ls -la /var/lib/ABBYY/SDK/12/Licenses/ 2>/dev/null || echo "[startup] WARNING: /var/lib/ABBYY/SDK/12/Licenses/ not found!"

# ── 1. CodeMeterLin (license daemon) ─────────────────────────────────────────
echo "[startup] Starting CodeMeterLin..."
if command -v CodeMeterLin >/dev/null 2>&1; then
    CodeMeterLin -f > "$CODEMETER_LOG" 2>&1 &
    echo "[startup] CodeMeterLin launched (PID $!)"
else
    echo "[startup] WARNING: CodeMeterLin not found — ABBYY license will fail"
fi

# ── 2. LicensingService: local or external ────────────────────────────────────
if [ -n "${ABBYY_LS_HOST:-}" ]; then
    # ── External License Server mode ──────────────────────────────────────────
    # FREngine will connect to the external LicensingService at ABBYY_LS_HOST:3023
    echo "[startup] ABBYY_LS_HOST=${ABBYY_LS_HOST} — using EXTERNAL LicensingService"
    echo "[startup] Patching all LicensingSettings.xml: ServerAddress → ${ABBYY_LS_HOST}"
    patched=0
    for settings in "${LICENSING_SETTINGS_FILES[@]}"; do
        if [ -f "$settings" ]; then
            sed -i "s/ServerAddress=\"[^\"]*\"/ServerAddress=\"${ABBYY_LS_HOST}\"/" "$settings"
            echo "[startup] ✓ patched $settings"
            patched=$((patched+1))
        else
            echo "[startup] (skip) not found: $settings"
        fi
    done
    if [ "$patched" -eq 0 ]; then
        echo "[startup] WARNING: No LicensingSettings.xml patched — conversions will fail"
    fi
    echo "[startup] Skipping local LicensingService start (using external)"
    echo "[startup] Testing connectivity to ${ABBYY_LS_HOST}:3023..."
    timeout 5 bash -c "cat < /dev/tcp/${ABBYY_LS_HOST}/3023" 2>/dev/null \
        && echo "[startup] ✓ External LicensingService REACHABLE at ${ABBYY_LS_HOST}:3023" \
        || echo "[startup] ✗ WARNING: Cannot reach ${ABBYY_LS_HOST}:3023 — conversions will fail"
else
    # ── Local LicensingService mode (default) ─────────────────────────────────
    echo "[startup] ABBYY_LS_HOST not set — starting LOCAL LicensingService"
    if [ -x "$LICENSING_SVC" ]; then
        "$LICENSING_SVC" /standalone > "$LICENSING_LOG" 2>&1 &
        echo "[startup] LicensingService launched (PID $!)"
    else
        echo "[startup] WARNING: LicensingService not found at $LICENSING_SVC"
    fi
fi

# ── Short wait for services to initialise before Streamlit checks them ────────
echo "[startup] Waiting 10s for licensing services to initialise..."
sleep 10

# ── Log service status (non-blocking diagnostics) ────────────────────────────
echo "[startup] === Service status at $(date -u) ==="
pgrep -a CodeMeterLin 2>/dev/null && echo "[startup] ✓ CodeMeterLin running" || echo "[startup] ✗ CodeMeterLin NOT running"
pgrep -a LicensingService 2>/dev/null && echo "[startup] ✓ LicensingService running" || echo "[startup] ✗ LicensingService NOT running"
echo "[startup] LicensingService log:"
cat "$LICENSING_LOG" 2>/dev/null || echo "(no log yet)"

# ── 3. Streamlit (always starts regardless of licensing status) ───────────────
echo "[startup] Starting Streamlit on port 8501..."
exec streamlit run /app/streamlit_app.py \
    --server.port=8501 \
    --server.address=0.0.0.0 \
    --server.headless=true
