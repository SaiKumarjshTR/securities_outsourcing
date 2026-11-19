#!/bin/bash
# =============================================================================
# docker_entrypoint.sh — starts CodeMeter + ABBYY LicensingService, then
# Streamlit in the foreground (so the container exits if Streamlit dies).
# =============================================================================
set -e

# ABBYY_LICENSES_DIR normally lives under /tmp (see Dockerfile comment: Plexus
# enforces a read-only root filesystem, so /var/lib/... isn't writable there).
# /tmp itself is usually a fresh tmpfs mount supplied by the platform at
# container start, which wipes anything baked into that path at image build
# time — so the directory must be (re)created here, every startup, not in the
# Dockerfile. mkdir -p is a no-op if it already exists (e.g. local dev without
# a tmpfs mount).
mkdir -p "${ABBYY_LICENSES_DIR:-/tmp/abbyy_licenses}"
chmod 777 "${ABBYY_LICENSES_DIR:-/tmp/abbyy_licenses}"

# Restore ABBYY/CodeMeter mutable state from the build-time seed (see
# Dockerfile: /var/lib/ABBYY, /var/lib/CodeMeter, /var/log/CodeMeter,
# /etc/wibu/CodeMeter, and a few config files under /opt/ABBYY/FREngine12
# are all symlinked into /tmp/abbyy_runtime/ so activation/CodeMeter can
# write to them under a read-only root filesystem). A new pod starts with an
# empty /tmp, so every target is recreated from the seed. A container that is
# restarted inside the SAME pod normally finds its /tmp again: only targets
# that are missing are seeded, so an activated license is kept (the start-up
# job below checks that ABBYY still accepts it before it activates anything).
mkdir -p /tmp/abbyy_runtime
for _name in var_lib_ABBYY var_lib_CodeMeter var_log_CodeMeter etc_wibu_CodeMeter \
             Bin_LicensingSettings.xml CommonBin_LicensingSettings.xml \
             SamplesConfig.h SamplesConfig.java; do
    _target="/tmp/abbyy_runtime/$_name"
    if [ ! -e "$_target" ]; then
        cp -a "/opt/abbyy-seed/$_name" "$_target"
    fi
done
date +%s > /tmp/abbyy_runtime/boot_ts

# Count the container starts inside this pod, so the platform log and the License
# Setup page can tell a restarted container (count > 1, /tmp survived) from a new
# pod (count 1, empty /tmp). Plus the facts that decide whether ABBYY can run here.
_starts=0
if [ -r /tmp/abbyy_runtime/start_count ]; then
    _starts=$(cat /tmp/abbyy_runtime/start_count 2>/dev/null || echo 0)
fi
case "$_starts" in ''|*[!0-9]*) _starts=0 ;; esac
_starts=$((_starts + 1))
echo "$_starts" > /tmp/abbyy_runtime/start_count
_mem=unknown
if [ -r /sys/fs/cgroup/memory.max ]; then
    _mem=$(cat /sys/fs/cgroup/memory.max 2>/dev/null || echo unknown)
elif [ -r /sys/fs/cgroup/memory/memory.limit_in_bytes ]; then
    _mem=$(cat /sys/fs/cgroup/memory/memory.limit_in_bytes 2>/dev/null || echo unknown)
fi
case "$_mem" in ''|*[!0-9]*) ;; *) _mem="$((_mem / 1048576)) MiB" ;; esac
_cpu=unknown
if [ -r /sys/fs/cgroup/cpu.max ]; then _cpu=$(cat /sys/fs/cgroup/cpu.max 2>/dev/null || echo unknown); fi
_shm=$(df -m /dev/shm 2>/dev/null | awk 'NR==2 {print $2}')
if [ "$_starts" -gt 1 ]; then
    echo "[entrypoint] Container start #$_starts in this pod (host ${HOSTNAME:-unknown}): /tmp survived, an activated license is kept."
else
    echo "[entrypoint] Container start #1 in this pod (host ${HOSTNAME:-unknown}): fresh /tmp."
fi
echo "[entrypoint] Limits: memory ${_mem}, cpu.max ${_cpu}, /dev/shm ${_shm:-unknown} MiB (ABBYY needs at least 1024)."

echo "[entrypoint] Starting CodeMeter runtime..."
/usr/sbin/CodeMeterLin -f &
sleep 2

echo "[entrypoint] Starting ABBYY LicensingService..."
/usr/local/bin/ABBYY/SDK/12/Licensing/LicensingService /standalone &
sleep 2

if pgrep -f LicensingService >/dev/null 2>&1; then
    echo "[entrypoint] LicensingService is running."
else
    echo "[entrypoint] WARNING: LicensingService did not start — PDF conversion will fail until a license is activated via the License Setup page and the service is restarted."
fi

# Default production license from Plexus (ABBYY_SERIAL_NUMBER + ABBYY_CUSTOMER_PROJECT_ID).
# Runs in the background so Streamlit and the Plexus health probes are never delayed;
# it logs status only, never the license values.
python3 /app/abbyy_license.py autoactivate &

echo "[entrypoint] Starting Streamlit..."
exec streamlit run streamlit_app.py \
    --server.port=8501 \
    --server.address=0.0.0.0 \
    --server.headless=true
