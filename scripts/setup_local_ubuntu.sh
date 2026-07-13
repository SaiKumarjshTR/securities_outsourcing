#!/usr/bin/env bash
# =============================================================================
# setup_local_ubuntu.sh  — Install & run SGML Pipeline on a bare Ubuntu machine
#
# For business users on Ubuntu 20.04 / 22.04 / 24.04 (x86-64).
# Run ONCE to install everything. After that use: sgml-pipeline start
#
# Installs:
#   • Python 3.12 + app dependencies
#   • ABBYY FineReader Engine 12 (from abbyy_bundle/ in this repo, OR
#     from a portable .tar.gz drop shipped to the user)
#   • Systemd user service for automatic startup
#   • Desktop launcher shortcut
#
# Usage:
#   git clone ... /opt/sgml-pipeline && cd /opt/sgml-pipeline
#   sudo bash scripts/setup_local_ubuntu.sh
#
# Or, if deployed as a portable bundle:
#   tar -xzf sgml-pipeline-ubuntu.tar.gz
#   sudo bash sgml-pipeline/scripts/setup_local_ubuntu.sh
#
# After install, access the app at: http://localhost:8501
# =============================================================================
set -euo pipefail

INSTALL_DIR="${INSTALL_DIR:-/opt/sgml-pipeline}"
APP_PORT="${APP_PORT:-8501}"
ABBYY_INSTALL_DIR="${ABBYY_INSTALL_DIR:-/opt/ABBYY/FREngine12}"
ABBYY_LS_HOST="${ABBYY_LS_HOST:-}"   # Leave empty — local machine HAS internet

# Colour helpers
RED='\033[0;31m'; GREEN='\033[0;32m'; YELLOW='\033[1;33m'; NC='\033[0m'
ok()   { echo -e "${GREEN}  ✓ $*${NC}"; }
warn() { echo -e "${YELLOW}  ⚠ $*${NC}"; }
err()  { echo -e "${RED}  ✗ $*${NC}"; exit 1; }
step() { echo -e "\n${YELLOW}[$(date '+%H:%M:%S')] $*${NC}"; }

echo ""
echo "============================================================"
echo "  SGML Pipeline — Ubuntu Local Installation"
echo "  Install dir : $INSTALL_DIR"
echo "  App port    : $APP_PORT"
echo "  ABBYY dir   : $ABBYY_INSTALL_DIR"
echo "============================================================"
echo ""

# ── Must run as root ──────────────────────────────────────────────────────────
[[ "$EUID" -ne 0 ]] && err "Please run as root: sudo bash $0"

# ── Detect repo root ──────────────────────────────────────────────────────────
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
BUNDLE_DIR="$REPO_DIR/abbyy_bundle"

# =============================================================================
# STEP 1 — System dependencies
# =============================================================================
step "1/8  Installing system dependencies..."
apt-get update -qq
apt-get install -y --no-install-recommends \
    python3.12 \
    python3.12-venv \
    python3-pip \
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
    procps \
    libusb-1.0-0 \
    iproute2 \
    curl \
    ca-certificates \
    >/dev/null 2>&1

locale-gen en_US.UTF-8 >/dev/null 2>&1
localedef -i en_US -f UTF-8 en_US.UTF-8 2>/dev/null || true
update-locale LANG=en_US.UTF-8 LC_ALL=en_US.UTF-8
ok "System dependencies installed"

# =============================================================================
# STEP 2 — Install CodeMeter (ABBYY license manager)
# =============================================================================
step "2/8  Installing CodeMeter..."
CODEMETER_DEB="$BUNDLE_DIR/codemeter.deb"
if [ -f "$CODEMETER_DEB" ]; then
    dpkg -i --force-depends "$CODEMETER_DEB" >/dev/null 2>&1 || true
    ok "CodeMeter installed"
else
    warn "codemeter.deb not found at $CODEMETER_DEB — skipping (may cause license issues)"
fi

# =============================================================================
# STEP 3 — Install ABBYY FREngine12
# =============================================================================
step "3/8  Installing ABBYY FREngine12..."

if [ ! -d "$BUNDLE_DIR/FREngine12/Bin" ]; then
    err "abbyy_bundle/FREngine12 not found in $BUNDLE_DIR\n  Run: bash scripts/bundle_abbyy.sh  OR  extract the portable bundle"
fi

mkdir -p "$ABBYY_INSTALL_DIR"
cp -a "$BUNDLE_DIR/FREngine12/." "$ABBYY_INSTALL_DIR/"
chmod +x \
    "$ABBYY_INSTALL_DIR/Samples/CommandLineInterface/CommandLineInterface" \
    "$ABBYY_INSTALL_DIR/Bin/FREngineProcessor" \
    "$ABBYY_INSTALL_DIR/Bin/LicenseManager.Console"
find "$ABBYY_INSTALL_DIR" -name "*.sh" -exec chmod +x {} \; 2>/dev/null || true
rm -f "$ABBYY_INSTALL_DIR/Bin/libProtection.Developer.so" 2>/dev/null || true
ok "FREngine12 installed to $ABBYY_INSTALL_DIR"

# ── Install LicensingService ──────────────────────────────────────────────────
mkdir -p /usr/local/bin/ABBYY/SDK/12/Licensing
mkdir -p /usr/local/lib/ABBYY/SDK/12/Licensing
mkdir -p /var/lib/ABBYY/SDK/12/Licenses

cp "$BUNDLE_DIR/LicensingService" /usr/local/bin/ABBYY/SDK/12/Licensing/LicensingService
chmod +x /usr/local/bin/ABBYY/SDK/12/Licensing/LicensingService
cp -a "$BUNDLE_DIR/usr_local_lib_abbyy/." /usr/local/lib/ABBYY/
cp -a "$BUNDLE_DIR/var_lib_abbyy/ABBYY/." /var/lib/ABBYY/
chmod -R 777 /var/lib/ABBYY/
ok "LicensingService installed"

# ── Trust TR/Zscaler CA certs ─────────────────────────────────────────────────
if ls "$BUNDLE_DIR/tr_certs/"*.crt >/dev/null 2>&1; then
    cp "$BUNDLE_DIR/tr_certs/"*.crt /usr/local/share/ca-certificates/ 2>/dev/null || true
    update-ca-certificates >/dev/null 2>&1 || true
    ok "TR/Zscaler CA certs installed"
fi

# ── Run activatefre.sh to configure LicensingSettings.xml ────────────────────
ACTIVATE="$ABBYY_INSTALL_DIR/activatefre.sh"
if [ -x "$ACTIVATE" ]; then
    export RootScriptLaunched=1
    if [ -n "$ABBYY_LS_HOST" ]; then
        "$ACTIVATE" -- \
            --install-dir "$ABBYY_INSTALL_DIR" \
            --service-address "$ABBYY_LS_HOST" \
            --skip-local-service-installation \
            --skip-local-license-activation 2>/dev/null || true
    else
        # Local internet machine: LicensingService runs locally, connects to account.abbyy.com
        "$ACTIVATE" -- \
            --install-dir "$ABBYY_INSTALL_DIR" \
            --skip-local-service-installation \
            --skip-local-license-activation 2>/dev/null || true
    fi
    ok "LicensingSettings.xml configured"
fi

# =============================================================================
# STEP 4 — Install app to INSTALL_DIR
# =============================================================================
step "4/8  Installing SGML Pipeline app to $INSTALL_DIR..."
mkdir -p "$INSTALL_DIR"

# Copy app files
cp -a "$REPO_DIR/app"           "$INSTALL_DIR/"
cp -a "$REPO_DIR/pipeline"      "$INSTALL_DIR/"
cp -a "$REPO_DIR/streamlit_app.py"  "$INSTALL_DIR/"
cp -a "$REPO_DIR/abbyy_convert.py"  "$INSTALL_DIR/"
cp -a "$REPO_DIR/requirements.txt"  "$INSTALL_DIR/"
cp -a "$REPO_DIR/scripts/docker_start.sh" "$INSTALL_DIR/start_services.sh"
[ -d "$REPO_DIR/.streamlit" ]   && cp -a "$REPO_DIR/.streamlit" "$INSTALL_DIR/"
[ -d "$REPO_DIR/data" ]         && cp -a "$REPO_DIR/data"       "$INSTALL_DIR/"
[ -d "$REPO_DIR/pages" ]        && cp -a "$REPO_DIR/pages"      "$INSTALL_DIR/" 2>/dev/null || true
[ -d "$REPO_DIR/validator" ]    && cp -a "$REPO_DIR/validator"  "$INSTALL_DIR/" 2>/dev/null || true

mkdir -p "$INSTALL_DIR/tmp" /tmp/sgml_pipeline
ok "App files copied to $INSTALL_DIR"

# =============================================================================
# STEP 5 — Python virtual environment + dependencies
# =============================================================================
step "5/8  Creating Python virtual environment..."
python3.12 -m venv "$INSTALL_DIR/.venv"
source "$INSTALL_DIR/.venv/bin/activate"
pip3 install --quiet --no-cache-dir --timeout 300 -r "$INSTALL_DIR/requirements.txt"
ok "Python environment ready ($(python3 --version))"

# =============================================================================
# STEP 6 — Create launch script
# =============================================================================
step "6/8  Creating launch script..."

cat > /usr/local/bin/sgml-pipeline << LAUNCH
#!/usr/bin/env bash
# SGML Pipeline launcher
INSTALL_DIR="$INSTALL_DIR"
VENV="\$INSTALL_DIR/.venv"
export LD_LIBRARY_PATH="$ABBYY_INSTALL_DIR/Bin:/usr/local/lib/ABBYY/SDK/12/Licensing:\$LD_LIBRARY_PATH"
export ABBYY_CLI="$ABBYY_INSTALL_DIR/Samples/CommandLineInterface/CommandLineInterface"
export ABBYY_LIB="$ABBYY_INSTALL_DIR/Bin"
export TEMP_DIR="/tmp/sgml_pipeline"
export LANG=en_US.UTF-8

case "\${1:-start}" in
  start)
    # Start CodeMeterLin
    command -v CodeMeterLin >/dev/null 2>&1 && CodeMeterLin -f >/dev/null 2>&1 &
    sleep 2

    # Start LicensingService
    LS_BIN="/usr/local/bin/ABBYY/SDK/12/Licensing/LicensingService"
    if ! pgrep -f LicensingService >/dev/null 2>&1; then
        LD_LIBRARY_PATH=/usr/local/lib/ABBYY/SDK/12/Licensing "\$LS_BIN" /standalone &
        sleep 3
    fi

    echo "Starting SGML Pipeline on http://localhost:$APP_PORT ..."
    mkdir -p "\$TEMP_DIR"

    # Open browser in background (works on Ubuntu desktop)
    (sleep 4 && xdg-open "http://localhost:$APP_PORT" 2>/dev/null) &

    source "\$VENV/bin/activate"
    cd "\$INSTALL_DIR"
    exec streamlit run streamlit_app.py \
        --server.port=$APP_PORT \
        --server.address=0.0.0.0 \
        --server.headless=true
    ;;

  stop)
    pkill -f "streamlit run" 2>/dev/null && echo "Streamlit stopped" || echo "Not running"
    pkill -f LicensingService  2>/dev/null && echo "LicensingService stopped" || true
    ;;

  status)
    pgrep -a streamlit && echo "✅ Streamlit running" || echo "❌ Streamlit not running"
    pgrep -a LicensingService && echo "✅ LicensingService running" || echo "❌ LicensingService not running"
    ;;

  diag)
    source "\$VENV/bin/activate"
    cd "\$INSTALL_DIR"
    python3 abbyy_convert.py --diag
    ;;

  *)
    echo "Usage: sgml-pipeline {start|stop|status|diag}"
    ;;
esac
LAUNCH

chmod +x /usr/local/bin/sgml-pipeline
ok "Launch script: sgml-pipeline {start|stop|status|diag}"

# =============================================================================
# STEP 7 — Systemd service (auto-start on boot)
# =============================================================================
step "7/8  Registering systemd service..."

cat > /etc/systemd/system/sgml-pipeline.service << SVCFILE
[Unit]
Description=SGML Pipeline — PDF to DOCX/SGML Conversion
After=network.target

[Service]
Type=simple
User=root
WorkingDirectory=$INSTALL_DIR
Environment="LD_LIBRARY_PATH=$ABBYY_INSTALL_DIR/Bin:/usr/local/lib/ABBYY/SDK/12/Licensing"
Environment="ABBYY_CLI=$ABBYY_INSTALL_DIR/Samples/CommandLineInterface/CommandLineInterface"
Environment="ABBYY_LIB=$ABBYY_INSTALL_DIR/Bin"
Environment="TEMP_DIR=/tmp/sgml_pipeline"
Environment="LANG=en_US.UTF-8"
ExecStartPre=-/usr/sbin/CodeMeterLin -f
ExecStartPre=/bin/sleep 2
ExecStartPre=/bin/bash -c 'LD_LIBRARY_PATH=/usr/local/lib/ABBYY/SDK/12/Licensing /usr/local/bin/ABBYY/SDK/12/Licensing/LicensingService /standalone &'
ExecStartPre=/bin/sleep 3
ExecStart=$INSTALL_DIR/.venv/bin/streamlit run $INSTALL_DIR/streamlit_app.py --server.port=$APP_PORT --server.address=0.0.0.0 --server.headless=true
Restart=on-failure
RestartSec=10
StandardOutput=journal
StandardError=journal

[Install]
WantedBy=multi-user.target
SVCFILE

systemctl daemon-reload
systemctl enable sgml-pipeline >/dev/null 2>&1
ok "Systemd service registered (sgml-pipeline)"

# =============================================================================
# STEP 8 — Desktop shortcut (Ubuntu desktop)
# =============================================================================
step "8/8  Creating desktop shortcut..."

DESKTOP_DIR="/usr/share/applications"
cat > "$DESKTOP_DIR/sgml-pipeline.desktop" << DESKTOP
[Desktop Entry]
Name=SGML Pipeline
Comment=PDF to DOCX/SGML Conversion (ABBYY FREngine12)
Exec=/usr/local/bin/sgml-pipeline start
Icon=utilities-terminal
Terminal=false
Type=Application
Categories=Office;
StartupNotify=true
DESKTOP

ok "Desktop shortcut created"

# =============================================================================
# DONE
# =============================================================================
echo ""
echo "============================================================"
echo -e "${GREEN}  Installation Complete!${NC}"
echo "============================================================"
echo ""
echo "  Commands:"
echo "    sgml-pipeline start   → Start app (opens browser)"
echo "    sgml-pipeline stop    → Stop app"
echo "    sgml-pipeline status  → Check running services"
echo "    sgml-pipeline diag    → Run ABBYY diagnostics"
echo ""
echo "  Web UI:  http://localhost:$APP_PORT"
echo "  Logs:    journalctl -u sgml-pipeline -f"
echo ""
echo "  Starting now..."
echo ""
systemctl start sgml-pipeline 2>/dev/null || /usr/local/bin/sgml-pipeline start &
sleep 5
echo -e "  ${GREEN}✓ App is running at http://localhost:$APP_PORT${NC}"
echo ""
echo "  NOTE: ABBYY Online License requires internet access to"
echo "        account.abbyy.com:443 — ensure this is reachable."
echo "============================================================"
