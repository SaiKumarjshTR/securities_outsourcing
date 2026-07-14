#!/usr/bin/env bash
# =============================================================================
# create_portable_bundle.sh
#
# Creates a portable, self-contained distributable bundle:
#   sgml-pipeline-bundle.tar.gz    (for Ubuntu / WSL2 direct install)
#   sgml-pipeline-ubuntu.run      (self-extracting single .run for Ubuntu)
#
# WHAT'S INCLUDED:
#   • Python application code (app/, pipeline/, validator/, pages/, streamlit_app.py, etc.)
#   • ABBYY FREngine12 binaries + LicensingService (from abbyy_bundle/)
#   • requirements.txt + setup scripts
#
# WHAT'S NOT INCLUDED (installed at runtime):
#   • Python 3.12 (system package)
#   • pip packages (downloaded via pip from PyPI)
#   • System libs (apt packages installed by setup_local_ubuntu.sh)
#
# USAGE:
#   bash scripts/create_portable_bundle.sh
#
# OUTPUT:
#   dist/sgml-pipeline-bundle.tar.gz   → Extract and run setup_local_ubuntu.sh
#   dist/sgml-pipeline-ubuntu.run      → Run directly on Ubuntu
#
# DISTRIBUTION:
#   - Ubuntu:  scp sgml-pipeline-ubuntu.run user@machine:~ && bash ~/sgml-pipeline-ubuntu.run
#   - Windows: Copy sgml-pipeline-bundle.tar.gz + setup_windows_wsl.ps1
#              Run setup_windows_wsl.ps1 as Administrator
# =============================================================================
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
DIST_DIR="$REPO_DIR/dist"
BUNDLE_NAME="sgml-pipeline-bundle"
VERSION="${VERSION:-$(date +%Y%m%d)}"

GREEN='\033[0;32m'; YELLOW='\033[1;33m'; RED='\033[0;31m'; NC='\033[0m'
ok()   { echo -e "${GREEN}  ✓ $*${NC}"; }
warn() { echo -e "${YELLOW}  ⚠ $*${NC}"; }
step() { echo -e "\n${YELLOW}[$*]${NC}"; }

echo ""
echo "============================================================"
echo "  SGML Pipeline — Portable Bundle Creator"
echo "  Repo : $REPO_DIR"
echo "  Dist : $DIST_DIR"
echo "  Ver  : $VERSION"
echo "============================================================"

mkdir -p "$DIST_DIR"
TMP="$DIST_DIR/.bundle-tmp"
BUNDLE_ROOT="$TMP/$BUNDLE_NAME"
rm -rf "$TMP"
mkdir -p "$BUNDLE_ROOT"

# =============================================================================
# CHECK ABBYY bundle exists
# =============================================================================
step "Checking abbyy_bundle..."
if [ ! -d "$REPO_DIR/abbyy_bundle/FREngine12/Bin" ]; then
    echo -e "${RED}ERROR: abbyy_bundle/FREngine12/Bin not found.${NC}"
    echo "Run: bash scripts/bundle_abbyy.sh   (or ensure abbyy_bundle/ is complete)"
    exit 1
fi
BUNDLE_SIZE_GB=$(du -sh "$REPO_DIR/abbyy_bundle" 2>/dev/null | awk '{print $1}')
ok "abbyy_bundle found (~$BUNDLE_SIZE_GB)"

# =============================================================================
# STEP 1 — Copy app code
# =============================================================================
step "1/4  Copying application code..."

for d in app pipeline validator pages data; do
    [ -d "$REPO_DIR/$d" ] && cp -a "$REPO_DIR/$d" "$BUNDLE_ROOT/"
done

for f in streamlit_app.py abbyy_convert.py requirements.txt hitl_review.py excel_hitl.py config.py validator_app.py entities_list.txt; do
    [ -f "$REPO_DIR/$f" ] && cp "$REPO_DIR/$f" "$BUNDLE_ROOT/"
done

[ -d "$REPO_DIR/.streamlit" ] && cp -a "$REPO_DIR/.streamlit" "$BUNDLE_ROOT/"
cp -a "$REPO_DIR/scripts" "$BUNDLE_ROOT/"

ok "App code copied"

# =============================================================================
# STEP 2 — Copy ABBYY bundle (the big part — ~3.3GB)
# =============================================================================
step "2/4  Copying ABBYY FREngine12 bundle (~3.3GB)..."
echo "  This may take a few minutes..."
cp -a "$REPO_DIR/abbyy_bundle" "$BUNDLE_ROOT/"
ok "ABBYY bundle copied"

# =============================================================================
# STEP 3 — Create a bundle manifest
# =============================================================================
step "3/4  Writing bundle info..."

cat > "$BUNDLE_ROOT/BUNDLE_INFO.txt" << INFO
SGML Pipeline Bundle
Version     : $VERSION
Created     : $(date -u '+%Y-%m-%d %H:%M:%S UTC')
Build host  : $(hostname)
ABBYY ver   : FREngine12 (Linux x86-64)
License type: Online (requires internet access to account.abbyy.com:443)

CONTENTS:
  app/                 Application code (FastAPI backend)
  pipeline/            ABBYY conversion pipeline
  validator/           SGML validator
  pages/               Streamlit pages
  streamlit_app.py     Main Streamlit UI
  abbyy_bundle/        ABBYY FREngine12 + LicensingService
  requirements.txt     Python dependencies
  scripts/             Setup + launcher scripts

INSTALL:
  Ubuntu (direct):
    sudo bash scripts/setup_local_ubuntu.sh

  Windows (via WSL2, run as Administrator):
    Right-click scripts\setup_windows_wsl.ps1 → Run with PowerShell

  Manual (advanced):
    Extraction only: tar -xzf sgml-pipeline-bundle.tar.gz
    Then run setup_local_ubuntu.sh

ACCESS:
  http://localhost:8501

NOTES:
  - ABBYY Online License requires internet to account.abbyy.com:443
  - Works on any machine with internet (business laptops, workstations)
  - Tested on Ubuntu 20.04, 22.04, 24.04 and Windows 11 WSL2
INFO

ok "Bundle info written"

# =============================================================================
# STEP 4 — Create compressed archives
# =============================================================================
step "4/4  Creating archives..."

cd "$TMP"
# Remove any existing machine-specific ABBYY license tokens before packaging.
# Business users must activate their own license — our tokens are machine-bound
# and will NOT work on other machines. Keeping them would only confuse users.
find "$BUNDLE_NAME/abbyy_bundle/" -name "*.ActivationToken" -delete 2>/dev/null || true
find "$BUNDLE_NAME/abbyy_bundle/" -name "*.ActivationToken.bak" -delete 2>/dev/null || true
# Remove Python caches (not needed in distribution)
find "$BUNDLE_NAME/" -name '__pycache__' -type d -exec rm -rf {} + 2>/dev/null || true
find "$BUNDLE_NAME/" -name '*.pyc' -delete 2>/dev/null || true
TARBALL="$DIST_DIR/${BUNDLE_NAME}-${VERSION}.tar.gz"
echo "  Creating $TARBALL ..."
tar -czf "$TARBALL" "$BUNDLE_NAME/"
TARBALL_SIZE=$(du -sh "$TARBALL" | awk '{print $1}')
ok "Created: $TARBALL ($TARBALL_SIZE)"

# Also create a "latest" symlink
ln -sf "${BUNDLE_NAME}-${VERSION}.tar.gz" "$DIST_DIR/${BUNDLE_NAME}.tar.gz"
ok "Symlink: $DIST_DIR/${BUNDLE_NAME}.tar.gz → ${BUNDLE_NAME}-${VERSION}.tar.gz"

# ── Create self-extracting .run for Ubuntu ────────────────────────────────────
RUN_FILE="$DIST_DIR/${BUNDLE_NAME}-${VERSION}.run"
echo "  Creating self-extracting Ubuntu installer: $RUN_FILE ..."

# Header script that gets prepended to the base64-encoded tarball
cat > "$RUN_FILE" << 'RUNHEADER'
#!/usr/bin/env bash
# =============================================================================
# SGML Pipeline Self-Extracting Installer for Ubuntu
# Usage: bash sgml-pipeline-ubuntu.run [--extract-only] [--install-dir DIR]
# =============================================================================
set -euo pipefail

INSTALL_DIR="${SGML_INSTALL_DIR:-}"
EXTRACT_ONLY=0

for arg in "$@"; do
    case "$arg" in
        --extract-only) EXTRACT_ONLY=1 ;;
        --install-dir=*) INSTALL_DIR="${arg#*=}" ;;
    esac
done

# Find the line where the tarball data starts (after TARBALL_BEGIN marker)
TARBALL_LINE=$(awk '/^TARBALL_BEGIN$/{print NR+1; exit}' "$0")
EXTRACT_DIR="$(mktemp -d /tmp/sgml-install.XXXXXX)"

echo ""
echo "  SGML Pipeline Self-Extracting Installer"
echo "  Extracting to $EXTRACT_DIR ..."

# Decode and extract
tail -n +"$TARBALL_LINE" "$0" | base64 -d | tar -xz -C "$EXTRACT_DIR"

if [ "$EXTRACT_ONLY" -eq 1 ]; then
    echo "  Extracted to $EXTRACT_DIR"
    exit 0
fi

# Run setup
export INSTALL_DIR="${INSTALL_DIR:-/opt/sgml-pipeline}"
[[ "$EUID" -ne 0 ]] && SUDO="sudo" || SUDO=""
$SUDO bash "$EXTRACT_DIR/sgml-pipeline-bundle/scripts/setup_local_ubuntu.sh"
rm -rf "$EXTRACT_DIR"
exit 0

TARBALL_BEGIN
RUNHEADER

# Append base64-encoded tarball
base64 "$TARBALL" >> "$RUN_FILE"
chmod +x "$RUN_FILE"
RUN_SIZE=$(du -sh "$RUN_FILE" | awk '{print $1}')
ok "Created: $RUN_FILE ($RUN_SIZE)"

# Also create "latest" .run symlink
ln -sf "${BUNDLE_NAME}-${VERSION}.run" "$DIST_DIR/${BUNDLE_NAME}-ubuntu.run"
ok "Symlink: $DIST_DIR/${BUNDLE_NAME}-ubuntu.run → ${BUNDLE_NAME}-${VERSION}.run"

rm -rf "$TMP"

# =============================================================================
# Summary
# =============================================================================
echo ""
echo "============================================================"
echo -e "${GREEN}  Bundle creation complete!${NC}"
echo "============================================================"
echo ""
echo "  Output files:"
echo "    $DIST_DIR/${BUNDLE_NAME}-${VERSION}.tar.gz  ($TARBALL_SIZE) — for Windows/WSL2"
echo "    $DIST_DIR/${BUNDLE_NAME}-${VERSION}.run               — Ubuntu self-extracting"
echo ""
echo "  DISTRIBUTION GUIDE:"
echo ""
echo "  Ubuntu user:"
echo "    1. Copy ${BUNDLE_NAME}-ubuntu.run to the user's machine"
echo "    2. bash sgml-pipeline-bundle-ubuntu.run"
echo "    3. Access http://localhost:8501"
echo ""
echo "  Windows user:"
echo "    1. Copy these files to the user's Windows machine:"
echo "         ${BUNDLE_NAME}.tar.gz"
echo "         scripts/setup_windows_wsl.ps1"
echo "    2. Right-click setup_windows_wsl.ps1 → Run with PowerShell (as Admin)"
echo "    3. Desktop shortcut 'SGML Pipeline' is created automatically"
echo ""
echo "  IMPORTANT: Both paths require internet access for ABBYY license."
echo "             Business user machines have internet — this WILL work."
echo "============================================================"
