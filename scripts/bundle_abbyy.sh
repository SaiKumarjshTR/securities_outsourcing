#!/bin/bash
# ─────────────────────────────────────────────────────────────────────────────
# bundle_abbyy.sh
# Copies the local ABBYY FREngine 12 installation into abbyy_bundle/ so the
# Dockerfile can COPY it into the Docker image at build time.
#
# Run once before each Docker build:
#   bash scripts/bundle_abbyy.sh
#
# The abbyy_bundle/ directory is .gitignored (contains licensed binaries).
# ─────────────────────────────────────────────────────────────────────────────
set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"
BUNDLE_DIR="$PROJECT_DIR/abbyy_bundle"

echo "=== ABBYY FREngine 12 Bundle Script ==="
echo "  Project : $PROJECT_DIR"
echo "  Bundle  : $BUNDLE_DIR"
echo ""

# Validate source paths
for path in \
    "/opt/ABBYY/FREngine12" \
    "/usr/local/bin/ABBYY/SDK/12/Licensing/LicensingService" \
    "/var/lib/ABBYY/SDK/12/Licenses" \
    "/opt/ABBYY/FREngine12/CodeMeter/codemeter_7.30.4811.500_amd64.deb"; do
    if [ ! -e "$path" ]; then
        echo "ERROR: Required path not found: $path"
        echo "       Run this script on the Ubuntu host where ABBYY is installed."
        exit 1
    fi
done

# Clean and recreate bundle directory
echo "[prep] Cleaning existing bundle..."
rm -rf "$BUNDLE_DIR"
mkdir -p "$BUNDLE_DIR"

# 1. FREngine12 (~3.3 GB — the full SDK including all .so libs and CLI)
echo ""
echo "[1/4] Copying FREngine12 (~3.3 GB, takes 2-5 minutes)..."
cp -r /opt/ABBYY/FREngine12 "$BUNDLE_DIR/FREngine12"
echo "      Done: $(du -sh "$BUNDLE_DIR/FREngine12" | cut -f1)"

# 2. CodeMeter deb package (license manager daemon)
echo ""
echo "[2/4] Copying CodeMeter package..."
cp /opt/ABBYY/FREngine12/CodeMeter/codemeter_7.30.4811.500_amd64.deb \
   "$BUNDLE_DIR/codemeter.deb"
echo "      Done: $(du -sh "$BUNDLE_DIR/codemeter.deb" | cut -f1)"

# 3. LicensingService binary
echo ""
echo "[3/4] Copying ABBYY LicensingService binary..."
cp /usr/local/bin/ABBYY/SDK/12/Licensing/LicensingService \
   "$BUNDLE_DIR/LicensingService"
chmod +x "$BUNDLE_DIR/LicensingService"
echo "      Done"

# 4. Full ABBYY license state (activation tokens + activated Protection.ccf + counters)
# Protection.ccf is the encrypted activated license — required for LicensingService to start.
echo ""
echo "[4/4] Copying full ABBYY license state (/var/lib/ABBYY/)..."
mkdir -p "$BUNDLE_DIR/var_lib_abbyy"
# Use -T (--no-target-directory) so cp copies CONTENTS of ABBYY/ into var_lib_abbyy/
# without creating an extra ABBYY/ subdirectory level
cp -rT /var/lib/ABBYY/ "$BUNDLE_DIR/var_lib_abbyy/"
echo "      Done: $(du -sh "$BUNDLE_DIR/var_lib_abbyy" | cut -f1)"

echo ""
echo "=== Bundle complete ==="
du -sh "$BUNDLE_DIR"/*
echo ""
echo "Total: $(du -sh "$BUNDLE_DIR" | cut -f1)"
echo ""
echo "Next: podman build (via scripts/build_push_v014.sh)"
