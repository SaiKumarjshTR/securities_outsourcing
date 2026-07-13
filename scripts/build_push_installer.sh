#!/bin/bash
# ─────────────────────────────────────────────────────────────────────────────
# build_push_installer.sh — Build and push using Dockerfile.installer
#
# This uses ABBYY's official installer approach (activatefre.sh) rather than
# manually copying files. LicensingSettings.xml is configured at BUILD TIME
# when SERVICE_ADDRESS is provided — no runtime XML patching required.
#
# Prerequisites:
#   Either provide the official FRE*.sh from ABBYY portal, OR run:
#     bash scripts/create_fre_bundle_installer.sh
#   to create FRE_bundle.sh from the existing bundle.
#
# Usage:
#   # Mode A — Fixed LS address baked into image (recommended for Plexus):
#   bash scripts/build_push_installer.sh \
#     C303180 <jfrog_token> sgml-abbyy-ls.default.svc.cluster.local
#
#   # Mode B — Dynamic LS address (set ABBYY_LS_HOST env var at deployment):
#   bash scripts/build_push_installer.sh C303180 <jfrog_token>
#
#   # With LICENSE_PASSWORD (to activate online license during build):
#   bash scripts/build_push_installer.sh C303180 <jfrog_token> "" <license_pwd>
# ─────────────────────────────────────────────────────────────────────────────
set -e

JFROG_USER="${1:-C303180}"
JFROG_TOKEN="${2:-}"
SERVICE_ADDRESS="${3:-}"    # Optional: hard-code LS address at build time
LICENSE_PASSWORD="${4:-}"   # Optional: needed if activating online license at build

VERSION="0.0.22-installer"
IMAGE_NAME="sgml-pipeline-prod"
ECR="127288631409.dkr.ecr.us-east-1.amazonaws.com/a207870-ml-model-registry-model-registry-prod-use1"
ECR_HOST="127288631409.dkr.ecr.us-east-1.amazonaws.com"

cd /home/securities_outsourcing

echo "================================================"
echo " SGML Pipeline ${VERSION} — Installer-based"
echo " (activatefre.sh build-time LS configuration)"
echo "================================================"
echo ""

if [ -n "$SERVICE_ADDRESS" ]; then
    echo " Mode A: SERVICE_ADDRESS=$SERVICE_ADDRESS (baked in)"
else
    echo " Mode B: No SERVICE_ADDRESS — runtime ABBYY_LS_HOST required"
fi
echo ""

# ── Pre-flight checks ─────────────────────────────────────────────────────────
if [ ! -f Dockerfile.installer ]; then
    echo "ERROR: Dockerfile.installer not found."
    exit 1
fi

if [ ! -d abbyy_bundle/FREngine12/Bin ]; then
    echo "ERROR: abbyy_bundle/FREngine12 is missing."
    echo "       Run: bash scripts/bundle_abbyy.sh"
    exit 1
fi

echo "[pre-flight] Checking for FRE installer..."
if [ -f FRE12.sh ]; then
    echo "[pre-flight] ✓ Official FRE12.sh found — will use ABBYY Option 1 in Dockerfile.installer"
    echo "             NOTE: Ensure Option 1 is uncommented in Dockerfile.installer"
elif [ -f FRE_bundle.sh ]; then
    echo "[pre-flight] ✓ FRE_bundle.sh found (created from existing bundle)"
    echo "             NOTE: Ensure Option 1 is uncommented in Dockerfile.installer"
else
    echo "[pre-flight] No FRE*.sh found — using Option 2 (bundle copy + activatefre.sh)"
    echo "             To create FRE_bundle.sh: bash scripts/create_fre_bundle_installer.sh"
fi

# ── ECR login ─────────────────────────────────────────────────────────────────
echo ""
echo "=== STEP 1/3: ECR login ==="
aws --profile tr-aiml-hackathon-prod ecr get-login-password --region us-east-1 | \
    docker login --username AWS --password-stdin "$ECR_HOST"

# ── Build ─────────────────────────────────────────────────────────────────────
echo ""
echo "=== STEP 2/3: Building ${IMAGE_NAME}:${VERSION} ==="
echo "Log: /tmp/build_installer.log"

BUILD_ARGS=(
    --build-arg "TR_JFROG_USERNAME=$JFROG_USER"
    --build-arg "TR_JFROG_TOKEN=$JFROG_TOKEN"
)
[ -n "$SERVICE_ADDRESS" ] && BUILD_ARGS+=(--build-arg "SERVICE_ADDRESS=$SERVICE_ADDRESS")
[ -n "$LICENSE_PASSWORD" ] && BUILD_ARGS+=(--build-arg "LICENSE_PASSWORD=$LICENSE_PASSWORD")

docker build \
    -f Dockerfile.installer \
    "${BUILD_ARGS[@]}" \
    -t "${IMAGE_NAME}:${VERSION}" \
    . 2>&1 | tee /tmp/build_installer.log

echo "Build exit: ${PIPESTATUS[0]}"

# ── Verify LicensingSettings.xml was configured correctly ─────────────────────
echo ""
echo "=== Verifying LicensingSettings.xml in built image ==="
docker run --rm "${IMAGE_NAME}:${VERSION}" \
    bash -c 'echo "=== Bin/LicensingSettings.xml ===" && \
             cat /opt/ABBYY/FREngine12/Bin/LicensingSettings.xml && \
             echo "" && \
             echo "=== CommonBin/Licensing/LicensingSettings.xml ===" && \
             cat /opt/ABBYY/FREngine12/CommonBin/Licensing/LicensingSettings.xml 2>/dev/null || true'

# ── Tag and push ──────────────────────────────────────────────────────────────
echo ""
echo "=== STEP 3/3: Tagging and pushing to ECR ==="
ECR_TAG="${ECR}:${IMAGE_NAME}-${VERSION}"
docker tag "${IMAGE_NAME}:${VERSION}" "$ECR_TAG"
docker push "$ECR_TAG" 2>&1 | tee /tmp/push_installer.log

echo ""
echo "================================================"
echo " SUCCESS — ${VERSION} pushed"
echo " ECR: $ECR_TAG"
if [ -n "$SERVICE_ADDRESS" ]; then
    echo ""
    echo " Mode A: LicensingService address is BAKED IN:"
    echo "   ServerAddress = $SERVICE_ADDRESS"
    echo "   No ABBYY_LS_HOST env var needed in Plexus!"
    echo "   (docker_start.sh patching is a no-op when already set)"
else
    echo ""
    echo " Mode B: Set ABBYY_LS_HOST in Plexus deployment:"
    echo "   ABBYY_LS_HOST=sgml-abbyy-ls.default.svc.cluster.local"
fi
echo ""
echo " CRITICAL: Also set /dev/shm ≥ 1GB in Plexus pod spec:"
echo "   volumes: [{name: dshm, emptyDir: {medium: Memory, sizeLimit: 1Gi}}]"
echo "   volumeMounts: [{mountPath: /dev/shm, name: dshm}]"
echo "================================================"
