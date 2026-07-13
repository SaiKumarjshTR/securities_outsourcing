#!/bin/bash
# ─────────────────────────────────────────────────────────────────────────────
# build_push_v016.sh — Build and push sgml-pipeline-prod:0.0.21
#
# v0.0.21: FIX for external License Server (Option B / ABBYY_LS_HOST).
#          docker_start.sh now patches ServerAddress in ALL LicensingSettings.xml
#          copies the FREngine CLI actually reads:
#            - /opt/ABBYY/FREngine12/Bin/LicensingSettings.xml
#            - /opt/ABBYY/FREngine12/CommonBin/Licensing/LicensingSettings.xml
#            - /usr/local/lib/ABBYY/SDK/12/Licensing/LicensingSettings.xml
#          Previously only the /usr/local/lib copy was patched, so the engine
#          kept using 127.0.0.1 and conversions failed even with a working
#          EC2 License Server.
#
# Deploy in Plexus with env var:  ABBYY_LS_HOST=<ec2-private-ip>
#
# Usage:
#   bash scripts/build_push_v016.sh [JFROG_USER] [JFROG_TOKEN]
#
# AWS credentials must already be active (run cloud-tool --region us-east-1 login).
# ─────────────────────────────────────────────────────────────────────────────
set -e

JFROG_USER="${1:-C303180}"
JFROG_TOKEN="${2:-}"
VERSION="0.0.21"
IMAGE_NAME="sgml-pipeline-prod"
ECR="127288631409.dkr.ecr.us-east-1.amazonaws.com/a207870-ml-model-registry-model-registry-prod-use1"
ECR_HOST="127288631409.dkr.ecr.us-east-1.amazonaws.com"

cd /home/securities_outsourcing

echo "================================================"
echo " SGML Pipeline v${VERSION} — External LS fix"
echo "================================================"
echo ""

# ── Pre-flight: ABBYY bundle must exist ──────────────────────────────────────
if [ ! -d abbyy_bundle/FREngine12 ] || [ ! -f abbyy_bundle/LicensingService ]; then
    echo "ERROR: abbyy_bundle/ is incomplete. Run: bash scripts/bundle_abbyy.sh"
    exit 1
fi

# ── Step 1: ECR login ─────────────────────────────────────────────────────────
echo "=== STEP 1/3: ECR login ==="
aws --profile tr-aiml-hackathon-prod ecr get-login-password --region us-east-1 | \
    docker login --username AWS --password-stdin "$ECR_HOST"
echo ""

# ── Step 2: Build ─────────────────────────────────────────────────────────────
echo "=== STEP 2/3: Building ${IMAGE_NAME}:${VERSION} ==="
echo "Log: /tmp/build_v016.log"
echo ""

docker build \
    --build-arg TR_JFROG_USERNAME="$JFROG_USER" \
    --build-arg TR_JFROG_TOKEN="$JFROG_TOKEN" \
    -t "${IMAGE_NAME}:${VERSION}" \
    . 2>&1 | tee /tmp/build_v016.log

echo "Build exit: ${PIPESTATUS[0]}"
echo ""

# ── Step 3: Tag and push ──────────────────────────────────────────────────────
echo "=== STEP 3/3: Tagging and pushing to ECR ==="
ECR_TAG="${ECR}:${IMAGE_NAME}-${VERSION}"
docker tag "${IMAGE_NAME}:${VERSION}" "$ECR_TAG"
echo "Pushing $ECR_TAG ..."
docker push "$ECR_TAG" 2>&1 | tee /tmp/push_v016.log

echo ""
echo "================================================"
echo " SUCCESS — v${VERSION} pushed to ECR"
echo " ECR: $ECR_TAG"
echo ""
echo " Plexus next steps:"
echo "  1. Model Registry → sgml-pipeline-prod → Add Version ${VERSION}"
echo "     Image: $ECR_TAG"
echo "     Health: /_stcore/health  Port: 8501"
echo "  2. Deployment → set env var ABBYY_LS_HOST=<ec2-private-ip>"
echo "  3. Activate v${VERSION}"
echo "================================================"
