#!/bin/bash
# ─────────────────────────────────────────────────────────────────────────────
# build_push_v013.sh — Build and push sgml-pipeline-prod:0.0.13
#
# v0.0.13: ABBYY CLI called directly from Python (no HTTP bridge).
#          procps + libusb-1.0-0 added so CodeMeter installs cleanly.
#          Startup: CodeMeterLin → LicensingService → Streamlit (3 steps only).
#
# Usage:
#   bash scripts/build_push_v013.sh [JFROG_USER] [JFROG_TOKEN]
#
# AWS credentials must already be active (run cloud-tool --region us-east-1 login).
# ─────────────────────────────────────────────────────────────────────────────
set -e

JFROG_USER="${1:-C303180}"
JFROG_TOKEN="${2:-}"
VERSION="0.0.13"
IMAGE_NAME="sgml-pipeline-prod"
ECR="127288631409.dkr.ecr.us-east-1.amazonaws.com/a207870-ml-model-registry-model-registry-prod-use1"
ECR_HOST="127288631409.dkr.ecr.us-east-1.amazonaws.com"

cd /home/securities_outsourcing

echo "================================================"
echo " SGML Pipeline v${VERSION} — Direct ABBYY CLI"
echo "================================================"
echo ""

# ── Step 1: Bundle ABBYY into build context ──────────────────────────────────
echo "=== STEP 1/4: Bundle ABBYY (copies ~3.3 GB from /opt/ABBYY) ==="
bash scripts/bundle_abbyy.sh
echo ""

# ── Step 2: ECR login (credentials assumed already active) ───────────────────
echo "=== STEP 2/4: ECR login ==="
aws --profile tr-aiml-hackathon-prod ecr get-login-password --region us-east-1 | \
    podman login --username AWS --password-stdin "$ECR_HOST" --tls-verify=false
echo ""

# ── Step 3: Build ─────────────────────────────────────────────────────────────
echo "=== STEP 3/4: Building ${IMAGE_NAME}:${VERSION} ==="
echo "Log: /tmp/build_v013.log"
echo ""

podman build --no-cache \
    --build-arg TR_JFROG_USERNAME="$JFROG_USER" \
    --build-arg TR_JFROG_TOKEN="$JFROG_TOKEN" \
    -t "${IMAGE_NAME}:${VERSION}" \
    . 2>&1 | tee /tmp/build_v013.log

echo "Build exit: ${PIPESTATUS[0]}"
echo ""

# ── Step 4: Tag and push ──────────────────────────────────────────────────────
echo "=== STEP 4/4: Tagging and pushing to ECR ==="
ECR_TAG="${ECR}:${IMAGE_NAME}-${VERSION}"
podman tag "${IMAGE_NAME}:${VERSION}" "$ECR_TAG"
echo "Pushing $ECR_TAG ..."
podman push "$ECR_TAG" --tls-verify=false 2>&1 | tee /tmp/push_v013.log

echo ""
echo "================================================"
echo " SUCCESS — v${VERSION} pushed to ECR"
echo " ECR: $ECR_TAG"
echo ""
echo " Plexus next steps:"
echo "  1. Model Registry → sgml-pipeline-prod → Add Version ${VERSION}"
echo "     Image: $ECR_TAG"
echo "     Health: /_stcore/health  Port: 8501"
echo "  2. Deployment → Securities Commission Conversion → Activate v${VERSION}"
echo "================================================"
