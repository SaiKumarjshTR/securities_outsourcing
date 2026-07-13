#!/bin/bash
# ─────────────────────────────────────────────────────────────────────────────
# build_push_v012.sh — Build and push sgml-pipeline-prod:0.0.12
#
# v0.0.12: ABBYY FREngine 12 bundled inside the image.
#          No external bridge server. No Windows port proxy. Fully self-contained.
#
# Usage:
#   bash scripts/build_push_v012.sh <JFROG_USER> <JFROG_TOKEN>
#
# AWS credentials must already be active (run cloud-tool --region us-east-1 login).
# ─────────────────────────────────────────────────────────────────────────────
set -e

JFROG_USER="${1:-C303180}"
JFROG_TOKEN="${2:-}"
VERSION="0.0.12"
IMAGE_NAME="sgml-pipeline-prod"
ECR="127288631409.dkr.ecr.us-east-1.amazonaws.com/a207870-ml-model-registry-model-registry-prod-use1"
ECR_HOST="127288631409.dkr.ecr.us-east-1.amazonaws.com"

cd /home/securities_outsourcing

echo "================================================"
echo " SGML Pipeline v${VERSION} — ABBYY Bundled Build"
echo "================================================"
echo ""

# ── Step 1: Bundle ABBYY into build context ──────────────────────────────────
echo "=== STEP 1/5: Bundle ABBYY (copies ~3.3 GB from /opt/ABBYY) ==="
bash scripts/bundle_abbyy.sh
echo ""

# ── Step 2: AWS login ─────────────────────────────────────────────────────────
echo "=== STEP 2/5: AWS login ==="
echo "Logging in via cloud-tool (requires MGMT credentials interactively)..."
/root/cloud-tool-env/bin/cloud-tool --region us-east-1 login
echo ""

# ── Step 3: ECR docker login ──────────────────────────────────────────────────
echo "=== STEP 3/5: ECR login ==="
aws ecr get-login-password --region us-east-1 | \
    podman login --username AWS --password-stdin "$ECR_HOST"
echo ""

# ── Step 4: Build ─────────────────────────────────────────────────────────────
echo "=== STEP 4/5: Building ${IMAGE_NAME}:${VERSION} ==="
echo "NOTE: Build context is ~3.5 GB (ABBYY bundle). This may take 10-20 minutes."
echo "Log: /tmp/build_v012.log"
echo ""

podman build --no-cache \
    --build-arg TR_JFROG_USERNAME="$JFROG_USER" \
    --build-arg TR_JFROG_TOKEN="$JFROG_TOKEN" \
    -t "${IMAGE_NAME}:${VERSION}" \
    . 2>&1 | tee /tmp/build_v012.log

echo ""
echo "Build exit: ${PIPESTATUS[0]}"

# ── Step 5: Tag and push ──────────────────────────────────────────────────────
echo "=== STEP 5/5: Tagging and pushing to ECR ==="
ECR_TAG="${ECR}:${IMAGE_NAME}-${VERSION}"
podman tag "${IMAGE_NAME}:${VERSION}" "$ECR_TAG"
echo "Pushing $ECR_TAG ..."
podman push "$ECR_TAG" --tls-verify=false 2>&1 | tee /tmp/push_v012.log

echo ""
echo "================================================"
echo " SUCCESS — Image pushed to ECR"
echo " Tag: ${IMAGE_NAME}-${VERSION}"
echo " ECR: $ECR_TAG"
echo ""
echo " Plexus next steps:"
echo "  1. Model Registry → sgml-pipeline-prod → Add Version 0.0.12"
echo "     Image: $ECR_TAG"
echo "     Health: /_stcore/health  Port: 8501"
echo "  2. Deployment → Securities Commission Conversion → Activate v0.0.12"
echo ""
echo " No Windows port proxy needed. ABBYY runs inside the container."
echo "================================================"
