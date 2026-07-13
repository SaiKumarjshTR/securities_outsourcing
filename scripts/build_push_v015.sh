#!/bin/bash
# ─────────────────────────────────────────────────────────────────────────────
# build_push_v015.sh — Build and push sgml-pipeline-prod:0.0.20
#
# v0.0.20: Added missing dependencies to requirements.txt:
#          chromadb>=0.5.0, fastapi>=0.111.0, flask>=3.0.0
#          Fixed abbyy_convert.py to use /proc/net/tcp fallback (no ss needed)
#          Added iproute2 to apt packages
#
# Usage:
#   bash scripts/build_push_v015.sh [JFROG_USER] [JFROG_TOKEN]
#
# AWS credentials must already be active (run cloud-tool --region us-east-1 login).
# ─────────────────────────────────────────────────────────────────────────────
set -e

JFROG_USER="${1:-C303180}"
JFROG_TOKEN="${2:-}"
VERSION="0.0.20"
IMAGE_NAME="sgml-pipeline-prod"
ECR="127288631409.dkr.ecr.us-east-1.amazonaws.com/a207870-ml-model-registry-model-registry-prod-use1"
ECR_HOST="127288631409.dkr.ecr.us-east-1.amazonaws.com"

cd /home/securities_outsourcing

echo "================================================"
echo " SGML Pipeline v${VERSION} — Direct ABBYY CLI"
echo "================================================"
echo ""

# ── Step 1: ECR login ─────────────────────────────────────────────────────────
echo "=== STEP 1/3: ECR login ==="
aws --profile tr-aiml-hackathon-prod ecr get-login-password --region us-east-1 | \
    docker login --username AWS --password-stdin "$ECR_HOST"
echo ""

# ── Step 2: Build ─────────────────────────────────────────────────────────────
echo "=== STEP 2/3: Building ${IMAGE_NAME}:${VERSION} ==="
echo "Log: /tmp/build_v015.log"
echo ""

docker build \
    --build-arg TR_JFROG_USERNAME="$JFROG_USER" \
    --build-arg TR_JFROG_TOKEN="$JFROG_TOKEN" \
    -t "${IMAGE_NAME}:${VERSION}" \
    . 2>&1 | tee /tmp/build_v015.log

echo "Build exit: ${PIPESTATUS[0]}"
echo ""

# ── Step 3: Tag and push ──────────────────────────────────────────────────────
echo "=== STEP 3/3: Tagging and pushing to ECR ==="
ECR_TAG="${ECR}:${IMAGE_NAME}-${VERSION}"
docker tag "${IMAGE_NAME}:${VERSION}" "$ECR_TAG"
echo "Pushing $ECR_TAG ..."
docker push "$ECR_TAG" 2>&1 | tee /tmp/push_v015.log

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
