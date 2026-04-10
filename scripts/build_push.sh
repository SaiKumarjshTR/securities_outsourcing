#!/usr/bin/env bash
# =============================================================================
# build_push.sh — Build Docker image and push to TR ECR
# Usage: bash build_push.sh <employee_id> <jfrog_token>
# Run from: sgml-pipeline-deployment/ directory in WSL
# =============================================================================
set -euo pipefail

JFROG_USER="${1:-}"
JFROG_TOKEN="${2:-}"

if [[ -z "$JFROG_USER" || -z "$JFROG_TOKEN" ]]; then
  echo "Usage: bash scripts/build_push.sh <employee_id> <jfrog_token>"
  exit 1
fi

# ── Configuration (edit if needed) ───────────────────────────────────────────
MODEL_NAME="sgml-pipeline-prod"
VERSION="0.0.1"
AWS_REGION="us-east-1"
ECR_REGISTRY="127288631409.dkr.ecr.us-east-1.amazonaws.com"
ECR_REPO="a207870-ml-model-registry-model-registry-prod-use1"
FULL_IMAGE_TAG="${ECR_REGISTRY}/${ECR_REPO}:${MODEL_NAME}-${VERSION}"

echo ""
echo "========================================================"
echo "  SGML Pipeline — Docker Build & Push"
echo "  Model : ${MODEL_NAME}"
echo "  Version: ${VERSION}"
echo "  Target : ${FULL_IMAGE_TAG}"
echo "========================================================"
echo ""

# ── Step 1: ECR Login ─────────────────────────────────────────────────────────
echo "[1/4] Logging in to AWS ECR..."
aws ecr get-login-password --region "${AWS_REGION}" \
  | docker login --username AWS --password-stdin "${ECR_REGISTRY}"
echo "  ECR login OK"

# ── Step 2: Build ─────────────────────────────────────────────────────────────
echo "[2/4] Building Docker image (no-cache)..."
export TR_JFROG_USERNAME="${JFROG_USER}"
export TR_JFROG_TOKEN="${JFROG_TOKEN}"

docker build --no-cache \
  -t "${MODEL_NAME}" \
  --build-arg TR_JFROG_USERNAME="${TR_JFROG_USERNAME}" \
  --build-arg TR_JFROG_TOKEN="${TR_JFROG_TOKEN}" \
  --file Dockerfile .

echo "  Build OK: ${MODEL_NAME}"

# ── Step 3: Local health check ────────────────────────────────────────────────
echo "[3/4] Testing image locally..."
CONTAINER_ID=$(docker run -d -p 8501:8501 -e USE_LLM=false -e RAG_ENABLED=false "${MODEL_NAME}")
sleep 5

HTTP_STATUS=$(curl -s -o /dev/null -w "%{http_code}" http://localhost:8501/health || echo "000")
docker stop "${CONTAINER_ID}" > /dev/null

if [[ "${HTTP_STATUS}" == "200" ]]; then
  echo "  Health check PASSED (HTTP ${HTTP_STATUS})"
else
  echo "  WARNING: Health check returned HTTP ${HTTP_STATUS} — continuing anyway"
fi

# ── Step 4: Tag and Push ──────────────────────────────────────────────────────
echo "[4/4] Tagging and pushing to ECR..."
docker tag "${MODEL_NAME}" "${FULL_IMAGE_TAG}"
docker push "${FULL_IMAGE_TAG}"

echo ""
echo "========================================================"
echo "  PUSH COMPLETE"
echo ""
echo "  Image ARN (copy for Model Registry):"
echo "  ${FULL_IMAGE_TAG}"
echo ""
echo "  Next: Register in Model Registry:"
echo "  https://contentconsole.thomsonreuters.com/ai-platform/registry/model-registry/models"
echo "========================================================"
