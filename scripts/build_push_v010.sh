#!/usr/bin/env bash
# =============================================================================
# build_push_v010.sh  —  Build and push sgml-pipeline-prod:0.0.10
# =============================================================================
# WHAT'S NEW in v0.0.10:
#   - batch_runner_standalone.py updated (Apr 2026):
#       extract_inline_formatting: anchor each run via str.find() on para.text
#       instead of accumulating len(run.text) — prevents BOLD/EM span drift
#       caused by hidden Unicode chars (hyperlink wrappers, w:noBreakHyphen,
#       w:sym), fixing empty-parentheses bug (e.g. '()' → '(<BOLD>PPM</BOLD>)')
#   - SequentialSGMLLayer v14: Structure → Inline → Validate agents
#   - SGMLGenerator v5.1: container blocks, full entity map, table improvements
#
# PLEXUS DEPLOYMENT GUIDE (summary):
#   Step 1: Run this script from WSL (provides interactive cloud-tool login)
#   Step 2: Register new version 0.0.10 in Content Console Model Registry
#   Step 3: Activate Deployment Job (select version 0.0.10)
#
# USAGE:
#   bash scripts/build_push_v010.sh C303180 <JFROG_TOKEN>
#
# PREREQUISITES:
#   - WSL Ubuntu with cloud-tool installed in /root/cloud-tool-env/bin/
#   - podman available
#   - VPN connected (TR internal)
#   - Valid MGMT account for cloud-tool login
# =============================================================================

set -e

# ── Project path ──────────────────────────────────────────────────────────────
PROJ="/mnt/c/Users/C303180/OneDrive - Thomson Reuters Incorporated/Desktop/TR/sgml-pipeline-deployment"

# ── Image configuration ───────────────────────────────────────────────────────
MODEL_NAME="sgml-pipeline-prod"
VERSION="0.0.10"
ECR_REGISTRY="127288631409.dkr.ecr.us-east-1.amazonaws.com/a207870-ml-model-registry-model-registry-prod-use1"
IMAGE_TAG="${MODEL_NAME}-${VERSION}"
FULL_IMAGE_URI="${ECR_REGISTRY}:${IMAGE_TAG}"

# ── JFrog credentials (positional args or env vars) ───────────────────────────
TR_JFROG_USERNAME="${1:-C303180}"
TR_JFROG_TOKEN="${2:-}"

if [ -z "$TR_JFROG_TOKEN" ]; then
    echo "ERROR: JFrog token required as second argument."
    echo "Usage: bash build_push_v010.sh <jfrog_username> <jfrog_token>"
    echo ""
    echo "Get your token:"
    echo "  https://trten.sharepoint.com/sites/intr-artifactory-cop/SitePages/Updating-to-Access-Token-from-API-Key-for-JFrog-Artifactory.aspx"
    exit 1
fi

echo "======================================================"
echo "  SGML Pipeline v${VERSION} — Build & Push to ECR"
echo "  NEW: extract_inline_formatting Apr 2026 fix (span drift)"
echo "======================================================"
echo ""

# ── Step 1: Refresh AWS credentials via cloud-tool ────────────────────────────
echo "STEP 1/5: Refreshing AWS credentials (cloud-tool login)..."
/root/cloud-tool-env/bin/cloud-tool --region us-east-1 login
echo ""

# ── Step 2: Authenticate to ECR ──────────────────────────────────────────────
echo "STEP 2/5: Logging into ECR..."
AWS_PROFILE=tr-aiml-hackathon-prod \
  aws ecr get-login-password --region us-east-1 \
  | podman login --username AWS --password-stdin "${ECR_REGISTRY}" --tls-verify=false
echo "ECR login OK"
echo ""

# ── Step 3: Build Docker image ────────────────────────────────────────────────
echo "STEP 3/5: Building Docker image ${IMAGE_TAG}..."
cd "${PROJ}"
podman build \
  --no-cache \
  -t "${MODEL_NAME}:${VERSION}" \
  --build-arg TR_JFROG_USERNAME="${TR_JFROG_USERNAME}" \
  --build-arg TR_JFROG_TOKEN="${TR_JFROG_TOKEN}" \
  --file Dockerfile \
  .
echo "Build complete: ${MODEL_NAME}:${VERSION}"
echo ""

# ── Step 4: Tag image for ECR ─────────────────────────────────────────────────
echo "STEP 4/5: Tagging image for ECR..."
podman tag "${MODEL_NAME}:${VERSION}" "${FULL_IMAGE_URI}"
echo "Tagged → ${FULL_IMAGE_URI}"
echo ""

# ── Step 5: Push to ECR ───────────────────────────────────────────────────────
echo "STEP 5/5: Pushing to ECR..."
podman push "${FULL_IMAGE_URI}" --tls-verify=false
echo "Push complete!"
echo ""

# ── Summary ───────────────────────────────────────────────────────────────────
echo "======================================================"
echo "  SUCCESS! Image pushed to ECR"
echo ""
echo "  Image ARN:"
echo "  ${FULL_IMAGE_URI}"
echo ""
echo "  NEXT STEPS (browser):"
echo ""
echo "  1. Model Registry:"
echo "  https://contentconsole.thomsonreuters.com/ai-platform/registry/model-registry/models"
echo "     → sgml-pipeline-prod → Add Version"
echo "     → Version: 0.0.10"
echo "     → Image ARN: ${FULL_IMAGE_URI}"
echo "     → Health Endpoint Path: /_stcore/health"
echo "     → Port: 8501"
echo "     → Save → Approve (DEVELOPMENT → PRODUCTION)"
echo ""
echo "  2. Deployment:"
echo "  https://contentconsole.thomsonreuters.com/ai-platform/deployment"
echo "     → Job: Securities Commission Conversion"
echo "     → Select version 0.0.10 → Activate"
echo "======================================================"
