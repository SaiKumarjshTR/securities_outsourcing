#!/usr/bin/env bash
# =============================================================================
# build_push_v007.sh  —  Build and push sgml-pipeline-prod:0.0.7
# =============================================================================
# WHAT'S NEW in v0.0.7:
#   - Enterprise SessionManager: per-user isolated temp folders (input/output/logs)
#   - Background session cleanup thread (24-hour TTL, auto-expiry)
#   - "New Session" button now properly deletes prior session's scratch files
#   - Sidebar active-sessions / processing counts sourced from SessionManager
#   - SESSIONS_TEMP_DIR env var supported for custom session storage location
#
# PLEXUS DEPLOYMENT GUIDE (summary):
#   Step 1: WSL + cloud-tool setup (see full guide: Plexus_Deployment_Guide.docx)
#   Step 2: docker/podman build with JFrog credentials
#   Step 3: Tag + push to ECR:
#             127288631409.dkr.ecr.us-east-1.amazonaws.com/
#             a207870-ml-model-registry-model-registry-prod-use1
#   Step 4: Register image URI in TR Content Console → AI Platform Model Registry
#   Step 5: Activate Deployment Job in Content Console
#
# USAGE:
#   bash scripts/build_push_v007.sh C303180 <JFROG_TOKEN>
#
# PREREQUISITES:
#   - WSL Ubuntu with cloud-tool installed in /root/cloud-tool-env/bin/
#   - Docker / podman available
#   - VPN connected (TR internal)
#   - Valid MGMT account for cloud-tool login
# =============================================================================

set -e

# ── Project path ──────────────────────────────────────────────────────────────
PROJ="/mnt/c/Users/C303180/OneDrive - Thomson Reuters Incorporated/Desktop/TR/sgml-pipeline-deployment"

# ── Image configuration ───────────────────────────────────────────────────────
MODEL_NAME="sgml-pipeline-prod"
VERSION="0.0.7"
ECR_REGISTRY="127288631409.dkr.ecr.us-east-1.amazonaws.com/a207870-ml-model-registry-model-registry-prod-use1"
IMAGE_TAG="${MODEL_NAME}-${VERSION}"
FULL_IMAGE_URI="${ECR_REGISTRY}:${IMAGE_TAG}"

# ── JFrog credentials (positional args or env vars) ───────────────────────────
TR_JFROG_USERNAME="${1:-C303180}"
TR_JFROG_TOKEN="${2:-}"

if [ -z "$TR_JFROG_TOKEN" ]; then
    echo "ERROR: JFrog token required as second argument."
    echo "Usage: bash build_push_v007.sh <jfrog_username> <jfrog_token>"
    echo ""
    echo "Get your token:"
    echo "  https://trten.sharepoint.com/sites/intr-artifactory-cop/SitePages/Updating-to-Access-Token-from-API-Key-for-JFrog-Artifactory.aspx"
    exit 1
fi

echo "======================================================"
echo "  SGML Pipeline v${VERSION} — Build & Push to ECR"
echo "  NEW: Enterprise SessionManager (per-user isolation)"
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
echo "  BUILD & PUSH SUCCESSFUL"
echo "======================================================"
echo ""
echo "  Image URI (for Content Console):"
echo "  ${FULL_IMAGE_URI}"
echo ""
echo "  NEXT STEPS — Plexus deployment:"
echo "  1. Open TR Content Console:"
echo "     https://contentconsole.thomsonreuters.com/ai-platform/registry/model-registry/models"
echo "  2. Select model: ${MODEL_NAME}"
echo "  3. Add Model Version: ${VERSION}"
echo "  4. Paste Image URI above → Save"
echo "  5. Deployment Jobs → Actions (...) → Activate Deployment Job"
echo ""
echo "  Health endpoint (Streamlit built-in):"
echo "  GET /_stcore/health"
echo ""
echo "  Port: 8501"
echo "======================================================"
