#!/usr/bin/env bash
# =============================================================================
# build_push_v006.sh  —  Build and push sgml-pipeline-prod:0.0.6
# =============================================================================
# KEY FIX in v0.0.6:
#   - Removed --server.baseUrlPath from CMD (was causing 307 redirects on
#     Plexus readiness probe GET / → MonitoringError)
#   - Plexus ingress handles URL prefix at load balancer level
#   - Increased HEALTHCHECK start-period to 120s and retries to 5
#
# USAGE:
#   1. Open WSL Ubuntu terminal
#   2. Run: bash /mnt/c/Users/C303180/OneDrive\ -\ Thomson\ Reuters\ Incorporated/Desktop/TR/sgml-pipeline-deployment/scripts/build_push_v006.sh C303180 <JFROG_TOKEN>
#   3. Enter MGMT password when prompted by cloud-tool login
# =============================================================================

set -e

PROJ="/mnt/c/Users/C303180/OneDrive - Thomson Reuters Incorporated/Desktop/TR/sgml-pipeline-deployment"
MODEL_NAME="sgml-pipeline-prod"
VERSION="0.0.6"
ECR_REGISTRY="127288631409.dkr.ecr.us-east-1.amazonaws.com/a207870-ml-model-registry-model-registry-prod-use1"
IMAGE_TAG="${MODEL_NAME}-${VERSION}"

# JFrog credentials — pass as positional args or set as env vars
TR_JFROG_USERNAME="${1:-C303180}"
TR_JFROG_TOKEN="${2:-}"   # REQUIRED — get from JFrog UI → your profile → Identity Token

if [ -z "$TR_JFROG_TOKEN" ]; then
    echo "ERROR: JFrog token required as second argument"
    echo "Usage: bash build_push_v006.sh C303180 <your-jfrog-token>"
    exit 1
fi

echo "======================================================"
echo "  SGML Pipeline v${VERSION} — Build & Push to ECR"
echo "  FIX: Removed --server.baseUrlPath (probe 307 fix)"
echo "======================================================"
echo ""

# Step 1: Login to cloud-tool (AWS credentials)
echo "STEP 1/4: Refreshing AWS credentials (cloud-tool login)..."
/root/cloud-tool-env/bin/cloud-tool --region us-east-1 login
echo ""

# Step 2: ECR login via podman
echo "STEP 2/4: Logging into ECR..."
AWS_PROFILE=tr-aiml-hackathon-prod aws ecr get-login-password --region us-east-1 \
  | podman login --username AWS --password-stdin "${ECR_REGISTRY}" --tls-verify=false
echo "ECR login OK"
echo ""

# Step 3: Build Docker image
echo "STEP 3/4: Building Docker image ${IMAGE_TAG}..."
cd "${PROJ}"
podman build \
  --no-cache \
  -t "${MODEL_NAME}:${VERSION}" \
  --build-arg TR_JFROG_USERNAME="${TR_JFROG_USERNAME}" \
  --build-arg TR_JFROG_TOKEN="${TR_JFROG_TOKEN}" \
  --file Dockerfile \
  .
echo "Build complete!"
echo ""

# Step 4: Tag and push to ECR
echo "STEP 4/4: Pushing to ECR..."
podman tag "${MODEL_NAME}:${VERSION}" "${ECR_REGISTRY}:${IMAGE_TAG}"
podman push "${ECR_REGISTRY}:${IMAGE_TAG}" --tls-verify=false

echo ""
echo "======================================================"
echo "  SUCCESS! Image pushed to ECR"
echo ""
echo "  Image ARN:"
echo "  ${ECR_REGISTRY}:${IMAGE_TAG}"
echo ""
echo "  NEXT STEPS (in browser):"
echo "  1. https://contentconsole.thomsonreuters.com/ai-platform/registry/model-registry/models"
echo "     → sgml-pipeline-prod → Add Version"
echo "     → Version: 0.0.6"
echo "     → Image ARN: ${ECR_REGISTRY}:${IMAGE_TAG}"
echo "     → Health Endpoint Path: /_stcore/health    ← CRITICAL"
echo "     → Port: 8501"
echo "     → Save"
echo ""
echo "  2. On the version page → click 'Approve' to move DEVELOPMENT → PRODUCTION"
echo ""
echo "  3. https://contentconsole.thomsonreuters.com/ai-platform/deployment"
echo "     → Job: Securities Commission Conversion"
echo "     → Select version 0.0.6 → Activate"
echo "======================================================"
