#!/usr/bin/env bash
# =============================================================================
# build_push_v005.sh  —  Build and push sgml-pipeline-prod:0.0.5
# =============================================================================
# USAGE:
#   1. Open WSL Ubuntu terminal
#   2. Run: bash /mnt/c/Users/C303180/OneDrive\ -\ Thomson\ Reuters\ Incorporated/Desktop/TR/sgml-pipeline-deployment/scripts/build_push_v005.sh
#   3. Enter MGMT password when prompted by cloud-tool login
# =============================================================================

set -e

PROJ="/mnt/c/Users/C303180/OneDrive - Thomson Reuters Incorporated/Desktop/TR/sgml-pipeline-deployment"
MODEL_NAME="sgml-pipeline-prod"
VERSION="0.0.5"
ECR_REGISTRY="127288631409.dkr.ecr.us-east-1.amazonaws.com/a207870-ml-model-registry-model-registry-prod-use1"
IMAGE_TAG="${MODEL_NAME}-${VERSION}"

# JFrog credentials — pass as env vars or positional args:
#   bash build_push_v005.sh  C303180  <your-jfrog-token>
TR_JFROG_USERNAME="${1:-C303180}"
TR_JFROG_TOKEN="${2:-}"   # REQUIRED — get from JFrog UI under your profile

echo "======================================================"
echo "  SGML Pipeline v${VERSION} — Build & Push to ECR"
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
echo "  1. Go to Model Registry:"
echo "     https://contentconsole.thomsonreuters.com/ai-platform/registry/model-registry/models"
echo "  2. Find 'sgml-pipeline-prod' → Add Model Version"
echo "     - Version: 0.0.5"
echo "     - Image ARN: ${ECR_REGISTRY}:${IMAGE_TAG}"
echo ""
echo "  3. Go to Deployment:"
echo "     https://contentconsole.thomsonreuters.com/ai-platform/deployment"
echo "  4. Update deployment job 'Securities Commission Conversion'"
echo "     to use version 0.0.5"
echo ""
echo "  IMPORTANT: Update liveness probe path to:"
echo "     /_stcore/health"
echo "  (in the deployment job settings)"
echo "======================================================"
