#!/usr/bin/env bash
# =============================================================================
# build_push_v011.sh  —  Build and push sgml-pipeline-prod:0.0.11
# =============================================================================
# WHAT'S NEW in v0.0.11:
#   - ABBYY FREngine 12 (Ubuntu/Linux) replaces Windows FRS14 HTTP bridge
#       • New bridge: pipeline/abbyy_linux_bridge.py  (port 7091 on Ubuntu host)
#       • AbbyyLinuxConverter added to batch_runner_deploy.py
#       • FRS14_SERVER_URL updated to http://172.25.16.1:7091 (Ubuntu bridge)
#       • No Windows FRS14 machine required
#   - test_abbyy_linux.py: standalone diagnostic + integration test
#
# PRE-REQUISITES (one-time, on Ubuntu host):
#   1. Start ABBYY bridge:
#      nohup python3 /home/securities_outsourcing/pipeline/abbyy_linux_bridge.py \
#          --host 0.0.0.0 --port 7091 > /var/log/abbyy_bridge.log 2>&1 &
#   2. Verify bridge is up:
#      curl http://localhost:7091/health
#
# PLEXUS DEPLOYMENT STEPS (after running this script):
#   Step 1: Run this script (builds + pushes Docker image to ECR)
#   Step 2: Register version 0.0.11 in Content Console Model Registry
#   Step 3: Activate Deployment Job (select version 0.0.11)
#
# USAGE:
#   bash scripts/build_push_v011.sh C303180 <JFROG_TOKEN>
#
# PREREQUISITES:
#   - WSL Ubuntu with cloud-tool installed in /root/cloud-tool-env/bin/
#   - podman available
#   - VPN connected (TR internal)
#   - Valid MGMT account for cloud-tool login
#   - ABBYY bridge running on port 7091 (see PRE-REQUISITES above)
# =============================================================================

set -e

# ── Project path ──────────────────────────────────────────────────────────────
PROJ="/home/securities_outsourcing"

# ── Image configuration ───────────────────────────────────────────────────────
MODEL_NAME="sgml-pipeline-prod"
VERSION="0.0.11"
ECR_REGISTRY="127288631409.dkr.ecr.us-east-1.amazonaws.com/a207870-ml-model-registry-model-registry-prod-use1"
IMAGE_TAG="${MODEL_NAME}-${VERSION}"
FULL_IMAGE_URI="${ECR_REGISTRY}:${IMAGE_TAG}"

# ── JFrog credentials ─────────────────────────────────────────────────────────
TR_JFROG_USERNAME="${1:-C303180}"
TR_JFROG_TOKEN="${2:-}"

if [ -z "$TR_JFROG_TOKEN" ]; then
    echo "ERROR: JFrog token required as second argument."
    echo "Usage: bash build_push_v011.sh <jfrog_username> <jfrog_token>"
    echo ""
    echo "Get your token:"
    echo "  https://trten.sharepoint.com/sites/intr-artifactory-cop/SitePages/Updating-to-Access-Token-from-API-Key-for-JFrog-Artifactory.aspx"
    exit 1
fi

echo "======================================================"
echo "  SGML Pipeline v${VERSION} — Build & Push to ECR"
echo "  NEW: Ubuntu ABBYY FREngine 12 bridge (port 7091)"
echo "======================================================"
echo ""

# ── Step 0: Pre-flight — verify ABBYY bridge is running ──────────────────────
echo "STEP 0/5: Checking ABBYY Linux bridge..."
if curl -s --max-time 5 http://localhost:7091/health | grep -q '"status"'; then
    echo "  OK — ABBYY bridge running on port 7091"
else
    echo "  WARNING: ABBYY bridge not detected on port 7091"
    echo "  Start it with:"
    echo "    nohup python3 ${PROJ}/pipeline/abbyy_linux_bridge.py --host 0.0.0.0 --port 7091 > /var/log/abbyy_bridge.log 2>&1 &"
    read -p "  Continue anyway? (y/N): " cont
    [[ "$cont" =~ ^[Yy]$ ]] || exit 1
fi
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
echo "  Image URI:"
echo "  ${FULL_IMAGE_URI}"
echo ""
echo "  ─────────────────────────────────────────────────────"
echo "  NEXT STEPS (browser)"
echo "  ─────────────────────────────────────────────────────"
echo ""
echo "  1. Open Model Registry:"
echo "     https://contentconsole.thomsonreuters.com/ai-platform/registry/model-registry/models"
echo "     → sgml-pipeline-prod → Add Model Version"
echo "     → Version: ${VERSION}"
echo "     → Image ARN: ${FULL_IMAGE_URI}"
echo "     → Health Endpoint Path: /_stcore/health"
echo "     → Port: 8501"
echo "     → Save"
echo ""
echo "  2. Open Deployment:"
echo "     https://contentconsole.thomsonreuters.com/ai-platform/deployment"
echo "     → Job: Securities Commission Conversion"
echo "     → 3 dots (Actions) → Edit"
echo "     → Select version ${VERSION}"
echo "     → Save → Activate Deployment Job"
echo "======================================================"
