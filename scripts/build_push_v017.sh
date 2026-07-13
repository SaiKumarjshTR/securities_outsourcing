#!/bin/bash
# ─────────────────────────────────────────────────────────────────────────────
# build_push_v017.sh — Build and push sgml-pipeline-prod:0.0.22
#
# v0.0.22 changes (vs v0.0.21):
#   - Added missing system libraries required by ABBYY official Docker docs:
#       libx11-6, libfreetype6, libice6, locales (en_US.UTF-8)
#   - Removed libProtection.Developer.so from runtime image (dev-only library)
#   - Added /dev/shm size check in docker_start.sh (ABBYY requires ≥ 1GB;
#     Kubernetes default is 64MB which causes silent OCR failures)
#   - Added check_shm() to abbyy_convert.py --diag
#   - Added Dockerfile.licensing for two-container Plexus deployment
#
# TWO deployment paths are now supported:
#   Option A: Set ABBYY_LS_HOST=<cluster-ip-of-sgml-abbyy-ls>
#             (Deploy Dockerfile.licensing as a separate Plexus service)
#   Option B: Set ABBYY_LS_HOST=<ec2-private-ip> (EC2 license server)
#
# Usage:
#   bash scripts/build_push_v017.sh [JFROG_USER] [JFROG_TOKEN]
# ─────────────────────────────────────────────────────────────────────────────
set -e

JFROG_USER="${1:-C303180}"
JFROG_TOKEN="${2:-}"
VERSION="0.0.22"
IMAGE_NAME="sgml-pipeline-prod"
ECR="127288631409.dkr.ecr.us-east-1.amazonaws.com/a207870-ml-model-registry-model-registry-prod-use1"
ECR_HOST="127288631409.dkr.ecr.us-east-1.amazonaws.com"

cd /home/securities_outsourcing

echo "================================================"
echo " SGML Pipeline v${VERSION} — Runtime lib fixes"
echo " + shm check + two-container LS support"
echo "================================================"
echo ""

# ── Pre-flight ────────────────────────────────────────────────────────────────
if [ ! -d abbyy_bundle/FREngine12 ] || [ ! -f abbyy_bundle/LicensingService ]; then
    echo "ERROR: abbyy_bundle/ is incomplete. Run: bash scripts/bundle_abbyy.sh"
    exit 1
fi

# ── ECR login ─────────────────────────────────────────────────────────────────
echo "=== STEP 1/3: ECR login ==="
aws --profile tr-aiml-hackathon-prod ecr get-login-password --region us-east-1 | \
    docker login --username AWS --password-stdin "$ECR_HOST"
echo ""

# ── Build ─────────────────────────────────────────────────────────────────────
echo "=== STEP 2/3: Building ${IMAGE_NAME}:${VERSION} ==="
echo "Log: /tmp/build_v017.log"
echo ""

docker build \
    --build-arg TR_JFROG_USERNAME="$JFROG_USER" \
    --build-arg TR_JFROG_TOKEN="$JFROG_TOKEN" \
    -t "${IMAGE_NAME}:${VERSION}" \
    . 2>&1 | tee /tmp/build_v017.log

echo "Build exit: ${PIPESTATUS[0]}"
echo ""

# ── Tag and push ──────────────────────────────────────────────────────────────
echo "=== STEP 3/3: Tagging and pushing to ECR ==="
ECR_TAG="${ECR}:${IMAGE_NAME}-${VERSION}"
docker tag "${IMAGE_NAME}:${VERSION}" "$ECR_TAG"
echo "Pushing $ECR_TAG ..."
docker push "$ECR_TAG" 2>&1 | tee /tmp/push_v017.log

echo ""
echo "================================================"
echo " SUCCESS — v${VERSION} pushed to ECR"
echo " ECR: $ECR_TAG"
echo ""
echo " Plexus deployment:"
echo "  1. Model Registry → sgml-pipeline-prod → Add Version ${VERSION}"
echo "     Image: $ECR_TAG"
echo "     Health: /_stcore/health  Port: 8501"
echo ""
echo "  2. CRITICAL — /dev/shm (Kubernetes shm fix):"
echo "     In Plexus pod spec add:"
echo "       volumes: [{name: dshm, emptyDir: {medium: Memory, sizeLimit: 1Gi}}]"
echo "       volumeMounts: [{mountPath: /dev/shm, name: dshm}]"
echo "     Without this, ABBYY OCR will fail (needs 1GB shared memory)."
echo ""
echo "  3. Choose licensing option:"
echo "     Option A — Separate LS container (no EC2 needed):"
echo "       bash scripts/build_push_licensing_v001.sh"
echo "       Deploy sgml-abbyy-ls, then set:"
echo "       ABBYY_LS_HOST=sgml-abbyy-ls.default.svc.cluster.local"
echo ""
echo "     Option B — EC2 License Server (if already running):"
echo "       ABBYY_LS_HOST=<ec2-private-ip>"
echo "================================================"
