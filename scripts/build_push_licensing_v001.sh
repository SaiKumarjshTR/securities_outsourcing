#!/bin/bash
# ─────────────────────────────────────────────────────────────────────────────
# build_push_licensing_v001.sh — Build and push sgml-abbyy-ls:1.0.0
#
# This builds the ABBYY LicensingService-only container (Dockerfile.licensing).
# Deploy this as a SEPARATE Plexus service so the main app container can
# connect to it via ABBYY_LS_HOST=<cluster-ip-of-this-service>
#
# This implements the ABBYY-official two-container pattern:
#   https://docs.abbyy.com/fine-reader/engine/distribution/distribution-linux/
#   running-abbyy-finereader-engine-12-inside-a-docker-container-2
#
# Usage:
#   bash scripts/build_push_licensing_v001.sh
#
# AWS credentials must already be active (run cloud-tool --region us-east-1 login).
# ─────────────────────────────────────────────────────────────────────────────
set -e

VERSION="1.0.0"
IMAGE_NAME="sgml-abbyy-ls"
ECR="127288631409.dkr.ecr.us-east-1.amazonaws.com/a207870-ml-model-registry-model-registry-prod-use1"
ECR_HOST="127288631409.dkr.ecr.us-east-1.amazonaws.com"

cd /home/securities_outsourcing

echo "================================================"
echo " ABBYY LicensingService Container v${VERSION}"
echo " Two-container pattern — Plexus deployment"
echo "================================================"
echo ""
echo " Architecture:"
echo "  [App Pod]  --TCP:3023-->  [This LS Pod]  --HTTPS:443-->  account.abbyy.com"
echo ""

# ── Pre-flight checks ─────────────────────────────────────────────────────────
if [ ! -f Dockerfile.licensing ]; then
    echo "ERROR: Dockerfile.licensing not found in $(pwd)"
    exit 1
fi

if [ ! -f abbyy_bundle/LicensingService ]; then
    echo "ERROR: abbyy_bundle/ is incomplete. Run: bash scripts/bundle_abbyy.sh"
    exit 1
fi

echo "Checking license files..."
ls -la abbyy_bundle/var_lib_abbyy/ABBYY/SDK/12/Licenses/*.ABBYY.ActivationToken 2>/dev/null \
    || { echo "ERROR: No .ABBYY.ActivationToken found in abbyy_bundle/var_lib_abbyy/"; exit 1; }

# ── ECR login ─────────────────────────────────────────────────────────────────
echo ""
echo "=== STEP 1/3: ECR login ==="
aws --profile tr-aiml-hackathon-prod ecr get-login-password --region us-east-1 | \
    docker login --username AWS --password-stdin "$ECR_HOST"

# ── Build ─────────────────────────────────────────────────────────────────────
echo ""
echo "=== STEP 2/3: Building ${IMAGE_NAME}:${VERSION} ==="
echo "Log: /tmp/build_licensing_v001.log"
echo ""

docker build \
    -f Dockerfile.licensing \
    -t "${IMAGE_NAME}:${VERSION}" \
    . 2>&1 | tee /tmp/build_licensing_v001.log

echo "Build exit: ${PIPESTATUS[0]}"

# ── Tag and push ──────────────────────────────────────────────────────────────
echo ""
echo "=== STEP 3/3: Tagging and pushing to ECR ==="
ECR_TAG="${ECR}:${IMAGE_NAME}-${VERSION}"
docker tag "${IMAGE_NAME}:${VERSION}" "$ECR_TAG"
echo "Pushing $ECR_TAG ..."
docker push "$ECR_TAG" 2>&1 | tee /tmp/push_licensing_v001.log

echo ""
echo "================================================"
echo " SUCCESS — LicensingService image pushed"
echo " ECR: $ECR_TAG"
echo ""
echo " NEXT STEPS (Plexus):"
echo ""
echo "  1. Model Registry → Add a NEW model 'sgml-abbyy-ls'"
echo "     Image : $ECR_TAG"
echo "     Port  : 3023"
echo "     Health: (none — TCP check only)"
echo ""
echo "  2. Deploy 'sgml-abbyy-ls' with:"
echo "     - Outbound internet allowed (account.abbyy.com:443)"
echo "     - No shm requirement (LS doesn't need FRE shared memory)"
echo "     - Replica: 1 (one LS per Online License)"
echo ""
echo "  3. Get the cluster-internal IP/hostname of sgml-abbyy-ls"
echo "     (check Plexus service discovery / kubectl get svc)"
echo ""
echo "  4. In main 'sgml-pipeline-prod' deployment, set env var:"
echo "     ABBYY_LS_HOST=<cluster-ip-or-hostname-of-sgml-abbyy-ls>"
echo "     e.g.  ABBYY_LS_HOST=sgml-abbyy-ls.default.svc.cluster.local"
echo ""
echo "  5. Also fix /dev/shm in the main app pod (Kubernetes YAML):"
echo "     volumes:"
echo "       - name: dshm"
echo "         emptyDir:"
echo "           medium: Memory"
echo "           sizeLimit: 1Gi"
echo "     volumeMounts:"
echo "       - mountPath: /dev/shm"
echo "         name: dshm"
echo "================================================"
