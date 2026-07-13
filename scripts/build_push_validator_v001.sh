#!/bin/bash
# ─────────────────────────────────────────────────────────────────────────────
# build_push_validator_v001.sh — Build and push sgml-validator-prod:0.0.1
#
# Plexus model: Securities_Commission_Conversion_Validator
# Version:      0.0.1
#
# Usage:
#   bash scripts/build_push_validator_v001.sh
#
# AWS credentials must already be active (run cloud-tool --region us-east-1 login).
# ─────────────────────────────────────────────────────────────────────────────
set -e

VERSION="0.0.1"
IMAGE_NAME="sgml-validator-prod"
ECR="127288631409.dkr.ecr.us-east-1.amazonaws.com/a207870-ml-model-registry-model-registry-prod-use1"
ECR_HOST="127288631409.dkr.ecr.us-east-1.amazonaws.com"

cd /home/securities_outsourcing

# Copy validator requirements.txt from the GitHub repo (separate from pipeline)
cp /tmp/validator_latest/requirements.txt ./requirements_validator.txt
cp /tmp/validator_latest/.streamlit/config.toml ./.streamlit/config.toml 2>/dev/null || true

echo "=============================================="
echo " TR SGML Validator v${VERSION}"
echo "=============================================="
echo ""

# ── Step 1: ECR login ─────────────────────────────────────────────────────────
echo "=== STEP 1/3: ECR login ==="
aws --profile tr-aiml-hackathon-prod ecr get-login-password --region us-east-1 | \
    podman login --username AWS --password-stdin "$ECR_HOST" --tls-verify=false
echo ""

# ── Step 2: Build ─────────────────────────────────────────────────────────────
echo "=== STEP 2/3: Building ${IMAGE_NAME}:${VERSION} ==="
echo "Log: /tmp/build_validator_v001.log"
echo ""

podman build --no-cache \
    -f Dockerfile.validator \
    -t "${IMAGE_NAME}:${VERSION}" \
    . 2>&1 | tee /tmp/build_validator_v001.log

echo "Build exit: ${PIPESTATUS[0]}"
echo ""

# ── Step 3: Tag and push ──────────────────────────────────────────────────────
echo "=== STEP 3/3: Tagging and pushing to ECR ==="
ECR_TAG="${ECR}:${IMAGE_NAME}-${VERSION}"
podman tag "${IMAGE_NAME}:${VERSION}" "$ECR_TAG"
echo "Pushing $ECR_TAG ..."
podman push "$ECR_TAG" --tls-verify=false 2>&1 | tee /tmp/push_validator_v001.log

echo ""
echo "=============================================="
echo " SUCCESS — Validator v${VERSION} pushed to ECR"
echo " ECR: $ECR_TAG"
echo ""
echo " Plexus next steps:"
echo "  1. Model Registry → Securities_Commission_Conversion_Validator"
echo "     → Add Version 0.0.1"
echo "     Image: $ECR_TAG"
echo "     Health: /_stcore/health  Port: 8501"
echo "  2. Deployment → Activate v0.0.1"
echo "=============================================="
