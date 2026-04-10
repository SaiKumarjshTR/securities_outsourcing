#!/usr/bin/env bash
# =============================================================================
# deploy_all.sh — Master deployment script: WSL setup → Docker build → ECR push
# Run ONCE from inside WSL after restart:
#   bash scripts/deploy_all.sh <employee_id> <jfrog_token>
# =============================================================================
set -euo pipefail

JFROG_USER="${1:-}"
JFROG_TOKEN="${2:-}"

if [[ -z "$JFROG_USER" || -z "$JFROG_TOKEN" ]]; then
  echo ""
  echo "Usage: bash scripts/deploy_all.sh <employee_id> <jfrog_token>"
  echo ""
  echo "Where:"
  echo "  employee_id  = Your TR employee ID (same as JFrog username)"
  echo "  jfrog_token  = Token from https://trten.sharepoint.com/sites/intr-artifactory-cop"
  echo ""
  exit 1
fi

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"

echo ""
echo "========================================================"
echo "  SGML Pipeline — Full Deployment"
echo "  Project: ${PROJECT_DIR}"
echo "========================================================"

# Phase 2+3+4: WSL Environment Setup
echo ""
echo "[Phase 2-4] Setting up WSL environment..."
bash "${SCRIPT_DIR}/setup_wsl.sh" "${JFROG_USER}" "${JFROG_TOKEN}"

# Phase 5: Docker Build + Push
echo ""
echo "[Phase 5] Building and pushing Docker image..."
cd "${PROJECT_DIR}"
bash "${SCRIPT_DIR}/build_push.sh" "${JFROG_USER}" "${JFROG_TOKEN}"

echo ""
echo "========================================================"
echo "  ALL AUTOMATED PHASES COMPLETE"
echo ""
echo "  MANUAL STEPS REMAINING (browser):"
echo ""
echo "  [Phase 6] Model Registry:"
echo "  https://contentconsole.thomsonreuters.com/ai-platform/registry/model-registry/models"
echo "  → Add Model: sgml-pipeline-prod"
echo "  → Add Version: 0.0.1 with Image ARN shown above"
echo ""
echo "  [Phase 7] Deployment:"
echo "  https://contentconsole.thomsonreuters.com/ai-platform/deployment"
echo "  → Create Deployment Job for sgml-pipeline-prod"
echo "  → Activate Deployment Job"
echo "  → Copy the endpoint URL (share with business)"
echo "========================================================"
