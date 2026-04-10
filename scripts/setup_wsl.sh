#!/usr/bin/env bash
# =============================================================================
# setup_wsl.sh  — Run ONCE inside WSL Ubuntu after wsl --install + restart
# Purpose: Install Python venv, cloud-tool, Docker, AWS CLI
# Usage:   bash setup_wsl.sh <YOUR_EMPLOYEE_ID> <YOUR_JFROG_TOKEN>
# =============================================================================
set -euo pipefail

JFROG_USER="${1:-}"
JFROG_TOKEN="${2:-}"

if [[ -z "$JFROG_USER" || -z "$JFROG_TOKEN" ]]; then
  echo "Usage: bash setup_wsl.sh <employee_id> <jfrog_token>"
  echo "Get token from: https://trten.sharepoint.com/sites/intr-artifactory-cop/SitePages/How-to-Log-in-to-Artifactory.aspx"
  exit 1
fi

echo ""
echo "========================================================"
echo "  SGML Pipeline — WSL Environment Setup"
echo "========================================================"
echo ""

# ── Step 1: Python venv + cloud-tool ─────────────────────────────────────────
echo "[1/4] Installing Python venv and cloud-tool..."
sudo apt-get update -qq
sudo apt-get install -y python3.12-venv --quiet

python3 -m venv ~/cloud-tool
source ~/cloud-tool/bin/activate

pip3 install cloud-tool \
  --extra-index-url "https://${JFROG_USER}:${JFROG_TOKEN}@tr1.jfrog.io/tr1/api/pypi/pypi-local/simple" \
  --quiet

echo "  cloud-tool installed: $(cloud-tool --version 2>&1 | head -1)"

# ── Step 2: Docker ────────────────────────────────────────────────────────────
echo "[2/4] Installing Docker..."
sudo apt-get install -y apt-transport-https ca-certificates curl gnupg lsb-release --quiet
curl -fsSL https://download.docker.com/linux/ubuntu/gpg \
  | sudo gpg --dearmor -o /usr/share/keyrings/docker-archive-keyring.gpg

echo "deb [arch=$(dpkg --print-architecture) signed-by=/usr/share/keyrings/docker-archive-keyring.gpg] \
https://download.docker.com/linux/ubuntu $(lsb_release -cs) stable" \
  | sudo tee /etc/apt/sources.list.d/docker.list > /dev/null

sudo apt-get update -qq
sudo apt-get install -y docker-ce docker-ce-cli containerd.io --quiet
sudo usermod -aG docker "$USER"
sudo service docker start

echo "  Docker: $(docker --version)"

# ── Step 3: AWS CLI ───────────────────────────────────────────────────────────
echo "[3/4] Installing AWS CLI..."
sudo apt-get install -y unzip --quiet
curl -sS "https://awscli.amazonaws.com/awscli-exe-linux-x86_64.zip" -o /tmp/awscliv2.zip
unzip -q /tmp/awscliv2.zip -d /tmp
sudo /tmp/aws/install --update
echo "  AWS CLI: $(aws --version)"

# ── Step 4: cloud-tool login ──────────────────────────────────────────────────
echo "[4/4] Logging in to cloud-tool (TR AWS)..."
echo "  You will be prompted for:"
echo "    - MGMT username"
echo "    - Vault password"
echo "    - Select role: option 2 (aiml)"
echo ""
export AWS_PROFILE="default"
source ~/cloud-tool/bin/activate
cloud-tool --region us-east-1 login

echo ""
echo "========================================================"
echo "  Setup complete! Proceed to: bash build_push.sh"
echo "========================================================"
