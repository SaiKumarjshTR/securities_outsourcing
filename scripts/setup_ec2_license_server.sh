#!/usr/bin/env bash
# =============================================================================
# setup_ec2_license_server.sh
#
# Sets up an ABBYY LicensingService on an EC2 instance in the same AWS VPC
# as Plexus. This allows Plexus containers to get ABBYY licenses without
# needing direct internet access to account.abbyy.com.
#
# Architecture:
#   [Plexus Pod] --TCP:3023--> [EC2 License Server] --HTTPS:443--> account.abbyy.com
#
# Prerequisites:
#   - EC2 instance running Amazon Linux 2 or Ubuntu (x86-64) in same VPC
#   - EC2 security group: inbound TCP 3023 from Plexus subnet CIDR
#   - EC2 outbound TCP 443 open (default in AWS - no NAT gateway restrictions)
#   - SSH access to the EC2 instance
#
# Usage (run from WSL, NOT on EC2):
#   bash scripts/setup_ec2_license_server.sh <ec2-user>@<ec2-ip>
#
# After setup:
#   - Deploy Plexus container with env var: ABBYY_LS_HOST=<ec2-private-ip>
#   - The container's LicensingSettings.xml will be auto-patched at startup
# =============================================================================
set -euo pipefail

EC2_HOST="${1:-}"
if [ -z "$EC2_HOST" ]; then
    echo "Usage: $0 <ec2-user>@<ec2-ip>"
    echo "Example: $0 ec2-user@10.0.1.50"
    exit 1
fi

BUNDLE_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)/abbyy_bundle"

echo "========================================"
echo "  ABBYY License Server → EC2 Setup"
echo "  Target: $EC2_HOST"
echo "========================================"

# ── Step 1: Copy LicensingService binary + libs to EC2 ──────────────────────
echo ""
echo "[1/5] Copying LicensingService binary..."
ssh "$EC2_HOST" "sudo mkdir -p /usr/local/bin/ABBYY/SDK/12/Licensing \
    && sudo mkdir -p /usr/local/lib/ABBYY/SDK/12/Licensing \
    && sudo mkdir -p /var/lib/ABBYY/SDK/12/Licenses"

scp "$BUNDLE_DIR/LicensingService" \
    "$EC2_HOST:/tmp/LicensingService"

ssh "$EC2_HOST" "sudo mv /tmp/LicensingService /usr/local/bin/ABBYY/SDK/12/Licensing/LicensingService \
    && sudo chmod +x /usr/local/bin/ABBYY/SDK/12/Licensing/LicensingService"

# ── Step 2: Copy shared libraries ────────────────────────────────────────────
echo ""
echo "[2/5] Copying shared libraries..."
# Create tar of libs for efficient transfer
tar -czf /tmp/abbyy_ls_libs.tar.gz \
    -C "$BUNDLE_DIR/usr_local_lib_abbyy" .

scp /tmp/abbyy_ls_libs.tar.gz "$EC2_HOST:/tmp/abbyy_ls_libs.tar.gz"
ssh "$EC2_HOST" "sudo tar -xzf /tmp/abbyy_ls_libs.tar.gz -C /usr/local/lib/ABBYY/ \
    && rm /tmp/abbyy_ls_libs.tar.gz"

# ── Step 3: Copy license files ────────────────────────────────────────────────
echo ""
echo "[3/5] Copying ABBYY license files..."
tar -czf /tmp/abbyy_licenses.tar.gz \
    -C "$BUNDLE_DIR/var_lib_abbyy/ABBYY/SDK/12/Licenses" .

scp /tmp/abbyy_licenses.tar.gz "$EC2_HOST:/tmp/abbyy_licenses.tar.gz"
ssh "$EC2_HOST" "sudo tar -xzf /tmp/abbyy_licenses.tar.gz -C /var/lib/ABBYY/SDK/12/Licenses/ \
    && sudo chmod 700 /var/lib/ABBYY/SDK/12/Licenses/ \
    && rm /tmp/abbyy_licenses.tar.gz"

# ── Step 4: Create systemd service ───────────────────────────────────────────
echo ""
echo "[4/5] Creating systemd service..."
ssh "$EC2_HOST" "sudo tee /etc/systemd/system/abbyy-licensing.service > /dev/null" << 'SYSTEMD'
[Unit]
Description=ABBYY FREngine12 LicensingService
After=network-online.target
Wants=network-online.target

[Service]
Type=simple
ExecStart=/usr/local/bin/ABBYY/SDK/12/Licensing/LicensingService /standalone
Environment=LD_LIBRARY_PATH=/usr/local/lib/ABBYY/SDK/12/Licensing
Restart=always
RestartSec=5
StandardOutput=journal
StandardError=journal

[Install]
WantedBy=multi-user.target
SYSTEMD

ssh "$EC2_HOST" "sudo systemctl daemon-reload \
    && sudo systemctl enable abbyy-licensing \
    && sudo systemctl start abbyy-licensing"

# ── Step 5: Verify ────────────────────────────────────────────────────────────
echo ""
echo "[5/5] Verifying service..."
sleep 5
ssh "$EC2_HOST" "sudo systemctl status abbyy-licensing --no-pager && \
    ss -tlnp | grep 3023 && echo '✅ Port 3023 LISTENING'"

EC2_PRIVATE_IP=$(ssh "$EC2_HOST" "hostname -I | awk '{print \$1}'" 2>/dev/null)

echo ""
echo "========================================"
echo "  ✅ EC2 License Server Setup Complete"
echo ""
echo "  EC2 private IP: ${EC2_PRIVATE_IP:-<check manually>}"
echo ""
echo "  Next steps:"
echo "  1. Confirm Plexus security group allows TCP 3023 from:"
echo "     the EC2 instance's private IP to Plexus pod CIDR"
echo "  2. Deploy Plexus container with env var:"
echo "     ABBYY_LS_HOST=${EC2_PRIVATE_IP:-<ec2-private-ip>}"
echo "  3. Build + push v0.0.19 (already updated in Dockerfile)"
echo "========================================"
