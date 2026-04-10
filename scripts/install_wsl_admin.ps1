# =============================================================================
# install_wsl_admin.ps1
# Run this script as Administrator to install WSL + Ubuntu
# Right-click PowerShell → "Run as Administrator" → paste path of this file
# =============================================================================

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "========================================================"
Write-Host "  Installing Windows Subsystem for Linux (WSL)"
Write-Host "========================================================"
Write-Host ""

# Check if already installed
$wslCheck = wsl --status 2>&1
if ($LASTEXITCODE -eq 0 -and $wslCheck -notmatch "not installed") {
    Write-Host "  WSL is already installed."
    wsl --list --verbose
    Write-Host ""
    Write-Host "  If Ubuntu is not shown above, run:"
    Write-Host "  wsl --install -d Ubuntu"
    exit 0
}

Write-Host "  Running: wsl --install (installs WSL + Ubuntu)"
wsl --install

Write-Host ""
Write-Host "========================================================"
Write-Host "  WSL INSTALLED"
Write-Host ""
Write-Host "  IMPORTANT: You MUST restart your machine now."
Write-Host "  After restart, Ubuntu will finish setting up."
Write-Host "  Then run: scripts\setup_wsl.sh from inside WSL"
Write-Host "========================================================"
Write-Host ""

$restart = Read-Host "Restart now? (y/n)"
if ($restart -eq "y") {
    Restart-Computer -Force
}
