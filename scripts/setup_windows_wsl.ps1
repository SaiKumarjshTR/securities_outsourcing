# =============================================================================
# setup_windows_wsl.ps1
#
# SGML Pipeline — Windows Installer via WSL2
#
# Since ABBYY FREngine12 is a Linux x86-64 binary, it CANNOT run natively on
# Windows. WSL2 (Windows Subsystem for Linux) runs Ubuntu as a real Linux
# kernel — the SAME ABBYY binary works identically.
#
# WHAT THIS DOES:
#   1. Checks Windows version compatibility (Win 10 2004+ / Win 11)
#   2. Enables WSL2 feature (requires admin + restart on first run)
#   3. Installs Ubuntu 22.04 via WSL
#   4. Copies the portable bundle into Ubuntu WSL
#   5. Runs setup_local_ubuntu.sh inside WSL
#   6. Creates Desktop shortcut + Start Menu entry
#   7. Creates launch/stop scripts in the current directory
#
# REQUIREMENTS:
#   • Windows 10 version 2004+ (Build 19041+) or Windows 11
#   • Run as Administrator
#   • Internet access (for WSL2 Ubuntu download ~500MB, pip packages)
#   • The sgml-pipeline portable bundle in the same directory as this script
#     (sgml-pipeline-bundle.tar.gz — created by scripts/create_portable_bundle.sh)
#
# USAGE:
#   Right-click this file → Run with PowerShell (as Administrator)
#   Or from Admin PowerShell:
#     Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
#     .\scripts\setup_windows_wsl.ps1
#
# AFTER INSTALL:
#   Double-click "SGML Pipeline" on Desktop → browser opens to localhost:8501
# =============================================================================

# ── Self-elevate to Administrator if needed ──────────────────────────────────
# This replaces #Requires -RunAsAdministrator so the window does NOT silently
# close when run without admin rights — instead it re-launches with UAC prompt.
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Requesting Administrator privileges..." -ForegroundColor Yellow
    Write-Host "A UAC (User Account Control) prompt will appear — click Yes." -ForegroundColor White
    Start-Process PowerShell.exe -Verb RunAs `
        -ArgumentList "-ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Path)`""
    exit
}

# Now running as Administrator
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
$ErrorActionPreference = "Continue"   # Don't silently exit on every error

# ── Config ────────────────────────────────────────────────────────────────────
$UBUNTU_DISTRO  = "Ubuntu-22.04"
$WSL_APP_DIR    = "/opt/sgml-pipeline"
$WIN_APP_DIR    = "$env:LOCALAPPDATA\SGMLPipeline"
$APP_PORT       = 8501
$BUNDLE_NAME    = "sgml-pipeline-bundle.tar.gz"
$SCRIPT_DIR     = Split-Path -Parent $MyInvocation.MyCommand.Path
$REPO_DIR       = Split-Path -Parent $SCRIPT_DIR

function Write-Step   { param($msg) Write-Host "`n[$([datetime]::Now.ToString('HH:mm:ss'))] $msg" -ForegroundColor Yellow }
function Write-Ok     { param($msg) Write-Host "  ✓ $msg" -ForegroundColor Green }
function Write-Warn   { param($msg) Write-Host "  ⚠ $msg" -ForegroundColor Yellow }
function Write-Error2 { param($msg) Write-Host "  ✗ $msg" -ForegroundColor Red; exit 1 }

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  SGML Pipeline — Windows Installer (WSL2 Ubuntu)" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  ABBYY FREngine12 is Linux x86-64 only." -ForegroundColor White
Write-Host "  This installer deploys it inside WSL2 Ubuntu." -ForegroundColor White
Write-Host "  Your business machine HAS internet — Online License will work." -ForegroundColor White
Write-Host ""

# ── Check Windows version ─────────────────────────────────────────────────────
Write-Step "1/7  Checking Windows version..."
$WinVer = [System.Environment]::OSVersion.Version
$WinBuild = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").CurrentBuild
Write-Host "  Windows build: $WinBuild" -ForegroundColor Gray
if ([int]$WinBuild -lt 19041) {
    Write-Error2 "Windows 10 Build 19041 (version 2004) or later required for WSL2.`nYour build: $WinBuild — please update Windows."
}
Write-Ok "Windows version OK ($WinBuild)"

# ── Check/Enable WSL2 ─────────────────────────────────────────────────────────
Write-Step "2/7  Checking WSL2..."
$wslInstalled = $false
try {
    $wslInfo = wsl --status 2>&1
    if ($LASTEXITCODE -eq 0) { $wslInstalled = $true }
} catch {}

if (-not $wslInstalled) {
    Write-Warn "WSL2 not installed. Installing now..."
    Write-Host "  This enables the Windows Subsystem for Linux 2 feature." -ForegroundColor Gray
    
    # Enable WSL and Virtual Machine Platform features
    dism.exe /online /enable-feature /featurename:Microsoft-Windows-Subsystem-Linux /all /norestart | Out-Null
    dism.exe /online /enable-feature /featurename:VirtualMachinePlatform /all /norestart | Out-Null
    
    # Download and install WSL2 kernel update
    Write-Host "  Downloading WSL2 kernel update..." -ForegroundColor Gray
    $KernelUrl = "https://wslstorestorage.blob.core.windows.net/wslblob/wsl_update_x64.msi"
    $KernelMsi = "$env:TEMP\wsl_update_x64.msi"
    Invoke-WebRequest -Uri $KernelUrl -OutFile $KernelMsi -UseBasicParsing
    Start-Process msiexec.exe -ArgumentList "/i $KernelMsi /quiet" -Wait
    
    # Set WSL2 as default
    wsl --set-default-version 2
    
    Write-Warn "RESTART REQUIRED after WSL2 feature enablement."
    Write-Host ""
    Write-Host "  After restart, run this script again to continue." -ForegroundColor Cyan
    $restart = Read-Host "  Restart now? (y/n)"
    if ($restart -eq "y") { Restart-Computer -Force }
    exit 0
} else {
    # Set WSL2 as default (safe to re-run)
    wsl --set-default-version 2 2>&1 | Out-Null
    Write-Ok "WSL2 available"
}

# ── Install Ubuntu 22.04 via WSL ──────────────────────────────────────────────
Write-Step "3/7  Installing Ubuntu 22.04..."
$distros = wsl --list --quiet 2>&1 | Where-Object { $_ -match "Ubuntu" }
if ($distros) {
    Write-Ok "Ubuntu already installed: $($distros -join ', ')"
} else {
    Write-Host "  Installing Ubuntu 22.04 (~500MB download)..." -ForegroundColor Gray
    wsl --install -d $UBUNTU_DISTRO --no-launch
    
    # First launch to initialize (non-interactive)
    Write-Host "  Initializing Ubuntu (first run setup)..." -ForegroundColor Gray
    wsl -d $UBUNTU_DISTRO -- bash -c "echo 'Ubuntu initialized'" 2>&1 | Out-Null
    Write-Ok "Ubuntu 22.04 installed"
}

# Set as default distro
wsl --set-default $UBUNTU_DISTRO 2>&1 | Out-Null

# ── Find or create bundle ──────────────────────────────────────────────────────
Write-Step "4/7  Preparing SGML Pipeline bundle..."

# Look for bundle in several places
$BundlePath = $null
foreach ($loc in @("$SCRIPT_DIR\..\$BUNDLE_NAME", "$REPO_DIR\$BUNDLE_NAME", "$env:USERPROFILE\Downloads\$BUNDLE_NAME")) {
    if (Test-Path $loc) { $BundlePath = (Resolve-Path $loc).Path; break }
}

if (-not $BundlePath) {
    Write-Warn "Bundle $BUNDLE_NAME not found."
    Write-Host ""
    Write-Host "  To create the bundle, run on the Ubuntu build machine:" -ForegroundColor Cyan
    Write-Host "    bash scripts/create_portable_bundle.sh" -ForegroundColor White
    Write-Host "  Then copy sgml-pipeline-bundle.tar.gz to this Windows machine" -ForegroundColor White
    Write-Host "  and run this script again." -ForegroundColor White
    Write-Host ""
    
    # Fallback: try to use the repo directory directly if running from the repo
    $LocalSetup = "$REPO_DIR\scripts\setup_local_ubuntu.sh"
    if (Test-Path $LocalSetup) {
        Write-Warn "No bundle found — attempting direct install from repo in WSL..."
        # Convert Windows path to WSL path
        $WslRepoPath = wsl -- wslpath "'$REPO_DIR'" 2>&1
        $BundlePath = "REPO:$REPO_DIR"
    } else {
        Write-Error2 "Cannot find bundle or repo. Download sgml-pipeline-bundle.tar.gz and place next to this script."
    }
}

if ($BundlePath -notlike "REPO:*") {
    # Copy bundle into WSL
    Write-Host "  Copying bundle to WSL ($([int]($(Get-Item $BundlePath).Length / 1MB))MB)..." -ForegroundColor Gray
    $WslTmp = "/tmp/sgml-pipeline-bundle.tar.gz"
    wsl -- cp (wsl -- wslpath "'$BundlePath'") $WslTmp
    
    # Extract and install inside WSL
    wsl -- bash -c "
        set -e
        echo '[WSL] Extracting bundle...'
        mkdir -p /tmp/sgml-install
        tar -xzf $WslTmp -C /tmp/sgml-install
        echo '[WSL] Running installer...'
        sudo bash /tmp/sgml-install/sgml-pipeline/scripts/setup_local_ubuntu.sh
        rm -rf /tmp/sgml-install $WslTmp
    "
} else {
    # Use repo directly
    $RepoDir = $BundlePath.Substring(5)
    $WslRepo = (wsl -- wslpath "'$RepoDir'") 2>&1
    wsl -- bash -c "sudo bash '$WslRepo/scripts/setup_local_ubuntu.sh'"
}

Write-Ok "SGML Pipeline installed in WSL"

# ── Configure WSL2 networking (mirrored mode) ────────────────────────────────
Write-Step "5a/7  Configuring WSL2 networking..."
$wslConfigPath = "$env:USERPROFILE\.wslconfig"
$wslConfigContent = @"
[wsl2]
# Mirrored networking: WSL2 shares the Windows network stack.
# localhost in WSL2 = localhost in Windows — no port forwarding needed.
# Required for the SGML Pipeline browser to reach the app.
networkingMode=mirrored
localhostForwarding=true
"@
Set-Content -Path $wslConfigPath -Value $wslConfigContent -Encoding UTF8
Write-Ok "WSL2 networking configured (mirrored mode) at $wslConfigPath"

# ── Add Windows Firewall inbound rule for port 8501 ───────────────────────────
Write-Step "5b/7  Adding Windows Firewall rule for port $APP_PORT..."
try {
    # Remove old rule if exists
    Remove-NetFirewallRule -DisplayName "SGML Pipeline*" -ErrorAction SilentlyContinue
    # Add new inbound rule allowing localhost→WSL2 traffic on port 8501
    New-NetFirewallRule `
        -DisplayName "SGML Pipeline (Streamlit port $APP_PORT)" `
        -Direction Inbound `
        -Action Allow `
        -Protocol TCP `
        -LocalPort $APP_PORT `
        -Profile Any `
        -Description "SGML Pipeline Streamlit app running in WSL2 Ubuntu" | Out-Null
    Write-Ok "Firewall rule added: allow inbound TCP port $APP_PORT"
} catch {
    Write-Warn "Could not add firewall rule: $_ — localhost may not work, use WSL2 IP instead."
}

# ── Restart WSL2 to apply networking config ───────────────────────────────────
Write-Host "  Restarting WSL2 to apply networking changes..." -ForegroundColor Gray
wsl --shutdown 2>&1 | Out-Null
Start-Sleep -Seconds 3
Write-Ok "WSL2 restarted with mirrored networking"

# ── Create Windows launcher scripts ───────────────────────────────────────────
Write-Step "5c/7  Creating Windows launchers..."

New-Item -ItemType Directory -Force -Path $WIN_APP_DIR | Out-Null

# start.bat
@"
@echo off
title SGML Pipeline
echo Starting SGML Pipeline...
echo.
wsl -d Ubuntu -- bash -c "nohup sgml-pipeline start > /tmp/sgml-pipeline.log 2>&1 & sleep 2"
timeout /t 8 /nobreak > nul
start "" "http://localhost:$APP_PORT"
echo App running at http://localhost:$APP_PORT
echo Close this window to keep the app running.
pause
"@ | Set-Content "$WIN_APP_DIR\start.bat" -Encoding ASCII

# stop.bat  
@"
@echo off
title SGML Pipeline - Stop
echo Stopping SGML Pipeline...
wsl -d Ubuntu -- bash -c "sgml-pipeline stop"
echo Done.
pause
"@ | Set-Content "$WIN_APP_DIR\stop.bat" -Encoding ASCII

# open.bat (just opens browser, assumes app already running)
@"
@echo off
start "" "http://localhost:$APP_PORT"
"@ | Set-Content "$WIN_APP_DIR\open.bat" -Encoding ASCII

Write-Ok "Launcher scripts created in $WIN_APP_DIR"

# ── Desktop shortcut ──────────────────────────────────────────────────────────
Write-Step "6/7  Creating Desktop shortcut..."

$WshShell = New-Object -ComObject WScript.Shell

# Desktop shortcut
$Shortcut = $WshShell.CreateShortcut("$env:USERPROFILE\Desktop\SGML Pipeline.lnk")
$Shortcut.TargetPath = "$WIN_APP_DIR\start.bat"
$Shortcut.WorkingDirectory = $WIN_APP_DIR
$Shortcut.Description = "Start SGML Pipeline (PDF to DOCX/SGML)"
$Shortcut.WindowStyle = 7   # Minimized
$Shortcut.Save()

# Start Menu shortcut
$StartMenuDir = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\SGML Pipeline"
New-Item -ItemType Directory -Force -Path $StartMenuDir | Out-Null

$Shortcut2 = $WshShell.CreateShortcut("$StartMenuDir\SGML Pipeline.lnk")
$Shortcut2.TargetPath = "$WIN_APP_DIR\start.bat"
$Shortcut2.Description = "Start SGML Pipeline"
$Shortcut2.WindowStyle = 7
$Shortcut2.Save()

$Shortcut3 = $WshShell.CreateShortcut("$StartMenuDir\Stop SGML Pipeline.lnk")
$Shortcut3.TargetPath = "$WIN_APP_DIR\stop.bat"
$Shortcut3.Description = "Stop SGML Pipeline"
$Shortcut3.Save()

Write-Ok "Desktop shortcut: 'SGML Pipeline'"
Write-Ok "Start Menu: SGML Pipeline"

# ── Start the app ─────────────────────────────────────────────────────────────
Write-Step "7/7  Starting SGML Pipeline..."
wsl -- bash -c "sgml-pipeline start > /tmp/sgml-pipeline.log 2>&1 &"
Start-Sleep -Seconds 8

# Open browser
Start-Process "http://localhost:$APP_PORT"
Write-Ok "App launched in browser"

# ── Done ──────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  Installation Complete!" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  App URL  : http://localhost:$APP_PORT" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Shortcuts:" -ForegroundColor White
Write-Host "    Desktop      → 'SGML Pipeline' (starts app + opens browser)"
Write-Host "    Start Menu   → SGML Pipeline"
Write-Host "    Launcher dir → $WIN_APP_DIR"
Write-Host ""
Write-Host "  WSL2 commands (from PowerShell/CMD):" -ForegroundColor White
Write-Host "    wsl -- sgml-pipeline start"
Write-Host "    wsl -- sgml-pipeline stop"
Write-Host "    wsl -- sgml-pipeline status"
Write-Host "    wsl -- sgml-pipeline diag"
Write-Host ""
Write-Host "  NOTE: ABBYY license needs internet to account.abbyy.com:443" -ForegroundColor Yellow
Write-Host "        This works on business machines (no firewall blocking)." -ForegroundColor Yellow
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""
Read-Host "Press Enter to exit"
