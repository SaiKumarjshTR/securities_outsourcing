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
#   3. Installs Ubuntu 24.04 via WSL
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

#Requires -RunAsAdministrator
param([string]$BundlePath = "")
$ErrorActionPreference = "Stop"

# ── Config ────────────────────────────────────────────────────────────────────────────────────────
$UBUNTU_DISTRO  = "Ubuntu-24.04"
$WSL_APP_DIR    = "/opt/sgml-pipeline"
$WIN_APP_DIR    = "$env:LOCALAPPDATA\SGMLPipeline"
$APP_PORT       = 8501
# Explicit fallback list (kept for clarity/logging) — the glob search below is the real
# safety net, so a new bundle version can never again "drift" out of sync with run_setup.bat.
$BUNDLE_NAMES   = @("sgml-pipeline-bundle-v12.tar.gz", "sgml-pipeline-bundle-v11.tar.gz", "sgml-pipeline-bundle-v10.tar.gz",
                    "sgml-pipeline-bundle-v9.tar.gz", "sgml-pipeline-bundle-v5.tar.gz", "sgml-pipeline-bundle-v4.tar.gz",
                    "sgml-pipeline-bundle-v3.tar.gz", "sgml-pipeline-bundle-v2.tar.gz", "sgml-pipeline-bundle.tar.gz")
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

# ── Helper: check hardware virtualisation (needed for WSL2) ───────────────────
function Test-VirtualizationEnabled {
    try {
        $cpu = Get-WmiObject -Class Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($null -ne $cpu -and $cpu.VirtualizationFirmwareEnabled -eq $true) { return $true }
    } catch {}
    # Fallback: try reading the Hyper-V requirement from systeminfo
    try {
        $si = (systeminfo 2>&1) -join "`n"
        if ($si -match "A hypervisor has been detected|Hyper-V Requirements:.*Yes") { return $true }
    } catch {}
    return $false
}

$wslInstalled = $false
try {
    $wslInfo = wsl --status 2>&1
    if ($LASTEXITCODE -eq 0) { $wslInstalled = $true }
} catch {}

# Additional check: wsl --list succeeds even when status doesn't
if (-not $wslInstalled) {
    try {
        $null = wsl --list --quiet 2>&1
        if ($LASTEXITCODE -eq 0) { $wslInstalled = $true }
    } catch {}
}

if (-not $wslInstalled) {
    Write-Warn "WSL2 not installed. Enabling now..."

    # ── Pre-flight: check BIOS virtualisation ─────────────────────────────────
    Write-Host "  Checking hardware virtualisation..." -ForegroundColor Gray
    if (-not (Test-VirtualizationEnabled)) {
        Write-Host "" 
        Write-Host "  ╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Red
        Write-Host "  ║  VIRTUALISATION IS DISABLED ON THIS MACHINE                 ║" -ForegroundColor Red
        Write-Host "  ╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Red
        Write-Host ""
        Write-Host "  WSL2 requires Intel VT-x or AMD-V to be enabled in BIOS/UEFI." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  HOW TO FIX:" -ForegroundColor Cyan
        Write-Host "    1. Restart your computer" -ForegroundColor White
        Write-Host "    2. Enter BIOS/UEFI settings during startup" -ForegroundColor White
        Write-Host "       (press Del, F2, F10, or F12 — varies by manufacturer)" -ForegroundColor Gray
        Write-Host "    3. Look for: Virtualization Technology / Intel VT-x / AMD-V" -ForegroundColor White
        Write-Host "    4. Set it to: Enabled" -ForegroundColor White
        Write-Host "    5. Save and Exit, then run this installer again" -ForegroundColor White
        Write-Host ""
        Write-Host "  If this is a corporate/managed machine, ask IT to enable" -ForegroundColor Yellow
        Write-Host "  hardware virtualisation in BIOS for your device." -ForegroundColor Yellow
        Write-Host ""
        Read-Host "  Press Enter to close"
        exit 1
    }
    Write-Host "  Hardware virtualisation: available" -ForegroundColor Gray

    # ── Try modern approach first (Windows 10 21H2+ / Windows 11) ─────────────
    Write-Host "  Enabling WSL2 features (1-2 minutes)..." -ForegroundColor Gray
    $useModern = ([int]$WinBuild -ge 20000)  # Windows 11+

    if ($useModern) {
        # Windows 11: wsl --install handles everything (no kernel MSI needed)
        $wslOut = wsl --install --no-distribution 2>&1
        $wslExit = $LASTEXITCODE
    } else {
        # Windows 10: enable features via dism
        $d1 = dism.exe /online /enable-feature /featurename:Microsoft-Windows-Subsystem-Linux /all /norestart 2>&1
        $e1 = $LASTEXITCODE
        $d2 = dism.exe /online /enable-feature /featurename:VirtualMachinePlatform /all /norestart 2>&1
        $e2 = $LASTEXITCODE
        # dism returns 3010 for "success, restart required" — treat as success
        $wslExit = if (($e1 -eq 0 -or $e1 -eq 3010) -and ($e2 -eq 0 -or $e2 -eq 3010)) { 0 } else { 1 }

        if ($wslExit -eq 0) {
            # Also download the WSL2 kernel for older Windows 10 builds
            if ([int]$WinBuild -lt 21996) {
                Write-Host "  Downloading WSL2 kernel update..." -ForegroundColor Gray
                try {
                    $KernelUrl = "https://wslstorestorage.blob.core.windows.net/wslblob/wsl_update_x64.msi"
                    $KernelMsi = "$env:TEMP\wsl_update_x64.msi"
                    Invoke-WebRequest -Uri $KernelUrl -OutFile $KernelMsi -UseBasicParsing -TimeoutSec 60
                    Start-Process msiexec.exe -ArgumentList "/i `"$KernelMsi`" /quiet" -Wait
                } catch {
                    Write-Warn "Could not download WSL2 kernel update — will retry after restart."
                }
            }
        }
    }

    if ($wslExit -ne 0) {
        Write-Host ""
        Write-Host "  ╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Red
        Write-Host "  ║  WSL2 COULD NOT BE ENABLED (exit $wslExit)                      ║" -ForegroundColor Red
        Write-Host "  ╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Red
        Write-Host ""
        Write-Host "  Common causes and fixes:" -ForegroundColor Yellow
        Write-Host "    • Virtualisation disabled in BIOS — enable Intel VT-x / AMD-V" -ForegroundColor White
        Write-Host "    • Corporate Group Policy blocking Hyper-V — contact IT" -ForegroundColor White
        Write-Host "    • Machine is a VM without nested virtualisation support" -ForegroundColor White
        Write-Host "    • Pending Windows Update must be applied first — run Windows Update" -ForegroundColor White
        Write-Host ""
        Write-Host "  Manual WSL2 install guide:" -ForegroundColor Cyan
        Write-Host "    https://learn.microsoft.com/en-us/windows/wsl/install" -ForegroundColor Cyan
        Write-Host ""
        Read-Host "  Press Enter to close"
        exit 1
    }

    wsl --set-default-version 2 2>&1 | Out-Null

    Write-Warn "RESTART REQUIRED — WSL2 feature was just enabled."
    Write-Host ""
    Write-Host "  After restart, run this installer again to continue." -ForegroundColor Cyan
    Write-Host ""
    $restart = Read-Host "  Restart now? (y/n)"
    if ($restart -eq "y") { Restart-Computer -Force }
    exit 0
} else {
    wsl --set-default-version 2 2>&1 | Out-Null
    Write-Ok "WSL2 available"
}

# ── Configure .wslconfig (idle-timeout + networking) ──────────────────────────
Write-Step "2b/7  Configuring WSL global settings (.wslconfig)..."
$WslConfigPath = "$env:USERPROFILE\.wslconfig"
$WslConfigChanged = $false
if (Test-Path $WslConfigPath) {
    $WslConfigContent = Get-Content $WslConfigPath -Raw
} else {
    $WslConfigContent = "[wsl2]`r`n"
    $WslConfigChanged = $true
}
if ($WslConfigContent -notmatch "vmIdleTimeout") {
    $WslConfigContent = $WslConfigContent.TrimEnd() + "`r`nvmIdleTimeout=-1`r`n"
    $WslConfigChanged = $true
}
# "mirrored" networking mode needs a fairly recent WSL package (>= 2.0.0) --
# forcing it on an older/managed machine can silently fail to apply, leaving
# localhost port-forwarding broken (app runs fine INSIDE WSL, but Windows can
# never reach localhost:8501 -- exactly matches "install succeeded, app never
# becomes ready" reports). Only force mirrored mode when the installed WSL
# package actually supports it; otherwise stick to the universally-supported
# default (NAT) with localhostForwarding=true, which has worked reliably for
# years on every WSL2 version.
$SupportsMirrored = $false
try {
    $wslVerRaw = ((wsl --version 2>&1) -join "`n") -replace "`0", ""
    if ($wslVerRaw -match "WSL[^\d]*(\d+)\.(\d+)\.(\d+)") {
        $verMajor = [int]$Matches[1]; $verMinor = [int]$Matches[2]
        if ($verMajor -gt 2 -or ($verMajor -eq 2 -and $verMinor -ge 0)) { $SupportsMirrored = $true }
    }
} catch {}
if ($WslConfigContent -notmatch "networkingMode") {
    if ($SupportsMirrored) {
        $WslConfigContent = $WslConfigContent.TrimEnd() + "`r`nnetworkingMode=mirrored`r`nlocalhostForwarding=true`r`n"
    } else {
        Write-Warn "WSL package too old for mirrored networking -- using default NAT mode instead"
        $WslConfigContent = $WslConfigContent.TrimEnd() + "`r`nlocalhostForwarding=true`r`n"
    }
    $WslConfigChanged = $true
} elseif (-not $SupportsMirrored -and $WslConfigContent -match "networkingMode\s*=\s*mirrored") {
    Write-Warn "WSL package too old for mirrored networking -- removing it to restore localhost access"
    $WslConfigContent = $WslConfigContent -replace "networkingMode\s*=\s*mirrored\r?\n?", ""
    if ($WslConfigContent -notmatch "localhostForwarding") {
        $WslConfigContent = $WslConfigContent.TrimEnd() + "`r`nlocalhostForwarding=true`r`n"
    }
    $WslConfigChanged = $true
}
if ($WslConfigChanged) {
    Set-Content -Path $WslConfigPath -Value $WslConfigContent -Encoding ASCII
    Write-Ok "Updated $WslConfigPath (disabled VM idle-shutdown, configured networking)"
    Write-Host "  Restarting WSL to apply..." -ForegroundColor Gray
    wsl --shutdown
    Start-Sleep -Seconds 2
} else {
    Write-Ok "WSL global settings already configured"
}

# NOTE: a Windows Defender exclusion step (Add-MpPreference) used to run here.
# Removed -- on machines managed by a PAM/privilege-elevation agent (e.g.
# Thycotic/Delinea ComElevateHost), EACH Add-MpPreference call triggers its own
# separate elevation-approval popup, so this fired 5-6 confirmation dialogs in
# a row during install -- a bad experience for a "just double-click" installer,
# and on managed machines it would be blocked by policy anyway. Not worth it.

# ── Install Ubuntu 24.04 via WSL ──────────────────────────────────────────────
Write-Step "3/7  Installing Ubuntu 24.04..."
# wsl --list output is UTF-16 LE; strip null bytes before matching
$distroRaw = wsl --list --quiet 2>&1 | ForEach-Object { $_ -replace '\x00', '' }
$ubuntuInstalled = ($distroRaw | Where-Object { $_ -match 'Ubuntu' }).Count -gt 0
if ($ubuntuInstalled) {
    Write-Ok "Ubuntu already installed — skipping"
} else {
    Write-Host "  Installing Ubuntu 24.04 (~500MB download)..." -ForegroundColor Gray
    wsl --install -d $UBUNTU_DISTRO --no-launch
    if ($LASTEXITCODE -ne 0) {
        Write-Error2 "Ubuntu install failed. Ensure internet access and run again."
    }
    # First launch to initialize (non-interactive)
    Write-Host "  Initializing Ubuntu (first run setup)..." -ForegroundColor Gray
    wsl -d $UBUNTU_DISTRO -- bash -c "echo 'Ubuntu initialized'" 2>&1 | Out-Null
    Write-Ok "Ubuntu 24.04 installed"
}

# Set as default distro
wsl --set-default $UBUNTU_DISTRO 2>&1 | Out-Null

# ── Find or create bundle ──────────────────────────────────────────────────────
Write-Step "4/7  Preparing SGML Pipeline bundle..."

if ($BundlePath) {
    Write-Ok "Using bundle path passed in: $BundlePath"
} else {
    # Look for bundle in several places, trying all known bundle filenames
    foreach ($name in $BUNDLE_NAMES) {
        foreach ($loc in @("$SCRIPT_DIR\..\$name", "$REPO_DIR\$name", "$env:USERPROFILE\Downloads\$name", "$SCRIPT_DIR\$name")) {
            if (Test-Path $loc) { $BundlePath = (Resolve-Path $loc).Path; break }
        }
        if ($BundlePath) { break }
    }
    # Catch-all: glob for ANY versioned bundle name so a future version bump can never
    # drift out of sync with this list again (this was the actual v12 installer bug).
    if (-not $BundlePath) {
        $glob = Get-ChildItem -Path $SCRIPT_DIR, "$SCRIPT_DIR\..", $REPO_DIR, "$env:USERPROFILE\Downloads" `
                    -Filter "sgml-pipeline-bundle*.tar.gz" -ErrorAction SilentlyContinue |
                Sort-Object Name -Descending | Select-Object -First 1
        if ($glob) { $BundlePath = $glob.FullName }
    }
}
$BUNDLE_NAME = if ($BundlePath) { Split-Path -Leaf $BundlePath } else { $BUNDLE_NAMES[0] }

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
        echo '[WSL] Repairing any broken packages from previous runs...'
        sudo apt-get -f install -y -qq 2>/dev/null || true
        echo '[WSL] Checking Python 3.12 availability...'
        if apt-cache show python3.12 >/dev/null 2>&1; then
            echo '[WSL] Python 3.12 available in distro repos'
        else
            echo '[WSL] Adding Python 3.12 PPA...'
            sudo apt-get install -y -qq software-properties-common 2>/dev/null
            sudo add-apt-repository -y ppa:deadsnakes/ppa 2>/dev/null
            sudo apt-get update -qq 2>/dev/null
        fi
        echo '[WSL] Running installer...'
        sudo bash /tmp/sgml-install/sgml-pipeline-bundle/scripts/setup_local_ubuntu.sh
        rm -rf /tmp/sgml-install $WslTmp
    "
} else {
    # Use repo directly
    $RepoDir = $BundlePath.Substring(5)
    $WslRepo = (wsl -- wslpath "'$RepoDir'") 2>&1
    wsl -- bash -c "sudo bash '$WslRepo/scripts/setup_local_ubuntu.sh'"
}

Write-Ok "SGML Pipeline installed in WSL"

# ── Create Windows launcher scripts ───────────────────────────────────────────
Write-Step "5/7  Creating Windows launchers..."

New-Item -ItemType Directory -Force -Path $WIN_APP_DIR | Out-Null

# start.bat
@"
@echo off
title SGML Pipeline
echo.
echo  ============================================================
echo   SGML Pipeline - Starting up
echo  ============================================================
echo.
echo  Starting services, please wait (up to 30 seconds)...
echo.

REM Use systemctl, NOT the wrapper's own "start" subcommand (bypasses systemd
REM entirely). sgml-pipeline is an ENABLED systemd service that auto-starts on
REM every WSL boot -- a raw "nohup sgml-pipeline start &" here can race that
REM auto-start and spawn a SECOND, untracked streamlit fighting the systemd-
REM managed one for port 8501 (confirmed real crash-loop). systemctl start is
REM idempotent -- a safe no-op if already running/starting.
wsl -d $UBUNTU_DISTRO -- systemctl start sgml-pipeline
call :ensure_keepalive

REM Bounded, two-stage polling loop (was an infinite "Still starting..." loop
REM with zero diagnostics -- confirmed to hang forever on a machine slower to
REM boot WSL2/ABBYY licensing for the first time than the dev box). Stage 1
REM (~1 min) shows a one-time internal status check without giving up; stage 2
REM (~6 min total) dumps full service/log diagnostics and lets the user choose
REM to keep waiting instead of staring at a blank loop with no information.
set ATTEMPTS=0
set DIAG_SHOWN=0
:check_port
timeout /t 4 /nobreak >nul 2>&1
set /a ATTEMPTS+=1
powershell -NoProfile -Command "try{`$c=New-Object Net.Sockets.TcpClient;`$c.Connect('localhost',$APP_PORT);`$c.Close();exit 0}catch{exit 1}" >nul 2>&1
if %errorlevel% == 0 goto :open
if %ATTEMPTS% EQU 15 if "%DIAG_SHOWN%"=="0" call :show_diagnostics
if %ATTEMPTS% GEQ 90 goto :long_timeout
echo  Still starting... (%ATTEMPTS%)
goto :check_port

:open
echo.
echo  App is ready!
echo.
start "" "http://localhost:$APP_PORT"
echo  Running at http://localhost:$APP_PORT
echo  Keep this window open while using the app.
echo  Press any key to STOP the app and close.
pause >nul
wsl -d $UBUNTU_DISTRO -- systemctl stop sgml-pipeline
taskkill /FI "WINDOWTITLE eq SGML Pipeline Keep-Alive*" /T /F >nul 2>&1
exit /b 0

:show_diagnostics
set DIAG_SHOWN=1
echo.
echo  This is taking longer than usual. Checking status inside WSL...
echo  ------------------------------------------------------------
wsl -d $UBUNTU_DISTRO -- bash -c "curl -s -o /dev/null http://localhost:$APP_PORT && echo Internal check: app IS responding inside WSL || echo Internal check: app not responding yet inside WSL"
wsl -d $UBUNTU_DISTRO -- systemctl is-active sgml-pipeline
echo  ------------------------------------------------------------
echo  Still working -- first-time startup can take a few minutes
echo  (ABBYY license check needs internet access). Continuing...
echo.
goto :eof

:long_timeout
echo.
echo  ============================================================
echo   Taking much longer than expected
echo  ============================================================
echo.
echo  Service status:
wsl -d $UBUNTU_DISTRO -- systemctl status sgml-pipeline --no-pager -l
echo.
echo  Recent log entries:
wsl -d $UBUNTU_DISTRO -- journalctl -u sgml-pipeline --no-pager -n 40
echo.
echo  ------------------------------------------------------------
echo   TROUBLESHOOTING
echo   1. Check your internet connection (ABBYY license needs
echo      access to account.abbyy.com:443).
echo   2. Copy the text above and send it to your administrator.
echo   3. Try restarting your computer, then run this shortcut again.
echo  ------------------------------------------------------------
echo.
echo  Press any key to keep waiting, or close this window to stop.
pause >nul
set ATTEMPTS=0
goto :check_port

:: WSL2 tears the whole VM (and the app with it) down shortly after the last
:: attached client disconnects. The nohup line above dispatches and returns
:: immediately, so without this the app dies within a minute or two of
:: start.bat launching it, even though nothing crashed. This keeps one
:: lightweight WSL client attached for as long as the app should run.
:: Idempotent -- skips if one is already running.
:ensure_keepalive
wsl -d $UBUNTU_DISTRO -- pgrep -f "sleep infinity" >nul 2>&1
if %errorlevel% equ 0 goto :eof
start "SGML Pipeline Keep-Alive" /min wsl -d $UBUNTU_DISTRO -- sleep infinity
goto :eof
"@ -replace "`r?`n","`r`n" | Set-Content "$WIN_APP_DIR\start.bat" -Encoding ASCII

# stop.bat  
@"
@echo off
title SGML Pipeline - Stop
echo Stopping SGML Pipeline...
wsl -- systemctl stop sgml-pipeline
taskkill /FI "WINDOWTITLE eq SGML Pipeline Keep-Alive*" /T /F >nul 2>&1
echo Done.
pause
"@ -replace "`r?`n","`r`n" | Set-Content "$WIN_APP_DIR\stop.bat" -Encoding ASCII

# open.bat (just opens browser, assumes app already running)
@"
@echo off
start "" "http://localhost:$APP_PORT"
"@ -replace "`r?`n","`r`n" | Set-Content "$WIN_APP_DIR\open.bat" -Encoding ASCII

Write-Ok "Launcher scripts created in $WIN_APP_DIR"

# ── Desktop shortcut ──────────────────────────────────────────────────────────
Write-Step "6/7  Creating Desktop shortcut..."

$WshShell = New-Object -ComObject WScript.Shell

# Desktop shortcut — use Windows API to handle redirected/OneDrive desktops
$DesktopPath = [Environment]::GetFolderPath("Desktop")
try {
    $Shortcut = $WshShell.CreateShortcut("$DesktopPath\SGML Pipeline.lnk")
    $Shortcut.TargetPath = "$WIN_APP_DIR\start.bat"
    $Shortcut.WorkingDirectory = $WIN_APP_DIR
    $Shortcut.Description = "Start SGML Pipeline (PDF to DOCX/SGML)"
    $Shortcut.WindowStyle = 7   # Minimized
    $Shortcut.Save()
} catch {
    Write-Warn "Could not create Desktop shortcut: $_"
}

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
# systemctl start, NOT the wrapper's own "start" subcommand -- sgml-pipeline is
# an ENABLED systemd service that ALSO auto-starts on its own on every WSL
# boot. A raw "nohup sgml-pipeline start &" here can race that auto-start and
# spawn a second, untracked streamlit fighting the systemd-managed one for
# port 8501 (confirmed real crash-loop). systemctl start is idempotent.
# NOTE: ABBYY/CodeMeter now start via their own decoupled systemd service
# (sgml-pipeline-abbyy.service) that sgml-pipeline.service does not block on --
# confirmed on 2 real machines that letting ABBYY's first-run slowness gate
# THIS service's own start transaction caused "Job for sgml-pipeline.service
# failed because a timeout was exceeded" here, every single time.
wsl -d $UBUNTU_DISTRO -- systemctl start sgml-pipeline
$startExit = $LASTEXITCODE

# Keep one lightweight WSL client attached so WSL2 doesn't tear the VM (and
# the app with it) down shortly after this installer window closes -- the
# nohup dispatch above disconnects immediately. start.bat/stop.bat manage the
# same keep-alive for subsequent launches.
Start-Process -FilePath "cmd.exe" -ArgumentList "/c start ""SGML Pipeline Keep-Alive"" /min wsl -d $UBUNTU_DISTRO -- sleep infinity" -WindowStyle Hidden

# Poll until port is open (up to ~2 minutes -- streamlit no longer waits on
# ABBYY, so this should normally succeed in a few seconds; the extra margin
# is just for a slow first-time venv/Python warm-up on some machines).
$ready = $false
for ($i = 0; $i -lt 30; $i++) {
    Start-Sleep -Seconds 4
    try {
        $tcp = New-Object Net.Sockets.TcpClient
        $tcp.Connect("localhost", $APP_PORT)
        $tcp.Close()
        $ready = $true
        break
    } catch {}
}

if ($ready) {
    Start-Process "http://localhost:$APP_PORT"
    Write-Ok "App launched in browser"
} else {
    Write-Warn "App did not respond yet (this can happen on first run)."
    Write-Host "  Double-click the 'SGML Pipeline' Desktop shortcut -- it will" -ForegroundColor Yellow
    Write-Host "  keep checking and show you exactly what's happening if it" -ForegroundColor Yellow
    Write-Host "  needs more time." -ForegroundColor Yellow
    if ($startExit -ne 0) {
        Write-Host "  (systemctl start reported an error above -- the Desktop" -ForegroundColor Yellow
        Write-Host "  shortcut's diagnostics will show the real service status.)" -ForegroundColor Yellow
    }
}

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
