@echo off
:: =============================================================================
:: launch_sgml_pipeline.bat
::
:: SGML Pipeline - Windows Launcher (runs app inside WSL2 Ubuntu)
::
:: REQUIREMENTS:
::   • WSL2 with Ubuntu installed (run setup_windows_wsl.ps1 first)
::   • SGML Pipeline installed in WSL2 (done by setup_windows_wsl.ps1)
::
:: USAGE:
::   Double-click this file   - or -   Call from CMD / PowerShell
::
:: The app runs entirely inside Ubuntu WSL2.
:: Your browser opens to http://localhost:8501
:: =============================================================================

title SGML Pipeline Launcher
setlocal

set APP_PORT=8501
set APP_URL=http://localhost:%APP_PORT%

echo.
echo ============================================================
echo   SGML Pipeline ^| PDF to DOCX/SGML Converter
echo ============================================================
echo.

:: ── Check WSL2 is installed ───────────────────────────────────────────────────
wsl --status >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo  ERROR: WSL2 is not installed or not running.
    echo.
    echo  Please run setup_windows_wsl.ps1 as Administrator first.
    echo  Right-click setup_windows_wsl.ps1 ^> Run with PowerShell
    echo.
    pause
    exit /b 1
)

:: ── Check if SGML Pipeline is installed in WSL ───────────────────────────────
wsl -d Ubuntu -- bash -c "test -f /usr/local/bin/sgml-pipeline" >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo  ERROR: sgml-pipeline not found in WSL Ubuntu.
    echo.
    echo  Please run setup_windows_wsl.ps1 as Administrator first.
    echo  Right-click setup_windows_wsl.ps1 ^> Run with PowerShell
    echo.
    pause
    exit /b 1
)

:: ── Check if already running ──────────────────────────────────────────────────
wsl -d Ubuntu -- pgrep -f "streamlit run" >nul 2>&1
if %ERRORLEVEL% equ 0 (
    echo  App is already running. Opening browser...
    timeout /t 1 /nobreak >nul
    start "" "%APP_URL%"
    echo.
    echo  Access: %APP_URL%
    echo.
    goto :end
)

:: ── Start the app ─────────────────────────────────────────────────────────────
echo  Starting ABBYY LicensingService and Streamlit app...
echo  This takes ~10 seconds on first launch.
echo.

:: Start inside WSL2 using nohup so Streamlit survives after the WSL session exits.
:: Without nohup the process gets SIGHUP-killed when the calling shell closes.
wsl -d Ubuntu -- bash -c "nohup sgml-pipeline start > /tmp/sgml-pipeline.log 2>&1 & sleep 2 && echo 'started' > /tmp/sgml-pipeline.pid"

:: Wait for startup
echo  Waiting for app to start.
set /a i=0
:wait_loop
    timeout /t 2 /nobreak >nul
    set /a i+=1
    echo  . (%i%0s elapsed)
    wsl -d Ubuntu -- curl -sf "http://localhost:%APP_PORT%" >nul 2>&1
    if %ERRORLEVEL% equ 0 goto :ready
    if %i% geq 15 goto :timeout_warn
    goto :wait_loop

:timeout_warn
    echo.
    echo  App may still be starting. Opening browser anyway...
    goto :open_browser

:ready
    echo.
    echo  App is ready!

:open_browser
    :: Open browser
    start "" "%APP_URL%"
    echo.
    echo ============================================================
    echo   SGML Pipeline is running at %APP_URL%
    echo ============================================================
    echo.
    echo  Tips:
    echo    - Upload a PDF file in the browser
    echo    - Select conversion type (DOCX / SGML)
    echo    - Download the converted file
    echo.
    echo  To STOP the app, run: stop_sgml_pipeline.bat
    echo  Or from PowerShell:   wsl -- sgml-pipeline stop
    echo.

:end
pause
exit /b 0
