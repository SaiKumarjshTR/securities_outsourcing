@echo off
:: ============================================================
:: SGML Pipeline — Installer Launcher
::
:: HOW TO USE:
::   1. Download sgml-pipeline-bundle-v12.tar.gz from SharePoint
::   2. Put this file in the SAME folder as the bundle
::   3. Double-click this file -> click Yes on the UAC prompt
::   4. Wait — everything installs automatically
:: ============================================================
setlocal enabledelayedexpansion

:: ── If re-launched elevated by VBScript, go straight to install ──
if /i "%~1"=="ELEVATED" goto :run_as_admin

:: ── First run: elevate via VBScript, passing ELEVATED marker ──
set "VBS=%TEMP%\sgml_elevate_%RANDOM%.vbs"
echo Set oShell = CreateObject("Shell.Application")                        > "%VBS%"
echo oShell.ShellExecute "cmd.exe", "/c ""%~s0"" ELEVATED", "%~sdp0", "runas", 1 >> "%VBS%"
cscript //nologo "%VBS%"
del "%VBS%" 2>nul
exit /b

:run_as_admin
cd /d "%~dp0"

echo.
echo  ============================================================
echo   SGML Pipeline — Windows Installer
echo  ============================================================
echo.

:: ── Locate the bundle file ─────────────────────────────────────
set "BUNDLE="

for %%F in (
    "%~dp0sgml-pipeline-bundle.tar.gz"
    "%USERPROFILE%\Downloads\sgml-pipeline-bundle.tar.gz"
    "%~dp0sgml-pipeline-bundle-v12.tar.gz"
    "%~dp0sgml-pipeline-bundle-v11.tar.gz"
    "%~dp0sgml-pipeline-bundle-v10.tar.gz"
    "%~dp0sgml-pipeline-bundle-v9.tar.gz"
    "%~dp0sgml-pipeline-bundle-v5.tar.gz"
    "%~dp0sgml-pipeline-bundle-v4.tar.gz"
    "%~dp0sgml-pipeline-bundle-v3.tar.gz"
    "%USERPROFILE%\Downloads\sgml-pipeline-bundle-v12.tar.gz"
    "%USERPROFILE%\Downloads\sgml-pipeline-bundle-v11.tar.gz"
    "%USERPROFILE%\Downloads\sgml-pipeline-bundle-v10.tar.gz"
    "%USERPROFILE%\Downloads\sgml-pipeline-bundle-v9.tar.gz"
    "%USERPROFILE%\Downloads\sgml-pipeline-bundle-v5.tar.gz"
    "%USERPROFILE%\Downloads\sgml-pipeline-bundle-v4.tar.gz"
    "%USERPROFILE%\Downloads\sgml-pipeline-bundle-v3.tar.gz"
) do (
    if "!BUNDLE!"=="" if exist %%F set "BUNDLE=%%~fF"
)

if "!BUNDLE!"=="" (
    echo.
    echo  ERROR: Cannot find sgml-pipeline-bundle-v12.tar.gz
    echo.
    echo  Make sure RUN_SETUP.bat and sgml-pipeline-bundle-v12.tar.gz
    echo  are in the SAME folder, then double-click RUN_SETUP.bat again.
    echo.
    pause
    exit /b 1
)

echo  Found bundle: !BUNDLE!
echo  Extracting installer script...
echo.

:: ── Extract setup_windows_wsl.ps1 from the bundle ─────────────
set "TMP_DIR=%TEMP%\sgml_setup_%RANDOM%"
mkdir "%TMP_DIR%"

tar -xzf "!BUNDLE!" -C "%TMP_DIR%" "sgml-pipeline-bundle/scripts/setup_windows_wsl.ps1"
if %errorlevel% neq 0 (
    echo.
    echo  ERROR: Failed to extract installer from bundle.
    echo  The file may be corrupted. Re-download and try again.
    echo.
    pause
    exit /b 1
)

:: Symlink bundle where PS1 expects it (../bundle relative to scripts/)
mklink "%TMP_DIR%\sgml-pipeline-bundle\sgml-pipeline-bundle.tar.gz" "!BUNDLE!" >nul 2>&1

:: ── Run the main installer ────────────────────────────────────
:: Pass the resolved bundle path through explicitly — the PS1's own name-based
:: search only knew about older bundle versions and would fail to find it otherwise.
powershell -ExecutionPolicy Bypass -File "%TMP_DIR%\sgml-pipeline-bundle\scripts\setup_windows_wsl.ps1" -BundlePath "!BUNDLE!"

rmdir /S /Q "%TMP_DIR%" 2>nul

echo.
pause
