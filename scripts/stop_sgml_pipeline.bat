@echo off
:: =============================================================================
:: stop_sgml_pipeline.bat
:: SGML Pipeline - Stop all services in WSL2
:: =============================================================================
title SGML Pipeline - Stop
echo.
echo Stopping SGML Pipeline...
wsl -- bash -c "sgml-pipeline stop"
echo.
echo Done. App has been stopped.
echo.
pause
