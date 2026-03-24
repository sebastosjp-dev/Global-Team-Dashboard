@echo off
cd /d "%~dp0"
echo ===================================================
echo Starting Global Business Dashboard Local Server
echo ===================================================
echo.
echo This window needs to stay open while you use the dashboard!
echo Do not close it.
echo.
PowerShell -NoProfile -ExecutionPolicy Bypass -Command "& '%~dp0server.ps1'"
pause
