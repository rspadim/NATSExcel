@echo off
rem Double-click helper: validate NATSExcel.json against the shipped schema.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0validate-config.ps1" %*
echo.
pause
