@echo off
rem Double-click helper: validate NATSExcel.json against the shipped schema.
rem The script lives in the scripts subfolder (the installer puts every .ps1 there),
rem so this file can sit at the install root where the README points users.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\validate-config.ps1" %*
echo.
pause
