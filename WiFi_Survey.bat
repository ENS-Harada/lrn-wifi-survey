@echo off
chcp 65001 >nul 2>&1
title LRN WiFi Survey Tool
echo.
echo  ================================================
echo    LRN Wi-Fi Survey Tool
echo    Starting...
echo  ================================================
echo.

set "PS_SCRIPT=%~dp0WiFi_Survey.ps1"

if not exist "%PS_SCRIPT%" (
    echo  [ERROR] WiFi_Survey.ps1 not found.
    echo  Place this bat file and WiFi_Survey.ps1 in the same folder.
    echo.
    pause
    exit /b 1
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%"

echo.
pause
