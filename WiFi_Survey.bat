@echo off
chcp 65001 >nul 2>&1
title LRN WiFi Survey Tool

set "PS_SCRIPT=%~dp0WiFi_Survey.ps1"

if not exist "%PS_SCRIPT%" (
    echo  [ERROR] WiFi_Survey.ps1 not found.
    echo  Place this bat file and WiFi_Survey.ps1 in the same folder.
    pause
    exit /b 1
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "& '%PS_SCRIPT%'"
