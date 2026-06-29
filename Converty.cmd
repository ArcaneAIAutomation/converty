@echo off
REM Converty launcher - runs the GUI in Windows PowerShell 5.1 (STA mode).
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Converty.ps1"
