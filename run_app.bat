@echo off
cd /d "%~dp0"
if not exist "%~dp0.tmp" mkdir "%~dp0.tmp"
set "TEMP=%~dp0.tmp"
set "TMP=%~dp0.tmp"
set OPENBLAS_NUM_THREADS=1
set OMP_NUM_THREADS=1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0start_app.ps1"
if errorlevel 1 pause

