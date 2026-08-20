@echo off
setlocal
cd /d "%~dp0"

if not exist "runtime\python.exe" (
    echo Portable Python runtime not found.
    echo This package must be built by GitHub Actions first.
    pause
    exit /b 1
)

runtime\python.exe app\shipment_gui.py
if errorlevel 1 (
    echo.
    echo Shipment Automation stopped with an error.
    pause
)
