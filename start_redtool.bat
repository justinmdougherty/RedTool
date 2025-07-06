@echo off
REM RedTool Python Application Launcher
REM Checks for Python and dependencies before starting

echo RedTool - BOLT Terminal ^& Configurator
echo ========================================

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.7+ from https://python.org
    pause
    exit /b 1
)

echo Python found: 
python --version

REM Check if pyserial is installed
python -c "import serial" >nul 2>&1
if errorlevel 1 (
    echo Installing required dependencies...
    python -m pip install pyserial>=3.5
    if errorlevel 1 (
        echo ERROR: Failed to install dependencies
        echo Please run: pip install pyserial
        pause
        exit /b 1
    )
)

echo Starting RedTool...
echo.

REM Start the application
python redtool.py

if errorlevel 1 (
    echo.
    echo Application exited with error
    pause
)
