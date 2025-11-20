@echo off
REM Launcher for D365 Python Comparison Tool
REM Double-click this file to run the application

title D365 Python Comparison Tool

REM Check if Python is available
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo ERROR: Python not found!
    echo.
    echo Please install Python 3.8 or higher from:
    echo https://www.python.org/downloads/
    echo.
    echo After installing Python, run setup.bat first.
    echo.
    pause
    exit /b 1
)

REM Check if dependencies are installed
python -c "import requests" >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo ERROR: Required packages not installed!
    echo.
    echo Please run setup.bat first to install dependencies.
    echo.
    pause
    exit /b 1
)

REM Run the main application
cls
python main.py

REM Keep window open if there was an error
if %errorlevel% neq 0 (
    echo.
    echo.
    echo Application exited with an error.
    pause
)
