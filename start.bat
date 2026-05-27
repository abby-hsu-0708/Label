@echo off
title MISUMI Label Tool Starter
cd /d "%~dp0"
echo ==========================================
echo Starting MISUMI Label Tool...
echo ==========================================
echo.

if not exist .venv (
    echo [ERROR] .venv folder not found!
    echo Please make sure you are in the correct directory.
    pause
    exit /b
)

.\.venv\Scripts\python.exe main.py

if errorlevel 1 (
    echo.
    echo [ERROR] Application crashed.
    echo Please take a screenshot of this window and send it to me.
    echo.
    pause
)
