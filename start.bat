@echo off
title Shipping Label Tool Starter
cd /d "%~dp0"
echo ==========================================
echo Starting Shipping Label Tool...
echo ==========================================
echo.

if not exist .venv (
    echo [ERROR] .venv folder not found!
    echo Please make sure you are in the correct directory.
    pause
    exit /b
)

.\.venv\Scripts\streamlit.exe run streamlit_app.py

if errorlevel 1 (
    echo.
    echo [ERROR] Application crashed.
    echo Please take a screenshot of this window and send it to me.
    echo.
    pause
)
