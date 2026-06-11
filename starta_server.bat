@echo off
title Serviceprotokoll – Liljegrens
cd /d "%~dp0"

echo Kontrollerar Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo FEL: Python hittades inte.
    echo Installera Python fran https://www.python.org och borja om.
    pause
    exit /b 1
)

echo Installerar/kontrollerar beroenden...
python -m pip install -q flask openpyxl

echo.
echo ============================================
echo   Serviceprotokoll-server startar...
echo ============================================
echo.
python server.py
pause
