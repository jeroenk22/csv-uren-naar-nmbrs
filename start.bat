@echo off
chcp 65001 >nul
title Nmbrs Uren Invullen

echo Opstarten...

:: Python aanwezig?
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo Python is niet geinstalleerd.
    echo Download via: https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

:: python-dotenv aanwezig?
python -c "import dotenv" >nul 2>&1
if errorlevel 1 (
    echo python-dotenv installeren...
    pip install python-dotenv
)

:: playwright aanwezig?
python -c "import playwright" >nul 2>&1
if errorlevel 1 (
    echo playwright installeren...
    pip install playwright
)

:: Chromium aanwezig?
python -c "from playwright.sync_api import sync_playwright; import os; p=sync_playwright().start(); ok=os.path.exists(p.chromium.executable_path); p.stop(); exit(0 if ok else 1)" >nul 2>&1
if errorlevel 1 (
    echo Chromium installeren...
    playwright install chromium
)

:: Start!
python nmbrs_uren_invullen.py
