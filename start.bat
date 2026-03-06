@echo off
chcp 65001 >nul
title Nmbrs Uren Invullen

echo Opstarten...

:: Python aanwezig?
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo Python is niet geinstalleerd.
    set /p INSTALLEER="Wil je Python nu automatisch installeren? (J/N): "
    if /i "%INSTALLEER%"=="J" (
        winget --version >nul 2>&1
        if errorlevel 1 (
            echo.
            echo Automatisch installeren lukt niet op dit systeem.
            echo Download Python handmatig via: https://www.python.org/downloads/
            echo Vergeet niet "Add Python to PATH" aan te vinken!
            echo.
            pause
            exit /b 1
        )
        echo Python installeren via winget...
        winget install --id Python.Python.3 --source winget --silent --accept-package-agreements --accept-source-agreements
        python --version >nul 2>&1
        if errorlevel 1 (
            echo.
            echo Installatie mislukt. Probeer het handmatig via: https://www.python.org/downloads/
            echo.
            pause
            exit /b 1
        )
        echo Python succesvol geinstalleerd!
    ) else (
        echo.
        echo Download Python handmatig via: https://www.python.org/downloads/
        echo Vergeet niet "Add Python to PATH" aan te vinken!
        echo.
        pause
        exit /b 1
    )
)

:: pip bijwerken
python -m pip install --upgrade pip >nul 2>&1

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

:: Start! (pythonw = geen consolevenster, bat sluit direct)
pythonw nmbrs_uren_invullen.py
