@echo off
cd /d "%~dp0"
setlocal EnableExtensions
title pdf-image-ocr-to-markdown
echo ============================================================
echo   pdf-image-ocr-to-markdown
echo   Converts PDFs, images, and DOCX files in THIS folder
echo   Output goes to: md_output\
echo ============================================================
echo.

set "PROGRESS=10"
call :progress "Checking Python"
python --version >nul 2>&1
if errorlevel 1 (
    echo Python not found. Attempting automatic install...
    echo.

    winget --version >nul 2>&1
    if errorlevel 1 (
        echo ERROR: winget is not available on this machine.
        echo Please install Python manually from https://www.python.org/downloads/
        echo Make sure to check "Add Python to PATH" during install, then re-run this file.
        pause
        exit /b 1
    )

    echo Installing Python 3.11 via winget...
    winget install Python.Python.3.11 --silent --accept-package-agreements --accept-source-agreements
    if errorlevel 1 (
        echo.
        echo ERROR: Automatic Python install failed.
        echo Please install Python manually from https://www.python.org/downloads/
        pause
        exit /b 1
    )

    echo.
    echo Python installed successfully.
    echo Please close this window and run CONVERT_TO_MARKDOWN.bat again.
    pause
    exit /b 0
)

set "PROGRESS=35"
call :progress "Installing Python libraries"
python -m pip install --upgrade pip --quiet --no-warn-script-location
python -m pip install -r requirements.txt --quiet --no-warn-script-location
if errorlevel 1 (
    echo.
    echo ERROR: Failed to install required libraries.
    pause
    exit /b 1
)

set "PROGRESS=65"
call :progress "Preparing Docling OCR models"
python prepare_models.py "%~dp0docling_models"
if errorlevel 1 (
    echo.
    echo ERROR: Failed to prepare Docling models.
    pause
    exit /b 1
)

set "PROGRESS=85"
call :progress "Starting conversion"
echo Tip: Press Ctrl+C if conversion is too slow and you want to stop.
echo.
python convert_to_markdown.py
if errorlevel 1 (
    echo.
    echo ERROR: Conversion failed. See message above.
    pause
    exit /b 1
)

set "PROGRESS=100"
call :progress "Finished"
echo ============================================================
echo   Finished! Press any key to open the output folder...
echo ============================================================
pause
if exist "md_output" explorer "md_output"
exit /b 0

:progress
echo [%PROGRESS%%%] %~1
echo.
exit /b 0
