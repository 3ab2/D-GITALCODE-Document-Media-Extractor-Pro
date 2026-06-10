@echo off
REM ============================================================================
REM  D-GITALCODE Document Media Extractor Pro v2.0.0 - EXE Build Script
REM  Author: ELKARCH ABDELHAMID  ^|  Brand: D-GITALCODE
REM
REM  Builds a single windowed executable with the D-GITALCODE icon.
REM  Output: dist\D-GITALCODE-Document-Media-Extractor-Pro.exe
REM ============================================================================
setlocal

echo.
echo  ============================================================
echo   D-GITALCODE Document Media Extractor Pro v2.0.0 - EXE Builder
echo   Author: ELKARCH ABDELHAMID
echo  ============================================================
echo.

REM Use the project virtual environment when available.
set "PYTHON=python"
if exist "venv\Scripts\python.exe" set "PYTHON=venv\Scripts\python.exe"

echo [1/3] Checking PyInstaller...
"%PYTHON%" -m PyInstaller --version >nul 2>&1
if errorlevel 1 (
    echo        Installing PyInstaller...
    "%PYTHON%" -m pip install pyinstaller
    if errorlevel 1 (
        echo [ERROR] Failed to install PyInstaller.
        exit /b 1
    )
)

echo [2/3] Building executable (single file, no console)...
"%PYTHON%" -m PyInstaller ^
    --noconfirm ^
    --clean ^
    --onefile ^
    --windowed ^
    --name "D-GITALCODE-Document-Media-Extractor-Pro" ^
    --icon "resources\icons\logo.ico" ^
    --add-data "resources;resources" ^
    --collect-all customtkinter ^
    --collect-all tkinterdnd2 ^
    main.py
if errorlevel 1 (
    echo [ERROR] Build failed. See PyInstaller output above.
    exit /b 1
)

echo [3/3] Done.
echo.
echo  Executable: dist\D-GITALCODE-Document-Media-Extractor-Pro.exe
echo  D-GITALCODE (c) 2026 - Professional Document Media Extraction
echo.
endlocal
