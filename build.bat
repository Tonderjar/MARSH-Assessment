@echo off
setlocal
title Marsh FP^&A — Build Executable

echo.
echo  ============================================
echo   MARSH  ^|  FP^&A Report Generator
echo   Building standalone Windows executable...
echo  ============================================
echo.

:: ── 1. Install / upgrade PyInstaller ────────────────────────────────────────
echo [1/2]  Checking PyInstaller...
pip install --quiet --upgrade pyinstaller
if errorlevel 1 (
    echo.
    echo  ERROR: pip failed. Make sure Python ^& pip are on your PATH.
    pause & exit /b 1
)
echo        OK.
echo.

:: ── 2. Build ─────────────────────────────────────────────────────────────────
echo [2/2]  Building executable...
echo.

pyinstaller ^
    --onefile ^
    --windowed ^
    --name "MARSH_FPA_Report_Generator" ^
    --paths src ^
    --hidden-import engine ^
    --hidden-import openpyxl.styles ^
    --hidden-import openpyxl.chart ^
    --hidden-import openpyxl.chart.bar_chart ^
    --hidden-import openpyxl.chart.line_chart ^
    --collect-all openpyxl ^
    src\main.py

if errorlevel 1 (
    echo.
    echo  ============================================
    echo   BUILD FAILED — see output above.
    echo  ============================================
    pause & exit /b 1
)

echo.
echo  ============================================
echo   SUCCESS
echo   Executable: dist\MARSH_FPA_Report_Generator.exe
echo  ============================================
echo.
pause
