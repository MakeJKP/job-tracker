@echo off
:: ============================================================
:: build_windows.bat
:: Packages job_tracker.py into a standalone Windows .exe
:: using PyInstaller. Run this once; share the output .exe.
::
:: Requirements:
::   pip install pyinstaller python-docx
:: ============================================================

echo.
echo ============================================
echo  Job Tracker -- Windows .exe Builder
echo ============================================
echo.

:: Check Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Install from https://python.org
    pause
    exit /b 1
)

:: Install dependencies
echo Installing dependencies...
pip install pyinstaller python-docx --quiet

:: Build the exe
echo.
echo Building .exe (this takes ~30 seconds)...
pyinstaller ^
    --onefile ^
    --windowed ^
    --name "JobTracker" ^
    --add-data "windows_utils.py;." ^
    job_tracker.py

:: Result
if exist "dist\JobTracker.exe" (
    echo.
    echo ============================================
    echo  SUCCESS: dist\JobTracker.exe is ready
    echo  You can copy this single file anywhere.
    echo ============================================
) else (
    echo.
    echo ERROR: Build failed. Check output above.
)

echo.
pause
