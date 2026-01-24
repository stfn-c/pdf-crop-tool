@echo off
echo ========================================
echo   EnmeiSharkTankPitch Windows Builder
echo ========================================
echo.

:: Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo.
    echo Please install Python 3.10+ from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation
    pause
    exit /b 1
)

echo [1/4] Creating virtual environment...
python -m venv venv
if errorlevel 1 (
    echo ERROR: Failed to create virtual environment
    pause
    exit /b 1
)

echo [2/4] Activating venv and installing dependencies...
call venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install PyMuPDF Pillow customtkinter numpy pyinstaller

if errorlevel 1 (
    echo ERROR: Failed to install dependencies
    pause
    exit /b 1
)

echo [3/4] Building executable...
pyinstaller --onefile --windowed --name "EnmeiSharkTankPitch" --noconfirm pdf_cropper.py

if errorlevel 1 (
    echo ERROR: Build failed
    pause
    exit /b 1
)

echo [4/4] Done!
echo.
echo ========================================
echo   Build complete!
echo   Your exe is at: dist\EnmeiSharkTankPitch.exe
echo ========================================
echo.

:: Open the dist folder
explorer dist

pause
