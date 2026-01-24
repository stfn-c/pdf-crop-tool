@echo off
echo ========================================
echo   EnmeiSharkTankPitch Updater
echo ========================================
echo.

git --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Git is not installed or not in PATH
    echo Please install Git from https://git-scm.com/downloads
    pause
    exit /b 1
)

echo [1/3] Pulling latest changes...
git pull
if errorlevel 1 (
    echo ERROR: Git pull failed
    pause
    exit /b 1
)

echo [2/3] Activating venv and updating dependencies...
if not exist venv (
    echo Creating virtual environment first...
    python -m venv venv
)
call venv\Scripts\activate.bat
pip install --upgrade PyMuPDF Pillow customtkinter numpy pyinstaller

echo [3/3] Building executable...
pyinstaller --onefile --windowed --name "EnmeiSharkTankPitch" --noconfirm pdf_cropper.py

if errorlevel 1 (
    echo ERROR: Build failed
    pause
    exit /b 1
)

echo.
echo ========================================
echo   Update complete!
echo   Your exe is at: dist\EnmeiSharkTankPitch.exe
echo ========================================
echo.

explorer dist
pause
