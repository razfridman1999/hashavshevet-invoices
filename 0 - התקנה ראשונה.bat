@echo off
chcp 65001 >nul
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo.
    echo Python is not installed on this computer.
    echo Install it from https://www.python.org/downloads/  and tick "Add python.exe to PATH".
    echo Then run this file again.
    pause
    exit /b 1
)

python -c "import fitz, openpyxl, win32com.client" 2>nul
if errorlevel 1 (
    echo Installing the components the tool needs...
    python -m pip install --user -r requirements.txt
    if errorlevel 1 (
        echo.
        echo The installation failed. Take a screenshot of this window and send it to Raz.
        pause
        exit /b 1
    )
)

python -c "import fitz, openpyxl, win32com.client; print('OK')"
python run.py setup
pause
