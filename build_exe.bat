@echo off
echo ============================================
echo  DocType Billing Report - EXE Builder
echo ============================================

where python >nul 2>nul
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH.
    echo Install Python 3.10+ from https://www.python.org/downloads/ and re-run this file.
    pause
    exit /b 1
)

echo Installing required packages...
pip install --upgrade pip
pip install -r requirements.txt

echo.
echo Building DocTypeBillingReport.exe ...
pyinstaller --onefile --windowed --name SPS_TDM_DocTypeBillingReport --icon=Swift_Prosys.ico --add-data "Swift-Prosys.png;." --add-data "Swift_Prosys.ico;." --add-data "bg_dark.png;." --add-data "bg_light.png;." --collect-all customtkinter app.py

echo.
echo ============================================
echo  DONE! Your exe is in the "dist" folder:
echo  dist\DocTypeBillingReport.exe
echo ============================================
pause
