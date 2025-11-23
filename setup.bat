@echo off
setlocal enabledelayedexpansion

echo ========================================
echo  ODIN - Setup Script
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python not found!
    echo.
    choice /C YN /M "Download and install Python automatically"
    if !errorlevel! equ 1 (
        echo.
        echo Downloading Python installer...
        set PYTHON_VERSION=3.12.0
        set PYTHON_URL=https://www.python.org/ftp/python/!PYTHON_VERSION!/python-!PYTHON_VERSION!-amd64.exe
        set INSTALLER=%TEMP%\python_installer.exe
        
        curl -L -o "!INSTALLER!" "!PYTHON_URL!"
        
        echo Installing Python !PYTHON_VERSION! with PATH...
        "!INSTALLER!" /quiet InstallAllUsers=0 PrependPath=1 Include_test=0
        
        echo Cleaning up...
        del "!INSTALLER!"
        
        echo.
        echo Python installed. Please restart this script.
        pause
        exit /b
    ) else (
        echo Please install Python manually from https://www.python.org/downloads/
        pause
        exit /b 1
    )
)

echo Python found: 
python --version
echo.

echo Upgrading pip...
pip install --upgrade pip

echo.
echo ========================================
echo Installing Core Dependencies...
echo ========================================
echo.

REM Data processing packages
echo [1/8] Installing pandas...
pip install pandas

echo [2/8] Installing openpyxl...
pip install openpyxl

echo [3/8] Installing xlrd...
pip install xlrd

REM Windows integration
echo [4/8] Installing pywin32...
pip install pywin32

REM Image processing
echo [5/8] Installing Pillow...
pip install pillow

REM UI packages for original version
echo [6/8] Installing tkcalendar...
pip install tkcalendar

REM Modern UI framework
echo [7/8] Installing customtkinter...
pip install customtkinter

REM Additional dependencies for customtkinter
echo [8/8] Installing darkdetect...
pip install darkdetect

echo.
echo ========================================
echo Installation Complete!
echo ========================================
pause