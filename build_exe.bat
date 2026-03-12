@echo off
chcp 65001 >nul
echo ============================================
echo   Building format_docx.exe for Windows
echo ============================================
echo.

where python >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python not found. Please install Python 3.10+ first.
    echo         https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [1/3] Installing dependencies...
pip install python-docx lxml pyinstaller
if %errorlevel% neq 0 (
    echo [ERROR] Failed to install dependencies.
    pause
    exit /b 1
)

echo.
echo [2/3] Building exe with PyInstaller...
pyinstaller ^
    --noconfirm ^
    --onefile ^
    --windowed ^
    --name "FormatDocx" ^
    --add-data "format_docx.py;." ^
    --hidden-import lxml._elementpath ^
    --hidden-import lxml.etree ^
    format_docx_gui.py

if %errorlevel% neq 0 (
    echo [ERROR] PyInstaller build failed.
    pause
    exit /b 1
)

echo.
echo [3/3] Done!
echo.
echo   Output:  dist\FormatDocx.exe
echo.
echo   Usage:   Double-click FormatDocx.exe, select a .docx file,
echo            and the adjusted file will be saved next to the original.
echo.
pause
