@echo off
chcp 65001 >nul
echo ============================================
echo   Building format_docx.exe for Windows
echo ============================================
echo.

py -3 --version >nul 2>&1
if %errorlevel% neq 0 (
    where python >nul 2>&1
    if %errorlevel% neq 0 (
        echo [ERROR] Python not found. Install Python 3.10+ and the "py" launcher.
        echo         https://www.python.org/downloads/
        pause
        exit /b 1
    )
    echo [WARN] Using "python" on PATH; ensure the same interpreter runs PyInstaller below.
    set "PYINSTALL=python -m PyInstaller"
    echo [1/3] Installing dependencies...
    python -m pip install python-docx lxml pyinstaller pywin32
) else (
    set "PYINSTALL=py -3 -m PyInstaller"
    echo [1/3] Installing dependencies...
    py -3 -m pip install python-docx lxml pyinstaller pywin32
)
if %errorlevel% neq 0 (
    echo [ERROR] Failed to install dependencies.
    pause
    exit /b 1
)

echo.
echo [2/3] Building exe with PyInstaller...
%PYINSTALL% ^
    --noconfirm ^
    --onefile ^
    --windowed ^
    --name "FormatDocx" ^
    --add-data "format_docx.py;." ^
    --add-data "rtf_to_docx.py;." ^
    --collect-all pywin32 ^
    --hidden-import lxml._elementpath ^
    --hidden-import lxml.etree ^
    --hidden-import win32com.client ^
    --hidden-import pythoncom ^
    --hidden-import pywintypes ^
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
echo   Usage:   Double-click FormatDocx.exe, select a .docx or .rtf file.
echo            RTF requires Microsoft Word ^(COM via pywin32^).
echo.
pause
