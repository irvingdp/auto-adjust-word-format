@echo off
chcp 65001 >nul
echo ============================================
echo   Building format_docx.exe for Windows
echo ============================================
echo.

where py >nul 2>&1
if not "%errorlevel%"=="0" (
    echo [ERROR] Python Launcher 'py' not found. Please install Python 3.10+ first.
    echo         https://www.python.org/downloads/
    pause
    exit /b 1
)

set "PY=py -3.12"

echo Cleaning old build outputs...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist FormatDocx.spec del /q FormatDocx.spec

echo [1/3] Installing dependencies...
%PY% -m pip install -r requirements.txt
if not "%errorlevel%"=="0" (
    echo [ERROR] Failed to install dependencies.
    pause
    exit /b 1
)

echo.
echo [2/3] Building exe with PyInstaller...
%PY% -m PyInstaller ^
    --noconfirm ^
    --onefile ^
    --windowed ^
    --name "FormatDocx" ^
    --add-data "format_docx.py;." ^
    format_docx_gui.py

if not "%errorlevel%"=="0" (
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
exit /b 0
