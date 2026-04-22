@echo off
echo ========================================
echo   Smart Factory Dashboard - Build EXE
echo ========================================
echo.

python --version 2>nul
if errorlevel 1 (
    echo [ERROR] Python chua duoc cai dat hoac khong trong PATH
    pause & exit /b 1
)

echo [1/3] Cai dat dependencies...
pip install pyqt5 matplotlib pandas xlrd openpyxl pyinstaller Pillow --quiet
if errorlevel 1 (
    echo [ERROR] Cai dat that bai
    pause & exit /b 1
)

echo [2/3] Build EXE...
pyinstaller --clean SmartFactory.spec

if errorlevel 1 (
    echo [ERROR] Build that bai
    pause & exit /b 1
)

echo [3/3] Hoan thanh!
echo File EXE: dist\SmartFactory.exe
dir dist\SmartFactory.exe | find "SmartFactory.exe"
echo.
pause
