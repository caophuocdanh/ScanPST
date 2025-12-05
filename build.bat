@echo off
REM ========================================
REM SCANPST Auto Repair Tool - Build EXE
REM ========================================

echo.
echo ========================================
echo  SCANPST - Build Executable
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python khong tim thay hoac khong co trong PATH
    echo Vui long cai dat Python tu: https://www.python.org/
    echo Nhap Enter de thoat...
    pause
    exit /b 1
)

echo [OK] Python da duoc phat hien
python --version
echo.

pip install -r requirements.txt

REM Check if PyInstaller is installed
python -c "import PyInstaller" >nul 2>&1
if %errorlevel% neq 0 (
    echo [*] PyInstaller chua duoc cai dat
    echo [*] Dang cai dat PyInstaller...
    pip install pyinstaller
    if %errorlevel% neq 0 (
        echo.
        echo ERROR: Co loi khi cai dat PyInstaller!
        echo Nhap Enter de thoat...
        pause
        exit /b 1
    )
)

echo [OK] PyInstaller da san sang
echo.

REM Build the executable
echo [*] Dang build executable...
echo.

REM Build with PyInstaller
pyinstaller --onefile --windowed --icon=icon.ico --add-data "icon.ico;." --name "SCANPST" scanpst.py

if %errorlevel% neq 0 (
    echo.
    echo ERROR: Co loi trong qua trinh build!
    echo Nhap Enter de thoat...
    pause
    exit /b 1
)

REM Clean up build artifacts
echo.
echo [*] Dang xoa cac file tam...

REM Delete SCANPST.spec
if exist SCANPST.spec (
    del SCANPST.spec >nul 2>&1
    echo [OK] Xoa SCANPST.spec
)

REM Delete build directory
if exist build (
    rmdir /s /q build >nul 2>&1
    echo [OK] Xoa thu muc build
)

echo.
echo ========================================
echo [OK] BUILD THANH CONG!
echo ========================================
echo.
echo Executable file:
echo   dist\SCANPST.exe
echo.
echo Ban co the chay file nay bang:
echo   dist\SCANPST.exe
echo   hoac Double-click dist\SCANPST.exe
echo.
