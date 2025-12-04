@echo off
echo ========================================
echo Building Rider Receipt Merger...
echo ========================================
echo.

REM Check if PyInstaller is installed
python -c "import PyInstaller" 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo PyInstaller is not installed. Installing...
    pip install pyinstaller
    if %ERRORLEVEL% NEQ 0 (
        echo Failed to install PyInstaller!
        pause
        exit /b 1
    )
)

echo.
echo Build options:
echo - Executable name: RiderReceiptMerger.exe
echo - Single file mode: Enabled
echo - GUI mode: Enabled
echo.

REM Clean up existing build files
if exist dist\RiderReceiptMerger.exe (
    echo Deleting existing executable...
    del /f /q dist\RiderReceiptMerger.exe
)

if exist build (
    echo Deleting existing build folder...
    rmdir /s /q build
)

echo.
echo Running PyInstaller...
echo.

pyinstaller --name="RiderReceiptMerger" --windowed --onefile --icon=NONE --add-data="config;config" --add-data="mapping_config.json;." --hidden-import=openpyxl --hidden-import=msoffcrypto --hidden-import=PySide6 --collect-all=openpyxl --collect-all=msoffcrypto main.py

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================
    echo Build completed successfully!
    echo ========================================
    echo.
    echo Executable location: dist\RiderReceiptMerger.exe
    echo.
    echo When deploying, also copy these files:
    echo   - config\ folder
    echo   - mapping_config.json
    echo.
) else (
    echo.
    echo ========================================
    echo Build failed!
    echo ========================================
    echo.
    echo An error occurred. Please check the error messages above.
    echo.
)

pause
