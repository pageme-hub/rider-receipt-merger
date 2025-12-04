@echo off
chcp 65001 >nul
echo ========================================
echo Rider Receipt Merger 빌드 시작...
echo ========================================
echo.

REM PyInstaller 설치 확인
python -c "import PyInstaller" 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo PyInstaller가 설치되지 않았습니다. 설치 중...
    pip install pyinstaller
    if %ERRORLEVEL% NEQ 0 (
        echo PyInstaller 설치 실패!
        pause
        exit /b 1
    )
)

echo.
echo 빌드 옵션:
echo - 실행 파일명: RiderReceiptMerger.exe
echo - 단일 파일 모드: 활성화
echo - GUI 모드: 활성화
echo.

REM 기존 빌드 파일 정리
if exist dist\RiderReceiptMerger.exe (
    echo 기존 실행 파일 삭제 중...
    del /f /q dist\RiderReceiptMerger.exe
)

if exist build (
    echo 기존 빌드 폴더 삭제 중...
    rmdir /s /q build
)

echo.
echo PyInstaller 실행 중...
echo.

pyinstaller --name="RiderReceiptMerger" --windowed --onefile --icon=NONE --add-data="config;config" --add-data="mapping_config.json;." --hidden-import=openpyxl --hidden-import=msoffcrypto --hidden-import=PySide6 --collect-all=openpyxl --collect-all=msoffcrypto main.py

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================
    echo 빌드 완료!
    echo ========================================
    echo.
    echo 실행 파일 위치: dist\RiderReceiptMerger.exe
    echo.
    echo 배포 시 다음 파일들도 함께 복사하세요:
    echo   - config\ 폴더
    echo   - mapping_config.json
    echo.
) else (
    echo.
    echo ========================================
    echo 빌드 실패!
    echo ========================================
    echo.
    echo 오류가 발생했습니다. 위의 오류 메시지를 확인하세요.
    echo.
)

pause

