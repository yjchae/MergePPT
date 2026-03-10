@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo === PPT 병합기 Windows 빌드 ===

:: 이전 빌드 정리
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

:: PyInstaller 빌드
pyinstaller PPTMerger.spec

if exist "dist\PPT병합기.exe" (
    echo.
    echo [성공] dist\PPT병합기.exe 생성 완료
) else (
    echo [실패] 빌드 오류 발생
    exit /b 1
)

pause
