@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo === PPT 병합기 Windows 빌드 ===
echo.

:: ── 1. LibreOffice MSI 확인 ──────────────────────────────────────
if not exist "deps" mkdir deps

if not exist "deps\LibreOffice_Win_x86-64.msi" (
    echo [정보] deps\LibreOffice_Win_x86-64.msi 파일이 없습니다.
    echo.
    echo LibreOffice MSI 파일을 아래 주소에서 다운로드하세요:
    echo   https://www.libreoffice.org/download/libreoffice-fresh/
    echo   (Windows, x86_64, MSI 선택)
    echo.
    echo 다운로드 후 파일명을 다음으로 변경하여 deps\ 폴더에 넣으세요:
    echo   LibreOffice_Win_x86-64.msi
    echo.
    set /p DOWNLOAD="PowerShell로 최신 버전 자동 다운로드할까요? [y/N]: "
    if /i "%DOWNLOAD%"=="y" (
        echo [다운로드 중] LibreOffice 최신 버전 확인 중...
        powershell -NoProfile -ExecutionPolicy Bypass -Command ^
            "$ver = (Invoke-WebRequest 'https://download.documentfoundation.org/libreoffice/stable/' -UseBasicParsing).Links.href | Where-Object {$_ -match '^\d+\.\d+\.\d+/$'} | Sort-Object -Descending | Select-Object -First 1; $ver = $ver.TrimEnd('/'); $url = \"https://download.documentfoundation.org/libreoffice/stable/$ver/win/x86_64/LibreOffice_${ver}_Win_x86-64.msi\"; Write-Host \"다운로드: $url\"; Invoke-WebRequest -Uri $url -OutFile 'deps\LibreOffice_Win_x86-64.msi' -UseBasicParsing"
        if errorlevel 1 (
            echo [오류] 자동 다운로드 실패. 위 링크에서 수동으로 다운로드해주세요.
            pause
            exit /b 1
        )
        echo [완료] LibreOffice MSI 다운로드 완료
    ) else (
        echo 수동으로 파일을 준비한 뒤 다시 실행해주세요.
        pause
        exit /b 1
    )
)

echo [확인] deps\LibreOffice_Win_x86-64.msi 존재함
echo.

:: ── 2. 이전 빌드 정리 ────────────────────────────────────────────
echo [1/3] 이전 빌드 정리 중...
if exist build rmdir /s /q build
if exist "dist\PPT병합기.exe" del /q "dist\PPT병합기.exe"

:: ── 3. PyInstaller 빌드 ──────────────────────────────────────────
echo [2/3] PyInstaller 빌드 중...
pyinstaller PPTMerger.spec

if not exist "dist\PPT병합기.exe" (
    echo [실패] PyInstaller 빌드 오류 발생
    pause
    exit /b 1
)
echo [완료] dist\PPT병합기.exe 생성됨
echo.

:: ── 4. Inno Setup 컴파일 ─────────────────────────────────────────
echo [3/3] 설치 파일 생성 중 (Inno Setup)...

:: Inno Setup 경로 탐색
set ISCC=""
if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" (
    set ISCC="C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
) else if exist "C:\Program Files\Inno Setup 6\ISCC.exe" (
    set ISCC="C:\Program Files\Inno Setup 6\ISCC.exe"
) else (
    for /f "delims=" %%i in ('where ISCC 2^>nul') do set ISCC="%%i"
)

if %ISCC%=="" (
    echo [경고] Inno Setup을 찾을 수 없습니다.
    echo   https://jrsoftware.org/isdl.php 에서 설치 후 다시 실행하세요.
    echo.
    echo   PyInstaller 빌드 결과물: dist\PPT병합기.exe
    pause
    exit /b 1
)

%ISCC% installer.iss

if exist "dist\PPT병합기_Setup_v1.0.exe" (
    echo.
    echo ========================================
    echo  [성공] 설치 파일 생성 완료!
    echo  dist\PPT병합기_Setup_v1.0.exe
    echo ========================================
) else (
    echo [실패] Inno Setup 컴파일 오류 발생
    pause
    exit /b 1
)

pause
