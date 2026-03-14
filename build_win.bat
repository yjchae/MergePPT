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
    set /p DOWNLOAD="PowerShell로 최신 버전 자동 다운로드할까요? [y/N]: "
    if /i "%DOWNLOAD%"=="y" (
        echo [다운로드 중] LibreOffice 최신 버전 확인 중...
        powershell -NoProfile -ExecutionPolicy Bypass -Command ^
            "$ver = (Invoke-WebRequest 'https://download.documentfoundation.org/libreoffice/stable/' -UseBasicParsing).Links.href | Where-Object {$_ -match '^\d+\.\d+\.\d+/$'} | Sort-Object -Descending | Select-Object -First 1; $ver = $ver.TrimEnd('/'); $url = \"https://download.documentfoundation.org/libreoffice/stable/$ver/win/x86_64/LibreOffice_${ver}_Win_x86-64.msi\"; Write-Host \"다운로드: $url\"; Invoke-WebRequest -Uri $url -OutFile 'deps\LibreOffice_Win_x86-64.msi' -UseBasicParsing"
        if errorlevel 1 (
            echo [오류] 자동 다운로드 실패.
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

:: ── 2. 아이콘 생성 (없을 경우) ───────────────────────────────────
if not exist "icon.ico" (
    echo [정보] icon.ico 없음, 자동 생성 중...
    pip install pillow -q
    python -c "from PIL import Image, ImageDraw; img = Image.new('RGBA', (256,256), (0,99,204,255)); draw = ImageDraw.Draw(img); draw.rectangle([40,60,216,196], fill=(255,255,255,255)); draw.rectangle([60,80,196,176], fill=(0,99,204,255)); img.save('icon.ico', format='ICO', sizes=[(256,256),(128,128),(64,64),(32,32),(16,16)]); print('icon.ico 생성 완료')"
) else (
    echo [확인] icon.ico 존재함
)
echo.

:: ── 3. 이전 빌드 정리 ────────────────────────────────────────────
echo [1/3] 이전 빌드 정리 중...
if exist build rmdir /s /q build
if exist "dist\PPT병합기.exe" del /q "dist\PPT병합기.exe"

:: ── 4. PyInstaller 빌드 ──────────────────────────────────────────
echo [2/3] PyInstaller 빌드 중...
pyinstaller PPTMerger.spec

if not exist "dist\PPT병합기.exe" (
    echo [실패] PyInstaller 빌드 오류 발생
    pause
    exit /b 1
)
echo [완료] dist\PPT병합기.exe 생성됨
echo.

:: ── 5. Inno Setup 컴파일 ─────────────────────────────────────────
echo [3/3] 설치 파일 생성 중 (Inno Setup)...

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