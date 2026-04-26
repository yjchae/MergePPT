#!/bin/bash
set -e
cd "$(dirname "$0")"

echo "=== PPT 병합기 macOS 빌드 ==="
echo ""

# ── 1. LibreOffice 확인 ───────────────────────────────────────────
LO_APP="/Applications/LibreOffice.app"

if [ ! -d "$LO_APP" ]; then
    echo "⚠️  LibreOffice가 설치되어 있지 않습니다."
    echo "   빌드 머신에 LibreOffice가 있어야 앱 번들에 포함됩니다."
    echo ""
    read -p "LibreOffice를 지금 다운로드하여 설치할까요? [y/N]: " INSTALL_LO
    if [[ "$INSTALL_LO" =~ ^[Yy]$ ]]; then
        echo "[다운로드] LibreOffice 최신 버전 확인 중..."
        # 최신 안정 버전 URL 조회
        LO_VER=$(curl -s https://download.documentfoundation.org/libreoffice/stable/ \
            | grep -oE '[0-9]+\.[0-9]+\.[0-9]+/' | sort -V | tail -1 | tr -d '/')
        LO_DMG="LibreOffice_${LO_VER}_MacOS_aarch64.dmg"
        LO_URL="https://download.documentfoundation.org/libreoffice/stable/${LO_VER}/mac/aarch64/${LO_DMG}"

        echo "[다운로드] $LO_URL"
        curl -L -o "/tmp/${LO_DMG}" "$LO_URL"

        echo "[설치] LibreOffice DMG 마운트 중..."
        hdiutil attach "/tmp/${LO_DMG}" -mountpoint /Volumes/LibreOffice -nobrowse -quiet
        cp -R "/Volumes/LibreOffice/LibreOffice.app" /Applications/
        hdiutil detach /Volumes/LibreOffice -quiet
        rm "/tmp/${LO_DMG}"
        echo "✅ LibreOffice 설치 완료"
    else
        echo ""
        echo "LibreOffice 없이 빌드를 계속합니다."
        echo "(.ppt 변환 기능은 앱 번들에 포함되지 않습니다)"
    fi
fi

echo ""

# ── 2. 이전 빌드 정리 ────────────────────────────────────────────
echo "[1/3] 이전 빌드 정리 중..."
rm -rf build dist

# ── 3. PyInstaller 빌드 ──────────────────────────────────────────
echo "[2/3] PyInstaller 빌드 중..."
if [ -d "$LO_APP" ]; then
    echo "      LibreOffice.app 번들 포함 (용량이 커서 시간이 걸릴 수 있습니다)"
fi
pyinstaller PPTMerger.spec

APP_PATH="dist/PPT병합기.app"
if [ ! -d "$APP_PATH" ]; then
    echo "❌ 빌드 실패"
    exit 1
fi
echo "✅ 빌드 성공: $APP_PATH"
echo ""

# ── 4. DMG 생성 ──────────────────────────────────────────────────
echo "[3/3] DMG 생성 중..."
DMG_PATH="dist/PPT병합기_v1.0_mac.dmg"

# 앱 번들 크기 측정 (용량에 따라 DMG 여유 공간 조정)
APP_SIZE_KB=$(du -sk "$APP_PATH" | awk '{print $1}')
DMG_SIZE_MB=$(( (APP_SIZE_KB / 1024) + 100 ))

hdiutil create -volname "PPT병합기" \
    -srcfolder "$APP_PATH" \
    -ov -format UDZO \
    "$DMG_PATH"

echo ""

# ── 5. 코드 서명 (ad-hoc) ─────────────────────────────────────────
echo "[+] 코드 서명 중..."

# LibreOffice Info.plist 심볼릭 링크 → 실제 파일로 교체
LO_BUNDLE="$APP_PATH/Contents/Frameworks/LibreOffice.app"
if [ -d "$LO_BUNDLE" ]; then
    find "$LO_BUNDLE" -type l | while read symlink; do
        real=$(readlink -f "$symlink" 2>/dev/null || true)
        if [ -n "$real" ] && [ -f "$real" ]; then
            cp -f "$real" "$symlink.tmp" && mv -f "$symlink.tmp" "$symlink"
        fi
    done
    # LibreOffice 내부 바이너리 서명
    find "$LO_BUNDLE/Contents/MacOS" -type f | while read f; do
        codesign --force --sign "-" "$f" 2>/dev/null || true
    done
    find "$LO_BUNDLE/Contents/Frameworks" -name "*.dylib" -type f | while read f; do
        codesign --force --sign "-" "$f" 2>/dev/null || true
    done
    codesign --force --sign "-" "$LO_BUNDLE" 2>/dev/null || true
fi

# 앱 전체 서명
codesign --force --sign "-" "$APP_PATH" 2>&1 | grep -v "replacing existing" || true
echo "✅ 서명 완료"
echo ""

# ── 6. DMG 재생성 (서명된 앱으로) ────────────────────────────────
echo "[재생성] 서명된 앱으로 DMG 갱신 중..."
hdiutil create -volname "PPT병합기" \
    -srcfolder "$APP_PATH" \
    -ov -format UDZO \
    "$DMG_PATH"

echo ""
echo "========================================"
echo " ✅ 빌드 완료!"
echo "    $DMG_PATH"
APP_SIZE_MB=$((APP_SIZE_KB / 1024))
echo "    앱 크기: ~${APP_SIZE_MB}MB"
echo "========================================"
