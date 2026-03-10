#!/bin/bash
set -e
cd "$(dirname "$0")"

echo "=== PPT 병합기 macOS 빌드 ==="

# 이전 빌드 정리
rm -rf build dist

# PyInstaller 빌드
pyinstaller PPTMerger.spec

# dist 폴더에 .app 생성됨
APP_PATH="dist/PPT병합기.app"

if [ -d "$APP_PATH" ]; then
    echo ""
    echo "✅ 빌드 성공: $APP_PATH"

    # DMG 생성 (hdiutil은 macOS 기본 내장)
    echo "📦 DMG 생성 중..."
    DMG_PATH="dist/PPT병합기_v1.0_mac.dmg"
    hdiutil create -volname "PPT병합기" \
        -srcfolder "$APP_PATH" \
        -ov -format UDZO \
        "$DMG_PATH"
    echo "✅ DMG 생성 완료: $DMG_PATH"
else
    echo "❌ 빌드 실패"
    exit 1
fi
