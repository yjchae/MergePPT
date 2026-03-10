# -*- mode: python ; coding: utf-8 -*-
import sys, os

block_cipher = None
APP_NAME = 'PPT병합기'

a = Analysis(
    ['mergeppt.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'PySide6.QtWidgets',
        'PySide6.QtCore',
        'PySide6.QtGui',
        'PySide6.QtSvg',
        'lxml',
        'lxml.etree',
        'pptx',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

if sys.platform == 'darwin':
    # ── macOS ──────────────────────────────────────────────
    exe = EXE(
        pyz, a.scripts, [],
        exclude_binaries=True,
        name=APP_NAME,
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=True,
        console=False,
        icon='icon.icns' if os.path.exists('icon.icns') else None,
    )
    coll = COLLECT(
        exe, a.binaries, a.zipfiles, a.datas,
        strip=False, upx=True,
        name=APP_NAME,
    )
    app = BUNDLE(
        coll,
        name=f'{APP_NAME}.app',
        icon='icon.icns' if os.path.exists('icon.icns') else None,
        bundle_identifier='com.zionp.pptmerger',
        info_plist={
            'NSHighResolutionCapable': True,
            'CFBundleShortVersionString': '1.0.0',
            'CFBundleVersion': '1.0.0',
            'CFBundleDisplayName': APP_NAME,
            'LSMinimumSystemVersion': '11.0',
            'NSRequiresAquaSystemAppearance': False,  # 다크모드 지원
        },
    )

else:
    # ── Windows ────────────────────────────────────────────
    exe = EXE(
        pyz, a.scripts, a.binaries, a.zipfiles, a.datas, [],
        name=APP_NAME,
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=True,
        upx_exclude=[],
        runtime_tmpdir=None,
        console=False,
        disable_windowed_traceback=False,
        target_arch=None,
        codesign_identity=None,
        entitlements_file=None,
        icon='icon.ico' if os.path.exists('icon.ico') else None,
        version_file=None,
    )
