# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

hiddenimports = []
hiddenimports += collect_submodules('openpyxl')

# tzdata is pure data (no importable code to auto-detect), and report times
# are rendered in US Eastern via zoneinfo — bundle its IANA database files so
# the frozen app doesn't depend on the OS having one (Windows doesn't).
datas = collect_data_files('tzdata')

a = Analysis(
    ['src/webapp.py'],
    pathex=['src'],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='ShipReportBuilder',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='ShipReportBuilder',
)
app = BUNDLE(
    coll,
    name='ShipReportBuilder.app',
    icon='appicon.icns',
    bundle_identifier='com.blues.shipreportbuilder',
    info_plist={
        # Tk 8.5 renders a black window under macOS dark mode; force light Aqua.
        'NSRequiresAquaSystemAppearance': True,
        'NSHighResolutionCapable': True,
        'CFBundleName': 'ShipReportBuilder',
        'CFBundleDisplayName': 'Ship Report Builder',
    },
)
