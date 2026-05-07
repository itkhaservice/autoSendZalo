# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['app_gui.py'],
    pathex=[],
    binaries=[],
    datas=[('C:\\Users\\qlnha\\AppData\\Roaming\\Python\\Python314\\site-packages\\customtkinter', 'customtkinter/'), ('C:\\Users\\qlnha\\AppData\\Roaming\\Python\\Python314\\site-packages\\playwright', 'playwright/'), ('ms-playwright', 'ms-playwright/'), ('poppler', 'poppler/'), ('Logo512.png', '.'), ('Logo512.ico', '.')],
    hiddenimports=['win32com.client', 'pythoncom'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'notebook', 'scipy', 'torch', 'torchvision'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='ZaloAutoSender',
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
    icon=['Logo512.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='ZaloAutoSender',
)
