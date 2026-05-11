# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['app_gui.py'],
    pathex=[],
    binaries=[],
    datas=[('C:\\Users\\PC\\AppData\\Local\\Programs\\Python\\Python312\\Lib\\site-packages\\customtkinter', 'customtkinter/'), ('C:\\Users\\PC\\AppData\\Local\\Programs\\Python\\Python312\\Lib\\site-packages\\playwright', 'playwright/'), ('D:\\CODE\\autoSendZalo\\pw-browsers', 'pw-browsers/'), ('D:\\CODE\\autoSendZalo\\poppler', 'poppler/'), ('D:\\CODE\\autoSendZalo\\mailauto', 'mailauto/'), ('D:\\CODE\\autoSendZalo\\Logo512.png', '.'), ('D:\\CODE\\autoSendZalo\\Logo512.ico', '.')],
    hiddenimports=['win32com.client', 'pythoncom', 'email_core', 'zalo_core'],
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
    icon=['D:\\CODE\\autoSendZalo\\Logo512.ico'],
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
