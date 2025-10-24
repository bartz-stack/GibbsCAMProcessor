# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['processor.py', 'config.py', 'logging_setup.py', 'notifications.py', 'ncf_parser.py', 'excel_mapper.py', 'window_detector.py', 'screenshot_gui.py', 'screenshot_capture.py', 'screenshot_colors.py'],
    pathex=['.'],
    binaries=[],
    datas=[('config.ini', '.'), ('Gibbscam.ico', '.')],
    hiddenimports=['win32com.client', 'win32gui', 'win32process', 'psutil', 'pandas', 'openpyxl', 'windows_toasts', 'PIL', 'PIL.Image', 'PIL.ImageTk', 'PIL.ImageGrab', 'tkinter', 'tkinter.ttk', 'tkinter.messagebox'],
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
    a.binaries,
    a.datas,
    [],
    name='GibbsCAMProcessor',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['Gibbscam.ico'],
)
