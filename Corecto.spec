# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['app_analyses_anomalies.py'],
    pathex=[],
    binaries=[],
    datas=[('logique_controles.py', '.'), ('main.py', '.'), ('regles_config.py', '.')],
    hiddenimports=['pandas', 'openpyxl', 'tkinterdnd2', 'xlrd', 'tkinter', 'tkinter.filedialog', 'tkinter.messagebox', 'threading', 'os', 'sys', 'logique_controles', 'main', 'regles_config'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'scipy', 'PIL', 'setuptools', 'unittest', 'test'],
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
    name='Corecto',
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
    icon='NONE',
)
