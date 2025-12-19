# build.spec
block_cipher = None

a = Analysis(
    ['win11_optimizer.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('optimizer.ico', '.')
    ],
    hiddenimports=[
        'tkinter',
        'subprocess', 
        'os',
        'datetime',
        'threading',
        'sys',
        'ctypes',
        'winreg',
        'psutil'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='Windows 11 Optimizer v4.7',
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
    icon=['optimizer.ico'],
    version='version.rc'
)