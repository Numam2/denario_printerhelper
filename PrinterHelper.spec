# -*- mode: python ; coding: utf-8 -*-
# Build:  .\build.ps1   (o: pyinstaller --noconfirm --clean PrinterHelper.spec)


a = Analysis(
    ['PrinterHelper.py'],
    pathex=[],
    binaries=[],
    datas=[('icon.ico', '.')],
    hiddenimports=['win32com.client'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # Paquetes del entorno global que se colaban por imports opcionales y no se usan.
    excludes=[
        'flask_cors', 'tkinter', 'numpy', 'scipy', 'pandas', 'matplotlib',
        'IPython', 'cryptography', 'win32ui', 'pywin', 'pythonwin',
    ],
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
    name='DenarioPrinter',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,  # UPX aumenta los falsos positivos de antivirus
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['icon.ico'],
    version='version_info.txt',
)
