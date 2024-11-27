# -*- mode: python ; coding: utf-8 -*-

# Adjust Analysis to add paths if needed
a = Analysis(
    ['excel_list_cleaner.py'],
    pathex=['.'],  # Specify the path if additional imports are in subdirectories
    binaries=[], 
    datas=[('scribe-logo-final.png', '.'), ('scribe-icon.ico', '.')],  # data files like logo and icon
    hiddenimports=[],  # include dependencies explicitly
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

# Standard PYZ creation
pyz = PYZ(a.pure)

# Define EXE options, including icon path, name, etc.
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='excel_column_cleaner',  # Executable name
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Change to True if a console is needed for debugging
    onefile=False,  # True for single-file output; False to retain individual files in output folder
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='scribe-icon.ico',  # icon file for the executable
)
