# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main.py'],  # Ersetze mit dem Namen deines Python-Skripts
    pathex=['.'],  # Verzeichnis, in dem sich dein Skript befindet
    binaries=[],
    datas=[
        ('app-icon.ico', '.'),  # Füge hier dein Symbol hinzu
    ],
    hiddenimports=[],  # Füge hier versteckte Imports hinzu, falls erforderlich
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],  # Module ausschließen, die nicht benötigt werden
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Rechnungspruefung v1.0',  # Ersetze mit dem gewünschten Namen deiner .exe
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,  # UPX-Komprimierung aktivieren
    console=False,  # Verhindert das Öffnen einer Konsole (bei GUI-Anwendungen)
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Rechnungspruefung v1.0',  # Ersetze mit dem gewünschten Namen deines Builds
)
