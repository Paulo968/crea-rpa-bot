# -- mode: python ; coding: utf-8 --

from PyInstaller.utils.hooks import collect_submodules
from PyInstaller.building.build_main import Analysis, PYZ, EXE, COLLECT

hidden_imports = collect_submodules("customtkinter")
block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=['C:\\Users\\Faturamento-Passos\\Downloads\\crea_bot_REAL_COMPLETO (1)\\crea_bot_REAL_COMPLETO'],
    binaries=[],
    datas=[
        ('C:\\Users\\Faturamento-Passos\\Downloads\\crea_bot_REAL_COMPLETO (1)\\crea_bot_REAL_COMPLETO\\icon.ico', '.'),
        # Adicione essa linha abaixo (ajuste o caminho se necessário):
        ('C:\\Users\\Faturamento-Passos\\Downloads\\crea_bot_REAL_COMPLETO (1)\\crea_bot_REAL_COMPLETO\\venv\\Lib\\site-packages\\validate_docbr', 'validate_docbr'),
    ],
    hiddenimports=hidden_imports,
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
)

pyz = PYZ(
    a.pure,
    a.zipped_data,
    cipher=block_cipher
)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='CREA-BOT',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon='C:\\Users\\Faturamento-Passos\\Downloads\\crea_bot_REAL_COMPLETO (1)\\crea_bot_REAL_COMPLETO\\icon.ico'
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='CREA-BOT'
)
