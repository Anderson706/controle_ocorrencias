# updater.spec — PyInstaller spec para CCTV_Updater
# Compilar: $venv_py -m PyInstaller updater.spec --noconfirm
from PyInstaller.utils.hooks import collect_all, collect_data_files, collect_submodules

block_cipher = None

# ── Dados ─────────────────────────────────────────────────────────────────────
datas = []
datas += collect_data_files('oracledb')

# ── Hidden imports ─────────────────────────────────────────────────────────────
hiddenimports = (
    collect_submodules('oracledb')
    + collect_submodules('cryptography')
    + [
        'cryptography.hazmat.primitives.kdf.pbkdf2',
        'cryptography.hazmat.primitives.kdf.scrypt',
        'cryptography.hazmat.backends.openssl',
        'tkinter',
        'tkinter.ttk',
        'tkinter.messagebox',
        # stdlib que o oracledb importa e o PyInstaller pode perder
        'getpass',
        'readline',
        'pwd',
        'grp',
        'termios',
        'tty',
    ]
)

# ── Analysis ──────────────────────────────────────────────────────────────────
a = Analysis(
    ['updater.py'],
    pathex=['.'],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'pytest', 'unittest', 'doctest', 'pdb',
        'flask', 'flask_sqlalchemy', 'sqlalchemy',
        'webview', 'pythonnet', 'clr_loader',
        'reportlab', 'docx', 'openpyxl',
        'PIL', 'pillow',
    ],
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# ── EXE --onefile ─────────────────────────────────────────────────────────────
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='CCTV_Updater',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,        # GUI — usa tkinter (sem janela de terminal)
    icon='static/icone.ico',
)
