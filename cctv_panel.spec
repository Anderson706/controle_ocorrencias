# cctv_panel.spec  — PyInstaller spec para CCTV Control Panel
from PyInstaller.utils.hooks import collect_all, collect_data_files, collect_submodules

block_cipher = None

# ── Dados ────────────────────────────────────────────────────────────────────
datas = [
    ('templates', 'templates'),
    ('static',    'static'),
]
datas += collect_data_files('reportlab')
datas += collect_data_files('docx')
datas += collect_data_files('oracledb')
# webview: inclui apenas os arquivos que realmente existem
_wv = 'venv/Lib/site-packages/webview/lib'
datas += [
    (f'{_wv}/Microsoft.Web.WebView2.Core.dll',           'webview/lib'),
    (f'{_wv}/Microsoft.Web.WebView2.WinForms.dll',       'webview/lib'),
    (f'{_wv}/WebBrowserInterop.x64.dll',                 'webview/lib'),
    (f'{_wv}/WebBrowserInterop.x86.dll',                 'webview/lib'),
    (f'{_wv}/runtimes/win-x64/native/WebView2Loader.dll','webview/lib/runtimes/win-x64/native'),
    (f'{_wv}/runtimes/win-x86/native/WebView2Loader.dll','webview/lib/runtimes/win-x86/native'),
    (f'{_wv}/runtimes/win-arm64/native/WebView2Loader.dll','webview/lib/runtimes/win-arm64/native'),
]

# ── Hidden imports ────────────────────────────────────────────────────────────
hiddenimports = (
    collect_submodules('oracledb')
    + collect_submodules('cryptography')
    + collect_submodules('sqlalchemy')
    + collect_submodules('sqlalchemy.dialects.oracle')
    + collect_submodules('flask_sqlalchemy')
    + collect_submodules('reportlab')
    + collect_submodules('docx')
    + collect_submodules('openpyxl')
    + collect_submodules('webview')
    + [
        'webview.platforms.winforms',
        'clr',
        'pythonnet',
        'werkzeug',
        'werkzeug.security',
        'werkzeug.serving',
        'jinja2',
        'jinja2.ext',
        'click',
        'pkg_resources',
        'pkg_resources.extern',
        'cryptography.hazmat.primitives.kdf.pbkdf2',
        'cryptography.hazmat.primitives.kdf.scrypt',
        'cryptography.hazmat.backends.openssl',
    ]
)

# ── Binários do pythonnet (CLR) ───────────────────────────────────────────────
binaries = []

# ── Analysis ──────────────────────────────────────────────────────────────────
a = Analysis(
    ['launcher.py'],
    pathex=['.'],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['pytest', 'unittest', 'doctest', 'pdb'],
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# ── Splash nativo (aparece antes da extração) ─────────────────────────────────
splash = Splash(
    'static/splash.png',
    binaries=a.binaries,
    datas=a.datas,
    text_pos=(240, 200),
    text_size=10,
    text_color='#888888',
    minify_script=True,
    always_on_top=True,
)

# ── EXE --onefile ─────────────────────────────────────────────────────────────
exe = EXE(
    pyz,
    a.scripts,
    splash,
    splash.binaries,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='CCTV_ControlPanel',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    icon='static/icone.ico',
)
