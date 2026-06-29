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
datas += collect_data_files('pptx')          # template .pptx interno do python-pptx
datas += collect_data_files('oracledb')
datas += collect_data_files('clr_loader')
datas += collect_data_files('pythonnet')
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
    + collect_submodules('sqlalchemy.dialects.sqlite')
    + collect_submodules('flask_sqlalchemy')
    + collect_submodules('reportlab')
    + collect_submodules('docx')
    + collect_submodules('openpyxl')
    + collect_submodules('pptx')
    + collect_submodules('webview')
    + collect_submodules('clr_loader')
    + collect_submodules('pythonnet')
    + [
        'webview.platforms.winforms',
        'clr',
        'clr_loader',
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
        # Controle de Chaves
        'chaves_blueprint',
        # Achados e Perdidos
        'achados_blueprint',
        # Controle de Ativos (redundância via Supabase)
        'ativos_blueprint',
    ]
)

# ── Binários do pythonnet (CLR) ───────────────────────────────────────────────
binaries = []

# ── Stack do Supabase (Controle de Ativos / Usuários Ativos via REST) ─────────
# supabase 2.31 usa sub-pacotes próprios (postgrest, supabase_auth, etc.) com
# imports dinâmicos — collect_all garante módulos + dados + binários de cada um.
hiddenimports = list(hiddenimports)
for _pkg in ('supabase', 'postgrest', 'supabase_auth', 'supabase_functions',
             'realtime', 'storage3', 'httpx', 'httpcore', 'h2', 'hpack',
             'websockets', 'gotrue', 'supafunc',
             'truststore', 'certifi'):   # truststore = confia na CA do proxy DHL
    try:
        _d, _b, _h = collect_all(_pkg)
        datas += _d
        binaries += _b
        hiddenimports += _h
    except Exception:
        pass

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

# ── Splash screen (aparece enquanto o EXE extrai e inicializa) ────────────────
splash = Splash(
    'static/splash.png',
    binaries=a.binaries,
    datas=a.datas,
    text_pos=None,
    text_color='black',
    text_size=12,
    minify_script=True,
    always_on_top=True,
)

# ── EXE --onedir (one-folder) ─────────────────────────────────────────────────
# Modo pasta: o EXE NÃO re-extrai ~63 MB no temp a cada abertura (inicialização
# muito mais rápida). 'splash' fica no EXE; binários/dados vão para o COLLECT.
exe = EXE(
    pyz,
    a.scripts,
    splash,
    [],
    exclude_binaries=True,
    name='CCTV_ControlPanel',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    icon='static/icone.ico',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    splash.binaries,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='CCTV_ControlPanel',
)
