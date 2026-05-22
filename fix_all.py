"""
fix_all.py — corrige todos os problemas de uma vez:
  1. base.html        → adiciona htmx, panel.js, hx-boost, loading-overlay
  2. Chaves templates → adiciona loading-overlay HTML
  3. Chart.register   → envolve em try/catch (evita erro "already registered")
  4. overview.html    → corrige charts se houver
"""
import re, os

BASE = "templates"

# ─── LOADING OVERLAY HTML ────────────────────────────────────────────────────
OVERLAY_HTML = '<div id="loading-overlay"><div class="loading-box"><div class="spinner"></div><p>Processando...</p></div></div>'

OVERLAY_CSS = """
    #loading-overlay{ display:none; position:fixed; inset:0; z-index:99999;
      background:rgba(15,23,42,.52); backdrop-filter:blur(3px);
      align-items:center; justify-content:center; }
    #loading-overlay.show{ display:flex; }
    .loading-box{ background:#fff; border-radius:20px; padding:36px 52px;
      text-align:center; box-shadow:0 24px 60px rgba(0,0,0,.22); }
    .spinner{ width:46px; height:46px; border:4px solid #e5e7eb;
      border-top-color:#d40511; border-radius:50%;
      animation:spin .75s linear infinite; margin:0 auto 18px; }
    @keyframes spin{ to{ transform:rotate(360deg); } }
    .loading-box p{ margin:0; font-size:14px; font-weight:800;
      color:#374151; letter-spacing:.3px; }"""


# ══════════════════════════════════════════════════════════════════════════════
# 1. base.html — adiciona htmx + panel.js + hx-boost + overlay
# ══════════════════════════════════════════════════════════════════════════════
base_path = os.path.join(BASE, "base.html")
with open(base_path, encoding="utf-8") as f:
    raw = f.read()

changed = False

if "htmx.min.js" not in raw:
    raw = raw.replace("</head>",
        '    <script src="/static/htmx.min.js"></script>\n'
        '    <script src="/static/panel.js"></script>\n'
        "</head>")
    changed = True

if "hx-boost" not in raw:
    raw = re.sub(r'<body([^>]*)>', r'<body\1 hx-boost="true">', raw, count=1)
    changed = True

if "loading-overlay" not in raw:
    # Adiciona CSS no <style> existente e HTML após <body>
    raw = raw.replace("</style>", OVERLAY_CSS + "\n    </style>", 1)
    raw = re.sub(r'(<body[^>]*>)', r'\1\n' + OVERLAY_HTML, raw, count=1)
    changed = True

if changed:
    with open(base_path, "w", encoding="utf-8") as f:
        f.write(raw)
    print("  OK: base.html")
else:
    print("  skip: base.html (já atualizado)")


# ══════════════════════════════════════════════════════════════════════════════
# 2. Chaves templates — adiciona loading-overlay
# ══════════════════════════════════════════════════════════════════════════════
chaves_files = [
    os.path.join(BASE, "chaves", "meu_claviculario.html"),
    os.path.join(BASE, "chaves", "realizar_devolucao.html"),
    os.path.join(BASE, "chaves", "realizar_retirada.html"),
]

for path in chaves_files:
    with open(path, encoding="utf-8") as f:
        raw = f.read()

    if "loading-overlay" in raw:
        print(f"  skip: {os.path.basename(path)} (já tem overlay)")
        continue

    # Adiciona CSS do overlay no primeiro </style>
    raw = raw.replace("</style>", OVERLAY_CSS + "\n    </style>", 1)

    # Adiciona HTML após <body ...>
    raw = re.sub(r'(<body[^>]*>)', r'\1\n' + OVERLAY_HTML, raw, count=1)

    with open(path, "w", encoding="utf-8") as f:
        f.write(raw)
    print(f"  OK: {os.path.basename(path)}")


# ══════════════════════════════════════════════════════════════════════════════
# 3. Chart.register — envolve em try/catch nos 3 dashboards + overview
# ══════════════════════════════════════════════════════════════════════════════
chart_files = [
    os.path.join(BASE, "dashboard.html"),
    os.path.join(BASE, "dashboard_analise.html"),
    os.path.join(BASE, "dashboard_anc.html"),
    os.path.join(BASE, "overview.html"),
]

OLD_REG = "Chart.register(ChartDataLabels);"
NEW_REG = "try { Chart.register(ChartDataLabels); } catch(e) {}"

for path in chart_files:
    if not os.path.exists(path):
        continue
    with open(path, encoding="utf-8") as f:
        raw = f.read()

    if OLD_REG not in raw:
        print(f"  skip: {os.path.basename(path)} (Chart.register não encontrado ou já corrigido)")
        continue

    raw = raw.replace(OLD_REG, NEW_REG)
    with open(path, "w", encoding="utf-8") as f:
        f.write(raw)
    print(f"  OK: {os.path.basename(path)} (Chart.register try/catch)")


# ══════════════════════════════════════════════════════════════════════════════
# 4. Garante que overview.html tem chart.umd.min.js no head (se usar Chart.js)
# ══════════════════════════════════════════════════════════════════════════════
overview_path = os.path.join(BASE, "overview.html")
with open(overview_path, encoding="utf-8") as f:
    raw = f.read()

changed = False
if "new Chart(" in raw and "chart.umd.min.js" not in raw:
    raw = raw.replace("</head>",
        '  <script src="/static/chart.umd.min.js"></script>\n'
        '  <script src="/static/chartjs-plugin-datalabels.min.js"></script>\n'
        "</head>")
    changed = True

# Envolve chart init em RAF se ainda não tiver
if "new Chart(" in raw and "requestAnimationFrame" not in raw:
    pattern = re.compile(r'(<script>\s*)((?:try\s*\{[^}]*\}[^}]*|Chart\.[^\n]+\n|const\s+\w+\s*=[\s\S]*?(?=new Chart))*new Chart[\s\S]*?(?=</script>))', re.DOTALL)
    def wrap_raf(m):
        inner = m.group(2).strip()
        return m.group(1) + "requestAnimationFrame(function(){\n" + inner + "\n});\n"
    new_raw, n = pattern.subn(wrap_raf, raw, count=1)
    if n:
        raw = new_raw
        changed = True

if changed:
    with open(overview_path, "w", encoding="utf-8") as f:
        f.write(raw)
    print(f"  OK: overview.html")

print("\nConcluído.")
