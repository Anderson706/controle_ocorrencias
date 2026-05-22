"""Move blocos <style> do <head> para o <body> em todos os templates do painel."""
import os, re, glob

BASE = "templates"
SKIP = {"login.html","lgpd_aceite.html","esqueci_senha.html",
        "redefinir_senha.html","versao_bloqueada.html","base.html","_translate.html"}

files = (
    glob.glob(os.path.join(BASE, "*.html")) +
    glob.glob(os.path.join(BASE, "chaves", "*.html"))
)

modified = 0
for path in sorted(files):
    name = os.path.basename(path)
    if name in SKIP:
        continue

    with open(path, encoding="utf-8") as f:
        raw = f.read()

    head_end = raw.find("</head>")
    if head_end < 0:
        continue

    # 1. Coleta blocos <style>...</style> que estão DENTRO do <head>
    style_re = re.compile(r'<style>.*?</style>', re.DOTALL)
    styles_in_head = [m for m in style_re.finditer(raw) if m.start() < head_end]
    if not styles_in_head:
        continue

    # 2. Remove do head em ordem reversa
    new_raw = raw
    collected = []
    for m in reversed(styles_in_head):
        collected.insert(0, m.group())
        new_raw = new_raw[:m.start()] + new_raw[m.end():]

    style_block = "\n".join(collected)

    # 3. Scripts no <head> se necessário
    head_end2 = new_raw.find("</head>")
    scripts = ""
    if "htmx.min.js" not in new_raw:
        scripts += '  <script src="/static/htmx.min.js"></script>\n'
    if "panel.js" not in new_raw:
        scripts += '  <script src="/static/panel.js"></script>\n'
    if scripts:
        new_raw = new_raw[:head_end2] + scripts + new_raw[head_end2:]

    # 4. hx-boost no <body>
    if "hx-boost" not in new_raw:
        new_raw = re.sub(r'<body([^>]*)>', r'<body\1 hx-boost="true">', new_raw, count=1)

    # 5. Insere <style> logo após <body ...> usando str.replace simples
    body_match = re.search(r'<body[^>]*>', new_raw)
    if body_match:
        insert_pos = body_match.end()
        new_raw = new_raw[:insert_pos] + "\n" + style_block + new_raw[insert_pos:]

    with open(path, "w", encoding="utf-8") as f:
        f.write(new_raw)
    modified += 1
    print(f"  OK: {name}")

print(f"\nTotal: {modified} arquivos modificados")
