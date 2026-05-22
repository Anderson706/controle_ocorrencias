"""
fix_charts_pager.py
  1. Gráficos: envolve init em requestAnimationFrame (resolve timing htmx)
  2. Tabelas: injeta paginador de 10 linhas em todas as páginas de lista
"""
import re, os

BASE = "templates"

# ─── CSS DO PAGINADOR ─────────────────────────────────────────────────────────
PAGER_CSS = """
    /* ── Paginador de tabela ── */
    .tbl-pager{display:flex;align-items:center;justify-content:center;gap:10px;
               padding:14px 0 4px;flex-wrap:wrap;}
    .tbl-pager-btn{border:1px solid #e5e7eb;background:#fff;border-radius:8px;
                   padding:6px 16px;font-size:12px;font-weight:800;cursor:pointer;
                   color:#1f2937;transition:.15s;font-family:inherit;}
    .tbl-pager-btn:hover:not(:disabled){background:#d40511;color:#fff;border-color:#d40511;}
    .tbl-pager-btn:disabled{opacity:.35;cursor:default;}
    .tbl-pager-info{font-size:12px;font-weight:700;color:#6b7280;white-space:nowrap;}
"""

# ─── HTML DO PAGINADOR ────────────────────────────────────────────────────────
PAGER_HTML = """
<div class="tbl-pager">
  <button class="tbl-pager-btn" id="pager-prev" onclick="pagerPrev()">‹ Anterior</button>
  <span class="tbl-pager-info" id="pager-info">carregando...</span>
  <button class="tbl-pager-btn" id="pager-next" onclick="pagerNext()">Próxima ›</button>
</div>
"""

# ─── JS DO PAGINADOR (auto-contido, injeta após a tabela) ─────────────────────
PAGER_JS = """<script>
/* Paginador de tabela — 10 linhas por página */
(function(){
  var PER = 10, pg = 1, snap = null;
  var tb = document.querySelector('tbody');
  if(!tb) return;

  function dataRows(){
    return Array.from(tb.querySelectorAll('tr')).filter(function(tr){
      return !(tr.children.length === 1 && tr.children[0].colSpan > 1);
    });
  }

  function render(){
    if(snap === null){
      snap = dataRows().filter(function(tr){ return tr.style.display !== 'none'; });
    }
    var tot = snap.length;
    var totPg = Math.max(1, Math.ceil(tot / PER));
    if(pg > totPg) pg = totPg;
    var s = (pg - 1) * PER, e = s + PER;

    dataRows().forEach(function(tr){ tr.style.display = 'none'; });
    snap.forEach(function(tr, i){ if(i >= s && i < e) tr.style.display = ''; });

    var info = document.getElementById('pager-info');
    var prev = document.getElementById('pager-prev');
    var next = document.getElementById('pager-next');
    if(info) info.textContent = 'Pág. ' + pg + ' / ' + totPg + '  ·  ' + tot + ' registro' + (tot !== 1 ? 's' : '');
    if(prev) prev.disabled = pg <= 1;
    if(next) next.disabled = pg >= totPg;
  }

  window.pagerPrev  = function(){ if(pg > 1){ pg--; render(); } };
  window.pagerNext  = function(){ pg++; render(); };
  window.pagerReset = function(){ snap = null; pg = 1; render(); };

  render();
})();
</script>"""

# ─── ONDE INJETAR O PAGINADOR HTML (após </table> ou </tbody>) ───────────────
# Cada entrada: (arquivo, âncora_após_tabela)
LIST_TEMPLATES = [
    ("ocorrencias.html",        "</table>"),
    ("analises.html",           "</table>"),
    ("anc.html",                "</table>"),
    ("sh_registrar.html",       "</table>"),
    ("admin_usuarios.html",     "</table>"),
    ("admin_solicitacoes.html", "</table>"),
]

# ─── DASHBOARD: arquivos e trecho antes do primeiro new Chart ────────────────
DASH_TEMPLATES = [
    "dashboard.html",
    "dashboard_analise.html",
    "dashboard_anc.html",
]


def fix_charts(path):
    """Envolve toda a inicialização de Chart.js em requestAnimationFrame."""
    with open(path, encoding="utf-8") as f:
        raw = f.read()

    # Já corrigido?
    if "requestAnimationFrame" in raw:
        print(f"  skip (já tem RAF): {os.path.basename(path)}")
        return

    # Localiza o bloco <script> que contém 'new Chart('
    # Envolve tudo entre Chart.register(...) e o fechamento do script em RAF
    pattern = re.compile(
        r'(<script>\s*)(Chart\.register[\s\S]*?)(</script>)',
        re.DOTALL
    )

    def wrap_raf(m):
        inner = m.group(2).strip()
        return (
            m.group(1)
            + "requestAnimationFrame(function(){\n"
            + inner + "\n"
            + "});\n"
            + m.group(3)
        )

    new_raw, count = pattern.subn(wrap_raf, raw, count=1)
    if count:
        with open(path, "w", encoding="utf-8") as f:
            f.write(new_raw)
        print(f"  OK (charts RAF): {os.path.basename(path)}")
    else:
        print(f"  AVISO: padrão Chart não encontrado em {os.path.basename(path)}")


def add_pager(filename, anchor):
    path = os.path.join(BASE, filename)
    if not os.path.exists(path):
        # tenta chaves/
        path2 = os.path.join(BASE, "chaves", filename)
        if os.path.exists(path2):
            path = path2
        else:
            print(f"  skip (não encontrado): {filename}")
            return

    with open(path, encoding="utf-8") as f:
        raw = f.read()

    changed = False

    # 1. CSS — injeta no bloco <style> existente no body
    if "tbl-pager" not in raw:
        raw = raw.replace("</style>", PAGER_CSS + "\n    </style>", 1)
        changed = True

    # 2. HTML — injeta após a primeira ocorrência da âncora
    if "pager-prev" not in raw:
        # Insere após a PRIMEIRA ocorrência do anchor
        idx = raw.find(anchor)
        if idx >= 0:
            insert_at = idx + len(anchor)
            raw = raw[:insert_at] + "\n" + PAGER_HTML + raw[insert_at:]
            changed = True
        else:
            print(f"  AVISO: âncora '{anchor}' não encontrada em {filename}")

    # 3. JS — injeta antes de </body>
    if "pagerReset" not in raw:
        raw = raw.replace("</body>", PAGER_JS + "\n</body>")
        changed = True

    # 4. Para analises.html: patcha filtrar() para chamar pagerReset
    if filename == "analises.html" and "pagerReset" in raw:
        # Adiciona chamada ao pagerReset no final da função filtrar
        # O padrão: countEl.textContent = ... seguido de nada (fim do filtrar)
        old_count = "if(countEl) countEl.textContent = visible + ' de ' + total + ' registros';"
        new_count = old_count + "\n    if(window.pagerReset) window.pagerReset();"
        if old_count in raw and "if(window.pagerReset)" not in raw:
            raw = raw.replace(old_count, new_count)
            changed = True

    if changed:
        with open(path, "w", encoding="utf-8") as f:
            f.write(raw)
        print(f"  OK (pager): {filename}")
    else:
        print(f"  skip (sem mudança): {filename}")


# ─── EXECUÇÃO ─────────────────────────────────────────────────────────────────
print("=== Corrigindo gráficos (requestAnimationFrame) ===")
for name in DASH_TEMPLATES:
    fix_charts(os.path.join(BASE, name))

print("\n=== Adicionando paginador às tabelas ===")
for name, anchor in LIST_TEMPLATES:
    add_pager(name, anchor)

print("\nConcluído.")
