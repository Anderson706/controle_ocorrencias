/* table-paginate.js — paginação client-side leve para tabelas grandes.
 *
 * Uso: marque a <table> com  data-paginate="50"  (linhas por página).
 * O script esconde as linhas além da página atual e injeta controles de
 * navegação logo abaixo da tabela. Reduz o custo de renderização no DOM
 * (importante em tablets) sem mudar nada no servidor.
 *
 * Reage a swaps do HTMX (hx-boost) reescaneando o conteúdo novo.
 */
(function () {
  function _btn(txt, fn) {
    var b = document.createElement('button');
    b.type = 'button';
    b.textContent = txt;
    b.style.cssText = 'border:1px solid #e5e7eb;background:#fff;border-radius:8px;padding:6px 14px;' +
                      'font-size:12px;font-weight:800;cursor:pointer;font-family:inherit;color:#1f2937;';
    b.addEventListener('click', fn);
    return b;
  }

  function paginar(table) {
    if (table._tpDone) return;
    var size = parseInt(table.getAttribute('data-paginate'), 10) || 50;
    var tbody = table.tBodies && table.tBodies[0];
    if (!tbody) return;
    var rows = Array.prototype.slice.call(tbody.rows);
    if (rows.length <= size) return;   // não precisa paginar

    table._tpDone = true;
    var page = 0;
    var pages = Math.ceil(rows.length / size);

    var pager = document.createElement('div');
    pager.className = 'tp-pager';
    pager.style.cssText = 'display:flex;align-items:center;justify-content:center;gap:12px;' +
                          'padding:14px 8px;flex-wrap:wrap;';

    function render() {
      var ini = page * size, fim = ini + size;
      for (var i = 0; i < rows.length; i++) {
        rows[i].style.display = (i >= ini && i < fim) ? '' : 'none';
      }
      pager.innerHTML = '';
      var prev = _btn('‹ Anterior', function () { if (page > 0) { page--; render(); _scrollTop(table); } });
      prev.disabled = (page === 0);
      prev.style.opacity = prev.disabled ? '.4' : '1';
      var info = document.createElement('span');
      info.style.cssText = 'font-size:12px;font-weight:700;color:#6b7280;white-space:nowrap;';
      info.textContent = 'Página ' + (page + 1) + ' de ' + pages + '  ·  ' + rows.length + ' registros';
      var next = _btn('Próxima ›', function () { if (page < pages - 1) { page++; render(); _scrollTop(table); } });
      next.disabled = (page === pages - 1);
      next.style.opacity = next.disabled ? '.4' : '1';
      pager.appendChild(prev);
      pager.appendChild(info);
      pager.appendChild(next);
    }

    function _scrollTop(el) {
      try { el.scrollIntoView({ block: 'nearest', behavior: 'smooth' }); } catch (e) {}
    }

    // Insere o pager logo após a tabela (ou após o wrapper de scroll, se houver)
    var alvo = table;
    if (table.parentNode && /scroll|table-wrap|table-card/i.test(table.parentNode.className || '')) {
      alvo = table.parentNode;
    }
    alvo.parentNode.insertBefore(pager, alvo.nextSibling);
    render();
  }

  function escanear(root) {
    var alvo = root || document;
    if (!alvo.querySelectorAll) return;
    Array.prototype.forEach.call(alvo.querySelectorAll('table[data-paginate]'), paginar);
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', function () { escanear(document); });
  } else {
    escanear(document);
  }
  document.addEventListener('htmx:afterSwap', function (e) { escanear(e.target || document); });
  document.addEventListener('htmx:load', function (e) { escanear(e.target || document); });

  window.TablePaginate = { escanear: escanear };
})();
