/* filters.js — Persistência de filtros de página via sessionStorage
   ────────────────────────────────────────────────────────────────────
   Comportamento:
     • Salva os params GET sempre que a página carrega com filtros na URL
     • Restaura os filtros quando a página carrega sem parâmetros
       (situação típica após POST → redirect de Flask)
     • Limpa APENAS quando o usuário clica num elemento com [data-clear-filter]

   Setup:
     • Inclua este script em base_app.html (já feito)
     • Adicione data-clear-filter ao botão "Limpar" de cada página filtrada
     • Qualquer <form method="get"> dentro de .filter-card é monitorado
       automaticamente
*/
(function () {
  'use strict';

  var _PREFIX = 'pf_';

  /* Chave única por rota — usa o pathname sem trailing slash */
  function _chave() {
    return _PREFIX + window.location.pathname.replace(/\/$/, '');
  }

  /* Retorna true se a querystring contém ao menos um parâmetro não-vazio */
  function _temFiltros(qs) {
    if (!qs || qs === '?' || qs.length <= 1) return false;
    try {
      var p = new URLSearchParams(qs);
      var tem = false;
      p.forEach(function (v) { if (v && v.trim()) tem = true; });
      return tem;
    } catch (e) {
      return qs.length > 1;
    }
  }

  function _salvar() {
    try { sessionStorage.setItem(_chave(), window.location.search); } catch (e) {}
  }

  function _limpar() {
    try { sessionStorage.removeItem(_chave()); } catch (e) {}
  }

  function _obterSalvo() {
    try { return sessionStorage.getItem(_chave()) || ''; } catch (e) { return ''; }
  }

  /* ── Core ───────────────────────────────────────────────────────── */
  function init() {
    /* Só atua em páginas que tenham um form de filtro GET dentro de .filter-card */
    if (!document.querySelector('.filter-card form[method="GET"]')) return;

    var qs = window.location.search;

    if (_temFiltros(qs)) {
      /* Página carregou com filtros na URL → persiste no storage */
      _salvar();
    } else {
      /* Página carregou limpa (ex: após POST redirect) → tenta restaurar */
      var salvo = _obterSalvo();
      if (_temFiltros(salvo)) {
        window.location.replace(window.location.pathname + salvo);
        return; /* a página vai recarregar — interrompe o init atual */
      }
    }

    /* Registra listeners nos botões "Limpar" desta página */
    document.querySelectorAll('[data-clear-filter]').forEach(function (el) {
      /* _pfBound evita empilhar listeners duplicados em re-inits HTMX */
      if (el._pfBound) return;
      el._pfBound = true;
      el.addEventListener('click', _limpar, true);
    });
  }

  /* ── Inicialização ──────────────────────────────────────────────── */
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }

  /* Re-executa após cada swap HTMX (novos elementos no DOM) */
  document.addEventListener('htmx:afterSettle', function () {
    /* Reseta _pfBound nos novos elementos (DOM reconstruído pelo HTMX) */
    document.querySelectorAll('[data-clear-filter]').forEach(function (el) {
      el._pfBound = false;
    });
    init();
  });

})();
