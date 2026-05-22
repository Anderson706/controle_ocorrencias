/* panel.js — comportamentos globais do painel (htmx-aware)
   Carregado em todas as páginas via <head> ou base.html.
   ─────────────────────────────────────────────────────────
   Responsabilidades:
     1. Loading overlay padronizado (estilo tela de login)
     2. Intercepta TODOS os cliques em links e submits de form
     3. Integração com htmx
     4. Expõe window.mostrarLoading() globalmente
*/
(function () {
  'use strict';

  /* ── URLs que NÃO devem mostrar loading (downloads) ─────────────────────── */
  var SKIP = [
    'exportar','download','anexo','comprovante',
    'export','excel','pdf','termo',
  ];

  function isSkip(str) {
    if (!str) return false;
    var s = str.toLowerCase();
    /* ignora âncoras e javascript */
    if (s === '#' || s.startsWith('javascript:')) return true;
    return SKIP.some(function(k){ return s.indexOf(k) !== -1; });
  }

  /* ── CSS padrão (estilo login: fundo escuro + spinner amarelo) ───────────── */
  var OVERLAY_CSS = [
    '#loading-overlay{',
    '  position:fixed!important;inset:0!important;',
    '  background:rgba(15,23,42,.65)!important;',
    '  z-index:2147483647!important;',   /* máximo z-index */
    '  display:none;',
    '  align-items:center!important;justify-content:center!important;',
    '  flex-direction:column!important;',
    '  backdrop-filter:blur(4px)!important;',
    '}',
    '#loading-overlay.show{display:flex!important;}',

    /* spinner amarelo */
    '#loading-overlay .ov-spinner{',
    '  width:52px;height:52px;',
    '  border:5px solid rgba(255,255,255,.20);',
    '  border-top-color:#ffcc00;',
    '  border-radius:50%;',
    '  animation:_panel_spin .75s linear infinite;',
    '}',
    '@keyframes _panel_spin{to{transform:rotate(360deg);}}',

    /* texto branco */
    '#loading-overlay .ov-text{',
    '  margin:20px 0 0;font-size:14px;font-weight:800;',
    '  color:#fff;letter-spacing:.3px;font-family:Arial,sans-serif;',
    '}',

    /* anula estilos legado (.loading-box com fundo branco / spinner vermelho) */
    '#loading-overlay .loading-box{',
    '  background:transparent!important;box-shadow:none!important;',
    '  padding:0!important;border-radius:0!important;',
    '  display:flex!important;flex-direction:column!important;',
    '  align-items:center!important;',
    '}',
    '#loading-overlay .loading-box .spinner,',
    '#loading-overlay>.spinner{',
    '  display:none!important;',
    '}',
    '#loading-overlay .loading-box p,',
    '#loading-overlay>p{',
    '  display:none!important;',
    '}',
  ].join('');

  /* HTML injetado dentro do overlay */
  var OV_INNER = '<div class="ov-spinner"></div><p class="ov-text">Processando, aguarde...</p>';

  /* ── Helpers ──────────────────────────────────────────────────────────────── */
  function getOv() {
    return document.getElementById('loading-overlay');
  }

  function show() {
    var ov = getOv();
    if (ov) ov.classList.add('show');
  }

  function hide() {
    var ov = getOv();
    if (ov) ov.classList.remove('show');
  }

  /* Expõe globalmente para uso inline nos templates */
  window.mostrarLoading = show;
  window.esconderLoading = hide;

  /* ── Inicialização ───────────────────────────────────────────────────────── */
  function applyStyles() {
    /* Remove versão anterior do CSS injetado */
    var old = document.getElementById('_panel_ov_css');
    if (old) old.parentNode.removeChild(old);

    var s = document.createElement('style');
    s.id  = '_panel_ov_css';
    s.textContent = OVERLAY_CSS;
    document.head.appendChild(s);
  }

  function ensureOverlay() {
    var ov = getOv();
    if (!ov) {
      ov    = document.createElement('div');
      ov.id = 'loading-overlay';
      document.body.insertBefore(ov, document.body.firstChild);
    }
    /* Garante que o HTML interno padrão existe */
    if (!ov.querySelector('.ov-spinner')) {
      ov.insertAdjacentHTML('beforeend', OV_INNER);
    }
  }

  function init() {
    applyStyles();
    ensureOverlay();
  }

  /* ── Eventos htmx ───────────────────────────────────────────────────────── */
  /* Mostra loading para QUALQUER requisição htmx (boost ou parcial) */
  document.addEventListener('htmx:beforeRequest', show);

  /* Esconde quando o conteúdo assentar */
  document.addEventListener('htmx:afterSettle', function () {
    hide();
    /* Reaplica estilos e garante overlay após swap de página */
    applyStyles();
    ensureOverlay();
  });

  /* Esconde em erros */
  document.addEventListener('htmx:responseError', hide);
  document.addEventListener('htmx:sendError',     hide);

  /* ── Cliques em links ───────────────────────────────────────────────────── */
  document.addEventListener('click', function (e) {
    /* Sobe até o <a> mais próximo */
    var el = e.target;
    while (el && el.tagName !== 'A') el = el.parentElement;
    if (!el) return;

    var href = el.getAttribute('href') || '';
    if (isSkip(href)) return;
    if (href === '' || href === '#') return;

    /* Ignora target="_blank" */
    if (el.target === '_blank') return;

    show();
  }, true); /* capture=true — antes do htmx processar */

  /* ── Submits de formulário ──────────────────────────────────────────────── */
  document.addEventListener('submit', function (e) {
    var form   = e.target;
    var action = form.getAttribute('action') || '';
    if (isSkip(action)) return;
    show();
  }, true);

  /* ── Segurança: timeout de 15 s esconde caso a página trave ────────────── */
  var _safeTimer = null;
  document.addEventListener('htmx:beforeRequest', function () {
    clearTimeout(_safeTimer);
    _safeTimer = setTimeout(hide, 15000);
  });
  document.addEventListener('click', function () {
    clearTimeout(_safeTimer);
    _safeTimer = setTimeout(hide, 15000);
  }, true);

  /* ── Bootstrap ──────────────────────────────────────────────────────────── */
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }

})();
