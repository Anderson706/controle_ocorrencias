/* panel.js — indicadores de carregamento globais (htmx-aware)
   ─────────────────────────────────────────────────────────────
   Navegação (GET)  → barra fina no topo (estilo GitHub/Linear)
   Submit (POST)    → overlay escuro com spinner
   Expõe: window.mostrarLoading(), window.esconderLoading()
*/
(function () {
  'use strict';

  /* ═══════════════════════════════════════════════════════════
     1. TOP PROGRESS BAR — para navegações GET
  ═══════════════════════════════════════════════════════════ */
  var _bar       = null;
  var _barTimer  = null;
  var _barActive = false;

  function _createBar() {
    var b = document.createElement('div');
    b.id = '_top_progress';
    b.style.cssText = [
      'position:fixed', 'top:0', 'left:0',
      'height:3px', 'width:0',
      'z-index:2147483646',
      'pointer-events:none',
      'opacity:0',
      'background:linear-gradient(90deg,#d40511,#ff6b35)',
      'box-shadow:0 0 10px 1px rgba(212,5,17,.6),0 0 4px 0 #ffcc00',
      'transition:none',
    ].join(';');
    document.body.insertBefore(b, document.body.firstChild);
    return b;
  }

  function _getBar() {
    if (!_bar || !document.body.contains(_bar)) {
      _bar = document.getElementById('_top_progress') || _createBar();
    }
    return _bar;
  }

  function showBar() {
    if (_barActive) return;
    _barActive = true;
    clearTimeout(_barTimer);
    var b = _getBar();
    /* Reset sem transição */
    b.style.transition = 'none';
    b.style.width = '0';
    b.style.opacity = '1';
    /* Aguarda o browser aplicar o reset, depois anima até 80% */
    requestAnimationFrame(function () {
      requestAnimationFrame(function () {
        b.style.transition = 'width 12s cubic-bezier(.05,.7,.1,1)';
        b.style.width = '80%';
      });
    });
  }

  function hideBar() {
    if (!_barActive) return;
    _barActive = false;
    clearTimeout(_barTimer);
    var b = _getBar();
    /* Completa rapidamente até 100% */
    b.style.transition = 'width .18s ease,opacity .35s ease .18s';
    b.style.width = '100%';
    /* Fade out e reset */
    _barTimer = setTimeout(function () {
      b.style.opacity = '0';
      _barTimer = setTimeout(function () {
        b.style.transition = 'none';
        b.style.width = '0';
      }, 380);
    }, 180);
  }


  /* ═══════════════════════════════════════════════════════════
     2. OVERLAY — para submits POST (salvar / excluir)
  ═══════════════════════════════════════════════════════════ */
  var _OV_CSS = [
    '#loading-overlay{',
    '  position:fixed!important;inset:0!important;',
    '  background:rgba(15,23,42,.68)!important;',
    '  z-index:2147483647!important;',
    '  display:none;',
    '  align-items:center!important;justify-content:center!important;',
    '  flex-direction:column!important;',
    '  backdrop-filter:blur(3px)!important;',
    '}',
    '#loading-overlay.show{display:flex!important;}',
    '#loading-overlay .ov-spinner{',
    '  width:44px;height:44px;',
    '  border:4px solid rgba(255,255,255,.18);',
    '  border-top-color:#ffcc00;',
    '  border-radius:50%;',
    '  animation:_pnl_spin .65s linear infinite;',
    '}',
    '@keyframes _pnl_spin{to{transform:rotate(360deg)}}',
    '#loading-overlay .ov-text{',
    '  margin:16px 0 0;font-size:13px;font-weight:800;',
    '  color:#fff;letter-spacing:.3px;font-family:Arial,sans-serif;',
    '}',
  ].join('');

  var _OV_HTML = '<div class="ov-spinner"></div><p class="ov-text">Processando, aguarde...</p>';

  function _applyOvStyles() {
    var old = document.getElementById('_pnl_ov_css');
    if (old) old.parentNode.removeChild(old);
    var s = document.createElement('style');
    s.id = '_pnl_ov_css';
    s.textContent = _OV_CSS;
    document.head.appendChild(s);
  }

  function _ensureOverlay() {
    var ov = document.getElementById('loading-overlay');
    if (!ov) {
      ov = document.createElement('div');
      ov.id = 'loading-overlay';
      document.body.insertBefore(ov, document.body.firstChild);
    }
    if (!ov.querySelector('.ov-spinner')) {
      ov.insertAdjacentHTML('beforeend', _OV_HTML);
    }
    return ov;
  }

  function showOv() {
    _ensureOverlay().classList.add('show');
  }

  function hideOv() {
    var ov = document.getElementById('loading-overlay');
    if (ov) ov.classList.remove('show');
  }

  /* API pública — usada em templates (onclick="mostrarLoading()") */
  window.mostrarLoading  = showOv;
  window.esconderLoading = hideOv;


  /* ═══════════════════════════════════════════════════════════
     3. SKIP — URLs que não devem disparar indicadores
  ═══════════════════════════════════════════════════════════ */
  var _SKIP = ['exportar','download','anexo','comprovante','export','excel','pdf','termo'];

  function _isSkip(str) {
    if (!str) return false;
    var s = str.toLowerCase();
    if (s === '#' || s.indexOf('javascript:') === 0) return true;
    return _SKIP.some(function (k) { return s.indexOf(k) !== -1; });
  }


  /* ═══════════════════════════════════════════════════════════
     4. HTMX — integração com os eventos
  ═══════════════════════════════════════════════════════════ */
  document.addEventListener('htmx:beforeRequest', function (e) {
    var verb = (e.detail && e.detail.requestConfig && e.detail.requestConfig.verb) || 'get';
    if (verb === 'get') {
      showBar();
    } else {
      /* POST/PUT/DELETE → overlay */
      showOv();
    }
  });

  document.addEventListener('htmx:afterSettle', function () {
    hideBar();
    hideOv();
    /* Reaplica estilos e garante elementos após cada swap */
    _applyOvStyles();
    _ensureOverlay();
    _getBar();
  });

  document.addEventListener('htmx:responseError', function () { hideBar(); hideOv(); });
  document.addEventListener('htmx:sendError',     function () { hideBar(); hideOv(); });


  /* ═══════════════════════════════════════════════════════════
     5. CLIQUES EM LINKS — barra de progresso
  ═══════════════════════════════════════════════════════════ */
  document.addEventListener('click', function (e) {
    var el = e.target;
    while (el && el.tagName !== 'A') el = el.parentElement;
    if (!el) return;
    var href = el.getAttribute('href') || '';
    if (_isSkip(href) || href === '' || href === '#') return;
    if (el.target === '_blank') return;
    showBar();
  }, true);


  /* ═══════════════════════════════════════════════════════════
     6. SUBMITS DE FORMULÁRIO — overlay
  ═══════════════════════════════════════════════════════════ */
  document.addEventListener('submit', function (e) {
    var action = (e.target && e.target.getAttribute('action')) || '';
    if (_isSkip(action)) return;
    showOv();
  }, true);


  /* ═══════════════════════════════════════════════════════════
     7. SEGURANÇA — esconde tudo após 15 s (evita tela travada)
  ═══════════════════════════════════════════════════════════ */
  var _safeTimer = null;

  function _resetSafe() {
    clearTimeout(_safeTimer);
    _safeTimer = setTimeout(function () { hideBar(); hideOv(); }, 15000);
  }

  document.addEventListener('htmx:beforeRequest', _resetSafe);
  document.addEventListener('click', _resetSafe, true);
  document.addEventListener('submit', _resetSafe, true);


  /* ═══════════════════════════════════════════════════════════
     8. INICIALIZAÇÃO
  ═══════════════════════════════════════════════════════════ */
  function _init() {
    _applyOvStyles();
    _ensureOverlay();
    _getBar();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', _init);
  } else {
    _init();
  }

})();


/* ═══════════════════════════════════════════════════════════
   AUTO-FILTER — submete formulários de filtro GET automaticamente
   Selects: submit imediato ao mudar
   Inputs texto/data: submit com debounce de 400ms
   Ativa em qualquer <form data-autofilter> ou <form class="auto-filter">
═══════════════════════════════════════════════════════════ */
(function () {
  'use strict';

  var _timers = new WeakMap();

  function _doSubmit(form) {
    if (form._afSubmitting) return;
    form._afSubmitting = true;
    if (window.mostrarLoading) window.mostrarLoading();
    form.submit();
  }

  function _debounced(form, delay) {
    var t = _timers.get(form);
    if (t) clearTimeout(t);
    _timers.set(form, setTimeout(function () { _doSubmit(form); }, delay));
  }

  function _isFilterForm(form) {
    if ((form.getAttribute('method') || 'get').toUpperCase() !== 'GET') return false;
    return 'autofilter' in form.dataset || form.classList.contains('auto-filter');
  }

  function _isSubmitBtn(el) {
    if (el.tagName === 'INPUT' && el.type === 'submit') return true;
    if (el.tagName === 'BUTTON') {
      var t = (el.getAttribute('type') || 'submit').toLowerCase();
      return t === 'submit';
    }
    return false;
  }

  function _setup() {
    document.querySelectorAll('form').forEach(function (form) {
      if (!_isFilterForm(form)) return;

      // Selects → submit imediato
      form.querySelectorAll('select').forEach(function (el) {
        el.addEventListener('change', function () { _doSubmit(form); });
      });

      // Textos e datas → submit com debounce
      form.querySelectorAll('input').forEach(function (el) {
        var t = (el.getAttribute('type') || 'text').toLowerCase();
        if (['text','search','date','month','number'].indexOf(t) !== -1) {
          el.addEventListener('input', function () { _debounced(form, 400); });
        }
      });

      // Oculta botões de submit
      form.querySelectorAll('button, input[type=submit]').forEach(function (el) {
        if (_isSubmitBtn(el)) el.style.display = 'none';
      });
    });
  }

  // panel.js é carregado no fim do <body> — DOM já pronto na maioria dos casos
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', _setup);
  } else {
    _setup();
  }
})();
