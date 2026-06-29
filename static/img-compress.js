/* img-compress.js — compressão automática de imagens antes do upload.
 *
 * Anexa-se a TODOS os <input type="file"> que aceitam imagem e, ao selecionar
 * uma foto, redimensiona (máx. 2000px) e recomprime em JPEG (~0.82) no próprio
 * navegador. Uma foto de tablet/celular de dezenas de MB vira poucos MB — reduz
 * o armazenamento no Oracle, a transferência via VPN e o tempo de geração de PDF,
 * além de evitar o erro de tamanho de upload.
 *
 * Como desativar num input específico: adicione o atributo  data-no-compress
 * (ex.: o checklist de câmeras, que tem lógica própria).
 */
(function () {
  var MAX_DIM = 2000;     // maior dimensão (px) após redimensionar
  var QUALITY = 0.82;     // qualidade JPEG

  function _aceitaImagem(input) {
    if (input.hasAttribute('data-no-compress')) return false;
    var acc = (input.getAttribute('accept') || '').toLowerCase();
    if (acc === '') return true; // sem accept → tenta (só processa se o arquivo for imagem)
    return acc.indexOf('image/') !== -1 ||
           /\.(png|jpe?g|webp|gif|bmp|heic|heif)/.test(acc);
  }

  function comprimir(file) {
    return new Promise(function (resolve) {
      if (!file || !file.type || file.type.indexOf('image/') !== 0) { resolve(file); return; }
      // SVG e GIF animado não passam bem por canvas → mantém original
      if (file.type === 'image/svg+xml' || file.type === 'image/gif') { resolve(file); return; }

      var url = URL.createObjectURL(file);
      var img = new Image();
      img.onload = function () {
        URL.revokeObjectURL(url);
        var scale = Math.min(1, MAX_DIM / Math.max(img.width, img.height));
        var cw = Math.max(1, Math.round(img.width * scale));
        var ch = Math.max(1, Math.round(img.height * scale));
        var canvas = document.createElement('canvas');
        canvas.width = cw; canvas.height = ch;
        var ctx = canvas.getContext('2d');
        ctx.fillStyle = '#ffffff';                 // fundo branco (JPEG não tem alfa)
        ctx.fillRect(0, 0, cw, ch);
        ctx.drawImage(img, 0, 0, cw, ch);
        canvas.toBlob(function (blob) {
          if (!blob || blob.size >= file.size) { resolve(file); return; }  // só usa se ficou menor
          var nome = file.name.replace(/\.[^.]+$/, '') + '.jpg';
          resolve(new File([blob], nome, { type: 'image/jpeg', lastModified: Date.now() }));
        }, 'image/jpeg', QUALITY);
      };
      img.onerror = function () { URL.revokeObjectURL(url); resolve(file); };
      img.src = url;
    });
  }

  function anexar(input) {
    if (input._imgcAttached) return;
    input._imgcAttached = true;

    input.addEventListener('change', function () {
      if (input._imgcBusy) return;
      if (!input.files || !input.files.length) return;
      var files = Array.prototype.slice.call(input.files);
      var temImagem = files.some(function (f) { return f.type && f.type.indexOf('image/') === 0; });
      if (!temImagem) return;

      input._imgcBusy = true;
      Promise.all(files.map(comprimir)).then(function (out) {
        var mudou = out.some(function (f, i) { return f !== files[i]; });
        if (mudou && window.DataTransfer) {
          try {
            var dt = new DataTransfer();
            out.forEach(function (f) { dt.items.add(f); });
            input.files = dt.files;   // não dispara novo 'change'
          } catch (e) { /* navegador sem suporte → mantém original */ }
        }
        input._imgcBusy = false;
      }).catch(function () { input._imgcBusy = false; });
    }, false);
  }

  function escanear(root) {
    var alvo = root || document;
    var inputs = alvo.querySelectorAll ? alvo.querySelectorAll('input[type=file]') : [];
    Array.prototype.forEach.call(inputs, function (inp) {
      if (_aceitaImagem(inp)) anexar(inp);
    });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', function () { escanear(document); });
  } else {
    escanear(document);
  }
  // Conteúdo trocado via htmx (hx-boost) — reescaneia os novos inputs
  document.addEventListener('htmx:afterSwap', function (e) { escanear(e.target || document); });
  document.addEventListener('htmx:load', function (e) { escanear(e.target || document); });

  window.ImgCompress = { comprimir: comprimir, escanear: escanear };
})();
