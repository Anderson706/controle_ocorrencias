import sys
import os
import socket
import threading
import time
import urllib.request
import urllib.error

if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

os.chdir(BASE_DIR)

PORT = 5000
URL  = f"http://127.0.0.1:{PORT}"

# ── Flask roda em thread background — carrega app.py lá dentro ───────────────
def _run_flask():
    from app import app  # carregado só aqui, em paralelo com import webview
    app.run(host="127.0.0.1", port=PORT, debug=False, use_reloader=False, threaded=True)


def _aguardar_flask(timeout=20):
    """Aguarda o Flask abrir a porta TCP — sem fazer requisições HTTP.
    Checagem por socket evita disparar queries Oracle antes da janela abrir."""
    inicio = time.time()
    while time.time() - inicio < timeout:
        try:
            s = socket.create_connection(("127.0.0.1", PORT), timeout=0.5)
            s.close()
            return True
        except OSError:
            time.sleep(0.1)
    return False


# ── API exposta ao JavaScript ─────────────────────────────────────────────────
class Api:
    def download(self, url, filename, cookies=""):
        """Chamado pelo JS ao clicar num link de exportação."""
        import tkinter as tk
        from tkinter import filedialog

        ext = os.path.splitext(filename)[1].lower() or ".bin"
        tipos = {
            ".pdf":  [("PDF",        "*.pdf"),  ("Todos os arquivos", "*.*")],
            ".xlsx": [("Excel",      "*.xlsx"), ("Todos os arquivos", "*.*")],
            ".docx": [("Word",       "*.docx"), ("Todos os arquivos", "*.*")],
            ".pptx": [("PowerPoint", "*.pptx"), ("Todos os arquivos", "*.*")],
        }
        filetypes = tipos.get(ext, [("Todos os arquivos", "*.*")])

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        save_path = filedialog.asksaveasfilename(
            title="Salvar arquivo como",
            initialfile=filename,
            defaultextension=ext,
            filetypes=filetypes,
        )
        root.destroy()

        if not save_path:
            return {"ok": False}

        req = urllib.request.Request(f"{URL}{url}")
        if cookies:
            req.add_header("Cookie", cookies)
        try:
            resp = urllib.request.urlopen(req, timeout=30)
            data = resp.read()
            with open(save_path, "wb") as f:
                f.write(data)
            return {"ok": True, "path": save_path}
        except Exception as exc:
            return {"ok": False, "error": str(exc)}


# ── JS injetado em cada página para interceptar downloads ────────────────────
_DOWNLOAD_JS = r"""
(function () {
  function interceptar() {
    document.addEventListener('click', function (e) {
      var el = e.target;
      while (el && el.tagName !== 'A') el = el.parentElement;
      if (!el) return;

      var href = el.getAttribute('href') || '';
      if (!href || href.startsWith('#') || href.startsWith('javascript:')) return;

      var ehDownload = href.indexOf('exportar')    !== -1
                    || href.indexOf('download')    !== -1
                    || href.indexOf('comprovante') !== -1
                    || href.indexOf('export')      !== -1
                    || href.indexOf('termo')       !== -1
                    || href.indexOf('/pdf')        !== -1;
      if (!ehDownload) return;

      e.preventDefault();
      e.stopPropagation();

      var nome = el.getAttribute('data-filename') || '';
      if (!nome) {
        if (href.indexOf('excel') !== -1 || href.indexOf('xlsx') !== -1)
          nome = 'exportacao.xlsx';
        else if (href.indexOf('comprovante') !== -1)
          nome = 'comprovante_retirada.pdf';
        else if (href.indexOf('pdf') !== -1)
          nome = 'relatorio.pdf';
        else if (href.indexOf('download') !== -1)
          nome = 'analise.docx';
        else
          nome = 'arquivo';
      }

      window.pywebview.api.download(href, nome, document.cookie);
    }, true);
  }

  if (window.pywebview && window.pywebview.api) {
    interceptar();
  } else {
    window.addEventListener('pywebviewready', interceptar);
  }
})();
"""


# ── Cronometragem de inicialização (grava num log p/ diagnóstico) ─────────────
def _startup_log_path():
    try:
        if getattr(sys, 'frozen', False):
            d = os.path.dirname(sys.executable)
            if os.path.basename(d).lower() == "app":
                d = os.path.dirname(d)   # install root (sobrevive a updates)
            return os.path.join(d, "startup_timing.log")
    except Exception:
        pass
    return os.path.join(BASE_DIR, "startup_timing.log")


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    _t0 = time.time()
    _marcas = []
    def _marca(nome):
        _marcas.append((nome, time.time() - _t0))

    # 1. Inicia Flask em background (carrega app.py lá dentro)
    threading.Thread(target=_run_flask, daemon=True).start()
    _marca("flask_thread_iniciada")

    # 2. Enquanto Flask carrega app.py, carregamos webview aqui em paralelo
    #    (webview/CLR é pesado — ~5-10s — e agora sobrepõe com o import do Flask)
    import webview
    _marca("webview_importado")

    # 3. Aguarda Flask estar pronto (provavelmente já está)
    _aguardar_flask()
    _marca("flask_pronto")

    # 4. Migrações de schema — síncrono antes de abrir a janela. Com marcador local,
    #    no caso comum NÃO toca no banco (zero round-trip Oracle via VPN).
    from app import _init_db
    _init_db()
    _marca("init_db")

    # Grava o log de tempos (ajuda a diagnosticar lentidão de startup)
    try:
        from datetime import datetime as _dt
        with open(_startup_log_path(), "a", encoding="utf-8") as _lf:
            _lf.write(_dt.now().strftime("%Y-%m-%d %H:%M:%S") + "  ")
            _lf.write("  ".join(f"{n}={d:.2f}s" for n, d in _marcas) + "\n")
    except Exception:
        pass

    # 5. Fecha o splash nativo do PyInstaller
    try:
        import pyi_splash
        pyi_splash.close()
    except Exception:
        pass

    # 6. Abre a janela
    api = Api()

    window = webview.create_window(
        title="CCTV Control Panel — DHL Security",
        url=URL,
        width=1440,
        height=900,
        min_size=(1000, 650),
        resizable=True,
        text_select=False,
        confirm_close=False,
        js_api=api,
    )

    def on_loaded():
        window.evaluate_js(_DOWNLOAD_JS)
        try:
            from datetime import datetime as _dt
            with open(_startup_log_path(), "a", encoding="utf-8") as _lf:
                _lf.write(_dt.now().strftime("%Y-%m-%d %H:%M:%S") +
                          f"  janela_carregada={time.time() - _t0:.2f}s (total ate a 1a tela)\n")
        except Exception:
            pass

    window.events.loaded += on_loaded

    webview.start()


if __name__ == "__main__":
    main()
