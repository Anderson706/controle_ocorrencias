import sys
import os
import threading
import time
import urllib.request
import tkinter as tk
from tkinter import filedialog
import webview

if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

os.chdir(BASE_DIR)

from app import app  # noqa: E402

PORT = 5000
URL  = f"http://127.0.0.1:{PORT}"


# ── API exposta ao JavaScript ─────────────────────────────────────────────────
class Api:
    def download(self, url, filename, cookies=""):
        """Chamado pelo JS ao clicar num link de exportação."""
        ext = os.path.splitext(filename)[1].lower() or ".bin"
        tipos = {
            ".pdf":  [("PDF",   "*.pdf"),  ("Todos os arquivos", "*.*")],
            ".xlsx": [("Excel", "*.xlsx"), ("Todos os arquivos", "*.*")],
            ".docx": [("Word",  "*.docx"), ("Todos os arquivos", "*.*")],
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
      // sobe até encontrar o <a>
      var el = e.target;
      while (el && el.tagName !== 'A') el = el.parentElement;
      if (!el) return;

      var href = el.getAttribute('href') || '';
      if (!href || href.startsWith('#') || href.startsWith('javascript:')) return;

      var ehDownload = href.indexOf('exportar') !== -1 || href.indexOf('download') !== -1;
      if (!ehDownload) return;

      e.preventDefault();
      e.stopPropagation();

      // Usa data-filename se disponível, senão fallback pelo tipo de URL
      var nome = el.getAttribute('data-filename') || '';
      if (!nome) {
        if (href.indexOf('excel') !== -1 || href.indexOf('xlsx') !== -1)
          nome = 'exportacao.xlsx';
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


# ── Flask + WebView ───────────────────────────────────────────────────────────
def _run_flask():
    app.run(host="127.0.0.1", port=PORT, debug=False, use_reloader=False)


def _aguardar_flask(timeout=15):
    inicio = time.time()
    while time.time() - inicio < timeout:
        try:
            urllib.request.urlopen(URL, timeout=1)
            return True
        except Exception:
            time.sleep(0.1)
    return False


def main():
    threading.Thread(target=_run_flask, daemon=True).start()
    _aguardar_flask()

    # Fecha o splash nativo do PyInstaller
    try:
        import pyi_splash
        pyi_splash.close()
    except Exception:
        pass

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

    window.events.loaded += on_loaded

    webview.start()


if __name__ == "__main__":
    main()
