import threading
import webview
from app import app


def start_flask():
    app.run(host="127.0.0.1", port=5050, debug=False, use_reloader=False)


if __name__ == "__main__":
    t = threading.Thread(target=start_flask, daemon=True)
    t.start()

    webview.create_window(
        "DHL Security — Gerenciador de Usuários",
        "http://127.0.0.1:5050",
        width=1280,
        height=800,
        resizable=True,
        min_size=(1024, 600),
    )
    webview.start()
