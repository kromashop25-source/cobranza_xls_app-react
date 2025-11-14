# run_app.py
import os
import sys
import time
import threading
import socket
import webbrowser
import traceback
from pathlib import Path

LOG = Path("CobranzaApp_error.log")

def _log(msg: str) -> None:
    try:
        LOG.write_text((LOG.read_text(encoding="utf-8") if LOG.exists() else "") + msg + "\n", encoding="utf-8")
    except Exception:
        pass

def base_dir_app() -> Path:
    """
    Carpeta /app con los assets (index.html, static/, data/)
    tanto en desarrollo como cuando está empaquetado (PyInstaller onefile).
    """
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS) / "app"  # type: ignore[attr-defined]
    return Path(__file__).parent / "app"

def find_free_port(preferred: int = 8010) -> int:
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    try:
        s.bind(("127.0.0.1", preferred))
        port = preferred
    except OSError:
        s.bind(("127.0.0.1", 0))
        port = s.getsockname()[1]
    finally:
        s.close()
    return port

def serve(port: int):
    try:
        base = base_dir_app()
        os.environ["COBRANZA_BASE_DIR"] = str(base)
        _log(f"[serve] COBRANZA_BASE_DIR={base}")

        idx1 = base / "index.html"
        idx2 = base / "static" / "index.html"
        _log(f"[serve] index candidates -> {idx1} (exists={idx1.exists()}), {idx2} (exists={idx2.exists()})")
        _log(f"[serve] static dir -> {(base / 'static')} (exists={(base / 'static').exists()})")
        _log(f"[serve] data dir   -> {(base / 'data')} (exists={(base / 'data').exists()})")

        # Pre-import para capturar y loguear errores de importación
        import importlib
        importlib.import_module("app.main")
        _log("[serve] import app.main OK")

        import uvicorn
        uvicorn.run(
            "app.main:app",
            host="127.0.0.1",
            port=port,
            lifespan="on",
            log_config=None,      # ⬅️ clave: evita DefaultFormatter (sys.stderr es None en --noconsole)
            log_level="warning",
            access_log=False,     # opcional: menos ruido
        )

    except Exception:
        _log("[serve] FATAL\n" + traceback.format_exc())

def wait_ready(port: int, timeout: float = 40.0) -> bool:
    import http.client
    t0 = time.time()
    while time.time() - t0 < timeout:
        try:
            conn = http.client.HTTPConnection("127.0.0.1", port, timeout=0.6)
            conn.request("GET", "/")
            resp = conn.getresponse()
            if 200 <= resp.status < 500:
                _log(f"[ready] OK status={resp.status}")
                return True
        except Exception:
            time.sleep(0.25)
    _log("[ready] TIMEOUT")
    return False

if __name__ == "__main__":
    port = find_free_port(8010)
    _log(f"[main] chosen port={port}")
    th = threading.Thread(target=serve, args=(port,), daemon=True)
    th.start()

    if wait_ready(port):
        try:
            webbrowser.open(f"http://127.0.0.1:{port}/")
        except Exception:
            pass
        th.join()
    else:
        _log("No se pudo iniciar el servidor en el tiempo esperado.")
        sys.exit(1)
