# -*- coding: utf-8 -*-
import os
import sys
import threading
import time
import webbrowser
import socket
import traceback
from datetime import datetime
from pathlib import Path
from typing import Optional
from urllib.request import urlopen, Request
from urllib.error import URLError

import pystray
from PIL import Image
import uvicorn

# ---------------- Helpers de logging ----------------
def _log_path() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent / "tray.log"
    return Path(__file__).resolve().parent / "tray.log"

def _write_tray_log(message: str) -> None:
    try:
        log_file = _log_path()
        log_file.parent.mkdir(parents=True, exist_ok=True)
        with log_file.open("a", encoding="utf-8") as fh:
            fh.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {message}\n")
    except Exception:
        pass

_write_tray_log("=== Iniciando CobranzaTray ===")

# En EXE windowed (sin consola), sys.stdout/stderr pueden ser None.
# Asegura que existan para evitar fallos en librerías que asumen streams válidos.
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w")
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w")

# ---------------- Import del backend ----------------
try:
    from backend.app.main import app  # noqa
except Exception:
    _write_tray_log("Falla import backend.app.main:\n" + traceback.format_exc())
    raise

# ---------------- Parámetros base ----------------
HOST = "127.0.0.1"
PORT = int(os.getenv("COBRANZA_PORT", "8010"))
URL = f"http://{HOST}:{PORT}/"
server: Optional[uvicorn.Server] = None

# Forzar ruta del frontend cuando se ejecuta compilado
if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
    os.environ.setdefault(
        "COBRANZA_FRONTEND_DIST",
        str(Path(getattr(sys, "_MEIPASS")) / "frontend" / "dist")
    )

# ---------------- Arranque del server ----------------
def _run_server():
    global server
    try:
        config = uvicorn.Config(
        app,
        host=HOST,
        port=PORT,
        # Clave: evitar el formatter por defecto que usa sys.stderr.isatty
        log_config=None,   # evita usar sys.stderr.isatty en windowed EXE
           access_log=False,
           workers=1,
    )
        server = uvicorn.Server(config)
        if hasattr(server, "install_signal_handlers"):
            try:
               server.install_signal_handlers = (lambda: None)  # type: ignore[attr-defined]
            except Exception:
               pass
        _write_tray_log("Arrancando servidor FastAPI...")
        server.run()
        _write_tray_log("Servidor FastAPI finalizó.")
    except Exception:
        _write_tray_log("Excepción en servidor:\n" + traceback.format_exc())
        raise

def start_server_in_thread():
    thread = threading.Thread(target=_run_server, daemon=True)
    thread.start()
    return thread

def wait_for_server(timeout: float = 20.0) -> bool:
    end = time.time() + timeout
    while time.time() < end:
        try:
            with socket.create_connection((HOST, PORT), timeout=0.6):
                return True
        except OSError:
            time.sleep(0.3)
    return False

def stop_server():
    global server
    if server is not None:
        server.should_exit = True

# ---------------- UI helpers ----------------
def open_ui(_icon=None, _item=None):
    webbrowser.open(URL)

def copy_url(_icon=None, _item=None):
    try:
        import pyperclip
        pyperclip.copy(URL)
    except Exception:
        pass
    webbrowser.open(URL)

def on_quit(icon, _item):
    stop_server()
    try:
        icon.visible = False
    except Exception:
        pass
    icon.stop()

# ---------------- Iconos ----------------
def _exe_dir() -> Path:
    return Path(sys.executable).resolve().parent

def _meipass_dir() -> Optional[Path]:
    return Path(getattr(sys, "_MEIPASS")) if getattr(sys, "frozen", False) else None

def _dev_static_dir() -> Path:
    here = Path(__file__).resolve().parent
    for p in (
        here / "app" / "static",
        here / "backend" / "app" / "static",
        here.parent / "backend" / "app" / "static",
    ):
        if p.is_dir():
            return p
    return here

def load_icon_image() -> Image.Image:
    exe_dir = _exe_dir()
    candidates = [exe_dir / "app" / "static" / "cobranza.ico"]

    meipass = _meipass_dir()
    if meipass:
        candidates.append(meipass / "app" / "static" / "cobranza.ico")

    dev_static = _dev_static_dir()
    candidates += [dev_static / "cobranza.ico", dev_static / "tray.png", dev_static / "logo.png"]

    for p in candidates:
        try:
            if p.is_file():
                _write_tray_log(f"Usando icono: {p}")
                return Image.open(str(p))
        except Exception:
            continue

    _write_tray_log("Icono no encontrado; usando fallback.")
    return Image.new("RGBA", (64, 64), (20, 30, 50, 255))

# ---------------- Main ----------------
def main():
    start_server_in_thread()

    if wait_for_server():
        _write_tray_log(f"Servidor disponible en {URL}")
    else:
        # doble verificación por HTTP para dejar rastro del error real si lo hubiese
        try:
            req = Request(URL, headers={"User-Agent": "CobranzaTray"})
            with urlopen(req, timeout=2):
                pass
        except URLError as e:
            _write_tray_log(f"No se pudo verificar el servidor: {getattr(e, 'reason', repr(e))}")
        except Exception as e:
            _write_tray_log(f"No se pudo verificar el servidor: {e!r}")

        _write_tray_log(f"Advertencia: no se pudo verificar el servidor en {URL} dentro del timeout.")

    icon = pystray.Icon(
        "CobranzaXLS",
        load_icon_image(),
        "Cobranza XLS App",
        menu=pystray.Menu(
            pystray.MenuItem("Abrir", open_ui),
            pystray.MenuItem("Copiar URL y abrir", copy_url),
            pystray.MenuItem("Salir", on_quit),
        ),
    )
    icon.run()

if __name__ == "__main__":
    main()
