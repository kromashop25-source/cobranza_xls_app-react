# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import sys
import tempfile
import threading
import uuid
from pathlib import Path
from typing import Callable, Dict, Optional

import logging

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Response
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from .routers import pdf as pdf_router

from .services.excel_copy import copy_first_sheet_exact, ExcelCopyError

def app_path(*parts: str) -> Path:
    """
    Devuelve la ruta correcta tanto en desarrollo como en ejecutable (PyInstaller).
    Prioriza la variable de entorno COBRANZA_BASE_DIR si existe (la seteará el launcher).
    """
    base_env = os.getenv("COBRANZA_BASE_DIR")
    if base_env:
        return Path(base_env).joinpath(*parts)
    if getattr(sys, "frozen", False):
        base_root = Path(sys._MEIPASS)  # type: ignore[attr-defined]
        base = base_root / "app"
        if not base.exists():
            base = base_root
    else:
        base = Path(__file__).parent
    return base.joinpath(*parts)


def resolve_default_master() -> Path:
    """
    Busca la ruta del maestro por defecto con reglas configurables.
    """
    env_override = os.getenv("COBRANZA_DEFAULT_MASTER")
    if env_override:
        env_path = Path(env_override)
        if env_path.is_file():
            return env_path.resolve()

    data_dir = app_path("data")
    for name in ("COBRANZA-formateado.XLS", "COBRANZA-formateado.xls"):
        candidate = data_dir / name
        if candidate.is_file():
            return candidate.resolve()

    try:
        for candidate in sorted(data_dir.glob("COBRANZA-formateado.*")):
            if candidate.is_file():
                return candidate.resolve()
    except FileNotFoundError:
        pass

    return (data_dir / "COBRANZA-formateado.XLS").resolve()


# Ruta del maestro por defecto
DEFAULT_MASTER_PATH = resolve_default_master()

logger = logging.getLogger("cobranza.app")

app = FastAPI()
app.include_router(pdf_router.router)


class SPAStaticFiles(StaticFiles):
    """
    Variante de StaticFiles que hace fallback a index.html (SPA) ante 404.
    """

    def __init__(self, *args, spa_index: str = "index.html", **kwargs):
        self.spa_index = spa_index
        super().__init__(*args, **kwargs)

    async def get_response(self, path: str, scope):
        try:
            return await super().get_response(path, scope)
        except HTTPException as exc:
            if exc.status_code == 404 and scope["method"] in ("GET", "HEAD"):
                try:
                    return await super().get_response(self.spa_index, scope)
                except HTTPException:
                    pass
            raise


def resolve_frontend_dist() -> Optional[Path]:
    """
    Intenta ubicar frontend/dist tanto en dev como en ejecutable (PyInstaller).
    """
    env_path = os.getenv("COBRANZA_FRONTEND_DIST")
    candidates = []
    if env_path:
        candidates.append(Path(env_path))

    if getattr(sys, "frozen", False):
        base = Path(getattr(sys, "_MEIPASS"))  # type: ignore[attr-defined]
        candidates.extend(
            [
                base / "frontend" / "dist",
                base / "app" / "frontend" / "dist",
                base / "app" / "dist",
                base / "dist",
            ]
        )

    module_dir = Path(__file__).resolve().parent
    repo_root = module_dir.parents[2] if len(module_dir.parents) >= 3 else module_dir
    candidates.extend(
        [
            module_dir / "frontend_dist",
            module_dir / "dist",
            repo_root / "frontend" / "dist",
            app_path("..", "..", "frontend", "dist"),
        ]
    )

    for candidate in candidates:
        try:
            if candidate and candidate.is_dir():
                return candidate.resolve()
        except Exception:
            continue
    return None


FRONTEND_DIST_PATH = resolve_frontend_dist()
LEGACY_STATIC_PATH = app_path("static")

if LEGACY_STATIC_PATH.exists():
    app.mount("/static", StaticFiles(directory=str(LEGACY_STATIC_PATH)), name="static")
else:
    logger.warning("Legacy static directory not found at %s. Skipping /static mount.", LEGACY_STATIC_PATH)


# Borrar 6 filas en la primera hoja luego del copiado
DELETE_ROWS_AFTER_PASTE = 6


# -------------------------------------------------
#                  HOME
# -------------------------------------------------




# -------------------------------------------------
#            GESTIÓN DE PROGRESO (MEMORIA)
# -------------------------------------------------
class ProgressState(Dict[str, object]):
    pass


_progress: Dict[str, ProgressState] = {}
_progress_lock = threading.Lock()


def _set_progress(
    job_id: str,
    pct: int,
    msg: str,
    status: str = "running",
    out_path: Optional[str] = None,
):
    with _progress_lock:
        st: ProgressState = _progress.setdefault(job_id, ProgressState())
        st["pct"] = int(max(0, min(100, pct)))
        st["msg"] = msg
        st["status"] = status
        if out_path is not None:
            st["out_path"] = out_path


def _progress_cb_factory(job_id: str) -> Callable[[int, str], None]:
    def cb(pct: int, msg: str) -> None:
        _set_progress(job_id, pct, msg, status="running")
    return cb


def _save_upload_to_tmp(upload: UploadFile) -> str:
    if upload is None:
        raise HTTPException(status_code=400, detail="Archivo no recibido.")
    suffix = os.path.splitext(upload.filename or "")[-1] or ".xls"
    fd, path = tempfile.mkstemp(prefix="up_", suffix=suffix)
    with os.fdopen(fd, "wb") as f:
        f.write(upload.file.read())
    return path


def _worker(
    job_id: str,
    src_path: str,
    mst_path: str,
    hdr_date: Optional[str],
    desired_name: Optional[str] = None,
) -> None:
    """Hilo que ejecuta el copiado y va reportando progreso."""
    try:
        cb = _progress_cb_factory(job_id)
        _set_progress(job_id, 1, "Preparando archivos…", status="running")
        out_path = copy_first_sheet_exact(
            src_path,
            mst_path,
            header_date=hdr_date,                       # <- recibe 'hdr_date'
            delete_first_rows=DELETE_ROWS_AFTER_PASTE,  # <- 6 filas
            progress_cb=cb,
        )
        # --- Normalizar nombre final al del archivo de origen (sin prefijos) ---
        if desired_name:
            # Asegura sólo el nombre base y extensión .xls si faltara
            base = os.path.basename(desired_name).strip() or "COBRANZA.xls"
            root, ext = os.path.splitext(base)
            if not ext:
                base = root + ".xls"
            target = os.path.join(os.path.dirname(out_path), base)
            try:
                if os.path.normcase(os.path.basename(out_path)) != os.path.normcase(base):
                    os.replace(out_path, target)
                    out_path = target
            except Exception as e:
                # No es fatal; seguimos con el path original
                _set_progress(job_id, 95, f"No se pudo renombrar el archivo: {e}", status="running")
        _set_progress(job_id, 100, "Completado.", status="done", out_path=out_path)
    except ExcelCopyError as e:
        _set_progress(job_id, 100, f"Error de Excel: {e}", status="error")
    except Exception as e:
        _set_progress(job_id, 100, f"Error: {e}", status="error")
    finally:
        # Limpieza de temporales subidos
        abs_default = os.path.abspath(str(DEFAULT_MASTER_PATH))
        for p in (src_path, mst_path):
            try:
                if p and os.path.isfile(p) and os.path.abspath(p) != abs_default:
                    os.remove(p)
            except Exception:
                pass


# -------------------------------------------------
#              ENDPOINTS LARGOS (con progreso)
# -------------------------------------------------
@app.post("/start-merge")
def start_merge(
    source: UploadFile = File(...),
    master: Optional[UploadFile] = File(default=None),
    hdr_date: Optional[str] = Form(None),
    use_default_master: int = Form(0),
):

    """
    Inicia el copiado en un hilo. Guarda el nombre ORIGINAL del archivo de
    origen para usarlo al descargar en /download/{job_id}.
    """
    orig_name = source.filename or "COBRANZA.xls"

    try:
        src_path = _save_upload_to_tmp(source)

        if use_default_master:
            # Usar el maestro alojado
            if not DEFAULT_MASTER_PATH.exists():
                raise HTTPException(status_code=500, detail="No se encuentra el maestro por defecto en app/data.")
            mst_path = str(DEFAULT_MASTER_PATH)
        else:
            # Requiere archivo maestro subido por el usuario
            if master is None:
                raise HTTPException(
                    status_code=400,
                    detail="Sube un maestro o activa 'Usar maestro por defecto'."
                )
            mst_path = _save_upload_to_tmp(master)

    finally:
        try:
            source.file.close()
        except Exception:
            pass
        try:
            if master is not None:
                master.file.close()
        except Exception:
            pass
    
    # --- Preparar hilo y devolver job_id (ESTO DEBE ESTAR DENTRO DEL ENDPOINT) ---
    job_id = uuid.uuid4().hex[:12]
    with _progress_lock:
        st: ProgressState = _progress.setdefault(job_id, ProgressState())
        st["orig_name"] = orig_name

    _set_progress(job_id, 0, "Iniciando…", status="running")
    t = threading.Thread(
        target=_worker,
        args=(job_id, src_path, mst_path, hdr_date, orig_name),  # <- pasamos el nombre deseado
        daemon=True,
    )
    t.start()
    return {"job_id": job_id}


@app.get("/progress/{job_id}")
def get_progress(job_id: str):
    with _progress_lock:
        st = _progress.get(job_id)
        if not st:
            return JSONResponse(
                {"pct": 0, "msg": "No existe el proceso.", "status": "unknown"},
                status_code=404,
            )
        return st


@app.get("/download/{job_id}")
def download(job_id: str):
    with _progress_lock:
        st = _progress.get(job_id)
        if not st:
            raise HTTPException(status_code=404, detail="Proceso no encontrado.")
        if st.get("status") != "done":
            raise HTTPException(status_code=409, detail="El proceso aún no ha finalizado.")

        out_path_obj = st.get("out_path")
        if not isinstance(out_path_obj, str) or not os.path.isfile(out_path_obj):
            raise HTTPException(status_code=404, detail="Archivo no disponible.")
        out_path: str = out_path_obj

        # ← aquí está el fix de tipado
        fname_obj = st.get("orig_name")
        if isinstance(fname_obj, str) and fname_obj.strip():
            # asegúrate de que sea solo el nombre base
            fname: str = os.path.basename(fname_obj.strip())
        else:
            fname = os.path.basename(out_path)

    return FileResponse(out_path, filename=fname, media_type="application/vnd.ms-excel")


# -------------------------------------------------
#      COMPATIBILIDAD CON /merge (sin progreso)
# -------------------------------------------------
@app.post("/merge")
def merge_compat(
    source: UploadFile = File(...),
    master: Optional[UploadFile] = File(default=None),
    hdr_date: Optional[str] = Form(None),
    use_default_master: int = Form(0),
):

    """
    Endpoint simple de compatibilidad (sin progreso).
    Devuelve el archivo usando el **nombre del archivo origen**.
    """
    orig_name = source.filename or "COBRANZA.xls"

    try:
        src_path = _save_upload_to_tmp(source)
        if use_default_master:
            mst_path = str(DEFAULT_MASTER_PATH)
        else:
            if master is None:
                raise HTTPException(
                    status_code=400,
                    detail="Sube un maestro o activa 'Usar maestro por defecto'."
                )
            mst_path = _save_upload_to_tmp(master)
        
        out_path = copy_first_sheet_exact(
            src_path,
            mst_path,
            header_date=hdr_date,                       # <- recibe 'hdr_date'
            delete_first_rows=DELETE_ROWS_AFTER_PASTE,  # <- 6 filas
            progress_cb=None,
        )
        # Descargar con el **nombre original** del archivo de origen (nombre base)
        return FileResponse(
            out_path,
            filename=os.path.basename(orig_name.strip() or "COBRANZA.xls"),
            media_type="application/vnd.ms-excel",
        )
    except ExcelCopyError as e:
        raise HTTPException(status_code=500, detail=f"Error de Excel: {e}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        try:
            source.file.close()
        except Exception:
            pass
        try:
            if master is not None:
                master.file.close()
        except Exception:
            pass

@app.get("/master/default-info")
def master_default_info():
    path = Path(DEFAULT_MASTER_PATH)
    exists = path.is_file()
    payload: Dict[str, Optional[str] | bool] = {
        "exists": exists,
        "name": path.name if exists else None,
    }
    if os.getenv("COBRANZA_DEBUG") == "1":
        payload["debug_path"] = str(path)
    return payload


# -------------------------------------------------
#              SPA / FRONTEND STATIC BUILD
# -------------------------------------------------
SPA_ROOT_DIR: Optional[Path] = None
for candidate in (FRONTEND_DIST_PATH, LEGACY_STATIC_PATH):
    if candidate and candidate.exists():
        SPA_ROOT_DIR = candidate
        break

if SPA_ROOT_DIR is not None:
    app.mount(
        "/",
        SPAStaticFiles(directory=str(SPA_ROOT_DIR), html=True),
        name="frontend-app",
    )
else:
    logger.error(
        "No frontend assets found. Expected build under %s or %s. React UI will not be served.",
        FRONTEND_DIST_PATH,
        LEGACY_STATIC_PATH,
    )

if (
    FRONTEND_DIST_PATH
    and FRONTEND_DIST_PATH.exists()
    and LEGACY_STATIC_PATH.exists()
):
    app.mount(
        "/legacy",
        StaticFiles(directory=str(LEGACY_STATIC_PATH), html=True),
        name="legacy-ui",
    )

