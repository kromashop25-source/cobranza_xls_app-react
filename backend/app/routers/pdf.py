# app/routers/pdf.py
import io
import traceback
import zipfile
from datetime import datetime
from pathlib import Path
from tempfile import TemporaryDirectory

from fastapi import (
    APIRouter,
    File,
    Form,
    HTTPException,
    Response,
    UploadFile,
)

from ..services.pdf_export_service import export_vendor_pdfs

router = APIRouter(prefix="/pdf", tags=["pdf"])

LOG_FILE = Path("pdf_export.log")


def _log(msg: str) -> None:
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
    with LOG_FILE.open("a", encoding="utf-8") as fh:
        fh.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}\n")


@router.post("/export")
def exportar_pdfs(file_path: str, hoja_base: str | None = None, carpeta_salida: str | None = None):
    """
    Genera PDFs por vendedor. SUR y NORTE se exportan completas.
    - file_path: ruta completa al Excel.
    - hoja_base: limita el procesamiento a esa hoja; si None, busca en todas (excepto SUR/NORTE).
    - carpeta_salida: si no se indica, crea 'PDFS' junto al Excel.
    """
    try:
        xls = Path(file_path).resolve()
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Ruta invalida: {exc}")

    if not xls.exists():
        raise HTTPException(status_code=400, detail=f"No existe el archivo: {xls}")

    if carpeta_salida:
        try:
            out = Path(carpeta_salida).resolve()
        except Exception as exc:
            raise HTTPException(status_code=400, detail=f"Carpeta de salida invalida: {exc}")
    else:
        out = xls.parent / "PDFS"

    try:
        out.mkdir(parents=True, exist_ok=True)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"No se pudo crear la carpeta de salida: {exc}")

    try:
        pdfs = export_vendor_pdfs(
            xls_path=xls,
            out_dir=out,
            hojas_completas=("SUR", "NORTE"),
            hoja_base=hoja_base
        )
    except (ValueError, RuntimeError) as exc:
        _log(f"ERROR export: {exc}")
        raise HTTPException(status_code=400, detail=str(exc))
    except Exception as exc:
        tb = traceback.format_exc()
        _log(f"ERROR export: {exc}\n{tb}")
        raise HTTPException(status_code=500, detail=f"Fallo exportando PDFs: {exc}")

    _log(f"OK export: {xls} -> {out} | {len(pdfs)} archivos")
    return {"status": "ok", "count": len(pdfs), "out_dir": str(out), "files": [str(p) for p in pdfs]}


@router.post("/export-upload")
async def export_pdfs_upload(
    excel: UploadFile = File(..., description="XLS formateado"),
    hoja_base: str | None = Form(None),
):
    """
    Recibe un XLS formateado subido por el usuario y devuelve un ZIP con los PDFs.
    """
    fname = (excel.filename or "").strip()
    if not fname.lower().endswith(".xls"):
        raise HTTPException(status_code=400, detail="Se espera un .XLS")

    with TemporaryDirectory(prefix="pdf_exp_") as td:
        tmp_dir = Path(td)
        safe_name = Path(fname).name or "COBRANZA.xls"
        xls_path = tmp_dir / safe_name
        xls_path.write_bytes(await excel.read())

        out_dir = xls_path.parent / "PDFS"
        out_dir.mkdir(parents=True, exist_ok=True)

        files = export_vendor_pdfs(
            xls_path=xls_path,
            out_dir=out_dir,
            hojas_completas=("SUR", "NORTE"),
            hoja_base=hoja_base,
        )
        if not files:
            raise HTTPException(
                status_code=409,
                detail="No se detectaron bloques de vendedores ni hojas SUR/NORTE.",
            )

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
            for file_path in files:
                z.write(file_path, arcname=file_path.name)
        buf.seek(0)

    zip_name = f"PDFS_{Path(safe_name).stem}.zip"
    return Response(
        content=buf.read(),
        media_type="application/zip",
        headers={"Content-Disposition": f'attachment; filename="{zip_name}"'},
    )


@router.get("/debug-blocks")
def debug_blocks(file_path: str, hoja: str):
    """
    Devuelve los bloques detectados (Vendedor... -> Saldo para...) en la hoja indicada.
    """
    try:
        from ..services.pdf_export_service import _find_vendor_blocks
        import win32com.client as win32
        import pythoncom

        xls = Path(file_path).resolve()
        if not xls.exists():
            raise HTTPException(status_code=400, detail=f"No existe el archivo: {xls}")

        initialized = False
        try:
            pythoncom.CoInitialize()
            initialized = True
        except pythoncom.com_error as exc:
            raise HTTPException(status_code=500, detail=f'No se pudo inicializar COM: {exc}')

        excel = None
        wb = None
        try:
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            try:
                wb = excel.Workbooks.Open(str(xls))
                try:
                    ws = wb.Worksheets(hoja)
                except Exception:
                    raise HTTPException(status_code=400, detail=f"No existe la hoja: {hoja}")
                blocks = _find_vendor_blocks(ws)
                out = [{"vendor": b["vendor_name"], "row_start": b["row_start"], "row_end": b["row_end"]} for b in blocks]
                return {"sheet": hoja, "count": len(out), "blocks": out}
            finally:
                if wb is not None:
                    try:
                        wb.Close(SaveChanges=False)
                    except Exception:
                        pass
        finally:
            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass
            if initialized:
                pythoncom.CoUninitialize()
    except HTTPException:
        raise
    except Exception as exc:
        tb = traceback.format_exc()
        _log(f"ERROR debug-blocks: {exc}\n{tb}")
        raise HTTPException(status_code=500, detail=f"Fallo en debug-blocks: {exc}")
