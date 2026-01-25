# -*- coding: utf-8 -*-
"""
excel_copy.py
-------------
Copia hoja 1 del XLS de origen al XLS maestro (con formatos), elimina 6
primeras filas, actualiza fechas (encabezado de la 1.ª hoja y títulos de
SUR/NORTE/GENERAL) y guarda el archivo con el **mismo nombre** del archivo
ORIGEN. Reporta progreso por callback.

Requiere: pywin32 (pythoncom, win32com.client).
"""
from __future__ import annotations

import os
import re
import tempfile
import uuid
import time
import unicodedata
# Fuerza inclusión en el .exe
import win32timezone  # noqa: F401

from datetime import datetime
from typing import Optional, Callable, Dict, List, Tuple

import pythoncom
from win32com.client import DispatchEx


# -------------------------------------------------------------------
# Excepción propia
# -------------------------------------------------------------------
class ExcelCopyError(Exception):
    """Errores específicos del proceso de copiado Excel."""


# -------------------------------------------------------------------
# Configuración
# -------------------------------------------------------------------
DELETE_FIRST_ROWS = 6            # N filas a borrar en la Hoja1 tras el pegado
HEADER_SCAN_ROWS = 8             # Cuántas filas superiores considerar "encabezado"
HEADER_SCAN_COLS = 60            # Cuántas columnas a revisar para fecha en celdas
TARGET_NAME_COL = 2              # Columna B para nombres de vendedor en hojas destino
TARGET_VAL_COLS = (3, 4)         # Columnas C (Importe) y D (A cuenta)
MAX_LOOKAHEAD_VALUES = 20        # Cuántas columnas hacia la derecha buscar valores
PASTE_VALUES_PER_VENDOR = 2      # Solo los 2 primeros valores (Importe, A cuenta)

# Alias de "Saldo para ..." -> cómo aparece el vendedor en col B de la pestaña destino
ALIAS_MAP = {
    "OFICINA (VES)": "OFICINA",
    "MANUEL CARRASCO": "MANUEL",
    "PITER HUAYTA": "PITER",
    "BEATRIZ ROJAS": "BEATRIZ",
    "LEONEL MEZA": "LEONEL",
    "CAÑETE - (MANUEL)": "CAÑETE",
    "CIUDAD - (ROSA)": "CIUDAD",
    "LURIN - (ROSA)": "LURIN",
    "MANCHAY - (ROSA)": "MANCHAY",
    "UNICACHI SUR - (ROSA)": "UNICACHI",
    "SURQ/SURCO - (OSCAR)": "SURQ/SURCO",      # va a pestaña SURQUILLO
    "SAN LUIS (OSCAR)": "SAN LUIS",
    "CAQUETA (ROSA)": "CAQUETA",
    "SURQUILLO (OSCAR)": "SURQUILLO", 
    "SURCO (OSCAR)": "SURCO", # texto en col B
    "NORTE - ROSA": "NORTE",
    "RAUL ARROYO": "RAUL",
}

DATE_RE = re.compile(r"\b(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})\b", re.IGNORECASE)
SALDO_PARA_RE = re.compile(r"^\s*saldo\s+para\s+(.+?)\s*$", re.IGNORECASE)


def _progress_notify(progress_cb: Callable[[int, str], None] | None, pct: int, msg: str) -> None:
    if progress_cb is None:
        return
    try:
        progress_cb(int(max(0, min(100, pct))), msg)
    except Exception:
        pass

def _paste_all_robust(excel, src_ws, dst_ws):
    """Estrategia robusta de pegado."""
    last = ""
    try:
        src_ws.UsedRange.Copy(Destination=dst_ws.Range("A1"))
        return "Copy(Destination)"
    except Exception as e1:
        last = f"Copy(Destination) -> {e1}"

    try:
        excel.CutCopyMode = False
        src_ws.UsedRange.Copy()
        time.sleep(0.2)
        dst_ws.Paste(Destination=dst_ws.Range("A1"))
        return "Worksheet.Paste"
    except Exception as e2:
        last += f" | Worksheet.Paste -> {e2}"

    try:
        excel.CutCopyMode = False
        src_ws.UsedRange.Copy()
        time.sleep(0.2)
        dst_ws.Range("A1").PasteSpecial(Paste=-4104)  # xlPasteAll
        return "Range.PasteSpecial(xlPasteAll)"
    except Exception as e3:
        last += f" | PasteSpecial(xlPasteAll) -> {e3}"
        raise ExcelCopyError(last)


def _iso_to_es_ddmmyyyy(iso_date: str) -> str | None:
    try:
        dt = datetime.strptime(iso_date, "%Y-%m-%d")
        return dt.strftime("%d/%m/%Y")
    except Exception:
        return None


def _update_header_date_in_cells(dst_ws, header_date_iso: str, header_rows: int, header_cols: int):
    """Reemplaza fechas en TEXTO dentro de celdas del encabezado."""
    es_date = _iso_to_es_ddmmyyyy(header_date_iso)
    if not es_date:
        return
    used = dst_ws.UsedRange
    rows = min(header_rows, used.Rows.Count)
    cols = min(header_cols, used.Columns.Count)

    for r in range(1, rows + 1):
        for c in range(1, cols + 1):
            cell = dst_ws.Cells(r, c)
            val = cell.Value
            if isinstance(val, str) and DATE_RE.search(val):
                try:
                    cell.Value = DATE_RE.sub(es_date, val)
                except Exception:
                    pass


def _replace_date_in_shapes(dst_ws, es_date: str):
    """Reemplaza fechas en cuadros de texto (Shapes)."""
    try:
        for shp in dst_ws.Shapes:
            # TextFrame2 (Office 2010+)
            try:
                if shp.TextFrame2.HasText:
                    t = shp.TextFrame2.TextRange.Text
                    if t and DATE_RE.search(t):
                        shp.TextFrame2.TextRange.Text = DATE_RE.sub(es_date, t)
                        continue
            except Exception:
                pass
            # TextFrame clásico
            try:
                t = shp.TextFrame.Characters().Text
                if t and DATE_RE.search(t):
                    shp.TextFrame.Characters().Text = DATE_RE.sub(es_date, t)
            except Exception:
                pass
    except Exception:
        pass


def _replace_date_in_page_headers(dst_ws, es_date: str):
    """Reemplaza fechas en encabezados/pies de página (PageSetup)."""
    try:
        ps = dst_ws.PageSetup
        for attr in ("LeftHeader", "CenterHeader", "RightHeader",
                     "LeftFooter", "CenterFooter", "RightFooter"):
            try:
                s = getattr(ps, attr, "")
                if isinstance(s, str) and s and DATE_RE.search(s):
                    setattr(ps, attr, DATE_RE.sub(es_date, s))
            except Exception:
                pass
    except Exception:
        pass


def _norm(s: str) -> str:
    """Normaliza: quita acentos, mayúsculas, colapsa espacios."""
    t = unicodedata.normalize("NFD", s)
    t = "".join(ch for ch in t if unicodedata.category(ch) != "Mn")
    t = " ".join(t.split())
    return t.upper()

def _tokens(norm_text: str) -> list[str]:
    return [t for t in re.split(r"[^A-Z0-9]+", norm_text) if len(t) >= 3]

def _is_fuzzy_match(cell_norm: str, vendor_norm: str) -> bool:
    # Match exacto solamente (nombres fijos del sistema).
    return cell_norm == vendor_norm

def _alias_map_norm() -> dict[str, str]:
    return { _norm(k): _norm(v) for k, v in ALIAS_MAP.items() }

def _map_vendor_key_to_target(vendor_key_norm: str) -> str:
    return _alias_map_norm().get(vendor_key_norm, vendor_key_norm)




def _try_number(val) -> float | None:
    """Devuelve número float si val ya es numérico o si es texto con miles/decimal (es-PE), si no None."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    if isinstance(val, str):
        # Extrae primer número estilo 1.234,56 (o 1234,56 / 1.234 / 1234)
        m = re.search(r"[-+]?\d{1,3}(?:\.\d{3})*(?:,\d+)?|[-+]?\d+(?:,\d+)?", val)
        if not m:
            return None
        txt = m.group(0)
        txt = txt.replace(".", "").replace(",", ".")
        try:
            return float(txt)
        except Exception:
            return None
    return None


def _collect_vendor_totals_from_sheet1(dst_ws) -> dict[str, tuple[float | None, float | None, float | None]]:
    """
    Busca filas con "Saldo para <VENDEDOR>" y recoge hasta 3 valores numéricos
    en esa misma fila, hacia la derecha del rótulo.
    Devuelve dict: { vendedor_normalizado: (v1, v2, v3) }
    """
    results: dict[str, tuple[float | None, float | None, float | None]] = {}
    used = dst_ws.UsedRange
    rows = used.Rows.Count
    cols = used.Columns.Count

    for r in range(1, rows + 1):
        for c in range(1, cols + 1):
            cell = dst_ws.Cells(r, c)
            val = cell.Value
            if isinstance(val, str):
                m = SALDO_PARA_RE.match(val)
                if m:
                    raw_name = m.group(1)
                    vend_key = _norm(raw_name)
                    # Buscar hasta 3 valores numéricos en la misma fila, a la derecha
                    values = []
                    for j in range(c + 1, min(cols, c + 1 + MAX_LOOKAHEAD_VALUES)):
                        num = _try_number(dst_ws.Cells(r, j).Value)
                        if num is not None:
                            values.append(num)
                        if len(values) >= 3:
                            break
                    # Rellenar con None si faltan
                    while len(values) < 3:
                        values.append(None)
                    results[vend_key] = (values[0], values[1], values[2])
                    break  # pasa a siguiente fila tras hallar el rótulo
    return results


def _write_vendor_values_to_other_sheets(dst_wb, vendor_map: dict[str, tuple[float | None, float | None, float | None]]):
    """
    Escribe v1,v2 (Importe, Cuenta) en columnas C y D de la fila cuyo col B coincida
    con el vendedor destino. Usa alias y match difuso. Recorre en orden: SUR → NORTE → SURQUILLO → resto.
    """
    if not vendor_map:
        return

    # 1) Normaliza claves del vendor_map y aplica alias → target
    # vendor_map viene con claves normalizadas? Si no, normalizamos aquí por seguridad.
    norm_vendor_map = { _norm(k): v for k, v in vendor_map.items() }
    mapped: dict[str, tuple[float | None, float | None, float | None]] = {}
    for key_norm, vals in norm_vendor_map.items():
        target_norm = _map_vendor_key_to_target(key_norm)
        # Si dos orígenes mapean al mismo destino, el último gana (si quieres sumar, lo cambio).
        mapped[target_norm] = vals

    # 2) Orden de recorrido de hojas
    total = dst_wb.Worksheets.Count
    preferred = ["SUR", "NORTE", "SURQUILLO"]  # <--- añadimos SURQUILLO
    indices_ordenados: list[int] = []
    vistos = set()

    def _find_idx(name_norm: str) -> int | None:
        for i in range(2, total + 1):
            ws = dst_wb.Worksheets(i)
            if _norm(ws.Name) == name_norm:
                return i
        return None

    for p in preferred:
        idx = _find_idx(_norm(p))
        if idx is not None:
            indices_ordenados.append(idx)
            vistos.add(idx)
    for i in range(2, total + 1):
        if i not in vistos:
            indices_ordenados.append(i)

    # 3) Escribir valores
    target_keys = list(mapped.keys())  # ya normalizadas
    for i in indices_ordenados:
        ws = dst_wb.Worksheets(i)
        used = ws.UsedRange
        rows = used.Rows.Count

        for r in range(1, rows + 1):
            raw = ws.Cells(r, TARGET_NAME_COL).Value
            if not isinstance(raw, str):
                continue
            cell_norm = _norm(raw)

            # Busca coincidencia exacta o difusa con keys destino
            match_key = None
            if cell_norm in mapped:
                match_key = cell_norm
            else:
                for tk in target_keys:
                    if _is_fuzzy_match(cell_norm, tk):
                        match_key = tk
                        break

            if match_key:
                v1, v2, _ = mapped.get(match_key, (None, None, None))
                try:
                    if v1 is not None:
                        ws.Cells(r, TARGET_VAL_COLS[0]).Value = v1  # C
                    if v2 is not None:
                        ws.Cells(r, TARGET_VAL_COLS[1]).Value = v2  # D
                except Exception:
                    pass
                # Sigue buscando otras filas (por si hay varias áreas con el mismo vendedor).



def copy_first_sheet_exact(
    source_xls_path: str,
    master_xls_path: str,
    *,
    header_date: str | None = None,
    delete_first_rows: int = DELETE_FIRST_ROWS,
    progress_cb: Callable[[int, str], None] | None = None,
) -> str:
    """Pipeline principal."""
    if not os.path.isfile(source_xls_path):
        raise ExcelCopyError(f"No existe el archivo origen: {source_xls_path}")
    if not os.path.isfile(master_xls_path):
        raise ExcelCopyError(f"No existe el archivo maestro: {master_xls_path}")

    XL_XLS_FORMAT = 56  # .xls

    def notify(pct: int, message: str) -> None:
        _progress_notify(progress_cb, pct, message)

    notify(5, "Inicializando Excel...")


    pythoncom.CoInitialize()
    excel = None
    src_wb = None
    dst_wb = None
    try:
        excel = DispatchEx("Excel.Application")  # instancia nueva evita conflictos
        excel.Visible = False
        excel.DisplayAlerts = False

        src_wb = excel.Workbooks.Open(source_xls_path, UpdateLinks=0, ReadOnly=True)
        dst_wb = excel.Workbooks.Open(master_xls_path, UpdateLinks=0, ReadOnly=False)
        notify(25, "Abriendo libros en Excel...")

        src_ws = src_wb.Worksheets(1)
        dst_ws = dst_wb.Worksheets(1)
        notify(45, "Copiando hoja de origen...")

        # Limpiar y pegar robusto
        dst_ws.Cells.Clear()
        _paste_all_robust(excel, src_ws, dst_ws)
        notify(60, "Pegado completo. Aplicando ajustes...")

        # Borrar primeras N filas
        if delete_first_rows and delete_first_rows > 0:
            try:
                dst_ws.Rows(f"1:{delete_first_rows}").Delete()
            except Exception:
                try:
                    dst_ws.UsedRange.UnMerge()
                    dst_ws.Rows(f"1:{delete_first_rows}").Delete()
                except Exception:
                    pass

        # Reemplazo de fecha en encabezado
        if header_date:
            try:
                _update_header_date_in_cells(dst_ws, header_date, HEADER_SCAN_ROWS, HEADER_SCAN_COLS)
                es_date = _iso_to_es_ddmmyyyy(header_date)
                if es_date:
                    _replace_date_in_shapes(dst_ws, es_date)
                    _replace_date_in_page_headers(dst_ws, es_date)
            except Exception:
                pass

        # Ajuste básico de anchos de columnas según origen (opcional)
        try:
            used_src = src_ws.UsedRange
            used_dst = dst_ws.UsedRange
            cols = min(used_dst.Columns.Count, used_src.Columns.Count)
            for c in range(1, cols + 1):
                dst_ws.Columns(c).ColumnWidth = src_ws.Columns(c).ColumnWidth
        except Exception:
            pass

        # === NUEVO: leer “Saldo para <VENDEDOR>” de Hoja1 y escribir a otras hojas
        vendor_map = _collect_vendor_totals_from_sheet1(dst_ws)
        # Solo los dos primeros valores; el tercero (saldo) lo calculan fórmulas en destino
        _write_vendor_values_to_other_sheets(dst_wb, vendor_map)
        notify(80, "Actualizando hojas destino...")
         # --- Actualiza el título de SUR y NORTE con la fecha seleccionada ---
        if header_date:
            try:
                for sheet_name in ("SUR", "NORTE"):
                    ws_title = _find_ws_by_name_norm(dst_wb, sheet_name)
                    if ws_title is not None:
                        _update_sheet_title_cobranza(ws_title, header_date, search_rows=6, search_cols=30)
            except Exception:
                pass

        excel.CutCopyMode = False

        # Guardar con nombre único
        out_dir = tempfile.mkdtemp(prefix="cobranza_xls_")
        ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        base_name = f"maestro_copiado_{ts}.xls"
        out_path = os.path.join(out_dir, base_name)
        try:
            notify(90, "Guardando archivo resultado...")
            dst_wb.SaveAs(out_path, FileFormat=XL_XLS_FORMAT)
        except Exception:
            base_name = f"maestro_copiado_{ts}_{uuid.uuid4().hex[:8]}.xls"
            out_path = os.path.join(out_dir, base_name)
            dst_wb.SaveAs(out_path, FileFormat=XL_XLS_FORMAT)

        notify(99, "Archivo listo.")
        return out_path

    except Exception as e:
        raise ExcelCopyError(str(e))
    finally:
        try:
            if src_wb is not None:
                src_wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if dst_wb is not None:
                dst_wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()
# --- Meses en español en MAYÚSCULAS (con SETIEMBRE como en Perú) ---
_ES_MESES = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
    9: "SETIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
}

def _es_title_from_iso(iso_date: str) -> str | None:
    """
    'YYYY-MM-DD' -> 'COBRANZA AL DD MES YYYY' (todo mayúsculas)
    """
    try:
        dt = datetime.strptime(iso_date, "%Y-%m-%d")
        mes = _ES_MESES.get(dt.month, "")
        return f"COBRANZA AL {dt.strftime('%d')} {mes} {dt.year}"
    except Exception:
        return None

def _update_sheet_title_cobranza(ws, iso_date: str, search_rows: int = 6, search_cols: int = 20):
    """
    Busca en las primeras filas un texto que contenga 'COBRANZA AL' (en celdas y en shapes)
    y lo reemplaza COMPLETO por el título target.
    """
    title = _es_title_from_iso(iso_date)
    if not title:
        return

    # 1) En celdas (primeras filas/columnas)
    try:
        used = ws.UsedRange
        rows = min(search_rows, used.Rows.Count)
        cols = min(search_cols, used.Columns.Count)
        for r in range(1, rows + 1):
            for c in range(1, cols + 1):
                cell = ws.Cells(r, c)
                v = cell.Value
                if isinstance(v, str) and "COBRANZA AL" in v.upper():
                    cell.Value = title
    except Exception:
        pass

    # 2) En shapes (cuadros de texto)
    try:
        for shp in ws.Shapes:
            repl = False
            # TextFrame2
            try:
                if shp.TextFrame2.HasText:
                    t = shp.TextFrame2.TextRange.Text or ""
                    if "COBRANZA AL" in t.upper():
                        shp.TextFrame2.TextRange.Text = title
                        repl = True
            except Exception:
                pass
            if repl:
                continue
            # TextFrame clásico
            try:
                t = shp.TextFrame.Characters().Text or ""
                if "COBRANZA AL" in t.upper():
                    shp.TextFrame.Characters().Text = title
            except Exception:
                pass
    except Exception:
        pass

def _find_ws_by_name_norm(wb, name: str):
    tgt = _norm(name)
    for i in range(1, wb.Worksheets.Count + 1):
        ws = wb.Worksheets(i)
        if _norm(ws.Name) == tgt:
            return ws
    return None
