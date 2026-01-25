# -*- coding: utf-8 -*-
"""
Exporta un PDF por vendedor preservando el estilo de impresión del Excel.
- Requiere: pywin32 (win32com), Microsoft Excel (Windows)
- Integra: coloca este archivo en app/services/ y llama a export_vendor_pdfs(...)
"""

from __future__ import annotations
from pathlib import Path
from typing import List, Tuple, Optional, Dict
import re
import time
import unicodedata
from datetime import datetime

import pythoncom
import win32com.client as win32  # si no estaba
from pypdf import PdfReader, PdfWriter

def _col_to_index(col: str) -> int:
    """Convierte letras de columna (por ej. 'AA') a índice numérico (1-based)."""
    col = col.strip().upper()
    value = 0
    for ch in col:
        if not ch.isalpha():
            break
        value = value * 26 + (ord(ch) - 64)
    return value

def _parse_a1_bounds(a1_range: str) -> Optional[Tuple[int, int, int, int]]:
    """
    Extrae límites numéricos de un rango A1.
    Retorna (row_start, col_start, row_end, col_end) o None si no aplica.
    """
    if not a1_range:
        return None
    matches = re.findall(r"\$?([A-Z]+)\$?(\d+)", a1_range)
    if len(matches) < 2:
        return None
    (c1, r1), (c2, r2) = matches[0], matches[-1]
    return int(r1), _col_to_index(c1), int(r2), _col_to_index(c2)

def _merge_pdf_files(pdf_paths: List[Path], output_path: Path) -> Optional[Path]:
    """
    Une los PDFs en el orden recibido y devuelve la ruta del consolidado.
    """
    writer = PdfWriter()
    for pdf_file in pdf_paths:
        reader = PdfReader(str(pdf_file))
        for page in reader.pages:
            writer.add_page(page)
    if len(writer.pages) == 0:
        return None
    with output_path.open("wb") as fh:
        writer.write(fh)
    return output_path

def aplicar_layout_modelo(
    src_ws,
    tmp_ws,
    block_start: int,
    block_end: int,
    header_rows: Optional[Tuple[int, int]] = None,
) -> None:
    ps_src = src_ws.PageSetup
    bounds = _parse_a1_bounds(ps_src.PrintArea)
    if bounds:
        _, _, _, last_col = bounds
    else:
        _, _, _, last_col = _get_used_range(src_ws)
    last_col = max(1, last_col)

    header_count = 0
    if header_rows:
        header_start, header_end = header_rows
        header_start = max(1, header_start)
        header_end = max(header_start, header_end)
        header_count = header_end - header_start + 1
        src_hdr = src_ws.Range(src_ws.Cells(header_start, 1), src_ws.Cells(header_end, last_col))
        src_hdr.Copy(tmp_ws.Cells(1, 1))
        for offset, src_row in enumerate(range(header_start, header_end + 1), start=1):
            tmp_ws.Rows(offset).RowHeight = src_ws.Rows(src_row).RowHeight

    dest_block_row = header_count + 1 if header_count else 1
    block_count = max(0, block_end - block_start + 1)
    if block_count > 0:
        src_blk = src_ws.Range(src_ws.Cells(block_start, 1), src_ws.Cells(block_end, last_col))
        src_blk.Copy(tmp_ws.Cells(dest_block_row, 1))
        for idx, src_row in enumerate(range(block_start, block_end + 1)):
            dst_row = dest_block_row + idx
            tmp_ws.Rows(dst_row).RowHeight = src_ws.Rows(src_row).RowHeight

    for col_idx in range(1, last_col + 1):
        tmp_ws.Columns(col_idx).ColumnWidth = src_ws.Columns(col_idx).ColumnWidth

    ps_dst = tmp_ws.PageSetup
    ps_dst.Orientation = ps_src.Orientation
    ps_dst.PaperSize = ps_src.PaperSize
    ps_dst.Zoom = ps_src.Zoom
    ps_dst.FitToPagesWide = ps_src.FitToPagesWide
    ps_dst.FitToPagesTall = ps_src.FitToPagesTall
    ps_dst.LeftMargin = ps_src.LeftMargin
    ps_dst.RightMargin = ps_src.RightMargin
    ps_dst.TopMargin = ps_src.TopMargin
    ps_dst.BottomMargin = ps_src.BottomMargin
    ps_dst.HeaderMargin = ps_src.HeaderMargin
    ps_dst.FooterMargin = ps_src.FooterMargin
    ps_dst.CenterHorizontally = ps_src.CenterHorizontally
    ps_dst.CenterVertically = ps_src.CenterVertically
    ps_dst.PrintHeadings = ps_src.PrintHeadings
    ps_dst.PrintGridlines = ps_src.PrintGridlines
    ps_dst.PrintTitleColumns = ps_src.PrintTitleColumns
    if header_count > 0:
        ps_dst.PrintTitleRows = f"$1:${header_count}"
    else:
        ps_dst.PrintTitleRows = ""

    try:
        ps_dst.OddAndEvenPagesHeaderFooter = ps_src.OddAndEvenPagesHeaderFooter
        ps_dst.DifferentFirstPageHeaderFooter = ps_src.DifferentFirstPageHeaderFooter
    except Exception:
        pass

    try:
        ps_dst.ScaleWithDocHeaderFooter = ps_src.ScaleWithDocHeaderFooter
        ps_dst.AlignMarginsHeaderFooter = ps_src.AlignMarginsHeaderFooter
    except Exception:
        pass

    try:
        ps_dst.LeftHeader = ps_src.LeftHeader
        ps_dst.CenterHeader = ps_src.CenterHeader
        ps_dst.RightHeader = ps_src.RightHeader
        ps_dst.LeftFooter = ps_src.LeftFooter
        ps_dst.CenterFooter = ps_src.CenterFooter
        ps_dst.RightFooter = ps_src.RightFooter
    except Exception:
        pass

    total_rows = max(1, header_count + block_count)
    ps_dst.PrintArea = tmp_ws.Range(tmp_ws.Cells(1, 1), tmp_ws.Cells(total_rows, last_col)).Address

    tmp_ws.Application.CutCopyMode = False



# ------------------------
# Utilidades de texto
# ------------------------
RE_VENDEDOR = re.compile(r"\bVendedor\b[:\s]*", re.IGNORECASE)
RE_SALDO_PARA = re.compile(r"\bSaldo\s*para\b[:\s]*(.+?)\s*$", re.IGNORECASE)

def _sanitize(name: str) -> str:
    placeholder_upper = "__TILDE_N_UPPER__"
    placeholder_lower = "__TILDE_N_LOWER__"
    name_with_placeholders = (
        name.replace("Ñ", placeholder_upper).replace("ñ", placeholder_lower)
    )
    normalized = unicodedata.normalize("NFKD", name_with_placeholders)
    ascii_name = normalized.encode("ascii", "ignore").decode("ascii")
    ascii_name = (
        ascii_name.replace(placeholder_upper, "Ñ").replace(placeholder_lower, "ñ")
    )
    clean = re.sub(r'[^A-Za-z0-9 _\-\(\)Ññ]+', '_', ascii_name)
    clean = re.sub(r'\s+', ' ', clean).strip()
    return clean or re.sub(r'[^A-Za-z0-9Ññ]+', '_', name).strip()

# ------------------------
# Helpers extra
# ------------------------
def _strip_leading_code(name: str) -> str:
    """Quita prefijos numéricos tipo '000012 ' del nombre."""
    return re.sub(r"^\s*\d+\s+", "", name or "").strip()

def _date_tag_from_iso(iso_date: Optional[str]) -> Optional[str]:
    """'YYYY-MM-DD' -> 'DD-MM-YYYY' (para nombre de archivo)."""
    if not iso_date:
        return None
    try:
        dt = datetime.strptime(iso_date, "%Y-%m-%d")
        return dt.strftime("%d-%m-%Y")
    except Exception:
        return None

def _with_date_suffix(base: str, date_tag: Optional[str]) -> str:
    if not date_tag:
        return base
    return f"{base} {date_tag}"

def _make_block_id(sheet_name: str, row_start: int, row_end: int, vendor_name: str) -> str:
    return f"{sheet_name}|{row_start}-{row_end}|{vendor_name}"

SALDOS_BLOCK_ID = "__SALDOS_COBRANZA__"

# ------------------------
# Lectura de celdas (rápida)
# ------------------------
def _get_used_range(ws) -> Tuple[int, int, int, int]:
    ur = ws.UsedRange
    r1 = ur.Row
    c1 = ur.Column
    r2 = r1 + ur.Rows.Count - 1
    c2 = c1 + ur.Columns.Count - 1
    return r1, c1, r2, c2

def _cells(ws, r1: int, c1: int, r2: int, c2: int):
    # Devuelve una matriz (lista de listas) con los valores de celdas
    return ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2)).Value

# ------------------------
# Detección de bloques por vendedor
# ------------------------
def _find_vendor_blocks(ws) -> List[Dict]:
    """
    Encuentra bloques:
      - inicio: fila con "Vendedor ..."
      - fin: fila con "Saldo para <lo que sea>" (incluida)
    Retorna: lista de dict con {vendor_name, row_start, row_end}
    """
    r1, c1, r2, c2 = _get_used_range(ws)
    if r2 < r1 or c2 < c1:
        return []
    data = _cells(ws, r1, c1, r2, c2)
    if data is None:
        return []
    if not isinstance(data, (list, tuple)):
        data = [data]

    header_row = 1

    blocks: List[Dict] = []
    current_start: Optional[int] = None
    current_vendor: Optional[str] = None

    for i, row in enumerate(data, start=r1):
        row_values = row if isinstance(row, (list, tuple)) else [row]
        row_texts = [str(x) if x is not None else "" for x in row_values]
        joined = " ".join(row_texts).strip()
        if not joined and current_start is None:
            continue

        if current_start is None:
            vendor_match = RE_VENDEDOR.search(joined)
            if vendor_match:
                after = joined[vendor_match.end():].strip()
                current_vendor = after if after else joined
                current_start = i
        else:
            saldo_match = RE_SALDO_PARA.search(joined)
            if saldo_match:
                saldo_vendor = saldo_match.group(1).strip()
                vendor_name = (current_vendor or saldo_vendor or joined).strip()
                blocks.append({
                    "vendor_name": vendor_name,
                    "row_start": current_start,
                    "row_end": i,
                    "header_row": header_row
                })
                current_start = None
                current_vendor = None

    if current_start is not None:
        vendor_name = (current_vendor or f"Vendedor_{current_start}").strip()
        blocks.append({
            "vendor_name": vendor_name,
            "row_start": current_start,
            "row_end": r2,
            "header_row": header_row
        })

    return blocks

def _scan_vendor_blocks(
    wb,
    hojas_completas_set: set[str],
    target_sheet_name: Optional[str],
) -> Tuple[List[Dict], Dict[str, Optional[Tuple[int, int]]]]:
    """
    Devuelve lista de bloques con id y un dict sheet->header_rows.
    Preserva el orden de las hojas y filas.
    """
    blocks_out: List[Dict] = []
    header_rows_map: Dict[str, Optional[Tuple[int, int]]] = {}

    for ws in wb.Worksheets:
        name = str(ws.Name)
        name_clean = name.strip()

        if name_clean in hojas_completas_set:
            continue
        if target_sheet_name and name != target_sheet_name:
            continue

        blocks = _find_vendor_blocks(ws)
        if not blocks:
            continue

        first_vendor_row = min(blk["row_start"] for blk in blocks)
        header_rows: Optional[Tuple[int, int]] = None
        if first_vendor_row > 1:
            header_rows = (1, first_vendor_row - 1)
        header_rows_map[name] = header_rows

        for blk in blocks:
            block_id = _make_block_id(name, blk["row_start"], blk["row_end"], blk["vendor_name"])
            blocks_out.append({
                "id": block_id,
                "vendor_name": blk["vendor_name"],
                "row_start": blk["row_start"],
                "row_end": blk["row_end"],
                "sheet_name": name,
            })

    return blocks_out, header_rows_map

def _apply_order(preferred_ids: Optional[List[str]], available_ids: List[str]) -> List[str]:
    """
    Aplica un orden preferido, sin dejar huecos si faltan IDs.
    Los que no estÃ¡n en preferred_ids se agregan al final en su orden original.
    """
    if not preferred_ids:
        return list(available_ids)
    available_set = set(available_ids)
    ordered = [pid for pid in preferred_ids if pid in available_set]
    ordered_set = set(ordered)
    ordered.extend([vid for vid in available_ids if vid not in ordered_set])
    return ordered

def list_vendor_blocks(
    xls_path: Path,
    hojas_completas: Tuple[str, ...] = ("SUR", "NORTE", "IMPORTE CUENTA SALDO"),
    hoja_base: Optional[str] = None,
    include_saldos: bool = True,
) -> List[Dict]:
    """
    Devuelve la lista de bloques detectados (sin exportar PDFs).
    Cada item incluye: id, vendor_name, row_start, row_end, sheet_name.
    """
    initialized = False
    try:
        pythoncom.CoInitialize()
        initialized = True
    except pythoncom.com_error as exc:
        raise RuntimeError(f'No se pudo inicializar COM: {exc}') from exc

    excel = None
    wb = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            wb = excel.Workbooks.Open(str(xls_path))
        except Exception as exc:
            raise RuntimeError(f"No se pudo abrir el archivo de Excel: {exc}") from exc

        hojas_completas_set = {alias.strip() for alias in hojas_completas}
        hoja_lookup = {ws.Name.strip().lower(): ws.Name for ws in wb.Worksheets}

        hoja_base_normalized: Optional[str] = None
        if hoja_base:
            hoja_base_normalized = hoja_base.strip().lower()
            if hoja_base_normalized not in hoja_lookup:
                raise ValueError(f"No existe la hoja solicitada: {hoja_base}")

        target_sheet_name = hoja_lookup.get(hoja_base_normalized) if hoja_base_normalized else None
        blocks, _ = _scan_vendor_blocks(wb, hojas_completas_set, target_sheet_name)
        if include_saldos:
            blocks = [
                {
                    "id": SALDOS_BLOCK_ID,
                    "vendor_name": "SALDOS COBRANZA",
                    "row_start": 0,
                    "row_end": 0,
                    "sheet_name": "GENERAL",
                }
            ] + blocks
        return blocks
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        if excel is not None:
            time.sleep(0.2)
            excel.Quit()
        if initialized:
            pythoncom.CoUninitialize()


# ------------------------
# Exportación
# ------------------------
def export_vendor_pdfs(
    xls_path: Path,
    out_dir: Path,
    hojas_completas: Tuple[str, ...] = ("SUR", "NORTE", "IMPORTE CUENTA SALDO"),
    hoja_base: Optional[str] = None,
    orden_ids: Optional[List[str]] = None,
    excluir_ids: Optional[List[str]] = None,
    pdf_date: Optional[str] = None,
) -> List[Path]:
    """
    xls_path: ruta del Excel origen.
    out_dir: carpeta donde guardar los PDFs.
    hojas_completas: hojas que se exportan tal cual.
    hoja_base: si se especifica, solo procesa esa hoja para bloques de vendedor;
               si es None, intenta procesar todas las hojas no incluidas en hojas_completas.
    orden_ids: orden preferido (IDs de bloques) para el consolidado y numeracion.
    excluir_ids: IDs de bloques a excluir del consolidado (no afecta los PDFs individuales).
    pdf_date: fecha ISO (YYYY-MM-DD) para agregar al nombre de archivos.
    """
    out_dir.mkdir(parents=True, exist_ok=True)
    date_tag = _date_tag_from_iso(pdf_date)

    initialized = False
    try:
        pythoncom.CoInitialize()
        initialized = True
    except pythoncom.com_error as exc:
        raise RuntimeError(f'No se pudo inicializar COM: {exc}') from exc

    excel = None
    wb = None
    generated: List[Path] = []
    pdf_by_id: Dict[str, Path] = {}

    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        try:
            wb = excel.Workbooks.Open(str(xls_path))
        except Exception as exc:
            raise RuntimeError(f"No se pudo abrir el archivo de Excel: {exc}") from exc

        hojas_completas_set = {alias.strip() for alias in hojas_completas}
        hoja_lookup = {ws.Name.strip().lower(): ws.Name for ws in wb.Worksheets}

        hoja_base_normalized: Optional[str] = None
        if hoja_base:
            hoja_base_normalized = hoja_base.strip().lower()
            if hoja_base_normalized not in hoja_lookup:
                raise ValueError(f"No existe la hoja solicitada: {hoja_base}")

        target_sheet_name = hoja_lookup.get(hoja_base_normalized) if hoja_base_normalized else None

        for ws in wb.Worksheets:
            name = str(ws.Name)
            name_clean = name.strip()

            if name_clean in hojas_completas_set:
                pdf_base = f"COBRANZA_{_sanitize(name_clean)}"
                pdf_name = f"{_with_date_suffix(pdf_base, date_tag)}.pdf"
                pdf_path = out_dir / pdf_name
                ws.ExportAsFixedFormat(Type=0, Filename=str(pdf_path), Quality=0, IncludeDocProperties=True, IgnorePrintAreas=False, OpenAfterPublish=False)
                generated.append(pdf_path)

        blocks, header_rows_map = _scan_vendor_blocks(wb, hojas_completas_set, target_sheet_name)
        block_ids = [blk["id"] for blk in blocks]
        if orden_ids:
            ordered_block_ids = _apply_order(orden_ids, block_ids)
        else:
            ordered_block_ids = list(block_ids)
        seq_map = {bid: idx + 1 for idx, bid in enumerate(ordered_block_ids)}

        available_ids_for_merge = [SALDOS_BLOCK_ID] + block_ids
        if orden_ids:
            ordered_ids_for_merge = _apply_order(orden_ids, available_ids_for_merge)
        else:
            ordered_ids_for_merge = list(available_ids_for_merge)

        for blk in blocks:
            ws = wb.Worksheets(blk["sheet_name"])
            header_rows = header_rows_map.get(blk["sheet_name"])
            vendor_name_raw = blk["vendor_name"].strip()
            vendor_display = _strip_leading_code(vendor_name_raw)
            vendor = _sanitize(vendor_display) or "SIN_NOMBRE"
            row_start = blk["row_start"]
            row_end = blk["row_end"]
            seq = seq_map.get(blk["id"], 0)

            tmp = wb.Worksheets.Add(After=ws)
            tmp_base = f"_tmp_{vendor[:20] or 'VEN'}"
            tmp_name = tmp_base
            suffix = 1
            while True:
                try:
                    tmp.Name = tmp_name
                    break
                except Exception:
                    tmp_name = f"{tmp_base[:18]}_{suffix}"
                    suffix += 1

            try:
                aplicar_layout_modelo(ws, tmp, row_start, row_end, header_rows=header_rows)
                prefix = f"{seq:06d} " if seq else ""
                pdf_base = f"COBRANZA_{prefix}{vendor}"
                pdf_name = f"{_with_date_suffix(pdf_base, date_tag)}.pdf"
                pdf_path = out_dir / pdf_name
                tmp.ExportAsFixedFormat(Type=0, Filename=str(pdf_path), Quality=0, IncludeDocProperties=True, IgnorePrintAreas=False, OpenAfterPublish=False)
                generated.append(pdf_path)
                pdf_by_id[blk["id"]] = pdf_path
            finally:
                tmp.Delete()
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        if excel is not None:
            time.sleep(0.2)
            excel.Quit()
        if initialized:
            pythoncom.CoUninitialize()

    # Merging especial: SALDOS COBRANZA (IMPORTE CUENTA SALDO + NORTE + SUR)
    saldos_components = [
        "COBRANZA_IMPORTE CUENTA SALDO",
        "COBRANZA_SUR",
        "COBRANZA_NORTE",
    ]
    generated_by_name = {p.name.lower(): p for p in generated}
    saldos_paths: List[Path] = []
    for base in saldos_components:
        expected = f"{_with_date_suffix(base, date_tag)}.pdf".lower()
        if expected in generated_by_name:
            saldos_paths.append(generated_by_name[expected])

    saldos_path = None
    if saldos_paths:
        saldos_base = "SALDOS COBRANZA"
        saldos_name = f"{_with_date_suffix(saldos_base, date_tag)}.pdf"
        saldos_path = out_dir / saldos_name
        merged_saldos = _merge_pdf_files(saldos_paths, saldos_path)
        if merged_saldos:
            generated = [p for p in generated if p not in saldos_paths]
            generated.append(merged_saldos)
            saldos_path = merged_saldos

    # Consolidado con orden y exclusion
    exclude_set = set(excluir_ids or [])
    merge_candidates: List[Path] = []
    for bid in ordered_ids_for_merge:
        if bid in exclude_set:
            continue
        if bid == SALDOS_BLOCK_ID:
            if saldos_path and saldos_path.exists():
                merge_candidates.append(saldos_path)
            continue
        if bid in pdf_by_id and pdf_by_id[bid].exists():
            merge_candidates.append(pdf_by_id[bid])

    consolidated_base = "COBRANZA_CONSOLIDADO"
    consolidated_name = f"{_with_date_suffix(consolidated_base, date_tag)}.pdf"
    merged_path = None
    if merge_candidates:
        consolidated_path = out_dir / consolidated_name
        merged_path = _merge_pdf_files(merge_candidates, consolidated_path)
    if merged_path and merged_path not in generated:
        generated.append(merged_path)

    return generated
