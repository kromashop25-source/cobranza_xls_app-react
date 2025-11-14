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


# ------------------------
# Exportación
# ------------------------
def export_vendor_pdfs(
    xls_path: Path,
    out_dir: Path,
    hojas_completas: Tuple[str, ...] = ("SUR", "NORTE"),
    hoja_base: Optional[str] = None
) -> List[Path]:
    """
    xls_path: ruta del Excel origen.
    out_dir: carpeta donde guardar los PDFs.
    hojas_completas: hojas que se exportan tal cual.
    hoja_base: si se especifica, solo procesa esa hoja para bloques de vendedor;
               si es None, intenta procesar todas las hojas no incluidas en hojas_completas.
    """
    out_dir.mkdir(parents=True, exist_ok=True)

    initialized = False
    try:
        pythoncom.CoInitialize()
        initialized = True
    except pythoncom.com_error as exc:
        raise RuntimeError(f'No se pudo inicializar COM: {exc}') from exc

    excel = None
    wb = None
    generated: List[Path] = []

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
                pdf_name = f"COBRANZA_{_sanitize(name_clean)}.pdf"
                pdf_path = out_dir / pdf_name
                ws.ExportAsFixedFormat(Type=0, Filename=str(pdf_path), Quality=0, IncludeDocProperties=True, IgnorePrintAreas=False, OpenAfterPublish=False)
                generated.append(pdf_path)
                continue

            if target_sheet_name and name != target_sheet_name:
                continue

            if hoja_base is None and name_clean in hojas_completas_set:
                continue

            blocks = _find_vendor_blocks(ws)
            if not blocks:
                continue

            header_rows: Optional[Tuple[int, int]] = None
            first_vendor_row = min(blk["row_start"] for blk in blocks)
            if first_vendor_row > 1:
                header_rows = (1, first_vendor_row - 1)

            for blk in blocks:
                vendor_name_raw = blk["vendor_name"].strip()
                vendor = _sanitize(vendor_name_raw) or "SIN_NOMBRE"
                row_start = blk["row_start"]
                row_end = blk["row_end"]

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
                    pdf_name = f"COBRANZA_{vendor}.pdf"
                    pdf_path = out_dir / pdf_name
                    tmp.ExportAsFixedFormat(Type=0, Filename=str(pdf_path), Quality=0, IncludeDocProperties=True, IgnorePrintAreas=False, OpenAfterPublish=False)
                    generated.append(pdf_path)
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

    excluded_merge_names = {
        "COBRANZA_IMPORTE CUENTA SALDO.pdf",
        "COBRANZA_000020 SURQUILLO_SURCO - (OSCAR).pdf",
        "COBRANZA_000000 OFICINA (VES).pdf",
        # Agrega aquí los nombres exactos que quieres omitir
        # "COBRANZA_000002 PITER HUAYTA.pdf",
    }
    excluded_merge = {name.strip().lower() for name in excluded_merge_names}

    consolidated_name = "COBRANZA_CONSOLIDADO.pdf"
    merge_candidates = [
        pdf for pdf in generated
        if pdf.exists()
        and pdf.name.lower() not in excluded_merge
        and pdf.name.lower() != consolidated_name.lower()
    ]
    merged_path = None
    if merge_candidates:
        consolidated_path = out_dir / consolidated_name
        merged_path = _merge_pdf_files(merge_candidates, consolidated_path)
    if merged_path and merged_path not in generated:
        generated.append(merged_path)

    return generated
