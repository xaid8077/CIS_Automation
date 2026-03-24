"""
excel_writer.py
───────────────
Converts the flat payload from app.py into a styled .xlsx workbook.

Sheet layout
────────────
  0  Header          — project metadata as a key/value table
  1  Field Instruments — Section 1 (full process + signal data)
  2  Electrical       — Section 2 (Tag/Instrument/Service merged across signal rows)
  3  MOV              — Section 3 (same merge logic as Electrical)

Every data sheet opens with a compact header block showing project info,
then the column-header row, then data rows.

Merge logic (Sections 2 & 3)
─────────────────────────────
Consecutive rows that share the same Tag No are considered one tag group.
Cells in columns "Tag No", "Instrument Name", and "Service Description"
are merged vertically across that group and centred.
"""

from io import BytesIO
from pathlib import Path
from typing import Any, Dict, List, Union

from openpyxl import Workbook
from openpyxl.styles import (
    Alignment, Border, Font, PatternFill, Side
)
from openpyxl.utils import get_column_letter


# ─── Style constants ──────────────────────────────────────────────────────────

# Colours  (hex, no leading #)
C_HEADER_BG   = "1E3A5F"   # dark navy  — project info block bg
C_HEADER_FG   = "FFFFFF"   # white      — project info block text
C_COL_HEAD_BG = "D9EAF7"   # light blue — column header row
C_COL_HEAD_FG = "0F1117"   # near-black
C_ALT_ROW     = "F4F7FB"   # very light — alternating data row tint
C_MERGE_BG    = "EBF3FB"   # slightly deeper tint for merged identity cells
C_ACCENT_LEFT = "4F8FFF"   # blue left border accent on merged cells

def _fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)

def _border(
    left="thin", right="thin", top="thin", bottom="thin"
) -> Border:
    def s(style): return Side(style=style) if style else Side()
    return Border(left=s(left), right=s(right), top=s(top), bottom=s(bottom))

THIN_BORDER   = _border()
MEDIUM_BORDER = _border("medium", "medium", "medium", "medium")

CENTER_WRAP = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT_WRAP   = Alignment(horizontal="left",   vertical="center", wrap_text=True)
LEFT_TOP    = Alignment(horizontal="left",   vertical="top",    wrap_text=True)

# Columns that are merged vertically across a tag group in Sections 2 & 3
MERGE_KEYS = {"Tag No", "Instrument Name", "Service Description"}


# ─── Internal helpers ─────────────────────────────────────────────────────────

def _write_project_block(ws, header: Dict[str, str]) -> int:
    """
    Write a 2-column key/value block at the top of a sheet.
    Returns the row index of the NEXT empty row after the block.
    """
    fields = [
        ("Project",      header.get("projectName",    "")),
        ("Document",     header.get("documentName",   "")),
        ("Doc No",       header.get("documentNumber", "")),
        ("Client",       header.get("client",         "")),
        ("Consultant",   header.get("consultant",     "")),
        ("Date",         header.get("date",           "")),
        ("Prepared by",  header.get("preparedBy",     "")),
        ("Checked by",   header.get("checkedBy",      "")),
        ("Approved by",  header.get("approvedBy",     "")),
        ("Revision",     header.get("revision",       "")),
    ]
    # Only include filled fields
    fields = [(k, v) for k, v in fields if v]

    start_row = ws.max_row if ws.max_row > 1 else 1

    for key, val in fields:
        ws.append([key, val])
        r = ws.max_row
        for col_idx in (1, 2):
            cell = ws.cell(row=r, column=col_idx)
            cell.fill      = _fill(C_HEADER_BG)
            cell.font      = Font(color=C_HEADER_FG,
                                  bold=(col_idx == 1),
                                  size=10)
            cell.alignment = LEFT_WRAP
            cell.border    = THIN_BORDER

    # blank spacer
    ws.append([])
    return ws.max_row + 1


def _write_column_headers(ws, headers: List[str], row: int) -> None:
    for col_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=col_idx, value=h)
        cell.fill      = _fill(C_COL_HEAD_BG)
        cell.font      = Font(bold=True, size=10, color=C_COL_HEAD_FG)
        cell.alignment = CENTER_WRAP
        cell.border    = THIN_BORDER


def _write_data_rows(ws, rows: List[Dict], headers: List[str],
                     start_row: int, alt: bool = True) -> None:
    for i, row_dict in enumerate(rows):
        r = start_row + i
        use_alt = alt and (i % 2 == 1)
        for col_idx, h in enumerate(headers, start=1):
            val  = row_dict.get(h, "")
            cell = ws.cell(row=r, column=col_idx, value=val)
            cell.border    = THIN_BORDER
            cell.alignment = CENTER_WRAP if col_idx == 1 else LEFT_WRAP
            if use_alt:
                cell.fill = _fill(C_ALT_ROW)


def _apply_merged_rows(ws, rows: List[Dict], headers: List[str],
                       data_start_row: int) -> None:
    """
    For consecutive rows sharing the same Tag No, merge the identity columns
    (Tag No, Instrument Name, Service Description) vertically.
    Also applies a left accent border to the Tag No column of each group.
    """
    if not rows:
        return

    tag_col   = headers.index("Tag No") + 1 if "Tag No" in headers else None
    merge_col_indices = [
        (headers.index(k) + 1)
        for k in MERGE_KEYS
        if k in headers
    ]

    # Build groups: list of (start_row, end_row, tag_value)
    groups: List[tuple] = []
    i = 0
    while i < len(rows):
        tag = rows[i].get("Tag No", "").strip()
        j   = i + 1
        while j < len(rows) and rows[j].get("Tag No", "").strip() == tag and tag:
            j += 1
        groups.append((data_start_row + i, data_start_row + j - 1, tag))
        i = j

    for (row_start, row_end, tag) in groups:
        if row_start == row_end:
            # Single row — just style the tag cell
            if tag_col:
                cell = ws.cell(row=row_start, column=tag_col)
                cell.font      = Font(bold=True, size=10)
                cell.alignment = CENTER_WRAP
            continue

        # Multi-row group → merge identity columns
        for col_idx in merge_col_indices:
            ws.merge_cells(
                start_row=row_start, start_column=col_idx,
                end_row=row_end,     end_column=col_idx,
            )
            top_cell = ws.cell(row=row_start, column=col_idx)
            top_cell.alignment = CENTER_WRAP
            top_cell.fill      = _fill(C_MERGE_BG)
            top_cell.border    = THIN_BORDER

            if col_idx == tag_col:
                top_cell.font = Font(bold=True, size=10)

        # Redraw borders around the whole group block for clean appearance
        for r in range(row_start, row_end + 1):
            for col_idx, _ in enumerate(headers, start=1):
                cell = ws.cell(row=r, column=col_idx)
                left   = "medium" if col_idx == 1                    else "thin"
                right  = "medium" if col_idx == len(headers)         else "thin"
                top    = "medium" if r == row_start                  else "thin"
                bottom = "medium" if r == row_end                    else "thin"
                cell.border = _border(left, right, top, bottom)


def _auto_col_widths(ws, min_w: int = 12, max_w: int = 45) -> None:
    for col in ws.columns:
        best = 0
        letter = col[0].column_letter
        for cell in col:
            if cell.value:
                # account for merged cells (value only on top cell)
                try:
                    best = max(best, len(str(cell.value)))
                except Exception:
                    pass
        ws.column_dimensions[letter].width = max(min_w, min(best + 3, max_w))


# ─── Sheet writers ────────────────────────────────────────────────────────────

def _write_header_sheet(ws, header: Dict[str, str]) -> None:
    ws.title = "Header"
    fields = [
        ("Project Name",   header.get("projectName",    "")),
        ("Document Name",  header.get("documentName",   "")),
        ("Document No",    header.get("documentNumber", "")),
        ("Client",         header.get("client",         "")),
        ("Consultant",     header.get("consultant",     "")),
        ("Date",           header.get("date",           "")),
        ("Prepared By",    header.get("preparedBy",     "")),
        ("Checked By",     header.get("checkedBy",      "")),
        ("Approved By",    header.get("approvedBy",     "")),
        ("Revision",       header.get("revision",       "")),
    ]
    # Title row
    ws.append(["Document Header"])
    title_cell = ws.cell(row=1, column=1)
    title_cell.font      = Font(bold=True, size=13, color=C_HEADER_FG)
    title_cell.fill      = _fill(C_HEADER_BG)
    title_cell.alignment = CENTER_WRAP
    ws.merge_cells("A1:B1")
    ws.row_dimensions[1].height = 22

    for key, val in fields:
        ws.append([key, val])
        r = ws.max_row
        key_cell = ws.cell(row=r, column=1)
        val_cell = ws.cell(row=r, column=2)
        key_cell.font      = Font(bold=True, size=10, color=C_HEADER_FG)
        key_cell.fill      = _fill(C_HEADER_BG)
        key_cell.alignment = LEFT_WRAP
        key_cell.border    = THIN_BORDER
        val_cell.font      = Font(size=10)
        val_cell.fill      = _fill("EBF3FB")
        val_cell.alignment = LEFT_WRAP
        val_cell.border    = THIN_BORDER

    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 50


def _write_data_sheet(
    ws,
    title: str,
    rows: List[Dict],
    header: Dict[str, str],
    merge_identity: bool = False,
) -> None:
    ws.title = title

    # 1. Project info block
    _write_project_block(ws, header)

    if not rows:
        ws.append(["No data available for this section."])
        return

    headers = list(rows[0].keys())

    # 2. Column header row
    col_header_row = ws.max_row
    _write_column_headers(ws, headers, col_header_row)
    ws.row_dimensions[col_header_row].height = 28

    # 3. Data rows
    data_start = col_header_row + 1
    _write_data_rows(ws, rows, headers, data_start, alt=not merge_identity)

    # 4. Merge identity columns for sections with grouped tags
    if merge_identity:
        _apply_merged_rows(ws, rows, headers, data_start)

    # 5. Freeze panes below column headers
    ws.freeze_panes = ws.cell(row=data_start, column=1)

    # 6. Auto column widths
    _auto_col_widths(ws)

    # 7. Set a reasonable row height for data rows
    for r in range(data_start, data_start + len(rows)):
        ws.row_dimensions[r].height = 18


# ─── Public entry point ───────────────────────────────────────────────────────

def write_workbook(
    payload: Dict[str, Any],
    destination: Union[BytesIO, Path],
) -> None:
    """
    Build and save the workbook.

    Parameters
    ----------
    payload      : dict with keys 'header', 'field_instruments',
                   'electrical', 'mov'
    destination  : BytesIO (for in-memory Flask response) or Path (file save)
    """
    wb = Workbook()

    header = payload.get("header", {})

    # Sheet 0 — Header metadata
    _write_header_sheet(wb.active, header)

    # Sheet 1 — Field Instruments (no merging — each row is unique)
    _write_data_sheet(
        wb.create_sheet(),
        "Field Instruments",
        payload.get("field_instruments", []),
        header,
        merge_identity=False,
    )

    # Sheet 2 — Electrical (merge Tag/Instrument/Service across signal rows)
    _write_data_sheet(
        wb.create_sheet(),
        "Electrical",
        payload.get("electrical", []),
        header,
        merge_identity=True,
    )

    # Sheet 3 — MOV (same merge logic)
    _write_data_sheet(
        wb.create_sheet(),
        "MOV",
        payload.get("mov", []),
        header,
        merge_identity=True,
    )

    wb.save(destination)
