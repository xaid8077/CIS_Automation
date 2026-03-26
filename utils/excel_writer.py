"""
excel_writer.py
───────────────
Writes payload into the IL.xlsx template (templates/IL.xlsx).

Cover sheet     → fills 9 fixed cells with header metadata.
Instrument List → writes Section 1 data starting at row 6,
                  columns C–S (fixed mapping per spec).
                  Sections 2 & 3 reserved for a future update.
"""

from io import BytesIO
from pathlib import Path
from typing import Any, Dict, List, Union

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, PatternFill, Side

# ─── Style helpers ────────────────────────────────────────────────────────────

def _fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)

def _border(left="thin", right="thin", top="thin", bottom="thin") -> Border:
    def s(st): return Side(style=st) if st else Side()
    return Border(left=s(left), right=s(right), top=s(top), bottom=s(bottom))

THIN        = _border()
CENTER_WRAP = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT_WRAP   = Alignment(horizontal="left",   vertical="center", wrap_text=True)

# ─── Cover sheet ──────────────────────────────────────────────────────────────

def _fill_cover(ws, header: Dict[str, str]) -> None:
    """Write header metadata into fixed Cover sheet cells."""
    mapping = {
        "AI6":  header.get("date",           ""),
        "AI10": header.get("revision",       ""),
        "AI12": header.get("client",         ""),
        "AI13": header.get("consultant",     ""),
        "AI11": header.get("projectName",    ""),
        "AI7":  header.get("preparedBy",     ""),
        "AI8":  header.get("checkedBy",      ""),
        "AI9":  header.get("approvedBy",     ""),
        "AI14": header.get("documentNumber", ""),
    }
    for cell_ref, value in mapping.items():
        ws[cell_ref] = value

    text = ws["AI14"].value or ""
    for i, ch in enumerate(str(text)):
        ws.cell(row=44, column=4+i, value=ch)

# ─── Instrument List — Section 1 (fixed column mapping) ───────────────────────

# Column letter → row-dict key mapping, starting row = 6
FI_COL_MAP = {
    "C": "Tag No",
    "D": "Instrument Name",
    "E": "Service Description",
    "F": "Line Size",
    "G": "Medium",
    "H": "Specification",
    "J": "Process Conn",       # note: I is skipped per spec
    "K": "Work Press",
    "L": "Work Flow",
    "M": "Work Level",
    "N": "Design Press",
    "O": "Design Flow",
    "P": "Design Level",
    "Q": "Set-point",
    "R": "Range",
    "S": "UOM",
}
FI_START_ROW = 6

def _write_fi_fixed(ws, rows: List[Dict]) -> None:
    """
    Write Section 1 rows into the fixed column layout of 'Instrument List'.
    Starts at row FI_START_ROW, increments down for each data row.
    Does NOT write any column headers — those already exist in the template.
    """
    for i, row_dict in enumerate(rows):
        r = FI_START_ROW + i
        for col_letter, field_key in FI_COL_MAP.items():
            cell = ws[f"{col_letter}{r}"]
            cell.value     = row_dict.get(field_key, "")
            cell.alignment = CENTER_WRAP if col_letter == "C" else LEFT_WRAP
            cell.border    = THIN

# ─── Sections 2 & 3 — reserved for future use ────────────────────────────────

def _fill_instrument_list(ws, payload: Dict[str, Any]) -> None:
    """
    Write Section 1 (Field Instruments) into the 'Instrument List' sheet.
    Fixed columns C:S starting at row 6.
    Column headers already exist in the template — not rewritten here.
    Sections 2 & 3 will be added in a future update.
    """
    _write_fi_fixed(ws, payload.get("field_instruments", []))

# ─── Public entry point ───────────────────────────────────────────────────────

BASE_DIR      = Path(__file__).resolve().parent.parent
TEMPLATE_PATH = BASE_DIR / "templates" / "IL.xlsx"

def write_workbook(
    payload: Dict[str, Any],
    destination: Union[BytesIO, Path],
) -> None:
    """
    Load IL.xlsx template, fill Cover + Instrument List, save to destination.
    Raises FileNotFoundError clearly if the template is missing.
    """
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(
            f"Template not found at {TEMPLATE_PATH}. "
            "Ensure IL.xlsx is placed in the templates/ folder."
        )

    wb = load_workbook(TEMPLATE_PATH)

    # ── Cover ─────────────────────────────────────────────────────────────────
    if "Cover" in wb.sheetnames:
        _fill_cover(wb["Cover"], payload.get("header", {}))

    # ── Instrument List ───────────────────────────────────────────────────────
    if "Instrument List" in wb.sheetnames:
        _fill_instrument_list(wb["Instrument List"], payload)
    else:
        ws = wb.create_sheet("Instrument List")
        _fill_instrument_list(ws, payload)

    wb.save(destination)