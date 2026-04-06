"""
excel_writer.py
───────────────
Writes payload into the IL.xlsx template (templates/IL.xlsx).

Cover sheet     → fills project metadata into fixed cells.
                  AI14 receives the Instrument List document number
                  from fi_meta["docNumber"], not from the global header.

Instrument List → Section 1 data starting at row 6.
                  Column A   : serial numbers (1, 2, 3 …)
                  Columns C–S: instrument data (fixed mapping per spec).
                  Column I   : intentionally skipped (spacer in template).
                  Columns U–V: velocity & optimised diameter for flowmeters.
                  Sections 2 & 3 reserved for a future update.
"""

import math
import re
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, List, Tuple, Union

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


# ─── Flowmeter keywords ───────────────────────────────────────────────────────
# Any instrument whose name contains one of these strings (case-insensitive)
# will have velocity and optimised diameter calculated and written to U & V.
# Add more keywords here as your naming conventions require.

FLOW_KEYWORDS = [
    "flow transmitter",
    "flowmeter",
    "flow meter",
    "flow element",
    "magnetic flowmeter",
    "electromagnetic flow",
    "vortex flow",
    "turbine flow",
    "ultrasonic flow",
    "coriolis",
    "rotameter",
]

# Standard nominal bore sizes in mm (DN series).
# The velocity optimiser steps through this list rather than jumping
# in arbitrary 50 mm increments, so every result lands on a real pipe size.
STANDARD_DN = [
    15, 20, 25, 32, 40, 50, 65, 80, 100, 125, 150,
    200, 250, 300, 350, 400, 450, 500, 600, 700, 800,
    900, 1000, 1200, 1400, 1600, 1800, 2000,
]

# Velocity band for liquid service (m/s).
V_MIN = 0.2
V_MAX = 5.0

# Cover sheet: starting column index (1-based) for character-by-character
# document number writing in row 44.  Column D = 4.
# *** Verify this against your template before using. ***
DOC_NUMBER_ROW        = 44
DOC_NUMBER_START_COL  = 4   # D


# ─── Cover sheet ──────────────────────────────────────────────────────────────

def _fill_cover(ws, header: Dict[str, str], fi_meta: Dict[str, str]) -> None:
    """
    Write project metadata into fixed cells on the Cover sheet.

    Global project fields come from `header`.
    The Instrument List document number comes from `fi_meta["docNumber"]`
    and is written to AI14.  It is also written character-by-character
    across row DOC_NUMBER_ROW starting at column DOC_NUMBER_START_COL,
    because the template uses individual cells for each character of the
    document number in the cover page layout.
    """
    mapping = {
        "AI6":  header.get("date",        ""),
        "AI7":  header.get("preparedBy",  ""),
        "AI8":  header.get("checkedBy",   ""),
        "AI9":  header.get("approvedBy",  ""),
        "AI10": header.get("revision",    ""),
        "AI11": header.get("projectName", ""),
        "AI12": header.get("client",      ""),
        "AI13": header.get("consultant",  ""),
        "AI14": fi_meta.get("docNumber",  ""),
    }

    for cell_ref, value in mapping.items():
        ws[cell_ref] = value

    # Character-by-character write for the styled document number box.
    doc_number = str(fi_meta.get("docNumber", "") or "")
    for i, ch in enumerate(doc_number):
        ws.cell(row=DOC_NUMBER_ROW, column=DOC_NUMBER_START_COL + i, value=ch)


# ─── Velocity optimiser ───────────────────────────────────────────────────────

def _velocity_at(dia_mm: float, flow_m3h: float) -> float:
    """
    Return fluid velocity (m/s) for a circular pipe.

    Formula:  v = Q / A
              Q = flow_m3h / 3600          [m³/s]
              A = π × (dia_mm / 1000)² / 4 [m²]
    """
    area = math.pi * ((dia_mm / 1000) ** 2) / 4
    return (flow_m3h / 3600) / area


def _calculate_optimized_velocity(
    initial_dia_mm: float,
    flow_m3h: float,
) -> Tuple[float, float]:
    """
    Walk the standard DN series to find the smallest bore whose velocity
    falls within [V_MIN, V_MAX] m/s.

    Strategy
    --------
    1. Calculate velocity at the user-supplied diameter.
    2. If velocity is already in range, return immediately.
    3. If velocity > V_MAX (too fast → pipe too small), step UP through
       the DN list until velocity ≤ V_MAX.
    4. If velocity < V_MIN (too slow → pipe too large), step DOWN through
       the DN list until velocity ≥ V_MIN.
    5. If no standard size brings velocity into range, return the closest
       result with the final diameter so the engineer can see the value
       rather than getting a blank cell.

    Parameters
    ----------
    initial_dia_mm : user-entered line size (mm).
    flow_m3h       : design flow rate (m³/h).

    Returns
    -------
    (velocity_m_per_s, recommended_dia_mm) — both rounded to 3 d.p.
    """
    if flow_m3h <= 0 or initial_dia_mm <= 0:
        return 0.0, initial_dia_mm

    # Snap the initial diameter to the nearest standard DN size.
    # abs(dn - initial_dia_mm) finds the distance from each standard size;
    # min(..., key=...) picks the closest one.
    current_dia = min(STANDARD_DN, key=lambda dn: abs(dn - initial_dia_mm))
    idx         = STANDARD_DN.index(current_dia)

    velocity = _velocity_at(current_dia, flow_m3h)

    if V_MIN <= velocity <= V_MAX:
        # Already in range at the snapped size — return immediately.
        return round(velocity, 3), float(current_dia)

    if velocity > V_MAX:
        # Too fast — step up to larger bores.
        for dn in STANDARD_DN[idx + 1:]:
            v = _velocity_at(dn, flow_m3h)
            if v <= V_MAX:
                return round(v, 3), float(dn)
        # Still too fast even at the largest standard size — return largest.
        largest = STANDARD_DN[-1]
        return round(_velocity_at(largest, flow_m3h), 3), float(largest)

    else:
        # Too slow — step down to smaller bores.
        for dn in reversed(STANDARD_DN[:idx]):
            v = _velocity_at(dn, flow_m3h)
            if v >= V_MIN:
                return round(v, 3), float(dn)
        # Still too slow even at the smallest standard size — return smallest.
        smallest = STANDARD_DN[0]
        return round(_velocity_at(smallest, flow_m3h), 3), float(smallest)


def _extract_first_number(text: str) -> float:
    """
    Extract the first numeric value (integer or decimal) from a string.
    Examples:
        "DN50"      → 50.0
        "12.5 m³/h" → 12.5
        "2 inch"    → 2.0    ← caller must handle unit conversion if needed
    Raises ValueError if no number is found.
    """
    matches = re.findall(r"[-+]?\d*\.\d+|\d+", str(text))
    if not matches:
        raise ValueError(f"No numeric value found in {text!r}")
    return float(matches[0])


# ─── Instrument List sheet — Section 1 ───────────────────────────────────────

FI_COL_MAP = {
    "C": "Tag No",
    "D": "Instrument Name",
    "E": "Service Description",
    "F": "Line Size",
    "G": "Medium",
    "H": "Specification",
    # Column I intentionally skipped — spacer/merged region in template
    "J": "Process Conn",
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
    Write Section 1 (Field Instruments) rows into the Instrument List sheet.

    Per row:
      Column A   : serial number (1-based)
      Columns C–S: instrument data per FI_COL_MAP
      Columns U–V: (flowmeters only) calculated velocity and optimised DN

    Rows 1–5 (template header block) are never touched.
    """
    for i, row_dict in enumerate(rows):
        r      = FI_START_ROW + i
        serial = i + 1

        # ── Column A: Serial number ───────────────────────────────────────
        sn_cell           = ws[f"A{r}"]
        sn_cell.value     = serial
        sn_cell.alignment = CENTER_WRAP
        sn_cell.border    = THIN

        # ── Columns C–S: Instrument data ──────────────────────────────────
        for col_letter, field_key in FI_COL_MAP.items():
            cell           = ws[f"{col_letter}{r}"]
            cell.value     = row_dict.get(field_key, "")
            cell.alignment = CENTER_WRAP if col_letter == "C" else LEFT_WRAP
            cell.border    = THIN

        # ── Columns U–V: Flowmeter velocity calculation ───────────────────
        instr_name = str(row_dict.get("Instrument Name", "")).lower()
        is_flowmeter = any(kw in instr_name for kw in FLOW_KEYWORDS)

        if is_flowmeter:
            try:
                dia_mm   = _extract_first_number(row_dict.get("Line Size",   "0"))
                flow_m3h = _extract_first_number(row_dict.get("Design Flow", "0"))

                v_final, d_final = _calculate_optimized_velocity(dia_mm, flow_m3h)

                ws[f"U{r}"].value = v_final
                ws[f"V{r}"].value = d_final

                for col in ("U", "V"):
                    ws[f"{col}{r}"].alignment = CENTER_WRAP
                    ws[f"{col}{r}"].border    = THIN

            except (ValueError, ZeroDivisionError):
                # Parsing failed — write ERR so the engineer knows to check
                # the Line Size or Design Flow input, rather than seeing
                # a blank and assuming no calculation was attempted.
                for col in ("U", "V"):
                    cell           = ws[f"{col}{r}"]
                    cell.value     = "ERR"
                    cell.alignment = CENTER_WRAP
                    cell.border    = THIN


# ─── Instrument List orchestrator ─────────────────────────────────────────────

def _fill_instrument_list(ws, payload: Dict[str, Any]) -> None:
    """
    Entry point for writing the 'Instrument List' sheet.
    Sections 2 & 3 will be added in a future update.
    """
    _write_fi_fixed(ws, payload.get("field_instruments", []))


# ─── Sections 2 & 3 — reserved for future implementation ─────────────────────


# ─── Template path ────────────────────────────────────────────────────────────

BASE_DIR      = Path(__file__).resolve().parent.parent
TEMPLATE_PATH = BASE_DIR / "templates" / "IL.xlsx"


# ─── Public entry point ───────────────────────────────────────────────────────

def write_workbook(
    payload: Dict[str, Any],
    destination: Union[BytesIO, Path],
) -> None:
    """
    Load IL.xlsx template, fill Cover + Instrument List, save to destination.

    Parameters
    ----------
    payload:
        Dictionary from app.py._build_payload(). Expected keys:
          "header"            – global project fields (projectName, client, …)
          "fi_meta"           – {"docName": …, "docNumber": …} for this workbook
          "field_instruments" – list of row dicts for Section 1
          "electrical"        – list of row dicts for Section 2 (future)
          "mov"               – list of row dicts for Section 3 (future)

    destination:
        BytesIO for in-memory streaming to the browser, or a Path to save
        directly to disk.

    Raises
    ------
    FileNotFoundError if IL.xlsx is not found at TEMPLATE_PATH.
    """
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(
            f"Template not found at {TEMPLATE_PATH}.\n"
            "Place IL.xlsx inside the templates/ folder of the project."
        )

    wb = load_workbook(TEMPLATE_PATH)

    header  = payload.get("header",  {})
    fi_meta = payload.get("fi_meta", {})

    # ── Cover sheet ───────────────────────────────────────────────────────────
    if "Cover" in wb.sheetnames:
        _fill_cover(wb["Cover"], header, fi_meta)

    # ── Instrument List sheet ─────────────────────────────────────────────────
    if "Instrument List" in wb.sheetnames:
        _fill_instrument_list(wb["Instrument List"], payload)
    else:
        ws = wb.create_sheet("Instrument List")
        _fill_instrument_list(ws, payload)

    wb.save(destination)