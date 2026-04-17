"""
excel_writer.py
───────────────
Writes payload into the IL.xlsx and IO.xlsx templates.

─── Instrument List (IL.xlsx) ───────────────────────────────────────────────
Cover sheet     → fills project metadata into fixed cells.
                  AI14 = fi_meta["docNumber"]
                  AI15 = header["location"]

Instrument List → Section 1 data starting at row 6.
                  Column A   : serial numbers (1, 2, 3 …)
                  Column B   : generated instrument code from Instrument Name
                  Columns C–S: instrument data (fixed mapping per spec).
                  Column I   : intentionally skipped (spacer in template).
                  Columns U–V: velocity & optimised diameter for flowmeters.
                  Sections 2 & 3 reserved for a future update.

─── IO List (IO.xlsx) ───────────────────────────────────────────────────────
Cover sheet     → same cell mapping as IL cover.
                  AI14 = io_meta["docNumber"]
                  AI15 = header["location"]

IO List sheet   → Three section blocks starting at row 6.
                  Each block opens with a merged header row (A:N).
                  Column A : serial number (continuous across all sections)
                  Column B : Tag No
                  Column C : Instrument Name
                  Column D : Service Description
                              (Electrical & MOV: "Service Desc - Signal Desc")
                  Column E : Signal Type
                  Column F : Source
                  Column G : Destination
                  Column H : DI  — 1 if Signal == "DI", else blank
                  Column I : DO  — 1 if Signal == "DO", else blank
                  Column J : AI  — 1 if Signal == "AI", else blank
                  Column K : AO  — 1 if Signal == "AO", else blank
                  Column L : skipped (blank, bordered)
                  Column M : TRIP — "TRIP" if "trip" in Service Description
"""

import json
import math
import re
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


# ─── Style helpers ────────────────────────────────────────────────────────────

def _fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)


def _border(left="thin", right="thin", top="thin", bottom="thin") -> Border:
    def s(st):
        return Side(style=st) if st else Side()
    return Border(left=s(left), right=s(right), top=s(top), bottom=s(bottom))


THIN        = _border()
CENTER_WRAP = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT_WRAP   = Alignment(horizontal="left",   vertical="center", wrap_text=True)


# ─── Paths ────────────────────────────────────────────────────────────────────

BASE_DIR         = Path(__file__).resolve().parent.parent
TEMPLATE_PATH    = BASE_DIR / "templates" / "IL.xlsx"
CODE_MAP_PATH    = BASE_DIR / "templates" / "instrument_code_map.json"
IO_TEMPLATE_PATH = BASE_DIR / "templates" / "IO.xlsx"


# ─── Flowmeter keywords ───────────────────────────────────────────────────────

FLOW_KEYWORDS = [
    "flow transmitter",
    "flowmeter",
    "flow meter",
    "flow element",
    "magnetic flowmeter",
    "magnetic flow meter",
    "electromagnetic flow",
    "vortex flow",
    "turbine flow",
    "ultrasonic flow",
    "coriolis",
    "rotameter",
]

STANDARD_DN = [
    15, 20, 25, 32, 40, 50, 65, 80, 100, 125, 150,
    200, 250, 300, 350, 400, 450, 500, 600, 700, 800,
    900, 1000, 1200, 1400, 1600, 1800, 2000, 2500, 3000,
]

V_MIN = 0.2
V_MAX = 5.0

DOC_NUMBER_ROW       = 44
DOC_NUMBER_START_COL = 4   # column D


# ─── Instrument code config — lazy loader ─────────────────────────────────────

_instrument_code_config: Optional[Dict[str, Any]] = None


def _normalize_text(text: str) -> str:
    text = str(text or "").strip().lower()
    text = text.replace("&", " and ")
    text = re.sub(r"[-_/(),]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text


def _load_instrument_code_config(path: Path) -> Dict[str, Any]:
    if not path.exists():
        raise FileNotFoundError(
            f"Instrument code map not found at {path}.\n"
            "Place instrument_code_map.json inside the templates/ folder."
        )
    with path.open("r", encoding="utf-8") as f:
        config = json.load(f)

    if not isinstance(config, dict):
        raise ValueError("instrument_code_map.json must be a JSON object.")

    config.setdefault("exact_map",      {})
    config.setdefault("contains_rules", [])
    config.setdefault("stop_words",     [])

    if not isinstance(config["exact_map"], dict):
        raise ValueError("'exact_map' must be a JSON object.")
    if not isinstance(config["contains_rules"], list):
        raise ValueError("'contains_rules' must be a JSON array.")
    if not isinstance(config["stop_words"], list):
        raise ValueError("'stop_words' must be a JSON array.")

    normalised_exact: Dict[str, str] = {
        _normalize_text(k): str(v)
        for k, v in config["exact_map"].items()
    }

    normalised_rules: List[Tuple[str, str]] = []
    for item in config["contains_rules"]:
        if not isinstance(item, list) or len(item) != 2:
            raise ValueError(
                "Each entry in 'contains_rules' must be a two-item array: "
                "[keyword, code]"
            )
        keyword, code = item
        normalised_rules.append((_normalize_text(keyword), str(code)))

    normalised_stops: set = {_normalize_text(w) for w in config["stop_words"]}

    return {
        "exact_map":      normalised_exact,
        "contains_rules": normalised_rules,
        "stop_words":     normalised_stops,
    }


def _get_code_config() -> Dict[str, Any]:
    global _instrument_code_config
    if _instrument_code_config is None:
        _instrument_code_config = _load_instrument_code_config(CODE_MAP_PATH)
    return _instrument_code_config


# ─── Instrument code generation ───────────────────────────────────────────────

def _fallback_instrument_code(instr_name: str, stop_words: set) -> str:
    words = [
        w for w in _normalize_text(instr_name).split()
        if w and w not in stop_words
    ]
    if not words:
        return ""
    parts = []
    for w in words[:3]:
        if re.search(r"\d", w):
            parts.append(w.upper())
        else:
            parts.append(w[0].upper())
    return "".join(parts)


def _generate_instrument_code(instr_name: str) -> str:
    text = _normalize_text(instr_name)
    if not text:
        return ""

    config         = _get_code_config()
    exact_map      = config["exact_map"]
    contains_rules = config["contains_rules"]
    stop_words     = config["stop_words"]

    if text in exact_map:
        return exact_map[text]

    for keyword, code in contains_rules:
        if keyword and keyword in text:
            return code

    return _fallback_instrument_code(text, stop_words)


# ─── Cover sheet ──────────────────────────────────────────────────────────────

def _fill_cover(ws, header: Dict[str, str], meta: Dict[str, str]) -> None:
    """
    Write project metadata into fixed cells on a Cover sheet.

    AI6–AI13 : global project fields from `header`.
    AI14      : document number from `meta["docNumber"]`.
    AI15      : location from `header["location"]`.

    The document number is also written character-by-character across
    row DOC_NUMBER_ROW starting at column DOC_NUMBER_START_COL (D44).

    Shared by both the IL and IO workbook writers — pass fi_meta for the
    Instrument List, io_meta for the IO List.
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
        "AI14": meta.get("docNumber",     ""),
        "AI15": header.get("location",    ""),
    }

    for cell_ref, value in mapping.items():
        ws[cell_ref] = value

    # Character-by-character write for the styled document number row.
    doc_number = str(meta.get("docNumber", "") or "")
    for i, ch in enumerate(doc_number):
        ws.cell(row=DOC_NUMBER_ROW, column=DOC_NUMBER_START_COL + i, value=ch)


# ─── Velocity optimiser ───────────────────────────────────────────────────────

def _velocity_at(dia_mm: float, flow_m3h: float) -> float:
    area = math.pi * ((dia_mm / 1000) ** 2) / 4
    return (flow_m3h / 3600) / area


def _calculate_optimized_velocity(
    initial_dia_mm: float,
    flow_m3h: float,
) -> Tuple[float, float]:
    if flow_m3h <= 0 or initial_dia_mm <= 0:
        return 0.0, initial_dia_mm

    current_dia = min(STANDARD_DN, key=lambda dn: abs(dn - initial_dia_mm))
    idx         = STANDARD_DN.index(current_dia)
    velocity    = _velocity_at(current_dia, flow_m3h)

    if V_MIN <= velocity <= V_MAX:
        return round(velocity, 3), float(current_dia)

    if velocity > V_MAX:
        for dn in STANDARD_DN[idx + 1:]:
            v = _velocity_at(dn, flow_m3h)
            if v <= V_MAX:
                return round(v, 3), float(dn)
        largest = STANDARD_DN[-1]
        return round(_velocity_at(largest, flow_m3h), 3), float(largest)

    for dn in reversed(STANDARD_DN[:idx]):
        v = _velocity_at(dn, flow_m3h)
        if v >= V_MIN:
            return round(v, 3), float(dn)
    smallest = STANDARD_DN[0]
    return round(_velocity_at(smallest, flow_m3h), 3), float(smallest)


def _extract_first_number(text: str) -> float:
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
    # Column I intentionally skipped — spacer/merged region in template.
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


def _write_fi_fixed(ws, rows: List[Dict[str, Any]]) -> None:
    for i, row_dict in enumerate(rows):
        r      = FI_START_ROW + i
        serial = i + 1

        sn_cell           = ws[f"A{r}"]
        sn_cell.value     = serial
        sn_cell.alignment = CENTER_WRAP
        sn_cell.border    = THIN

        instrument_name     = row_dict.get("Instrument Name", "")
        code_cell           = ws[f"B{r}"]
        code_cell.value     = _generate_instrument_code(instrument_name)
        code_cell.alignment = CENTER_WRAP
        code_cell.border    = THIN

        for col_letter, field_key in FI_COL_MAP.items():
            cell           = ws[f"{col_letter}{r}"]
            cell.value     = row_dict.get(field_key, "")
            cell.alignment = CENTER_WRAP if col_letter == "C" else LEFT_WRAP
            cell.border    = THIN

        instr_lower  = str(instrument_name).lower()
        is_flowmeter = any(kw in instr_lower for kw in FLOW_KEYWORDS)

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
                for col in ("U", "V"):
                    cell           = ws[f"{col}{r}"]
                    cell.value     = "ERR"
                    cell.alignment = CENTER_WRAP
                    cell.border    = THIN


# ─── Instrument List orchestrator ─────────────────────────────────────────────

def _fill_instrument_list(ws, payload: Dict[str, Any]) -> None:
    _write_fi_fixed(ws, payload.get("field_instruments", []))


# ─── IL Public entry point ────────────────────────────────────────────────────

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
          "header"            – global project fields incl. "location"
          "fi_meta"           – {"docNumber": …} for the Instrument List
          "field_instruments" – list of row dicts for Section 1
          "electrical"        – reserved for Section 2 (future)
          "mov"               – reserved for Section 3 (future)
    destination:
        BytesIO for in-memory streaming, or a Path to save to disk.
    """
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(
            f"Template not found at {TEMPLATE_PATH}.\n"
            "Place IL.xlsx inside the templates/ folder of the project."
        )

    _get_code_config()

    wb      = load_workbook(TEMPLATE_PATH)
    header  = payload.get("header",  {})
    fi_meta = payload.get("fi_meta", {})

    if "Cover" in wb.sheetnames:
        _fill_cover(wb["Cover"], header, fi_meta)

    if "Instrument List" in wb.sheetnames:
        _fill_instrument_list(wb["Instrument List"], payload)
    else:
        ws = wb.create_sheet("Instrument List")
        _fill_instrument_list(ws, payload)

    wb.save(destination)


# ═════════════════════════════════════════════════════════════════════════════
# IO List
# ═════════════════════════════════════════════════════════════════════════════

IO_START_ROW = 6

# Block header: accent-dim blue fill (#1e3a6e) with white bold text so the
# section label is clearly readable against the dark background.
_IO_BLOCK_FILL  = _fill("1e3a6e")
_IO_BLOCK_FONT  = Font(bold=True, color="FFFFFF")
_IO_BLOCK_ALIGN = Alignment(horizontal="left", vertical="center",
                             indent=1, wrap_text=False)


def _io_cell(ws, row: int, col: int, value: Any, alignment: Alignment) -> None:
    """Write value + border + alignment into a single IO List data cell."""
    c           = ws.cell(row=row, column=col)
    c.value     = value
    c.alignment = alignment
    c.border    = THIN


def _write_io_block_header(ws, row: int, label: str) -> None:
    """
    Write a merged section-label row spanning columns A (1) through N (14).

    openpyxl requires the border to be applied to every cell in the merged
    range individually — otherwise only the outer edge renders correctly.
    """
    ws.merge_cells(f"A{row}:N{row}")

    hdr           = ws.cell(row=row, column=1)
    hdr.value     = label
    hdr.fill      = _IO_BLOCK_FILL
    hdr.font      = _IO_BLOCK_FONT
    hdr.alignment = _IO_BLOCK_ALIGN
    hdr.border    = THIN

    for col in range(2, 15):   # B(2) … N(14)
        ws.cell(row=row, column=col).border = THIN


def _write_io_list_sheet(ws, payload: Dict[str, Any]) -> None:
    """
    Write all three instrument sections into the 'IO List' sheet.

    Section order:
      1. Field Instruments   (service description as-is)
      2. Electrical Equipment (service + " - " + signal description)
      3. Motor Operated Valves (service + " - " + signal description)

    The TRIP flag in column M is evaluated against the raw service
    description before concatenation, so the keyword match is unambiguous.

    Column layout (1-based):
      1  A  Serial number  (continuous across all sections)
      2  B  Tag No
      3  C  Instrument Name
      4  D  Service Description (El/MOV: "Service Desc - Signal Desc")
      5  E  Signal Type
      6  F  Source
      7  G  Destination
      8  H  DI  — 1 if Signal == "DI", else blank
      9  I  DO  — 1 if Signal == "DO", else blank
     10  J  AI  — 1 if Signal == "AI", else blank
     11  K  AO  — 1 if Signal == "AO", else blank
     12  L  Skipped (blank, bordered)
     13  M  TRIP — "TRIP" if "trip" in Service Description, else blank
    """
    # concat_sig_desc flag: True for Electrical and MOV, False for FI
    sections = [
        ("Field Instruments",     payload.get("field_instruments", []), False),
        ("Electrical Equipment",  payload.get("electrical",        []), True),
        ("Motor Operated Valves", payload.get("mov",               []), True),
    ]

    row    = IO_START_ROW
    serial = 1

    for block_label, rows, concat_sig_desc in sections:
        _write_io_block_header(ws, row, block_label)
        row += 1

        for rd in rows:
            tag      = rd.get("Tag No",              "")
            instr    = rd.get("Instrument Name",     "")
            service  = rd.get("Service Description", "")
            sig_desc = rd.get("Signal Description",  "").strip()
            sig_t    = rd.get("Signal Type",         "")
            source   = rd.get("Source",              "")
            dest     = rd.get("Destination",         "")
            signal   = rd.get("Signal",              "").strip().upper()

            if concat_sig_desc and sig_desc:
                service_display = f"{service} - {sig_desc}" if service else sig_desc
            else:
                service_display = service

            trip = "TRIP" if "trip" in service+" - "+sig_desc.lower() else ""

            _io_cell(ws, row,  1, serial,                          CENTER_WRAP)
            _io_cell(ws, row,  2, tag,                             CENTER_WRAP)
            _io_cell(ws, row,  3, instr,                           LEFT_WRAP)
            _io_cell(ws, row,  4, service_display,                 LEFT_WRAP)
            _io_cell(ws, row,  5, sig_t,                           CENTER_WRAP)
            _io_cell(ws, row,  6, source,                          CENTER_WRAP)
            _io_cell(ws, row,  7, dest,                            CENTER_WRAP)
            _io_cell(ws, row,  8, 1 if signal == "DI" else "",     CENTER_WRAP)
            _io_cell(ws, row,  9, 1 if signal == "DO" else "",     CENTER_WRAP)
            _io_cell(ws, row, 10, 1 if signal == "AI" else "",     CENTER_WRAP)
            _io_cell(ws, row, 11, 1 if signal == "AO" else "",     CENTER_WRAP)
            _io_cell(ws, row, 12, "",                              CENTER_WRAP)
            _io_cell(ws, row, 13, trip,                            CENTER_WRAP)

            serial += 1
            row    += 1


# ─── IO Public entry point ────────────────────────────────────────────────────

def write_io_workbook(
    payload: Dict[str, Any],
    destination: Union[BytesIO, Path],
) -> None:
    """
    Load IO.xlsx template, fill Cover + IO List sheet, save to destination.

    Parameters
    ----------
    payload:
        Dictionary from app.py._build_payload(). Expected keys:
          "header"            – global project fields incl. "location"
          "io_meta"           – {"docNumber": …} for the IO List
          "field_instruments" – list of row dicts (Section 1)
          "electrical"        – list of row dicts (Section 2)
          "mov"               – list of row dicts (Section 3)
    destination:
        BytesIO for in-memory streaming, or a Path to save to disk.
    """
    if not IO_TEMPLATE_PATH.exists():
        raise FileNotFoundError(
            f"IO template not found at {IO_TEMPLATE_PATH}.\n"
            "Place IO.xlsx inside the templates/ folder of the project."
        )

    wb      = load_workbook(IO_TEMPLATE_PATH)
    header  = payload.get("header",  {})
    io_meta = payload.get("io_meta", {})

    if "Cover" in wb.sheetnames:
        _fill_cover(wb["Cover"], header, io_meta)

    if "IO List" in wb.sheetnames:
        _write_io_list_sheet(wb["IO List"], payload)
    else:
        ws = wb.create_sheet("IO List")
        _write_io_list_sheet(ws, payload)

    wb.save(destination)