# utils/excel_writer.py
"""
excel_writer.py
───────────────
Low-level Excel generation via direct XML manipulation.

Both workbooks (IL and IO) are stored as base-64-encoded zip archives in
utils/embedded_templates.py.  This module:

  1. Decodes the template.
  2. Reads and patches the XML inside the zip in-memory.
  3. Returns a BytesIO stream the caller can send directly to the browser.

No openpyxl dependency — all manipulation is done with stdlib xml.etree,
zipfile, and base64.

Public API
──────────
    stream = write_instrument_list(payload)   → BytesIO
    stream = write_io_workbook(payload)        → BytesIO

Template cell conventions (Cover / sheet1)
──────────────────────────────────────────
    AI6  = date (written as Excel date serial so TEXT() formulas work)
    AI7  = prepared by
    AI8  = checked by
    AI9  = approved by
    AI10 = revision
    AI11 = project name
    AI12 = client
    AI13 = consultant
    AI14 = document number
    AI15 = location / project description
"""

import base64
import zipfile
from datetime import datetime, date
from io import BytesIO
from typing import Any, Dict, List, Optional, Tuple
from xml.etree import ElementTree as ET

from utils.embedded_templates import IL_TEMPLATE_B64, IO_TEMPLATE_B64

# ─────────────────────────────────────────────────────────────────────────────
# Namespace constants
# ─────────────────────────────────────────────────────────────────────────────

_NSMAP = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r":    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

# Register so ET.tostring() keeps prefixes intact
for _prefix, _uri in _NSMAP.items():
    ET.register_namespace(_prefix, _uri)
ET.register_namespace("", _NSMAP["main"])   # default namespace

_MAIN_NS = _NSMAP["main"]
_REL_NS  = _NSMAP["r"]


def _ns(tag: str) -> str:
    return f"{{{_MAIN_NS}}}{tag}"


# ─────────────────────────────────────────────────────────────────────────────
# Date helper
# ─────────────────────────────────────────────────────────────────────────────

# Excel epoch: day 1 = 1900-01-01.
# Python's date(1899, 12, 30) acts as day 0 because Excel incorrectly
# treats 1900 as a leap year (Lotus 1-2-3 compatibility bug).
_EXCEL_EPOCH = date(1899, 12, 30)


def _to_excel_date(date_str: str) -> Optional[int]:
    """
    Convert an ISO date string (YYYY-MM-DD) to an Excel date serial.

    Returns None when the string cannot be parsed so callers can fall back
    to writing the raw string.
    """
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            d = datetime.strptime(date_str, fmt).date()
            return (d - _EXCEL_EPOCH).days
        except ValueError:
            continue
    return None


# ─────────────────────────────────────────────────────────────────────────────
# Shared-strings helpers
# ─────────────────────────────────────────────────────────────────────────────

def _load_shared_strings(zipf: zipfile.ZipFile) -> Tuple[Optional[ET.Element], List[str]]:
    """
    Parse xl/sharedStrings.xml and return (root_element, list_of_strings).
    Returns (None, []) when the file does not exist in the archive.
    """
    try:
        root = ET.fromstring(zipf.read("xl/sharedStrings.xml"))
        strings: List[str] = []
        for si in root.findall(_ns("si")):
            # Simple <t> element
            t = si.find(_ns("t"))
            if t is not None:
                strings.append(t.text or "")
                continue
            # Rich-text <r><t> elements — concatenate all runs
            parts = [r.text or "" for r in si.findall(f".//{_ns('t')}")]
            strings.append("".join(parts))
        return root, strings
    except KeyError:
        return None, []


def _get_shared_idx(root: ET.Element, strings: List[str], val: str) -> int:
    """
    Return the shared-string index for *val*, appending it if absent.
    Mutates *root* and *strings* in-place.
    """
    if val in strings:
        return strings.index(val)
    idx = len(strings)
    strings.append(val)
    si = ET.SubElement(root, _ns("si"))
    t  = ET.SubElement(si, _ns("t"))
    t.text = val
    # Keep the count attribute accurate
    root.set("count",       str(int(root.get("count",       "0")) + 1))
    root.set("uniqueCount", str(int(root.get("uniqueCount", "0")) + 1))
    return idx


# ─────────────────────────────────────────────────────────────────────────────
# Sheet-map helper
# ─────────────────────────────────────────────────────────────────────────────

def _map_sheets(zipf: zipfile.ZipFile) -> Dict[str, str]:
    """
    Return {sheet_name: "xl/worksheets/sheetN.xml"} by parsing workbook.xml.
    """
    wb   = ET.fromstring(zipf.read("xl/workbook.xml"))
    rels = ET.fromstring(zipf.read("xl/_rels/workbook.xml.rels"))

    rel_map: Dict[str, str] = {
        r.attrib["Id"]: r.attrib["Target"]
        for r in rels
    }

    mapping: Dict[str, str] = {}
    for s in wb.findall(f".//{_ns('sheet')}"):
        name = s.attrib["name"]
        rid  = s.attrib[f"{{{_REL_NS}}}id"]
        target = rel_map[rid]
        # Target may already start with "worksheets/"
        path = f"xl/{target}" if not target.startswith("xl/") else target
        mapping[name] = path

    return mapping


# ─────────────────────────────────────────────────────────────────────────────
# Cell read / write
# ─────────────────────────────────────────────────────────────────────────────

def _col_letter_to_index(col: str) -> int:
    """Convert column letter(s) to a 1-based integer (A→1, Z→26, AA→27…)."""
    result = 0
    for ch in col.upper():
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result


def _parse_ref(ref: str) -> Tuple[str, int]:
    """Split "AB12" → ("AB", 12)."""
    col = "".join(c for c in ref if c.isalpha())
    row = int("".join(c for c in ref if c.isdigit()))
    return col, row


def _set_cell(
    sheet_root: ET.Element,
    ref: str,
    val: Any,
    shared_root: ET.Element,
    shared_strings: List[str],
) -> None:
    """
    Write *val* to cell *ref* in *sheet_root*.

    Behaviour
    ─────────
    - If the cell already exists: replaces its value and removes any formula.
    - If the cell does not exist: creates it in the correct row, creating the
      row itself if necessary.
    - Strings are stored as shared strings (t="s").
    - Integers / floats are stored as inline numbers.
    - Empty string → cell is cleared (value element removed, t removed).
    """
    col_str, row_num = _parse_ref(ref)
    col_idx = _col_letter_to_index(col_str)

    sheet_data = sheet_root.find(_ns("sheetData"))
    if sheet_data is None:
        sheet_data = ET.SubElement(sheet_root, _ns("sheetData"))

    # ── Find or create the row ────────────────────────────────────────────────
    target_row: Optional[ET.Element] = None
    rows = sheet_data.findall(_ns("row"))
    insert_pos = len(rows)

    for i, r in enumerate(rows):
        r_num = int(r.attrib.get("r", 0))
        if r_num == row_num:
            target_row = r
            break
        if r_num > row_num:
            insert_pos = i
            break

    if target_row is None:
        target_row = ET.Element(_ns("row"), {"r": str(row_num), "spans": f"{col_idx}:{col_idx}"})
        sheet_data.insert(insert_pos, target_row)

    # ── Find or create the cell ───────────────────────────────────────────────
    target_cell: Optional[ET.Element] = None
    cells = target_row.findall(_ns("c"))
    cell_insert_pos = len(cells)

    for i, c in enumerate(cells):
        c_ref = c.attrib.get("r", "")
        c_col = "".join(ch for ch in c_ref if ch.isalpha())
        if c_ref == ref:
            target_cell = c
            break
        if _col_letter_to_index(c_col) > col_idx:
            cell_insert_pos = i
            break

    if target_cell is None:
        target_cell = ET.Element(_ns("c"), {"r": ref})
        target_row.insert(cell_insert_pos, target_cell)

    # ── Remove existing formula and value ─────────────────────────────────────
    for child_tag in (_ns("f"), _ns("v"), _ns("is")):
        for child in target_cell.findall(child_tag):
            target_cell.remove(child)
    target_cell.attrib.pop("t", None)

    # ── Write new value ───────────────────────────────────────────────────────
    if val == "" or val is None:
        # Empty cell — leave it value-less (just the element)
        return

    if isinstance(val, str):
        idx = _get_shared_idx(shared_root, shared_strings, val)
        target_cell.set("t", "s")
        v = ET.SubElement(target_cell, _ns("v"))
        v.text = str(idx)
    elif isinstance(val, bool):
        # bool must come before int (bool subclasses int in Python)
        target_cell.set("t", "b")
        v = ET.SubElement(target_cell, _ns("v"))
        v.text = "1" if val else "0"
    elif isinstance(val, int):
        v = ET.SubElement(target_cell, _ns("v"))
        v.text = str(val)
    elif isinstance(val, float):
        v = ET.SubElement(target_cell, _ns("v"))
        v.text = repr(val)


# ─────────────────────────────────────────────────────────────────────────────
# Merge-cell helper
# ─────────────────────────────────────────────────────────────────────────────

def _apply_merges(sheet_root: ET.Element, merges: List[Tuple[str, str]]) -> None:
    """Append merge-cell entries to the sheet's <mergeCells> element."""
    if not merges:
        return

    merge_cells = sheet_root.find(_ns("mergeCells"))
    if merge_cells is None:
        # Insert after <sheetData> if present
        sheet_data = sheet_root.find(_ns("sheetData"))
        idx = list(sheet_root).index(sheet_data) + 1 if sheet_data is not None else len(sheet_root)
        merge_cells = ET.Element(_ns("mergeCells"))
        sheet_root.insert(idx, merge_cells)

    existing_refs = {mc.attrib.get("ref") for mc in merge_cells.findall(_ns("mergeCell"))}

    for start, end in merges:
        ref = f"{start}:{end}"
        if ref not in existing_refs:
            ET.SubElement(merge_cells, _ns("mergeCell"), {"ref": ref})
            existing_refs.add(ref)

    merge_cells.set("count", str(len(merge_cells.findall(_ns("mergeCell")))))


# ─────────────────────────────────────────────────────────────────────────────
# Cover-sheet header writer (shared by IL and IO)
# ─────────────────────────────────────────────────────────────────────────────

def _build_cover_updates(payload: Dict[str, Any], doc_number: str) -> Dict[str, Any]:
    """
    Return a dict of {cell_ref: value} for the Cover sheet header area.

    All callers pass the full payload; the relevant doc_number key (fi_meta
    or io_meta) is resolved before calling.
    """
    hdr = payload.get("header", {})

    date_str = hdr.get("date", "")
    excel_date = _to_excel_date(date_str)

    return {
        # Input column — formulas in visible cells reference these
        "AI6":  excel_date if excel_date is not None else date_str,
        "AI7":  hdr.get("preparedBy",  ""),
        "AI8":  hdr.get("checkedBy",   ""),
        "AI9":  hdr.get("approvedBy",  ""),
        "AI10": hdr.get("revision",    ""),
        "AI11": hdr.get("projectName", ""),
        "AI12": hdr.get("client",      ""),
        "AI13": hdr.get("consultant",  ""),
        "AI14": doc_number,
        "AI15": hdr.get("location",    ""),
    }


# ─────────────────────────────────────────────────────────────────────────────
# IL data builder
# ─────────────────────────────────────────────────────────────────────────────

# Instrument List column layout (data sheet)
# A  = S.No         B  = Tag No       C  = Instrument Description
# D  = Service Desc E  = Line Size     F  = Medium
# G  = Specification H = Process Conn  I  = Working Pressure
# J  = Working Flow  K = Working Level L  = Design Pressure
# M  = Design Flow   N = Design Level  O  = Set-point
# P  = Range         Q = UOM           R  = Signal Type
# S  = Source        T = Destination   U  = Signal

_IL_DATA_START_ROW = 6   # first data row in the IL data sheet


def _build_il_data_updates(
    payload: Dict[str, Any],
) -> Tuple[Dict[str, Any], List[Tuple[str, str]]]:
    """
    Build cell updates and merge ranges for all three IL sections.

    Returns
    -------
    (updates, merges)
      updates : {cell_ref: value}
      merges  : [(start_ref, end_ref)]
    """
    updates: Dict[str, Any]          = {}
    merges:  List[Tuple[str, str]]  = []
    row    = _IL_DATA_START_ROW
    serial = 1

    # Each section is (label, rows, extra_cols_from_payload)
    sections = [
        (
            "Section 1 — Field Instruments",
            payload.get("field_instruments", []),
            "fi",
        ),
        (
            "Section 2 — Electrical Equipment",
            payload.get("electrical", []),
            "el",
        ),
        (
            "Section 3 — Motor Operated Valves",
            payload.get("mov", []),
            "mov",
        ),
    ]

    for label, rows, section_type in sections:
        if not rows:
            continue

        # Section-header row (merged A→U)
        updates[f"A{row}"] = label
        merges.append((f"A{row}", f"U{row}"))
        row += 1

        run_start = row
        prev_tag  = None

        for rd in rows:
            tag   = rd.get("Tag No",              "")
            instr = rd.get("Instrument Name",     "")
            svc   = rd.get("Service Description", "")

            updates[f"A{row}"] = serial
            updates[f"B{row}"] = tag
            updates[f"C{row}"] = instr
            updates[f"D{row}"] = svc

            if section_type == "fi":
                updates[f"E{row}"] = rd.get("Line Size",       "")
                updates[f"F{row}"] = rd.get("Medium",          "")
                updates[f"G{row}"] = rd.get("Specification",   "")
                updates[f"H{row}"] = rd.get("Process Conn",    "")
                updates[f"I{row}"] = rd.get("Work Press",      "")
                updates[f"J{row}"] = rd.get("Work Flow",       "")
                updates[f"K{row}"] = rd.get("Work Level",      "")
                updates[f"L{row}"] = rd.get("Design Press",    "")
                updates[f"M{row}"] = rd.get("Design Flow",     "")
                updates[f"N{row}"] = rd.get("Design Level",    "")
                updates[f"O{row}"] = rd.get("Set-point",       "")
                updates[f"P{row}"] = rd.get("Range",           "")
                updates[f"Q{row}"] = rd.get("UOM",             "")
                updates[f"R{row}"] = rd.get("Signal Type",     "")
                updates[f"S{row}"] = rd.get("Source",          "")
                updates[f"T{row}"] = rd.get("Destination",     "")
                updates[f"U{row}"] = rd.get("Signal",          "")
            else:
                # Electrical / MOV — fewer process columns
                sig_desc = rd.get("Signal Description", "")
                if sig_desc:
                    svc_full = f"{svc} — {sig_desc}" if svc else sig_desc
                    updates[f"D{row}"] = svc_full
                updates[f"R{row}"] = rd.get("Signal Type",  "")
                updates[f"S{row}"] = rd.get("Source",       "")
                updates[f"T{row}"] = rd.get("Destination",  "")
                updates[f"U{row}"] = rd.get("Signal",       "")

            # Tag-group merge tracking
            if tag != prev_tag and prev_tag is not None:
                if row - run_start > 1:
                    # Merge tag-stable columns for the completed run
                    for col in ("B", "C"):
                        merges.append((f"{col}{run_start}", f"{col}{row - 1}"))
                run_start = row

            prev_tag = tag or None
            row     += 1
            serial  += 1

        # Flush final run
        if row - run_start > 1:
            for col in ("B", "C"):
                merges.append((f"{col}{run_start}", f"{col}{row - 1}"))

    return updates, merges


# ─────────────────────────────────────────────────────────────────────────────
# IO data builder
# ─────────────────────────────────────────────────────────────────────────────

# IO List column layout (sheet2)
# A = S.No   B = Tag No   C = Instrument Desc   D = Service Desc
# E = Signal Type   F = Source   G = Destination
# H = DI   I = DO   J = AI   K = AO
# L = (spare)   M = Trip/Alarm flag   N = Remarks

_IO_DATA_START_ROW = 6


def _build_io_data_updates(
    payload: Dict[str, Any],
) -> Tuple[Dict[str, Any], List[Tuple[str, str]]]:
    """
    Build cell updates and merge ranges for the IO List data sheet.
    """
    updates: Dict[str, Any]         = {}
    merges: List[Tuple[str, str]]   = []
    row    = _IO_DATA_START_ROW
    serial = 1

    sections = [
        ("Section 1 — Field Instruments",       payload.get("field_instruments", []), False),
        ("Section 2 — Electrical Equipment",     payload.get("electrical",        []), True),
        ("Section 3 — Motor Operated Valves",    payload.get("mov",               []), True),
    ]

    for label, rows, concat_sig_desc in sections:
        if not rows:
            continue

        # Section header merged across all columns
        updates[f"A{row}"] = label
        merges.append((f"A{row}", f"N{row}"))
        row += 1

        run_start = row
        prev_tag  = None

        for rd in rows:
            tag      = rd.get("Tag No",              "")
            instr    = rd.get("Instrument Name",     "")
            svc      = rd.get("Service Description", "")
            sig_desc = rd.get("Signal Description",  "")
            sig_type = rd.get("Signal Type",         "")
            source   = rd.get("Source",              "")
            dest     = rd.get("Destination",         "")
            signal   = (rd.get("Signal") or "").strip().upper()

            if concat_sig_desc and sig_desc:
                svc = f"{svc} — {sig_desc}" if svc else sig_desc

            updates[f"A{row}"] = serial
            updates[f"B{row}"] = tag
            updates[f"C{row}"] = instr
            updates[f"D{row}"] = svc
            updates[f"E{row}"] = sig_type
            updates[f"F{row}"] = source
            updates[f"G{row}"] = dest

            updates[f"H{row}"] = 1 if signal == "DI" else ""
            updates[f"I{row}"] = 1 if signal == "DO" else ""
            updates[f"J{row}"] = 1 if signal == "AI" else ""
            updates[f"K{row}"] = 1 if signal == "AO" else ""

            updates[f"M{row}"] = "TRIP" if "trip" in svc.lower() else ""

            # Tag-group merge tracking
            if tag != prev_tag and prev_tag is not None:
                if row - run_start > 1:
                    for col in ("B", "C", "F", "G"):
                        merges.append((f"{col}{run_start}", f"{col}{row - 1}"))
                run_start = row

            prev_tag = tag or None
            row     += 1
            serial  += 1

        # Flush final run
        if row - run_start > 1:
            for col in ("B", "C", "F", "G"):
                merges.append((f"{col}{run_start}", f"{col}{row - 1}"))

    return updates, merges


# ─────────────────────────────────────────────────────────────────────────────
# Core processing engine
# ─────────────────────────────────────────────────────────────────────────────

def _process(
    template_b64: str,
    updates_by_sheet: Dict[str, Dict[str, Any]],
    merges_by_sheet:  Dict[str, List[Tuple[str, str]]],
) -> BytesIO:
    """
    Decode the base-64 template, apply all cell updates and merges, and
    return the finished workbook as a BytesIO stream.

    Parameters
    ----------
    template_b64      : base-64-encoded bytes of the xlsx template.
    updates_by_sheet  : {"xl/worksheets/sheetN.xml": {"A1": value, …}}
    merges_by_sheet   : {"xl/worksheets/sheetN.xml": [("A1","B2"), …]}
    """
    template_bytes = base64.b64decode(template_b64)

    zin         = zipfile.ZipFile(BytesIO(template_bytes), "r")
    zout_buffer = BytesIO()
    zout        = zipfile.ZipFile(zout_buffer, "w", zipfile.ZIP_DEFLATED)

    # Parse shared strings once — all sheet writers share the same pool
    shared_root, shared_strings = _load_shared_strings(zin)

    for item in zin.infolist():
        data = zin.read(item.filename)

        if item.filename in updates_by_sheet:
            # Register namespace so it survives round-trip serialisation
            root = ET.fromstring(data)

            sheet_updates = updates_by_sheet[item.filename]
            sheet_merges  = merges_by_sheet.get(item.filename, [])

            for ref, val in sheet_updates.items():
                _set_cell(root, ref, val, shared_root, shared_strings)

            if sheet_merges:
                _apply_merges(root, sheet_merges)

            data = ET.tostring(root, encoding="unicode", xml_declaration=False).encode("utf-8")

        elif item.filename == "xl/sharedStrings.xml" and shared_root is not None:
            # Rewrite with accumulated strings
            data = ET.tostring(shared_root, encoding="unicode", xml_declaration=False).encode("utf-8")

        zout.writestr(item, data)

    zin.close()
    zout.close()
    zout_buffer.seek(0)
    return zout_buffer


# ─────────────────────────────────────────────────────────────────────────────
# Public writers
# ─────────────────────────────────────────────────────────────────────────────

def write_instrument_list(payload: Dict[str, Any]) -> BytesIO:
    """
    Generate an Instrument List workbook from *payload*.

    Writes header data to the Cover sheet (sheet1.xml / "Cover") and
    instrument-list data to the IL data sheet (sheet2.xml).

    Returns a BytesIO ready for send_file().
    """
    zin_tmp = zipfile.ZipFile(BytesIO(base64.b64decode(IL_TEMPLATE_B64)), "r")
    sheet_map = _map_sheets(zin_tmp)
    zin_tmp.close()

    cover_path = sheet_map.get("Cover",           "xl/worksheets/sheet1.xml")
    data_path  = sheet_map.get("Instrument List", "xl/worksheets/sheet2.xml")

    doc_number = (payload.get("fi_meta") or {}).get("docNumber", "")

    cover_updates              = _build_cover_updates(payload, doc_number)
    il_updates, il_merges      = _build_il_data_updates(payload)

    return _process(
        IL_TEMPLATE_B64,
        updates_by_sheet={
            cover_path: cover_updates,
            data_path:  il_updates,
        },
        merges_by_sheet={
            data_path: il_merges,
        },
    )


def write_io_workbook(payload: Dict[str, Any]) -> BytesIO:
    """
    Generate an IO List workbook from *payload*.

    Writes header data to the Cover sheet and IO data to the IO List sheet.

    Returns a BytesIO ready for send_file().
    """
    zin_tmp = zipfile.ZipFile(BytesIO(base64.b64decode(IO_TEMPLATE_B64)), "r")
    sheet_map = _map_sheets(zin_tmp)
    zin_tmp.close()

    cover_path = sheet_map.get("Cover",   "xl/worksheets/sheet1.xml")
    data_path  = sheet_map.get("IO List", "xl/worksheets/sheet2.xml")

    doc_number = (payload.get("io_meta") or {}).get("docNumber", "")

    cover_updates              = _build_cover_updates(payload, doc_number)
    io_updates, io_merges      = _build_io_data_updates(payload)

    return _process(
        IO_TEMPLATE_B64,
        updates_by_sheet={
            cover_path: cover_updates,
            data_path:  io_updates,
        },
        merges_by_sheet={
            data_path: io_merges,
        },
    )