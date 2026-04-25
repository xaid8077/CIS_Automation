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

IL column layout (data sheet)
──────────────────────────────
    A  = S.No
    B  = Instrument Code  (auto-generated from Instrument Name)
    C  = Tag No
    D  = Instrument Name
    E  = Service Description
    F  = Line Size
    G  = Medium
    H  = Specification
    I  = Process Conn
    J  = Working Pressure
    K  = Working Flow
    L  = Working Level
    M  = Design Pressure
    N  = Design Flow
    O  = Design Level
    P  = Set-point
    Q  = Range
    R  = UOM
    S  = Signal Type
    T  = Velocity          (flowmeters only — FM instruments)
    U  = FM Size / NB      (flowmeters only)

Cover / sheet1 input cells
──────────────────────────────────────────
    AI6  = date (Excel date serial so TEXT() formulas work)
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
import functools
import json
import math
import os
import re
import zipfile
from copy import deepcopy
from datetime import datetime, date
from io import BytesIO
from typing import Any, Dict, List, Optional, Tuple
from xml.etree import ElementTree as ET

# Change the import line near the top:
from utils.embedded_templates import IL_TEMPLATE_B64, IO_TEMPLATE_B64, CS_TEMPLATE_B64

# ─────────────────────────────────────────────────────────────────────────────
# Namespace constants
# ─────────────────────────────────────────────────────────────────────────────

# ─────────────────────────────────────────────────────────────────────────────
# Namespace constants  — ALL namespaces found in xlsx XML must be registered
# so ElementTree preserves their prefixes on serialisation.
# Missing registrations cause ns0:/ns1: mangling which corrupts the file.
# ─────────────────────────────────────────────────────────────────────────────

_ALL_NS = {
    "":      "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r":     "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc":    "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "x14ac": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
    "xr":    "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
    "xr2":   "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2",
    "xr3":   "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3",
    "x14":   "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
    "x15":   "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main",
    "x15ac": "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac",
    "xr6":   "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6",
    "xr10":  "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10",
    "a":     "http://schemas.openxmlformats.org/drawingml/2006/main",
    "xdr":   "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a14":   "http://schemas.microsoft.com/office/drawing/2010/main",
    "a16":   "http://schemas.microsoft.com/office/drawing/2014/main",
    "cp":    "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc":    "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
    "vt":    "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
    "xcalcf": "http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures",
}

for _prefix, _uri in _ALL_NS.items():
    ET.register_namespace(_prefix, _uri)

_MAIN_NS = _ALL_NS[""]
_REL_NS  = _ALL_NS["r"]
_MC_NS   = _ALL_NS["mc"]
_PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_CALC_CHAIN_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain"
)
_URI_TO_PREFIX = {
    uri: prefix
    for prefix, uri in _ALL_NS.items()
    if prefix
}

# Every XML file inside an xlsx must start with this declaration.
# ET.tostring strips it — we prepend it manually after serialisation.
_XML_DECL = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
_WRITTEN_CELL_FONT_NAME = "Aptos"
_WRITTEN_CELL_FONT_SIZE = "18"
_WRITTEN_ROW_HEIGHT = "35"


def _ns(tag: str) -> str:
    return f"{{{_MAIN_NS}}}{tag}"


def _pkg_rel_ns(tag: str) -> str:
    return f"{{{_PKG_REL_NS}}}{tag}"


def _content_types_ns(tag: str) -> str:
    return f"{{{_CONTENT_TYPES_NS}}}{tag}"

# ─────────────────────────────────────────────────────────────────────────────
# Instrument code map loader
# ─────────────────────────────────────────────────────────────────────────────

@functools.lru_cache(maxsize=1)
def _load_code_map() -> dict:
    """
    Load instrument_code_map.json once and cache it for the process lifetime.

    Searches in templates/ relative to this file's location, then relative
    to the current working directory.  Returns an empty dict on any error
    so the writer degrades gracefully (empty code column) rather than
    crashing document generation.
    """
    candidates = [
        os.path.join(os.path.dirname(__file__), "..", "templates", "instrument_code_map.json"),
        os.path.join("templates", "instrument_code_map.json"),
    ]
    for path in candidates:
        try:
            with open(os.path.normpath(path), encoding="utf-8") as fh:
                return json.load(fh)
        except (FileNotFoundError, json.JSONDecodeError):
            continue
    return {}


def _get_instrument_code(instrument_name: str) -> str:
    """
    Derive the instrument code (e.g. "FM", "PT", "LT") from a free-text
    instrument name string.

    Lookup order
    ─────────────
    1. Exact match on the full lowercased name.
    2. Exact match after removing stop-words.
    3. First matching contains-rule (order-sensitive — longer/more specific
       patterns must appear first in the JSON, which they do).

    Returns an empty string when no match is found.
    """
    if not instrument_name or not instrument_name.strip():
        return ""

    code_map   = _load_code_map()
    exact_map  = code_map.get("exact_map",     {})
    stop_words = set(code_map.get("stop_words", []))
    rules      = code_map.get("contains_rules", [])

    name_lower = instrument_name.lower().strip()

    # 1. Exact match
    if name_lower in exact_map:
        return exact_map[name_lower]

    # 2. Exact match after stop-word removal
    filtered = " ".join(w for w in name_lower.split() if w not in stop_words)
    if filtered and filtered in exact_map:
        return exact_map[filtered]

    # 3. Contains rules (first match wins)
    for pattern, code in rules:
        if pattern in name_lower:
            return code

    return ""


# ─────────────────────────────────────────────────────────────────────────────
# Flowmeter velocity / NB helpers
# ─────────────────────────────────────────────────────────────────────────────

# Common nominal-bore sizes in mm (DN series) used for snapping.
_NB_SERIES = [
    15, 20, 25, 32, 40, 50, 65, 80, 100, 125, 150, 200, 250,
    300, 350, 400, 450, 500, 600, 700, 800, 900, 1000, 1200,
]


def _snap_to_nb(diameter_mm: float) -> int:
    """Return the closest standard NB size for a computed pipe diameter."""
    return min(_NB_SERIES, key=lambda nb: abs(nb - diameter_mm))


def _parse_number(text: str) -> Optional[float]:
    """
    Extract the first numeric value from a free-text string.
    Returns None when no number can be found.
    Examples: "100 m³/h" → 100.0 ; "DN150" → 150.0 ; "0.5 bar" → 0.5
    """
    m = re.search(r"[-+]?\d+(?:\.\d+)?", str(text))
    return float(m.group()) if m else None


def _compute_fm_velocity_and_nb(
    work_flow: str,
    line_size: str,
) -> Tuple[str, str]:
    """
    Compute the flow velocity and snap to the nearest nominal bore (NB) for a
    flowmeter row.

    Parameters
    ----------
    work_flow : str   Working flow value (e.g. "100 m³/h", "27.8 l/s")
    line_size : str   Pipe size string  (e.g. "150 mm", "6 inch", "DN200")

    Returns
    -------
    (velocity_str, nb_str)
        velocity_str — formatted as "X.XX m/s" or "" on failure
        nb_str       — formatted as "DN XXX" or "" on failure

    Notes
    ─────
    - Flow is assumed to be in m³/h if the unit string contains "m³" or "m3"
      or no recognisable unit; in l/s if "l/s" or "lps"; in l/min if "l/min".
    - Pipe size is assumed mm unless "inch" or '"' is present (converted
      to mm: 1 inch = 25.4 mm).
    - Velocity = Q / A  where A = π/4 × D²  (D in metres, Q in m³/s).
    """
    q_raw = _parse_number(work_flow)
    d_raw = _parse_number(line_size)

    if q_raw is None or d_raw is None or q_raw <= 0 or d_raw <= 0:
        return "", ""

    # ── Unit normalisation: flow → m³/s ──────────────────────────────────────
    wf_lower = work_flow.lower()
    if "l/s" in wf_lower or "lps" in wf_lower:
        q_m3s = q_raw / 1000.0
    elif "l/min" in wf_lower or "lpm" in wf_lower:
        q_m3s = q_raw / 60_000.0
    elif "l/h" in wf_lower or "lph" in wf_lower:
        q_m3s = q_raw / 3_600_000.0
    else:
        # Default: m³/h
        q_m3s = q_raw / 3600.0

    # ── Unit normalisation: diameter → mm ────────────────────────────────────
    ls_lower = line_size.lower()
    if "inch" in ls_lower or '"' in ls_lower:
        d_mm = d_raw * 25.4
    else:
        d_mm = d_raw          # assume mm

    nb = _snap_to_nb(d_mm)
    d_m = d_mm / 1000.0

    area = math.pi / 4 * d_m ** 2
    if area <= 0:
        return "", ""

    velocity = q_m3s / area
    return f"{velocity:.2f} m/s", f"DN {nb}"

# ─────────────────────────────────────────────────────────────────────────────
# Date helper
# ─────────────────────────────────────────────────────────────────────────────

_EXCEL_EPOCH = date(1899, 12, 30)


def _to_excel_date(date_str: str) -> Optional[int]:
    """
    Convert an ISO date string (YYYY-MM-DD) to an Excel date serial.
    Returns None when the string cannot be parsed.
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
            t = si.find(_ns("t"))
            if t is not None:
                strings.append(t.text or "")
                continue
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
    root.set("count",       str(int(root.get("count",       "0")) + 1))
    root.set("uniqueCount", str(int(root.get("uniqueCount", "0")) + 1))
    return idx


def _load_styles(zipf: zipfile.ZipFile) -> Optional[ET.Element]:
    """Parse xl/styles.xml and return its root element when present."""
    try:
        return ET.fromstring(zipf.read("xl/styles.xml"))
    except KeyError:
        return None


def _ensure_written_cell_font(styles_root: ET.Element) -> int:
    """Return the fontId for the workbook's Aptos 18 normal font."""
    fonts = styles_root.find(_ns("fonts"))
    if fonts is None:
        fonts = ET.SubElement(styles_root, _ns("fonts"), {"count": "0"})

    for idx, font in enumerate(fonts.findall(_ns("font"))):
        name = font.find(_ns("name"))
        size = font.find(_ns("sz"))
        if (
            name is not None
            and size is not None
            and name.attrib.get("val") == _WRITTEN_CELL_FONT_NAME
            and size.attrib.get("val") == _WRITTEN_CELL_FONT_SIZE
            and font.find(_ns("b")) is None
            and font.find(_ns("i")) is None
            and font.find(_ns("u")) is None
        ):
            return idx

    font = ET.SubElement(fonts, _ns("font"))
    ET.SubElement(font, _ns("sz"), {"val": _WRITTEN_CELL_FONT_SIZE})
    ET.SubElement(font, _ns("name"), {"val": _WRITTEN_CELL_FONT_NAME})
    ET.SubElement(font, _ns("family"), {"val": "2"})
    fonts.set("count", str(len(fonts.findall(_ns("font")))))
    return len(fonts.findall(_ns("font"))) - 1


def _ensure_written_cell_border(styles_root: ET.Element) -> int:
    """Return the borderId for a thin border on all four cell sides."""
    borders = styles_root.find(_ns("borders"))
    if borders is None:
        borders = ET.SubElement(styles_root, _ns("borders"), {"count": "0"})

    for idx, border in enumerate(borders.findall(_ns("border"))):
        left = border.find(_ns("left"))
        right = border.find(_ns("right"))
        top = border.find(_ns("top"))
        bottom = border.find(_ns("bottom"))
        if not all((left, right, top, bottom)):
            continue
        if (
            left.attrib.get("style") == "thin"
            and right.attrib.get("style") == "thin"
            and top.attrib.get("style") == "thin"
            and bottom.attrib.get("style") == "thin"
        ):
            return idx

    border = ET.SubElement(borders, _ns("border"))
    for side in ("left", "right", "top", "bottom"):
        side_elem = ET.SubElement(border, _ns(side), {"style": "thin"})
        ET.SubElement(side_elem, _ns("color"), {"indexed": "64"})
    ET.SubElement(border, _ns("diagonal"))
    borders.set("count", str(len(borders.findall(_ns("border")))))
    return len(borders.findall(_ns("border"))) - 1


def _get_cell_style_index(cell: ET.Element, row: ET.Element) -> int:
    """Return the base style index inherited by a cell."""
    cell_style = cell.attrib.get("s")
    if cell_style is not None:
        try:
            return int(cell_style)
        except ValueError:
            pass

    row_style = row.attrib.get("s")
    if row_style is not None and row.attrib.get("customFormat") == "1":
        try:
            return int(row_style)
        except ValueError:
            pass

    return 0


def _get_written_cell_style_index(
    styles_root: ET.Element,
    base_style_idx: int,
    written_font_id: int,
    written_border_id: int,
    style_cache: Dict[int, int],
) -> int:
    """
    Clone the base cell style and swap only its font to the Aptos 18 font.
    """
    cached_idx = style_cache.get(base_style_idx)
    if cached_idx is not None:
        return cached_idx

    cell_xfs = styles_root.find(_ns("cellXfs"))
    if cell_xfs is None:
        cell_xfs = ET.SubElement(styles_root, _ns("cellXfs"), {"count": "0"})

    xfs = cell_xfs.findall(_ns("xf"))
    if not xfs:
        xfs = [ET.SubElement(cell_xfs, _ns("xf"), {
            "numFmtId": "0",
            "fontId": "0",
            "fillId": "0",
            "borderId": "0",
            "xfId": "0",
        })]

    if base_style_idx < 0 or base_style_idx >= len(xfs):
        base_style_idx = 0

    new_xf = deepcopy(xfs[base_style_idx])
    new_xf.set("fontId", str(written_font_id))
    new_xf.set("borderId", str(written_border_id))
    new_xf.set("applyFont", "1")
    new_xf.set("applyBorder", "1")

    alignment = new_xf.find(_ns("alignment"))
    if alignment is None:
        alignment = ET.SubElement(new_xf, _ns("alignment"))
    alignment.set("horizontal", "center")
    alignment.set("vertical", "center")
    new_xf.set("applyAlignment", "1")

    cell_xfs.append(new_xf)
    cell_xfs.set("count", str(len(cell_xfs.findall(_ns("xf")))))

    new_idx = len(cell_xfs.findall(_ns("xf"))) - 1
    style_cache[base_style_idx] = new_idx
    return new_idx


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
        name   = s.attrib["name"]
        rid    = s.attrib[f"{{{_REL_NS}}}id"]
        target = rel_map[rid]
        path   = f"xl/{target}" if not target.startswith("xl/") else target
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


def _col_index_to_letter(idx: int) -> str:
    """Convert a 1-based integer to Excel column letters."""
    result = []
    while idx > 0:
        idx, remainder = divmod(idx - 1, 26)
        result.append(chr(ord("A") + remainder))
    return "".join(reversed(result))


def _sanitize_excel_text(value: str) -> str:
    """
    Remove characters that are invalid in XML 1.0 and can make Excel repair
    the workbook on open.
    """
    if not value:
        return ""

    cleaned: List[str] = []
    for ch in value:
        code = ord(ch)
        if (
            code in (0x09, 0x0A, 0x0D)
            or 0x20 <= code <= 0xD7FF
            or 0xE000 <= code <= 0xFFFD
            or 0x10000 <= code <= 0x10FFFF
        ):
            cleaned.append(ch)
    return "".join(cleaned)


def _row_text(row: Dict[str, Any], *keys: str) -> str:
    """Return the first non-empty row value among the provided keys."""
    for key in keys:
        if key not in row:
            continue
        value = row.get(key)
        if value is None:
            continue
        text = _sanitize_excel_text(str(value)).strip()
        if text:
            return text
    return ""


def _is_il_section_1_row(row: Dict[str, Any]) -> bool:
    """
    Restrict the IL sheet to explicit Section 1 rows when section metadata is
    present. If no section marker exists, keep the row.
    """
    markers = (
        "Section",
        "Section Name",
        "Section Type",
        "section",
        "section_name",
        "section_type",
        "Category",
        "category",
    )
    section_value = _row_text(row, *markers).lower()
    if not section_value:
        return True
    normalized = section_value.replace("_", " ").replace("-", " ")
    if normalized in {"1", "section 1", "field instrument", "field instruments"}:
        return True
    return "section 1" in normalized or "field instrument" in normalized


def _get_existing_cell(sheet_root: ET.Element, ref: str) -> Optional[ET.Element]:
    """Return the existing cell for *ref* without creating rows/cells."""
    _, row_num = _parse_ref(ref)
    sheet_data = sheet_root.find(_ns("sheetData"))
    if sheet_data is None:
        return None

    for row in sheet_data.findall(_ns("row")):
        if int(row.attrib.get("r", 0)) != row_num:
            continue
        for cell in row.findall(_ns("c")):
            if cell.attrib.get("r") == ref:
                return cell
        return None
    return None


def _clear_existing_cell_value(sheet_root: ET.Element, ref: str) -> None:
    """Clear only the value/formula payload of an existing cell."""
    target_cell = _get_existing_cell(sheet_root, ref)
    if target_cell is None:
        return

    for child_tag in (_ns("f"), _ns("v"), _ns("is")):
        for child in target_cell.findall(child_tag):
            target_cell.remove(child)
    target_cell.attrib.pop("t", None)


def _iter_merge_covered_refs(start_ref: str, end_ref: str) -> List[str]:
    """Return all covered cell refs except the top-left merge anchor."""
    start_col, start_row = _parse_ref(start_ref)
    end_col, end_row = _parse_ref(end_ref)
    start_idx = _col_letter_to_index(start_col)
    end_idx = _col_letter_to_index(end_col)

    refs: List[str] = []
    for row_num in range(start_row, end_row + 1):
        for col_idx in range(start_idx, end_idx + 1):
            ref = f"{_col_index_to_letter(col_idx)}{row_num}"
            if ref != start_ref:
                refs.append(ref)
    return refs


def _collect_used_namespace_prefixes(root: ET.Element) -> set[str]:
    """Collect namespace prefixes actually used in element/attribute names."""
    used: set[str] = set()

    for elem in root.iter():
        if elem.tag.startswith("{"):
            uri = elem.tag[1:].split("}", 1)[0]
            prefix = _URI_TO_PREFIX.get(uri)
            if prefix:
                used.add(prefix)

        for attr_name in elem.attrib:
            if not attr_name.startswith("{"):
                continue
            uri = attr_name[1:].split("}", 1)[0]
            prefix = _URI_TO_PREFIX.get(uri)
            if prefix:
                used.add(prefix)

    return used


def _normalize_ignorable_prefixes(root: ET.Element) -> None:
    """
    ElementTree drops unused namespace declarations on serialisation. Remove
    any now-undeclared prefixes from mc:Ignorable so Excel does not reject the
    worksheet/workbook root element.
    """
    ignorable_attr = f"{{{_MC_NS}}}Ignorable"
    ignorable = root.attrib.get(ignorable_attr)
    if not ignorable:
        return

    used_prefixes = _collect_used_namespace_prefixes(root)
    kept_prefixes = [prefix for prefix in ignorable.split() if prefix in used_prefixes]

    if kept_prefixes:
        root.set(ignorable_attr, " ".join(kept_prefixes))
    else:
        root.attrib.pop(ignorable_attr, None)


def _serialize_xml(root: ET.Element) -> bytes:
    """Serialise an XML root with workbook-safe namespace cleanup."""
    _normalize_ignorable_prefixes(root)
    body = ET.tostring(root, encoding="unicode", xml_declaration=False)
    return _XML_DECL + body.encode("utf-8")


def _remove_calc_chain_content_type(data: bytes) -> bytes:
    """Remove the calcChain content-type override from [Content_Types].xml."""
    root = ET.fromstring(data)

    for override in list(root.findall(_content_types_ns("Override"))):
        if override.attrib.get("PartName") == "/xl/calcChain.xml":
            root.remove(override)

    return _serialize_xml(root)


def _remove_calc_chain_relationship(data: bytes) -> bytes:
    """Remove the workbook relationship pointing to calcChain.xml."""
    root = ET.fromstring(data)

    for rel in list(root.findall(_pkg_rel_ns("Relationship"))):
        if (
            rel.attrib.get("Type") == _CALC_CHAIN_REL_TYPE
            or rel.attrib.get("Target") == "calcChain.xml"
        ):
            root.remove(rel)

    return _serialize_xml(root)


def _set_cell(
    sheet_root: ET.Element,
    ref: str,
    val: Any,
    shared_root: ET.Element,
    shared_strings: List[str],
    styles_root: Optional[ET.Element],
    written_font_id: Optional[int],
    written_border_id: Optional[int],
    style_cache: Dict[int, int],
) -> None:
    """
    Write *val* to cell *ref* in *sheet_root*.

    - Strings  → shared strings (t="s")
    - int/float → inline number
    - Empty string / None → cell cleared
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
        target_row = ET.Element(
            _ns("row"),
            {"r": str(row_num), "spans": f"{col_idx}:{col_idx}"},
        )
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
        return

    target_row.set("ht", _WRITTEN_ROW_HEIGHT)
    target_row.set("customHeight", "1")

    if (
        styles_root is not None
        and written_font_id is not None
        and written_border_id is not None
    ):
        base_style_idx = _get_cell_style_index(target_cell, target_row)
        style_idx = _get_written_cell_style_index(
            styles_root,
            base_style_idx,
            written_font_id,
            written_border_id,
            style_cache,
        )
        target_cell.set("s", str(style_idx))

    if isinstance(val, str):
        val = _sanitize_excel_text(val)
        if not val:
            return
        idx = _get_shared_idx(shared_root, shared_strings, val)
        target_cell.set("t", "s")
        v = ET.SubElement(target_cell, _ns("v"))
        v.text = str(idx)
    elif isinstance(val, bool):
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
        sheet_data = sheet_root.find(_ns("sheetData"))
        idx = list(sheet_root).index(sheet_data) + 1 if sheet_data is not None else len(sheet_root)
        merge_cells = ET.Element(_ns("mergeCells"))
        sheet_root.insert(idx, merge_cells)

    existing_refs = {mc.attrib.get("ref") for mc in merge_cells.findall(_ns("mergeCell"))}

    for start, end in merges:
        ref = f"{start}:{end}"
        if ref not in existing_refs:
            for covered_ref in _iter_merge_covered_refs(start, end):
                _clear_existing_cell_value(sheet_root, covered_ref)
            ET.SubElement(merge_cells, _ns("mergeCell"), {"ref": ref})
            existing_refs.add(ref)

    merge_cells.set("count", str(len(merge_cells.findall(_ns("mergeCell")))))


# ─────────────────────────────────────────────────────────────────────────────
# Cover-sheet header writer (shared by IL and IO)
# ─────────────────────────────────────────────────────────────────────────────

def _build_cover_updates(payload: Dict[str, Any], doc_number: str) -> Dict[str, Any]:
    """
    Return a dict of {cell_ref: value} for the Cover sheet header area.
    """
    hdr = payload.get("header", {})

    date_str   = hdr.get("date", "")
    excel_date = _to_excel_date(date_str)

    return {
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
# IL data builder  — Section 1 (Field Instruments) ONLY
# ─────────────────────────────────────────────────────────────────────────────
#
# IL column layout
# ────────────────
#   A  = S.No
#   B  = Instrument Code  (auto-generated)
#   C  = Tag No
#   D  = Instrument Name
#   E  = Service Description
#   F  = Line Size
#   G  = Medium
#   H  = Specification
#   I  = Process Conn
#   J  = Working Pressure
#   K  = Working Flow
#   L  = Working Level
#   M  = Design Pressure
#   N  = Design Flow
#   O  = Design Level
#   P  = Set-point
#   Q  = Range
#   R  = UOM
#   S  = Signal Type
#   T  = Velocity          (FM instruments only)
#   U  = FM Size / NB      (FM instruments only)

_IL_DATA_START_ROW = 6   # first data row in the IL data sheet

# Columns that are tag-stable and should be merged within a tag group.
# (Tag No and Instrument Name stay the same across multi-signal rows.)
_IL_TAG_STABLE_COLS = ("C", "D")


def _build_il_data_updates(
    payload: Dict[str, Any],
) -> Tuple[Dict[str, Any], List[Tuple[str, str]]]:
    """
    Build cell updates and merge ranges for the Instrument List data sheet.

    Only Section 1 (field_instruments) is written.
    Electrical equipment and MOVs do NOT appear in the IL.

    Column mapping
    ──────────────
    A  S.No  |  B  Code  |  C  Tag No  |  D  Instrument Name
    E  Service Desc  |  F  Line Size  |  G  Medium  |  H  Specification
    I  Process Conn  |  J  Work Press  |  K  Work Flow  |  L  Work Level
    M  Des Press  |  N  Des Flow  |  O  Des Level  |  P  Set-point
    Q  Range  |  R  UOM  |  S  Signal Type  |  T  Velocity  |  U  FM NB

    Returns
    -------
    (updates, merges)
      updates : {cell_ref: value}
      merges  : [(start_ref, end_ref)]
    """
    updates: Dict[str, Any]         = {}
    merges:  List[Tuple[str, str]]  = []

    fi_rows = [
        row_data
        for row_data in payload.get("field_instruments", [])
        if _is_il_section_1_row(row_data)
    ]
    if not fi_rows:
        return updates, merges

    row    = _IL_DATA_START_ROW
    serial = 1

    # Track tag-group runs for merging tag-stable columns
    run_start = row
    prev_tag  = None

    for rd in fi_rows:
        tag = _row_text(rd, "Tag No", "Tag Number", "Tag Numbers")
        instr = _row_text(rd, "Instrument Name", "Instrument Names")
        svc = _row_text(rd, "Service Description", "Description")

        # ── Instrument code (column B) ────────────────────────────────────────
        code = _get_instrument_code(instr) or _row_text(rd, "Instrument Code", "Code")

        # ── Flowmeter velocity + NB (columns T, U) ───────────────────────────
        # Only computed when the instrument resolves to the "FM" code.
        velocity_str = ""
        nb_str       = ""
        if code == "FM":
            work_flow = _row_text(rd, "Working Flow", "Work Flow")
            line_size = _row_text(rd, "Line Size")
            velocity_str, nb_str = _compute_fm_velocity_and_nb(work_flow, line_size)

        if not velocity_str:
            velocity_str = _row_text(rd, "Velocity")
        if not nb_str:
            nb_str = _row_text(rd, "FM Size", "FM Sizes", "FM NB", "NB", "NB Size")

        # ── Tag-group merge tracking (flush completed run) ────────────────────
        if tag != prev_tag and prev_tag is not None:
            if row - run_start > 1:
                for col in _IL_TAG_STABLE_COLS:
                    merges.append((f"{col}{run_start}", f"{col}{row - 1}"))
            run_start = row

        # ── Write cells ───────────────────────────────────────────────────────
        updates[f"A{row}"] = serial
        updates[f"B{row}"] = code
        updates[f"C{row}"] = tag
        updates[f"D{row}"] = instr
        updates[f"E{row}"] = svc
        updates[f"F{row}"] = _row_text(rd, "Line Size")
        updates[f"G{row}"] = _row_text(rd, "Medium")
        updates[f"H{row}"] = _row_text(rd, "Specification")
        updates[f"I{row}"] = _row_text(rd, "Process connection", "Process Connection", "Process Conn")
        updates[f"J{row}"] = _row_text(rd, "Working Pressure", "Work Press")
        updates[f"K{row}"] = _row_text(rd, "Working Flow", "Work Flow")
        updates[f"L{row}"] = _row_text(rd, "Working Level", "Work Level")
        updates[f"M{row}"] = _row_text(rd, "Design Pressure", "Design Press")
        updates[f"N{row}"] = _row_text(rd, "Design Flow")
        updates[f"O{row}"] = _row_text(rd, "Design Level")
        updates[f"P{row}"] = _row_text(rd, "Setpoint", "Set-point")
        updates[f"Q{row}"] = _row_text(rd, "Instrument Range", "Range")
        updates[f"R{row}"] = _row_text(rd, "UOM")
        updates[f"S{row}"] = _row_text(rd, "Signal Type")
        updates[f"T{row}"] = velocity_str
        updates[f"U{row}"] = nb_str

        prev_tag = tag or None
        row     += 1
        serial  += 1

    # Flush the final tag-group run
    if prev_tag is not None and row - run_start > 1:
        for col in _IL_TAG_STABLE_COLS:
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

# Columns that are tag-stable within a tag group in the IO sheet
_IO_TAG_STABLE_COLS = ("B", "C", "F", "G")


def _build_io_data_updates(
    payload: Dict[str, Any],
) -> Tuple[Dict[str, Any], List[Tuple[str, str]]]:
    """
    Build cell updates and merge ranges for the IO List data sheet.

    All three sections (field instruments, electrical, MOVs) appear in the
    IO List — unlike the IL which only shows field instruments.

    Column mapping
    ──────────────
    A  S.No  |  B  Tag No  |  C  Instrument Name  |  D  Service Description
    E  Signal Type  |  F  Source  |  G  Destination
    H  DI  |  I  DO  |  J  AI  |  K  AO
    M  Trip flag  |  N  Remarks

    Returns
    -------
    (updates, merges)
      updates : {cell_ref: value}
      merges  : [(start_ref, end_ref)]
    """
    updates: Dict[str, Any]        = {}
    merges: List[Tuple[str, str]]  = []
    row    = _IO_DATA_START_ROW
    serial = 1

    sections = [
        (
            "Section 1 — Field Instruments",
            payload.get("field_instruments", []),
            False,   # concat_sig_desc
        ),
        (
            "Section 2 — Electrical Equipment",
            payload.get("electrical", []),
            True,
        ),
        (
            "Section 3 — Motor Operated Valves",
            payload.get("mov", []),
            True,
        ),
    ]

    for label, rows, concat_sig_desc in sections:
        if not rows:
            continue

        # Section header row — merged across all IO columns (A→N)
        updates[f"A{row}"] = label
        merges.append((f"A{row}", f"N{row}"))
        row += 1

        run_start = row
        prev_tag  = None

        for rd in rows:
            tag      = rd.get("Tag No",              "").strip()
            instr    = rd.get("Instrument Name",     "").strip()
            svc      = rd.get("Service Description", "").strip()
            sig_desc = rd.get("Signal Description",  "").strip()
            sig_type = rd.get("Signal Type",         "").strip()
            source   = rd.get("Source",              "").strip()
            dest     = rd.get("Destination",         "").strip()
            signal   = (rd.get("Signal") or "").strip().upper()

            # For electrical / MOV rows, append signal description to service
            if concat_sig_desc and sig_desc:
                svc = f"{svc} — {sig_desc}" if svc else sig_desc

            # Tag-group merge tracking (flush completed run)
            if tag != prev_tag and prev_tag is not None:
                if row - run_start > 1:
                    for col in _IO_TAG_STABLE_COLS:
                        merges.append((f"{col}{run_start}", f"{col}{row - 1}"))
                run_start = row

            updates[f"A{row}"] = serial
            updates[f"B{row}"] = tag
            updates[f"C{row}"] = instr
            updates[f"D{row}"] = svc
            updates[f"E{row}"] = sig_type
            updates[f"F{row}"] = source
            updates[f"G{row}"] = dest

            # Signal tick marks
            updates[f"H{row}"] = 1 if signal == "DI" else ""
            updates[f"I{row}"] = 1 if signal == "DO" else ""
            updates[f"J{row}"] = 1 if signal == "AI" else ""
            updates[f"K{row}"] = 1 if signal == "AO" else ""

            # Trip flag
            updates[f"M{row}"] = "TRIP" if "trip" in svc.lower() else ""

            prev_tag = tag or None
            row     += 1
            serial  += 1

        # Flush the final tag-group run for this section
        if prev_tag is not None and row - run_start > 1:
            for col in _IO_TAG_STABLE_COLS:
                merges.append((f"{col}{run_start}", f"{col}{row - 1}"))

    return updates, merges


# ─────────────────────────────────────────────────────────────────────────────
# CS data builder
# ─────────────────────────────────────────────────────────────────────────────
#
# CS column layout (data sheet named "CS")
# ─────────────────────────────────────────────────────────────────────────────
#   A  = S.No
#   B  = Tag No
#   C  = Cable Type          (deferred — left blank for now)
#   D  = Service Description
#   E  = Instrument Name
#   F  = From                (deferred)
#   G  = To                  (deferred)
#   H  = Run                 (deferred)
#   I  = Route Length        (deferred)
#   J  = Total Length        (deferred)

_CS_DATA_START_ROW = 6   # first data row — adjust if template differs


def _build_cs_data_updates(
    payload: Dict[str, Any],
) -> Tuple[Dict[str, Any], List[Tuple[str, str]]]:
    """
    Build cell updates for the Cable Schedule data sheet.

    All three sections (field instruments, electrical, MOVs) contribute
    rows — every tagged instrument corresponds to at least one cable entry.

    Columns C and F-J are deferred and left blank until their logic is
    finalised.

    Returns
    -------
    (updates, merges)
      updates : {cell_ref: value}
      merges  : [(start_ref, end_ref)]   — empty for now
    """
    updates: Dict[str, Any]       = {}
    merges: List[Tuple[str, str]] = []

    row    = _CS_DATA_START_ROW
    serial = 1

    sections = [
        payload.get("field_instruments", []),
        payload.get("electrical",        []),
        payload.get("mov",               []),
    ]

    for section_rows in sections:
        for rd in section_rows:
            tag   = rd.get("Tag No",              "").strip()
            svc   = rd.get("Service Description", "").strip()
            instr = rd.get("Instrument Name",     "").strip()

            # Skip completely empty rows (schema already filters most,
            # but guard here in case a section has stale blanks).
            if not any((tag, svc, instr)):
                continue

            updates[f"A{row}"] = serial
            updates[f"B{row}"] = tag
            # Column C (Cable Type) — deferred, left blank intentionally
            updates[f"D{row}"] = svc
            updates[f"E{row}"] = instr
            # Columns F-J — deferred, left blank intentionally

            row    += 1
            serial += 1

    return updates, merges


def write_cable_schedule(payload: Dict[str, Any]) -> BytesIO:
    """
    Generate a Cable Schedule workbook from *payload*.

    Cover sheet  → header metadata (AI6–AI15), same mapping as IL and IO.
    CS data sheet → all tagged instruments, columns A B D E populated.

    Returns a BytesIO ready for send_file().
    """
    zin_tmp   = zipfile.ZipFile(BytesIO(base64.b64decode(CS_TEMPLATE_B64)), "r")
    sheet_map = _map_sheets(zin_tmp)
    zin_tmp.close()

    cover_path = sheet_map.get("Cover",          "xl/worksheets/sheet1.xml")
    data_path  = sheet_map.get("CS",             "xl/worksheets/sheet2.xml")

    doc_number = (payload.get("cs_meta") or {}).get("docNumber", "")

    cover_updates         = _build_cover_updates(payload, doc_number)
    cs_updates, cs_merges = _build_cs_data_updates(payload)

    return _process(
        CS_TEMPLATE_B64,
        updates_by_sheet={
            cover_path: cover_updates,
            data_path:  cs_updates,
        },
        merges_by_sheet={
            data_path: cs_merges,
        },
    )


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

    Key correctness requirements
    ────────────────────────────
    1. Every XML file written back into the zip MUST start with the XML
       declaration (<?xml version="1.0" encoding="UTF-8" standalone="yes"?>).
       Excel marks files without it as damaged.
    2. All namespace prefixes must be registered before parsing so ET
       preserves them on serialisation.  Unregistered namespaces get
       renamed to ns0:/ns1: which also corrupts the file.
    3. Zip entries that we do NOT modify are copied byte-for-byte from
       the template — no re-compression, no re-serialisation.
    """
    template_bytes = base64.b64decode(template_b64)

    zin         = zipfile.ZipFile(BytesIO(template_bytes), "r")
    zout_buffer = BytesIO()
    zout        = zipfile.ZipFile(zout_buffer, "w", zipfile.ZIP_DEFLATED)

    # Parse shared strings once — all sheet writers share the same pool.
    shared_root, shared_strings = _load_shared_strings(zin)
    styles_root = _load_styles(zin)
    written_font_id = (
        _ensure_written_cell_font(styles_root)
        if styles_root is not None
        else None
    )
    written_border_id = (
        _ensure_written_cell_border(styles_root)
        if styles_root is not None
        else None
    )
    written_style_cache: Dict[int, int] = {}

    for item in zin.infolist():
        if item.filename == "xl/calcChain.xml":
            continue

        if item.filename == "xl/sharedStrings.xml":
            continue

        if item.filename == "xl/styles.xml":
            continue

        data = zin.read(item.filename)

        if item.filename in updates_by_sheet:
            # ── Parse → mutate → serialise ────────────────────────────────────
            root = ET.fromstring(data)

            sheet_updates = updates_by_sheet[item.filename]
            sheet_merges  = merges_by_sheet.get(item.filename, [])

            for ref, val in sheet_updates.items():
                _set_cell(
                    root,
                    ref,
                    val,
                    shared_root,
                    shared_strings,
                    styles_root,
                    written_font_id,
                    written_border_id,
                    written_style_cache,
                )

            if sheet_merges:
                _apply_merges(root, sheet_merges)

            data = _serialize_xml(root)

        elif item.filename == "[Content_Types].xml":
            data = _remove_calc_chain_content_type(data)

        elif item.filename == "xl/_rels/workbook.xml.rels":
            data = _remove_calc_chain_relationship(data)

        # All other files (styles, workbook, relationships, images, …) are
        # copied unchanged — do NOT re-parse or re-serialise them.
        zout.writestr(item, data)

    if shared_root is not None:
        zout.writestr("xl/sharedStrings.xml", _serialize_xml(shared_root))

    if styles_root is not None:
        zout.writestr("xl/styles.xml", _serialize_xml(styles_root))

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

    - Cover sheet  → header metadata (AI6–AI15)
    - IL data sheet → Section 1 field instruments ONLY
                      (Sections 2 & 3 are NOT written to the IL)

    Column layout: A=S.No, B=Code, C=Tag, D=Name, E=Desc, F=Line Size,
    G=Medium, H=Spec, I=Proc Conn, J=Work Press, K=Work Flow, L=Work Level,
    M=Des Press, N=Des Flow, O=Des Level, P=Set-point, Q=Range, R=UOM,
    S=Signal Type, T=Velocity (FM only), U=NB size (FM only)

    Returns a BytesIO ready for send_file().
    """
    zin_tmp   = zipfile.ZipFile(BytesIO(base64.b64decode(IL_TEMPLATE_B64)), "r")
    sheet_map = _map_sheets(zin_tmp)
    zin_tmp.close()

    cover_path = sheet_map.get("Cover",           "xl/worksheets/sheet1.xml")
    data_path  = sheet_map.get("Instrument List", "xl/worksheets/sheet2.xml")

    doc_number = (payload.get("fi_meta") or {}).get("docNumber", "")

    cover_updates         = _build_cover_updates(payload, doc_number)
    il_updates, il_merges = _build_il_data_updates(payload)

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

    - Cover sheet    → header metadata (AI6–AI15)
    - IO List sheet  → all three sections (FI + Electrical + MOV)

    Column layout: A=S.No, B=Tag, C=Name, D=Desc, E=Signal Type,
    F=Source, G=Dest, H=DI, I=DO, J=AI, K=AO, M=Trip flag, N=Remarks

    Returns a BytesIO ready for send_file().
    """
    zin_tmp   = zipfile.ZipFile(BytesIO(base64.b64decode(IO_TEMPLATE_B64)), "r")
    sheet_map = _map_sheets(zin_tmp)
    zin_tmp.close()

    cover_path = sheet_map.get("Cover",   "xl/worksheets/sheet1.xml")
    data_path  = sheet_map.get("IO List", "xl/worksheets/sheet2.xml")

    doc_number = (payload.get("io_meta") or {}).get("docNumber", "")

    cover_updates         = _build_cover_updates(payload, doc_number)
    io_updates, io_merges = _build_io_data_updates(payload)

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
