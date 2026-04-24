import base64
import zipfile
from io import BytesIO
from typing import Dict, Any, List, Tuple
from xml.etree import ElementTree as ET

from utils.embedded_templates import IL_TEMPLATE_B64, IO_TEMPLATE_B64


NS = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

def _ns(tag):
    return f"{{{NS['main']}}}{tag}"


# ─────────────────────────────────────────────────────────────
# SHARED STRINGS
# ─────────────────────────────────────────────────────────────

def _load_shared_strings(zipf):
    try:
        root = ET.fromstring(zipf.read("xl/sharedStrings.xml"))
        strings = [t.text if (t := si.find(_ns("t"))) is not None else ""
                   for si in root.findall(_ns("si"))]
        return root, strings
    except KeyError:
        return None, []


def _get_shared_idx(root, strings, val):
    if val in strings:
        return strings.index(val)
    idx = len(strings)
    strings.append(val)
    si = ET.SubElement(root, _ns("si"))
    t = ET.SubElement(si, _ns("t"))
    t.text = val
    return idx


# ─────────────────────────────────────────────────────────────
# CELL WRITE
# ─────────────────────────────────────────────────────────────

def _set_cell(sheet, ref, val, shared_root, shared_strings):
    for c in sheet.iter(_ns("c")):
        if c.attrib.get("r") == ref:
            for v in c.findall(_ns("v")):
                c.remove(v)

            if isinstance(val, str):
                idx = _get_shared_idx(shared_root, shared_strings, val)
                c.set("t", "s")
                v = ET.SubElement(c, _ns("v"))
                v.text = str(idx)
            else:
                c.attrib.pop("t", None)
                v = ET.SubElement(c, _ns("v"))
                v.text = str(val)
            return


# ─────────────────────────────────────────────────────────────
# MERGE CELLS (CRITICAL PART)
# ─────────────────────────────────────────────────────────────

def _apply_merges(sheet_root, merges: List[Tuple[str, str]]):
    mergeCells = sheet_root.find(_ns("mergeCells"))

    if mergeCells is None:
        mergeCells = ET.SubElement(sheet_root, _ns("mergeCells"))

    for start, end in merges:
        ET.SubElement(
            mergeCells,
            _ns("mergeCell"),
            {"ref": f"{start}:{end}"}
        )

    mergeCells.set("count", str(len(mergeCells)))


# ─────────────────────────────────────────────────────────────
# SHEET MAPPING
# ─────────────────────────────────────────────────────────────

def _map_sheets(zipf):
    wb = ET.fromstring(zipf.read("xl/workbook.xml"))
    rels = ET.fromstring(zipf.read("xl/_rels/workbook.xml.rels"))

    rel_map = {r.attrib["Id"]: r.attrib["Target"] for r in rels}

    mapping = {}
    for s in wb.find(_ns("sheets")):
        name = s.attrib["name"]
        rid = s.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
        mapping[name] = "xl/" + rel_map[rid]

    return mapping


# ─────────────────────────────────────────────────────────────
# IO LOGIC (MERGING + HEADERS)
# ─────────────────────────────────────────────────────────────

def _build_io_updates(payload):
    updates = {}
    merges = []

    row = 6
    serial = 1

    sections = [
        ("Field Instruments", payload.get("field_instruments", []), False),
        ("Electrical Equipment", payload.get("electrical", []), True),
        ("Motor Operated Valves", payload.get("mov", []), True),
    ]

    for label, rows, concat in sections:

        # Section header
        updates[f"A{row}"] = label
        merges.append((f"A{row}", f"N{row}"))
        row += 1

        start_run = row
        prev_tag = None

        for rd in rows:
            tag = rd.get("Tag No", "")
            instr = rd.get("Instrument Name", "")
            service = rd.get("Service Description", "")
            sig_desc = rd.get("Signal Description", "")
            sig_type = rd.get("Signal Type", "")
            source = rd.get("Source", "")
            dest = rd.get("Destination", "")
            signal = rd.get("Signal", "").upper()

            if concat and sig_desc:
                service = f"{service} - {sig_desc}" if service else sig_desc

            updates[f"A{row}"] = serial
            updates[f"B{row}"] = tag
            updates[f"C{row}"] = instr
            updates[f"D{row}"] = service
            updates[f"E{row}"] = sig_type
            updates[f"F{row}"] = source
            updates[f"G{row}"] = dest

            updates[f"H{row}"] = 1 if signal == "DI" else ""
            updates[f"I{row}"] = 1 if signal == "DO" else ""
            updates[f"J{row}"] = 1 if signal == "AI" else ""
            updates[f"K{row}"] = 1 if signal == "AO" else ""

            updates[f"M{row}"] = "TRIP" if "trip" in service.lower() else ""

            # MERGE LOGIC
            if tag != prev_tag and prev_tag is not None:
                if row - start_run > 1:
                    for col in ["B", "C", "F", "G"]:
                        merges.append((f"{col}{start_run}", f"{col}{row-1}"))
                start_run = row

            prev_tag = tag
            row += 1
            serial += 1

        # finalize last run
        if row - start_run > 1:
            for col in ["B", "C", "F", "G"]:
                merges.append((f"{col}{start_run}", f"{col}{row-1}"))

    return updates, merges


# ─────────────────────────────────────────────────────────────
# CORE ENGINE
# ─────────────────────────────────────────────────────────────

def _process(template_b64, updates_by_sheet, merges_by_sheet):
    template_bytes = base64.b64decode(template_b64)

    zin = zipfile.ZipFile(BytesIO(template_bytes))
    zout_buffer = BytesIO()
    zout = zipfile.ZipFile(zout_buffer, "w", zipfile.ZIP_DEFLATED)

    sheet_map = _map_sheets(zin)
    shared_root, shared_strings = _load_shared_strings(zin)

    for item in zin.infolist():
        data = zin.read(item.filename)

        if item.filename in updates_by_sheet:
            root = ET.fromstring(data)

            updates = updates_by_sheet[item.filename]
            merges = merges_by_sheet.get(item.filename, [])

            for ref, val in updates.items():
                _set_cell(root, ref, val, shared_root, shared_strings)

            if merges:
                _apply_merges(root, merges)

            data = ET.tostring(root)

        elif item.filename == "xl/sharedStrings.xml" and shared_root is not None:
            data = ET.tostring(shared_root)

        zout.writestr(item, data)

    zout.close()
    zout_buffer.seek(0)
    return zout_buffer


# ─────────────────────────────────────────────────────────────
# PUBLIC API
# ─────────────────────────────────────────────────────────────

def write_io_workbook(payload: Dict[str, Any]):
    updates, merges = _build_io_updates(payload)

    return _process(
        IO_TEMPLATE_B64,
        updates_by_sheet={
            "xl/worksheets/sheet2.xml": updates  # adjust if needed
        },
        merges_by_sheet={
            "xl/worksheets/sheet2.xml": merges
        }
    )