import zipfile
from io import BytesIO
from xml.etree import ElementTree as ET

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
}

def _ns(tag):
    return f"{{{NS['main']}}}{tag}"


# ─────────────────────────────────────────────────────────────
# Shared Strings Handling
# ─────────────────────────────────────────────────────────────

def _load_shared_strings(zipf):
    try:
        data = zipf.read("xl/sharedStrings.xml")
        root = ET.fromstring(data)

        strings = []
        for si in root.findall(_ns("si")):
            t = si.find(_ns("t"))
            strings.append(t.text if t is not None else "")

        return root, strings

    except KeyError:
        return None, []


def _get_or_add_shared_string(root, strings, value):
    if value in strings:
        return strings.index(value)

    idx = len(strings)
    strings.append(value)

    si = ET.SubElement(root, _ns("si"))
    t = ET.SubElement(si, _ns("t"))
    t.text = value

    return idx


# ─────────────────────────────────────────────────────────────
# Cell update
# ─────────────────────────────────────────────────────────────

def _set_cell_value(sheet_root, cell_ref, value, shared_root, shared_strings):
    for c in sheet_root.iter(_ns("c")):
        if c.attrib.get("r") == cell_ref:

            # Remove old value
            v = c.find(_ns("v"))
            if v is not None:
                c.remove(v)

            # Handle string
            if isinstance(value, str):
                idx = _get_or_add_shared_string(shared_root, shared_strings, value)

                c.set("t", "s")  # shared string
                v = ET.SubElement(c, _ns("v"))
                v.text = str(idx)

            else:
                c.attrib.pop("t", None)
                v = ET.SubElement(c, _ns("v"))
                v.text = str(value)

            return

    # If cell doesn't exist → create it (important)
    row_num = int(''.join(filter(str.isdigit, cell_ref)))

    sheetData = sheet_root.find(_ns("sheetData"))

    row = None
    for r in sheetData.findall(_ns("row")):
        if int(r.attrib["r"]) == row_num:
            row = r
            break

    if row is None:
        row = ET.SubElement(sheetData, _ns("row"), {"r": str(row_num)})

    c = ET.SubElement(row, _ns("c"), {"r": cell_ref})

    if isinstance(value, str):
        idx = _get_or_add_shared_string(shared_root, shared_strings, value)
        c.set("t", "s")
        v = ET.SubElement(c, _ns("v"))
        v.text = str(idx)
    else:
        v = ET.SubElement(c, _ns("v"))
        v.text = str(value)


# ─────────────────────────────────────────────────────────────
# Main patch engine
# ─────────────────────────────────────────────────────────────

def patch_excel(template_bytes, updates_by_sheet):
    zin = zipfile.ZipFile(BytesIO(template_bytes), "r")
    zout_buffer = BytesIO()
    zout = zipfile.ZipFile(zout_buffer, "w", zipfile.ZIP_DEFLATED)

    shared_root, shared_strings = _load_shared_strings(zin)

    for item in zin.infolist():
        data = zin.read(item.filename)

        if item.filename.startswith("xl/worksheets/sheet"):
            sheet_root = ET.fromstring(data)

            sheet_name = item.filename  # map manually later

            if sheet_name in updates_by_sheet:
                updates = updates_by_sheet[sheet_name]

                for cell_ref, value in updates.items():
                    _set_cell_value(
                        sheet_root,
                        cell_ref,
                        value,
                        shared_root,
                        shared_strings,
                    )

                data = ET.tostring(sheet_root)

        elif item.filename == "xl/sharedStrings.xml" and shared_root is not None:
            data = ET.tostring(shared_root)

        zout.writestr(item, data)

    zin.close()
    zout.close()

    zout_buffer.seek(0)
    return zout_buffer

def map_sheet_names(zipf):
    workbook_xml = ET.fromstring(zipf.read("xl/workbook.xml"))

    rels = ET.fromstring(zipf.read("xl/_rels/workbook.xml.rels"))

    rel_map = {
        r.attrib["Id"]: r.attrib["Target"]
        for r in rels
    }

    mapping = {}

    for sheet in workbook_xml.findall(_ns("sheets"))[0]:
        name = sheet.attrib["name"]
        rid = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
        target = rel_map[rid]
        mapping[name] = f"xl/{target}"

    return mapping