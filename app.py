from datetime import datetime
from pathlib import Path
import traceback
from flask import Flask, render_template, request, jsonify, send_file
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

app = Flask(__name__)

HEADER_FILL = PatternFill("solid", fgColor="D9EAF7")
HEADER_FONT = Font(bold=True)
WRAP_TOP = Alignment(vertical="top", wrap_text=True)

# -------------------------
# Helpers
# -------------------------

def clean(value):
    return value.strip() if isinstance(value, str) else value

def get_scalar(form, key):
    return clean(form.get(key, ""))

def get_list(form, key):
    return [clean(v) for v in form.getlist(key)]

# -------------------------
# Builders
# -------------------------

def build_field_instruments(form):
    tags = get_list(form, "fiTagNo[]")
    instruments = get_list(form, "fiInstrument[]")
    services = get_list(form, "fiServiceDescription[]")
    line_sizes = get_list(form, "fiLineSize[]")
    mediums = get_list(form, "fiMedium[]")
    type_specs = get_list(form, "fiTypeSpec[]")
    proc_conns = get_list(form, "fiProcessConnection[]")
    working_vals = get_list(form, "fiWorkingValues[]")
    design_vals = get_list(form, "fiDesignValues[]")
    set_points = get_list(form, "fiSetPoint[]")
    instr_ranges = get_list(form, "fiInstrumentRange[]")
    uoms = get_list(form, "fiUom[]")

    rows = []
    max_len = max([
        len(tags), len(instruments), len(services), len(line_sizes), len(mediums),
        len(type_specs), len(proc_conns), len(working_vals), len(design_vals),
        len(set_points), len(instr_ranges), len(uoms)
    ], default=0)

    for i in range(max_len):
        row = {
            "tag_no": tags[i] if i < len(tags) else "",
            "instrument": instruments[i] if i < len(instruments) else "",
            "service_description": services[i] if i < len(services) else "",
            "line_size_mm": line_sizes[i] if i < len(line_sizes) else "",
            "medium": mediums[i] if i < len(mediums) else "",
            "type_spec": type_specs[i] if i < len(type_specs) else "",
            "process_connection": proc_conns[i] if i < len(proc_conns) else "",
            "working_values": working_vals[i] if i < len(working_vals) else "",
            "design_values": design_vals[i] if i < len(design_vals) else "",
            "set_point": set_points[i] if i < len(set_points) else "",
            "instrument_range": instr_ranges[i] if i < len(instr_ranges) else "",
            "uom": uoms[i] if i < len(uoms) else "",
        }
        if any(row.values()):
            rows.append(row)
    return rows

def build_equipment_section(form, prefix):
    tags = get_list(form, f"{prefix}TagNo[]")
    instruments = get_list(form, f"{prefix}Instrument[]")
    services = get_list(form, f"{prefix}ServiceDescription[]")
    line_sizes = get_list(form, f"{prefix}LineSize[]")
    mediums = get_list(form, f"{prefix}Medium[]")
    type_specs = get_list(form, f"{prefix}TypeSpec[]")
    proc_conns = get_list(form, f"{prefix}ProcessConnection[]")
    working_vals = get_list(form, f"{prefix}WorkingValues[]")
    design_vals = get_list(form, f"{prefix}DesignValues[]")
    set_points = get_list(form, f"{prefix}SetPoint[]")
    instr_ranges = get_list(form, f"{prefix}InstrumentRange[]")
    uoms = get_list(form, f"{prefix}Uom[]")

    rows = []
    max_len = max([
        len(tags), len(instruments), len(services), len(line_sizes), len(mediums),
        len(type_specs), len(proc_conns), len(working_vals), len(design_vals),
        len(set_points), len(instr_ranges), len(uoms)
    ], default=0)

    for i in range(max_len):
        row = {
            "tag_no": tags[i] if i < len(tags) else "",
            "instrument": instruments[i] if i < len(instruments) else "",
            "service_description": services[i] if i < len(services) else "",
            "line_size_mm": line_sizes[i] if i < len(line_sizes) else "",
            "medium": mediums[i] if i < len(mediums) else "",
            "type_spec": type_specs[i] if i < len(type_specs) else "",
            "process_connection": proc_conns[i] if i < len(proc_conns) else "",
            "working_values": working_vals[i] if i < len(working_vals) else "",
            "design_values": design_vals[i] if i < len(design_vals) else "",
            "set_point": set_points[i] if i < len(set_points) else "",
            "instrument_range": instr_ranges[i] if i < len(instr_ranges) else "",
            "uom": uoms[i] if i < len(uoms) else "",
        }
        if any(row.values()):
            rows.append(row)
    return rows


def parse_form_data(form):
    return {
        "header": {
            "project_name": get_scalar(form, "projectName"),
            "document_name": get_scalar(form, "documentName"),
            "client": get_scalar(form, "client"),
            "consultant": get_scalar(form, "consultant"),
            "document_number": get_scalar(form, "documentNumber"),
            "date": get_scalar(form, "date"),
            "prepared_by": get_scalar(form, "preparedBy"),
            "checked_by": get_scalar(form, "checkedBy"),
            "approved_by": get_scalar(form, "approvedBy"),
            "revision": get_scalar(form, "revision"),
        },
        "field_instruments": build_field_instruments(form),
        "electrical_equipment": build_equipment_section(form, "el"),
        "mov_equipment": build_equipment_section(form, "mov"),
    }


def validate_payload(payload):
    errors = []
    header = payload.get("header", {})
    if not header.get("project_name"):
        errors.append("Project Name is required.")
    if not header.get("document_name"):
        errors.append("Document Name is required.")
    if not header.get("document_number"):
        errors.append("Document Number is required.")

    tag_set = set()

    # Validate Field Instruments
    for idx, row in enumerate(payload.get("field_instruments", []), start=1):
        tag = row.get("tag_no", "")
        if not tag:
            errors.append(f"Field Instrument row {idx}: Tag No is required.")
        elif tag in tag_set:
            errors.append(f"Duplicate Tag No found: {tag}")
        else:
            tag_set.add(tag)

    # Validate Electrical and MOV sections
    for section_name in ("electrical_equipment", "mov_equipment"):
        for idx, equipment in enumerate(payload.get(section_name, []), start=1):
            tag = equipment.get("tag_no", "")
            if not tag:
                errors.append(f"{section_name} row {idx}: Tag No is required.")
            elif tag in tag_set:
                errors.append(f"Duplicate Tag No found: {tag}")
            else:
                tag_set.add(tag)

    return errors


def write_sheet(ws, rows, title):
    ws.title = title
    if not rows:
        ws.append(["No data"])
        return

    headers = list(rows[0].keys())
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = WRAP_TOP

    for row in rows:
        ws.append([row.get(h, "") for h in headers])

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = WRAP_TOP

    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            value = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, len(value))
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 14), 45)


def build_master_list(payload):
    master = []
    for row in payload.get("field_instruments", []):
        master.append({"category": "Field Instrument", **row})
    for row in payload.get("electrical_equipment", []):
        master.append({"category": "Electrical", **row})
    for row in payload.get("mov_equipment", []):
        master.append({"category": "MOV", **row})
    return master


def write_workbook(payload, output_path):
    wb = Workbook()

    # Header
    ws_header = wb.active
    write_sheet(
        ws_header,
        [{"field": k, "value": v} for k, v in payload["header"].items()],
        "Header",
    )

    # Stage 1 sections
    ws_fi = wb.create_sheet("Field Instruments")
    write_sheet(ws_fi, payload["field_instruments"], "Field Instruments")

    ws_el = wb.create_sheet("Electrical Equip")
    write_sheet(ws_el, payload["electrical_equipment"], "Electrical Equip")

    ws_mov = wb.create_sheet("MOV Equip")
    write_sheet(ws_mov, payload["mov_equipment"], "MOV Equip")

    # Hidden master list
    ws_master = wb.create_sheet("Master List")
    write_sheet(ws_master, build_master_list(payload), "Master List")
    ws_master.sheet_state = 'hidden'

    wb.save(output_path)


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/preview", methods=["POST"])
def preview():
    try:
        payload = parse_form_data(request.form)
        errors = validate_payload(payload)
        return jsonify({
            "success": len(errors) == 0,
            "payload": payload,
            "errors": errors,
        })
    except Exception as exc:
        traceback.print_exc()
        return jsonify({
            "success": False,
            "message": "Preview failed.",
            "error": str(exc),
        }), 500


@app.route("/submit", methods=["POST"])
def submit():
    try:
        payload = parse_form_data(request.form)
        errors = validate_payload(payload)
        if errors:
            return jsonify({
                "success": False,
                "message": "Validation failed.",
                "errors": errors,
            }), 400

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"cis_output_{timestamp}.xlsx"
        output_path = OUTPUT_DIR / filename
        write_workbook(payload, output_path)

        return jsonify({
            "success": True,
            "message": "Workbook generated successfully.",
            "download_url": f"/download/{filename}",
        })
    except Exception as exc:
        traceback.print_exc()
        return jsonify({
            "success": False,
            "message": "Unexpected server error.",
            "error": str(exc),
        }), 500


@app.route("/download/<filename>", methods=["GET"])
def download_file(filename):
    file_path = OUTPUT_DIR / filename
    if not file_path.exists():
        return jsonify({
            "success": False,
            "message": "File not found.",
        }), 404
    return send_file(file_path, as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True)