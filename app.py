from datetime import datetime
from pathlib import Path
import traceback
from io import BytesIO
from flask import Flask, render_template, request, jsonify, send_file
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

app = Flask(__name__)

# Constants for Styling
HEADER_FILL = PatternFill("solid", fgColor="D9EAF7")
HEADER_FONT = Font(bold=True)
WRAP_TOP = Alignment(vertical="top", wrap_text=True)

def clean(value):
    return value.strip() if isinstance(value, str) else value

def get_list(form, key):
    return [clean(v) for v in form.getlist(key)]

def build_generic_section(form, prefix):
    """Refactored helper to handle all grid sections and unpack P/F/L values."""
    tags = get_list(form, f"{prefix}TagNo[]")
    instruments = get_list(form, f"{prefix}Instrument[]")
    services = get_list(form, f"{prefix}ServiceDescription[]")
    line_sizes = get_list(form, f"{prefix}LineSize[]")
    mediums = get_list(form, f"{prefix}Medium[]")
    specs = get_list(form, f"{prefix}TypeSpec[]")
    conns = get_list(form, f"{prefix}ProcessConnection[]")
    working_vals = get_list(form, f"{prefix}WorkingValues[]")
    design_vals = get_list(form, f"{prefix}DesignValues[]")
    set_points = get_list(form, f"{prefix}SetPoint[]")
    ranges = get_list(form, f"{prefix}InstrumentRange[]")
    uoms = get_list(form, f"{prefix}Uom[]")

    rows = []
    max_len = max(len(tags), len(working_vals), 0)

    for i in range(max_len):
        # Unpack P/F/L strings (e.g., "10/50/20")
        w_vals = working_vals[i].split('/') if i < len(working_vals) else ["", "", ""]
        d_vals = design_vals[i].split('/') if i < len(design_vals) else ["", "", ""]
        
        # Ensure we have at least 3 elements even if split fails
        w_vals += [""] * (3 - len(w_vals))
        d_vals += [""] * (3 - len(d_vals))

        row = {
            "Tag No": tags[i] if i < len(tags) else "",
            "Instrument": instruments[i] if i < len(instruments) else "",
            "Service": services[i] if i < len(services) else "",
            "Size (mm)": line_sizes[i] if i < len(line_sizes) else "",
            "Medium": mediums[i] if i < len(mediums) else "",
            "Spec": specs[i] if i < len(specs) else "",
            "Connection": conns[i] if i < len(conns) else "",
            "Work Press": w_vals[0], "Work Flow": w_vals[1], "Work Level": w_vals[2],
            "Design Press": d_vals[0], "Design Flow": d_vals[1], "Design Level": d_vals[2],
            "Set Point": set_points[i] if i < len(set_points) else "",
            "Range": ranges[i] if i < len(ranges) else "",
            "UOM": uoms[i] if i < len(uoms) else ""
        }
        if any(row.values()):
            rows.append(row)
    return rows

def write_sheet(ws, rows, title):
    ws.title = title
    if not rows:
        ws.append(["No data"])
        return
    headers = list(rows[0].keys())
    ws.append(headers)
    for cell in ws[1]:
        cell.fill, cell.font, cell.alignment = HEADER_FILL, HEADER_FONT, WRAP_TOP
    for row in rows:
        ws.append([row.get(h, "") for h in headers])
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 18

@app.route("/submit", methods=["POST"])
def submit():
    try:
        payload = {
            "header": {k: request.form.get(k, "") for k in ["projectName", "documentName", "client", "documentNumber"]},
            "field_instruments": build_generic_section(request.form, "fi"),
            "electrical": build_generic_section(request.form, "el"),
            "mov": build_generic_section(request.form, "mov"),
        }
        
        output = BytesIO()
        wb = Workbook()
        write_sheet(wb.active, payload["field_instruments"], "Field Instruments")
        write_sheet(wb.create_sheet("Electrical"), payload["electrical"], "Electrical")
        write_sheet(wb.create_sheet("MOV"), payload["mov"], "MOV")
        
        wb.save(output)
        output.seek(0)
        
        return send_file(output, as_attachment=True, download_name=f"Instrument_List_{datetime.now().strftime('%Y%m%d')}.xlsx")
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route("/")
def index(): return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True)