from datetime import datetime
from pathlib import Path
import traceback
from io import BytesIO
from flask import Flask, render_template, request, jsonify, send_file
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from utils.validator import validate_payload

app = Flask(__name__)

# Constants for Styling
HEADER_FILL = PatternFill("solid", fgColor="D9EAF7")
HEADER_FONT = Font(bold=True)
WRAP_TOP = Alignment(vertical="top", wrap_text=True)

def clean(value):
    return value.strip() if isinstance(value, str) else value

def get_list(form, key):
    return [clean(v) for v in form.getlist(key)]

def unpack_pfl(pfl_list, index):
    """Helper to unpack packed P/F/L strings (e.g., '10/50/20') safely."""
    vals = pfl_list[index].split('/') if index < len(pfl_list) else ["", "", ""]
    vals += [""] * (3 - len(vals)) # Ensure exactly 3 elements
    return vals

def build_fi_section(form, prefix):
    """Section 1: Field Instruments (Process Data + Signal IO Data)"""
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
    # New Signal IO fields
    sig_types = get_list(form, f"{prefix}SignalType[]")
    sources = get_list(form, f"{prefix}Source[]")
    destinations = get_list(form, f"{prefix}Destination[]")
    signals = get_list(form, f"{prefix}Signal[]")

    rows = []
    for i in range(len(tags)):
        w_vals = unpack_pfl(working_vals, i)
        d_vals = unpack_pfl(design_vals, i)

        row = {
            "Tag No": tags[i] if i < len(tags) else "",
            "Instrument Name": instruments[i] if i < len(instruments) else "",
            "Service Description": services[i] if i < len(services) else "",
            "Line Size": line_sizes[i] if i < len(line_sizes) else "",
            "Medium": mediums[i] if i < len(mediums) else "",
            "Work Press": w_vals[0], "Work Flow": w_vals[1], "Work Level": w_vals[2],
            "Design Press": d_vals[0], "Design Flow": d_vals[1], "Design Level": d_vals[2],
            "Specification": specs[i] if i < len(specs) else "",
            "Process Conn": conns[i] if i < len(conns) else "",
            "Set-point": set_points[i] if i < len(set_points) else "",
            "Range": ranges[i] if i < len(ranges) else "",
            "UOM": uoms[i] if i < len(uoms) else "",
            "Signal Type": sig_types[i] if i < len(sig_types) else "",
            "Source": sources[i] if i < len(sources) else "",
            "Destination": destinations[i] if i < len(destinations) else "",
            "Signal": signals[i] if i < len(signals) else ""
        }
        if any(row.values()):
            rows.append(row)
    return rows

def build_el_section(form, prefix):
    """Section 2: Electrical (Base Data + Signal IO Data ONLY)"""
    tags = get_list(form, f"{prefix}TagNo[]")
    instruments = get_list(form, f"{prefix}Instrument[]")
    services = get_list(form, f"{prefix}ServiceDescription[]")
    sig_types = get_list(form, f"{prefix}SignalType[]")
    sources = get_list(form, f"{prefix}Source[]")
    destinations = get_list(form, f"{prefix}Destination[]")
    signals = get_list(form, f"{prefix}Signal[]")

    rows = []
    for i in range(len(tags)):
        row = {
            "Tag No": tags[i] if i < len(tags) else "",
            "Instrument Name": instruments[i] if i < len(instruments) else "",
            "Service Description": services[i] if i < len(services) else "",
            "Signal Type": sig_types[i] if i < len(sig_types) else "",
            "Source": sources[i] if i < len(sources) else "",
            "Destination": destinations[i] if i < len(destinations) else "",
            "Signal": signals[i] if i < len(signals) else ""
        }
        if any(row.values()):
            rows.append(row)
    return rows

def build_mov_section(form, prefix):
    """Section 3: MOVs (Base Data + Process Data ONLY)"""
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
    for i in range(len(tags)):
        w_vals = unpack_pfl(working_vals, i)
        d_vals = unpack_pfl(design_vals, i)

        row = {
            "Tag No": tags[i] if i < len(tags) else "",
            "Instrument Name": instruments[i] if i < len(instruments) else "",
            "Service Description": services[i] if i < len(services) else "",
            "Line Size": line_sizes[i] if i < len(line_sizes) else "",
            "Medium": mediums[i] if i < len(mediums) else "",
            "Work Press": w_vals[0], "Work Flow": w_vals[1], "Work Level": w_vals[2],
            "Design Press": d_vals[0], "Design Flow": d_vals[1], "Design Level": d_vals[2],
            "Specification": specs[i] if i < len(specs) else "",
            "Process Conn": conns[i] if i < len(conns) else "",
            "Set-point": set_points[i] if i < len(set_points) else "",
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

@app.route("/preview", methods=["POST"])
def preview():
    try:
        # Build the exact same payload you do for submission
        payload = {
            "header": {k: request.form.get(k, "") for k in ["projectName", "documentName", "client", "documentNumber"]},
            "field_instruments": build_fi_section(request.form, "fi"),
            "electrical": build_el_section(request.form, "el"),
            "mov": build_mov_section(request.form, "mov"),
        }
        
        # Run it through your validator
        errors = validate_payload(payload)
        
        if errors:
            # If there are errors, send them back to the frontend to alert the user
            return jsonify({"success": False, "errors": errors}), 400
            
        # If no errors, let the frontend know the data is clean
        return jsonify({"success": True, "message": "Validation passed! No errors found."}), 200

    except Exception as e:
        traceback.print_exc()
        return jsonify({"success": False, "message": str(e)}), 500

@app.route("/submit", methods=["POST"])
def submit():
    try:
        payload = {
            "header": {k: request.form.get(k, "") for k in ["projectName", "documentName", "client", "documentNumber"]},
            "field_instruments": build_fi_section(request.form, "fi"),
            "electrical": build_el_section(request.form, "el"),
            "mov": build_mov_section(request.form, "mov"),
        }
        
        output = BytesIO()
        wb = Workbook()
        
        # Write the three distinct sheets
        write_sheet(wb.active, payload["field_instruments"], "Field Instruments")
        write_sheet(wb.create_sheet("Electrical"), payload["electrical"], "Electrical")
        write_sheet(wb.create_sheet("MOV"), payload["mov"], "MOV")
        
        wb.save(output)
        output.seek(0)
        
        return send_file(output, as_attachment=True, download_name=f"Instrument_List_{datetime.now().strftime('%Y%m%d')}.xlsx")
    except Exception as e:
        traceback.print_exc() # Useful for terminal debugging
        return jsonify({"success": False, "error": str(e)}), 500

@app.route("/")
def index(): 
    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True)