from datetime import datetime
from io import BytesIO
import traceback

from flask import Flask, render_template, request, jsonify, send_file

from utils.excel_writer import write_workbook
from utils.validator import validate_payload

app = Flask(__name__)

def _clean(value):
    return value.strip() if isinstance(value, str) else value

def _get_list(form, key):
    return [_clean(v) for v in form.getlist(key)]

def _unpack_pfl(packed_list, index):
    raw = packed_list[index] if index < len(packed_list) else ""
    parts = raw.split("/")
    parts += [""] * (3 - len(parts))
    return parts[:3]


def _build_fi(form):
    tags         = _get_list(form, "fiTagNo[]")
    instruments  = _get_list(form, "fiInstrument[]")
    services     = _get_list(form, "fiServiceDescription[]")
    line_sizes   = _get_list(form, "fiLineSize[]")
    mediums      = _get_list(form, "fiMedium[]")
    specs        = _get_list(form, "fiTypeSpec[]")
    conns        = _get_list(form, "fiProcessConnection[]")
    working_vals = _get_list(form, "fiWorkingValues[]")
    design_vals  = _get_list(form, "fiDesignValues[]")
    set_points   = _get_list(form, "fiSetPoint[]")
    ranges       = _get_list(form, "fiInstrumentRange[]")
    uoms         = _get_list(form, "fiUom[]")
    sig_types    = _get_list(form, "fiSignalType[]")
    sources      = _get_list(form, "fiSource[]")
    destinations = _get_list(form, "fiDestination[]")
    signals      = _get_list(form, "fiSignal[]")

    rows = []
    for i in range(len(tags)):
        w = _unpack_pfl(working_vals, i)
        d = _unpack_pfl(design_vals, i)
        row = {
            "Tag No":              tags[i]         if i < len(tags)         else "",
            "Instrument Name":     instruments[i]  if i < len(instruments)  else "",
            "Service Description": services[i]     if i < len(services)     else "",
            "Line Size":           line_sizes[i]   if i < len(line_sizes)   else "",
            "Medium":              mediums[i]      if i < len(mediums)      else "",
            "Specification":       specs[i]        if i < len(specs)        else "",
            "Process Conn":        conns[i]        if i < len(conns)        else "",
            "Work Press":          w[0],
            "Work Flow":           w[1],
            "Work Level":          w[2],
            "Design Press":        d[0],
            "Design Flow":         d[1],
            "Design Level":        d[2],
            "Set-point":           set_points[i]   if i < len(set_points)   else "",
            "Range":               ranges[i]       if i < len(ranges)       else "",
            "UOM":                 uoms[i]         if i < len(uoms)         else "",
            "Signal Type":         sig_types[i]    if i < len(sig_types)    else "",
            "Source":              sources[i]      if i < len(sources)      else "",
            "Destination":         destinations[i] if i < len(destinations) else "",
            "Signal":              signals[i]      if i < len(signals)      else "",
        }
        if any(row.values()):
            rows.append(row)
    return rows


def _build_flat(form, prefix):
    tags         = _get_list(form, f"{prefix}TagNo[]")
    instruments  = _get_list(form, f"{prefix}Instrument[]")
    services     = _get_list(form, f"{prefix}ServiceDescription[]")
    sig_types    = _get_list(form, f"{prefix}SignalType[]")
    sources      = _get_list(form, f"{prefix}Source[]")
    destinations = _get_list(form, f"{prefix}Destination[]")
    sig_descs    = _get_list(form, f"{prefix}SigDesc[]")
    signals      = _get_list(form, f"{prefix}Signal[]")

    rows = []
    for i in range(len(tags)):
        row = {
            "Tag No":              tags[i]         if i < len(tags)         else "",
            "Instrument Name":     instruments[i]  if i < len(instruments)  else "",
            "Service Description": services[i]     if i < len(services)     else "",
            "Signal Type":         sig_types[i]    if i < len(sig_types)    else "",
            "Source":              sources[i]      if i < len(sources)      else "",
            "Destination":         destinations[i] if i < len(destinations) else "",
            "Signal Description":  sig_descs[i]   if i < len(sig_descs)    else "",
            "Signal":              signals[i]      if i < len(signals)      else "",
        }
        if any(row.values()):
            rows.append(row)
    return rows


def _build_header(form):
    return {k: form.get(k, "") for k in [
        "projectName", "client", "consultant", "documentNumber",
        "date", "preparedBy", "checkedBy", "approvedBy", "revision"
    ]}


def _build_section_meta(form, prefix):
    """Per-section document name and number."""
    return {
        "docName":   _clean(form.get(f"{prefix}DocName",   "")),
        "docNumber": _clean(form.get(f"{prefix}DocNumber", "")),
    }


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/preview", methods=["POST"])
def preview():
    try:
        payload = {
            "header":            _build_header(request.form),
            "fi_meta":           _build_section_meta(request.form, "fi"),
            "el_meta":           _build_section_meta(request.form, "el"),
            "mov_meta":          _build_section_meta(request.form, "mov"),
            "field_instruments": _build_fi(request.form),
            "electrical":        _build_flat(request.form, "el"),
            "mov":               _build_flat(request.form, "mov"),
        }
        errors = validate_payload(payload)
        if errors:
            return jsonify({"success": False, "errors": errors}), 400
        return jsonify({"success": True, "message": "Validation passed — no errors found."}), 200
    except Exception:
        traceback.print_exc()
        return jsonify({"success": False, "message": "Server error during validation."}), 500


@app.route("/submit", methods=["POST"])
def submit():
    try:
        payload = {
            "header":            _build_header(request.form),
            "fi_meta":           _build_section_meta(request.form, "fi"),
            "el_meta":           _build_section_meta(request.form, "el"),
            "mov_meta":          _build_section_meta(request.form, "mov"),
            "field_instruments": _build_fi(request.form),
            "electrical":        _build_flat(request.form, "el"),
            "mov":               _build_flat(request.form, "mov"),
        }
        output = BytesIO()
        write_workbook(payload, output)
        output.seek(0)
        filename = f"Instrument_List_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        return send_file(
            output, as_attachment=True, download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception:
        traceback.print_exc()
        return jsonify({"success": False, "error": "Server error during workbook generation."}), 500


if __name__ == "__main__":
    app.run(debug=True)