# schemas/payload.py
"""
Marshmallow schemas for inbound JSON payload validation.

These replace the ad-hoc dict access scattered across app.py.
Every route that receives a payload calls PayloadSchema().load(raw)
and works exclusively with the clean, typed result.

Design decisions
────────────────
- load_default=''  on string fields means missing keys become ''
  rather than raising, which preserves the existing "empty is ok"
  behaviour for optional fields.
- VALID_SIGNALS / VALID_SIGNAL_TYPES are enforced here so that
  validator.py can stay focused on *cross-row* business rules
  (duplicate tags, cross-section conflicts) rather than field types.
- Row schemas use data_key so the internal Python name is snake_case
  while the JSON key stays exactly as the frontend sends it
  (e.g. "Tag No" → tag_no).
"""

from marshmallow import (
    Schema, fields, validates, validates_schema,
    ValidationError, pre_load, post_load, EXCLUDE,
)

# ─── Constants ────────────────────────────────────────────────────────────────

VALID_SIGNALS      = {"AI", "AO", "DI", "DO"}
VALID_SIGNAL_TYPES = {"4-20mA", "Potential Free Contact", "24V DC"}


# ─── Helpers ──────────────────────────────────────────────────────────────────

def _strip_str(value):
    """Strip whitespace from string values before field validation."""
    return value.strip() if isinstance(value, str) else value


class _BaseRowSchema(Schema):
    """Common behaviour for all instrument row schemas."""

    class Meta:
        unknown = EXCLUDE          # ignore extra keys from the frontend

    @pre_load
    def strip_strings(self, data, **kwargs):
        return {k: _strip_str(v) for k, v in data.items()}


# ─── Sub-schemas ──────────────────────────────────────────────────────────────

class HeaderSchema(Schema):
    """Global project / revision metadata block."""

    class Meta:
        unknown = EXCLUDE

    projectName = fields.Str(load_default="")
    client      = fields.Str(load_default="")
    consultant  = fields.Str(load_default="")
    location    = fields.Str(load_default="")
    date        = fields.Str(load_default="")
    preparedBy  = fields.Str(load_default="")
    checkedBy   = fields.Str(load_default="")
    approvedBy  = fields.Str(load_default="")
    revision    = fields.Str(load_default="")

    @validates("projectName")
    def require_project_name(self, value):
        if not value.strip():
            raise ValidationError("Project Name is required.")

    @pre_load
    def strip_strings(self, data, **kwargs):
        return {k: _strip_str(v) for k, v in data.items()}

    @post_load
    def ensure_strings(self, data, **kwargs):
        """Guarantee all values are strings (never None)."""
        return {k: (v or "") for k, v in data.items()}


class MetaSchema(Schema):
    """Per-document metadata for IL, IO, EL, MOV documents."""

    class Meta:
        unknown = EXCLUDE

    docNumber = fields.Str(load_default="")

    @pre_load
    def coerce_none(self, data, **kwargs):
        if data is None:
            return {}
        return {k: _strip_str(v) for k, v in data.items()}


class CsMetaSchema(Schema):
    """Per-document metadata for the Cable Schedule — cable specs and AJB config."""

    class Meta:
        unknown = EXCLUDE

    docNumber    = fields.Str(load_default="")
    analogCable  = fields.Str(load_default="")   # e.g. "2P × 1.0 sq.mm"
    digitalCable = fields.Str(load_default="")   # e.g. "10C × 1.0 sq.mm"
    ajbCapacity  = fields.Int(load_default=16)   # 4, 8, or 16 — defaults to 16-way

    @pre_load
    def coerce_none(self, data, **kwargs):
        if data is None:
            return {}
        out = {}
        for k, v in data.items():
            if k == "ajbCapacity":
                try:
                    capacity = int(v) if v else 16
                except (ValueError, TypeError):
                    capacity = 16
                out[k] = capacity if capacity in {4, 8, 16} else 16
            else:
                out[k] = _strip_str(v) if isinstance(v, str) else (v or "")
        return out


class FieldInstrumentRowSchema(_BaseRowSchema):
    """One row in the Field Instruments grid (Section 1)."""

    tag_no          = fields.Str(data_key="Tag No",              load_default="")
    instrument_name = fields.Str(data_key="Instrument Name",     load_default="")
    service_desc    = fields.Str(data_key="Service Description", load_default="")
    line_size       = fields.Str(data_key="Line Size",           load_default="")
    medium          = fields.Str(data_key="Medium",              load_default="")
    specification   = fields.Str(data_key="Specification",       load_default="")
    process_conn    = fields.Str(data_key="Process Conn",        load_default="")
    work_press      = fields.Str(data_key="Work Press",          load_default="")
    work_flow       = fields.Str(data_key="Work Flow",           load_default="")
    work_level      = fields.Str(data_key="Work Level",          load_default="")
    design_press    = fields.Str(data_key="Design Press",        load_default="")
    design_flow     = fields.Str(data_key="Design Flow",         load_default="")
    design_level    = fields.Str(data_key="Design Level",        load_default="")
    set_point       = fields.Str(data_key="Set-point",           load_default="")
    range_          = fields.Str(data_key="Range",               load_default="")
    uom             = fields.Str(data_key="UOM",                 load_default="")
    signal_type     = fields.Str(data_key="Signal Type",         load_default="")
    source          = fields.Str(data_key="Source",              load_default="")
    destination     = fields.Str(data_key="Destination",         load_default="")
    signal          = fields.Str(data_key="Signal",              load_default="")

    @validates("signal")
    def validate_signal(self, value):
        if value and value not in VALID_SIGNALS:
            raise ValidationError(
                f"Invalid Signal '{value}'. Must be one of {sorted(VALID_SIGNALS)}."
            )

    @validates("signal_type")
    def validate_signal_type(self, value):
        if value and value not in VALID_SIGNAL_TYPES:
            raise ValidationError(
                f"Invalid Signal Type '{value}'. "
                f"Must be one of {sorted(VALID_SIGNAL_TYPES)}."
            )

    @post_load
    def rebuild_original_keys(self, data, **kwargs):
        """
        Re-map snake_case back to the exact keys that excel_writer.py
        and validator.py expect ("Tag No", "Signal Type", …).
        The rest of the pipeline was written against those keys — we keep
        them so we don't have to rewrite excel_writer.
        """
        return {
            "Tag No":              data.get("tag_no",          ""),
            "Instrument Name":     data.get("instrument_name", ""),
            "Service Description": data.get("service_desc",    ""),
            "Line Size":           data.get("line_size",       ""),
            "Medium":              data.get("medium",          ""),
            "Specification":       data.get("specification",   ""),
            "Process Conn":        data.get("process_conn",    ""),
            "Work Press":          data.get("work_press",      ""),
            "Work Flow":           data.get("work_flow",       ""),
            "Work Level":          data.get("work_level",      ""),
            "Design Press":        data.get("design_press",    ""),
            "Design Flow":         data.get("design_flow",     ""),
            "Design Level":        data.get("design_level",    ""),
            "Set-point":           data.get("set_point",       ""),
            "Range":               data.get("range_",          ""),
            "UOM":                 data.get("uom",             ""),
            "Signal Type":         data.get("signal_type",     ""),
            "Source":              data.get("source",          ""),
            "Destination":         data.get("destination",     ""),
            "Signal":              data.get("signal",          ""),
        }


class ElectricalRowSchema(_BaseRowSchema):
    """One row in the Electrical Equipment grid (Section 2)."""

    tag_no          = fields.Str(data_key="Tag No",              load_default="")
    instrument_name = fields.Str(data_key="Instrument Name",     load_default="")
    service_desc    = fields.Str(data_key="Service Description", load_default="")
    signal_type     = fields.Str(data_key="Signal Type",         load_default="")
    source          = fields.Str(data_key="Source",              load_default="")
    destination     = fields.Str(data_key="Destination",         load_default="")
    signal_desc     = fields.Str(data_key="Signal Description",  load_default="")
    signal          = fields.Str(data_key="Signal",              load_default="")

    @validates("signal")
    def validate_signal(self, value):
        if value and value not in VALID_SIGNALS:
            raise ValidationError(
                f"Invalid Signal '{value}'. Must be one of {sorted(VALID_SIGNALS)}."
            )

    @validates("signal_type")
    def validate_signal_type(self, value):
        if value and value not in VALID_SIGNAL_TYPES:
            raise ValidationError(
                f"Invalid Signal Type '{value}'."
            )

    @post_load
    def rebuild_original_keys(self, data, **kwargs):
        return {
            "Tag No":              data.get("tag_no",          ""),
            "Instrument Name":     data.get("instrument_name", ""),
            "Service Description": data.get("service_desc",    ""),
            "Signal Type":         data.get("signal_type",     ""),
            "Source":              data.get("source",          ""),
            "Destination":         data.get("destination",     ""),
            "Signal Description":  data.get("signal_desc",     ""),
            "Signal":              data.get("signal",          ""),
        }


# MOVs share the exact same shape as Electrical rows
MOVRowSchema = ElectricalRowSchema


# ─── Top-level payload schema ─────────────────────────────────────────────────

class PayloadSchema(Schema):
    """
    Full inbound payload schema.

    Usage
    ─────
        schema  = PayloadSchema()
        payload = schema.load(request.get_json(force=True, silent=True) or {})
        # payload is guaranteed clean — no missing keys, no None strings.

    Raises marshmallow.ValidationError on hard failures.
    Row-level errors are collected and re-raised as a single error
    with a structured {field: [errors]} dict.
    """

    class Meta:
        unknown = EXCLUDE

    header            = fields.Nested(HeaderSchema,  load_default=dict)
    fi_meta           = fields.Nested(MetaSchema,    load_default=dict)
    el_meta           = fields.Nested(MetaSchema,    load_default=dict)
    mov_meta          = fields.Nested(MetaSchema,    load_default=dict)
    io_meta           = fields.Nested(MetaSchema,    load_default=dict)
    cs_meta           = fields.Nested(CsMetaSchema,  load_default=dict)   # ← NEW
    field_instruments = fields.List(
        fields.Nested(FieldInstrumentRowSchema), load_default=list
    )
    electrical        = fields.List(
        fields.Nested(ElectricalRowSchema),      load_default=list
    )
    mov               = fields.List(
        fields.Nested(MOVRowSchema),             load_default=list
    )

    @pre_load
    def coerce_none_root(self, data, **kwargs):
        """Guard against a completely null body."""
        return data or {}

    @post_load
    def filter_empty_rows(self, data, **kwargs):
        """
        Drop rows where every field is blank.
        This mirrors the previous getData() filter in the JS layer and
        prevents empty grid rows from being written to the DB or Excel.
        """
        def _has_content(row: dict) -> bool:
            return any(v for v in row.values())

        data["field_instruments"] = [r for r in data["field_instruments"] if _has_content(r)]
        data["electrical"]        = [r for r in data["electrical"]        if _has_content(r)]
        data["mov"]               = [r for r in data["mov"]               if _has_content(r)]
        return data


# ─── Public convenience ────────────────────────────────────────────────────────

_schema = PayloadSchema()


def load_payload(raw: dict) -> dict:
    """
    Deserialise and validate a raw JSON body.

    Returns a clean payload dict on success.
    Raises marshmallow.ValidationError on failure — callers should
    catch this and return a 422 response.
    """
    return _schema.load(raw or {})
