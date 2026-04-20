# app/services/document_service.py

from app.utils.validator import validate_payload


# ─────────────────────────────────────────────────────────────────────────────
# Payload Builder
# ─────────────────────────────────────────────────────────────────────────────

def build_payload(data: dict) -> dict:
    hdr = data.get("header") or {}

    def _s(d, k):
        return (d.get(k) or "").strip() if isinstance(d, dict) else ""

    payload = {
        "header": {k: _s(hdr, k) for k in [
            "projectName", "client", "consultant", "location",
            "date", "preparedBy", "checkedBy", "approvedBy", "revision",
        ]},
        "fi_meta":  {"docNumber": _s(data.get("fi_meta")  or {}, "docNumber")},
        "el_meta":  {"docNumber": _s(data.get("el_meta")  or {}, "docNumber")},
        "mov_meta": {"docNumber": _s(data.get("mov_meta") or {}, "docNumber")},
        "io_meta":  {"docNumber": _s(data.get("io_meta")  or {}, "docNumber")},
        "field_instruments": data.get("field_instruments") or [],
        "electrical":        data.get("electrical") or [],
        "mov":               data.get("mov") or [],
    }

    return payload


# ─────────────────────────────────────────────────────────────────────────────
# Validation
# ─────────────────────────────────────────────────────────────────────────────

def validate(payload, require_doc_numbers=False):
    return validate_payload(payload, require_doc_numbers=require_doc_numbers)