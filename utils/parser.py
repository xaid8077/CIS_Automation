from typing import Any, Dict, List
from werkzeug.datastructures import ImmutableMultiDict


def _clean(value: str) -> str:
    return value.strip() if isinstance(value, str) else value


def _get_scalar(form: ImmutableMultiDict, key: str) -> str:
    return _clean(form.get(key, ""))


def _get_list(form: ImmutableMultiDict, key: str) -> List[str]:
    return [_clean(v) for v in form.getlist(key)]


def _build_field_instruments(form: ImmutableMultiDict) -> List[Dict[str, str]]:
    tags = _get_list(form, "fiTagNo[]")
    instruments = _get_list(form, "fiInstrument[]")
    services = _get_list(form, "fiServiceDescription[]")
    io_types = _get_list(form, "fiIoType[]")

    rows: List[Dict[str, str]] = []
    max_len = max(len(tags), len(instruments), len(services), len(io_types), default=0)

    for i in range(max_len):
        row = {
            "tag_no": tags[i] if i < len(tags) else "",
            "instrument": instruments[i] if i < len(instruments) else "",
            "service_description": services[i] if i < len(services) else "",
            "io_type": io_types[i] if i < len(io_types) else "",
        }

        if any(row.values()):
            rows.append(row)

    return rows


def _build_parent_child_entities(
    form: ImmutableMultiDict,
    parent_tag_key: str,
    parent_instrument_key: str,
    parent_service_key: str,
    child_signal_key: str,
    child_io_key: str,
) -> List[Dict[str, Any]]:
    """
    Skeleton parser for parent-child sections.

    Important limitation:
    With current plain form-array submission, parent-child linkage is inferred by order.
    This is workable for version 1, but fragile for version 2.

    Safer long-term design:
    Give every parent row a unique ID in frontend and submit that with child rows.
    """
    parent_tags = _get_list(form, parent_tag_key)
    parent_instruments = _get_list(form, parent_instrument_key)
    parent_services = _get_list(form, parent_service_key)

    child_signals = _get_list(form, child_signal_key)
    child_ios = _get_list(form, child_io_key)

    parents: List[Dict[str, Any]] = []
    for idx in range(max(len(parent_tags), len(parent_instruments), len(parent_services), default=0)):
        parent = {
            "tag_no": parent_tags[idx] if idx < len(parent_tags) else "",
            "instrument": parent_instruments[idx] if idx < len(parent_instruments) else "",
            "service_description": parent_services[idx] if idx < len(parent_services) else "",
            "signals": [],
        }
        if parent["tag_no"] or parent["instrument"] or parent["service_description"]:
            parents.append(parent)

    """
    Current frontend does not explicitly submit mapping metadata for child rows.
    So here we use a simple sequential grouping assumption:
    - each parent has at least one child
    - children appear in order under each parent
    - frontend row order is preserved

    Because HTML arrays alone do not carry parent IDs, this is only a temporary skeleton.
    For now, we distribute child rows conservatively.
    """

    child_rows = []
    max_child_len = max(len(child_signals), len(child_ios), default=0)
    for idx in range(max_child_len):
        row = {
            "signal_description": child_signals[idx] if idx < len(child_signals) else "",
            "io_type": child_ios[idx] if idx < len(child_ios) else "",
        }
        if row["signal_description"] or row["io_type"]:
            child_rows.append(row)

    if not parents:
        return []

    # Temporary distribution rule:
    # Assign at least one child row to each parent first, then any remaining rows to the last parent.
    child_index = 0
    for parent in parents:
        if child_index < len(child_rows):
            parent["signals"].append(child_rows[child_index])
            child_index += 1

    while child_index < len(child_rows):
        parents[-1]["signals"].append(child_rows[child_index])
        child_index += 1

    return parents


def _build_cable_schedule(form: ImmutableMultiDict) -> List[Dict[str, str]]:
    cable_tags = _get_list(form, "cTag[]")
    instr_tags = _get_list(form, "cInstr[]")
    from_points = _get_list(form, "cFrom[]")
    to_points = _get_list(form, "cTo[]")
    from_terms = _get_list(form, "cFromTerm[]")
    to_terms = _get_list(form, "cToTerm[]")
    cable_types = _get_list(form, "cType[]")
    pairs = _get_list(form, "cPairs[]")
    used = _get_list(form, "cUsed[]")
    signals = _get_list(form, "cSignal[]")
    remarks = _get_list(form, "cRemarks[]")

    rows: List[Dict[str, str]] = []
    max_len = max(
        len(cable_tags),
        len(instr_tags),
        len(from_points),
        len(to_points),
        len(from_terms),
        len(to_terms),
        len(cable_types),
        len(pairs),
        len(used),
        len(signals),
        len(remarks),
        default=0,
    )

    for i in range(max_len):
        row = {
            "cable_tag": cable_tags[i] if i < len(cable_tags) else "",
            "instrument_tag": instr_tags[i] if i < len(instr_tags) else "",
            "from": from_points[i] if i < len(from_points) else "",
            "to": to_points[i] if i < len(to_points) else "",
            "from_terminal": from_terms[i] if i < len(from_terms) else "",
            "to_terminal": to_terms[i] if i < len(to_terms) else "",
            "cable_type": cable_types[i] if i < len(cable_types) else "",
            "pairs": pairs[i] if i < len(pairs) else "",
            "used": used[i] if i < len(used) else "",
            "signal": signals[i] if i < len(signals) else "",
            "remarks": remarks[i] if i < len(remarks) else "",
        }
        if any(row.values()):
            rows.append(row)

    return rows


def parse_form_data(form: ImmutableMultiDict) -> Dict[str, Any]:
    payload: Dict[str, Any] = {
        "header": {
            "project_name": _get_scalar(form, "projectName"),
            "document_name": _get_scalar(form, "documentName"),
            "client": _get_scalar(form, "client"),
            "consultant": _get_scalar(form, "consultant"),
            "document_number": _get_scalar(form, "documentNumber"),
            "date": _get_scalar(form, "date"),
            "prepared_by": _get_scalar(form, "preparedBy"),
            "checked_by": _get_scalar(form, "checkedBy"),
            "approved_by": _get_scalar(form, "approvedBy"),
            "revision": _get_scalar(form, "revision"),
        },
        "field_instruments": _build_field_instruments(form),
        "electrical_equipment": _build_parent_child_entities(
            form,
            "elTagNo[]",
            "elInstrument[]",
            "elServiceDescription[]",
            "elSignalDescription[]",
            "elIoType[]",
        ),
        "mov_equipment": _build_parent_child_entities(
            form,
            "movTagNo[]",
            "movInstrument[]",
            "movServiceDescription[]",
            "movSignalDescription[]",
            "movIoType[]",
        ),
        "cable_schedule": _build_cable_schedule(form),
    }

    return payload