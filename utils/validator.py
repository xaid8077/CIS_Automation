from typing import Any, Dict, List, Set


VALID_IO_TYPES = {"AI", "AO", "DI", "DO"}


def validate_payload(payload: Dict[str, Any]) -> List[str]:
    errors: List[str] = []

    header = payload.get("header", {})
    if not header.get("project_name"):
        errors.append("Project Name is required.")
    if not header.get("document_name"):
        errors.append("Document Name is required.")
    if not header.get("document_number"):
        errors.append("Document Number is required.")

    tag_set: Set[str] = set()

    for idx, row in enumerate(payload.get("field_instruments", []), start=1):
        tag = row.get("tag_no", "")
        io_type = row.get("io_type", "")
        if not tag:
            errors.append(f"Field Instrument row {idx}: Tag No is required.")
        elif tag in tag_set:
            errors.append(f"Duplicate Tag No found: {tag}")
        else:
            tag_set.add(tag)

        if io_type and io_type not in VALID_IO_TYPES:
            errors.append(f"Field Instrument row {idx}: Invalid I/O type '{io_type}'.")

    for section_name in ("electrical_equipment", "mov_equipment"):
        for idx, equipment in enumerate(payload.get(section_name, []), start=1):
            tag = equipment.get("tag_no", "")
            if not tag:
                errors.append(f"{section_name} row {idx}: Tag No is required.")
            elif tag in tag_set:
                errors.append(f"Duplicate Tag No found: {tag}")
            else:
                tag_set.add(tag)

            signals = equipment.get("signals", [])
            if not signals:
                errors.append(f"{section_name} row {idx}: At least one signal is required.")

            for s_idx, signal in enumerate(signals, start=1):
                io_type = signal.get("io_type", "")
                if io_type and io_type not in VALID_IO_TYPES:
                    errors.append(
                        f"{section_name} row {idx}, signal {s_idx}: Invalid I/O type '{io_type}'."
                    )

    for idx, row in enumerate(payload.get("cable_schedule", []), start=1):
        cable_tag = row.get("cable_tag", "")
        instr_tag = row.get("instrument_tag", "")
        if not cable_tag:
            errors.append(f"Cable Schedule row {idx}: Cable Tag is required.")
        if instr_tag and instr_tag not in tag_set:
            errors.append(
                f"Cable Schedule row {idx}: Instrument Tag '{instr_tag}' does not exist in Stage 1."
            )

    return errors