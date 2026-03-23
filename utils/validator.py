from typing import Any, Dict, List, Set

VALID_IO_TYPES = {"AI", "AO", "DI", "DO"}

def validate_payload(payload: Dict[str, Any]) -> List[str]:
    errors: List[str] = []

    # --- 1. Validate Header ---
    header = payload.get("header", {})
    if not header.get("projectName"):
        errors.append("Header: Project Name is required.")
    if not header.get("documentName"):
        errors.append("Header: Document Name is required.")
    # Assuming documentNumber is also mandatory based on your original logic
    if not header.get("documentNumber"):
        errors.append("Header: Document Number is required.")

    # Track tags globally across all sections to prevent duplicates
    tag_set: Set[str] = set()

    # --- 2. Validate Field Instruments (Section 1) ---
    for idx, row in enumerate(payload.get("field_instruments", []), start=1):
        tag = row.get("Tag No", "").strip()
        signal = row.get("Signal", "").strip()
        
        if not tag:
            errors.append(f"Section 1 (Row {idx}): Tag No is required.")
        elif tag in tag_set:
            errors.append(f"Duplicate Tag No found: {tag}")
        else:
            tag_set.add(tag)

        if signal and signal not in VALID_IO_TYPES:
            errors.append(f"Section 1 (Tag '{tag}'): Invalid Signal '{signal}'. Expected one of {VALID_IO_TYPES}.")

    # --- 3. Validate Electrical (Section 2) ---
    for idx, row in enumerate(payload.get("electrical", []), start=1):
        tag = row.get("Tag No", "").strip()
        signal = row.get("Signal", "").strip()

        if not tag:
            errors.append(f"Section 2 (Row {idx}): Tag No is required.")
        elif tag in tag_set:
            errors.append(f"Duplicate Tag No found: {tag}")
        else:
            tag_set.add(tag)

        if signal and signal not in VALID_IO_TYPES:
            errors.append(f"Section 2 (Tag '{tag}'): Invalid Signal '{signal}'. Expected one of {VALID_IO_TYPES}.")

    # --- 4. Validate MOV (Section 3) ---
    for idx, row in enumerate(payload.get("mov", []), start=1):
        tag = row.get("Tag No", "").strip()
        # Note: Your current frontend MOV section does not have an I/O signal dropdown

        if not tag:
            errors.append(f"Section 3 (Row {idx}): Tag No is required.")
        elif tag in tag_set:
            errors.append(f"Duplicate Tag No found: {tag}")
        else:
            tag_set.add(tag)

    return errors