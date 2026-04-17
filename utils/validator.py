from typing import Any, Dict, List, Set

VALID_SIGNALS      = {"AI", "AO", "DI", "DO"}
VALID_SIGNAL_TYPES = {"4-20mA", "Potential Free Contact", "24V DC"}


def validate_payload(payload: Dict[str, Any]) -> List[str]:
    errors: List[str] = []

    # ── Header ────────────────────────────────────────────────────────────────
    header = payload.get("header", {})
    if not header.get("projectName", "").strip():
        errors.append("Header: Project Name is required.")

    # ── Per-section doc number required ───────────────────────────────────────
    for prefix, label in [
        ("fi",  "Instrument List (Section 1)"),
        ("el",  "Instrument List (Section 2)"),
        ("mov", "Instrument List (Section 3)"),
        ("io",  "IO List"),
    ]:
        meta = payload.get(f"{prefix}_meta", {})
        if not meta.get("docNumber", "").strip():
            errors.append(f"{label}: Document Number is required.")

    # ── Section 1 — Field Instruments ────────────────────────────────────────
    fi_tags: Set[str] = set()
    for idx, row in enumerate(payload.get("field_instruments", []), start=1):
        tag    = row.get("Tag No", "").strip()
        signal = row.get("Signal", "").strip()
        sig_t  = row.get("Signal Type", "").strip()

        if not tag:
            errors.append(f"Section 1 Row {idx}: Tag No is required.")
        elif tag in fi_tags:
            errors.append(f"Section 1 Row {idx}: Duplicate Tag No '{tag}'.")
        else:
            fi_tags.add(tag)

        if signal and signal not in VALID_SIGNALS:
            errors.append(f"Section 1 Tag '{tag or idx}': Invalid Signal '{signal}'. Must be one of {sorted(VALID_SIGNALS)}.")
        if sig_t and sig_t not in VALID_SIGNAL_TYPES:
            errors.append(f"Section 1 Tag '{tag or idx}': Invalid Signal Type '{sig_t}'.")

    # ── Sections 2 & 3 — grouped (tag may repeat within section) ─────────────
    def _validate_grouped(rows, label, forbidden_tags):
        seen: Set[str] = set()
        for idx, row in enumerate(rows, start=1):
            tag    = row.get("Tag No", "").strip()
            signal = row.get("Signal", "").strip()
            sig_t  = row.get("Signal Type", "").strip()

            if not tag:
                errors.append(f"{label} Row {idx}: Tag No is required.")
                continue
            if tag in forbidden_tags:
                errors.append(f"{label} Row {idx}: Tag '{tag}' already used in another section.")
            seen.add(tag)

            if signal and signal not in VALID_SIGNALS:
                errors.append(f"{label} Tag '{tag}': Invalid Signal '{signal}'. Must be one of {sorted(VALID_SIGNALS)}.")
            if sig_t and sig_t not in VALID_SIGNAL_TYPES:
                errors.append(f"{label} Tag '{tag}': Invalid Signal Type '{sig_t}'.")
        return seen

    el_tags = _validate_grouped(payload.get("electrical", []), "Section 2 (Electrical)", set())
    _validate_grouped(payload.get("mov", []), "Section 3 (MOV)", el_tags)

    return errors