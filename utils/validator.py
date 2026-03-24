"""
validator.py
────────────
Validates the flat payload produced by app.py.

Key rules
─────────
• Header: projectName, documentName, documentNumber are required.
• Section 1 (Field Instruments): every row must have a Tag No; the same
  tag may NOT appear more than once (each FI row is a unique instrument).
• Sections 2 & 3 (Electrical, MOV): every row must have a Tag No; the
  same tag CAN repeat within a section (multiple signals per tag is
  intentional). However, a tag must NOT appear in more than one section
  (e.g. a tag cannot be both Electrical and MOV).
• Signal column (all sections): if present, must be one of AI/AO/DI/DO.
• Signal Type column (all sections): if present, must be one of the
  recognised signal type strings.
"""

from typing import Any, Dict, List, Set

VALID_SIGNALS      = {"AI", "AO", "DI", "DO"}
VALID_SIGNAL_TYPES = {"4-20mA", "Potential Free Contact", "24V DC"}


def validate_payload(payload: Dict[str, Any]) -> List[str]:
    errors: List[str] = []

    # ── 1. Header ────────────────────────────────────────────────────────────
    header = payload.get("header", {})
    for field, label in [
        ("projectName",    "Project Name"),
        ("documentName",   "Document Name"),
        ("documentNumber", "Document Number"),
    ]:
        if not header.get(field, "").strip():
            errors.append(f"Header: {label} is required.")

    # ── 2. Section 1 — Field Instruments ────────────────────────────────────
    # Each row is a unique instrument → duplicates within this section are errors.
    fi_tags: Set[str] = set()
    for idx, row in enumerate(payload.get("field_instruments", []), start=1):
        tag    = row.get("Tag No", "").strip()
        signal = row.get("Signal", "").strip()
        sig_t  = row.get("Signal Type", "").strip()

        if not tag:
            errors.append(f"Section 1 Row {idx}: Tag No is required.")
        elif tag in fi_tags:
            errors.append(
                f"Section 1 Row {idx}: Duplicate Tag No '{tag}' — "
                f"each Field Instrument must be unique."
            )
        else:
            fi_tags.add(tag)

        if signal and signal not in VALID_SIGNALS:
            errors.append(
                f"Section 1 Tag '{tag or idx}': "
                f"Invalid Signal '{signal}'. Must be one of {sorted(VALID_SIGNALS)}."
            )
        if sig_t and sig_t not in VALID_SIGNAL_TYPES:
            errors.append(
                f"Section 1 Tag '{tag or idx}': "
                f"Invalid Signal Type '{sig_t}'."
            )

    # ── 3. Sections 2 & 3 — Electrical / MOV ────────────────────────────────
    # A tag can repeat within its own section (multiple signals).
    # A tag must not appear in BOTH sections.

    def _validate_grouped_section(
        rows: List[Dict],
        section_label: str,
        forbidden_tags: Set[str],     # tags already claimed by the other section
    ) -> Set[str]:
        """
        Validate one flat section. Returns the set of unique tags found
        (so the caller can pass it as forbidden_tags to the sibling section).
        """
        seen_tags: Set[str] = set()

        for idx, row in enumerate(rows, start=1):
            tag    = row.get("Tag No", "").strip()
            signal = row.get("Signal", "").strip()
            sig_t  = row.get("Signal Type", "").strip()

            if not tag:
                errors.append(f"{section_label} Row {idx}: Tag No is required.")
                continue

            if tag in forbidden_tags:
                errors.append(
                    f"{section_label} Row {idx}: Tag '{tag}' already exists "
                    f"in a different section."
                )

            seen_tags.add(tag)

            if signal and signal not in VALID_SIGNALS:
                errors.append(
                    f"{section_label} Tag '{tag}': "
                    f"Invalid Signal '{signal}'. Must be one of {sorted(VALID_SIGNALS)}."
                )
            if sig_t and sig_t not in VALID_SIGNAL_TYPES:
                errors.append(
                    f"{section_label} Tag '{tag}': "
                    f"Invalid Signal Type '{sig_t}'."
                )

        return seen_tags

    el_tags  = _validate_grouped_section(
        payload.get("electrical", []), "Section 2 (Electrical)", forbidden_tags=set()
    )
    _validate_grouped_section(
        payload.get("mov", []),        "Section 3 (MOV)",        forbidden_tags=el_tags
    )

    return errors
