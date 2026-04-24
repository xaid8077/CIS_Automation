# services/excel_service.py
"""
excel_service.py
────────────────
Thin, stateless wrapper over utils/excel_writer.py.

Why a wrapper?
──────────────
Routes and document_service must never import excel_writer directly.
This indirection means:
  - The heavy openpyxl logic stays isolated in utils/excel_writer.py.
  - If we ever swap the writer (e.g. xlsxwriter, async Celery task),
    only this file changes — document_service is untouched.
  - Unit tests can mock excel_service.generate() without needing a
    real xlsx template on disk.

Public API
──────────
    stream = excel_service.generate(doc_type, payload)
    # Returns a BytesIO at offset 0, ready for Flask's send_file().
"""

from io import BytesIO
from typing import Callable, Dict

from utils.excel_writer import write_workbook, write_io_workbook


# ─── Registry ─────────────────────────────────────────────────────────────────
# Maps doc_type string → writer function.
# Adding a new document type = adding one entry here.

_WRITERS: Dict[str, Callable[[dict], BytesIO]] = {
    "Instrument List": write_instrument_list,
    "IO List": write_io_workbook,
}


# ─── Public entry point ───────────────────────────────────────────────────────

def generate(doc_type: str, payload: dict) -> BytesIO:
    writer = _WRITERS.get(doc_type)

    if writer is None:
        raise ValueError(
            f"No Excel writer registered for document type '{doc_type}'. "
            f"Supported types: {sorted(_WRITERS)}."
        )

    result = writer(payload)

    # Backward compatibility (handles both styles)
    if isinstance(result, BytesIO):
        result.seek(0)
        return result

    # fallback for legacy writers (if any left)
    output = BytesIO()
    writer(payload, output)
    output.seek(0)
    return output


def supported_types() -> list:
    """Return a sorted list of all registered document type strings."""
    return sorted(_WRITERS.keys())


def register_writer(doc_type: str, writer_fn: Callable) -> None:
    """
    Register a new writer at runtime (e.g. from a plugin or test).

    Parameters
    ----------
    doc_type  : str   — The document type key (e.g. "Cable Schedule").
    writer_fn : callable(payload: dict, destination: BytesIO) -> None
    """
    _WRITERS[doc_type] = writer_fn