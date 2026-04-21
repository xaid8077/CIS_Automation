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

_WRITERS: Dict[str, Callable[[dict, BytesIO], None]] = {
    "Instrument List": write_workbook,
    "IO List":         write_io_workbook,
}


# ─── Public entry point ───────────────────────────────────────────────────────

def generate(doc_type: str, payload: dict) -> BytesIO:
    """
    Generate an Excel workbook for the given document type.

    Parameters
    ----------
    doc_type : str
        One of the keys registered in _WRITERS
        (e.g. "Instrument List", "IO List").
    payload : dict
        Clean payload dict — output of PayloadSchema.load().
        Must contain at minimum the keys expected by the writer:
          "header", "fi_meta" / "io_meta", "field_instruments",
          "electrical", "mov".

    Returns
    -------
    BytesIO
        In-memory stream positioned at offset 0.
        Pass directly to Flask's send_file() — no seek() needed.

    Raises
    ------
    ValueError
        doc_type is not registered in _WRITERS.
    FileNotFoundError
        The underlying xlsx template is missing from templates/.
    Exception
        Any openpyxl error during workbook generation propagates as-is.
    """
    writer = _WRITERS.get(doc_type)
    if writer is None:
        raise ValueError(
            f"No Excel writer registered for document type '{doc_type}'. "
            f"Supported types: {sorted(_WRITERS)}."
        )

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