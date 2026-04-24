# services/excel_service.py
"""
excel_service.py
────────────────
Thin, stateless wrapper over utils/excel_writer.py.

Isolates the heavy XML-manipulation logic from the rest of the
service layer.  Routes and document_service call generate() here;
they never import excel_writer directly.

Adding a new document type:
  1. Write a writer function in utils/excel_writer.py.
  2. Add one entry to _WRITERS below.
  Nothing else changes.
"""

from io import BytesIO
from typing import Callable, Dict

from utils.excel_writer import write_instrument_list, write_io_workbook


# ─── Registry ─────────────────────────────────────────────────────────────────

_WRITERS: Dict[str, Callable[[dict], BytesIO]] = {
    "Instrument List": write_instrument_list,
    "IO List":         write_io_workbook,
}


# ─── Public entry point ───────────────────────────────────────────────────────

def generate(doc_type: str, payload: dict) -> BytesIO:
    """
    Generate an Excel workbook for the given doc_type.

    Parameters
    ----------
    doc_type : str
        Must be a key in _WRITERS (e.g. "Instrument List", "IO List").
    payload  : dict
        Clean payload dict produced by PayloadSchema.load().

    Returns
    -------
    BytesIO
        In-memory stream positioned at offset 0, ready for send_file().

    Raises
    ------
    ValueError
        If doc_type has no registered writer.
    Any exception from the writer propagates unchanged so the caller
    (document_service) can roll back the DB transaction before re-raising.
    """
    writer = _WRITERS.get(doc_type)

    if writer is None:
        raise ValueError(
            f"No Excel writer registered for document type '{doc_type}'. "
            f"Supported: {sorted(_WRITERS)}."
        )

    stream: BytesIO = writer(payload)

    # All writers must return a BytesIO; rewind just in case.
    stream.seek(0)
    return stream


def supported_types() -> list:
    """Return a sorted list of all registered document type strings."""
    return sorted(_WRITERS.keys())


def register_writer(doc_type: str, writer_fn: Callable) -> None:
    """
    Register an additional writer at runtime (tests / plugins).

    Parameters
    ----------
    doc_type  : str          — e.g. "Cable Schedule"
    writer_fn : callable     — fn(payload: dict) -> BytesIO
    """
    _WRITERS[doc_type] = writer_fn