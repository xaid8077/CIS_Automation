# services/document_service.py
"""
document_service.py
───────────────────
Orchestrates the full "save + generate file" workflow.

This is the only module that knows about both the revision layer
(revision_service) and the file-generation layer (excel_service).
Routes call exactly one function here and get back either a BytesIO
stream (for downloads) or a plain dict (for JSON responses).

Responsibilities
────────────────
- Coordinate revision_service + excel_service in the correct order
- Build the download filename from project/location/revision metadata
- Roll back the DB transaction if file generation fails (so a broken
  revision record is never left orphaned in the database)

Design notes
────────────
- No Flask objects (request, g, current_user) enter this module.
  All context the service needs is passed as plain arguments.
- Exceptions propagate to the caller (the route) which decides the
  HTTP response code and body.
"""

from datetime import datetime
from io import BytesIO
from typing import Tuple

from extensions import db
from models import Project, ProjectLocation
from services import revision_service
from services import excel_service


# ─── Supported document types ─────────────────────────────────────────────────

SUPPORTED_DOC_TYPES = {
    "Instrument List",
    "IO List",
    "Cable Schedule",   # ← NEW
}


# ─── Draft save ───────────────────────────────────────────────────────────────

def save_draft(
    project: Project,
    location: ProjectLocation,
    user_id: int,
    payload: dict,
) -> dict:
    """
    Persist payload as a draft revision (no file generated).

    Parameters
    ----------
    project:  Project ORM instance.
    location: ProjectLocation ORM instance.
    user_id:  ID of the currently logged-in user.
    payload:  Clean payload dict from PayloadSchema.load().

    Returns
    -------
    A plain dict suitable for jsonify():
        {"ok": True, "message": "..."}
    """
    revision_service.upsert_draft(
        project  = project,
        location = location,
        user_id  = user_id,
        payload  = payload,
    )
    return {"ok": True, "message": "Data saved successfully."}


# ─── Generate + save ──────────────────────────────────────────────────────────

def generate_and_save(
    project: Project,
    location: ProjectLocation,
    user_id: int,
    doc_type: str,
    payload: dict,
) -> Tuple[BytesIO, str]:
    """
    Create a published revision, generate the Excel file, and return
    the in-memory stream together with the recommended download filename.

    The DB write and file generation are intentionally ordered so that:
      1. The revision row is written first (gets an ID + revision_number).
      2. The Excel file is generated from the same payload.
      3. If Excel generation raises, the DB transaction is rolled back.

    This guarantees no orphaned revision record exists when file
    generation fails.

    Parameters
    ----------
    project:  Project ORM instance.
    location: ProjectLocation ORM instance.
    user_id:  ID of the currently logged-in user.
    doc_type: One of the keys in SUPPORTED_DOC_TYPES.
    payload:  Clean payload dict from PayloadSchema.load().

    Returns
    -------
    (stream, filename)
      stream   — BytesIO positioned at offset 0, ready for send_file().
      filename — Suggested attachment filename (e.g. "Plant_IL_Rev0_20250421.xlsx").

    Raises
    ------
    ValueError  — doc_type is not in SUPPORTED_DOC_TYPES.
    Exception   — any error from revision_service or excel_service;
                  the DB transaction is rolled back before re-raising.
    """
    if doc_type not in SUPPORTED_DOC_TYPES:
        raise ValueError(
            f"Unsupported document type '{doc_type}'. "
            f"Must be one of: {sorted(SUPPORTED_DOC_TYPES)}."
        )

    # ── 1. Persist the revision record ────────────────────────────────────────
    rev = revision_service.create_published_revision(
        project  = project,
        location = location,
        user_id  = user_id,
        doc_type = doc_type,
        payload  = payload,
    )

    # ── 2. Generate the Excel file ────────────────────────────────────────────
    try:
        stream = excel_service.generate(doc_type, payload)
    except Exception:
        # Roll back the revision row so the history table stays clean.
        db.session.rollback()
        raise

    # ── 3. Build the download filename ────────────────────────────────────────
    filename = _build_filename(project, location, doc_type, rev.revision_number)

    return stream, filename


# ─── Re-download from stored revision ────────────────────────────────────────

def regenerate_from_revision(
    project: Project,
    rev_id: int,
) -> Tuple[BytesIO, str]:
    """
    Re-generate an Excel file from a stored revision's payload.

    No DB write is performed — this is a pure read path.

    Parameters
    ----------
    project: Project ORM instance (used for the filename).
    rev_id:  Primary key of the DocumentRevision to regenerate.

    Returns
    -------
    (stream, filename) — same contract as generate_and_save().

    Raises
    ------
    404     — revision not found.
    ValueError — doc_type not supported for re-download.
    """
    from models import DocumentRevision

    rev = revision_service.get_revision_or_404(project.id, rev_id)

    if rev.doc_type not in SUPPORTED_DOC_TYPES:
        raise ValueError(
            f"Re-download not supported for document type '{rev.doc_type}'."
        )

    stream   = excel_service.generate(rev.doc_type, rev.data_payload)
    filename = _build_filename(
        project,
        rev.location,
        rev.doc_type,
        rev.revision_number,
        date_override=rev.created_at,
    )
    return stream, filename


# ─── Internal helpers ─────────────────────────────────────────────────────────

_DOC_TYPE_PREFIX = {
    "Instrument List": "Instrument_List",
    "IO List":         "IO_List",
}


def _loc_tag(location) -> str:
    """
    Return a filesystem-safe location tag for use in filenames.

    Prefers the location's short code; falls back to the first 12
    characters of the name with spaces replaced by underscores.
    Returns an empty string when location is None (legacy revisions).
    """
    if location is None:
        return ""
    if location.code and location.code.strip():
        return location.code.strip()
    return location.name[:12].replace(" ", "_")


def _build_filename(
    project,
    location,
    doc_type: str,
    rev_number: int,
    date_override=None,
) -> str:
    """
    Construct the recommended download filename.

    Format:
        {ProjectName}_{LocationTag}_{DocPrefix}_Rev{N}_{YYYYMMDD}.xlsx

    Examples:
        Chandrawal_WTP_PA-INT_Instrument_List_Rev2_20250421.xlsx
        Chandrawal_WTP_PA-INT_IO_List_Rev0_20250421.xlsx
    """
    date_str = (date_override or datetime.now()).strftime("%Y%m%d")
    prefix   = _DOC_TYPE_PREFIX.get(doc_type, doc_type.replace(" ", "_"))
    tag      = _loc_tag(location)
    tag_part = f"_{tag}" if tag else ""

    safe_project = project.display_name.replace(" ", "_")

    return f"{safe_project}{tag_part}_{prefix}_Rev{rev_number}_{date_str}.xlsx"