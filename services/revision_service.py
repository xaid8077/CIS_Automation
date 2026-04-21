# services/revision_service.py
"""
revision_service.py
───────────────────
All database operations for DocumentRevision live here.

Routes never touch DocumentRevision directly — they call these
functions and get back plain model instances or dicts.

Responsibilities
────────────────
- Create draft revisions (upsert — one per project+location)
- Create published revisions (incrementing revision number)
- Fetch the latest revision for pre-filling the editor
- Patch document numbers on an existing revision (admin action)
- Re-fetch a revision for re-download

Design notes
────────────
- Every function accepts an explicit db_session argument so that
  unit tests can inject a test session without touching the app context.
  In practice, routes call the module-level helpers that use the
  imported `db` session from extensions.py.
- No Flask request objects enter this module — input is always a
  clean Python dict (already validated by the schema layer).
- Exceptions are not caught here; callers decide how to handle them.
"""

from datetime import datetime, timezone
from typing import Optional

from extensions import db
from models import DocumentRevision, Project, ProjectLocation


# ─── Draft management ─────────────────────────────────────────────────────────

def upsert_draft(
    project: Project,
    location: ProjectLocation,
    user_id: int,
    payload: dict,
    *,
    session=None,
) -> DocumentRevision:
    """
    Save payload as a draft revision.

    Only one draft is kept per project+location.  If one already exists
    it is updated in-place; otherwise a new draft is created.
    The draft is committed before returning.

    Parameters
    ----------
    project:  Project ORM instance.
    location: ProjectLocation ORM instance.
    user_id:  ID of the currently logged-in user.
    payload:  Clean payload dict (output of PayloadSchema.load).
    session:  SQLAlchemy session — defaults to the extension's db.session.

    Returns
    -------
    The saved (or updated) DocumentRevision instance.
    """
    sess = session or db.session

    existing = (
        sess.query(DocumentRevision)
        .filter_by(
            project_id  = project.id,
            location_id = location.id,
            status      = "draft",
        )
        .first()
    )

    if existing:
        existing.data_payload = payload
        existing.user_id      = user_id
        existing.created_at   = datetime.now(timezone.utc)
        draft = existing
    else:
        draft = DocumentRevision(
            project_id      = project.id,
            user_id         = user_id,
            location_id     = location.id,
            doc_type        = "Draft",
            revision_number = 0,
            data_payload    = payload,
            status          = "draft",
        )
        sess.add(draft)

    sess.commit()
    return draft


# ─── Published revision management ───────────────────────────────────────────

def create_published_revision(
    project: Project,
    location: ProjectLocation,
    user_id: int,
    doc_type: str,
    payload: dict,
    *,
    session=None,
) -> DocumentRevision:
    """
    Create a new published revision with an auto-incremented revision number.

    Revision numbers are scoped to (project, location, doc_type) so that
    Instrument List Rev 0 and IO List Rev 0 can coexist independently.

    The new revision is committed before returning so the caller can
    immediately use its .revision_number for the download filename.

    Parameters
    ----------
    project:  Project ORM instance.
    location: ProjectLocation ORM instance.
    user_id:  ID of the currently logged-in user.
    doc_type: One of "Instrument List", "IO List" (validated by the route).
    payload:  Clean payload dict (output of PayloadSchema.load).
    session:  SQLAlchemy session — defaults to the extension's db.session.

    Returns
    -------
    The newly created and committed DocumentRevision instance.
    """
    sess = session or db.session

    latest = (
        sess.query(DocumentRevision)
        .filter_by(
            project_id  = project.id,
            location_id = location.id,
            doc_type    = doc_type,
            status      = "published",
        )
        .order_by(DocumentRevision.revision_number.desc())
        .first()
    )

    next_rev_num = (latest.revision_number + 1) if latest else 0

    rev = DocumentRevision(
        project_id      = project.id,
        user_id         = user_id,
        location_id     = location.id,
        doc_type        = doc_type,
        revision_number = next_rev_num,
        data_payload    = payload,
        status          = "published",
    )
    sess.add(rev)
    sess.commit()
    return rev


# ─── Fetching helpers ─────────────────────────────────────────────────────────

def get_latest_for_editor(
    project_id: int,
    location_id: int,
    *,
    session=None,
) -> Optional[DocumentRevision]:
    """
    Return the most recent revision to pre-fill the editor.

    Priority order:
      1. Unsaved draft — reflects the most recent work.
      2. Latest published revision — fallback if no draft exists.

    Returns None when no revision exists for this project+location.
    """
    sess = session or db.session

    # Prefer draft
    draft = (
        sess.query(DocumentRevision)
        .filter_by(
            project_id  = project_id,
            location_id = location_id,
            status      = "draft",
        )
        .first()
    )
    if draft:
        return draft

    # Fallback: latest published
    return (
        sess.query(DocumentRevision)
        .filter_by(
            project_id  = project_id,
            location_id = location_id,
            status      = "published",
        )
        .order_by(DocumentRevision.id.desc())
        .first()
    )


def get_revision_or_404(
    project_id: int,
    rev_id: int,
    *,
    session=None,
) -> DocumentRevision:
    """
    Fetch a specific revision by project + revision ID.

    Raises a 404 if not found, so routes can call this without
    their own existence check.
    """
    from flask import abort

    sess = session or db.session
    rev  = (
        sess.query(DocumentRevision)
        .filter_by(id=rev_id, project_id=project_id)
        .first()
    )
    if rev is None:
        abort(404)
    return rev


def get_published_revisions(
    project_id: int,
    location_id: Optional[int] = None,
    *,
    session=None,
) -> list:
    """
    Return all published revisions for a project, ordered newest-first.

    If location_id is provided, results are filtered to that location.
    """
    sess = session or db.session

    q = sess.query(DocumentRevision).filter_by(
        project_id = project_id,
        status     = "published",
    )
    if location_id is not None:
        q = q.filter_by(location_id=location_id)

    return q.order_by(DocumentRevision.created_at.desc()).all()


def get_drafts(
    project_id: int,
    location_id: Optional[int] = None,
    *,
    session=None,
) -> list:
    """
    Return all draft revisions for a project.

    If location_id is provided, results are filtered to that location.
    """
    sess = session or db.session

    q = sess.query(DocumentRevision).filter_by(
        project_id = project_id,
        status     = "draft",
    )
    if location_id is not None:
        q = q.filter_by(location_id=location_id)

    return q.order_by(DocumentRevision.created_at.desc()).all()


# ─── Admin: patch document numbers ────────────────────────────────────────────

def patch_doc_numbers(
    rev: DocumentRevision,
    *,
    fi_doc_number:  Optional[str] = None,
    el_doc_number:  Optional[str] = None,
    mov_doc_number: Optional[str] = None,
    io_doc_number:  Optional[str] = None,
    session=None,
) -> DocumentRevision:
    """
    Update document numbers inside a revision's stored payload.

    Only keys supplied as non-None arguments are patched.
    A copy of data_payload is mutated to ensure SQLAlchemy detects
    the change (avoids the silent in-place mutation gotcha on JSON
    columns with some backends).

    The revision is committed before returning.
    """
    sess    = session or db.session
    payload = dict(rev.data_payload)

    def _patch(meta_key: str, new_value: Optional[str]) -> None:
        if new_value is None:
            return
        meta = dict(payload.get(meta_key) or {})
        meta["docNumber"] = new_value.strip()
        payload[meta_key] = meta

    _patch("fi_meta",  fi_doc_number)
    _patch("el_meta",  el_doc_number)
    _patch("mov_meta", mov_doc_number)
    _patch("io_meta",  io_doc_number)

    rev.data_payload = payload
    sess.commit()
    return rev