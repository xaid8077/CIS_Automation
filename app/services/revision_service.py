# app/services/revision_service.py

from datetime import datetime, timezone
from sqlalchemy import func

from app.extensions import db
from app.models import DocumentRevision


# ─────────────────────────────────────────────────────────────────────────────
# Draft Handling (WITH LOCK)
# ─────────────────────────────────────────────────────────────────────────────

def save_draft(project, location, user_id, payload):
    existing = (
        DocumentRevision.query
        .with_for_update()
        .filter_by(
            project_id=project.id,
            location_id=location.id,
            status="draft",
        )
        .first()
    )

    if existing:
        existing.data_payload = payload
        existing.user_id = user_id
        existing.created_at = datetime.now(timezone.utc)
        return existing

    draft = DocumentRevision(
        project_id=project.id,
        user_id=user_id,
        location_id=location.id,
        doc_type="Draft",
        revision_number=0,
        data_payload=payload,
        status="draft",
    )

    db.session.add(draft)
    return draft


# ─────────────────────────────────────────────────────────────────────────────
# Revision Creation (SAFE NUMBERING)
# ─────────────────────────────────────────────────────────────────────────────

def create_revision(project, location, user_id, doc_type, payload):

    latest_rev = db.session.query(
        func.max(DocumentRevision.revision_number)
    ).filter_by(
        project_id=project.id,
        location_id=location.id,
        doc_type=doc_type,
        status="published",
    ).scalar()

    next_rev = (latest_rev + 1) if latest_rev is not None else 0

    new_rev = DocumentRevision(
        project_id=project.id,
        user_id=user_id,
        location_id=location.id,
        doc_type=doc_type,
        revision_number=next_rev,
        data_payload=payload,
        status="published",
    )

    db.session.add(new_rev)
    return new_rev, next_rev