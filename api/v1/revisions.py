# api/v1/revisions.py
"""
api/v1/revisions.py
───────────────────
Revision list, detail, and draft-save endpoints.

GET  /api/v1/revisions/projects/<id>/                     view_projects
GET  /api/v1/revisions/projects/<id>/locations/<lid>/     view_projects
GET  /api/v1/revisions/<rev_id>                           view_projects
GET  /api/v1/revisions/<rev_id>/payload                   view_projects
POST /api/v1/revisions/projects/<id>/locations/<lid>/draft generate_document

Design notes
────────────
- Listing endpoints support optional ?status= filter
  (published | draft | all, default: published).
- The payload endpoint is separate from the detail endpoint so that
  large JSON payloads are only transferred when explicitly requested,
  keeping list and detail responses lean.
- Draft save delegates to document_service.save_draft() which in turn
  calls revision_service.upsert_draft() — the same path as the web UI.
"""

import traceback

from flask import Blueprint, jsonify, request
from flask_jwt_extended import get_jwt_identity
from marshmallow import ValidationError

from extensions import db
from models import DocumentRevision, Project, ProjectLocation
from schemas.payload import load_payload
from services import document_service, revision_service
from utils.rbac import api_permission_required
from utils.validator import validate_payload

revisions_bp = Blueprint("revisions_api", __name__, url_prefix="/revisions")


# ─── List revisions for a project ────────────────────────────────────────────

@revisions_bp.route("/projects/<int:project_id>/", methods=["GET"])
@api_permission_required("view_projects")
def list_project_revisions(project_id):
    """
    GET /api/v1/revisions/projects/<id>/
    ─────────────────────────────────────
    Query params:
        status = published | draft | all  (default: published)

    Response 200: list of revision dicts (payload excluded).
    """
    Project.query.get_or_404(project_id)

    status = request.args.get("status", "published")
    q      = DocumentRevision.query.filter_by(project_id=project_id)

    if status != "all":
        q = q.filter_by(status=status)

    revisions = q.order_by(DocumentRevision.created_at.desc()).all()
    return jsonify([r.to_dict() for r in revisions]), 200


@revisions_bp.route(
    "/projects/<int:project_id>/locations/<int:loc_id>/",
    methods=["GET"],
)
@api_permission_required("view_projects")
def list_location_revisions(project_id, loc_id):
    """
    GET /api/v1/revisions/projects/<id>/locations/<lid>/
    ─────────────────────────────────────────────────────
    Query params:
        status = published | draft | all  (default: published)

    Response 200: list of revision dicts filtered to this location.
    """
    ProjectLocation.query.filter_by(
        id=loc_id, project_id=project_id
    ).first_or_404()

    status = request.args.get("status", "published")

    if status == "published":
        revisions = revision_service.get_published_revisions(project_id, loc_id)
    elif status == "draft":
        revisions = revision_service.get_drafts(project_id, loc_id)
    else:
        revisions = (
            DocumentRevision.query
            .filter_by(project_id=project_id, location_id=loc_id)
            .order_by(DocumentRevision.created_at.desc())
            .all()
        )

    return jsonify([r.to_dict() for r in revisions]), 200


# ─── Single revision ──────────────────────────────────────────────────────────

@revisions_bp.route("/<int:rev_id>", methods=["GET"])
@api_permission_required("view_projects")
def get_revision(rev_id):
    """
    GET /api/v1/revisions/<rev_id>
    ──────────────────────────────
    Response 200: revision dict (payload excluded).
    Response 404: revision not found.
    """
    rev = DocumentRevision.query.get_or_404(rev_id)
    return jsonify(rev.to_dict()), 200


@revisions_bp.route("/<int:rev_id>/payload", methods=["GET"])
@api_permission_required("view_projects")
def get_revision_payload(rev_id):
    """
    GET /api/v1/revisions/<rev_id>/payload
    ───────────────────────────────────────
    Returns the full stored JSON payload for a revision.
    Kept as a separate endpoint to avoid bloating list responses.

    Response 200: { revision meta + data_payload }
    """
    rev = DocumentRevision.query.get_or_404(rev_id)
    return jsonify(rev.to_dict(include_payload=True)), 200


# ─── Save draft ───────────────────────────────────────────────────────────────

@revisions_bp.route(
    "/projects/<int:project_id>/locations/<int:loc_id>/draft",
    methods=["POST"],
)
@api_permission_required("generate_document")
def save_draft(project_id, loc_id):
    """
    POST /api/v1/revisions/projects/<id>/locations/<lid>/draft
    ───────────────────────────────────────────────────────────
    Body (JSON): full CIS payload (same schema as the web UI).

    Persists the payload as a draft revision.
    Only one draft is kept per project+location — each POST overwrites
    the previous draft.

    Response 200: {"ok": true, "message": "..."}
    Response 422: schema or business-rule validation error.
    Response 500: unexpected server error.
    """
    project  = Project.query.get_or_404(project_id)
    location = ProjectLocation.query.filter_by(
        id=loc_id, project_id=project_id
    ).first_or_404()

    try:
        raw     = request.get_json(force=True, silent=True) or {}
        payload = load_payload(raw)
    except ValidationError as exc:
        return jsonify({"error": "Validation failed.", "detail": exc.messages}), 422

    try:
        identity = get_jwt_identity()
        result   = document_service.save_draft(
            project  = project,
            location = location,
            user_id  = identity["id"],
            payload  = payload,
        )
        return jsonify(result), 200

    except Exception as exc:
        traceback.print_exc()
        return jsonify({"error": str(exc)}), 500


# ─── Admin: patch doc numbers ─────────────────────────────────────────────────

@revisions_bp.route("/<int:rev_id>/doc-numbers", methods=["PATCH"])
@api_permission_required("patch_doc_numbers")
def patch_doc_numbers(rev_id):
    """
    PATCH /api/v1/revisions/<rev_id>/doc-numbers
    ─────────────────────────────────────────────
    Body (JSON — all fields optional):
        {
          "fi_doc_number":  "PRJ-IL-001",
          "io_doc_number":  "PRJ-IOL-001",
          "el_doc_number":  "PRJ-EL-001",
          "mov_doc_number": "PRJ-MOV-001"
        }

    Updates the document numbers stored inside the revision's payload.
    Does NOT regenerate the Excel file.

    Response 200: updated revision dict (payload excluded).
    """
    rev = DocumentRevision.query.get_or_404(rev_id)

    data = request.get_json(force=True, silent=True) or {}

    revision_service.patch_doc_numbers(
        rev,
        fi_doc_number  = data.get("fi_doc_number"),
        el_doc_number  = data.get("el_doc_number"),
        mov_doc_number = data.get("mov_doc_number"),
        io_doc_number  = data.get("io_doc_number"),
    )

    return jsonify(rev.to_dict()), 200