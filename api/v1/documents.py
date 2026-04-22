# api/v1/documents.py
"""
api/v1/documents.py
───────────────────
Document generation and re-download endpoints.

POST /api/v1/documents/projects/<id>/locations/<lid>/<doc_type>
    Validate payload, create a published revision, stream the Excel file.
    Requires: generate_document permission (admin + engineer).

GET  /api/v1/documents/revisions/<rev_id>/download
    Re-generate and stream an Excel file from a stored revision.
    Requires: download_revision permission (all roles).

Design notes
────────────
- doc_type is passed as a URL segment and validated against
  document_service.SUPPORTED_DOC_TYPES before any DB or file work.
- The generate endpoint validates both schema (Marshmallow) and
  business rules (validator.py) before writing anything to the DB.
- Both endpoints stream BytesIO directly — no temp files on disk.
- Content-Disposition filename is built by document_service so the
  naming convention is consistent between web UI and API downloads.
"""

import traceback

from flask import Blueprint, jsonify, request, send_file
from flask_jwt_extended import get_jwt_identity
from marshmallow import ValidationError

from models import Project, ProjectLocation
from schemas.payload import load_payload
from services import document_service
from utils.rbac import api_permission_required
from utils.validator import validate_payload

documents_bp = Blueprint("documents_api", __name__, url_prefix="/documents")

_XLSX_MIME = (
    "application/vnd.openxmlformats-officedocument"
    ".spreadsheetml.sheet"
)


# ─── Generate + save ──────────────────────────────────────────────────────────

@documents_bp.route(
    "/projects/<int:project_id>/locations/<int:loc_id>/<doc_type>",
    methods=["POST"],
)
@api_permission_required("generate_document")
def generate_document(project_id, loc_id, doc_type):
    """
    POST /api/v1/documents/projects/<id>/locations/<lid>/<doc_type>
    ───────────────────────────────────────────────────────────────
    <doc_type> must be one of:
        Instrument%20List   → "Instrument List"
        IO%20List           → "IO List"

    Body (JSON): full CIS payload.

    Response 200: application/vnd.openxmlformats… xlsx binary stream.
    Response 400: unsupported doc_type.
    Response 422: validation failed.
    Response 500: server error.

    On success the response includes:
        Content-Disposition: attachment; filename="..."
    so API clients can save the file with the correct name.
    """
    # URL-decode replaces %20 — Flask does this automatically,
    # but we normalise whitespace just in case.
    doc_type = " ".join(doc_type.split())

    if doc_type not in document_service.SUPPORTED_DOC_TYPES:
        return jsonify({
            "error":    f"Unsupported document type '{doc_type}'.",
            "supported": sorted(document_service.SUPPORTED_DOC_TYPES),
        }), 400

    project  = Project.query.get_or_404(project_id)
    location = ProjectLocation.query.filter_by(
        id=loc_id, project_id=project_id
    ).first_or_404()

    # ── Schema validation ─────────────────────────────────────────────────────
    try:
        raw     = request.get_json(force=True, silent=True) or {}
        payload = load_payload(raw)
    except ValidationError as exc:
        return jsonify({"error": "Validation failed.", "detail": exc.messages}), 422

    # ── Business-rule validation ──────────────────────────────────────────────
    errors = validate_payload(payload, require_doc_numbers=True)
    if errors:
        return jsonify({"ok": False, "errors": errors}), 422

    # ── Generate + save ───────────────────────────────────────────────────────
    try:
        identity         = get_jwt_identity()
        stream, filename = document_service.generate_and_save(
            project  = project,
            location = location,
            user_id  = identity["id"],
            doc_type = doc_type,
            payload  = payload,
        )
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        traceback.print_exc()
        return jsonify({"error": str(exc)}), 500

    return send_file(
        stream,
        as_attachment = True,
        download_name = filename,
        mimetype      = _XLSX_MIME,
    )


# ─── Re-download stored revision ──────────────────────────────────────────────

@documents_bp.route("/revisions/<int:rev_id>/download", methods=["GET"])
@api_permission_required("download_revision")
def download_revision(rev_id):
    """
    GET /api/v1/documents/revisions/<rev_id>/download
    ──────────────────────────────────────────────────
    Re-generates and streams an Excel file from a stored revision's
    payload.  Pure read path — no DB write.

    Response 200: xlsx binary stream.
    Response 400: doc_type not supported for re-download.
    Response 404: revision not found.
    Response 500: server error.
    """
    from models import DocumentRevision
    rev = DocumentRevision.query.get_or_404(rev_id)

    try:
        stream, filename = document_service.regenerate_from_revision(
            project = rev.project,
            rev_id  = rev_id,
        )
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        traceback.print_exc()
        return jsonify({"error": str(exc)}), 500

    return send_file(
        stream,
        as_attachment = True,
        download_name = filename,
        mimetype      = _XLSX_MIME,
    )