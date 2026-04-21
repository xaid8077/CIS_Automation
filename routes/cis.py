# routes/cis.py
"""
cis.py
──────
CIS blueprint — project list, dashboard, editor, and document download.

Every route follows the same three-step pattern:
  1. Authorise + fetch ORM objects.
  2. Delegate to a service (schema, revision_service, document_service).
  3. Return an HTTP response.

No Excel generation, no direct DB writes (except through services),
and no request parsing beyond extracting JSON live in this file.
"""

import traceback

from flask import (
    Blueprint, render_template, request,
    redirect, url_for, jsonify, send_file, abort,
)
from flask_login import login_required, current_user
from marshmallow import ValidationError

from models import Project, ProjectLocation, DocumentRevision
from schemas.payload import load_payload
from services import revision_service, document_service
from utils.validator import validate_payload

cis_bp = Blueprint("cis", __name__)

# MIME type constant so it is not repeated across download routes.
_XLSX_MIME = (
    "application/vnd.openxmlformats-officedocument"
    ".spreadsheetml.sheet"
)


# ─── Project list ─────────────────────────────────────────────────────────────

@cis_bp.route("/")
@login_required
def index():
    """Entry point after login — shows all projects."""
    projects = Project.query.order_by(Project.created_at.desc()).all()
    return render_template("project_list.html", projects=projects)


# ─── Project dashboard ────────────────────────────────────────────────────────

@cis_bp.route("/project/<int:project_id>")
@login_required
def project_dashboard(project_id):
    """
    Per-project revision history.

    An optional ?loc=<id> query-param filters the history table to
    a single location and reveals the "Generate New Revision" button.
    """
    project   = Project.query.get_or_404(project_id)
    locations = (
        ProjectLocation.query
        .filter_by(project_id=project.id)
        .order_by(ProjectLocation.name)
        .all()
    )

    loc_id   = request.args.get("loc", type=int)
    location = None
    if loc_id:
        location = ProjectLocation.query.filter_by(
            id=loc_id, project_id=project.id
        ).first()

    revisions = revision_service.get_published_revisions(
        project_id  = project.id,
        location_id = location.id if location else None,
    )
    drafts = revision_service.get_drafts(
        project_id  = project.id,
        location_id = location.id if location else None,
    )

    return render_template(
        "project_dashboard.html",
        project         = project,
        locations       = locations,
        active_location = location,
        revisions       = revisions,
        drafts          = drafts,
    )


# ─── Editor ───────────────────────────────────────────────────────────────────

@cis_bp.route("/project/<int:project_id>/location/<int:loc_id>/edit-docs")
@login_required
def edit_docs(project_id, loc_id):
    """
    Open the 5-step wizard for a specific project + location.

    Pre-fills grid data from the most recent draft (preferred) or
    the latest published revision if no draft exists.
    """
    project  = Project.query.get_or_404(project_id)
    location = ProjectLocation.query.filter_by(
        id=loc_id, project_id=project_id
    ).first_or_404()

    latest_rev = revision_service.get_latest_for_editor(
        project_id  = project.id,
        location_id = location.id,
    )
    previous_data = latest_rev.data_payload if latest_rev else None

    return render_template(
        "index.html",
        project       = project,
        location      = location,
        previous_data = previous_data,
    )


# ─── Validate (preview) ───────────────────────────────────────────────────────

@cis_bp.route("/preview", methods=["POST"])
@login_required
def preview():
    """
    Validate the submitted payload without writing to the DB or
    generating any file.  Called by the JS "Validate" button.

    Returns
    -------
    200 {"ok": True,  "message": "..."}  — validation passed.
    422 {"ok": False, "errors": [...]}   — validation failed.
    500 {"ok": False, "errors": [...]}   — unexpected server error.
    """
    try:
        raw     = request.get_json(force=True, silent=True) or {}
        payload = load_payload(raw)
    except ValidationError as exc:
        return jsonify({"ok": False, "errors": _flatten_errors(exc)}), 422

    try:
        # Business-rule validation (cross-row, duplicate tags, etc.)
        errors = validate_payload(payload, require_doc_numbers=False)
        if errors:
            return jsonify({"ok": False, "errors": errors}), 422

        row_counts = {
            "field_instruments": len(payload["field_instruments"]),
            "electrical":        len(payload["electrical"]),
            "mov":               len(payload["mov"]),
        }
        msg = (
            f"Validation passed — "
            f"{row_counts['field_instruments']} field instrument(s), "
            f"{row_counts['electrical']} electrical row(s), "
            f"{row_counts['mov']} MOV row(s)."
        )
        return jsonify({"ok": True, "message": msg}), 200

    except Exception as exc:
        traceback.print_exc()
        return jsonify({"ok": False, "errors": [str(exc)]}), 500


# ─── Save draft ───────────────────────────────────────────────────────────────

@cis_bp.route(
    "/project/<int:project_id>/location/<int:loc_id>/save-draft",
    methods=["POST"],
)
@login_required
def save_draft(project_id, loc_id):
    """
    Persist all grid data as a draft revision (no file generated).

    Returns
    -------
    200 {"ok": True,  "message": "..."}
    422 {"ok": False, "errors": [...]}  — schema validation failed.
    500 {"ok": False, "error":  "..."}  — unexpected server error.
    """
    project  = Project.query.get_or_404(project_id)
    location = ProjectLocation.query.filter_by(
        id=loc_id, project_id=project_id
    ).first_or_404()

    try:
        raw     = request.get_json(force=True, silent=True) or {}
        payload = load_payload(raw)
    except ValidationError as exc:
        return jsonify({"ok": False, "errors": _flatten_errors(exc)}), 422

    try:
        result = document_service.save_draft(
            project  = project,
            location = location,
            user_id  = current_user.id,
            payload  = payload,
        )
        return jsonify(result), 200

    except Exception as exc:
        traceback.print_exc()
        return jsonify({"ok": False, "error": str(exc)}), 500


# ─── Generate + download ──────────────────────────────────────────────────────

@cis_bp.route(
    "/project/<int:project_id>/location/<int:loc_id>/submit-doc/<doc_type>",
    methods=["POST"],
)
@login_required
def submit_and_save(project_id, loc_id, doc_type):
    """
    Validate, persist as a published revision, generate Excel, download.

    Returns
    -------
    200 application/vnd.openxmlformats… — xlsx binary stream.
    400 {"error": "..."}  — unsupported doc_type.
    422 {"error": "..."}  — schema / business-rule validation failed.
    500 {"error": "..."}  — unexpected server error.
    """
    project  = Project.query.get_or_404(project_id)
    location = ProjectLocation.query.filter_by(
        id=loc_id, project_id=project_id
    ).first_or_404()

    # ── Schema validation ─────────────────────────────────────────────────────
    try:
        raw     = request.get_json(force=True, silent=True) or {}
        payload = load_payload(raw)
    except ValidationError as exc:
        return jsonify({"error": _flatten_errors(exc)}), 422

    # ── Business-rule validation ──────────────────────────────────────────────
    errors = validate_payload(payload, require_doc_numbers=True)
    if errors:
        return jsonify({"ok": False, "errors": errors}), 422

    # ── Generate + save ───────────────────────────────────────────────────────
    try:
        stream, filename = document_service.generate_and_save(
            project  = project,
            location = location,
            user_id  = current_user.id,
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

@cis_bp.route(
    "/project/<int:project_id>/revision/<int:rev_id>/download",
    methods=["GET"],
)
@login_required
def download_revision(project_id, rev_id):
    """
    Re-generate an Excel file from a stored revision's payload.

    Pure read path — no DB write, no schema validation needed
    (the payload was validated when the revision was first created).

    Returns
    -------
    200 application/vnd.openxmlformats… — xlsx binary stream.
    400 {"error": "..."}  — doc_type not supported for re-download.
    500 {"error": "..."}  — unexpected server error.
    """
    project = Project.query.get_or_404(project_id)

    try:
        stream, filename = document_service.regenerate_from_revision(
            project = project,
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


# ─── Helpers ──────────────────────────────────────────────────────────────────

def _flatten_errors(exc: ValidationError) -> list:
    """
    Convert a Marshmallow ValidationError's nested messages dict
    into a flat list of human-readable strings.

    Example
    -------
    {"header": {"projectName": ["Project Name is required."]}}
    → ["header.projectName: Project Name is required."]
    """
    messages = exc.messages
    flat     = []

    def _walk(node, prefix=""):
        if isinstance(node, dict):
            for key, val in node.items():
                _walk(val, f"{prefix}.{key}" if prefix else str(key))
        elif isinstance(node, list):
            for item in node:
                if isinstance(item, str):
                    flat.append(f"{prefix}: {item}" if prefix else item)
                else:
                    _walk(item, prefix)
        elif isinstance(node, str):
            flat.append(f"{prefix}: {node}" if prefix else node)

    _walk(messages)
    return flat