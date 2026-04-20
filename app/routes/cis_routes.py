# app/routes/cis_routes.py

import logging
import traceback

from flask import (
    Blueprint, render_template, request,
    jsonify, send_file
)
from flask_login import login_required, current_user

from app.extensions import db
from app.models import Project, ProjectLocation, DocumentRevision

from app.services.document_service import build_payload, validate
from app.services.revision_service import save_draft, create_revision
from app.services.excel_service import generate_excel, build_filename


logger = logging.getLogger(__name__)

cis_bp = Blueprint("cis", __name__)


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def get_safe_json():
    data = request.get_json(silent=True)
    if not data:
        return None, jsonify({"error": "Invalid or empty JSON"}), 400
    return data, None, None


def check_payload_size(payload):
    if len(payload.get("field_instruments", [])) > 5000:
        return False
    return True


# ─────────────────────────────────────────────────────────────────────────────
# Routes
# ─────────────────────────────────────────────────────────────────────────────

@cis_bp.route("/")
@login_required
def index():
    projects = Project.query.order_by(Project.created_at.desc()).all()
    return render_template("project_list.html", projects=projects)


# ─────────────────────────────────────────────────────────────────────────────
# Preview
# ─────────────────────────────────────────────────────────────────────────────

@cis_bp.route("/preview", methods=["POST"])
@login_required
def preview():
    try:
        data, err_resp, code = get_safe_json()
        if err_resp:
            return err_resp, code

        payload = build_payload(data)

        if not check_payload_size(payload):
            return jsonify({"error": "Too much data"}), 413

        errors = validate(payload, require_doc_numbers=False)

        if errors:
            return jsonify({"ok": False, "errors": errors}), 422

        return jsonify({"ok": True}), 200

    except Exception:
        logger.exception("Preview failed")
        return jsonify({"ok": False, "errors": ["Internal error"]}), 500


# ─────────────────────────────────────────────────────────────────────────────
# Save Draft
# ─────────────────────────────────────────────────────────────────────────────

@cis_bp.route(
    "/project/<int:project_id>/location/<int:loc_id>/save-draft",
    methods=["POST"],
)
@login_required
def save_draft_route(project_id, loc_id):
    try:
        data, err_resp, code = get_safe_json()
        if err_resp:
            return err_resp, code

        project = Project.query.get_or_404(project_id)
        location = ProjectLocation.query.filter_by(
            id=loc_id, project_id=project_id
        ).first_or_404()

        payload = build_payload(data)

        if not check_payload_size(payload):
            return jsonify({"error": "Too much data"}), 413

        save_draft(project, location, current_user.id, payload)

        db.session.commit()

        return jsonify({"ok": True}), 200

    except Exception:
        logger.exception("Save draft failed")
        db.session.rollback()
        return jsonify({"ok": False, "error": "Internal error"}), 500


# ─────────────────────────────────────────────────────────────────────────────
# Submit & Generate
# ─────────────────────────────────────────────────────────────────────────────

@cis_bp.route(
    "/project/<int:project_id>/location/<int:loc_id>/submit-doc/<doc_type>",
    methods=["POST"],
)
@login_required
def submit_and_save(project_id, loc_id, doc_type):
    try:
        data, err_resp, code = get_safe_json()
        if err_resp:
            return err_resp, code

        project = Project.query.get_or_404(project_id)
        location = ProjectLocation.query.filter_by(
            id=loc_id, project_id=project_id
        ).first_or_404()

        payload = build_payload(data)

        if not check_payload_size(payload):
            return jsonify({"error": "Too much data"}), 413

        rev, rev_num = create_revision(
            project, location, current_user.id, doc_type, payload
        )

        db.session.commit()

        output, prefix = generate_excel(payload, doc_type)
        filename = build_filename(project, location, prefix, rev_num)

        return send_file(
            output,
            as_attachment=True,
            download_name=filename
        )

    except Exception:
        logger.exception("Submit failed")
        db.session.rollback()
        return jsonify({"error": "Internal error"}), 500


# ─────────────────────────────────────────────────────────────────────────────
# Download Existing Revision
# ─────────────────────────────────────────────────────────────────────────────

@cis_bp.route(
    "/project/<int:project_id>/revision/<int:rev_id>/download",
    methods=["GET"],
)
@login_required
def download_revision(project_id, rev_id):
    try:
        project = Project.query.get_or_404(project_id)

        rev = DocumentRevision.query.filter_by(
            id=rev_id, project_id=project_id
        ).first_or_404()

        output, prefix = generate_excel(rev.data_payload, rev.doc_type)

        filename = build_filename(
            project,
            rev.location,
            prefix,
            rev.revision_number
        )

        return send_file(
            output,
            as_attachment=True,
            download_name=filename
        )

    except Exception:
        logger.exception("Download failed")
        return jsonify({"error": "Internal error"}), 500