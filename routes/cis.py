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

# routes/cis.py
import threading
import traceback
import uuid
from datetime import datetime, timezone

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

_XLSX_MIME = (
    "application/vnd.openxmlformats-officedocument"
    ".spreadsheetml.sheet"
)

# ─── Async job store ──────────────────────────────────────────────────────────
#
# Maps job_id (UUID hex) → job dict.
# Only excel_service.generate() runs off the main thread; all DB writes
# happen on the main thread so no Flask app-context push is ever needed
# inside the worker.
#
# Job dict shape:
#   status      : "pending" | "ready" | "error"
#   stream      : BytesIO | None
#   error       : str | None
#   project_id  : int
#   loc_id      : int
#   user_id     : int
#   doc_type    : str
#   payload     : dict         (needed by job_download to create the revision)
#   created_at  : datetime (UTC)

_job_store: dict = {}
_job_lock         = threading.Lock()
_JOB_TTL_SECONDS  = 600          # 10 min — jobs older than this are swept


def _sweep_old_jobs() -> None:
    """Remove stale jobs. Called inside _job_lock before adding a new entry."""
    now   = datetime.now(timezone.utc)
    stale = [
        jid for jid, j in _job_store.items()
        if (now - j["created_at"]).total_seconds() > _JOB_TTL_SECONDS
    ]
    for jid in stale:
        _job_store.pop(jid, None)


def _worker_generate_excel(job_id: str, doc_type: str, payload: dict) -> None:
    """
    Background thread target.

    Runs openpyxl entirely off the main thread.  No SQLAlchemy calls —
    the DB write happens later in job_download() on the main thread.
    """
    try:
        from services import excel_service
        stream = excel_service.generate(doc_type, payload)
        with _job_lock:
            _job_store[job_id].update(status="ready", stream=stream)
    except Exception as exc:
        traceback.print_exc()
        with _job_lock:
            _job_store[job_id].update(status="error", error=str(exc))

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
    Validate payload, queue Excel generation in a background thread, and
    return 202 Accepted with a job_id for the client to poll.

    The revision record is NOT created here — it is created in job_download()
    once the file is ready and the user actually downloads it.  This keeps
    all DB writes on the main thread and avoids app-context management inside
    the worker thread.

    Returns
    -------
    202 {"ok": True, "job_id": "...", "doc_type": "..."}  — job queued.
    400 {"error": "..."}  — unsupported doc_type.
    422 {"error": "..."}  — schema / business-rule validation failed.
    500 {"error": "..."}  — unexpected error before thread spawn.
    """
    project  = Project.query.get_or_404(project_id)
    location = ProjectLocation.query.filter_by(
        id=loc_id, project_id=project_id
    ).first_or_404()

    if doc_type not in document_service.SUPPORTED_DOC_TYPES:
        return jsonify({
            "error":     f"Unsupported document type '{doc_type}'.",
            "supported": sorted(document_service.SUPPORTED_DOC_TYPES),
        }), 400

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

    # ── Create job entry and spawn worker thread ───────────────────────────────
    job_id = uuid.uuid4().hex

    with _job_lock:
        _sweep_old_jobs()
        _job_store[job_id] = {
            "status":     "pending",
            "stream":     None,
            "error":      None,
            "project_id": project.id,
            "loc_id":     location.id,
            "user_id":    current_user.id,
            "doc_type":   doc_type,
            "payload":    payload,
            "created_at": datetime.now(timezone.utc),
        }

    thread = threading.Thread(
        target  = _worker_generate_excel,
        args    = (job_id, doc_type, payload),
        daemon  = True,           # thread dies when the process exits
        name    = f"xlsx-{job_id[:8]}",
    )
    thread.start()

    return jsonify({
        "ok":       True,
        "job_id":   job_id,
        "doc_type": doc_type,
    }), 202


@cis_bp.route("/job/<job_id>/status", methods=["GET"])
@login_required
def job_status(job_id):
    """
    Poll endpoint for the async Excel generation job.

    Returns
    -------
    200 {"status": "pending" | "ready" | "error", "error": str | null}
    403 — job belongs to a different user.
    404 — job_id unknown or already consumed.
    """
    with _job_lock:
        job = _job_store.get(job_id)

    if job is None:
        return jsonify({"status": "not_found"}), 404

    if job["user_id"] != current_user.id:
        abort(403)

    return jsonify({
        "status": job["status"],
        "error":  job.get("error"),
    }), 200


@cis_bp.route("/job/<job_id>/download", methods=["GET"])
@login_required
def job_download(job_id):
    """
    Consume a completed job: create the revision record, stream the file,
    and remove the job from the store.

    The revision is created here (on the main thread) so the DB write is
    safely inside the Flask app context with no thread-safety concerns.

    Returns
    -------
    200 application/vnd.openxmlformats…  — xlsx binary stream.
    403 — job belongs to a different user.
    404 — job not found, not yet ready, or already downloaded.
    500 — revision creation or streaming error.
    """
    with _job_lock:
        job = _job_store.get(job_id)

    if job is None:
        return jsonify({"error": "Job not found or already downloaded."}), 404

    if job["user_id"] != current_user.id:
        abort(403)

    if job["status"] != "ready":
        return jsonify({
            "error":  f"Job is not ready (status: {job['status']}).",
            "status": job["status"],
        }), 404

    # ── Consume the job (pop before sending so double-download returns 404) ────
    with _job_lock:
        job = _job_store.pop(job_id, None)

    if job is None:
        # Lost a race with another request for the same job_id.
        return jsonify({"error": "Job already consumed."}), 404

    stream   = job["stream"]
    doc_type = job["doc_type"]
    payload  = job["payload"]

    # ── Create the revision record on the main thread ─────────────────────────
    try:
        project  = Project.query.get_or_404(job["project_id"])
        location = ProjectLocation.query.filter_by(
            id         = job["loc_id"],
            project_id = job["project_id"],
        ).first_or_404()

        rev = revision_service.create_published_revision(
            project  = project,
            location = location,
            user_id  = job["user_id"],
            doc_type = doc_type,
            payload  = payload,
        )

        filename = document_service._build_filename(
            project, location, doc_type, rev.revision_number
        )
    except Exception as exc:
        traceback.print_exc()
        return jsonify({"error": str(exc)}), 500

    stream.seek(0)
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