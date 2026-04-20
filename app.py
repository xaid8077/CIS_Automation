"""
app.py
──────
Application factory.

Changes vs previous version:
  - Admin routes for nickname editing and location management
  - cis_bp routes updated to handle location_id on DocumentRevision
  - preview route added (was missing, caused silent 404)
  - download_revision route added (re-generate from stored JSON)
  - edit_docs now receives location object so template can display it
  - submit_and_save now records location_id
"""

import os
import functools
import traceback
from datetime import datetime, timezone
from io import BytesIO
from flask import Request as _FlaskRequest

from flask import (
    Flask, Blueprint, render_template, request,
    redirect, url_for, flash, jsonify, send_file,
    abort, session,
)
from flask_login import (
    login_user, logout_user, login_required,
    current_user,
)

from config     import get_config
from extensions import db, login_manager, csrf, limiter
from models     import User, Project, ProjectLocation, DocumentRevision
from forms import (
    LoginForm, RegisterForm, EditUserForm,
    ProjectForm, EditProjectNameForm, ProjectNicknameForm,
    ProjectLocationForm, RevisionDocNumbersForm,
)
from utils.excel_writer import write_workbook, write_io_workbook
from utils.validator    import validate_payload


class _BigRequest(_FlaskRequest):
    """
    Werkzeug 3.x enforces form-size limits as class attributes on the
    Request object, independent of Flask config.  Override them here so
    large grid payloads are never rejected with 413.
    """
    max_content_length   = 16 * 1024 * 1024   # 16 MB total POST body
    max_form_memory_size = 16 * 1024 * 1024   # 16 MB per form field
    max_form_parts       = 10_000             # max number of form fields


# ─────────────────────────────────────────────────────────────────────────────
# Decorators
# ─────────────────────────────────────────────────────────────────────────────

def admin_required(f):
    @functools.wraps(f)
    @login_required
    def wrapped(*args, **kwargs):
        if not current_user.is_admin:
            abort(403)
        return f(*args, **kwargs)
    return wrapped


# ─────────────────────────────────────────────────────────────────────────────
# Auth Blueprint
# ─────────────────────────────────────────────────────────────────────────────

auth_bp = Blueprint("auth", __name__)


@auth_bp.route("/login", methods=["GET", "POST"])
@limiter.limit("10 per minute")
def login():
    if current_user.is_authenticated:
        return redirect(url_for("cis.index"))

    form = LoginForm()
    if form.validate_on_submit():
        user = User.query.filter_by(username=form.username.data).first()

        dummy_hash = "$argon2id$v=19$m=65536,t=3,p=2$dummy$dummy"
        if user is None:
            try:
                from argon2 import PasswordHasher
                PasswordHasher().verify(dummy_hash, form.password.data)
            except Exception:
                pass
            flash("Invalid username or password.", "danger")
            return render_template("login.html", form=form)

        if not user.is_active:
            flash("Your account is disabled. Contact an administrator.", "danger")
            return render_template("login.html", form=form)

        if not user.check_password(form.password.data):
            flash("Invalid username or password.", "danger")
            return render_template("login.html", form=form)

        login_user(user, remember=form.remember.data)
        user.last_login = datetime.now(timezone.utc)
        db.session.commit()

        next_page = request.args.get("next", "")
        if next_page and next_page.startswith("/") and not next_page.startswith("//"):
            return redirect(next_page)
        return redirect(url_for("cis.index"))

    return render_template("login.html", form=form)


@auth_bp.route("/logout", methods=["POST"])
@login_required
def logout():
    logout_user()
    session.clear()
    flash("You have been signed out.", "info")
    return redirect(url_for("auth.login"))


# ─────────────────────────────────────────────────────────────────────────────
# Admin Blueprint
# ─────────────────────────────────────────────────────────────────────────────

admin_bp = Blueprint("admin", __name__, url_prefix="/admin")


# REPLACE the existing dashboard() route with:
@admin_bp.route("/")
@admin_required
def dashboard():
    users             = User.query.order_by(User.created_at.desc()).all()
    projects          = Project.query.order_by(Project.created_at.desc()).all()
    user_form         = RegisterForm()
    project_form      = ProjectForm()
    nickname_form     = ProjectNicknameForm()
    edit_project_form = EditProjectNameForm()
    loc_form          = ProjectLocationForm()
    rev_doc_form      = RevisionDocNumbersForm()
    return render_template(
        "admin.html",
        users=users,
        projects=projects,
        user_form=user_form,
        project_form=project_form,
        nickname_form=nickname_form,
        edit_project_form=edit_project_form,
        loc_form=loc_form,
        rev_doc_form=rev_doc_form,
    )


@admin_bp.route("/users/create", methods=["POST"])
@admin_required
def create_user():
    form = RegisterForm()
    if form.validate_on_submit():
        user = User(
            username  = form.username.data,
            email     = form.email.data,
            role      = form.role.data,
            is_active = True,
        )
        user.set_password(form.password.data)
        db.session.add(user)
        db.session.commit()
        flash(f"User '{user.username}' created.", "success")
    else:
        for errs in form.errors.values():
            for err in errs:
                flash(err, "danger")
    return redirect(url_for("admin.dashboard"))


@admin_bp.route("/projects/create", methods=["POST"])
@admin_required
def create_project():
    form = ProjectForm()
    if form.validate_on_submit():
        project = Project(
            name       = form.name.data.strip(),
            nickname   = form.nickname.data.strip() if form.nickname.data else None,
            client     = form.client.data.strip(),
            consultant = form.consultant.data.strip() if form.consultant.data else None,
        )
        db.session.add(project)
        db.session.commit()
        flash(f"Project '{project.display_name}' created.", "success")
    else:
        for errs in form.errors.values():
            for err in errs:
                flash(err, "danger")
    return redirect(url_for("admin.dashboard"))


@admin_bp.route("/projects/<int:project_id>/nickname", methods=["POST"])
@admin_required
def set_nickname(project_id):
    """Set or update a project's nickname."""
    project = Project.query.get_or_404(project_id)
    form    = ProjectNicknameForm()
    if form.validate_on_submit():
        raw = form.nickname.data.strip() if form.nickname.data else ""
        project.nickname = raw if raw else None
        db.session.commit()
        flash(
            f"Nickname for '{project.name}' updated to "
            f"'{project.display_name}'.",
            "success",
        )
    else:
        for errs in form.errors.values():
            for err in errs:
                flash(err, "danger")
    return redirect(url_for("admin.dashboard"))


@admin_bp.route("/projects/<int:project_id>/locations/add", methods=["POST"])
@admin_required
def add_location(project_id):
    """Add a new location to a project."""
    project = Project.query.get_or_404(project_id)
    form    = ProjectLocationForm()
    if form.validate_on_submit():
        # Check uniqueness within the project
        existing = ProjectLocation.query.filter_by(
            project_id=project.id,
            name=form.name.data.strip(),
        ).first()
        if existing:
            flash(
                f"Location '{form.name.data.strip()}' already exists "
                f"in project '{project.display_name}'.",
                "warning",
            )
        else:
            loc = ProjectLocation(
                project_id = project.id,
                name       = form.name.data.strip(),
                code       = form.code.data.strip() if form.code.data else None,
            )
            db.session.add(loc)
            db.session.commit()
            flash(
                f"Location '{loc.display}' added to "
                f"'{project.display_name}'.",
                "success",
            )
    else:
        for errs in form.errors.values():
            for err in errs:
                flash(err, "danger")
    return redirect(url_for("admin.dashboard"))


@admin_bp.route(
    "/projects/<int:project_id>/locations/<int:loc_id>/delete",
    methods=["POST"],
)
@admin_required
def delete_location(project_id, loc_id):
    """
    Delete a location — only allowed if no revisions reference it.
    This prevents orphaned revision records.
    """
    loc = ProjectLocation.query.filter_by(
        id=loc_id, project_id=project_id
    ).first_or_404()

    if loc.revisions:
        flash(
            f"Cannot delete '{loc.display}' — it has "
            f"{len(loc.revisions)} revision(s) attached. "
            "Delete the revisions first or archive the location.",
            "danger",
        )
    else:
        name = loc.display
        db.session.delete(loc)
        db.session.commit()
        flash(f"Location '{name}' deleted.", "success")

    return redirect(url_for("admin.dashboard"))

@admin_bp.route("/projects/<int:project_id>/edit", methods=["POST"])
@admin_required
def edit_project(project_id):
    """
    Edit a project's full name, nickname, client, and consultant.
    No length cap on name — engineering project names can be long.
    Checks uniqueness only when the name actually changes.
    """
    project = Project.query.get_or_404(project_id)
    form    = EditProjectNameForm()

    if form.validate_on_submit():
        new_name = form.name.data.strip()

        # Uniqueness check — only relevant if the name changed
        if new_name != project.name:
            clash = Project.query.filter_by(name=new_name).first()
            if clash:
                flash(
                    f"Another project already has the name '{new_name}'.",
                    "danger",
                )
                return redirect(url_for("admin.dashboard"))

        project.name       = new_name
        project.nickname   = (form.nickname.data.strip()   or None) if form.nickname.data   else None
        project.client     = form.client.data.strip()
        project.consultant = (form.consultant.data.strip() or None) if form.consultant.data else None

        db.session.commit()
        flash(f"Project updated: '{project.display_name}'.", "success")
    else:
        for errs in form.errors.values():
            for err in errs:
                flash(err, "danger")

    return redirect(url_for("admin.dashboard"))


@admin_bp.route(
    "/projects/<int:project_id>/revisions/<int:rev_id>/doc-numbers",
    methods=["POST"],
)
@admin_required
def edit_revision_doc_numbers(project_id, rev_id):
    """
    Update document numbers stored inside an existing revision's data_payload.
    Does NOT regenerate the Excel file — only the stored JSON is patched.
    Useful for correcting a doc number after a revision has been published.
    """
    rev  = DocumentRevision.query.filter_by(
        id=rev_id, project_id=project_id
    ).first_or_404()
    form = RevisionDocNumbersForm()

    if form.validate_on_submit():
        # data_payload is a JSON column — mutate a copy so SQLAlchemy
        # detects the change reliably (avoids in-place mutation gotcha).
        payload = dict(rev.data_payload)

        def _patch_meta(key, value):
            if value is not None:
                meta          = dict(payload.get(key) or {})
                meta["docNumber"] = value.strip()
                payload[key]  = meta

        _patch_meta("fi_meta",  form.fi_doc_number.data)
        _patch_meta("el_meta",  form.el_doc_number.data)
        _patch_meta("mov_meta", form.mov_doc_number.data)
        _patch_meta("io_meta",  form.io_doc_number.data)

        rev.data_payload = payload
        db.session.commit()
        flash(
            f"Document numbers updated for revision {rev.revision_number} "
            f"({rev.doc_type}).",
            "success",
        )
    else:
        for errs in form.errors.values():
            for err in errs:
                flash(err, "danger")

    return redirect(url_for("admin.dashboard"))
# ─────────────────────────────────────────────────────────────────────────────
# CIS Blueprint
# ─────────────────────────────────────────────────────────────────────────────

cis_bp = Blueprint("cis", __name__)


# ── Payload builders ──────────────────────────────────────────────────────────

def _build_payload(data: dict) -> dict:
    """
    Build canonical payload from a parsed JSON request body.
    Called with:  request.get_json(force=True, silent=True) or {}
    """
    hdr = data.get("header") or {}

    def _s(d, k):
        return (d.get(k) or "").strip() if isinstance(d, dict) else ""

    return {
        "header": {k: _s(hdr, k) for k in [
            "projectName", "client", "consultant", "location",
            "date", "preparedBy", "checkedBy", "approvedBy", "revision",
        ]},
        "fi_meta":           {"docNumber": _s(data.get("fi_meta")  or {}, "docNumber")},
        "el_meta":           {"docNumber": _s(data.get("el_meta")  or {}, "docNumber")},
        "mov_meta":          {"docNumber": _s(data.get("mov_meta") or {}, "docNumber")},
        "io_meta":           {"docNumber": _s(data.get("io_meta")  or {}, "docNumber")},
        "field_instruments": data.get("field_instruments") or [],
        "electrical":        data.get("electrical")        or [],
        "mov":               data.get("mov")               or [],
    }


# ── Routes ────────────────────────────────────────────────────────────────────

@cis_bp.route("/")
@login_required
def index():
    """Project list — entry point after login."""
    projects = Project.query.order_by(Project.created_at.desc()).all()
    return render_template("project_list.html", projects=projects)


@cis_bp.route("/project/<int:project_id>")
@login_required
def project_dashboard(project_id):
    project   = Project.query.get_or_404(project_id)
    locations = ProjectLocation.query.filter_by(
        project_id=project.id
    ).order_by(ProjectLocation.name).all()

    loc_id   = request.args.get("loc", type=int)
    location = None
    if loc_id:
        location = ProjectLocation.query.filter_by(
            id=loc_id, project_id=project.id
        ).first()

    base_q = DocumentRevision.query.filter_by(project_id=project.id)
    if location:
        base_q = base_q.filter_by(location_id=location.id)

    # Published revisions shown in history table
    revisions = (
        base_q.filter_by(status="published")
        .order_by(DocumentRevision.created_at.desc())
        .all()
    )

    # Draft indicator — one per location at most
    drafts = (
        base_q.filter_by(status="draft")
        .order_by(DocumentRevision.created_at.desc())
        .all()
    )

    return render_template(
        "project_dashboard.html",
        project=project,
        locations=locations,
        active_location=location,
        revisions=revisions,
        drafts=drafts,
    )


@cis_bp.route("/project/<int:project_id>/location/<int:loc_id>/edit-docs")
@login_required
def edit_docs(project_id, loc_id):
    project  = Project.query.get_or_404(project_id)
    location = ProjectLocation.query.filter_by(
        id=loc_id, project_id=project_id
    ).first_or_404()

    # Prefer unsaved draft — it reflects the most recent work.
    # Fall back to latest published revision if no draft exists.
    latest_rev = DocumentRevision.query.filter_by(
        project_id  = project.id,
        location_id = location.id,
        status      = "draft",
    ).first()

    if not latest_rev:
        latest_rev = DocumentRevision.query.filter_by(
            project_id  = project.id,
            location_id = location.id,
            status      = "published",
        ).order_by(DocumentRevision.id.desc()).first()

    previous_data = latest_rev.data_payload if latest_rev else None

    return render_template(
        "index.html",
        project=project,
        location=location,
        previous_data=previous_data,
    )


@cis_bp.route("/preview", methods=["POST"])
@login_required
def preview():
    """
    Validate payload without writing to DB or generating a file.
    Called by the JS previewData() function.
    """
    try:
        payload = _build_payload(request.get_json(force=True, silent=True) or {})
        errors  = validate_payload(payload, require_doc_numbers=False)
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
            f"{row_counts['electrical']} electrical equipment row(s), "
            f"{row_counts['mov']} MOV row(s)."
        )
        return jsonify({"ok": True, "message": msg}), 200
    except Exception as e:
        traceback.print_exc()
        return jsonify({"ok": False, "errors": [str(e)]}), 500


@cis_bp.route(
    "/project/<int:project_id>/location/<int:loc_id>/save-draft",
    methods=["POST"],
)
@login_required
def save_draft(project_id, loc_id):
    """
    Save all grid + header data as a draft revision.
    Only one draft is kept per project+location — each save overwrites
    the previous draft so the history table stays clean.
    No file is generated.
    """
    project  = Project.query.get_or_404(project_id)
    location = ProjectLocation.query.filter_by(
        id=loc_id, project_id=project_id
    ).first_or_404()

    try:
        payload = _build_payload(request.get_json(force=True, silent=True) or {})

        existing_draft = DocumentRevision.query.filter_by(
            project_id  = project.id,
            location_id = location.id,
            status      = "draft",
        ).first()

        if existing_draft:
            existing_draft.data_payload = payload
            existing_draft.user_id      = current_user.id
            existing_draft.created_at   = datetime.now(timezone.utc)
        else:
            draft = DocumentRevision(
                project_id      = project.id,
                user_id         = current_user.id,
                location_id     = location.id,
                doc_type        = "Draft",
                revision_number = 0,
                data_payload    = payload,
                status          = "draft",
            )
            db.session.add(draft)

        db.session.commit()
        return jsonify({"ok": True, "message": "Data saved successfully."}), 200

    except Exception as e:
        traceback.print_exc()
        db.session.rollback()
        return jsonify({"ok": False, "error": str(e)}), 500


@cis_bp.route(
    "/project/<int:project_id>/location/<int:loc_id>/submit-doc/<doc_type>",
    methods=["POST"],
)
@login_required
def submit_and_save(project_id, loc_id, doc_type):
    try:
        project  = Project.query.get_or_404(project_id)
        location = ProjectLocation.query.filter_by(
            id=loc_id, project_id=project_id
        ).first_or_404()

        payload = _build_payload(request.get_json(force=True, silent=True) or {})

        # Revision numbers count only published revisions per doc_type
        latest_published = DocumentRevision.query.filter_by(
            project_id  = project.id,
            location_id = location.id,
            doc_type    = doc_type,
            status      = "published",
        ).order_by(DocumentRevision.revision_number.desc()).first()

        next_rev_num = (latest_published.revision_number + 1) if latest_published else 0

        new_rev = DocumentRevision(
            project_id      = project.id,
            user_id         = current_user.id,
            location_id     = location.id,
            doc_type        = doc_type,
            revision_number = next_rev_num,
            data_payload    = payload,
            status          = "published",
        )
        db.session.add(new_rev)
        db.session.commit()

        output = BytesIO()
        if doc_type == "Instrument List":
            write_workbook(payload, output)
            prefix = "Instrument_List"
        elif doc_type == "IO List":
            write_io_workbook(payload, output)
            prefix = "IO_List"
        else:
            return jsonify({"error": "Unknown document type."}), 400

        output.seek(0)
        loc_tag  = location.code.strip() if location.code else location.name[:12].replace(" ", "_")
        filename = (
            f"{project.display_name.replace(' ', '_')}"
            f"_{loc_tag}_{prefix}"
            f"_Rev{next_rev_num}"
            f"_{datetime.now().strftime('%Y%m%d')}.xlsx"
        )

        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype=(
                "application/vnd.openxmlformats-officedocument"
                ".spreadsheetml.sheet"
            ),
        )

    except Exception as e:
        traceback.print_exc()
        db.session.rollback()
        return jsonify({"error": str(e)}), 500


@cis_bp.route(
    "/project/<int:project_id>/revision/<int:rev_id>/download",
    methods=["GET"],
)
@login_required
def download_revision(project_id, rev_id):
    """
    Re-generate and stream an Excel file from a stored revision payload.
    Pure read — no DB write.
    """
    project = Project.query.get_or_404(project_id)
    rev     = DocumentRevision.query.filter_by(
        id=rev_id, project_id=project_id
    ).first_or_404()

    try:
        payload = rev.data_payload
        output  = BytesIO()

        if rev.doc_type == "Instrument List":
            write_workbook(payload, output)
            prefix = "Instrument_List"
        elif rev.doc_type == "IO List":
            write_io_workbook(payload, output)
            prefix = "IO_List"
        else:
            return jsonify({
                "error": f"Re-download not supported for '{rev.doc_type}'."
            }), 400

        output.seek(0)

        loc_tag = ""
        if rev.location:
            loc_tag = "_" + (
                rev.location.code.strip()
                if rev.location.code
                else rev.location.name[:12].replace(" ", "_")
            )

        filename = (
            f"{project.display_name.replace(' ', '_')}"
            f"{loc_tag}"
            f"_{prefix}"
            f"_Rev{rev.revision_number}"
            f"_{rev.created_at.strftime('%Y%m%d')}.xlsx"
        )

        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype=(
                "application/vnd.openxmlformats-officedocument"
                ".spreadsheetml.sheet"
            ),
        )

    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


# ─────────────────────────────────────────────────────────────────────────────
# Application factory
# ─────────────────────────────────────────────────────────────────────────────

def create_app():
    app = Flask(__name__)
    app.request_class = _BigRequest
    app.config.from_object(get_config())

    # Override upload/form-size limits explicitly here so stale
    # config .pyc files cannot shadow the values in config.py.
    # MAX_CONTENT_LENGTH  — total POST body Werkzeug will accept.
    # MAX_FORM_MEMORY_SIZE — per-form-field memory cap (Werkzeug 3.x
    #                        defaults this to 500 KB, which is too small
    #                        for large grids).
    # MAX_FORM_PARTS       — max number of form fields (Werkzeug 3.x
    #                        defaults to 1 000; large grids exceed this).
    app.config.setdefault("MAX_CONTENT_LENGTH",   16 * 1024 * 1024)
    app.config.setdefault("MAX_FORM_MEMORY_SIZE", 16 * 1024 * 1024)
    app.config.setdefault("MAX_FORM_PARTS",       10_000)

    db.init_app(app)
    login_manager.init_app(app)
    csrf.init_app(app)
    limiter.init_app(app)

    @login_manager.user_loader
    def load_user(user_id):
        return db.session.get(User, int(user_id))

    app.register_blueprint(auth_bp)
    app.register_blueprint(admin_bp)
    app.register_blueprint(cis_bp)

    with app.app_context():
        db.create_all()

    @app.errorhandler(403)
    def forbidden(e):
        return render_template("errors/403.html"), 403

    @app.errorhandler(404)
    def not_found(e):
        return render_template("errors/404.html"), 404

    @app.errorhandler(500)
    def server_error(e):
        return render_template("errors/500.html"), 500

    return app


app = create_app()

if __name__ == "__main__":
    app.run(debug=(os.environ.get("FLASK_ENV") == "development"))
