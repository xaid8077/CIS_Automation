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
from forms      import (
    LoginForm, RegisterForm, EditUserForm,
    ProjectForm, ProjectNicknameForm, ProjectLocationForm,
)
from utils.excel_writer import write_workbook, write_io_workbook
from utils.validator    import validate_payload


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


@admin_bp.route("/")
@admin_required
def dashboard():
    users         = User.query.order_by(User.created_at.desc()).all()
    projects      = Project.query.order_by(Project.created_at.desc()).all()
    user_form     = RegisterForm()
    project_form  = ProjectForm()
    nickname_form = ProjectNicknameForm()
    loc_form      = ProjectLocationForm()
    return render_template(
        "admin.html",
        users=users,
        projects=projects,
        user_form=user_form,
        project_form=project_form,
        nickname_form=nickname_form,
        loc_form=loc_form,
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


# ─────────────────────────────────────────────────────────────────────────────
# CIS Blueprint
# ─────────────────────────────────────────────────────────────────────────────

cis_bp = Blueprint("cis", __name__)


# ── Payload builders ──────────────────────────────────────────────────────────

def _clean(v):
    return v.strip() if isinstance(v, str) else v

def _get_list(form, key):
    return [_clean(v) for v in form.getlist(key)]

def _unpack_pfl(packed, idx):
    raw   = packed[idx] if idx < len(packed) else ""
    parts = raw.split("/")
    parts += [""] * (3 - len(parts))
    return parts[:3]


def _build_fi(form):
    tags         = _get_list(form, "fiTagNo[]")
    instruments  = _get_list(form, "fiInstrument[]")
    services     = _get_list(form, "fiServiceDescription[]")
    line_sizes   = _get_list(form, "fiLineSize[]")
    mediums      = _get_list(form, "fiMedium[]")
    specs        = _get_list(form, "fiTypeSpec[]")
    conns        = _get_list(form, "fiProcessConnection[]")
    working_vals = _get_list(form, "fiWorkingValues[]")
    design_vals  = _get_list(form, "fiDesignValues[]")
    set_points   = _get_list(form, "fiSetPoint[]")
    ranges       = _get_list(form, "fiInstrumentRange[]")
    uoms         = _get_list(form, "fiUom[]")
    sig_types    = _get_list(form, "fiSignalType[]")
    sources      = _get_list(form, "fiSource[]")
    destinations = _get_list(form, "fiDestination[]")
    signals      = _get_list(form, "fiSignal[]")

    rows = []
    for i in range(len(tags)):
        w = _unpack_pfl(working_vals, i)
        d = _unpack_pfl(design_vals,  i)
        row = {
            "Tag No":              tags[i]         if i < len(tags)         else "",
            "Instrument Name":     instruments[i]  if i < len(instruments)  else "",
            "Service Description": services[i]     if i < len(services)     else "",
            "Line Size":           line_sizes[i]   if i < len(line_sizes)   else "",
            "Medium":              mediums[i]       if i < len(mediums)      else "",
            "Specification":       specs[i]         if i < len(specs)        else "",
            "Process Conn":        conns[i]         if i < len(conns)        else "",
            "Work Press":  w[0], "Work Flow": w[1], "Work Level": w[2],
            "Design Press":d[0], "Design Flow":d[1], "Design Level":d[2],
            "Set-point":           set_points[i]   if i < len(set_points)   else "",
            "Range":               ranges[i]        if i < len(ranges)       else "",
            "UOM":                 uoms[i]          if i < len(uoms)         else "",
            "Signal Type":         sig_types[i]    if i < len(sig_types)    else "",
            "Source":              sources[i]       if i < len(sources)      else "",
            "Destination":         destinations[i]  if i < len(destinations) else "",
            "Signal":              signals[i]       if i < len(signals)      else "",
        }
        if any(row.values()):
            rows.append(row)
    return rows


def _build_flat(form, prefix):
    tags         = _get_list(form, f"{prefix}TagNo[]")
    instruments  = _get_list(form, f"{prefix}Instrument[]")
    services     = _get_list(form, f"{prefix}ServiceDescription[]")
    sig_types    = _get_list(form, f"{prefix}SignalType[]")
    sources      = _get_list(form, f"{prefix}Source[]")
    destinations = _get_list(form, f"{prefix}Destination[]")
    sig_descs    = _get_list(form, f"{prefix}SigDesc[]")
    signals      = _get_list(form, f"{prefix}Signal[]")

    rows = []
    for i in range(len(tags)):
        row = {
            "Tag No":              tags[i]         if i < len(tags)         else "",
            "Instrument Name":     instruments[i]  if i < len(instruments)  else "",
            "Service Description": services[i]     if i < len(services)     else "",
            "Signal Type":         sig_types[i]    if i < len(sig_types)    else "",
            "Source":              sources[i]       if i < len(sources)      else "",
            "Destination":         destinations[i]  if i < len(destinations) else "",
            "Signal Description":  sig_descs[i]    if i < len(sig_descs)    else "",
            "Signal":              signals[i]       if i < len(signals)      else "",
        }
        if any(row.values()):
            rows.append(row)
    return rows


def _build_header(form):
    return {k: form.get(k, "") for k in [
        "projectName", "client", "consultant", "location",
        "date", "preparedBy", "checkedBy", "approvedBy", "revision",
    ]}


def _build_section_meta(form, prefix):
    return {"docNumber": _clean(form.get(f"{prefix}DocNumber", ""))}


def _build_payload(form):
    return {
        "header":            _build_header(form),
        "fi_meta":           _build_section_meta(form, "fi"),
        "el_meta":           _build_section_meta(form, "el"),
        "mov_meta":          _build_section_meta(form, "mov"),
        "io_meta":           _build_section_meta(form, "io"),
        "field_instruments": _build_fi(form),
        "electrical":        _build_flat(form, "el"),
        "mov":               _build_flat(form, "mov"),
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
        payload = _build_payload(request.form)
        errors  = validate_payload(payload)
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
        payload = _build_payload(request.form)

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

        payload = _build_payload(request.form)

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
    app.config.from_object(get_config())

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