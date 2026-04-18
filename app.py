"""
app.py
──────
Application factory. Auth lives in the `auth` Blueprint.
All CIS data routes require login. Admin dashboard requires role=admin.
Now features Centralized Project Repositories and Document Revisions.
"""

import os
import functools
import traceback
from datetime import datetime, timezone
from io import BytesIO

from flask import (
    Flask, Blueprint, render_template, request,
    redirect, url_for, flash, jsonify, send_file,
    abort, session
)
from flask_login import (
    login_user, logout_user, login_required,
    current_user
)

from config      import get_config
from extensions  import db, login_manager, csrf, limiter
from models      import User, Project, DocumentRevision
from forms       import LoginForm, RegisterForm, EditUserForm, ProjectForm
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
    users = User.query.order_by(User.created_at.desc()).all()
    projects = Project.query.order_by(Project.created_at.desc()).all()
    user_form = RegisterForm()
    project_form = ProjectForm()
    return render_template("admin.html", users=users, projects=projects, 
                           user_form=user_form, project_form=project_form)

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
    return redirect(url_for("admin.dashboard"))

@admin_bp.route("/projects/create", methods=["POST"])
@admin_required
def create_project():
    form = ProjectForm()
    if form.validate_on_submit():
        project = Project(
            name       = form.name.data,
            client     = form.client.data,
            consultant = form.consultant.data,
            location   = form.location.data
        )
        db.session.add(project)
        db.session.commit()
        flash(f"Project '{project.name}' created successfully.", "success")
    else:
        for errs in form.errors.values():
            for err in errs:
                flash(err, "danger")
    return redirect(url_for("admin.dashboard"))


# ─────────────────────────────────────────────────────────────────────────────
# CIS Blueprint (Project-centric now)
# ─────────────────────────────────────────────────────────────────────────────

cis_bp = Blueprint("cis", __name__)

# ── Payload Builders ──────────────────────────────────────────────────────────
def _clean(v): return v.strip() if isinstance(v, str) else v
def _get_list(form, key): return [_clean(v) for v in form.getlist(key)]
def _unpack_pfl(packed, idx):
    raw   = packed[idx] if idx < len(packed) else ""
    parts = raw.split("/")
    parts += [""] * (3 - len(parts))
    return parts[:3]

def _build_fi(form):
    tags         = _get_list(form, "fiTagNo[]")
    instruments  = _get_list(form, "fiInstrument[]")
    # ... (Keep all your existing payload extraction logic exactly the same) ...
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
            "Medium":              mediums[i]      if i < len(mediums)      else "",
            "Specification":       specs[i]        if i < len(specs)        else "",
            "Process Conn":        conns[i]        if i < len(conns)        else "",
            "Work Press":  w[0], "Work Flow": w[1], "Work Level": w[2],
            "Design Press":d[0], "Design Flow":d[1],"Design Level":d[2],
            "Set-point":           set_points[i]   if i < len(set_points)   else "",
            "Range":               ranges[i]       if i < len(ranges)       else "",
            "UOM":                 uoms[i]         if i < len(uoms)         else "",
            "Signal Type":         sig_types[i]    if i < len(sig_types)    else "",
            "Source":              sources[i]      if i < len(sources)      else "",
            "Destination":         destinations[i] if i < len(destinations) else "",
            "Signal":              signals[i]      if i < len(signals)      else "",
        }
        if any(row.values()): rows.append(row)
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
            "Tag No": tags[i] if i < len(tags) else "",
            "Instrument Name": instruments[i] if i < len(instruments) else "",
            "Service Description": services[i] if i < len(services) else "",
            "Signal Type": sig_types[i] if i < len(sig_types) else "",
            "Source": sources[i] if i < len(sources) else "",
            "Destination": destinations[i] if i < len(destinations) else "",
            "Signal Description": sig_descs[i] if i < len(sig_descs) else "",
            "Signal": signals[i] if i < len(signals) else "",
        }
        if any(row.values()): rows.append(row)
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
    """Display list of projects user can access."""
    projects = Project.query.order_by(Project.created_at.desc()).all()
    return render_template("project_list.html", projects=projects)

@cis_bp.route("/project/<int:project_id>")
@login_required
def project_dashboard(project_id):
    """View a specific project and its document history."""
    project = Project.query.get_or_404(project_id)
    revisions = DocumentRevision.query.filter_by(project_id=project.id).order_by(DocumentRevision.created_at.desc()).all()
    return render_template("project_dashboard.html", project=project, revisions=revisions)

@cis_bp.route("/project/<int:project_id>/edit-docs")
@login_required
def edit_docs(project_id):
    """The main data entry interface, pre-loaded with project details."""
    project = Project.query.get_or_404(project_id)
    # Here you could potentially fetch the latest payload from DB and pass it to template
    # to pre-fill the form, fulfilling the "leverage later" requirement.
    latest_rev = DocumentRevision.query.filter_by(project_id=project.id).order_by(DocumentRevision.id.desc()).first()
    previous_data = latest_rev.data_payload if latest_rev else None
    
    return render_template("index.html", project=project, previous_data=previous_data)

@cis_bp.route("/project/<int:project_id>/submit-doc/<doc_type>", methods=["POST"])
@login_required
def submit_and_save(project_id, doc_type):
    """Saves the form data as a revision, then downloads the requested document."""
    try:
        project = Project.query.get_or_404(project_id)
        payload = _build_payload(request.form)
        
        # Determine Revision Number
        latest_rev = DocumentRevision.query.filter_by(
            project_id=project.id, doc_type=doc_type
        ).order_by(DocumentRevision.revision_number.desc()).first()
        
        next_rev_num = (latest_rev.revision_number + 1) if latest_rev else 0
        
        # Save payload to database
        new_rev = DocumentRevision(
            project_id=project.id,
            user_id=current_user.id,
            doc_type=doc_type,
            revision_number=next_rev_num,
            data_payload=payload
        )
        db.session.add(new_rev)
        db.session.commit()

        # Generate the Excel File
        output = BytesIO()
        if doc_type == "Instrument List":
            write_workbook(payload, output)
            prefix = "Instrument_List"
        elif doc_type == "IO List":
            write_io_workbook(payload, output)
            prefix = "IO_List"
        else:
            return jsonify({"success": False, "error": "Unknown document type"}), 400

        output.seek(0)
        filename = f"{project.name.replace(' ', '_')}_{prefix}_Rev{next_rev_num}_{datetime.now().strftime('%Y%m%d')}.xlsx"
        
        return send_file(
            output, as_attachment=True, download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        traceback.print_exc()
        db.session.rollback()
        return jsonify({"success": False, "error": str(e)}), 500


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
        return User.query.get(int(user_id))

    app.register_blueprint(auth_bp)
    app.register_blueprint(admin_bp)
    app.register_blueprint(cis_bp)

    with app.app_context():
        db.create_all()

    @app.errorhandler(403)
    def forbidden(e): return render_template("errors/403.html"), 403
    @app.errorhandler(404)
    def not_found(e): return render_template("errors/404.html"), 404
    @app.errorhandler(500)
    def server_error(e): return render_template("errors/500.html"), 500

    return app

app = create_app()

if __name__ == "__main__":
    app.run(debug=(os.environ.get("FLASK_ENV") == "development"))