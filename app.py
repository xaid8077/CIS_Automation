"""
app.py
──────
Application factory.  Auth lives in the `auth` Blueprint.
All CIS data routes require login.  Admin dashboard requires role=admin.
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
from models      import User
from forms       import LoginForm, RegisterForm, EditUserForm
from utils.excel_writer import write_workbook
from utils.validator    import validate_payload


# ─────────────────────────────────────────────────────────────────────────────
# Decorators
# ─────────────────────────────────────────────────────────────────────────────

def admin_required(f):
    """Route decorator: 403 unless current_user.is_admin."""
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
@limiter.limit("10 per minute")          # brute-force throttle on login endpoint
def login():
    if current_user.is_authenticated:
        return redirect(url_for("cis.index"))

    form = LoginForm()
    if form.validate_on_submit():
        user = User.query.filter_by(username=form.username.data).first()

        # Constant-time path: always call check_password even on unknown user
        # to prevent username enumeration via timing.
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

        # Safe redirect: only allow relative paths
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
    form  = RegisterForm()
    return render_template("admin.html", users=users, form=form)


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
        for field_errors in form.errors.values():
            for err in field_errors:
                flash(err, "danger")
    return redirect(url_for("admin.dashboard"))


@admin_bp.route("/users/<int:user_id>/edit", methods=["POST"])
@admin_required
def edit_user(user_id):
    # db.get_or_404 is the SQLAlchemy 2.x–compatible replacement for
    # the legacy Model.query.get_or_404() call.
    user = db.get_or_404(User, user_id)
    if user.id == current_user.id:
        flash("You cannot edit your own account here.", "warning")
        return redirect(url_for("admin.dashboard"))
    form = EditUserForm()
    if form.validate_on_submit():
        user.role      = form.role.data
        user.is_active = form.is_active.data
        db.session.commit()
        flash(f"User '{user.username}' updated.", "success")
    return redirect(url_for("admin.dashboard"))


@admin_bp.route("/users/<int:user_id>/delete", methods=["POST"])
@admin_required
def delete_user(user_id):
    user = db.get_or_404(User, user_id)
    if user.id == current_user.id:
        flash("You cannot delete your own account.", "warning")
        return redirect(url_for("admin.dashboard"))
    db.session.delete(user)
    db.session.commit()
    flash(f"User '{user.username}' deleted.", "success")
    return redirect(url_for("admin.dashboard"))


# ─────────────────────────────────────────────────────────────────────────────
# CIS Blueprint  (all routes require login)
# ─────────────────────────────────────────────────────────────────────────────

cis_bp = Blueprint("cis", __name__)


# ── Helpers ───────────────────────────────────────────────────────────────────

def _clean(v):
    return v.strip() if isinstance(v, str) else v

def _get_list(form, key):
    return [_clean(v) for v in form.getlist(key)]

def _safe(lst, i):
    """Return lst[i] if in bounds, else empty string.

    All array fields sent by the browser are kept in sync by the JS
    serialiser, so out-of-bounds access should never happen in practice.
    This guard is retained for defensive correctness.
    """
    return lst[i] if i < len(lst) else ""

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
            "Tag No":              _safe(tags,         i),
            "Instrument Name":     _safe(instruments,  i),
            "Service Description": _safe(services,     i),
            "Line Size":           _safe(line_sizes,   i),
            "Medium":              _safe(mediums,       i),
            "Specification":       _safe(specs,        i),
            "Process Conn":        _safe(conns,        i),
            "Work Press":  w[0], "Work Flow": w[1], "Work Level": w[2],
            "Design Press":d[0], "Design Flow":d[1],"Design Level":d[2],
            "Set-point":           _safe(set_points,   i),
            "Range":               _safe(ranges,       i),
            "UOM":                 _safe(uoms,         i),
            "Signal Type":         _safe(sig_types,    i),
            "Source":              _safe(sources,      i),
            "Destination":         _safe(destinations, i),
            "Signal":              _safe(signals,      i),
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
            "Tag No":              _safe(tags,         i),
            "Instrument Name":     _safe(instruments,  i),
            "Service Description": _safe(services,     i),
            "Signal Type":         _safe(sig_types,    i),
            "Source":              _safe(sources,      i),
            "Destination":         _safe(destinations, i),
            "Signal Description":  _safe(sig_descs,    i),
            "Signal":              _safe(signals,      i),
        }
        if any(row.values()):
            rows.append(row)
    return rows

def _build_header(form):
    return {k: form.get(k, "") for k in [
        "projectName", "client", "consultant",
        "date", "preparedBy", "checkedBy", "approvedBy", "revision"
    ]}

def _build_section_meta(form, prefix):
    return {
        "docName":   _clean(form.get(f"{prefix}DocName",   "")),
        "docNumber": _clean(form.get(f"{prefix}DocNumber", "")),
    }

def _build_payload(form):
    return {
        "header":            _build_header(form),
        "fi_meta":           _build_section_meta(form, "fi"),
        "el_meta":           _build_section_meta(form, "el"),
        "mov_meta":          _build_section_meta(form, "mov"),
        "field_instruments": _build_fi(form),
        "electrical":        _build_flat(form, "el"),
        "mov":               _build_flat(form, "mov"),
    }


# ── Routes ────────────────────────────────────────────────────────────────────

@cis_bp.route("/")
@login_required
def index():
    return render_template("index.html")


@cis_bp.route("/preview", methods=["POST"])
@login_required
def preview():
    try:
        errors = validate_payload(_build_payload(request.form))
        if errors:
            return jsonify({"success": False, "errors": errors}), 400
        return jsonify({"success": True, "message": "Validation passed — no errors found."}), 200
    except Exception:
        traceback.print_exc()
        return jsonify({"success": False, "message": "Server error during validation."}), 500


@cis_bp.route("/submit", methods=["POST"])
@login_required
def submit():
    try:
        payload = _build_payload(request.form)
        output  = BytesIO()
        write_workbook(payload, output)
        output.seek(0)
        filename = f"Instrument_List_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        return send_file(
            output, as_attachment=True, download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception:
        traceback.print_exc()
        return jsonify({"success": False, "error": "Server error during workbook generation."}), 500


# Placeholder routes for future documents — return 501 until implemented
@cis_bp.route("/download/io-list", methods=["POST"])
@login_required
def download_io_list():
    return jsonify({"success": False, "message": "IO List generation not yet implemented."}), 501


@cis_bp.route("/download/cable-schedule", methods=["POST"])
@login_required
def download_cable_schedule():
    return jsonify({"success": False, "message": "Cable Schedule generation not yet implemented."}), 501


@cis_bp.route("/download/cable-interconnection", methods=["POST"])
@login_required
def download_cable_interconnection():
    return jsonify({"success": False, "message": "Cable Interconnection Schedule generation not yet implemented."}), 501


# ─────────────────────────────────────────────────────────────────────────────
# Application factory
# ─────────────────────────────────────────────────────────────────────────────

def create_app():
    app = Flask(__name__)
    app.config.from_object(get_config())

    # Init extensions
    db.init_app(app)
    login_manager.init_app(app)
    csrf.init_app(app)
    limiter.init_app(app)

    # User loader for Flask-Login.
    # db.session.get() is the SQLAlchemy 2.x API; Session.get() replaces
    # the legacy Query.get() which emits deprecation warnings in SA 1.4+.
    from models import User as _User
    @login_manager.user_loader
    def load_user(user_id):
        return db.session.get(_User, int(user_id))

    # Register blueprints
    app.register_blueprint(auth_bp)
    app.register_blueprint(admin_bp)
    app.register_blueprint(cis_bp)

    # Create tables (use Flask-Migrate in production for schema changes)
    with app.app_context():
        db.create_all()

    # 403 / 404 / 500 error handlers
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