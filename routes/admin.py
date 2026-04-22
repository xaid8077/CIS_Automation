# routes/admin.py
"""
admin.py
────────
Admin blueprint — project, location, user, and revision management.

Design notes
────────────
- Every route is guarded by @admin_required which combines
  @login_required + an is_admin check in one decorator.
- No business logic lives here — routes validate the form, call
  a service or perform a direct DB write (for simple CRUD that
  does not warrant a service), then redirect with a flash message.
- Revision doc-number patching delegates to revision_service so the
  JSON-column mutation gotcha is handled in one place.
- All modal open/close logic stays in the Jinja template (admin.html)
  so this file stays free of any frontend concerns.
"""

import functools

from flask import (
    Blueprint, render_template, request,
    redirect, url_for, flash, abort,
)
from flask_login import login_required, current_user

from extensions import db
from forms import (
    RegisterForm, ProjectForm, EditProjectNameForm,
    ProjectNicknameForm, ProjectLocationForm, RevisionDocNumbersForm,
)
from models import User, Project, ProjectLocation, DocumentRevision
from services import revision_service

admin_bp = Blueprint("admin", __name__, url_prefix="/admin")


# ─── Decorator ────────────────────────────────────────────────────────────────

def admin_required(f):
    """Combine @login_required + admin role check into one decorator."""
    @functools.wraps(f)
    @login_required
    def wrapped(*args, **kwargs):
        if not current_user.is_admin:
            abort(403)
        return f(*args, **kwargs)
    return wrapped


# ─── Dashboard ────────────────────────────────────────────────────────────────

@admin_bp.route("/")
@admin_required
def dashboard():
    return render_template(
        "admin.html",
        users             = User.query.order_by(User.created_at.desc()).all(),
        projects          = Project.query.order_by(Project.created_at.desc()).all(),
        user_form         = RegisterForm(),
        project_form      = ProjectForm(),
        nickname_form     = ProjectNicknameForm(),
        edit_project_form = EditProjectNameForm(),
        loc_form          = ProjectLocationForm(),
        rev_doc_form      = RevisionDocNumbersForm(),
    )


# ─── User management ──────────────────────────────────────────────────────────

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
        _flash_form_errors(form)
    return redirect(url_for("admin.dashboard"))


# ─── Project management ───────────────────────────────────────────────────────

@admin_bp.route("/projects/create", methods=["POST"])
@admin_required
def create_project():
    form = ProjectForm()
    if form.validate_on_submit():
        project = Project(
            name       = form.name.data.strip(),
            nickname   = _optional(form.nickname.data),
            client     = form.client.data.strip(),
            consultant = _optional(form.consultant.data),
        )
        db.session.add(project)
        db.session.commit()
        flash(f"Project '{project.display_name}' created.", "success")
    else:
        _flash_form_errors(form)
    return redirect(url_for("admin.dashboard"))


@admin_bp.route("/projects/<int:project_id>/edit", methods=["POST"])
@admin_required
def edit_project(project_id):
    """
    Edit a project's full name, nickname, client, and consultant.

    Uniqueness is only enforced when the name actually changes so that
    submitting an unchanged name is never rejected.
    """
    project = Project.query.get_or_404(project_id)
    form    = EditProjectNameForm()

    if form.validate_on_submit():
        new_name = form.name.data.strip()

        if new_name != project.name:
            if Project.query.filter_by(name=new_name).first():
                flash(
                    f"Another project already has the name '{new_name}'.",
                    "danger",
                )
                return redirect(url_for("admin.dashboard"))

        project.name       = new_name
        project.nickname   = _optional(form.nickname.data)
        project.client     = form.client.data.strip()
        project.consultant = _optional(form.consultant.data)
        db.session.commit()
        flash(f"Project updated: '{project.display_name}'.", "success")
    else:
        _flash_form_errors(form)

    return redirect(url_for("admin.dashboard"))


@admin_bp.route("/projects/<int:project_id>/nickname", methods=["POST"])
@admin_required
def set_nickname(project_id):
    """Set or clear a project's short display name."""
    project = Project.query.get_or_404(project_id)
    form    = ProjectNicknameForm()

    if form.validate_on_submit():
        project.nickname = _optional(form.nickname.data)
        db.session.commit()
        flash(
            f"Nickname for '{project.name}' updated to "
            f"'{project.display_name}'.",
            "success",
        )
    else:
        _flash_form_errors(form)

    return redirect(url_for("admin.dashboard"))


# ─── Location management ──────────────────────────────────────────────────────

@admin_bp.route("/projects/<int:project_id>/locations/add", methods=["POST"])
@admin_required
def add_location(project_id):
    """Add a new location to a project."""
    project = Project.query.get_or_404(project_id)
    form    = ProjectLocationForm()

    if form.validate_on_submit():
        name = form.name.data.strip()
        if ProjectLocation.query.filter_by(project_id=project.id, name=name).first():
            flash(
                f"Location '{name}' already exists in "
                f"project '{project.display_name}'.",
                "warning",
            )
        else:
            loc = ProjectLocation(
                project_id = project.id,
                name       = name,
                code       = _optional(form.code.data),
            )
            db.session.add(loc)
            db.session.commit()
            flash(
                f"Location '{loc.display}' added to "
                f"'{project.display_name}'.",
                "success",
            )
    else:
        _flash_form_errors(form)

    return redirect(url_for("admin.dashboard"))


@admin_bp.route(
    "/projects/<int:project_id>/locations/<int:loc_id>/delete",
    methods=["POST"],
)
@admin_required
def delete_location(project_id, loc_id):
    """
    Delete a location.

    Blocked if any revisions reference it — orphaned revision records
    would lose their location FK and become unattributable.
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


# ─── Revision doc-number patch ────────────────────────────────────────────────

@admin_bp.route(
    "/projects/<int:project_id>/revisions/<int:rev_id>/doc-numbers",
    methods=["POST"],
)
@admin_required
def edit_revision_doc_numbers(project_id, rev_id):
    """
    Patch document numbers stored inside a revision's payload.

    Does NOT regenerate the Excel file — only the stored JSON is updated.
    Useful for correcting a doc number after a revision has been published.
    """
    rev  = DocumentRevision.query.filter_by(
        id=rev_id, project_id=project_id
    ).first_or_404()
    form = RevisionDocNumbersForm()

    if form.validate_on_submit():
        revision_service.patch_doc_numbers(
            rev,
            fi_doc_number  = form.fi_doc_number.data,
            el_doc_number  = form.el_doc_number.data,
            mov_doc_number = form.mov_doc_number.data,
            io_doc_number  = form.io_doc_number.data,
        )
        flash(
            f"Document numbers updated for revision {rev.revision_number} "
            f"({rev.doc_type}).",
            "success",
        )
    else:
        _flash_form_errors(form)

    return redirect(url_for("admin.dashboard"))


# ─── Helpers ──────────────────────────────────────────────────────────────────

def _optional(value) -> str | None:
    """Return stripped value or None when blank — for nullable DB columns."""
    if not value:
        return None
    stripped = value.strip()
    return stripped if stripped else None


def _flash_form_errors(form) -> None:
    """Flash every WTForms validation error as a danger message."""
    for errs in form.errors.values():
        for err in errs:
            flash(err, "danger")