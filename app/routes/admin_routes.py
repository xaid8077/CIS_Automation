# app/routes/admin_routes.py

from flask import Blueprint, render_template, redirect, url_for, flash
from flask_login import current_user

from app.extensions import db
from app.models import User, Project, ProjectLocation, DocumentRevision
from app.forms import (
    RegisterForm, ProjectForm, ProjectNicknameForm,
    EditProjectNameForm, ProjectLocationForm,
    RevisionDocNumbersForm,
)
from app.utils.decorators import admin_required


admin_bp = Blueprint("admin", __name__, url_prefix="/admin")


# ─────────────────────────────────────────────────────────────────────────────
# Dashboard
# ─────────────────────────────────────────────────────────────────────────────

@admin_bp.route("/")
@admin_required
def dashboard():
    users    = User.query.order_by(User.created_at.desc()).all()
    projects = Project.query.order_by(Project.created_at.desc()).all()

    return render_template(
        "admin.html",
        users=users,
        projects=projects,
        user_form=RegisterForm(),
        project_form=ProjectForm(),
        nickname_form=ProjectNicknameForm(),
        edit_project_form=EditProjectNameForm(),
        loc_form=ProjectLocationForm(),
        rev_doc_form=RevisionDocNumbersForm(),
    )


# ─────────────────────────────────────────────────────────────────────────────
# User Management
# ─────────────────────────────────────────────────────────────────────────────

@admin_bp.route("/users/create", methods=["POST"])
@admin_required
def create_user():
    form = RegisterForm()

    if form.validate_on_submit():
        user = User(
            username=form.username.data,
            email=form.email.data,
            role=form.role.data,
            is_active=True,
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


# ─────────────────────────────────────────────────────────────────────────────
# Project Management
# ─────────────────────────────────────────────────────────────────────────────

@admin_bp.route("/projects/create", methods=["POST"])
@admin_required
def create_project():
    form = ProjectForm()

    if form.validate_on_submit():
        project = Project(
            name=form.name.data.strip(),
            nickname=form.nickname.data.strip() if form.nickname.data else None,
            client=form.client.data.strip(),
            consultant=form.consultant.data.strip() if form.consultant.data else None,
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
    project = Project.query.get_or_404(project_id)
    form = ProjectNicknameForm()

    if form.validate_on_submit():
        raw = form.nickname.data.strip() if form.nickname.data else ""
        project.nickname = raw if raw else None

        db.session.commit()

        flash(
            f"Nickname for '{project.name}' updated to '{project.display_name}'.",
            "success",
        )
    else:
        for errs in form.errors.values():
            for err in errs:
                flash(err, "danger")

    return redirect(url_for("admin.dashboard"))


# ─────────────────────────────────────────────────────────────────────────────
# Location Management
# ─────────────────────────────────────────────────────────────────────────────

@admin_bp.route("/projects/<int:project_id>/locations/add", methods=["POST"])
@admin_required
def add_location(project_id):
    project = Project.query.get_or_404(project_id)
    form = ProjectLocationForm()

    if form.validate_on_submit():
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
                project_id=project.id,
                name=form.name.data.strip(),
                code=form.code.data.strip() if form.code.data else None,
            )

            db.session.add(loc)
            db.session.commit()

            flash(
                f"Location '{loc.display}' added to '{project.display_name}'.",
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
    loc = ProjectLocation.query.filter_by(
        id=loc_id, project_id=project_id
    ).first_or_404()

    if loc.revisions:
        flash(
            f"Cannot delete '{loc.display}' — it has "
            f"{len(loc.revisions)} revision(s) attached.",
            "danger",
        )
    else:
        name = loc.display
        db.session.delete(loc)
        db.session.commit()

        flash(f"Location '{name}' deleted.", "success")

    return redirect(url_for("admin.dashboard"))


# ─────────────────────────────────────────────────────────────────────────────
# Project Editing
# ─────────────────────────────────────────────────────────────────────────────

@admin_bp.route("/projects/<int:project_id>/edit", methods=["POST"])
@admin_required
def edit_project(project_id):
    project = Project.query.get_or_404(project_id)
    form = EditProjectNameForm()

    if form.validate_on_submit():
        new_name = form.name.data.strip()

        if new_name != project.name:
            clash = Project.query.filter_by(name=new_name).first()
            if clash:
                flash(
                    f"Another project already has the name '{new_name}'.",
                    "danger",
                )
                return redirect(url_for("admin.dashboard"))

        project.name = new_name
        project.nickname = (
            form.nickname.data.strip() or None
            if form.nickname.data else None
        )
        project.client = form.client.data.strip()
        project.consultant = (
            form.consultant.data.strip() or None
            if form.consultant.data else None
        )

        db.session.commit()

        flash(f"Project updated: '{project.display_name}'.", "success")
    else:
        for errs in form.errors.values():
            for err in errs:
                flash(err, "danger")

    return redirect(url_for("admin.dashboard"))


# ─────────────────────────────────────────────────────────────────────────────
# Revision Editing
# ─────────────────────────────────────────────────────────────────────────────

@admin_bp.route(
    "/projects/<int:project_id>/revisions/<int:rev_id>/doc-numbers",
    methods=["POST"],
)
@admin_required
def edit_revision_doc_numbers(project_id, rev_id):
    rev = DocumentRevision.query.filter_by(
        id=rev_id, project_id=project_id
    ).first_or_404()

    form = RevisionDocNumbersForm()

    if form.validate_on_submit():
        payload = dict(rev.data_payload)

        def patch(key, value):
            if value is not None:
                meta = dict(payload.get(key) or {})
                meta["docNumber"] = value.strip()
                payload[key] = meta

        patch("fi_meta", form.fi_doc_number.data)
        patch("el_meta", form.el_doc_number.data)
        patch("mov_meta", form.mov_doc_number.data)
        patch("io_meta", form.io_doc_number.data)

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