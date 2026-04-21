# api/v1/projects.py
"""
api/v1/projects.py
──────────────────
Project and location CRUD endpoints.

GET    /api/v1/projects/                       view_projects
GET    /api/v1/projects/<id>                   view_projects
POST   /api/v1/projects/                       manage_projects  (admin)
PATCH  /api/v1/projects/<id>                   manage_projects  (admin)

GET    /api/v1/projects/<id>/locations/        view_projects
POST   /api/v1/projects/<id>/locations/        manage_locations (admin)
DELETE /api/v1/projects/<id>/locations/<lid>   manage_locations (admin)

Design notes
────────────
- All write operations require the manage_projects / manage_locations
  permission (admin only per the permission matrix in models.py).
- Read operations require view_projects (all roles).
- Input is validated with inline Marshmallow schemas — lightweight
  because project fields are simple strings.
- No partial-update schema is used for PATCH; every editable field is
  optional, and only supplied fields are applied to the model.
"""

from flask import Blueprint, jsonify, request
from marshmallow import Schema, fields, ValidationError, EXCLUDE, validates

from extensions import db
from models import Project, ProjectLocation
from utils.rbac import api_permission_required

projects_bp = Blueprint("projects_api", __name__, url_prefix="/projects")


# ─── Input schemas ────────────────────────────────────────────────────────────

class ProjectCreateSchema(Schema):
    class Meta:
        unknown = EXCLUDE

    name       = fields.Str(required=True)
    nickname   = fields.Str(load_default=None)
    client     = fields.Str(required=True)
    consultant = fields.Str(load_default=None)

    @validates("name")
    def name_not_blank(self, value):
        if not value.strip():
            raise ValidationError("Project name cannot be blank.")

    @validates("client")
    def client_not_blank(self, value):
        if not value.strip():
            raise ValidationError("Client cannot be blank.")


class ProjectPatchSchema(Schema):
    """All fields optional — only supplied fields are applied."""
    class Meta:
        unknown = EXCLUDE

    name       = fields.Str(load_default=None)
    nickname   = fields.Str(load_default=None)
    client     = fields.Str(load_default=None)
    consultant = fields.Str(load_default=None)


class LocationCreateSchema(Schema):
    class Meta:
        unknown = EXCLUDE

    name = fields.Str(required=True)
    code = fields.Str(load_default=None)

    @validates("name")
    def name_not_blank(self, value):
        if not value.strip():
            raise ValidationError("Location name cannot be blank.")


_project_create_schema  = ProjectCreateSchema()
_project_patch_schema   = ProjectPatchSchema()
_location_create_schema = LocationCreateSchema()


# ─── Helper ───────────────────────────────────────────────────────────────────

def _optional(value) -> str | None:
    """Return stripped string or None when blank."""
    if value is None:
        return None
    stripped = str(value).strip()
    return stripped or None


# ─── Project endpoints ────────────────────────────────────────────────────────

@projects_bp.route("/", methods=["GET"])
@api_permission_required("view_projects")
def list_projects():
    """
    GET /api/v1/projects/
    ─────────────────────
    Response 200:
        [{ project dict with nested locations }, ...]
    """
    projects = (
        Project.query
        .order_by(Project.created_at.desc())
        .all()
    )
    return jsonify([p.to_dict() for p in projects]), 200


@projects_bp.route("/<int:project_id>", methods=["GET"])
@api_permission_required("view_projects")
def get_project(project_id):
    """
    GET /api/v1/projects/<id>
    ─────────────────────────
    Response 200: project dict with nested locations.
    Response 404: project not found.
    """
    project = Project.query.get_or_404(project_id)
    return jsonify(project.to_dict()), 200


@projects_bp.route("/", methods=["POST"])
@api_permission_required("manage_projects")
def create_project():
    """
    POST /api/v1/projects/
    ──────────────────────
    Body (JSON):
        {"name": "...", "client": "...", "nickname": "...", "consultant": "..."}

    Response 201: created project dict.
    Response 409: project name already exists.
    Response 422: validation error.
    """
    try:
        data = _project_create_schema.load(
            request.get_json(force=True, silent=True) or {}
        )
    except ValidationError as exc:
        return jsonify({"error": "Validation failed.", "detail": exc.messages}), 422

    name = data["name"].strip()
    if Project.query.filter_by(name=name).first():
        return jsonify({"error": f"A project named '{name}' already exists."}), 409

    project = Project(
        name       = name,
        nickname   = _optional(data.get("nickname")),
        client     = data["client"].strip(),
        consultant = _optional(data.get("consultant")),
    )
    db.session.add(project)
    db.session.commit()
    return jsonify(project.to_dict()), 201


@projects_bp.route("/<int:project_id>", methods=["PATCH"])
@api_permission_required("manage_projects")
def patch_project(project_id):
    """
    PATCH /api/v1/projects/<id>
    ───────────────────────────
    Body (JSON): any subset of {name, nickname, client, consultant}.
    Only supplied (non-null) fields are applied.

    Response 200: updated project dict.
    Response 409: name conflict with another project.
    Response 422: validation error.
    """
    project = Project.query.get_or_404(project_id)

    try:
        data = _project_patch_schema.load(
            request.get_json(force=True, silent=True) or {}
        )
    except ValidationError as exc:
        return jsonify({"error": "Validation failed.", "detail": exc.messages}), 422

    if data.get("name") is not None:
        new_name = data["name"].strip()
        if new_name and new_name != project.name:
            if Project.query.filter_by(name=new_name).first():
                return jsonify({"error": f"A project named '{new_name}' already exists."}), 409
            project.name = new_name

    if data.get("nickname") is not None:
        project.nickname = _optional(data["nickname"])

    if data.get("client") is not None:
        client = data["client"].strip()
        if client:
            project.client = client

    if data.get("consultant") is not None:
        project.consultant = _optional(data["consultant"])

    db.session.commit()
    return jsonify(project.to_dict()), 200


# ─── Location endpoints ───────────────────────────────────────────────────────

@projects_bp.route("/<int:project_id>/locations/", methods=["GET"])
@api_permission_required("view_projects")
def list_locations(project_id):
    """
    GET /api/v1/projects/<id>/locations/
    ─────────────────────────────────────
    Response 200: list of location dicts for the project.
    """
    project   = Project.query.get_or_404(project_id)
    locations = (
        ProjectLocation.query
        .filter_by(project_id=project.id)
        .order_by(ProjectLocation.name)
        .all()
    )
    return jsonify([loc.to_dict() for loc in locations]), 200


@projects_bp.route("/<int:project_id>/locations/", methods=["POST"])
@api_permission_required("manage_locations")
def add_location(project_id):
    """
    POST /api/v1/projects/<id>/locations/
    ──────────────────────────────────────
    Body (JSON):
        {"name": "Plant A – Intake", "code": "PA-INT"}

    Response 201: created location dict.
    Response 409: location name already exists for this project.
    Response 422: validation error.
    """
    project = Project.query.get_or_404(project_id)

    try:
        data = _location_create_schema.load(
            request.get_json(force=True, silent=True) or {}
        )
    except ValidationError as exc:
        return jsonify({"error": "Validation failed.", "detail": exc.messages}), 422

    name = data["name"].strip()
    if ProjectLocation.query.filter_by(project_id=project.id, name=name).first():
        return jsonify({
            "error": f"Location '{name}' already exists in this project."
        }), 409

    loc = ProjectLocation(
        project_id = project.id,
        name       = name,
        code       = _optional(data.get("code")),
    )
    db.session.add(loc)
    db.session.commit()
    return jsonify(loc.to_dict()), 201


@projects_bp.route(
    "/<int:project_id>/locations/<int:loc_id>",
    methods=["DELETE"],
)
@api_permission_required("manage_locations")
def delete_location(project_id, loc_id):
    """
    DELETE /api/v1/projects/<id>/locations/<lid>
    ─────────────────────────────────────────────
    Blocked if any revisions reference this location.

    Response 204: deleted successfully (no body).
    Response 409: location has attached revisions.
    """
    loc = ProjectLocation.query.filter_by(
        id=loc_id, project_id=project_id
    ).first_or_404()

    if loc.revisions:
        return jsonify({
            "error": (
                f"Cannot delete '{loc.display}' — "
                f"{len(loc.revisions)} revision(s) are attached. "
                "Delete the revisions first."
            )
        }), 409

    db.session.delete(loc)
    db.session.commit()
    return "", 204