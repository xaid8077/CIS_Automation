# api/v1/__init__.py
"""
api/v1/__init__.py
──────────────────
Registers all v1 sub-blueprints onto a single parent blueprint.

The parent blueprint is mounted at /api/v1 in app.py.
Each sub-blueprint owns its own URL prefix beneath that root:

    /api/v1/auth/        ← JWT login, refresh, me
    /api/v1/projects/    ← project + location CRUD
    /api/v1/revisions/   ← revision list, draft save
    /api/v1/documents/   ← generate + re-download

Adding a new resource group = create a new file, import it here,
register it on api_v1_bp.  Nothing else needs to change.
"""

from flask import Blueprint

from api.v1.auth      import auth_api_bp
from api.v1.projects  import projects_bp
from api.v1.revisions import revisions_bp
from api.v1.documents import documents_bp

# Parent blueprint — all v1 routes share this prefix.
api_v1_bp = Blueprint("api_v1", __name__, url_prefix="/api/v1")

api_v1_bp.register_blueprint(auth_api_bp)
api_v1_bp.register_blueprint(projects_bp)
api_v1_bp.register_blueprint(revisions_bp)
api_v1_bp.register_blueprint(documents_bp)