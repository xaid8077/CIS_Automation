# app.py
"""
app.py
──────
Application factory — wiring only.

All business logic, route handlers, and service code live in their
own modules.  This file's only job is to create the Flask app,
attach extensions, register blueprints, and wire error handlers.

Request size override
─────────────────────
Werkzeug 3.x enforces form-size limits as class-level attributes on
the Request object, independent of Flask config.  _BigRequest overrides
them so large JSON payloads are never rejected with a silent 413.
(JSON bodies are not subject to MAX_FORM_MEMORY_SIZE, but the override
is kept here as a defence-in-depth measure for any multipart routes
added in the future.)
"""

import os

from flask import Flask, Request as _FlaskRequest, render_template

from config     import get_config
from extensions import db, login_manager, csrf, limiter
from models     import User

# ── Blueprints ─────────────────────────────────────────────────────────────────
from routes.auth  import auth_bp
from routes.admin import admin_bp
from routes.cis   import cis_bp


# ─── Request subclass ─────────────────────────────────────────────────────────

class _BigRequest(_FlaskRequest):
    """Raise Werkzeug's built-in per-request size caps to 16 MB."""
    max_content_length   = 16 * 1024 * 1024
    max_form_memory_size = 16 * 1024 * 1024
    max_form_parts       = 10_000


# ─── Factory ──────────────────────────────────────────────────────────────────

def create_app() -> Flask:
    app = Flask(__name__)
    app.request_class = _BigRequest
    app.config.from_object(get_config())

    # Belt-and-suspenders: set size limits in config too so anything
    # reading from app.config (e.g. older Werkzeug middleware) agrees.
    app.config.setdefault("MAX_CONTENT_LENGTH",   16 * 1024 * 1024)
    app.config.setdefault("MAX_FORM_MEMORY_SIZE", 16 * 1024 * 1024)
    app.config.setdefault("MAX_FORM_PARTS",       10_000)

    # ── Extensions ────────────────────────────────────────────────────────────
    db.init_app(app)
    login_manager.init_app(app)
    csrf.init_app(app)
    limiter.init_app(app)

    @login_manager.user_loader
    def load_user(user_id: str):
        return db.session.get(User, int(user_id))

    # ── Blueprints ─────────────────────────────────────────────────────────────
    app.register_blueprint(auth_bp)
    app.register_blueprint(admin_bp)
    app.register_blueprint(cis_bp)

    # ── Database ───────────────────────────────────────────────────────────────
    with app.app_context():
        db.create_all()

    # ── Error handlers ─────────────────────────────────────────────────────────
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


# ─── Entry point ──────────────────────────────────────────────────────────────

app = create_app()

if __name__ == "__main__":
    app.run(debug=(os.environ.get("FLASK_ENV") == "development"))