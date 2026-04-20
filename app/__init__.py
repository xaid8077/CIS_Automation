# app/__init__.py

import os
from flask import Flask, render_template
from flask_login import LoginManager

from app.extensions import db, login_manager, csrf, limiter
from app.config import get_config
import logging


# ─────────────────────────────────────────────────────────────────────────────
# Custom Request Class (moved cleanly here)
# ─────────────────────────────────────────────────────────────────────────────

from flask import Request as _FlaskRequest

class BigRequest(_FlaskRequest):
    max_content_length   = 16 * 1024 * 1024
    max_form_memory_size = 16 * 1024 * 1024
    max_form_parts       = 10_000


# ─────────────────────────────────────────────────────────────────────────────
# Application Factory
# ─────────────────────────────────────────────────────────────────────────────

def create_app():
    app = Flask(__name__)
    app.request_class = BigRequest

    logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s"
    )

    # Load config
    app.config.from_object(get_config())

    # Force limits (override Werkzeug defaults)
    app.config.setdefault("MAX_CONTENT_LENGTH",   16 * 1024 * 1024)
    app.config.setdefault("MAX_FORM_MEMORY_SIZE", 16 * 1024 * 1024)
    app.config.setdefault("MAX_FORM_PARTS",       10_000)

    # Init extensions
    db.init_app(app)
    login_manager.init_app(app)
    csrf.init_app(app)
    limiter.init_app(app)

    # ── User loader ──────────────────────────────────────────────────────────
    from app.models import User

    @login_manager.user_loader
    def load_user(user_id):
        return db.session.get(User, int(user_id))

    # ── Register Blueprints ──────────────────────────────────────────────────
    from app.routes.auth_routes import auth_bp
    from app.routes.admin_routes import admin_bp
    from app.routes.cis_routes import cis_bp

    app.register_blueprint(auth_bp)
    app.register_blueprint(admin_bp)
    app.register_blueprint(cis_bp)

    # ── Create DB ────────────────────────────────────────────────────────────
    with app.app_context():
        db.create_all()

    # ── Error Handlers ───────────────────────────────────────────────────────
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