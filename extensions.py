# extensions.py
"""
extensions.py
─────────────
All Flask extension singletons live here.

Import from this module everywhere — never from app.py — to avoid
circular imports.
"""

from flask_sqlalchemy   import SQLAlchemy
from flask_login        import LoginManager
from flask_wtf.csrf     import CSRFProtect
from flask_limiter      import Limiter
from flask_limiter.util import get_remote_address
from flask_migrate      import Migrate
from flask_jwt_extended import JWTManager

db            = SQLAlchemy()
login_manager = LoginManager()
csrf          = CSRFProtect()
limiter       = Limiter(key_func=get_remote_address)
migrate       = Migrate()
jwt           = JWTManager()

# ── Flask-Login defaults ───────────────────────────────────────────────────────
login_manager.login_view             = "auth.login"
login_manager.login_message          = "Please log in to access this page."
login_manager.login_message_category = "warning"