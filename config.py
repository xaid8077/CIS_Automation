"""
config.py
─────────
Environment-based configuration.
Load the correct class by setting FLASK_ENV=development|production|testing.

All secrets MUST come from environment variables in production.
Never commit a real SECRET_KEY or DATABASE_URL to version control.
"""

import os
from datetime import timedelta


class Base:
    # ── Core ──────────────────────────────────────────────────────────────────
    SECRET_KEY = os.environ.get("SECRET_KEY", "change-me-before-production")
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    MAX_CONTENT_LENGTH = 16 * 1024 * 1024   # 16 MB upload limit

    # ── Session cookies ───────────────────────────────────────────────────────
    # HttpOnly: JS cannot read the cookie (blocks XSS theft).
    # Samesite: blocks CSRF from cross-site requests.
    # Secure defaults to True here so any env class that forgets to override
    # it stays safe.  Development and Testing explicitly set it to False.
    SESSION_COOKIE_HTTPONLY  = True
    SESSION_COOKIE_SAMESITE  = "Lax"
    SESSION_COOKIE_SECURE    = True   # overridden to False in dev/testing below
    PERMANENT_SESSION_LIFETIME = timedelta(hours=8)

    # ── CSRF ──────────────────────────────────────────────────────────────────
    WTF_CSRF_TIME_LIMIT = 3600   # token expires after 1 hour

    # ── Rate limiting ─────────────────────────────────────────────────────────
    RATELIMIT_DEFAULT          = "200 per day;50 per hour"
    RATELIMIT_STORAGE_URL      = "memory://"   # swap for redis:// in prod
    RATELIMIT_STRATEGY         = "fixed-window"
    RATELIMIT_HEADERS_ENABLED  = True


class Development(Base):
    DEBUG  = True
    TESTING = False
    SQLALCHEMY_DATABASE_URI = os.environ.get(
        "DATABASE_URL", "sqlite:///cis_dev.db"
    )
    SESSION_COOKIE_SECURE = False   # HTTP is fine locally


class Production(Base):
    DEBUG  = False
    TESTING = False
    SQLALCHEMY_DATABASE_URI = os.environ.get("DATABASE_URL")   # must be set
    # SESSION_COOKIE_SECURE inherited as True from Base — no override needed.
    WTF_CSRF_SSL_STRICT     = True

    @classmethod
    def validate(cls):
        """Call at startup to catch missing env vars early."""
        missing = []
        if cls.SECRET_KEY == "change-me-before-production":
            missing.append("SECRET_KEY")
        if not cls.SQLALCHEMY_DATABASE_URI:
            missing.append("DATABASE_URL")
        if missing:
            raise EnvironmentError(
                f"Production requires these env vars to be set: {missing}"
            )


class Testing(Base):
    TESTING = True
    SQLALCHEMY_DATABASE_URI = "sqlite:///:memory:"
    WTF_CSRF_ENABLED        = False
    SESSION_COOKIE_SECURE   = False


_map = {
    "development": Development,
    "production":  Production,
    "testing":     Testing,
}

def get_config():
    env = os.environ.get("FLASK_ENV", "development").lower()
    cfg = _map.get(env, Development)
    if env == "production":
        cfg.validate()
    return cfg
