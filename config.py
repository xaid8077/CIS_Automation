"""
config.py
─────────
Environment-based configuration.
Load the correct class by setting FLASK_ENV=development|production|testing.

All secrets MUST come from environment variables — never commit a real
SECRET_KEY or DATABASE_URL to version control.

Neon.tech PostgreSQL is used for both development and production.
SQLite is used only for testing (in-memory, no credentials needed).
"""

import os
from datetime import timedelta

# ── Load .env automatically if python-dotenv is installed ─────────────────────
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass   # dotenv not installed — rely on env vars being set externally


class Base:
    # ── Core ──────────────────────────────────────────────────────────────────
    SECRET_KEY = os.environ.get("SECRET_KEY", "change-me-before-production")
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    MAX_CONTENT_LENGTH = 16 * 1024 * 1024   # 16 MB upload limit

    # ── SQLAlchemy connection pool — tuned for serverless Postgres ────────────
    # Neon suspends the compute after ~5 min of inactivity; connections can go
    # stale.  pool_pre_ping sends a lightweight "SELECT 1" before reuse so
    # SQLAlchemy automatically discards dead connections instead of erroring.
    SQLALCHEMY_ENGINE_OPTIONS = {
        "pool_pre_ping": True,          # discard stale connections automatically
        "pool_recycle":  300,           # recycle connections every 5 min
        "pool_size":     5,             # keep up to 5 connections open
        "max_overflow":  10,            # allow up to 10 extra under load
        "connect_args": {
            "sslmode": "require",       # Neon requires TLS — enforce it here too
            "connect_timeout": 10,      # fail fast if Neon hasn't woken up yet
        },
    }

    # ── Session cookies ───────────────────────────────────────────────────────
    SESSION_COOKIE_HTTPONLY  = True
    SESSION_COOKIE_SAMESITE  = "Lax"
    SESSION_COOKIE_SECURE    = True   # overridden to False in dev/testing
    PERMANENT_SESSION_LIFETIME = timedelta(hours=8)

    # ── CSRF ──────────────────────────────────────────────────────────────────
    WTF_CSRF_TIME_LIMIT = 3600   # token expires after 1 hour

    # ── Rate limiting ─────────────────────────────────────────────────────────
    RATELIMIT_DEFAULT          = "200 per day;50 per hour"
    RATELIMIT_STORAGE_URL      = "memory://"   # swap for redis:// in prod
    RATELIMIT_STRATEGY         = "fixed-window"
    RATELIMIT_HEADERS_ENABLED  = True


class Development(Base):
    DEBUG   = True
    TESTING = False

    # Uses the same Neon DB as production — change to a separate Neon
    # branch/database if you want true environment isolation.
    SQLALCHEMY_DATABASE_URI = os.environ.get("DATABASE_URL")

    # HTTP is fine locally (no HTTPS) so the session cookie does not need
    # the Secure flag — browsers won't send it over plain HTTP otherwise.
    SESSION_COOKIE_SECURE = False

    @classmethod
    def validate(cls):
        if not cls.SQLALCHEMY_DATABASE_URI:
            raise EnvironmentError(
                "DATABASE_URL is not set. "
                "Add it to your .env file:\n"
                "  DATABASE_URL=postgresql://..."
            )


class Production(Base):
    DEBUG   = False
    TESTING = False
    SQLALCHEMY_DATABASE_URI = os.environ.get("DATABASE_URL")
    WTF_CSRF_SSL_STRICT     = True
    # SESSION_COOKIE_SECURE inherited as True from Base.

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
    # In-memory SQLite for unit tests — no Neon credentials needed.
    SQLALCHEMY_DATABASE_URI = "sqlite:///:memory:"
    WTF_CSRF_ENABLED        = False
    SESSION_COOKIE_SECURE   = False
    # Override pool options — SQLite doesn't support the same connect_args.
    SQLALCHEMY_ENGINE_OPTIONS = {
        "pool_pre_ping": True,
    }


_map = {
    "development": Development,
    "production":  Production,
    "testing":     Testing,
}


def get_config():
    env = os.environ.get("FLASK_ENV", "development").lower()
    cfg = _map.get(env, Development)
    # Always validate — catches missing DATABASE_URL early in all envs.
    cfg.validate()
    return cfg