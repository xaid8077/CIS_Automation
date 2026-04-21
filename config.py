# config.py
"""
config.py
─────────
Environment-based configuration.

python-dotenv loads .env automatically before any class attribute is
evaluated, so os.environ already contains the .env values by the time
Flask reads the config object.

Environment classes
───────────────────
Development  — local dev, SQLite fallback removed (use real Postgres).
Production   — strict validation at startup; refuses to boot with defaults.
Testing      — in-memory SQLite, CSRF off, short JWT lifetime.

New in this version
───────────────────
- python-dotenv auto-load at module import time.
- PostgreSQL as the only supported DB engine (no SQLite in dev/prod).
- JWT configuration (secret, token lifetimes, location).
- Redis URL for rate limiting (falls back to memory:// in dev).
- SESSION_COOKIE_SECURE read from env so it can be toggled without
  changing code.
"""

import os
from datetime import timedelta

from dotenv import load_dotenv

# Load .env before any os.environ access.
# override=False means real environment variables always win over .env,
# which is the correct behaviour for containerised / CI deployments.
load_dotenv(override=False)


# ─── Base ─────────────────────────────────────────────────────────────────────

class Base:

    # ── Core ─────────────────────────────────────────────────────────────────
    SECRET_KEY             = os.environ.get("SECRET_KEY", "change-me-before-production")
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    MAX_CONTENT_LENGTH     = 16 * 1024 * 1024   # 16 MB

    # ── Session cookies ───────────────────────────────────────────────────────
    SESSION_COOKIE_HTTPONLY  = True
    SESSION_COOKIE_SAMESITE  = "Lax"
    SESSION_COOKIE_SECURE    = os.environ.get("SESSION_COOKIE_SECURE", "true").lower() == "true"
    PERMANENT_SESSION_LIFETIME = timedelta(hours=8)

    # ── CSRF ──────────────────────────────────────────────────────────────────
    WTF_CSRF_TIME_LIMIT      = 3600

    # ── Rate limiting ─────────────────────────────────────────────────────────
    # In production, REDIS_URL should point to a real Redis instance.
    # "memory://" is single-process only and resets on every restart.
    RATELIMIT_DEFAULT        = "200 per day;50 per hour"
    RATELIMIT_STORAGE_URL    = os.environ.get("REDIS_URL", "memory://")
    RATELIMIT_STRATEGY       = "fixed-window"
    RATELIMIT_HEADERS_ENABLED = True

    # ── JWT ───────────────────────────────────────────────────────────────────
    JWT_SECRET_KEY           = os.environ.get("JWT_SECRET_KEY", "jwt-change-me-before-production")
    JWT_ACCESS_TOKEN_EXPIRES = timedelta(
        seconds=int(os.environ.get("JWT_ACCESS_TOKEN_EXPIRES", 3600))
    )
    JWT_REFRESH_TOKEN_EXPIRES = timedelta(
        seconds=int(os.environ.get("JWT_REFRESH_TOKEN_EXPIRES", 2592000))
    )
    # Tokens are read from the Authorization header as "Bearer <token>".
    # Cookie-based JWT is not used — the web UI uses session auth (Flask-Login).
    JWT_TOKEN_LOCATION       = ["headers"]
    JWT_HEADER_NAME          = "Authorization"
    JWT_HEADER_TYPE          = "Bearer"

    # ── SQLAlchemy connection pool ────────────────────────────────────────────
    # These defaults are sensible for PostgreSQL with a small team.
    # Increase pool_size / max_overflow for high-concurrency production.
    SQLALCHEMY_ENGINE_OPTIONS = {
        "pool_pre_ping": True,        # verify connection before checkout
        "pool_size":     5,
        "max_overflow":  10,
        "pool_timeout":  30,
        "pool_recycle":  1800,        # recycle connections every 30 min
    }


# ─── Development ──────────────────────────────────────────────────────────────

class Development(Base):
    DEBUG   = True
    TESTING = False

    SQLALCHEMY_DATABASE_URI = os.environ.get(
        "DATABASE_URL",
        # Explicit fallback so a missing .env gives a clear connection error
        # rather than silently falling back to SQLite.
        "postgresql+psycopg2://cis_user:cis_password@localhost:5432/cis_dev",
    )

    # HTTP is fine locally — no need for secure cookies.
    SESSION_COOKIE_SECURE = False

    # Shorter JWT lifetime is useful during development so token expiry
    # behaviour can be tested without waiting an hour.
    JWT_ACCESS_TOKEN_EXPIRES = timedelta(minutes=30)


# ─── Production ───────────────────────────────────────────────────────────────

class Production(Base):
    DEBUG   = False
    TESTING = False

    SQLALCHEMY_DATABASE_URI = os.environ.get("DATABASE_URL")

    # Force SSL-strict CSRF in production.
    WTF_CSRF_SSL_STRICT = True

    @classmethod
    def validate(cls) -> None:
        """
        Called at startup in production mode.
        Raises EnvironmentError if any critical env var is missing or
        still set to an insecure default.
        """
        missing = []

        if cls.SECRET_KEY == "change-me-before-production":
            missing.append("SECRET_KEY")

        if cls.JWT_SECRET_KEY == "jwt-change-me-before-production":
            missing.append("JWT_SECRET_KEY")

        if not cls.SQLALCHEMY_DATABASE_URI:
            missing.append("DATABASE_URL")

        # Warn explicitly when Redis is not configured — memory:// is not
        # suitable for production (resets on restart, not shared across workers).
        if cls.RATELIMIT_STORAGE_URL == "memory://":
            missing.append("REDIS_URL (memory:// is not safe for production)")

        if missing:
            raise EnvironmentError(
                "Production mode requires these environment variables to be "
                f"set correctly: {missing}"
            )


# ─── Testing ──────────────────────────────────────────────────────────────────

class Testing(Base):
    TESTING  = True
    DEBUG    = False

    # In-memory SQLite is fast for unit tests and requires no external service.
    SQLALCHEMY_DATABASE_URI  = "sqlite:///:memory:"

    # Disable CSRF so tests can POST without tokens.
    WTF_CSRF_ENABLED         = False
    SESSION_COOKIE_SECURE    = False

    # Short JWT lifetime to keep test runs fast.
    JWT_ACCESS_TOKEN_EXPIRES  = timedelta(minutes=5)
    JWT_REFRESH_TOKEN_EXPIRES = timedelta(minutes=10)

    # Disable connection-pool settings for SQLite.
    SQLALCHEMY_ENGINE_OPTIONS = {}


# ─── Registry ─────────────────────────────────────────────────────────────────

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