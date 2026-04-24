# config.py
"""
config.py
─────────
Environment-based configuration.

All secrets MUST come from environment variables — never commit a real
SECRET_KEY or DATABASE_URL to version control.
"""

import os
from datetime import timedelta

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass


class Base:
    # ── Core ──────────────────────────────────────────────────────────────────
    SECRET_KEY = os.environ.get("SECRET_KEY", "change-me-before-production")
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    MAX_CONTENT_LENGTH = 16 * 1024 * 1024   # 16 MB upload limit

    # ── SQLAlchemy connection pool ─────────────────────────────────────────────
    SQLALCHEMY_ENGINE_OPTIONS = {
        "pool_pre_ping": True,
        "pool_recycle":  1800,
        "pool_size":     5,
        "max_overflow":  10,
    }

    # ── Session cookies ───────────────────────────────────────────────────────
    SESSION_COOKIE_HTTPONLY  = True
    SESSION_COOKIE_SAMESITE  = "Lax"
    SESSION_COOKIE_SECURE    = True
    PERMANENT_SESSION_LIFETIME = timedelta(hours=8)

    # ── CSRF ──────────────────────────────────────────────────────────────────
    WTF_CSRF_TIME_LIMIT = 3600

    # ── Rate limiting ─────────────────────────────────────────────────────────
    RATELIMIT_DEFAULT         = "200 per day;50 per hour"
    RATELIMIT_STORAGE_URL     = "memory://"
    RATELIMIT_STRATEGY        = "fixed-window"
    RATELIMIT_HEADERS_ENABLED = True

    # ── JWT (Flask-JWT-Extended) ───────────────────────────────────────────────
    # Access tokens are short-lived; refresh tokens are long-lived.
    # Both durations are overridable via environment variables so production
    # deployments can tune them without a code change.
    JWT_SECRET_KEY             = os.environ.get("JWT_SECRET_KEY") or os.environ.get("SECRET_KEY", "change-me-jwt")
    JWT_ACCESS_TOKEN_EXPIRES   = timedelta(
        seconds=int(os.environ.get("JWT_ACCESS_EXPIRES_SECONDS",  "3600"))
    )
    JWT_REFRESH_TOKEN_EXPIRES  = timedelta(
        days=int(os.environ.get("JWT_REFRESH_EXPIRES_DAYS", "30"))
    )
    JWT_TOKEN_LOCATION         = ["headers"]
    JWT_HEADER_NAME            = "Authorization"
    JWT_HEADER_TYPE            = "Bearer"


class Development(Base):
    DEBUG   = True
    TESTING = False
    SQLALCHEMY_DATABASE_URI = os.environ.get(
        "DATABASE_URL",
        "postgresql://postgres:postgres@localhost:6543/cis_automation?sslmode=disable"
    )
    SESSION_COOKIE_SECURE = False

    @classmethod
    def validate(cls):
        if not cls.SQLALCHEMY_DATABASE_URI:
            raise EnvironmentError(
                "DATABASE_URL is not set. "
                "Add it to your .env file:\n"
                "  DATABASE_URL=postgresql://user:password@localhost:5432/cis_automation"
            )


class Production(Base):
    DEBUG   = False
    TESTING = False
    SQLALCHEMY_DATABASE_URI = os.environ.get("DATABASE_URL")
    WTF_CSRF_SSL_STRICT     = True

    @classmethod
    def validate(cls):
        missing = []
        if not cls.SECRET_KEY or cls.SECRET_KEY == "change-me-before-production":
            missing.append("SECRET_KEY")
        if not cls.SQLALCHEMY_DATABASE_URI:
            missing.append("DATABASE_URL")
        # Warn (not hard-fail) when JWT key falls back to SECRET_KEY
        if not os.environ.get("JWT_SECRET_KEY"):
            import warnings
            warnings.warn(
                "JWT_SECRET_KEY not set — falling back to SECRET_KEY. "
                "Set a dedicated JWT_SECRET_KEY in production.",
                stacklevel=2,
            )
        if missing:
            raise EnvironmentError(
                f"Production requires these env vars to be set: {missing}"
            )


class Testing(Base):
    TESTING = True
    SQLALCHEMY_DATABASE_URI = "sqlite:///:memory:"
    WTF_CSRF_ENABLED        = False
    SESSION_COOKIE_SECURE   = False
    # Disable token expiry in tests so fixtures never need to refresh
    JWT_ACCESS_TOKEN_EXPIRES  = False   # type: ignore[assignment]
    JWT_REFRESH_TOKEN_EXPIRES = False   # type: ignore[assignment]
    SQLALCHEMY_ENGINE_OPTIONS = {
        "pool_pre_ping": True,
    }

    @classmethod
    def validate(cls):
        pass   # Nothing to validate for the in-memory test DB


_map = {
    "development": Development,
    "production":  Production,
    "testing":     Testing,
}


def get_config():
    env = os.environ.get("FLASK_ENV", "development").lower()
    cfg = _map.get(env, Development)
    cfg.validate()
    return cfg