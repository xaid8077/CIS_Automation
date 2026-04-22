"""
config.py
─────────
Environment-based configuration.
Load the correct class by setting FLASK_ENV=development|production|testing.

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
    SECRET_KEY = os.environ.get("SECRET_KEY")
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    MAX_CONTENT_LENGTH = 16 * 1024 * 1024   # 16 MB upload limit

    # ── SQLAlchemy connection pool ─────────────────────────────────────────────
    SQLALCHEMY_ENGINE_OPTIONS = {
        "pool_pre_ping": True,    # discard stale connections automatically
        "pool_recycle":  1800,    # recycle connections every 30 min
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
    RATELIMIT_DEFAULT          = "200 per day;50 per hour"
    RATELIMIT_STORAGE_URL      = "memory://"
    RATELIMIT_STRATEGY         = "fixed-window"
    RATELIMIT_HEADERS_ENABLED  = True


class Development(Base):
    DEBUG   = True
    TESTING = False
    SQLALCHEMY_DATABASE_URI = os.environ.get(
        "DATABASE_URL", "postgresql://postgres:postgres@localhost:6543/cis_automation?sslmode=disable"
    )
    SESSION_COOKIE_SECURE = False

    @classmethod
    def validate(cls):
        if not cls.SQLALCHEMY_DATABASE_URI:
            raise EnvironmentError(
                "DATABASE_URL is not set. "
                "Add it to your .env file:\n"
                "  DATABASE_URL=postgresql://postgres:Minecraft%40007@localhost:6543/cis_automation"
            )


class Production(Base):
    DEBUG   = False
    TESTING = False
    SQLALCHEMY_DATABASE_URI = os.environ.get("DATABASE_URL")
    WTF_CSRF_SSL_STRICT     = True

    @classmethod
    def validate(cls):
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
    cfg.validate()
    return cfg