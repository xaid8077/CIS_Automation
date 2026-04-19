"""
models.py
─────────
SQLAlchemy Models for Users, Projects, ProjectLocations, and Document Revisions.

Schema changes vs previous version:
  - Project.nickname  : short display name, admin-only editable
  - ProjectLocation   : new model — one project → many locations
  - DocumentRevision.location_id : FK to ProjectLocation (nullable so old
    revisions without a location are not broken)
"""

from datetime import datetime, timezone
from argon2 import PasswordHasher
from argon2.exceptions import VerifyMismatchError, VerificationError, InvalidHashError
from flask_login import UserMixin
from extensions import db

_ph = PasswordHasher(
    time_cost=3,
    memory_cost=65536,
    parallelism=2,
    hash_len=32,
    salt_len=16,
)

ROLES = ("user", "admin")


# ─────────────────────────────────────────────────────────────────────────────
# User
# ─────────────────────────────────────────────────────────────────────────────

class User(UserMixin, db.Model):
    __tablename__ = "users"

    id            = db.Column(db.Integer,     primary_key=True)
    username      = db.Column(db.String(64),  unique=True, nullable=False, index=True)
    email         = db.Column(db.String(120), unique=True, nullable=False, index=True)
    password_hash = db.Column(db.String(256), nullable=False)
    role          = db.Column(db.String(16),  nullable=False, default="user")
    is_active     = db.Column(db.Boolean,     nullable=False, default=True)
    created_at    = db.Column(db.DateTime,    default=lambda: datetime.now(timezone.utc))
    last_login    = db.Column(db.DateTime,    nullable=True)

    # Relationships
    revisions = db.relationship("DocumentRevision", backref="author", lazy=True)

    def set_password(self, plaintext: str) -> None:
        if len(plaintext) < 8:
            raise ValueError("Password must be at least 8 characters.")
        self.password_hash = _ph.hash(plaintext)

    def check_password(self, plaintext: str) -> bool:
        try:
            _ph.verify(self.password_hash, plaintext)
        except VerifyMismatchError:
            return False
        except (VerificationError, InvalidHashError):
            return False

        if _ph.check_needs_rehash(self.password_hash):
            try:
                self.password_hash = _ph.hash(plaintext)
                db.session.commit()
            except Exception:
                db.session.rollback()

        return True

    @property
    def is_admin(self) -> bool:
        return self.role == "admin"

    def __repr__(self):
        return f"<User {self.username!r} role={self.role!r}>"


# ─────────────────────────────────────────────────────────────────────────────
# Project
# ─────────────────────────────────────────────────────────────────────────────

class Project(db.Model):
    __tablename__ = "projects"

    id         = db.Column(db.Integer,     primary_key=True)
    name       = db.Column(db.String(150), unique=True, nullable=False)
    nickname   = db.Column(db.String(40),  nullable=True)   # short display name
    client     = db.Column(db.String(150), nullable=False)
    consultant = db.Column(db.String(150), nullable=True)
    location   = db.Column(db.String(150), nullable=True)   # legacy field — kept for old rows
    created_at = db.Column(db.DateTime,    default=lambda: datetime.now(timezone.utc))

    # Relationships
    locations = db.relationship(
        "ProjectLocation",
        backref="project",
        lazy=True,
        cascade="all, delete-orphan",
        order_by="ProjectLocation.name",
    )
    revisions = db.relationship(
        "DocumentRevision",
        backref="project",
        lazy=True,
        cascade="all, delete-orphan",
    )

    @property
    def display_name(self) -> str:
        """Return nickname if set, otherwise the full project name."""
        return self.nickname.strip() if self.nickname and self.nickname.strip() else self.name

    def __repr__(self):
        return f"<Project {self.name!r}>"


# ─────────────────────────────────────────────────────────────────────────────
# ProjectLocation
# ─────────────────────────────────────────────────────────────────────────────

class ProjectLocation(db.Model):
    """
    One project can have many locations (e.g. different plant sites).
    Each DocumentRevision is tied to exactly one location.
    """
    __tablename__ = "project_locations"

    id         = db.Column(db.Integer,     primary_key=True)
    project_id = db.Column(db.Integer,     db.ForeignKey("projects.id"), nullable=False)
    name       = db.Column(db.String(150), nullable=False)   # e.g. "Plant A – Block 3"
    code       = db.Column(db.String(20),  nullable=True)    # short code, e.g. "PA-B3"
    created_at = db.Column(db.DateTime,    default=lambda: datetime.now(timezone.utc))

    # Relationships
    revisions = db.relationship(
        "DocumentRevision",
        backref="location",
        lazy=True,
    )

    __table_args__ = (
        db.UniqueConstraint("project_id", "name", name="uq_project_location_name"),
    )

    @property
    def display(self) -> str:
        """Return 'CODE — Name' if code exists, otherwise just the name."""
        if self.code and self.code.strip():
            return f"{self.code.strip()} — {self.name}"
        return self.name

    def __repr__(self):
        return f"<ProjectLocation {self.name!r} project_id={self.project_id}>"


# ─────────────────────────────────────────────────────────────────────────────
# DocumentRevision
# ─────────────────────────────────────────────────────────────────────────────

class DocumentRevision(db.Model):
    __tablename__ = "document_revisions"

    id              = db.Column(db.Integer, primary_key=True)
    project_id      = db.Column(db.Integer, db.ForeignKey("projects.id"), nullable=False)
    user_id         = db.Column(db.Integer, db.ForeignKey("users.id"),    nullable=False)

    # nullable so existing revisions (before locations were added) are preserved
    location_id     = db.Column(
        db.Integer,
        db.ForeignKey("project_locations.id"),
        nullable=True,
    )

    doc_type        = db.Column(db.String(50),  nullable=False)
    revision_number = db.Column(db.Integer,     nullable=False, default=0)
    data_payload    = db.Column(db.JSON,        nullable=False)
    created_at      = db.Column(db.DateTime,    default=lambda: datetime.now(timezone.utc))

    def __repr__(self):
        return (
            f"<DocRev {self.doc_type!r} "
            f"Rev:{self.revision_number} "
            f"Proj:{self.project_id} "
            f"Loc:{self.location_id}>"
        )