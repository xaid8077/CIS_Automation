# models.py
"""
models.py
─────────
SQLAlchemy models.

Role system (new in this version)
──────────────────────────────────
Three roles replace the previous binary admin / user split:

  admin     — full access: create projects, manage users, patch doc numbers.
  engineer  — create and download revisions; cannot manage users or projects.
  viewer    — read-only: browse projects and download existing revisions.

The permission matrix is enforced by the RBAC decorators in utils/rbac.py.
Models expose simple boolean properties (can_generate, can_manage_projects,
etc.) so that templates and services can make decisions without importing
the RBAC module.

Permission matrix
─────────────────
Action                        admin   engineer   viewer
─────────────────────────────────────────────────────────
view projects / history         ✓        ✓         ✓
download existing revision      ✓        ✓         ✓
save draft / generate doc       ✓        ✓         ✗
create / edit / delete project  ✓        ✗         ✗
manage users                    ✓        ✗         ✗
manage locations                ✓        ✗         ✗
patch revision doc numbers      ✓        ✗         ✗
"""

from datetime import datetime, timezone

from argon2 import PasswordHasher
from argon2.exceptions import InvalidHashError, VerificationError, VerifyMismatchError
from flask_login import UserMixin

from extensions import db

_ph = PasswordHasher(
    time_cost   = 3,
    memory_cost = 65536,
    parallelism = 2,
    hash_len    = 32,
    salt_len    = 16,
)

# All valid role strings — used by forms, validators, and the RBAC layer.
ROLES = ("admin", "engineer", "viewer")


# ─────────────────────────────────────────────────────────────────────────────
# Permission sets
# ─────────────────────────────────────────────────────────────────────────────

# Keep permission logic in one place so adding a new permission means
# editing exactly one dict, not hunting through decorators and templates.
_PERMISSIONS: dict[str, set[str]] = {
    # View project list and revision history.
    "view_projects":       {"admin", "engineer", "viewer"},

    # Download a previously generated revision file.
    "download_revision":   {"admin", "engineer", "viewer"},

    # Save a draft or generate a new revision (IL, IO List, etc.).
    "generate_document":   {"admin", "engineer"},

    # Create, edit, delete projects and assign locations.
    "manage_projects":     {"admin"},

    # Create, edit, deactivate users.
    "manage_users":        {"admin"},

    # Add / delete project locations.
    "manage_locations":    {"admin"},

    # Patch document numbers on an existing revision payload.
    "patch_doc_numbers":   {"admin"},
}


def has_permission(role: str, permission: str) -> bool:
    """
    Return True if *role* is granted *permission*.

    Parameters
    ----------
    role:       One of the strings in ROLES.
    permission: One of the keys in _PERMISSIONS.
    """
    return role in _PERMISSIONS.get(permission, set())


# ─────────────────────────────────────────────────────────────────────────────
# User
# ─────────────────────────────────────────────────────────────────────────────

class User(UserMixin, db.Model):
    __tablename__ = "users"

    id            = db.Column(db.Integer,     primary_key=True)
    username      = db.Column(db.String(64),  unique=True, nullable=False, index=True)
    email         = db.Column(db.String(120), unique=True, nullable=False, index=True)
    password_hash = db.Column(db.String(256), nullable=False)
    role          = db.Column(db.String(16),  nullable=False, default="viewer")
    is_active     = db.Column(db.Boolean,     nullable=False, default=True)
    created_at    = db.Column(
        db.DateTime(timezone=True),
        default=lambda: datetime.now(timezone.utc),
    )
    last_login    = db.Column(db.DateTime(timezone=True), nullable=True)

    # Relationships
    revisions = db.relationship(
        "DocumentRevision",
        backref = "author",
        lazy    = True,
    )

    # ── Password management ───────────────────────────────────────────────────

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

        # Opportunistic rehash when argon2 parameters are upgraded.
        if _ph.check_needs_rehash(self.password_hash):
            try:
                self.password_hash = _ph.hash(plaintext)
                db.session.commit()
            except Exception:
                db.session.rollback()

        return True

    # ── Permission shortcuts ──────────────────────────────────────────────────
    # Templates and services use these properties so they never import
    # the RBAC module or call has_permission() directly.

    @property
    def is_admin(self) -> bool:
        return self.role == "admin"

    @property
    def is_engineer(self) -> bool:
        return self.role in ("admin", "engineer")

    @property
    def can_view_projects(self) -> bool:
        return has_permission(self.role, "view_projects")

    @property
    def can_generate_document(self) -> bool:
        return has_permission(self.role, "generate_document")

    @property
    def can_download_revision(self) -> bool:
        return has_permission(self.role, "download_revision")

    @property
    def can_manage_projects(self) -> bool:
        return has_permission(self.role, "manage_projects")

    @property
    def can_manage_users(self) -> bool:
        return has_permission(self.role, "manage_users")

    @property
    def can_manage_locations(self) -> bool:
        return has_permission(self.role, "manage_locations")

    @property
    def can_patch_doc_numbers(self) -> bool:
        return has_permission(self.role, "patch_doc_numbers")

    # ── JWT identity ──────────────────────────────────────────────────────────

    def to_jwt_identity(self) -> dict:
        """
        Return the dict stored inside a JWT's 'sub' claim.
        Kept minimal — only what the API needs to authorise requests.
        """
        return {
            "id":       self.id,
            "username": self.username,
            "role":     self.role,
        }

    def to_dict(self) -> dict:
        """Public representation for API responses."""
        return {
            "id":         self.id,
            "username":   self.username,
            "email":      self.email,
            "role":       self.role,
            "is_active":  self.is_active,
            "created_at": self.created_at.isoformat() if self.created_at else None,
            "last_login": self.last_login.isoformat() if self.last_login else None,
        }

    def __repr__(self):
        return f"<User {self.username!r} role={self.role!r}>"


# ─────────────────────────────────────────────────────────────────────────────
# Project
# ─────────────────────────────────────────────────────────────────────────────

class Project(db.Model):
    __tablename__ = "projects"

    id         = db.Column(db.Integer,     primary_key=True)
    name       = db.Column(db.String(150), unique=True, nullable=False)
    nickname   = db.Column(db.String(40),  nullable=True)
    client     = db.Column(db.String(150), nullable=False)
    consultant = db.Column(db.String(150), nullable=True)
    created_at = db.Column(
        db.DateTime(timezone=True),
        default=lambda: datetime.now(timezone.utc),
    )

    # Relationships
    locations = db.relationship(
        "ProjectLocation",
        backref       = "project",
        lazy          = True,
        cascade       = "all, delete-orphan",
        order_by      = "ProjectLocation.name",
    )
    revisions = db.relationship(
        "DocumentRevision",
        backref  = "project",
        lazy     = True,
        cascade  = "all, delete-orphan",
    )

    @property
    def display_name(self) -> str:
        return self.nickname.strip() if self.nickname and self.nickname.strip() else self.name

    def to_dict(self) -> dict:
        return {
            "id":           self.id,
            "name":         self.name,
            "nickname":     self.nickname,
            "display_name": self.display_name,
            "client":       self.client,
            "consultant":   self.consultant,
            "created_at":   self.created_at.isoformat() if self.created_at else None,
            "locations":    [loc.to_dict() for loc in self.locations],
        }

    def __repr__(self):
        return f"<Project {self.name!r}>"


# ─────────────────────────────────────────────────────────────────────────────
# ProjectLocation
# ─────────────────────────────────────────────────────────────────────────────

class ProjectLocation(db.Model):
    __tablename__ = "project_locations"

    id         = db.Column(db.Integer,     primary_key=True)
    project_id = db.Column(db.Integer,     db.ForeignKey("projects.id"),  nullable=False)
    name       = db.Column(db.String(150), nullable=False)
    code       = db.Column(db.String(20),  nullable=True)
    created_at = db.Column(
        db.DateTime(timezone=True),
        default=lambda: datetime.now(timezone.utc),
    )

    revisions = db.relationship(
        "DocumentRevision",
        backref = "location",
        lazy    = True,
    )

    __table_args__ = (
        db.UniqueConstraint("project_id", "name", name="uq_project_location_name"),
    )

    @property
    def display(self) -> str:
        if self.code and self.code.strip():
            return f"{self.code.strip()} — {self.name}"
        return self.name

    def to_dict(self) -> dict:
        return {
            "id":         self.id,
            "project_id": self.project_id,
            "name":       self.name,
            "code":       self.code,
            "display":    self.display,
            "created_at": self.created_at.isoformat() if self.created_at else None,
        }

    def __repr__(self):
        return f"<ProjectLocation {self.name!r} project_id={self.project_id}>"


# ─────────────────────────────────────────────────────────────────────────────
# DocumentRevision
# ─────────────────────────────────────────────────────────────────────────────

class DocumentRevision(db.Model):
    __tablename__ = "document_revisions"

    id              = db.Column(db.Integer, primary_key=True)
    project_id      = db.Column(db.Integer, db.ForeignKey("projects.id"),          nullable=False)
    user_id         = db.Column(db.Integer, db.ForeignKey("users.id"),             nullable=False)
    location_id     = db.Column(db.Integer, db.ForeignKey("project_locations.id"), nullable=True)

    doc_type        = db.Column(db.String(50),  nullable=False)
    revision_number = db.Column(db.Integer,     nullable=False, default=0)
    data_payload    = db.Column(db.JSON,        nullable=False)

    status          = db.Column(db.String(16),  nullable=False, default="published")
    # 'draft'     — saved via "Save & Submit", no file generated.
    #               Only one draft kept per project+location (overwritten on each save).
    # 'published' — triggered by a download action, accumulates as numbered revisions.

    created_at      = db.Column(
        db.DateTime(timezone=True),
        default=lambda: datetime.now(timezone.utc),
    )

    def to_dict(self, include_payload: bool = False) -> dict:
        """
        Serialise to a dict safe for API responses.

        Parameters
        ----------
        include_payload:
            Set True only when the caller explicitly needs the full
            data_payload (e.g. the { } Data modal).  Defaults to False
            to keep list responses lean.
        """
        d = {
            "id":              self.id,
            "project_id":      self.project_id,
            "location_id":     self.location_id,
            "doc_type":        self.doc_type,
            "revision_number": self.revision_number,
            "status":          self.status,
            "created_at":      self.created_at.isoformat() if self.created_at else None,
            "author":          self.author.username if self.author else None,
            "location":        self.location.to_dict() if self.location else None,
        }
        if include_payload:
            d["data_payload"] = self.data_payload
        return d

    def __repr__(self):
        return (
            f"<DocRev {self.doc_type!r} "
            f"Rev:{self.revision_number} "
            f"Status:{self.status} "
            f"Proj:{self.project_id}>"
        )