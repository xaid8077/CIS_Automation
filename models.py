"""
models.py
─────────
SQLAlchemy Models for Users, Projects, and Document Revisions.
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

class User(UserMixin, db.Model):
    __tablename__ = "users"

    id           = db.Column(db.Integer,     primary_key=True)
    username     = db.Column(db.String(64),  unique=True, nullable=False, index=True)
    email        = db.Column(db.String(120), unique=True, nullable=False, index=True)
    password_hash= db.Column(db.String(256), nullable=False)
    role         = db.Column(db.String(16),  nullable=False, default="user")
    is_active    = db.Column(db.Boolean,     nullable=False, default=True)
    created_at   = db.Column(db.DateTime,    default=lambda: datetime.now(timezone.utc))
    last_login   = db.Column(db.DateTime,    nullable=True)
    
    # Relationships
    revisions    = db.relationship('DocumentRevision', backref='author', lazy=True)

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


class Project(db.Model):
    __tablename__ = "projects"

    id          = db.Column(db.Integer, primary_key=True)
    name        = db.Column(db.String(150), unique=True, nullable=False)
    client      = db.Column(db.String(150), nullable=False)
    consultant  = db.Column(db.String(150), nullable=True)
    location    = db.Column(db.String(150), nullable=True)
    created_at  = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    
    # Relationships
    revisions   = db.relationship('DocumentRevision', backref='project', lazy=True, cascade="all, delete-orphan")

    def __repr__(self):
        return f"<Project {self.name!r}>"


class DocumentRevision(db.Model):
    __tablename__ = "document_revisions"

    id              = db.Column(db.Integer, primary_key=True)
    project_id      = db.Column(db.Integer, db.ForeignKey('projects.id'), nullable=False)
    user_id         = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    
    doc_type        = db.Column(db.String(50), nullable=False) # e.g., "Instrument List", "IO List"
    revision_number = db.Column(db.Integer, nullable=False, default=0)
    
    # Store the entire form payload as JSON so it can be re-loaded or analyzed later
    data_payload    = db.Column(db.JSON, nullable=False)
    
    created_at      = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    def __repr__(self):
        return f"<DocRev {self.doc_type} Rev:{self.revision_number} Proj:{self.project_id}>"