"""
models.py
─────────
SQLAlchemy User model.

Password hashing: Argon2id via argon2-cffi.
  - Winner of the Password Hashing Competition (PHC).
  - Memory-hard: resists GPU/ASIC brute-force far better than bcrypt.
  - pip install argon2-cffi

Roles: "user" (read/generate) | "admin" (user management dashboard).
"""

from datetime import datetime, timezone
from argon2 import PasswordHasher
from argon2.exceptions import VerifyMismatchError, VerificationError, InvalidHashError
from flask_login import UserMixin
from extensions import db

_ph = PasswordHasher(
    time_cost=3,        # iterations
    memory_cost=65536,  # 64 MB RAM per hash — brute-force is expensive
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

    # ── Password handling ─────────────────────────────────────────────────────

    def set_password(self, plaintext: str) -> None:
        """Hash plaintext with Argon2id and store the result."""
        if len(plaintext) < 8:
            raise ValueError("Password must be at least 8 characters.")
        self.password_hash = _ph.hash(plaintext)

    def check_password(self, plaintext: str) -> bool:
        """
        Verify plaintext against the stored hash.
        Also transparently rehashes if the stored parameters are outdated.
        Returns True on success, False on mismatch.
        """
        try:
            _ph.verify(self.password_hash, plaintext)
        except VerifyMismatchError:
            return False
        except (VerificationError, InvalidHashError):
            return False

        # Argon2 recommends rehashing when parameters change
        if _ph.check_needs_rehash(self.password_hash):
            self.password_hash = _ph.hash(plaintext)
            db.session.commit()

        return True

    # ── Role helpers ──────────────────────────────────────────────────────────

    @property
    def is_admin(self) -> bool:
        return self.role == "admin"

    def __repr__(self):
        return f"<User {self.username!r} role={self.role!r}>"
