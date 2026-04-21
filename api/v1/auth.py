# api/v1/auth.py
"""
api/v1/auth.py
──────────────
JWT authentication endpoints.

POST /api/v1/auth/login
    Exchange username + password for an access token and a refresh token.

POST /api/v1/auth/refresh
    Exchange a valid refresh token for a new access token.
    Send the refresh token in the Authorization header as "Bearer <token>".

GET  /api/v1/auth/me
    Return the authenticated user's profile.

POST /api/v1/auth/logout
    Client-side logout hint — instructs the client to discard its tokens.
    Server-side blacklisting requires a Redis token store; that is noted
    as a TODO for when the deployment warrants it.

Design notes
────────────
- Access tokens are short-lived (default 1 hour, configurable via env).
- Refresh tokens are long-lived (default 30 days, configurable via env).
- The same constant-time dummy verification used in routes/auth.py
  is applied here to prevent username enumeration.
- Rate limiting is applied to /login to mitigate brute-force attacks.
"""

from datetime import datetime, timezone

from flask import Blueprint, request, jsonify
from flask_jwt_extended import (
    create_access_token,
    create_refresh_token,
    jwt_required,
    get_jwt_identity,
)
from marshmallow import Schema, fields, ValidationError, validates, EXCLUDE

from extensions import db, limiter
from models import User
from utils.rbac import api_permission_required

auth_api_bp = Blueprint("auth_api", __name__, url_prefix="/auth")


# ─── Request schemas ──────────────────────────────────────────────────────────

class LoginSchema(Schema):
    class Meta:
        unknown = EXCLUDE

    username = fields.Str(required=True)
    password = fields.Str(required=True, load_default="")

    @validates("username")
    def not_blank(self, value):
        if not value.strip():
            from marshmallow import ValidationError
            raise ValidationError("Username is required.")


_login_schema = LoginSchema()


# ─── Helpers ──────────────────────────────────────────────────────────────────

_DUMMY_HASH = (
    "$argon2id$v=19$m=65536,t=3,p=2"
    "$dummysaltdummysalt$dummyhashvalue0000000000000000"
)


def _dummy_verify(password: str) -> None:
    """Constant-time guard against user-enumeration via timing."""
    try:
        from argon2 import PasswordHasher
        PasswordHasher().verify(_DUMMY_HASH, password)
    except Exception:
        pass


def _make_tokens(user: User) -> dict:
    """Create a fresh access + refresh token pair for a user."""
    identity = user.to_jwt_identity()
    return {
        "access_token":  create_access_token(identity=identity),
        "refresh_token": create_refresh_token(identity=identity),
        "token_type":    "Bearer",
    }


# ─── Routes ───────────────────────────────────────────────────────────────────

@auth_api_bp.route("/login", methods=["POST"])
@limiter.limit("10 per minute")
def login():
    """
    POST /api/v1/auth/login
    ───────────────────────
    Body (JSON):
        {"username": "alice", "password": "secret"}

    Response 200:
        {
          "access_token":  "<jwt>",
          "refresh_token": "<jwt>",
          "token_type":    "Bearer",
          "user": { ... }
        }

    Response 401:
        {"error": "Invalid credentials."}
    """
    try:
        data = _login_schema.load(request.get_json(force=True, silent=True) or {})
    except ValidationError as exc:
        return jsonify({"error": "Validation failed.", "detail": exc.messages}), 422

    user = User.query.filter_by(username=data["username"].strip()).first()

    if user is None:
        _dummy_verify(data["password"])
        return jsonify({"error": "Invalid credentials."}), 401

    if not user.is_active:
        return jsonify({"error": "Account is disabled. Contact an administrator."}), 403

    if not user.check_password(data["password"]):
        return jsonify({"error": "Invalid credentials."}), 401

    # Record last login time.
    user.last_login = datetime.now(timezone.utc)
    db.session.commit()

    tokens = _make_tokens(user)
    return jsonify({**tokens, "user": user.to_dict()}), 200


@auth_api_bp.route("/refresh", methods=["POST"])
@jwt_required(refresh=True)
def refresh():
    """
    POST /api/v1/auth/refresh
    ─────────────────────────
    Authorization: Bearer <refresh_token>

    Response 200:
        {"access_token": "<new_jwt>", "token_type": "Bearer"}
    """
    identity     = get_jwt_identity()
    access_token = create_access_token(identity=identity)
    return jsonify({"access_token": access_token, "token_type": "Bearer"}), 200


@auth_api_bp.route("/me", methods=["GET"])
@api_permission_required("view_projects")
def me():
    """
    GET /api/v1/auth/me
    ───────────────────
    Authorization: Bearer <access_token>

    Response 200:
        { user profile dict }
    """
    identity = get_jwt_identity()
    user     = db.session.get(User, identity["id"])
    if user is None:
        return jsonify({"error": "User not found."}), 404
    return jsonify(user.to_dict()), 200


@auth_api_bp.route("/logout", methods=["POST"])
@jwt_required()
def logout():
    """
    POST /api/v1/auth/logout
    ────────────────────────
    Server-side JWT blacklisting requires a persistent token store
    (e.g. Redis).  For now, this endpoint returns a success hint so
    clients know they should discard their tokens locally.

    TODO: Implement token blacklisting with Redis when scaling up.

    Response 200:
        {"message": "Logged out. Discard your tokens on the client side."}
    """
    return jsonify({
        "message": (
            "Logged out successfully. "
            "Discard your access and refresh tokens on the client side."
        )
    }), 200