# utils/rbac.py
"""
utils/rbac.py
─────────────
Role-based access control decorators and helpers.

Two decorator families are provided:

  Web (session-based, Flask-Login)
  ─────────────────────────────────
      @permission_required("generate_document")
      def my_route(): ...

  API (JWT-based, Flask-JWT-Extended)
  ────────────────────────────────────
      @api_permission_required("manage_projects")
      def my_api_route(): ...

Both families share the same underlying _PERMISSIONS matrix defined in
models.py, so adding a new role or permission is always a single-file
change in models.py.

Helper utilities
────────────────
  current_user_has(permission)  — bool check for use in templates/services.
  require_same_user_or_admin()  — guards user-specific endpoints.
"""

import functools
from typing import Callable

from flask import abort, jsonify
from flask_login import current_user as _web_user, login_required
from flask_jwt_extended import (
    verify_jwt_in_request,
    get_jwt_identity,
    get_jwt,
)

from models import has_permission, ROLES


# ─── Web decorators (session / Flask-Login) ───────────────────────────────────

def permission_required(permission: str) -> Callable:
    """
    Decorator that enforces a named permission for session-authenticated
    (web UI) routes.

    Combines @login_required with a role check, returning 403 when the
    authenticated user's role does not include the permission.

    Usage
    ─────
        @cis_bp.route("/project/<int:project_id>/location/<int:loc_id>/save-draft",
                       methods=["POST"])
        @permission_required("generate_document")
        def save_draft(project_id, loc_id):
            ...
    """
    def decorator(f: Callable) -> Callable:
        @functools.wraps(f)
        @login_required
        def wrapped(*args, **kwargs):
            if not has_permission(_web_user.role, permission):
                abort(403)
            return f(*args, **kwargs)
        return wrapped
    return decorator


def admin_required(f: Callable) -> Callable:
    """
    Shortcut decorator — equivalent to @permission_required("manage_projects")
    but with a clearer name for routes that require full admin access.

    Kept as a standalone decorator (rather than an alias) so the intent
    is explicit at the call site.
    """
    @functools.wraps(f)
    @login_required
    def wrapped(*args, **kwargs):
        if not _web_user.is_admin:
            abort(403)
        return f(*args, **kwargs)
    return wrapped


# ─── API decorators (JWT / Flask-JWT-Extended) ────────────────────────────────

def api_permission_required(permission: str, optional: bool = False) -> Callable:
    """
    Decorator that enforces a named permission for JWT-authenticated
    (API) routes.

    On failure, returns a JSON error response rather than aborting with
    an HTML 403, which is appropriate for API consumers.

    Parameters
    ----------
    permission: Key from models._PERMISSIONS.
    optional:   If True, an anonymous request is allowed through but will
                not have identity information.  Used for public endpoints
                that provide richer data when authenticated.

    Usage
    ─────
        @projects_bp.route("/", methods=["GET"])
        @api_permission_required("view_projects")
        def list_projects():
            ...
    """
    def decorator(f: Callable) -> Callable:
        @functools.wraps(f)
        def wrapped(*args, **kwargs):
            # Validate the JWT and extract claims.
            try:
                verify_jwt_in_request(optional=optional)
            except Exception as exc:
                return jsonify({"error": "Invalid or missing token.", "detail": str(exc)}), 401

            identity = get_jwt_identity()
            if identity is None and optional:
                # Unauthenticated but optional — proceed without a role check.
                return f(*args, **kwargs)

            if identity is None:
                return jsonify({"error": "Authentication required."}), 401

            role = identity.get("role", "")
            if role not in ROLES:
                return jsonify({"error": f"Unknown role '{role}'."}), 403

            if not has_permission(role, permission):
                return jsonify({
                    "error":      "Permission denied.",
                    "required":   permission,
                    "your_role":  role,
                }), 403

            return f(*args, **kwargs)
        return wrapped
    return decorator


def api_admin_required(f: Callable) -> Callable:
    """
    Shortcut for API routes that require the admin role.
    Returns JSON errors (not HTML) on auth failure.
    """
    @functools.wraps(f)
    def wrapped(*args, **kwargs):
        try:
            verify_jwt_in_request()
        except Exception as exc:
            return jsonify({"error": "Invalid or missing token.", "detail": str(exc)}), 401

        identity = get_jwt_identity()
        if not identity or identity.get("role") != "admin":
            return jsonify({"error": "Admin access required."}), 403

        return f(*args, **kwargs)
    return wrapped


# ─── Convenience helpers ──────────────────────────────────────────────────────

def current_user_has(permission: str) -> bool:
    """
    Boolean check for use in Jinja templates and service code.

    Returns False when no user is authenticated.

    Example (template)
    ──────────────────
        {% if current_user_has('generate_document') %}
          <button>Generate</button>
        {% endif %}

    Example (service)
    ─────────────────
        from utils.rbac import current_user_has
        if not current_user_has("patch_doc_numbers"):
            raise PermissionError(...)
    """
    if not _web_user.is_authenticated:
        return False
    return has_permission(_web_user.role, permission)


def api_get_current_user_id() -> int | None:
    """
    Extract the current user's ID from a validated JWT.

    Returns None when no valid JWT is present.
    Safe to call inside routes that used @api_permission_required.
    """
    identity = get_jwt_identity()
    return identity.get("id") if identity else None


def api_get_current_role() -> str | None:
    """
    Extract the current user's role from a validated JWT.

    Returns None when no valid JWT is present.
    """
    identity = get_jwt_identity()
    return identity.get("role") if identity else None


def require_same_user_or_admin(target_user_id: int) -> None:
    """
    Abort with 403 if the web-session user is neither the target user
    nor an admin.

    Used to guard endpoints like "change my own password" where the
    engineer/viewer can act on their own record but not others'.

    Parameters
    ----------
    target_user_id: The user ID the route is acting on.
    """
    if not _web_user.is_authenticated:
        abort(401)
    if _web_user.id != target_user_id and not _web_user.is_admin:
        abort(403)