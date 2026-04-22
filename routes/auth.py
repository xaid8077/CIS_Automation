# routes/auth.py
"""
auth.py
───────
Authentication blueprint — login and logout only.

Design notes
────────────
- Timing-safe dummy hash verification on unknown usernames prevents
  user enumeration via response-time differences.
- Session is fully cleared on logout (not just the login cookie) so
  any server-side session data is also wiped.
- The "next" redirect is validated to be a local path to prevent
  open-redirect attacks.
"""

from datetime import datetime, timezone

from flask import (
    Blueprint, render_template, request,
    redirect, url_for, flash, session,
)
from flask_login import login_user, logout_user, login_required, current_user

from extensions import db, limiter
from forms import LoginForm
from models import User

auth_bp = Blueprint("auth", __name__)


# ─── Login ────────────────────────────────────────────────────────────────────

@auth_bp.route("/login", methods=["GET", "POST"])
@limiter.limit("10 per minute")
def login():
    if current_user.is_authenticated:
        return redirect(url_for("cis.index"))

    form = LoginForm()

    if form.validate_on_submit():
        user = User.query.filter_by(username=form.username.data).first()

        # ── Constant-time guard against user enumeration ──────────────────────
        # Whether the user exists or not, we always run a hash operation so
        # the response time is identical in both branches.
        if user is None:
            _dummy_verify(form.password.data)
            flash("Invalid username or password.", "danger")
            return render_template("login.html", form=form)

        if not user.is_active:
            flash(
                "Your account is disabled. Contact an administrator.",
                "danger",
            )
            return render_template("login.html", form=form)

        if not user.check_password(form.password.data):
            flash("Invalid username or password.", "danger")
            return render_template("login.html", form=form)

        # ── Successful login ──────────────────────────────────────────────────
        login_user(user, remember=form.remember.data)
        user.last_login = datetime.now(timezone.utc)
        db.session.commit()

        next_page = request.args.get("next", "")
        if _is_safe_redirect(next_page):
            return redirect(next_page)
        return redirect(url_for("cis.index"))

    return render_template("login.html", form=form)


# ─── Logout ───────────────────────────────────────────────────────────────────

@auth_bp.route("/logout", methods=["POST"])
@login_required
def logout():
    logout_user()
    session.clear()
    flash("You have been signed out.", "info")
    return redirect(url_for("auth.login"))


# ─── Helpers ──────────────────────────────────────────────────────────────────

_DUMMY_HASH = (
    "$argon2id$v=19$m=65536,t=3,p=2"
    "$dummysaltdummysalt$dummyhashvalue0000000000000000"
)


def _dummy_verify(password: str) -> None:
    """
    Run a throwaway argon2 verification so the login endpoint takes
    the same wall-clock time whether or not the username exists.
    Swallows all exceptions — the result is intentionally discarded.
    """
    try:
        from argon2 import PasswordHasher
        PasswordHasher().verify(_DUMMY_HASH, password)
    except Exception:
        pass


def _is_safe_redirect(target: str) -> bool:
    """
    Return True only for relative paths that don't start with '//'
    (which would be protocol-relative and could redirect off-site).
    """
    return bool(target) and target.startswith("/") and not target.startswith("//")