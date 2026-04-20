# app/routes/auth_routes.py

import traceback
from datetime import datetime, timezone

from flask import (
    Blueprint, render_template, request,
    redirect, url_for, flash, session
)
from flask_login import (
    login_user, logout_user,
    login_required, current_user
)

from app.extensions import db, limiter
from app.models import User
from app.forms import LoginForm


auth_bp = Blueprint("auth", __name__)


# ─────────────────────────────────────────────────────────────────────────────
# Login
# ─────────────────────────────────────────────────────────────────────────────

@auth_bp.route("/login", methods=["GET", "POST"])
@limiter.limit("10 per minute")
def login():
    if current_user.is_authenticated:
        return redirect(url_for("cis.index"))

    form = LoginForm()

    if form.validate_on_submit():
        user = User.query.filter_by(username=form.username.data).first()

        # Dummy hash to prevent timing attacks
        dummy_hash = "$argon2id$v=19$m=65536,t=3,p=2$dummy$dummy"

        if user is None:
            try:
                from argon2 import PasswordHasher
                PasswordHasher().verify(dummy_hash, form.password.data)
            except Exception:
                pass

            flash("Invalid username or password.", "danger")
            return render_template("login.html", form=form)

        if not user.is_active:
            flash("Your account is disabled. Contact an administrator.", "danger")
            return render_template("login.html", form=form)

        if not user.check_password(form.password.data):
            flash("Invalid username or password.", "danger")
            return render_template("login.html", form=form)

        login_user(user, remember=form.remember.data)

        user.last_login = datetime.now(timezone.utc)
        db.session.commit()

        next_page = request.args.get("next", "")

        if next_page and next_page.startswith("/") and not next_page.startswith("//"):
            return redirect(next_page)

        return redirect(url_for("cis.index"))

    return render_template("login.html", form=form)


# ─────────────────────────────────────────────────────────────────────────────
# Logout
# ─────────────────────────────────────────────────────────────────────────────

@auth_bp.route("/logout", methods=["POST"])
@login_required
def logout():
    logout_user()
    session.clear()

    flash("You have been signed out.", "info")
    return redirect(url_for("auth.login"))