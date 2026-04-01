"""
forms.py
────────
Flask-WTF forms.

Each form includes a CSRF hidden field automatically (Flask-WTF default).
Validators run server-side regardless of any client-side checks.
"""

from flask_wtf import FlaskForm
from wtforms import (
    StringField, PasswordField, BooleanField,
    SelectField, SubmitField
)
from wtforms.validators import (
    DataRequired, Email, Length,
    EqualTo, ValidationError
)
from models import User, ROLES


class LoginForm(FlaskForm):
    username  = StringField(
        "Username",
        validators=[DataRequired(), Length(min=3, max=64)],
        render_kw={"autocomplete": "username", "placeholder": "Username"},
    )
    password  = PasswordField(
        "Password",
        validators=[DataRequired(), Length(min=8, max=128)],
        render_kw={"autocomplete": "current-password", "placeholder": "Password"},
    )
    remember  = BooleanField("Remember me")
    submit    = SubmitField("Sign in")


class RegisterForm(FlaskForm):
    """Admin-only: create a new user account."""
    username  = StringField(
        "Username",
        validators=[DataRequired(), Length(min=3, max=64)],
        render_kw={"placeholder": "Username"},
    )
    email     = StringField(
        "Email",
        validators=[DataRequired(), Email(), Length(max=120)],
        render_kw={"placeholder": "user@example.com"},
    )
    password  = PasswordField(
        "Password",
        validators=[DataRequired(), Length(min=8, max=128)],
        render_kw={"placeholder": "Min 8 characters"},
    )
    password2 = PasswordField(
        "Confirm Password",
        validators=[DataRequired(), EqualTo("password", message="Passwords must match.")],
        render_kw={"placeholder": "Repeat password"},
    )
    role      = SelectField(
        "Role",
        choices=[(r, r.capitalize()) for r in ROLES],
        default="user",
    )
    submit    = SubmitField("Create User")

    def validate_username(self, field):
        if User.query.filter_by(username=field.data).first():
            raise ValidationError("Username already taken.")

    def validate_email(self, field):
        if User.query.filter_by(email=field.data).first():
            raise ValidationError("Email already registered.")


class EditUserForm(FlaskForm):
    """Admin-only: change a user's role or active status."""
    role      = SelectField(
        "Role",
        choices=[(r, r.capitalize()) for r in ROLES],
    )
    is_active = BooleanField("Active")
    submit    = SubmitField("Save Changes")
