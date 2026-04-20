"""
forms.py
────────
Flask-WTF forms.
"""

from flask_wtf import FlaskForm
from wtforms import (
    StringField, PasswordField, BooleanField,
    SelectField, SubmitField
)
from wtforms.validators import (
    DataRequired, Email, Length, Optional,
    EqualTo, ValidationError
)
from models import User, Project, ROLES


# ─────────────────────────────────────────────────────────────────────────────
# Auth
# ─────────────────────────────────────────────────────────────────────────────

class LoginForm(FlaskForm):
    username = StringField("Username",  validators=[DataRequired(), Length(min=3, max=64)])
    password = PasswordField("Password", validators=[DataRequired(), Length(min=8, max=128)])
    remember = BooleanField("Remember me")
    submit   = SubmitField("Sign in")


class RegisterForm(FlaskForm):
    username  = StringField("Username",           validators=[DataRequired(), Length(min=3, max=64)])
    email     = StringField("Email",              validators=[DataRequired(), Email(), Length(max=120)])
    password  = PasswordField("Password",         validators=[DataRequired(), Length(min=8, max=128)])
    password2 = PasswordField("Confirm Password", validators=[DataRequired(), EqualTo("password")])
    role      = SelectField("Role", choices=[(r, r.capitalize()) for r in ROLES], default="user")
    submit    = SubmitField("Create User")

    def validate_username(self, field):
        if User.query.filter_by(username=field.data).first():
            raise ValidationError("Username already taken.")

    def validate_email(self, field):
        if User.query.filter_by(email=field.data).first():
            raise ValidationError("Email already registered.")


class EditUserForm(FlaskForm):
    role      = SelectField("Role", choices=[(r, r.capitalize()) for r in ROLES])
    is_active = BooleanField("Active")
    submit    = SubmitField("Save Changes")


# ─────────────────────────────────────────────────────────────────────────────
# Project
# ─────────────────────────────────────────────────────────────────────────────

class ProjectForm(FlaskForm):
    """Admin-only: create a new project repository."""
    name       = StringField(
        "Project Name",
        validators=[DataRequired()],           # no length cap
        render_kw={"placeholder": "e.g., Chandrawal Water Treatment Plant"},
    )
    nickname   = StringField(
        "Nickname (short display name)",
        validators=[Optional(), Length(max=40)],
        render_kw={"placeholder": "e.g., Chandrawal WTP  (max 40 chars)"},
    )
    client     = StringField(
        "Client",
        validators=[DataRequired()],
        render_kw={"placeholder": "e.g., Delhi Jal Board"},
    )
    consultant = StringField(
        "Consultant",
        validators=[Optional()],
        render_kw={"placeholder": "Optional"},
    )
    submit     = SubmitField("Create Project")

    def validate_name(self, field):
        if Project.query.filter_by(name=field.data).first():
            raise ValidationError("A project with this name already exists.")


class EditProjectNameForm(FlaskForm):
    """
    Admin-only: edit a project's full name (and optionally nickname/client/
    consultant) after creation.  No length cap on name so long project titles
    used in engineering documents are not truncated.
    """
    name       = StringField(
        "Project Name",
        validators=[DataRequired()],
        render_kw={"placeholder": "Full project name as it appears on documents"},
    )
    nickname   = StringField(
        "Nickname",
        validators=[Optional(), Length(max=40)],
        render_kw={"placeholder": "Short display name (max 40 chars)"},
    )
    client     = StringField(
        "Client",
        validators=[DataRequired()],
        render_kw={"placeholder": "Client name"},
    )
    consultant = StringField(
        "Consultant",
        validators=[Optional()],
        render_kw={"placeholder": "Consultant (optional)"},
    )
    submit     = SubmitField("Save Changes")


class ProjectNicknameForm(FlaskForm):
    """
    Admin-only standalone form to set or update a project's nickname
    without touching any other project fields.
    """
    nickname = StringField(
        "Nickname",
        validators=[Optional(), Length(max=40)],
        render_kw={"placeholder": "Short display name (max 40 chars)"},
    )
    submit   = SubmitField("Save Nickname")


class RevisionDocNumbersForm(FlaskForm):
    """
    Admin-only: update the document numbers stored in an existing revision's
    data_payload without regenerating the file.  Useful for correcting a
    doc number after a revision has already been published.
    """
    fi_doc_number  = StringField("IL Doc Number",  validators=[Optional()])
    el_doc_number  = StringField("EL Doc Number",  validators=[Optional()])
    mov_doc_number = StringField("MOV Doc Number", validators=[Optional()])
    io_doc_number  = StringField("IO Doc Number",  validators=[Optional()])
    submit         = SubmitField("Update Doc Numbers")


# ─────────────────────────────────────────────────────────────────────────────
# Project Location
# ─────────────────────────────────────────────────────────────────────────────

class ProjectLocationForm(FlaskForm):
    """
    Admin-only: add a location to a project.
    """
    name = StringField(
        "Location Name",
        validators=[DataRequired()],
        render_kw={"placeholder": "e.g., Plant A – Intake Block"},
    )
    code = StringField(
        "Short Code",
        validators=[Optional(), Length(max=20)],
        render_kw={"placeholder": "e.g., PA-INT  (optional)"},
    )
    submit = SubmitField("Add Location")