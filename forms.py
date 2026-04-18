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
    DataRequired, Email, Length,
    EqualTo, ValidationError
)
from models import User, Project, ROLES


class LoginForm(FlaskForm):
    username  = StringField("Username", validators=[DataRequired(), Length(min=3, max=64)])
    password  = PasswordField("Password", validators=[DataRequired(), Length(min=8, max=128)])
    remember  = BooleanField("Remember me")
    submit    = SubmitField("Sign in")


class RegisterForm(FlaskForm):
    username  = StringField("Username", validators=[DataRequired(), Length(min=3, max=64)])
    email     = StringField("Email", validators=[DataRequired(), Email(), Length(max=120)])
    password  = PasswordField("Password", validators=[DataRequired(), Length(min=8, max=128)])
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


class ProjectForm(FlaskForm):
    """Admin-only: create a new project repository."""
    name       = StringField("Project Name", validators=[DataRequired(), Length(max=150)], render_kw={"placeholder": "e.g., Chandrawal W.T.P."})
    client     = StringField("Client", validators=[DataRequired(), Length(max=150)], render_kw={"placeholder": "e.g., Delhi Jal Board"})
    consultant = StringField("Consultant", validators=[Length(max=150)], render_kw={"placeholder": "Optional"})
    location   = StringField("Location", validators=[Length(max=150)], render_kw={"placeholder": "Site Location"})
    submit     = SubmitField("Create Project")

    def validate_name(self, field):
        if Project.query.filter_by(name=field.data).first():
            raise ValidationError("A project with this name already exists.")