"""
create_admin.py
───────────────
CLI script to create the first admin user.

Usage:
    python create_admin.py

Run once after first deployment.  Will prompt for credentials interactively
so nothing sensitive is ever written to shell history or logs.

⚠  DEV SEED (clearly marked):
   For local development only, you can set env vars to skip prompts:
       SEED_ADMIN=1 SEED_USER=admin SEED_EMAIL=admin@local SEED_PASS=adminpass123
   NEVER use this in production.
"""

import os
import sys
import getpass

# Ensure project root is on path
sys.path.insert(0, os.path.dirname(__file__))

from app import create_app
from extensions import db
from models import User


def create_admin_interactive():
    print("\n── Create First Admin ──────────────────────────")
    username = input("Admin username: ").strip()
    email    = input("Admin email:    ").strip()
    password = getpass.getpass("Admin password (min 8 chars): ")
    confirm  = getpass.getpass("Confirm password: ")

    if password != confirm:
        print("❌  Passwords do not match.")
        sys.exit(1)
    if len(password) < 8:
        print("❌  Password must be at least 8 characters.")
        sys.exit(1)

    return username, email, password


def create_admin_from_env():
    """
    ⚠  DEVELOPMENT SEED ONLY — NOT SAFE FOR PRODUCTION.
    Uses environment variables SEED_USER / SEED_EMAIL / SEED_PASS.
    """
    print("⚠  DEV SEED MODE — unsafe for production.")
    username = os.environ["SEED_USER"]
    email    = os.environ["SEED_EMAIL"]
    password = os.environ["SEED_PASS"]
    return username, email, password


def main():
    flask_app = create_app()

    with flask_app.app_context():
        if os.environ.get("SEED_ADMIN") == "1":
            username, email, password = create_admin_from_env()
        else:
            username, email, password = create_admin_interactive()

        if User.query.filter_by(username=username).first():
            print(f"❌  User '{username}' already exists.")
            sys.exit(1)

        admin = User(username=username, email=email, role="admin", is_active=True)
        admin.set_password(password)
        db.session.add(admin)
        db.session.commit()
        print(f"✅  Admin user '{username}' created successfully.")


if __name__ == "__main__":
    main()
