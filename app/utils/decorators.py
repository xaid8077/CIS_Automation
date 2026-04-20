# app/utils/decorators.py

import functools
from flask_login import login_required, current_user
from flask import abort


def admin_required(f):
    @functools.wraps(f)
    @login_required
    def wrapped(*args, **kwargs):
        if not current_user.is_admin:
            abort(403)
        return f(*args, **kwargs)
    return wrapped