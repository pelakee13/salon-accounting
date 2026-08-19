"""
auth.py — Authentication & Authorization for Salon Accounting (آکادمی هلیا)

Uses SQLite (users.db) for user storage + Flask Session for login state.
Roles: admin (full), reception (front-desk), employee (own commission only).
Backward compatible: app still uses Excel files for business data.
"""
import sqlite3
import os
from functools import wraps
from flask import session, redirect, url_for, request, flash
from werkzeug.security import generate_password_hash, check_password_hash

import sys as _sys
def _auth_base():
    if getattr(_sys, "frozen", False):
        return os.path.dirname(_sys.executable)
    return os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(_auth_base(), "data", "users.db")

ROLES = {
    "admin": "مالک / مدیر",
    "reception": "پذیرش",
    "employee": "کارمند",
}

# Routes each role is allowed to access (aside from shared pages)
ROLE_ALLOWED = {
    "admin": "*",                                   # everything
    "reception": {"/", "/submit_transaction", "/customers", "/customer/",
                 "/reports", "/transaction/delete", "/transaction/edit",
                 "/api/customers", "/api/customers/add"},
    "employee": {"/dashboard", "/reports", "/payroll", "/transaction/delete", "/transaction/edit",
                 "/api/customers", "/api/customers/add"},
}


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_auth():
    """Create users table and seed a default admin if empty."""
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = get_db()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            name TEXT,
            role TEXT NOT NULL DEFAULT 'employee',
            active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT
        )
    """)
    cur = conn.execute("SELECT COUNT(*) AS c FROM users")
    if cur.fetchone()["c"] == 0:
        from web_app import PersianDate
        conn.execute(
            "INSERT INTO users (username, password_hash, name, role, active, created_at) VALUES (?,?,?,?,?,?)",
            ("admin", generate_password_hash("admin123"), "مدیر سالن", "admin", 1, PersianDate.today_str()),
        )
        conn.commit()
    conn.close()


def authenticate(username, password):
    conn = get_db()
    row = conn.execute("SELECT * FROM users WHERE username=?", (username,)).fetchone()
    conn.close()
    if row and row["active"] and check_password_hash(row["password_hash"], password):
        return dict(row)
    return None


def check_password(password, password_hash):
    return check_password_hash(password_hash, password)


def update_password(username, new_password):
    conn = get_db()
    conn.execute(
        "UPDATE users SET password_hash=? WHERE username=?",
        (generate_password_hash(new_password), username),
    )
    conn.commit()
    conn.close()


def get_user(username):
    conn = get_db()
    row = conn.execute("SELECT * FROM users WHERE username=?", (username,)).fetchone()
    conn.close()
    return dict(row) if row else None


def create_user(username, password, name, role):
    conn = get_db()
    conn.execute(
        "INSERT INTO users (username, password_hash, name, role, active, created_at) VALUES (?,?,?,?,?,?)",
        (username, generate_password_hash(password), name, role, 1, __import__("web_app").PersianDate.today_str()),
    )
    conn.commit()
    conn.close()


def login_user(user):
    session["user"] = user["username"]
    session["role"] = user["role"]
    session["name"] = user["name"]


def logout_user():
    session.clear()


def current_user():
    if "user" not in session:
        return None
    return get_user(session["user"])


def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if "user" not in session:
            flash("لطفاً ابتدا وارد شوید", "error")
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated


def role_required(*roles):
    def decorator(f):
        @wraps(f)
        def decorated(*args, **kwargs):
            if "user" not in session:
                flash("لطفاً ابتدا وارد شوید", "error")
                return redirect(url_for("login"))
            if session["role"] not in roles:
                flash("شما دسترسی لازم برای این بخش را ندارید", "error")
                return redirect(url_for("dashboard"))
            return f(*args, **kwargs)
        return decorated
    return decorator


def role_can_access(role, path):
    """Check if a role may access a given path (used for nav rendering)."""
    allowed = ROLE_ALLOWED.get(role)
    if allowed == "*":
        return True
    for prefix in allowed:
        if prefix == "/":
            if path == "/" or path.startswith("/?"):
                return True
        elif path == prefix or path.startswith(prefix.rstrip("/") + "/") or path.startswith(prefix.rstrip("/")):
            return True
    return False
