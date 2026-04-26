"""Usuarios locales (SQLite) — contraseñas con werkzeug."""

from __future__ import annotations

import os
import sqlite3
from typing import Optional, Tuple
from werkzeug.security import generate_password_hash, check_password_hash

_DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "users.db")


def _conn():
    os.makedirs(os.path.dirname(_DB_PATH), exist_ok=True)
    return sqlite3.connect(_DB_PATH)


# Usuario base (solo se crea si aún no existe; no pisa otras claves)
_DEFAULT_USERNAME = "lucas"
_DEFAULT_PASSWORD = "admin"


def init_db():
    c = _conn()
    try:
        c.execute(
            """CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT NOT NULL UNIQUE COLLATE NOCASE,
                password_hash TEXT NOT NULL,
                created_at TEXT DEFAULT (datetime('now'))
            )"""
        )
        c.commit()
    finally:
        c.close()


def ensure_bootstrap_user() -> None:
    """Crea el usuario 'lucas' / 'admin' la primera vez (clave corta permitida solo aquí)."""
    if user_exists(_DEFAULT_USERNAME):
        return
    c = _conn()
    try:
        ph = generate_password_hash(_DEFAULT_PASSWORD)
        c.execute("INSERT INTO users (username, password_hash) VALUES (?, ?)", (_DEFAULT_USERNAME, ph))
        c.commit()
    except sqlite3.IntegrityError:
        pass
    finally:
        c.close()


def user_exists(username: str) -> bool:
    init_db()
    c = _conn()
    try:
        row = c.execute("SELECT 1 FROM users WHERE username = ? COLLATE NOCASE", (username.strip(),)).fetchone()
        return row is not None
    finally:
        c.close()


def create_user(username: str, password: str) -> Tuple[bool, str]:
    init_db()
    u = username.strip()
    if not u or len(u) < 2:
        return False, "El usuario debe tener al menos 2 caracteres."
    if len(password) < 6:
        return False, "La contraseña debe tener al menos 6 caracteres."
    ph = generate_password_hash(password)
    c = _conn()
    try:
        c.execute("INSERT INTO users (username, password_hash) VALUES (?, ?)", (u, ph))
        c.commit()
        return True, ""
    except sqlite3.IntegrityError:
        return False, "Ese nombre de usuario ya existe."
    finally:
        c.close()


def verify_user(username: str, password: str) -> Optional[int]:
    init_db()
    c = _conn()
    try:
        row = c.execute(
            "SELECT id, password_hash FROM users WHERE username = ? COLLATE NOCASE",
            (username.strip(),),
        ).fetchone()
        if not row:
            return None
        uid, ph = row
        if check_password_hash(ph, password):
            return uid
        return None
    finally:
        c.close()
