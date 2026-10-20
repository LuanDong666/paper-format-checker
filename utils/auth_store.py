"""Small local account store; no email delivery or external identity provider."""

import re
import secrets
import sqlite3
from datetime import datetime
from pathlib import Path

from werkzeug.security import check_password_hash, generate_password_hash


EMAIL_RE = re.compile(r"^[^\s@]+@[^\s@]+\.[^\s@]+$")
AVATARS = {
    "whale": {"symbol": "🐋", "label": "蓝鲸"},
    "book": {"symbol": "📘", "label": "书本"},
    "star": {"symbol": "⭐", "label": "星星"},
    "cap": {"symbol": "🎓", "label": "学士帽"},
    "leaf": {"symbol": "🍃", "label": "绿叶"},
}


def prepare_auth(data_folder: Path) -> tuple[Path, str]:
    data_folder.mkdir(parents=True, exist_ok=True)
    key_path = data_folder / "session.key"
    if not key_path.exists():
        try:
            with key_path.open("x", encoding="ascii") as file:
                file.write(secrets.token_hex(32))
        except FileExistsError:
            pass
    secret = key_path.read_text(encoding="ascii").strip()
    db_path = data_folder / "users.sqlite3"
    with sqlite3.connect(db_path) as db:
        db.execute("""CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        )""")
        columns = {row[1] for row in db.execute("PRAGMA table_info(users)")}
        if "display_name" not in columns:
            db.execute("ALTER TABLE users ADD COLUMN display_name TEXT NOT NULL DEFAULT ''")
        if "avatar" not in columns:
            db.execute("ALTER TABLE users ADD COLUMN avatar TEXT NOT NULL DEFAULT 'whale'")
        db.execute("""CREATE TABLE IF NOT EXISTS job_history (
            job_id TEXT PRIMARY KEY,
            user_id INTEGER NOT NULL,
            filename TEXT NOT NULL,
            action TEXT NOT NULL,
            template_name TEXT NOT NULL,
            before_count INTEGER NOT NULL,
            after_count INTEGER,
            fix_count INTEGER,
            created_at TEXT NOT NULL,
            FOREIGN KEY(user_id) REFERENCES users(id)
        )""")
        db.execute("CREATE INDEX IF NOT EXISTS job_history_user_time ON job_history(user_id, created_at DESC)")
    return db_path, secret


def normalize_email(email: str) -> str:
    return email.strip().lower()


def valid_email(email: str) -> bool:
    return 3 <= len(email) <= 254 and bool(EMAIL_RE.fullmatch(email))


def create_user(db_path: Path, email: str, password: str) -> int | None:
    try:
        with sqlite3.connect(db_path) as db:
            cursor = db.execute(
                "INSERT INTO users(email, password_hash) VALUES (?, ?)",
                (email, generate_password_hash(password, method="scrypt")),
            )
            return cursor.lastrowid
    except sqlite3.IntegrityError:
        return None


def authenticate(db_path: Path, email: str, password: str) -> dict | None:
    with sqlite3.connect(db_path) as db:
        row = db.execute("SELECT id, email, password_hash, display_name, avatar FROM users WHERE email = ?", (email,)).fetchone()
    if row and check_password_hash(row[2], password):
        return {"id": row[0], "email": row[1], "display_name": row[3] or row[1].split("@")[0], "avatar": row[4]}
    return None


def get_user(db_path: Path, user_id: int) -> dict | None:
    with sqlite3.connect(db_path) as db:
        row = db.execute("SELECT id, email, display_name, avatar FROM users WHERE id = ?", (user_id,)).fetchone()
    return {"id": row[0], "email": row[1], "display_name": row[2] or row[1].split("@")[0], "avatar": row[3]} if row else None


def update_profile(db_path: Path, user_id: int, display_name: str, avatar: str) -> None:
    with sqlite3.connect(db_path) as db:
        db.execute("UPDATE users SET display_name = ?, avatar = ? WHERE id = ?",
                   (display_name, avatar, user_id))


def record_job(db_path: Path, *, job_id: str, user_id: int, filename: str,
               action: str, template_name: str, before_count: int,
               after_count: int | None = None, fix_count: int | None = None) -> None:
    created_at = datetime.now().astimezone().strftime("%Y-%m-%d %H:%M %z")
    with sqlite3.connect(db_path) as db:
        db.execute("""INSERT INTO job_history
            (job_id, user_id, filename, action, template_name, before_count,
             after_count, fix_count, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (job_id, user_id, filename, action, template_name, before_count,
             after_count, fix_count, created_at))


def list_jobs(db_path: Path, user_id: int) -> list[dict]:
    with sqlite3.connect(db_path) as db:
        db.row_factory = sqlite3.Row
        rows = db.execute("""SELECT job_id, filename, action, template_name,
            before_count, after_count, fix_count, created_at
            FROM job_history WHERE user_id = ? ORDER BY created_at DESC, rowid DESC""",
            (user_id,)).fetchall()
    return [dict(row) for row in rows]


def delete_job_record(db_path: Path, user_id: int, job_id: str) -> bool:
    with sqlite3.connect(db_path) as db:
        cursor = db.execute("DELETE FROM job_history WHERE user_id = ? AND job_id = ?",
                            (user_id, job_id))
        return cursor.rowcount == 1
