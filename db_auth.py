import sqlite3
from pathlib import Path

DB_PATH = Path(__file__).resolve().parent / "facturador.db"

def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def ensure_users_table():
    with get_conn() as conn:
        conn.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            is_active INTEGER NOT NULL DEFAULT 1,
	    role TEXT NOT NULL DEFAULT 'admin',
            locked_emisor_id INTEGER,
            locked_receptor_id INTEGER
        );
        """)
        conn.commit()
def get_user_by_username(username: str):
    with get_conn() as conn:
        cur = conn.execute("SELECT * FROM users WHERE username = ?", (username,))
        return cur.fetchone()

def create_user(username: str, password_hash: str):
    with get_conn() as conn:
        conn.execute(
            "INSERT INTO users (username, password_hash) VALUES (?, ?)",
            (username, password_hash),
        )
        conn.commit()
def update_user_lock(user_id: int, locked_emisor_id=None, locked_receptor_id=None, role=None):
    fields = []
    params = []

    if role is not None:
        fields.append("role = ?")
        params.append(role)

    if locked_emisor_id is not None:
        fields.append("locked_emisor_id = ?")
        params.append(locked_emisor_id)

    if locked_receptor_id is not None:
        fields.append("locked_receptor_id = ?")
        params.append(locked_receptor_id)

    if not fields:
        return

    params.append(user_id)
    q = "UPDATE users SET " + ", ".join(fields) + " WHERE id = ?"

    with get_conn() as conn:
        conn.execute(q, params)
        conn.commit()

