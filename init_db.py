"""Run this once to initialise the database."""
import sqlite3, os

DATABASE = os.path.join(os.path.dirname(__file__), "instance", "qpapers.db")
os.makedirs(os.path.dirname(DATABASE), exist_ok=True)

db = sqlite3.connect(DATABASE)
db.executescript("""
    CREATE TABLE IF NOT EXISTS papers (
        id           INTEGER PRIMARY KEY AUTOINCREMENT,
        subject_name TEXT    NOT NULL,
        department   TEXT    NOT NULL,
        year         TEXT    NOT NULL,
        file_url     TEXT    NOT NULL,
        public_id    TEXT    NOT NULL DEFAULT '',
        uploaded_by  TEXT    NOT NULL,
        status       TEXT    NOT NULL DEFAULT 'pending',
        upload_date  TEXT    NOT NULL DEFAULT (datetime('now'))
    );
""")
cols = [r[1] for r in db.execute("PRAGMA table_info(papers)").fetchall()]
if "status" not in cols:
    db.execute("ALTER TABLE papers ADD COLUMN status TEXT NOT NULL DEFAULT 'pending'")
if "public_id" not in cols:
    db.execute("ALTER TABLE papers ADD COLUMN public_id TEXT NOT NULL DEFAULT ''")
db.commit()
db.close()
print("Database ready.")
