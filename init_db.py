"""Run this once to initialise the database."""
import os
import psycopg2
from dotenv import load_dotenv

load_dotenv()

DATABASE_URL = os.getenv(
    "DATABASE_URL",
    "postgresql://postgres:password@localhost:5432/qpapers_db"
)

conn = psycopg2.connect(DATABASE_URL)
cur  = conn.cursor()

cur.execute("""
    CREATE TABLE IF NOT EXISTS papers (
        id           SERIAL PRIMARY KEY,
        subject_name TEXT      NOT NULL,
        department   TEXT      NOT NULL,
        year         TEXT      NOT NULL,
        paper_type   TEXT      NOT NULL DEFAULT 'CIE',
        file_url     TEXT      NOT NULL,
        public_id    TEXT      NOT NULL DEFAULT '',
        uploaded_by  TEXT      NOT NULL,
        status       TEXT      NOT NULL DEFAULT 'pending',
        upload_date  TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
    )
""")


cur.execute("""
    CREATE TABLE IF NOT EXISTS activity_logs (
        id           SERIAL PRIMARY KEY,
        student_reg  TEXT      NOT NULL,
        student_name TEXT      NOT NULL,
        action       TEXT      NOT NULL,
        paper_id     INTEGER,
        timestamp    TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
    )
""")

conn.commit()
cur.close()
conn.close()
print("PostgreSQL initialised successfully.")
