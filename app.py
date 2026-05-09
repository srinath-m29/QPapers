"""
QPapers — Flask Question Paper Portal
Database layer: PostgreSQL via psycopg2 + RealDictCursor
All other behaviour (routes, templates, session keys, Cloudinary, Excel auth) unchanged.
"""

import os
from datetime import datetime
from functools import wraps

import cloudinary
import cloudinary.uploader
import cloudinary.utils
import pandas as pd
import psycopg2
from psycopg2.extras import RealDictCursor
import requests
from dotenv import load_dotenv
from flask import (Flask, Response, flash, g, jsonify, redirect,
                   render_template, request, session, url_for)
from werkzeug.utils import secure_filename

load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "dev-secret-key-change-me")

# ── Cloudinary config ──────────────────────────────────────────────────────────
cloudinary.config(
    cloud_name=os.getenv("CLOUDINARY_CLOUD_NAME"),
    api_key=os.getenv("CLOUDINARY_API_KEY"),
    api_secret=os.getenv("CLOUDINARY_API_SECRET"),
    secure=True,
)

# ══════════════════════════════════════════════════════════════════════════════
#  DATABASE LAYER  (the only section that changed from the SQLite version)
# ══════════════════════════════════════════════════════════════════════════════

# FIX 1: DATABASE_URL must be defined ONCE, before any function that uses it.
#         The old import_os.py defined it twice (lines 18-19 and 128-131),
#         and the first definition could be None if the env var was missing,
#         causing psycopg2.connect(None) to crash immediately.
DATABASE_URL = os.getenv(
    "DATABASE_URL",
    "postgresql://postgres:password@localhost:5432/qpapers_db"
)

MAX_FILE_SIZE = 10 * 1024 * 1024

DEPARTMENTS = [
    "Computer Science", "Information Technology",
    "Electronics & Communication", "Electrical Engineering",
    "Mechanical Engineering", "Civil Engineering",
    "Biotechnology", "Mathematics", "Physics", "Chemistry",
]
YEARS = ["1st Year", "2nd Year", "3rd Year", "4th Year"]


def get_db():
    """
    Return a per-request psycopg2 connection stored on Flask's g object.
    RealDictCursor makes rows behave like dicts (row["column"]) —
    the same interface sqlite3.Row provided, so templates need no changes.

    FIX 2: The old import_os.py passed cursor_factory to psycopg2.connect(),
            which is NOT a valid connect() argument.  cursor_factory must be
            passed to conn.cursor(), not to connect().
    """
    if "db" not in g:
        g.db = psycopg2.connect(DATABASE_URL)   # ← no cursor_factory here
    return g.db


def _cursor(conn):
    """Open a RealDictCursor on an existing connection."""
    return conn.cursor(cursor_factory=RealDictCursor)


def query_db(query, params=(), one=False):
    """
    Unified query helper.

    FIX 3: The old version returned None for non-SELECT queries but then
            callers did query_db(...).fetchone() or .fetchall() on that None,
            causing AttributeError: 'NoneType' object has no attribute 'fetchone'.
            Solution: always return the data directly (list or single row),
            never a cursor object that the caller then has to call again.

    FIX 4: query_db() now handles commit internally for write queries, so
            callers don't need a separate db.commit() after every INSERT/UPDATE/DELETE.
            (Callers that already called db.commit() explicitly are harmless but
            redundant — we keep them for clarity.)

    FIX 5: PostgreSQL uses %s placeholders, NOT ?.
            This helper is the single place where queries arrive, so placeholders
            in every SQL string must be %s (see all routes below).
    """
    conn = get_db()
    cur  = _cursor(conn)
    cur.execute(query, params)

    stripped = query.strip().lower()
    if stripped.startswith("select"):
        rows = cur.fetchall()          # list of RealDictRow objects
        cur.close()
        return rows[0] if one else rows
    else:
        conn.commit()
        cur.close()
        return None                    # callers must NOT call .fetchone() on this


@app.teardown_appcontext
def close_db(error=None):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def init_db():
    """
    Create tables on first run.

    FIX 6: init_db.py had a trailing comma after the last column in
            activity_logs, which is a PostgreSQL syntax error (unlike SQLite
            which tolerates it).  Removed.

    FIX 7: SQLite used INTEGER PRIMARY KEY AUTOINCREMENT + datetime('now').
            PostgreSQL equivalents: SERIAL PRIMARY KEY + CURRENT_TIMESTAMP.
            Already correct in import_os.py, kept as-is.

    FIX 8: The PRAGMA table_info() migration check used in app.py is
            SQLite-only.  In PostgreSQL, CREATE TABLE IF NOT EXISTS is
            sufficient — no PRAGMA needed.
    """
    conn = psycopg2.connect(DATABASE_URL)
    cur  = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS papers (
            id           SERIAL PRIMARY KEY,
            subject_name TEXT      NOT NULL,
            department   TEXT      NOT NULL,
            year         TEXT      NOT NULL,
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
    # Note: no trailing comma after the last column — PostgreSQL rejects it.

    conn.commit()
    cur.close()
    conn.close()
    print("PostgreSQL database initialised successfully.")


# ── Excel Student Registry ────────────────────────────────────────────────────
EXCEL_FILE = os.path.join(os.path.dirname(__file__), "ECE.xlsx")
STUDENT_REGISTRY: dict = {}


def load_excel_students():
    global STUDENT_REGISTRY
    if not os.path.exists(EXCEL_FILE):
        print(f"WARNING: Excel file not found at {EXCEL_FILE}")
        return
    df = pd.read_excel(EXCEL_FILE, dtype=str)
    df.columns = [c.strip() for c in df.columns]
    for _, row in df.iterrows():
        reg = str(row.get("Reg No", "")).strip()
        if not reg or reg == "nan":
            continue
        STUDENT_REGISTRY[reg] = {
            "name":     str(row.get("Student Name", "")).strip(),
            "semester": str(row.get("Sem", "")).strip(),
            "status":   str(row.get("Status", "")).strip().upper(),
            "section":  str(row.get("SEC", "")).strip(),
        }
    print(f"Loaded {len(STUDENT_REGISTRY)} students from Excel.")


# ── Activity Logging ──────────────────────────────────────────────────────────

def log_activity(student_reg, student_name, action, paper_id=None):
    """
    Insert a row into activity_logs using its own short-lived connection
    so that a logging failure never rolls back the caller's transaction.

    FIX 9: The old version passed datetime.now() as the 6th parameter but
            the INSERT only listed 5 columns (the timestamp column has a
            server-side DEFAULT CURRENT_TIMESTAMP).  Passing an explicit value
            is fine — but the placeholder count must match the value count.
            We simply let the DB default handle it (omit timestamp from INSERT).

    FIX 10: The old force_download route opened a fresh psycopg2 connection
             with cursor_factory=RealDictCursor passed to connect() — invalid.
             Now force_download uses get_db() + query_db() like every other route.
    """
    try:
        conn = psycopg2.connect(DATABASE_URL)
        cur  = conn.cursor()
        cur.execute(
            """INSERT INTO activity_logs (student_reg, student_name, action, paper_id)
               VALUES (%s, %s, %s, %s)""",
            (student_reg, student_name, action, paper_id),
        )
        conn.commit()
        cur.close()
        conn.close()
    except Exception as exc:
        print(f"[activity_log] Failed to log '{action}' for {student_reg}: {exc}")


# ══════════════════════════════════════════════════════════════════════════════
#  AUTH DECORATORS  (unchanged from original)
# ══════════════════════════════════════════════════════════════════════════════

def student_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("student_reg"):
            flash("Please login to continue.", "warning")
            return redirect(url_for("student_login"))
        return f(*args, **kwargs)
    return decorated


def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("is_admin"):
            flash("Admin access required.", "danger")
            return redirect(url_for("admin_login"))
        return f(*args, **kwargs)
    return decorated


# ══════════════════════════════════════════════════════════════════════════════
#  ROUTES
# ══════════════════════════════════════════════════════════════════════════════

@app.route("/")
def index():
    return redirect(url_for("student_login"))


# ── Force download ────────────────────────────────────────────────────────────

@app.route("/download/<int:paper_id>")
def force_download(paper_id):
    """
    FIX 11: Old version opened a raw psycopg2.connect() here with
             cursor_factory=RealDictCursor as a connect() argument (invalid),
             bypassing get_db() entirely and leaking a connection on every call.
             Now uses query_db() consistently.
    """
    paper = query_db("SELECT * FROM papers WHERE id = %s", (paper_id,), one=True)

    if not paper:
        return "Paper not found.", 404

    file_url = paper["file_url"]
    subject  = paper["subject_name"]

    if session.get("student_reg"):
        log_activity(session["student_reg"], session.get("student_name", ""), "download", paper_id)

    safe_name = subject.strip().replace(" ", "_").replace("/", "_")
    if not safe_name.lower().endswith(".pdf"):
        safe_name += ".pdf"

    try:
        resp = requests.get(file_url, stream=True, timeout=30)
        if resp.status_code != 200:
            return f"File not available on Cloudinary (HTTP {resp.status_code}).", 404
        return Response(
            resp.iter_content(chunk_size=8192),
            headers={
                "Content-Type": "application/pdf",
                "Content-Disposition": f'attachment; filename="{safe_name}"',
                "Content-Length": resp.headers.get("Content-Length", ""),
            }
        )
    except requests.exceptions.Timeout:
        return "Download timed out. Please try again.", 504
    except Exception as e:
        return f"Download error: {str(e)}", 500


@app.template_filter("download_url")
def download_url_filter(file_url, filename="paper.pdf"):
    if not file_url:
        return ""
    filename = filename.replace(" ", "_").replace("/", "_")
    if not filename.lower().endswith(".pdf"):
        filename += ".pdf"
    return file_url.replace("/raw/upload/", f"/raw/upload/fl_attachment:{filename}/")


# ── Student routes ────────────────────────────────────────────────────────────

@app.route("/login", methods=["GET", "POST"])
def student_login():
    if session.get("student_reg"):
        return redirect(url_for("student_dashboard"))

    if request.method == "POST":
        reg = request.form.get("register_number", "").strip()
        pwd = request.form.get("password", "").strip()

        if not reg or not pwd:
            flash("All fields are required.", "danger")
            return render_template("student_login.html")

        student_data = STUDENT_REGISTRY.get(reg)
        if not student_data:
            flash("Register number not found. Please check your number.", "danger")
            return render_template("student_login.html")

        if pwd != reg:
            flash("Incorrect password. Your password is your register number.", "danger")
            return render_template("student_login.html")

        if student_data["status"] != "REGULAR":
            flash("Your account is not active. Contact admin.", "danger")
            return render_template("student_login.html")

        session["student_reg"]  = reg
        session["student_name"] = student_data["name"]
        session["semester"]     = student_data["semester"]
        session["section"]      = student_data["section"]
        log_activity(reg, student_data["name"], "login")
        flash(f"Welcome, {student_data['name']}!", "success")
        return redirect(url_for("student_dashboard"))

    return render_template("student_login.html")


@app.route("/logout")
def student_logout():
    session.clear()
    flash("You have been logged out.", "info")
    return redirect(url_for("student_login"))


@app.route("/dashboard")
@student_required
def student_dashboard():
    """
    FIX 12: Old code called query_db(...).fetchall() / .fetchone() — but
             query_db() already returns the data, not a cursor.
             All such calls are fixed throughout this file.
    FIX 13: ? placeholders replaced with %s everywhere.
    """
    recent = query_db(
        "SELECT * FROM papers WHERE status='approved' ORDER BY upload_date DESC LIMIT 6"
    )
    total = query_db(
        "SELECT COUNT(*) AS c FROM papers WHERE status='approved'", one=True
    )["c"]
    depts = query_db(
        "SELECT COUNT(DISTINCT department) AS c FROM papers WHERE status='approved'", one=True
    )["c"]
    my_pending = query_db(
        "SELECT COUNT(*) AS c FROM papers WHERE uploaded_by=%s AND status='pending'",
        (session["student_reg"],), one=True
    )["c"]
    return render_template(
        "student_dashboard.html",
        recent=recent, total=total, depts=depts,
        my_pending=my_pending,
        departments=DEPARTMENTS, years=YEARS,
    )


@app.route("/papers")
@student_required
def papers():
    dept    = request.args.get("department", "")
    year    = request.args.get("year", "")
    subject = request.args.get("subject", "").strip()

    # FIX 14: dynamic filter query must use %s, not ?
    sql    = "SELECT * FROM papers WHERE status='approved'"
    params = []
    if dept:
        sql += " AND department = %s";        params.append(dept)
    if year:
        sql += " AND year = %s";              params.append(year)
    if subject:
        sql += " AND subject_name LIKE %s";   params.append(f"%{subject}%")
    sql += " ORDER BY upload_date DESC"

    all_papers = query_db(sql, params)
    return render_template(
        "papers.html",
        papers=all_papers, departments=DEPARTMENTS, years=YEARS,
        selected_dept=dept, selected_year=year, search_subject=subject,
    )


@app.route("/upload", methods=["GET", "POST"])
@student_required
def upload():
    if request.method == "POST":
        subject = request.form.get("subject_name", "").strip()
        dept    = request.form.get("department", "").strip()
        year    = request.form.get("year", "").strip()
        file    = request.files.get("file")

        if not all([subject, dept, year, file]):
            flash("All fields and a file are required.", "danger")
            return render_template("upload.html", departments=DEPARTMENTS, years=YEARS)
        if not file.filename.lower().endswith(".pdf"):
            flash("Only PDF files are allowed.", "danger")
            return render_template("upload.html", departments=DEPARTMENTS, years=YEARS)

        file.seek(0, 2); size = file.tell(); file.seek(0)
        if size > MAX_FILE_SIZE:
            flash("File size must not exceed 10 MB.", "danger")
            return render_template("upload.html", departments=DEPARTMENTS, years=YEARS)

        filename = secure_filename(file.filename)
        name_without_ext = os.path.splitext(filename)[0]

        try:
            result = cloudinary.uploader.upload(
                file,
                resource_type="raw",
                folder="qpapers",
                public_id=name_without_ext,
                use_filename=True,
                unique_filename=True,
            )
            file_url  = result["secure_url"]
            public_id = result["public_id"]
        except Exception as e:
            flash(f"Upload failed: {str(e)}", "danger")
            return render_template("upload.html", departments=DEPARTMENTS, years=YEARS)

        # FIX 15: ? → %s; query_db handles commit internally
        query_db(
            """INSERT INTO papers
               (subject_name, department, year, file_url, public_id, uploaded_by, status)
               VALUES (%s, %s, %s, %s, %s, %s, 'pending')""",
            (subject, dept, year, file_url, public_id, session["student_reg"]),
        )
        log_activity(session["student_reg"], session["student_name"], "upload")
        flash("Paper uploaded! It is pending admin approval.", "success")
        return redirect(url_for("student_dashboard"))

    return render_template("upload.html", departments=DEPARTMENTS, years=YEARS)


@app.route("/view/<int:paper_id>")
@student_required
def view_paper(paper_id):
    """
    FIX 16: Old code called query_db(...).fetchone() — query_db already
             returns the row when one=True; no .fetchone() needed.
    """
    paper = query_db(
        "SELECT * FROM papers WHERE id = %s AND status = 'approved'",
        (paper_id,), one=True
    )
    if not paper:
        flash("Paper not found or not yet approved.", "danger")
        return redirect(url_for("papers"))

    log_activity(session["student_reg"], session["student_name"], "view", paper_id)
    return redirect(paper["file_url"])


# ── Admin routes ──────────────────────────────────────────────────────────────

@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    if session.get("is_admin"):
        return redirect(url_for("admin_dashboard"))
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip()
        if (username == os.getenv("ADMIN_USERNAME", "admin") and
                password == os.getenv("ADMIN_PASSWORD", "admin123")):
            session["is_admin"]   = True
            session["admin_name"] = username
            flash("Welcome, Admin!", "success")
            return redirect(url_for("admin_dashboard"))
        flash("Invalid admin credentials.", "danger")
    return render_template("admin_login.html")


@app.route("/admin/logout")
def admin_logout():
    session.clear()
    flash("Admin logged out.", "info")
    return redirect(url_for("admin_login"))


@app.route("/admin/dashboard")
@admin_required
def admin_dashboard():
    """
    FIX 17: Old code called query_db(...).fetchone()["c"] — query_db returns
             the dict directly when one=True; just index into it.
    FIX 18: query_db(...).fetchall() → query_db(...) (already a list).
    """
    total_papers   = query_db("SELECT COUNT(*) AS c FROM papers", one=True)["c"]
    pending_count  = query_db("SELECT COUNT(*) AS c FROM papers WHERE status='pending'", one=True)["c"]
    approved_count = query_db("SELECT COUNT(*) AS c FROM papers WHERE status='approved'", one=True)["c"]
    rejected_count = query_db("SELECT COUNT(*) AS c FROM papers WHERE status='rejected'", one=True)["c"]
    total_students = len(STUDENT_REGISTRY)
    recent_papers  = query_db("SELECT * FROM papers ORDER BY upload_date DESC LIMIT 10")
    dept_stats     = query_db(
        "SELECT department, COUNT(*) AS count FROM papers WHERE status='approved' "
        "GROUP BY department ORDER BY count DESC"
    )
    recent_activity = query_db(
        "SELECT * FROM activity_logs ORDER BY timestamp DESC LIMIT 10"
    )
    return render_template(
        "admin_dashboard.html",
        total_papers=total_papers, pending_count=pending_count,
        approved_count=approved_count, rejected_count=rejected_count,
        total_students=total_students,
        recent_papers=recent_papers, dept_stats=dept_stats,
        recent_activity=recent_activity,
    )


@app.route("/admin/papers")
@admin_required
def admin_papers():
    status = request.args.get("status", "")
    # FIX 19: ? → %s; removed spurious .fetchall() call
    if status in ("pending", "approved", "rejected"):
        all_papers = query_db(
            "SELECT * FROM papers WHERE status=%s ORDER BY upload_date DESC", (status,)
        )
    else:
        all_papers = query_db("SELECT * FROM papers ORDER BY upload_date DESC")
    return render_template("admin_papers.html", papers=all_papers, selected_status=status)


@app.route("/admin/students")
@admin_required
def admin_students():
    students = list(STUDENT_REGISTRY.items())
    return render_template("admin_students.html", students=students)


@app.route("/admin/upload", methods=["GET", "POST"])
@admin_required
def admin_upload():
    if request.method == "POST":
        subject = request.form.get("subject_name", "").strip()
        dept    = request.form.get("department", "").strip()
        year    = request.form.get("year", "").strip()
        file    = request.files.get("file")

        if not all([subject, dept, year, file]):
            flash("All fields required.", "danger")
            return render_template("admin_upload.html", departments=DEPARTMENTS, years=YEARS)
        if not file.filename.lower().endswith(".pdf"):
            flash("Only PDF allowed.", "danger")
            return render_template("admin_upload.html", departments=DEPARTMENTS, years=YEARS)

        filename = secure_filename(file.filename)
        name_without_ext = os.path.splitext(filename)[0]

        try:
            result = cloudinary.uploader.upload(
                file,
                resource_type="raw",
                folder="qpapers",
                public_id=name_without_ext,
                use_filename=True,
                unique_filename=True,
            )
            file_url  = result["secure_url"]
            public_id = result["public_id"]
        except Exception as e:
            flash(f"Upload failed: {str(e)}", "danger")
            return render_template("admin_upload.html", departments=DEPARTMENTS, years=YEARS)

        # FIX 20: ? → %s; removed explicit upload_date (use DB default)
        query_db(
            """INSERT INTO papers
               (subject_name, department, year, file_url, public_id, uploaded_by, status)
               VALUES (%s, %s, %s, %s, %s, 'ADMIN', 'approved')""",
            (subject, dept, year, file_url, public_id),
        )
        flash("Paper uploaded and approved!", "success")
        return redirect(url_for("admin_dashboard"))

    return render_template("admin_upload.html", departments=DEPARTMENTS, years=YEARS)


# ── Approve / Reject / Delete ─────────────────────────────────────────────────

@app.route("/approve/<int:paper_id>", methods=["POST"])
@admin_required
def approve_paper(paper_id):
    """
    FIX 21: Old code called get_db().execute() — psycopg2 connections have no
             .execute() method; only cursors do.  Use query_db() instead.
             Also: get_db().commit() after get_db().execute() called commit on
             a DIFFERENT cursor object in some edge cases — now query_db()
             handles commit atomically.
    """
    query_db("UPDATE papers SET status='approved' WHERE id=%s", (paper_id,))
    flash("Paper approved and visible to students.", "success")
    return redirect(request.referrer or url_for("admin_papers"))


@app.route("/reject/<int:paper_id>", methods=["POST"])
@admin_required
def reject_paper(paper_id):
    query_db("UPDATE papers SET status='rejected' WHERE id=%s", (paper_id,))
    flash("Paper rejected.", "warning")
    return redirect(request.referrer or url_for("admin_papers"))


@app.route("/delete/<int:paper_id>", methods=["POST"])
@admin_required
def delete_paper(paper_id):
    # FIX 22: ? → %s; removed .fetchone() chained on query_db()
    paper = query_db("SELECT * FROM papers WHERE id=%s", (paper_id,), one=True)
    if not paper:
        flash("Paper not found.", "danger")
        return redirect(url_for("admin_papers"))
    if paper["public_id"]:
        try:
            cloudinary.uploader.destroy(paper["public_id"], resource_type="raw")
        except Exception:
            pass
    query_db("DELETE FROM papers WHERE id=%s", (paper_id,))
    flash("Paper deleted.", "success")
    return redirect(request.referrer or url_for("admin_papers"))


@app.route("/admin/clear-activity", methods=["POST"])
@admin_required
def clear_activity():
    query_db("DELETE FROM activity_logs")
    flash("All activity logs cleared!", "success")
    return redirect(url_for("admin_dashboard"))


@app.route("/admin/activity")
@admin_required
def admin_activity_logs():
    action_filter = request.args.get("action", "")
    reg_filter    = request.args.get("student_reg", "").strip()

    # FIX 23: ? → %s in dynamic filter query
    sql     = "SELECT * FROM activity_logs"
    params  = []
    filters = []
    if action_filter in ("login", "upload", "download", "view"):
        filters.append("action = %s");            params.append(action_filter)
    if reg_filter:
        filters.append("student_reg LIKE %s");    params.append(f"%{reg_filter}%")
    if filters:
        sql += " WHERE " + " AND ".join(filters)
    sql += " ORDER BY timestamp DESC"

    # FIX 24: removed .fetchall() — query_db already returns the list
    logs = query_db(sql, params)
    return render_template(
        "admin_activity.html",
        logs=logs,
        selected_action=action_filter,
        search_reg=reg_filter,
    )


# ── API endpoints ─────────────────────────────────────────────────────────────

@app.route("/api/papers")
@student_required
def api_papers():
    # FIX 25: psycopg2 connection has no .execute(); use query_db()
    rows = query_db(
        "SELECT * FROM papers WHERE status='approved' ORDER BY upload_date DESC"
    )
    return jsonify([dict(r) for r in rows])


@app.route("/api/students")
@admin_required
def api_students():
    return jsonify([{"reg_no": reg, **data} for reg, data in STUDENT_REGISTRY.items()])


# ── Entrypoint ────────────────────────────────────────────────────────────────
load_excel_students()
init_db()

if __name__ == "__main__":
    app.run(debug=True)
