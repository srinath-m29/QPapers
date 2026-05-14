

import os
import sys
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

# ── Load env ──────────────────────────────────────────────────────────────────
load_dotenv()

# ── DATABASE_URL at import time ───────────────────────────────────────────────
DATABASE_URL = os.getenv("DATABASE_URL")
if not DATABASE_URL:
    DATABASE_URL = "postgresql://postgres:password@localhost:5432/qpapers_db"
    print(
        "WARNING: DATABASE_URL not set. Using localhost fallback.\n"
        "         On Render, add DATABASE_URL in the Environment tab.",
        file=sys.stderr,
    )

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "dev-secret-key-change-me")

# ── Cloudinary ────────────────────────────────────────────────────────────────
cloudinary.config(
    cloud_name=os.getenv("CLOUDINARY_CLOUD_NAME"),
    api_key=os.getenv("CLOUDINARY_API_KEY"),
    api_secret=os.getenv("CLOUDINARY_API_SECRET"),
    secure=True,
)

MAX_FILE_SIZE = 10 * 1024 * 1024

DEPARTMENTS = [
    "Computer Science", "Information Technology",
    "Electronics & Communication", "Electrical & Electronics Engineering",
    "Mechanical Engineering", "Civil Engineering",
    "Bio Medical Engineering", "Computer Science & Engineering", "Fashion Technology",
]
YEARS = ["1st Year", "2nd Year", "3rd Year", "4th Year"]


# ══════════════════════════════════════════════════════════════════════════════
#  DATABASE LAYER
# ══════════════════════════════════════════════════════════════════════════════

def get_db():
    
    if "db" not in g:
        try:
            g.db = psycopg2.connect(DATABASE_URL)
        except psycopg2.OperationalError as exc:
            from flask import abort
            print(f"[DB] Could not connect: {exc}", file=sys.stderr)
            abort(503, description="Database temporarily unavailable.")
    return g.db


def _cursor(conn):
    """Open a RealDictCursor — rows behave like dicts, same as sqlite3.Row."""
    return conn.cursor(cursor_factory=RealDictCursor)


def query_db(query, params=(), one=False):
    
    conn = get_db()
    cur  = _cursor(conn)
    try:
        cur.execute(query, params if params else None)
    except Exception:
        conn.rollback()
        cur.close()
        raise

    if query.strip().upper().startswith("SELECT"):
        rows = cur.fetchall()
        cur.close()
        if one:
            return rows[0] if rows else None
        return rows
    else:
        conn.commit()
        cur.close()
        return None


@app.teardown_appcontext
def close_db(error=None):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def init_db():
    """
    Create tables if they don't exist.
    R1: Wrapped in try/except — startup DB errors are logged, not fatal.
    """
    try:
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
        print("PostgreSQL: tables ready.")

    except psycopg2.OperationalError as exc:
        print(
            f"[init_db] WARNING: Could not reach PostgreSQL at startup: {exc}\n"
            "          Tables will be created on first successful connection.",
            file=sys.stderr,
        )
    except Exception as exc:
        print(f"[init_db] Unexpected error: {exc}", file=sys.stderr)


# ── R3: Excel file path safe for all gunicorn modes ───────────────────────────
def _excel_path():
    try:
        base = os.path.dirname(os.path.abspath(__file__))
    except (NameError, TypeError):
        base = os.getcwd()
    return os.path.join(base, "ECE.xlsx")


EXCEL_FILE       = _excel_path()
STUDENT_REGISTRY: dict = {}


def load_excel_students():
    
    global STUDENT_REGISTRY
    if not os.path.exists(EXCEL_FILE):
        print(
            f"WARNING: ECE.xlsx not found at {EXCEL_FILE}.\n"
            "         Student login will be unavailable until the file is present.",
            file=sys.stderr,
        )
        return
    try:
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
    except Exception as exc:
        print(f"[load_excel_students] Failed to read Excel: {exc}", file=sys.stderr)


# ── Activity logging ───────────────────────────────────────────────────────────

def log_activity(student_reg, student_name, action, paper_id=None):
    """
    R7: Own short-lived connection; catches ALL exceptions.
    A log failure never disrupts normal request handling.
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
        print(f"[activity_log] Failed to log '{action}' for {student_reg}: {exc}",
              file=sys.stderr)


# ══════════════════════════════════════════════════════════════════════════════
#  AUTH DECORATORS
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

# ── R8: Health check ──────────────────────────────────────────────────────────
@app.route("/healthz")
def healthz():
    """
    Set Health Check Path to /healthz in Render service settings.
    Returns 200 when DB is reachable, 503 otherwise.
    """
    try:
        conn = psycopg2.connect(DATABASE_URL, connect_timeout=3)
        conn.close()
        return jsonify({"status": "ok"}), 200
    except Exception as exc:
        return jsonify({"status": "error", "detail": str(exc)}), 503


@app.route("/")
def index():
    return redirect(url_for("student_login"))


# ── Force download ────────────────────────────────────────────────────────────

@app.route("/download/<int:paper_id>")
def force_download(paper_id):
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
    paper_type = request.args.get("paper_type", "")
    sql    = "SELECT * FROM papers WHERE status='approved'"
    params = []
    if dept:
        sql += " AND department = %s";       params.append(dept)
    if year:
        sql += " AND year = %s";             params.append(year)
    if subject:
        sql += " AND subject_name LIKE %s";  params.append(f"%{subject}%")
    if paper_type:
        sql += " AND paper_type = %s";       params.append(paper_type)
    sql += " ORDER BY upload_date DESC"

    all_papers = query_db(sql, params)
    return render_template(
        "papers.html",
        papers=all_papers, departments=DEPARTMENTS, years=YEARS,
        selected_dept=dept, selected_year=year, search_subject=subject,
        selected_paper_type=paper_type
        )


@app.route("/upload", methods=["GET", "POST"])
@student_required
def upload():
    if request.method == "POST":
        subject = request.form.get("subject_name", "").strip()
        dept    = request.form.get("department", "").strip()
        year    = request.form.get("year", "").strip()
        paper_type = request.form.get("paper_type", "").strip()
        file    = request.files.get("file")

        if not all([subject, dept, year, paper_type, file]):
            flash("All fields and a file are required.", "danger")
            return render_template("upload.html", departments=DEPARTMENTS, years=YEARS)
        if not file.filename.lower().endswith(".pdf"):
            flash("Only PDF files are allowed.", "danger")
            return render_template("upload.html", departments=DEPARTMENTS, years=YEARS)

        file.seek(0, 2); size = file.tell(); file.seek(0)
        if size > MAX_FILE_SIZE:
            flash("File size must not exceed 10 MB.", "danger")
            return render_template("upload.html", departments=DEPARTMENTS, years=YEARS)

        filename         = secure_filename(file.filename)
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

        query_db(
            """INSERT INTO papers
               (subject_name, department, year, paper_type, file_url, public_id, uploaded_by, status)
                VALUES (%s, %s, %s, %s, %s, %s, %s, 'pending')""",
            (subject, dept, year, paper_type, file_url, public_id, session["student_reg"]),
        )
        log_activity(session["student_reg"], session["student_name"], "upload")
        flash("Paper uploaded! It is pending admin approval.", "success")
        return redirect(url_for("student_dashboard"))

    return render_template("upload.html", departments=DEPARTMENTS, years=YEARS)


@app.route("/view/<int:paper_id>")
@student_required
def view_paper(paper_id):
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
    total_papers   = query_db("SELECT COUNT(*) AS c FROM papers", one=True)["c"]
    pending_count  = query_db("SELECT COUNT(*) AS c FROM papers WHERE status='pending'",  one=True)["c"]
    approved_count = query_db("SELECT COUNT(*) AS c FROM papers WHERE status='approved'", one=True)["c"]
    rejected_count = query_db("SELECT COUNT(*) AS c FROM papers WHERE status='rejected'", one=True)["c"]
    total_students = len(STUDENT_REGISTRY)
    recent_papers  = query_db("SELECT * FROM papers ORDER BY upload_date DESC LIMIT 10")
    dept_stats     = query_db(
        "SELECT department, COUNT(*) AS count FROM papers WHERE status='approved' "
        "GROUP BY department ORDER BY count DESC"
    )
    recent_activity = query_db(
        "SELECT * FROM activity_logs ORDER BY timestamp DESC"
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
        paper_type = request.form.get("paper_type", "").strip()
        file    = request.files.get("file")

        if not all([subject, dept, year, paper_type, file]):
            flash("All fields required.", "danger")
            return render_template("admin_upload.html", departments=DEPARTMENTS, years=YEARS)
        if not file.filename.lower().endswith(".pdf"):
            flash("Only PDF allowed.", "danger")
            return render_template("admin_upload.html", departments=DEPARTMENTS, years=YEARS)

        filename         = secure_filename(file.filename)
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

        query_db(
            """INSERT INTO papers
               (subject_name, department, year, paper_type, file_url, public_id, uploaded_by, status)
               VALUES (%s, %s, %s, %s, %s, %s, 'ADMIN', 'approved')""",
            (subject, dept, year, paper_type, file_url, public_id),
        )
        flash("Paper uploaded and approved!", "success")
        return redirect(url_for("admin_dashboard"))

    return render_template("admin_upload.html", departments=DEPARTMENTS, years=YEARS)


@app.route("/approve/<int:paper_id>", methods=["POST"])
@admin_required
def approve_paper(paper_id):
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

    sql     = "SELECT * FROM activity_logs"
    params  = []
    filters = []
    if action_filter in ("login", "upload", "download", "view"):
        filters.append("action = %s");           params.append(action_filter)
    if reg_filter:
        filters.append("student_reg LIKE %s");   params.append(f"%{reg_filter}%")
    if filters:
        sql += " WHERE " + " AND ".join(filters)
    sql += " ORDER BY timestamp DESC"

    logs = query_db(sql, params)
    return render_template(
        "admin_activity.html",
        logs=logs,
        selected_action=action_filter,
        search_reg=reg_filter,
    )


# ── API ───────────────────────────────────────────────────────────────────────

@app.route("/api/papers")
@student_required
def api_papers():
    rows = query_db(
        "SELECT * FROM papers WHERE status='approved' ORDER BY upload_date DESC"
    )
    return jsonify([dict(r) for r in rows])


@app.route("/api/students")
@admin_required
def api_students():
    return jsonify([{"reg_no": reg, **data} for reg, data in STUDENT_REGISTRY.items()])


# ══════════════════════════════════════════════════════════════════════════════
#  STARTUP — both calls are exception-safe so gunicorn never exits status 1
# ══════════════════════════════════════════════════════════════════════════════

load_excel_students()
init_db()

if __name__ == "__main__":
    app.run(debug=True)
