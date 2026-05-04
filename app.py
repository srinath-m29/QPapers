import os
import sqlite3
from datetime import datetime
from functools import wraps
from werkzeug.utils import secure_filename
from flask import Response
import requests
import cloudinary
import cloudinary.uploader
import cloudinary.utils
import pandas as pd
from dotenv import load_dotenv
from flask import (Flask, flash, g, jsonify, redirect, render_template, request, session, url_for)
import psycopg2

load_dotenv()

DATABASE_URL = os.getenv("DATABASE_URL")
USE_POSTGRES = DATABASE_URL is not None

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "dev-secret-key-change-me")

# ── Cloudinary config ──────────────────────────────────────────────────────────
cloudinary.config(
    cloud_name=os.getenv("CLOUDINARY_CLOUD_NAME"),
    api_key=os.getenv("CLOUDINARY_API_KEY"),
    api_secret=os.getenv("CLOUDINARY_API_SECRET"),
    secure=True,
)

# ── Force download route ─────────────────────────────────────────────────────
@app.route("/download/<int:paper_id>")
def force_download(paper_id):
    """
    Look up the paper in the DB, use the exact file_url Cloudinary returned
    at upload time (which already has the correct path + extension), and
    stream it back with Content-Disposition: attachment so the browser
    saves it as <subject_name>.pdf instead of showing a raw blob.
    """
    # ── Look up paper ────────────────────────────────────────────────────────
    os.makedirs(os.path.dirname(DATABASE), exist_ok=True)
    db = sqlite3.connect(DATABASE)
    db.row_factory = sqlite3.Row
    paper = query_db("SELECT * FROM papers WHERE id = ?", (paper_id,)).fetchone()
    db.close()

    if not paper:
        return "Paper not found.", 404

    file_url = paper["file_url"]      # exact URL Cloudinary gave us — always correct
    subject  = paper["subject_name"]  # use subject name for the saved filename

    # ── Activity Tracking Added ──────────────────────────────────────────────
    if session.get("student_reg"):
        log_activity(session["student_reg"], session.get("student_name", ""), "download", paper_id)
    # ────────────────────────────────────────────────────────────────────────

    # ── Build a clean filename ───────────────────────────────────────────────
    safe_name = subject.strip().replace(" ", "_").replace("/", "_")
    if not safe_name.lower().endswith(".pdf"):
        safe_name += ".pdf"

    # ── Stream the file from Cloudinary ─────────────────────────────────────
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
# ── Jinja2 filter: forced-download URL ────────────────────────────────────────
@app.template_filter("download_url")
def download_url_filter(file_url, filename="paper.pdf"):
    if not file_url:
        return ""

    filename = filename.replace(" ", "_").replace("/", "_")

    if not filename.lower().endswith(".pdf"):
        filename += ".pdf"

    return file_url.replace(
        "/raw/upload/",
        f"/raw/upload/fl_attachment:{filename}/"
    )

# ── Excel Student Registry ─────────────────────────────────────────────────────
EXCEL_FILE = os.path.join(os.getcwd(), "ECE.xlsx")
STUDENT_REGISTRY = {}  # reg_no -> {name, semester, status, section}

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

# ── Database ───────────────────────────────────────────────────────────────────
DATABASE = os.path.join(os.getcwd(), "instance", "qpapers.db")
MAX_FILE_SIZE = 10 * 1024 * 1024

DEPARTMENTS = [
    "Computer Science", "Information Technology",
    "Electronics & Communication", "Electrical Engineering",
    "Mechanical Engineering", "Civil Engineering",
    "Biotechnology", "Mathematics", "Physics", "Chemistry",
]
YEARS = ["1st Year", "2nd Year", "3rd Year", "4th Year"]

def get_db():
    if "db" not in g:
        if USE_POSTGRES:
            g.db = psycopg2.connect(DATABASE_URL)
        else:
            DATABASE = os.path.join(os.getcwd(), "instance", "qpapers.db")
            os.makedirs(os.path.dirname(DATABASE), exist_ok=True)
            g.db = sqlite3.connect(DATABASE)
            g.db.row_factory = sqlite3.Row
    return g.db

def query_db(query, params=(), one=False):
    db = get_db()

    if USE_POSTGRES:
        cur = db.cursor()
        cur.execute(query, params)

        if query.strip().lower().startswith("select"):
            columns = [desc[0] for desc in cur.description]
            rows = [dict(zip(columns, r)) for r in cur.fetchall()]
            return rows[0] if one and rows else (rows if not one else None)
        else:
            db.commit()
            return None
    else:
        cur = query_db(query, params)
        rows = cur.fetchall()
        return (rows[0] if rows else None) if one else rows

@app.teardown_appcontext
def close_db(error):
    db = g.pop("db", None)
    if db is not None:
        db.close()

def init_db():
    os.makedirs(os.path.dirname(DATABASE), exist_ok=True)
    db = sqlite3.connect(DATABASE)
    query_dbscript("""
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

        -- ── Activity Tracking Added ──────────────────────────────────────────
        CREATE TABLE IF NOT EXISTS activity_logs (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            student_reg  TEXT    NOT NULL,
            student_name TEXT    NOT NULL,
            action       TEXT    NOT NULL,
            paper_id     INTEGER,
            timestamp    TEXT    NOT NULL DEFAULT (datetime('now'))
        );
        -- ────────────────────────────────────────────────────────────────────
    """)
    # Safe migration for old schemas
    cols = [row[1] for row in query_db("PRAGMA table_info(papers)").fetchall()]
    if "status" not in cols:
        query_db("ALTER TABLE papers ADD COLUMN status TEXT NOT NULL DEFAULT 'pending'")
    if "public_id" not in cols:
        query_db("ALTER TABLE papers ADD COLUMN public_id TEXT NOT NULL DEFAULT ''")
    db.commit()
    db.close()
    print("Database initialised.")

# ── Activity Tracking Added ────────────────────────────────────────────────────
def log_activity(student_reg, student_name, action, paper_id=None):
    """Insert a row into activity_logs. Silently swallows errors so that a
    logging failure never disrupts normal request handling."""
    try:
        os.makedirs(os.path.dirname(DATABASE), exist_ok=True)
        db = sqlite3.connect(DATABASE)
        query_db(
            """INSERT INTO activity_logs (student_reg, student_name, action, paper_id, timestamp)
               VALUES (?, ?, ?, ?, ?)""",
            (student_reg, student_name, action, paper_id,
             datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        )
        db.commit()
        db.close()
    except Exception as exc:
        print(f"[activity_log] Failed to log '{action}' for {student_reg}: {exc}")
# ── End Activity Tracking ──────────────────────────────────────────────────────

# ── Auth decorators ────────────────────────────────────────────────────────────
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

# ── Root ───────────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return redirect(url_for("student_login"))

# ════════════════════════════════════════════════════════════════════════════════
#  STUDENT ROUTES
# ════════════════════════════════════════════════════════════════════════════════

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
        # ── Activity Tracking Added ──────────────────────────────────────────
        log_activity(reg, student_data["name"], "login")
        # ────────────────────────────────────────────────────────────────────
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
    db = get_db()
    recent = query_db(
        "SELECT * FROM papers WHERE status='approved' ORDER BY upload_date DESC LIMIT 6"
    ).fetchall()
    total = query_db("SELECT COUNT(*) as c FROM papers WHERE status='approved'").fetchone()["c"]
    depts = query_db("SELECT COUNT(DISTINCT department) as c FROM papers WHERE status='approved'").fetchone()["c"]
    my_pending = query_db(
        "SELECT COUNT(*) as c FROM papers WHERE uploaded_by=? AND status='pending'",
        (session["student_reg"],)
    ).fetchone()["c"]
    return render_template(
        "student_dashboard.html",
        recent=recent, total=total, depts=depts,
        my_pending=my_pending,
        departments=DEPARTMENTS, years=YEARS,
    )


@app.route("/papers")
@student_required
def papers():
    db      = get_db()
    dept    = request.args.get("department", "")
    year    = request.args.get("year", "")
    subject = request.args.get("subject", "").strip()

    query  = "SELECT * FROM papers WHERE status='approved'"
    params = []
    if dept:
        query += " AND department = ?"; params.append(dept)
    if year:
        query += " AND year = ?"; params.append(year)
    if subject:
        query += " AND subject_name LIKE ?"; params.append(f"%{subject}%")
    query += " ORDER BY upload_date DESC"

    all_papers = query_db(query, params).fetchall()
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

        db = get_db()
        query_db(
            """INSERT INTO papers
               (subject_name, department, year, file_url, public_id, uploaded_by, status, upload_date)
               VALUES (?, ?, ?, ?, ?, ?, 'pending', ?)""",
            (subject, dept, year, file_url, public_id,
             session["student_reg"], datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        )
        db.commit()
        # ── Activity Tracking Added ──────────────────────────────────────────
        log_activity(session["student_reg"], session["student_name"], "upload")
        # ────────────────────────────────────────────────────────────────────
        flash("Paper uploaded! It is pending admin approval.", "success")
        return redirect(url_for("student_dashboard"))

    return render_template("upload.html", departments=DEPARTMENTS, years=YEARS)

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

        db = get_db()
        query_db(
            """INSERT INTO papers
               (subject_name, department, year, file_url, public_id, uploaded_by, status, upload_date)
               VALUES (?, ?, ?, ?, ?, ?, 'approved', ?)""",
            (
                subject, dept, year, file_url, public_id,
                "ADMIN",   
                datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            ),
        )
        db.commit()

        flash("Paper uploaded and approved!", "success")
        return redirect(url_for("admin_dashboard"))

    return render_template("admin_upload.html", departments=DEPARTMENTS, years=YEARS)

# ── Activity Tracking Added ────────────────────────────────────────────────────
@app.route("/view/<int:paper_id>")
@student_required
def view_paper(paper_id):
    """Log a 'view' event then redirect to the paper's Cloudinary URL."""
    db    = get_db()
    paper = query_db("SELECT * FROM papers WHERE id = ? AND status = 'approved'",
                       (paper_id,)).fetchone()
    if not paper:
        flash("Paper not found or not yet approved.", "danger")
        return redirect(url_for("papers"))

    log_activity(session["student_reg"], session["student_name"], "view", paper_id)
    return redirect(paper["file_url"])
# ── End Activity Tracking ──────────────────────────────────────────────────────


# ════════════════════════════════════════════════════════════════════════════════
#  ADMIN ROUTES
# ════════════════════════════════════════════════════════════════════════════════

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
    db = get_db()
    total_papers   = query_db("SELECT COUNT(*) as c FROM papers").fetchone()["c"]
    pending_count  = query_db("SELECT COUNT(*) as c FROM papers WHERE status='pending'").fetchone()["c"]
    approved_count = query_db("SELECT COUNT(*) as c FROM papers WHERE status='approved'").fetchone()["c"]
    rejected_count = query_db("SELECT COUNT(*) as c FROM papers WHERE status='rejected'").fetchone()["c"]
    total_students = len(STUDENT_REGISTRY)
    recent_papers  = query_db("SELECT * FROM papers ORDER BY upload_date DESC LIMIT 10").fetchall()
    dept_stats     = query_db(
        "SELECT department, COUNT(*) as count FROM papers WHERE status='approved' "
        "GROUP BY department ORDER BY count DESC"
    ).fetchall()
    # --- Activity Tracking Added ---
    recent_activity = query_db(
        "SELECT * FROM activity_logs ORDER BY timestamp DESC LIMIT 10"
    ).fetchall()
# --------------------------------
    return render_template(
    "admin_dashboard.html",
    total_papers=total_papers, pending_count=pending_count,
    approved_count=approved_count, rejected_count=rejected_count,
    total_students=total_students,
    recent_papers=recent_papers, dept_stats=dept_stats,
    recent_activity=recent_activity  
    )


@app.route("/admin/papers")
@admin_required
def admin_papers():
    db     = get_db()
    status = request.args.get("status", "")
    if status in ("pending", "approved", "rejected"):
        all_papers = query_db(
            "SELECT * FROM papers WHERE status=? ORDER BY upload_date DESC", (status,)
        ).fetchall()
    else:
        all_papers = query_db("SELECT * FROM papers ORDER BY upload_date DESC").fetchall()
    return render_template("admin_papers.html", papers=all_papers, selected_status=status)


@app.route("/admin/students")
@admin_required
def admin_students():
    students = list(STUDENT_REGISTRY.items())
    return render_template("admin_students.html", students=students)


# ── Approve / Reject / Delete ──────────────────────────────────────────────────

@app.route("/approve/<int:paper_id>", methods=["POST"])
@admin_required
def approve_paper(paper_id):
    get_db().execute("UPDATE papers SET status='approved' WHERE id=?", (paper_id,))
    get_db().commit()
    flash("Paper approved and visible to students.", "success")
    return redirect(request.referrer or url_for("admin_papers"))


@app.route("/reject/<int:paper_id>", methods=["POST"])
@admin_required
def reject_paper(paper_id):
    get_db().execute("UPDATE papers SET status='rejected' WHERE id=?", (paper_id,))
    get_db().commit()
    flash("Paper rejected.", "warning")
    return redirect(request.referrer or url_for("admin_papers"))


@app.route("/delete/<int:paper_id>", methods=["POST"])
@admin_required
def delete_paper(paper_id):
    db    = get_db()
    paper = query_db("SELECT * FROM papers WHERE id=?", (paper_id,)).fetchone()
    if not paper:
        flash("Paper not found.", "danger")
        return redirect(url_for("admin_papers"))
    if paper["public_id"]:
        try:
            cloudinary.uploader.destroy(paper["public_id"], resource_type="raw")
        except Exception:
            pass
    query_db("DELETE FROM papers WHERE id=?", (paper_id,))
    db.commit()
    flash("Paper deleted.", "success")
    return redirect(request.referrer or url_for("admin_papers"))

@app.route("/admin/clear-activity", methods=["POST"])
@admin_required
def clear_activity():
    db = get_db()
    query_db("DELETE FROM activity_logs")
    db.commit()
    flash("All activity logs cleared!", "success")
    return redirect(url_for("admin_dashboard"))

# ── Activity Tracking Added ────────────────────────────────────────────────────
@app.route("/admin/activity")
@admin_required
def admin_activity_logs():
    """Admin view: display all student activity logs, newest first."""
    db = get_db()

    action_filter = request.args.get("action", "")
    reg_filter    = request.args.get("student_reg", "").strip()

    query  = "SELECT * FROM activity_logs"
    params = []
    filters = []
    if action_filter in ("login", "upload", "download", "view"):
        filters.append("action = ?"); params.append(action_filter)
    if reg_filter:
        filters.append("student_reg LIKE ?"); params.append(f"%{reg_filter}%")
    if filters:
        query += " WHERE " + " AND ".join(filters)
    query += " ORDER BY timestamp DESC"

    logs = query_db(query, params).fetchall()
    return render_template(
        "admin_activity.html",
        logs=logs,
        selected_action=action_filter,
        search_reg=reg_filter,
    )
# ── End Activity Tracking ──────────────────────────────────────────────────────

# ── API ────────────────────────────────────────────────────────────────────────
@app.route("/api/papers")
@student_required
def api_papers():
    rows = get_db().execute(
        "SELECT * FROM papers WHERE status='approved' ORDER BY upload_date DESC"
    ).fetchall()
    return jsonify([dict(r) for r in rows])

@app.route("/api/students")
@admin_required
def api_students():
    return jsonify([{"reg_no": reg, **data} for reg, data in STUDENT_REGISTRY.items()])


# ── Entrypoint ─────────────────────────────────────────────────────────────────
load_excel_students()
init_db()

if __name__ == "__main__":
    app.run(debug=True)
