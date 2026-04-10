import os
import json
import secrets
import sqlite3
from datetime import datetime
from functools import wraps

from dotenv import load_dotenv
from flask import (
    Flask,
    flash,
    g,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename

load_dotenv()

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", secrets.token_hex(32))

if os.path.exists("/data"):
    DATABASE = "/data/club13_gmp.db"
    TEMPLATE_DIR = "/data/doc_templates"
    ATTACHMENTS_DIR = "/data/attachments"
else:
    DATABASE = os.path.join(os.path.dirname(__file__), "club13_gmp.db")
    TEMPLATE_DIR = os.path.join(os.path.dirname(__file__), "doc_templates")
    ATTACHMENTS_DIR = os.path.join(os.path.dirname(__file__), "attachments")

os.makedirs(TEMPLATE_DIR, exist_ok=True)
os.makedirs(ATTACHMENTS_DIR, exist_ok=True)

# Auto-seed bundled templates into the persistent data directory
BUNDLED_TEMPLATES = os.path.join(os.path.dirname(__file__), "doc_templates")
if os.path.isdir(BUNDLED_TEMPLATES) and BUNDLED_TEMPLATES != TEMPLATE_DIR:
    import shutil
    for fname in os.listdir(BUNDLED_TEMPLATES):
        src = os.path.join(BUNDLED_TEMPLATES, fname)
        dst = os.path.join(TEMPLATE_DIR, fname)
        if os.path.isfile(src) and not os.path.exists(dst):
            shutil.copy2(src, dst)

ALLOWED_ATTACHMENT_EXT = {".pdf", ".png", ".jpg", ".jpeg", ".docx", ".doc"}
MAX_ATTACHMENT_SIZE = 10 * 1024 * 1024  # 10 MB


def _get_attachment_paths(spec_type, spec_id):
    """Return list of (original_name, full_path) for a spec's attachments."""
    db = get_db()
    rows = db.execute(
        "SELECT filename, original_name FROM spec_attachments WHERE spec_type = ? AND spec_id = ? ORDER BY created_at",
        (spec_type, spec_id),
    ).fetchall()
    paths = []
    for r in rows:
        fp = os.path.join(ATTACHMENTS_DIR, r["filename"])
        if os.path.exists(fp):
            paths.append((r["original_name"], fp))
    return paths


# ── Database helpers ────────────────────────────────────────────────
def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(DATABASE)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys = ON")
    return g.db


@app.teardown_appcontext
def close_db(exception):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def hash_password(password):
    return generate_password_hash(password, method="pbkdf2:sha256")


def init_db():
    db = sqlite3.connect(DATABASE)
    db.row_factory = sqlite3.Row
    db.execute("PRAGMA foreign_keys = ON")
    db.executescript(
        """
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            full_name TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'staff',
            active INTEGER NOT NULL DEFAULT 1,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );

        -- Packaging Specifications
        CREATE TABLE IF NOT EXISTS packaging_specs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            spec_number TEXT NOT NULL UNIQUE,
            material_name TEXT NOT NULL,
            material_code TEXT,
            supplier TEXT,
            description TEXT,
            status TEXT NOT NULL DEFAULT 'draft',
            revision TEXT NOT NULL DEFAULT '00',
            effective_date TEXT,
            created_by INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (created_by) REFERENCES users(id)
        );

        CREATE TABLE IF NOT EXISTS pk_test_parameters (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            spec_id INTEGER NOT NULL,
            parameter_name TEXT NOT NULL,
            test_method TEXT,
            acceptance_criteria TEXT,
            sort_order INTEGER DEFAULT 0,
            FOREIGN KEY (spec_id) REFERENCES packaging_specs(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS pk_test_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            spec_id INTEGER NOT NULL,
            lot_number TEXT NOT NULL,
            quantity_received TEXT,
            date_received TEXT,
            date_tested TEXT,
            tested_by TEXT,
            approved_by TEXT,
            status TEXT NOT NULL DEFAULT 'pending',
            comments TEXT,
            created_by INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (spec_id) REFERENCES packaging_specs(id) ON DELETE CASCADE,
            FOREIGN KEY (created_by) REFERENCES users(id)
        );

        CREATE TABLE IF NOT EXISTS pk_test_results (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            record_id INTEGER NOT NULL,
            parameter_id INTEGER NOT NULL,
            result_value TEXT,
            pass_fail TEXT DEFAULT 'pending',
            FOREIGN KEY (record_id) REFERENCES pk_test_records(id) ON DELETE CASCADE,
            FOREIGN KEY (parameter_id) REFERENCES pk_test_parameters(id) ON DELETE CASCADE
        );

        -- Raw Material Specifications
        CREATE TABLE IF NOT EXISTS raw_material_specs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            spec_number TEXT NOT NULL UNIQUE,
            material_name TEXT NOT NULL,
            material_code TEXT,
            supplier TEXT,
            cas_number TEXT,
            description TEXT,
            status TEXT NOT NULL DEFAULT 'draft',
            revision TEXT NOT NULL DEFAULT '00',
            effective_date TEXT,
            created_by INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (created_by) REFERENCES users(id)
        );

        CREATE TABLE IF NOT EXISTS rm_test_parameters (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            spec_id INTEGER NOT NULL,
            parameter_name TEXT NOT NULL,
            test_method TEXT,
            acceptance_criteria TEXT,
            parameter_type TEXT NOT NULL DEFAULT 'direct',
            sort_order INTEGER DEFAULT 0,
            FOREIGN KEY (spec_id) REFERENCES raw_material_specs(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS rm_test_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            spec_id INTEGER NOT NULL,
            lot_number TEXT NOT NULL,
            coa_reference TEXT,
            quantity_received TEXT,
            date_received TEXT,
            date_tested TEXT,
            tested_by TEXT,
            approved_by TEXT,
            status TEXT NOT NULL DEFAULT 'pending',
            comments TEXT,
            created_by INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (spec_id) REFERENCES raw_material_specs(id) ON DELETE CASCADE,
            FOREIGN KEY (created_by) REFERENCES users(id)
        );

        CREATE TABLE IF NOT EXISTS rm_test_results (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            record_id INTEGER NOT NULL,
            parameter_id INTEGER NOT NULL,
            result_value TEXT,
            pass_fail TEXT DEFAULT 'pending',
            FOREIGN KEY (record_id) REFERENCES rm_test_records(id) ON DELETE CASCADE,
            FOREIGN KEY (parameter_id) REFERENCES rm_test_parameters(id) ON DELETE CASCADE
        );

        -- Standard Operating Procedures
        CREATE TABLE IF NOT EXISTS sops (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sop_number TEXT NOT NULL,
            title TEXT NOT NULL,
            department TEXT,
            revision TEXT NOT NULL DEFAULT '00',
            effective_date TEXT,
            review_date TEXT,
            purpose TEXT,
            scope TEXT,
            responsibilities TEXT,
            definitions TEXT,
            procedure_text TEXT,
            equipment_materials TEXT,
            references_text TEXT,
            status TEXT NOT NULL DEFAULT 'draft',
            supersedes_id INTEGER,
            created_by INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (supersedes_id) REFERENCES sops(id),
            FOREIGN KEY (created_by) REFERENCES users(id)
        );

        CREATE TABLE IF NOT EXISTS sop_revision_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sop_id INTEGER NOT NULL,
            revision TEXT NOT NULL,
            date TEXT,
            description TEXT,
            approved_by TEXT,
            FOREIGN KEY (sop_id) REFERENCES sops(id) ON DELETE CASCADE
        );

        -- Spec attachments (3rd party COAs, vendor specs)
        CREATE TABLE IF NOT EXISTS spec_attachments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            spec_type TEXT NOT NULL,
            spec_id INTEGER NOT NULL,
            filename TEXT NOT NULL,
            original_name TEXT NOT NULL,
            file_type TEXT NOT NULL,
            uploaded_by INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (uploaded_by) REFERENCES users(id)
        );
        CREATE INDEX IF NOT EXISTS idx_attachments_spec ON spec_attachments(spec_type, spec_id);

        -- Document generation log
        CREATE TABLE IF NOT EXISTS document_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            doc_type TEXT NOT NULL,
            doc_id INTEGER NOT NULL,
            record_id INTEGER,
            action TEXT NOT NULL,
            filename TEXT,
            user_id INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (user_id) REFERENCES users(id)
        );

        CREATE INDEX IF NOT EXISTS idx_pk_params_spec ON pk_test_parameters(spec_id);
        CREATE INDEX IF NOT EXISTS idx_pk_records_spec ON pk_test_records(spec_id);
        CREATE INDEX IF NOT EXISTS idx_pk_results_record ON pk_test_results(record_id);
        CREATE INDEX IF NOT EXISTS idx_rm_params_spec ON rm_test_parameters(spec_id);
        CREATE INDEX IF NOT EXISTS idx_rm_records_spec ON rm_test_records(spec_id);
        CREATE INDEX IF NOT EXISTS idx_rm_results_record ON rm_test_results(record_id);
        CREATE INDEX IF NOT EXISTS idx_sop_number ON sops(sop_number);
        CREATE INDEX IF NOT EXISTS idx_sop_status ON sops(status);
        CREATE INDEX IF NOT EXISTS idx_doclog_type ON document_log(doc_type);
        """
    )

    existing_user = db.execute("SELECT id FROM users LIMIT 1").fetchone()
    if not existing_user:
        db.execute(
            "INSERT INTO users (username, password_hash, full_name, role) VALUES (?, ?, ?, ?)",
            ("admin", hash_password("club13admin"), "GMP Admin", "admin"),
        )
        db.execute(
            "INSERT INTO users (username, password_hash, full_name, role) VALUES (?, ?, ?, ?)",
            ("manager", hash_password("club13manager"), "Quality Manager", "manager"),
        )
        db.execute(
            "INSERT INTO users (username, password_hash, full_name, role) VALUES (?, ?, ?, ?)",
            ("staff", hash_password("club13staff"), "QA Staff", "staff"),
        )

    db.commit()
    db.close()


# ── Auth helpers ───────────────────────────────────────────────────
def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if "user_id" not in session:
            flash("Please log in to continue.", "warning")
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated


def require_role(*roles):
    def decorator(f):
        @wraps(f)
        def decorated(*args, **kwargs):
            if session.get("role") not in roles:
                flash("Access denied.", "danger")
                return redirect(url_for("dashboard"))
            return f(*args, **kwargs)
        return decorated
    return decorator


@app.context_processor
def inject_helpers():
    return {"now": datetime.now()}


# ── Auth routes ────────────────────────────────────────────────────
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        db = get_db()
        user = db.execute(
            "SELECT * FROM users WHERE username = ? AND active = 1", (username,)
        ).fetchone()
        if user and check_password_hash(user["password_hash"], password):
            session["user_id"] = user["id"]
            session["username"] = user["username"]
            session["full_name"] = user["full_name"]
            session["role"] = user["role"]
            flash(f"Welcome, {user['full_name']}!", "success")
            return redirect(url_for("dashboard"))
        flash("Invalid credentials.", "danger")
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    flash("Signed out.", "success")
    return redirect(url_for("login"))


# ── Dashboard ──────────────────────────────────────────────────────
@app.route("/")
@login_required
def dashboard():
    db = get_db()
    pk_count = db.execute("SELECT COUNT(*) c FROM packaging_specs WHERE status = 'active'").fetchone()["c"]
    rm_count = db.execute("SELECT COUNT(*) c FROM raw_material_specs WHERE status = 'active'").fetchone()["c"]
    sop_count = db.execute("SELECT COUNT(*) c FROM sops WHERE status = 'active'").fetchone()["c"]
    pk_draft = db.execute("SELECT COUNT(*) c FROM packaging_specs WHERE status = 'draft'").fetchone()["c"]
    rm_draft = db.execute("SELECT COUNT(*) c FROM raw_material_specs WHERE status = 'draft'").fetchone()["c"]
    sop_draft = db.execute("SELECT COUNT(*) c FROM sops WHERE status = 'draft'").fetchone()["c"]
    recent_pk = db.execute(
        "SELECT r.*, s.spec_number, s.material_name FROM pk_test_records r JOIN packaging_specs s ON r.spec_id = s.id ORDER BY r.created_at DESC LIMIT 5"
    ).fetchall()
    recent_rm = db.execute(
        "SELECT r.*, s.spec_number, s.material_name FROM rm_test_records r JOIN raw_material_specs s ON r.spec_id = s.id ORDER BY r.created_at DESC LIMIT 5"
    ).fetchall()
    recent_docs = db.execute(
        "SELECT d.*, u.full_name FROM document_log d LEFT JOIN users u ON d.user_id = u.id ORDER BY d.created_at DESC LIMIT 10"
    ).fetchall()

    # Template status
    pk_template = os.path.exists(os.path.join(TEMPLATE_DIR, "PK Specification Test Record Template.dotx"))
    rm_template = os.path.exists(os.path.join(TEMPLATE_DIR, "RM Specification Test Record Template.dotx"))
    sop_template = os.path.exists(os.path.join(TEMPLATE_DIR, "SOP Temp.dotx"))

    return render_template(
        "dashboard.html",
        pk_count=pk_count, rm_count=rm_count, sop_count=sop_count,
        pk_draft=pk_draft, rm_draft=rm_draft, sop_draft=sop_draft,
        recent_pk=recent_pk, recent_rm=recent_rm, recent_docs=recent_docs,
        pk_template=pk_template, rm_template=rm_template, sop_template=sop_template,
    )


# ══════════════════════════════════════════════════════════════════
# PACKAGING SPECIFICATIONS
# ══════════════════════════════════════════════════════════════════

@app.route("/packaging-specs")
@login_required
def pk_specs_list():
    db = get_db()
    status_filter = request.args.get("status", "")
    search = request.args.get("q", "")
    query = "SELECT s.*, u.full_name as creator_name FROM packaging_specs s LEFT JOIN users u ON s.created_by = u.id WHERE 1=1"
    params = []
    if status_filter:
        query += " AND s.status = ?"
        params.append(status_filter)
    if search:
        query += " AND (s.spec_number LIKE ? OR s.material_name LIKE ? OR s.supplier LIKE ?)"
        params.extend([f"%{search}%"] * 3)
    query += " ORDER BY s.spec_number"
    specs = db.execute(query, params).fetchall()
    return render_template("packaging_specs.html", specs=specs, status_filter=status_filter, search=search)


@app.route("/packaging-specs/new", methods=["GET", "POST"])
@login_required
def pk_spec_new():
    if request.method == "POST":
        return _save_pk_spec(None)
    return render_template("packaging_spec_form.html", spec=None, parameters=[])


@app.route("/packaging-specs/<int:spec_id>/edit", methods=["GET", "POST"])
@login_required
def pk_spec_edit(spec_id):
    db = get_db()
    spec = db.execute("SELECT * FROM packaging_specs WHERE id = ?", (spec_id,)).fetchone()
    if not spec:
        flash("Specification not found.", "danger")
        return redirect(url_for("pk_specs_list"))
    if request.method == "POST":
        return _save_pk_spec(spec_id)
    parameters = db.execute(
        "SELECT * FROM pk_test_parameters WHERE spec_id = ? ORDER BY sort_order", (spec_id,)
    ).fetchall()
    return render_template("packaging_spec_form.html", spec=spec, parameters=parameters)


def _save_pk_spec(spec_id):
    db = get_db()
    data = {
        "spec_number": request.form.get("spec_number", "").strip(),
        "material_name": request.form.get("material_name", "").strip(),
        "material_code": request.form.get("material_code", "").strip(),
        "supplier": request.form.get("supplier", "").strip(),
        "description": request.form.get("description", "").strip(),
        "status": request.form.get("status", "draft"),
        "revision": request.form.get("revision", "00").strip(),
        "effective_date": request.form.get("effective_date", "").strip(),
    }
    if not data["spec_number"] or not data["material_name"]:
        flash("Spec number and material name are required.", "danger")
        return redirect(request.url)

    if spec_id is None:
        existing = db.execute("SELECT id FROM packaging_specs WHERE spec_number = ?", (data["spec_number"],)).fetchone()
        if existing:
            flash("A spec with that number already exists.", "danger")
            return redirect(request.url)
        db.execute(
            """INSERT INTO packaging_specs (spec_number, material_name, material_code, supplier, description, status, revision, effective_date, created_by)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (data["spec_number"], data["material_name"], data["material_code"], data["supplier"],
             data["description"], data["status"], data["revision"], data["effective_date"], session["user_id"]),
        )
        spec_id = db.execute("SELECT last_insert_rowid()").fetchone()[0]
        flash("Packaging specification created.", "success")
    else:
        db.execute(
            """UPDATE packaging_specs SET spec_number=?, material_name=?, material_code=?, supplier=?,
               description=?, status=?, revision=?, effective_date=?, updated_at=CURRENT_TIMESTAMP WHERE id=?""",
            (data["spec_number"], data["material_name"], data["material_code"], data["supplier"],
             data["description"], data["status"], data["revision"], data["effective_date"], spec_id),
        )
        flash("Packaging specification updated.", "success")

    # Save test parameters
    db.execute("DELETE FROM pk_test_parameters WHERE spec_id = ?", (spec_id,))
    param_names = request.form.getlist("param_name[]")
    param_methods = request.form.getlist("param_method[]")
    param_criteria = request.form.getlist("param_criteria[]")
    for i, name in enumerate(param_names):
        if name.strip():
            method = param_methods[i] if i < len(param_methods) else ""
            criteria = param_criteria[i] if i < len(param_criteria) else ""
            db.execute(
                "INSERT INTO pk_test_parameters (spec_id, parameter_name, test_method, acceptance_criteria, sort_order) VALUES (?, ?, ?, ?, ?)",
                (spec_id, name.strip(), method.strip(), criteria.strip(), i),
            )
    db.commit()
    return redirect(url_for("pk_spec_detail", spec_id=spec_id))


@app.route("/packaging-specs/<int:spec_id>")
@login_required
def pk_spec_detail(spec_id):
    db = get_db()
    spec = db.execute(
        "SELECT s.*, u.full_name as creator_name FROM packaging_specs s LEFT JOIN users u ON s.created_by = u.id WHERE s.id = ?",
        (spec_id,),
    ).fetchone()
    if not spec:
        flash("Specification not found.", "danger")
        return redirect(url_for("pk_specs_list"))
    parameters = db.execute(
        "SELECT * FROM pk_test_parameters WHERE spec_id = ? ORDER BY sort_order", (spec_id,)
    ).fetchall()
    has_template = os.path.exists(os.path.join(TEMPLATE_DIR, "PK Specification Test Record Template.dotx"))
    has_crr_template = os.path.exists(os.path.join(TEMPLATE_DIR, "CLB003 Component Receiving Record.docx"))
    attachments = db.execute(
        "SELECT a.*, u.full_name as uploader_name FROM spec_attachments a LEFT JOIN users u ON a.uploaded_by = u.id WHERE a.spec_type = 'pk' AND a.spec_id = ? ORDER BY a.created_at",
        (spec_id,),
    ).fetchall()
    return render_template("packaging_spec_detail.html", spec=spec, parameters=parameters, has_template=has_template, has_crr_template=has_crr_template, attachments=attachments)


@app.route("/packaging-specs/<int:spec_id>/delete", methods=["POST"])
@login_required
@require_role("admin", "manager")
def pk_spec_delete(spec_id):
    db = get_db()
    db.execute("DELETE FROM packaging_specs WHERE id = ?", (spec_id,))
    db.commit()
    flash("Packaging specification deleted.", "success")
    return redirect(url_for("pk_specs_list"))


@app.route("/packaging-specs/<int:spec_id>/download")
@login_required
def pk_spec_download(spec_id):
    from doc_generator import generate_pk_spec_record
    db = get_db()
    spec = db.execute("SELECT * FROM packaging_specs WHERE id = ?", (spec_id,)).fetchone()
    if not spec:
        flash("Specification not found.", "danger")
        return redirect(url_for("pk_specs_list"))
    parameters = db.execute(
        "SELECT * FROM pk_test_parameters WHERE spec_id = ? ORDER BY sort_order", (spec_id,)
    ).fetchall()
    template_path = os.path.join(TEMPLATE_DIR, "PK Specification Test Record Template.dotx")
    if not os.path.exists(template_path):
        flash("PK template not uploaded. Go to Settings > Templates.", "danger")
        return redirect(url_for("pk_spec_detail", spec_id=spec_id))
    output_path = generate_pk_spec_record(template_path, spec, parameters,
                                           attachment_paths=_get_attachment_paths("pk", spec_id))
    filename = f"PK-{spec['spec_number']}-Test-Record.docx"
    db.execute(
        "INSERT INTO document_log (doc_type, doc_id, action, filename, user_id) VALUES (?, ?, ?, ?, ?)",
        ("pk_spec", spec_id, "download", filename, session["user_id"]),
    )
    db.commit()
    return send_file(output_path, as_attachment=True, download_name=filename)


@app.route("/packaging-specs/<int:spec_id>/receiving-record")
@login_required
def pk_receiving_record(spec_id):
    from doc_generator import generate_receiving_record
    db = get_db()
    spec = db.execute("SELECT * FROM packaging_specs WHERE id = ?", (spec_id,)).fetchone()
    if not spec:
        flash("Specification not found.", "danger")
        return redirect(url_for("pk_specs_list"))
    template_path = os.path.join(TEMPLATE_DIR, "CLB003 Component Receiving Record.docx")
    if not os.path.exists(template_path):
        flash("Receiving record template not uploaded. Go to Settings > Templates.", "danger")
        return redirect(url_for("pk_spec_detail", spec_id=spec_id))
    output_path = generate_receiving_record(template_path, spec, "pk")
    filename = f"PK-{spec['spec_number']}-Receiving-Record.docx"
    db.execute(
        "INSERT INTO document_log (doc_type, doc_id, action, filename, user_id) VALUES (?, ?, ?, ?, ?)",
        ("pk_receiving", spec_id, "download", filename, session["user_id"]),
    )
    db.commit()
    return send_file(output_path, as_attachment=True, download_name=filename)


# ══════════════════════════════════════════════════════════════════
# RAW MATERIAL SPECIFICATIONS
# ══════════════════════════════════════════════════════════════════

@app.route("/raw-material-specs")
@login_required
def rm_specs_list():
    db = get_db()
    status_filter = request.args.get("status", "")
    search = request.args.get("q", "")
    query = "SELECT s.*, u.full_name as creator_name FROM raw_material_specs s LEFT JOIN users u ON s.created_by = u.id WHERE 1=1"
    params = []
    if status_filter:
        query += " AND s.status = ?"
        params.append(status_filter)
    if search:
        query += " AND (s.spec_number LIKE ? OR s.material_name LIKE ? OR s.supplier LIKE ?)"
        params.extend([f"%{search}%"] * 3)
    query += " ORDER BY s.spec_number"
    specs = db.execute(query, params).fetchall()
    return render_template("raw_material_specs.html", specs=specs, status_filter=status_filter, search=search)


@app.route("/raw-material-specs/new", methods=["GET", "POST"])
@login_required
def rm_spec_new():
    if request.method == "POST":
        return _save_rm_spec(None)
    return render_template("raw_material_spec_form.html", spec=None, parameters=[])


@app.route("/raw-material-specs/<int:spec_id>/edit", methods=["GET", "POST"])
@login_required
def rm_spec_edit(spec_id):
    db = get_db()
    spec = db.execute("SELECT * FROM raw_material_specs WHERE id = ?", (spec_id,)).fetchone()
    if not spec:
        flash("Specification not found.", "danger")
        return redirect(url_for("rm_specs_list"))
    if request.method == "POST":
        return _save_rm_spec(spec_id)
    parameters = db.execute(
        "SELECT * FROM rm_test_parameters WHERE spec_id = ? ORDER BY sort_order", (spec_id,)
    ).fetchall()
    return render_template("raw_material_spec_form.html", spec=spec, parameters=parameters)


def _save_rm_spec(spec_id):
    db = get_db()
    data = {
        "spec_number": request.form.get("spec_number", "").strip(),
        "material_name": request.form.get("material_name", "").strip(),
        "material_code": request.form.get("material_code", "").strip(),
        "supplier": request.form.get("supplier", "").strip(),
        "cas_number": request.form.get("cas_number", "").strip(),
        "description": request.form.get("description", "").strip(),
        "status": request.form.get("status", "draft"),
        "revision": request.form.get("revision", "00").strip(),
        "effective_date": request.form.get("effective_date", "").strip(),
    }
    if not data["spec_number"] or not data["material_name"]:
        flash("Spec number and material name are required.", "danger")
        return redirect(request.url)

    if spec_id is None:
        existing = db.execute("SELECT id FROM raw_material_specs WHERE spec_number = ?", (data["spec_number"],)).fetchone()
        if existing:
            flash("A spec with that number already exists.", "danger")
            return redirect(request.url)
        db.execute(
            """INSERT INTO raw_material_specs (spec_number, material_name, material_code, supplier, cas_number, description, status, revision, effective_date, created_by)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (data["spec_number"], data["material_name"], data["material_code"], data["supplier"],
             data["cas_number"], data["description"], data["status"], data["revision"], data["effective_date"], session["user_id"]),
        )
        spec_id = db.execute("SELECT last_insert_rowid()").fetchone()[0]
        flash("Raw material specification created.", "success")
    else:
        db.execute(
            """UPDATE raw_material_specs SET spec_number=?, material_name=?, material_code=?, supplier=?, cas_number=?,
               description=?, status=?, revision=?, effective_date=?, updated_at=CURRENT_TIMESTAMP WHERE id=?""",
            (data["spec_number"], data["material_name"], data["material_code"], data["supplier"],
             data["cas_number"], data["description"], data["status"], data["revision"], data["effective_date"], spec_id),
        )
        flash("Raw material specification updated.", "success")

    # Save test parameters
    db.execute("DELETE FROM rm_test_parameters WHERE spec_id = ?", (spec_id,))
    param_names = request.form.getlist("param_name[]")
    param_methods = request.form.getlist("param_method[]")
    param_criteria = request.form.getlist("param_criteria[]")
    param_types = request.form.getlist("param_type[]")
    for i, name in enumerate(param_names):
        if name.strip():
            method = param_methods[i] if i < len(param_methods) else ""
            criteria = param_criteria[i] if i < len(param_criteria) else ""
            ptype = param_types[i] if i < len(param_types) else "direct"
            if ptype not in ("direct", "coa"):
                ptype = "direct"
            db.execute(
                "INSERT INTO rm_test_parameters (spec_id, parameter_name, test_method, acceptance_criteria, parameter_type, sort_order) VALUES (?, ?, ?, ?, ?, ?)",
                (spec_id, name.strip(), method.strip(), criteria.strip(), ptype, i),
            )
    db.commit()
    return redirect(url_for("rm_spec_detail", spec_id=spec_id))


@app.route("/raw-material-specs/<int:spec_id>")
@login_required
def rm_spec_detail(spec_id):
    db = get_db()
    spec = db.execute(
        "SELECT s.*, u.full_name as creator_name FROM raw_material_specs s LEFT JOIN users u ON s.created_by = u.id WHERE s.id = ?",
        (spec_id,),
    ).fetchone()
    if not spec:
        flash("Specification not found.", "danger")
        return redirect(url_for("rm_specs_list"))
    parameters = db.execute(
        "SELECT * FROM rm_test_parameters WHERE spec_id = ? ORDER BY sort_order", (spec_id,)
    ).fetchall()
    has_template = os.path.exists(os.path.join(TEMPLATE_DIR, "RM Specification Test Record Template.dotx"))
    has_crr_template = os.path.exists(os.path.join(TEMPLATE_DIR, "CLB003 Component Receiving Record.docx"))
    attachments = db.execute(
        "SELECT a.*, u.full_name as uploader_name FROM spec_attachments a LEFT JOIN users u ON a.uploaded_by = u.id WHERE a.spec_type = 'rm' AND a.spec_id = ? ORDER BY a.created_at",
        (spec_id,),
    ).fetchall()
    return render_template("raw_material_spec_detail.html", spec=spec, parameters=parameters, has_template=has_template, has_crr_template=has_crr_template, attachments=attachments)


@app.route("/raw-material-specs/<int:spec_id>/delete", methods=["POST"])
@login_required
@require_role("admin", "manager")
def rm_spec_delete(spec_id):
    db = get_db()
    db.execute("DELETE FROM raw_material_specs WHERE id = ?", (spec_id,))
    db.commit()
    flash("Raw material specification deleted.", "success")
    return redirect(url_for("rm_specs_list"))


@app.route("/raw-material-specs/<int:spec_id>/download")
@login_required
def rm_spec_download(spec_id):
    from doc_generator import generate_rm_spec_record
    db = get_db()
    spec = db.execute("SELECT * FROM raw_material_specs WHERE id = ?", (spec_id,)).fetchone()
    if not spec:
        flash("Specification not found.", "danger")
        return redirect(url_for("rm_specs_list"))
    parameters = db.execute(
        "SELECT * FROM rm_test_parameters WHERE spec_id = ? ORDER BY sort_order", (spec_id,)
    ).fetchall()
    direct_params = [p for p in parameters if p["parameter_type"] == "direct"]
    coa_params = [p for p in parameters if p["parameter_type"] == "coa"]
    template_path = os.path.join(TEMPLATE_DIR, "RM Specification Test Record Template.dotx")
    if not os.path.exists(template_path):
        flash("RM template not uploaded. Go to Settings > Templates.", "danger")
        return redirect(url_for("rm_spec_detail", spec_id=spec_id))
    output_path = generate_rm_spec_record(
        template_path, spec, parameters,
        direct_params=direct_params, coa_params=coa_params,
        attachment_paths=_get_attachment_paths("rm", spec_id),
    )
    filename = f"RM-{spec['spec_number']}-Test-Record.docx"
    db.execute(
        "INSERT INTO document_log (doc_type, doc_id, action, filename, user_id) VALUES (?, ?, ?, ?, ?)",
        ("rm_spec", spec_id, "download", filename, session["user_id"]),
    )
    db.commit()
    return send_file(output_path, as_attachment=True, download_name=filename)


@app.route("/raw-material-specs/<int:spec_id>/receiving-record")
@login_required
def rm_receiving_record(spec_id):
    from doc_generator import generate_receiving_record
    db = get_db()
    spec = db.execute("SELECT * FROM raw_material_specs WHERE id = ?", (spec_id,)).fetchone()
    if not spec:
        flash("Specification not found.", "danger")
        return redirect(url_for("rm_specs_list"))
    template_path = os.path.join(TEMPLATE_DIR, "CLB003 Component Receiving Record.docx")
    if not os.path.exists(template_path):
        flash("Receiving record template not uploaded. Go to Settings > Templates.", "danger")
        return redirect(url_for("rm_spec_detail", spec_id=spec_id))
    output_path = generate_receiving_record(template_path, spec, "rm")
    filename = f"RM-{spec['spec_number']}-Receiving-Record.docx"
    db.execute(
        "INSERT INTO document_log (doc_type, doc_id, action, filename, user_id) VALUES (?, ?, ?, ?, ?)",
        ("rm_receiving", spec_id, "download", filename, session["user_id"]),
    )
    db.commit()
    return send_file(output_path, as_attachment=True, download_name=filename)


# ══════════════════════════════════════════════════════════════════
# STANDARD OPERATING PROCEDURES
# ══════════════════════════════════════════════════════════════════

@app.route("/sops")
@login_required
def sops_list():
    db = get_db()
    status_filter = request.args.get("status", "")
    search = request.args.get("q", "")
    dept_filter = request.args.get("department", "")
    query = "SELECT s.*, u.full_name as creator_name FROM sops s LEFT JOIN users u ON s.created_by = u.id WHERE 1=1"
    params = []
    if status_filter:
        query += " AND s.status = ?"
        params.append(status_filter)
    if dept_filter:
        query += " AND s.department = ?"
        params.append(dept_filter)
    if search:
        query += " AND (s.sop_number LIKE ? OR s.title LIKE ?)"
        params.extend([f"%{search}%"] * 2)
    query += " ORDER BY s.sop_number"
    sops = db.execute(query, params).fetchall()
    departments = db.execute("SELECT DISTINCT department FROM sops WHERE department IS NOT NULL AND department != '' ORDER BY department").fetchall()
    return render_template("sops.html", sops=sops, status_filter=status_filter, search=search, dept_filter=dept_filter, departments=departments)


@app.route("/sops/new", methods=["GET", "POST"])
@login_required
def sop_new():
    if request.method == "POST":
        return _save_sop(None)
    return render_template("sop_form.html", sop=None, revision_history=[])


@app.route("/sops/<int:sop_id>/edit", methods=["GET", "POST"])
@login_required
def sop_edit(sop_id):
    db = get_db()
    sop = db.execute("SELECT * FROM sops WHERE id = ?", (sop_id,)).fetchone()
    if not sop:
        flash("SOP not found.", "danger")
        return redirect(url_for("sops_list"))
    if request.method == "POST":
        return _save_sop(sop_id)
    revision_history = db.execute(
        "SELECT * FROM sop_revision_history WHERE sop_id = ? ORDER BY id", (sop_id,)
    ).fetchall()
    return render_template("sop_form.html", sop=sop, revision_history=revision_history)


def _save_sop(sop_id):
    db = get_db()
    data = {
        "sop_number": request.form.get("sop_number", "").strip(),
        "title": request.form.get("title", "").strip(),
        "department": request.form.get("department", "").strip(),
        "revision": request.form.get("revision", "00").strip(),
        "effective_date": request.form.get("effective_date", "").strip(),
        "review_date": request.form.get("review_date", "").strip(),
        "purpose": request.form.get("purpose", "").strip(),
        "scope": request.form.get("scope", "").strip(),
        "responsibilities": request.form.get("responsibilities", "").strip(),
        "definitions": request.form.get("definitions", "").strip(),
        "procedure_text": request.form.get("procedure_text", "").strip(),
        "equipment_materials": request.form.get("equipment_materials", "").strip(),
        "references_text": request.form.get("references_text", "").strip(),
        "status": request.form.get("status", "draft"),
    }
    if not data["sop_number"] or not data["title"]:
        flash("SOP number and title are required.", "danger")
        return redirect(request.url)

    if sop_id is None:
        db.execute(
            """INSERT INTO sops (sop_number, title, department, revision, effective_date, review_date,
               purpose, scope, responsibilities, definitions, procedure_text, equipment_materials,
               references_text, status, created_by) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (data["sop_number"], data["title"], data["department"], data["revision"],
             data["effective_date"], data["review_date"], data["purpose"], data["scope"],
             data["responsibilities"], data["definitions"], data["procedure_text"],
             data["equipment_materials"], data["references_text"], data["status"], session["user_id"]),
        )
        sop_id = db.execute("SELECT last_insert_rowid()").fetchone()[0]
        flash("SOP created.", "success")
    else:
        db.execute(
            """UPDATE sops SET sop_number=?, title=?, department=?, revision=?, effective_date=?, review_date=?,
               purpose=?, scope=?, responsibilities=?, definitions=?, procedure_text=?, equipment_materials=?,
               references_text=?, status=?, updated_at=CURRENT_TIMESTAMP WHERE id=?""",
            (data["sop_number"], data["title"], data["department"], data["revision"],
             data["effective_date"], data["review_date"], data["purpose"], data["scope"],
             data["responsibilities"], data["definitions"], data["procedure_text"],
             data["equipment_materials"], data["references_text"], data["status"], sop_id),
        )
        flash("SOP updated.", "success")

    # Save revision history entries
    db.execute("DELETE FROM sop_revision_history WHERE sop_id = ?", (sop_id,))
    rev_nums = request.form.getlist("rev_num[]")
    rev_dates = request.form.getlist("rev_date[]")
    rev_descs = request.form.getlist("rev_desc[]")
    rev_approved = request.form.getlist("rev_approved[]")
    for i, rev in enumerate(rev_nums):
        if rev.strip():
            db.execute(
                "INSERT INTO sop_revision_history (sop_id, revision, date, description, approved_by) VALUES (?, ?, ?, ?, ?)",
                (sop_id, rev.strip(),
                 rev_dates[i].strip() if i < len(rev_dates) else "",
                 rev_descs[i].strip() if i < len(rev_descs) else "",
                 rev_approved[i].strip() if i < len(rev_approved) else ""),
            )
    db.commit()
    return redirect(url_for("sop_detail", sop_id=sop_id))


@app.route("/sops/<int:sop_id>")
@login_required
def sop_detail(sop_id):
    db = get_db()
    sop = db.execute(
        "SELECT s.*, u.full_name as creator_name FROM sops s LEFT JOIN users u ON s.created_by = u.id WHERE s.id = ?",
        (sop_id,),
    ).fetchone()
    if not sop:
        flash("SOP not found.", "danger")
        return redirect(url_for("sops_list"))
    revision_history = db.execute(
        "SELECT * FROM sop_revision_history WHERE sop_id = ? ORDER BY id", (sop_id,)
    ).fetchall()
    has_template = os.path.exists(os.path.join(TEMPLATE_DIR, "SOP Temp.dotx"))
    return render_template("sop_detail.html", sop=sop, revision_history=revision_history, has_template=has_template)


@app.route("/sops/<int:sop_id>/delete", methods=["POST"])
@login_required
@require_role("admin", "manager")
def sop_delete(sop_id):
    db = get_db()
    db.execute("DELETE FROM sops WHERE id = ?", (sop_id,))
    db.commit()
    flash("SOP deleted.", "success")
    return redirect(url_for("sops_list"))


@app.route("/sops/<int:sop_id>/revise", methods=["GET", "POST"])
@login_required
def sop_revise(sop_id):
    db = get_db()
    sop = db.execute("SELECT * FROM sops WHERE id = ?", (sop_id,)).fetchone()
    if not sop:
        flash("SOP not found.", "danger")
        return redirect(url_for("sops_list"))

    if request.method == "POST":
        # Mark old as superseded and create new revision
        db.execute("UPDATE sops SET status = 'superseded', updated_at = CURRENT_TIMESTAMP WHERE id = ?", (sop_id,))

        old_rev = sop["revision"] or "00"
        try:
            new_rev = str(int(old_rev) + 1).zfill(2)
        except ValueError:
            new_rev = old_rev + ".1"

        db.execute(
            """INSERT INTO sops (sop_number, title, department, revision, effective_date, review_date,
               purpose, scope, responsibilities, definitions, procedure_text, equipment_materials,
               references_text, status, supersedes_id, created_by)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'draft', ?, ?)""",
            (sop["sop_number"], sop["title"], sop["department"], new_rev,
             request.form.get("effective_date", ""), request.form.get("review_date", ""),
             sop["purpose"], sop["scope"], sop["responsibilities"], sop["definitions"],
             sop["procedure_text"], sop["equipment_materials"], sop["references_text"],
             sop_id, session["user_id"]),
        )
        new_id = db.execute("SELECT last_insert_rowid()").fetchone()[0]

        # Copy revision history and add new entry
        old_history = db.execute("SELECT * FROM sop_revision_history WHERE sop_id = ? ORDER BY id", (sop_id,)).fetchall()
        for h in old_history:
            db.execute(
                "INSERT INTO sop_revision_history (sop_id, revision, date, description, approved_by) VALUES (?, ?, ?, ?, ?)",
                (new_id, h["revision"], h["date"], h["description"], h["approved_by"]),
            )
        db.execute(
            "INSERT INTO sop_revision_history (sop_id, revision, date, description, approved_by) VALUES (?, ?, ?, ?, ?)",
            (new_id, new_rev, request.form.get("effective_date", ""),
             request.form.get("revision_description", ""), request.form.get("revision_approved_by", "")),
        )
        db.commit()
        flash(f"New revision {new_rev} created as draft.", "success")
        return redirect(url_for("sop_edit", sop_id=new_id))

    return render_template("sop_revise.html", sop=sop)


@app.route("/sops/<int:sop_id>/download")
@login_required
def sop_download(sop_id):
    from doc_generator import generate_sop
    db = get_db()
    sop = db.execute("SELECT * FROM sops WHERE id = ?", (sop_id,)).fetchone()
    if not sop:
        flash("SOP not found.", "danger")
        return redirect(url_for("sops_list"))
    revision_history = db.execute(
        "SELECT * FROM sop_revision_history WHERE sop_id = ? ORDER BY id", (sop_id,)
    ).fetchall()

    template_path = os.path.join(TEMPLATE_DIR, "SOP Temp.dotx")
    if not os.path.exists(template_path):
        flash("SOP template not uploaded. Go to Settings > Templates.", "danger")
        return redirect(url_for("sop_detail", sop_id=sop_id))

    output_path = generate_sop(template_path, sop, revision_history)
    filename = f"SOP-{sop['sop_number']}-Rev{sop['revision']}.docx"
    db.execute(
        "INSERT INTO document_log (doc_type, doc_id, action, filename, user_id) VALUES (?, ?, ?, ?, ?)",
        ("sop", sop_id, "download", filename, session["user_id"]),
    )
    db.commit()
    return send_file(output_path, as_attachment=True, download_name=filename)


# ══════════════════════════════════════════════════════════════════
# USERS (Admin)
# ══════════════════════════════════════════════════════════════════

@app.route("/users")
@login_required
@require_role("admin")
def users_list():
    db = get_db()
    users = db.execute("SELECT * FROM users ORDER BY full_name").fetchall()
    return render_template("users.html", users=users)


@app.route("/users/new", methods=["GET", "POST"])
@login_required
@require_role("admin")
def user_new():
    if request.method == "POST":
        return _save_user(None)
    return render_template("user_form.html", user=None)


@app.route("/users/<int:user_id>/edit", methods=["GET", "POST"])
@login_required
@require_role("admin")
def user_edit(user_id):
    db = get_db()
    user = db.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()
    if not user:
        flash("User not found.", "danger")
        return redirect(url_for("users_list"))
    if request.method == "POST":
        return _save_user(user_id)
    return render_template("user_form.html", user=user)


def _save_user(user_id):
    db = get_db()
    username = request.form.get("username", "").strip()
    full_name = request.form.get("full_name", "").strip()
    role = request.form.get("role", "staff")
    active = 1 if request.form.get("active") else 0
    password = request.form.get("password", "")

    if not username or not full_name:
        flash("Username and full name are required.", "danger")
        return redirect(request.url)

    if user_id is None:
        if not password:
            flash("Password is required for new users.", "danger")
            return redirect(request.url)
        existing = db.execute("SELECT id FROM users WHERE username = ?", (username,)).fetchone()
        if existing:
            flash("Username already taken.", "danger")
            return redirect(request.url)
        db.execute(
            "INSERT INTO users (username, password_hash, full_name, role, active) VALUES (?, ?, ?, ?, ?)",
            (username, hash_password(password), full_name, role, 1),
        )
        flash("User created.", "success")
    else:
        if password:
            db.execute(
                "UPDATE users SET username=?, password_hash=?, full_name=?, role=?, active=? WHERE id=?",
                (username, hash_password(password), full_name, role, active, user_id),
            )
        else:
            db.execute(
                "UPDATE users SET username=?, full_name=?, role=?, active=? WHERE id=?",
                (username, full_name, role, active, user_id),
            )
        flash("User updated.", "success")
    db.commit()
    return redirect(url_for("users_list"))


# ══════════════════════════════════════════════════════════════════
# RECEIVING RECORDS
# ══════════════════════════════════════════════════════════════════

@app.route("/receiving-records")
@login_required
def receiving_records():
    db = get_db()
    pk_specs = db.execute(
        "SELECT id, spec_number, material_name, material_code, supplier FROM packaging_specs ORDER BY spec_number"
    ).fetchall()
    rm_specs = db.execute(
        "SELECT id, spec_number, material_name, material_code, supplier FROM raw_material_specs ORDER BY spec_number"
    ).fetchall()
    has_crr_template = os.path.exists(os.path.join(TEMPLATE_DIR, "CLB003 Component Receiving Record.docx"))
    return render_template("receiving_records.html", pk_specs=pk_specs, rm_specs=rm_specs, has_crr_template=has_crr_template)


# ══════════════════════════════════════════════════════════════════
# TEMPLATE MANAGEMENT
# ══════════════════════════════════════════════════════════════════

ALLOWED_TEMPLATE_NAMES = {
    "pk_template": "PK Specification Test Record Template.dotx",
    "rm_template": "RM Specification Test Record Template.dotx",
    "sop_template": "SOP Temp.dotx",
    "crr_template": "CLB003 Component Receiving Record.docx",
}


@app.route("/settings/templates")
@login_required
@require_role("admin", "manager")
def template_settings():
    templates = {}
    for key, filename in ALLOWED_TEMPLATE_NAMES.items():
        path = os.path.join(TEMPLATE_DIR, filename)
        templates[key] = {
            "filename": filename,
            "exists": os.path.exists(path),
            "size": os.path.getsize(path) if os.path.exists(path) else 0,
        }
    return render_template("template_settings.html", templates=templates)


@app.route("/settings/templates/upload", methods=["POST"])
@login_required
@require_role("admin", "manager")
def template_upload():
    template_type = request.form.get("template_type", "")
    if template_type not in ALLOWED_TEMPLATE_NAMES:
        flash("Invalid template type.", "danger")
        return redirect(url_for("template_settings"))

    file = request.files.get("template_file")
    if not file or not file.filename:
        flash("No file selected.", "danger")
        return redirect(url_for("template_settings"))

    if not file.filename.lower().endswith((".dotx", ".docx")):
        flash("Only .dotx and .docx files are allowed.", "danger")
        return redirect(url_for("template_settings"))

    target_name = ALLOWED_TEMPLATE_NAMES[template_type]
    file.save(os.path.join(TEMPLATE_DIR, target_name))
    flash(f"Template '{target_name}' uploaded successfully.", "success")
    return redirect(url_for("template_settings"))


# ══════════════════════════════════════════════════════════════════
# SPEC ATTACHMENTS (3rd party COAs / Vendor Specifications)
# ══════════════════════════════════════════════════════════════════

@app.route("/attachments/<spec_type>/<int:spec_id>/upload", methods=["POST"])
@login_required
def attachment_upload(spec_type, spec_id):
    if spec_type not in ("pk", "rm"):
        flash("Invalid spec type.", "danger")
        return redirect(url_for("dashboard"))

    detail_url = url_for("pk_spec_detail" if spec_type == "pk" else "rm_spec_detail", spec_id=spec_id)

    file = request.files.get("attachment")
    if not file or file.filename == "":
        flash("No file selected.", "danger")
        return redirect(detail_url)

    original_name = secure_filename(file.filename)
    ext = os.path.splitext(original_name)[1].lower()
    if ext not in ALLOWED_ATTACHMENT_EXT:
        flash(f"File type {ext} not allowed. Allowed: {', '.join(sorted(ALLOWED_ATTACHMENT_EXT))}", "danger")
        return redirect(detail_url)

    # Read and check size
    file_data = file.read()
    if len(file_data) > MAX_ATTACHMENT_SIZE:
        flash("File too large. Maximum 10 MB.", "danger")
        return redirect(detail_url)

    # Save with unique name
    unique_name = f"{spec_type}_{spec_id}_{secrets.token_hex(8)}{ext}"
    save_path = os.path.join(ATTACHMENTS_DIR, unique_name)
    with open(save_path, "wb") as f:
        f.write(file_data)

    db = get_db()
    db.execute(
        "INSERT INTO spec_attachments (spec_type, spec_id, filename, original_name, file_type, uploaded_by) VALUES (?, ?, ?, ?, ?, ?)",
        (spec_type, spec_id, unique_name, original_name, ext, session["user_id"]),
    )
    db.commit()
    flash(f"Attachment '{original_name}' uploaded.", "success")
    return redirect(detail_url)


@app.route("/attachments/<int:att_id>/delete", methods=["POST"])
@login_required
def attachment_delete(att_id):
    db = get_db()
    att = db.execute("SELECT * FROM spec_attachments WHERE id = ?", (att_id,)).fetchone()
    if not att:
        flash("Attachment not found.", "danger")
        return redirect(url_for("dashboard"))

    detail_url = url_for(
        "pk_spec_detail" if att["spec_type"] == "pk" else "rm_spec_detail",
        spec_id=att["spec_id"],
    )

    # Remove file
    file_path = os.path.join(ATTACHMENTS_DIR, att["filename"])
    if os.path.exists(file_path):
        os.remove(file_path)

    db.execute("DELETE FROM spec_attachments WHERE id = ?", (att_id,))
    db.commit()
    flash("Attachment deleted.", "success")
    return redirect(detail_url)


@app.route("/attachments/<int:att_id>/view")
@login_required
def attachment_view(att_id):
    db = get_db()
    att = db.execute("SELECT * FROM spec_attachments WHERE id = ?", (att_id,)).fetchone()
    if not att:
        flash("Attachment not found.", "danger")
        return redirect(url_for("dashboard"))
    file_path = os.path.join(ATTACHMENTS_DIR, att["filename"])
    if not os.path.exists(file_path):
        flash("File not found on disk.", "danger")
        return redirect(url_for("dashboard"))
    return send_file(file_path, download_name=att["original_name"])


# ══════════════════════════════════════════════════════════════════
# DOCUMENT LOG
# ══════════════════════════════════════════════════════════════════

@app.route("/document-log")
@login_required
def document_log():
    db = get_db()
    logs = db.execute(
        "SELECT d.*, u.full_name FROM document_log d LEFT JOIN users u ON d.user_id = u.id ORDER BY d.created_at DESC LIMIT 200"
    ).fetchall()
    return render_template("document_log.html", logs=logs)


# ══════════════════════════════════════════════════════════════════
# INIT
# ══════════════════════════════════════════════════════════════════

init_db()

if __name__ == "__main__":
    app.run(debug=True, port=5003)
