import os
import sqlite3
import secrets
import hmac
from functools import wraps
from datetime import datetime
from pathlib import Path
import io

from flask import (
    Flask, request, redirect, url_for, session, render_template_string,
    flash, abort, send_file, jsonify
)
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "factory-inventory-secret-" + secrets.token_hex(16))

# Path configuration (supports Render persistent disk mount /var/data if available)
DB_PATH = os.environ.get("DATABASE_PATH", "inventory.db")
XLSX_PATH = os.environ.get("WORKBOOK_PATH", "Inventory management.xlsx")

CATEGORIES = [
    "01 General Fabrication",
    "02 Refrigeration",
    "03 Heating",
    "04 De-humidity",
    "05 Altitude",
    "06 Electrical & Control",
    "07 Miscellaneous"
]

def db():
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    return con

def init_db():
    # Ensure directory exists for DB_PATH if subfolder specified
    db_parent = Path(DB_PATH).parent
    if db_parent and not db_parent.exists():
        db_parent.mkdir(parents=True, exist_ok=True)

    con = db()
    con.executescript("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        role TEXT NOT NULL CHECK(role IN ('admin', 'viewer')),
        display_name TEXT
    );

    CREATE TABLE IF NOT EXISTS items (
        code TEXT PRIMARY KEY,
        description TEXT,
        specification TEXT,
        size_rating TEXT,
        standard TEXT,
        uom TEXT,
        opening REAL DEFAULT 0,
        min_level REAL DEFAULT 0,
        bin_location TEXT,
        supplier TEXT,
        remarks TEXT,
        category TEXT
    );

    CREATE TABLE IF NOT EXISTS transactions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        direction TEXT NOT NULL CHECK(direction IN ('IN', 'OUT')),
        txn_date TEXT NOT NULL,
        doc_no TEXT,
        item_code TEXT NOT NULL,
        quantity REAL NOT NULL CHECK(quantity > 0),
        job_no TEXT,
        party TEXT,
        entered_by TEXT,
        remarks TEXT,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(item_code) REFERENCES items(code)
    );
    """)
    con.commit()
    con.close()
    seed_users()
    import_workbook_if_empty()

def seed_users():
    """
    Seed 1 Admin and 3 Viewer accounts.
    Allows overriding usernames and passwords via environment variables.
    """
    default_accounts = [
        (os.environ.get("ADMIN_ID", "admin"), os.environ.get("ADMIN_PASSWORD", "admin123"), "admin", "System Administrator"),
        (os.environ.get("VIEWER2_ID", "viewer1"), os.environ.get("VIEWER2_PASSWORD", "view123"), "viewer", "Factory Viewer 1"),
        (os.environ.get("VIEWER3_ID", "viewer2"), os.environ.get("VIEWER3_PASSWORD", "view123"), "viewer", "Factory Viewer 2"),
        (os.environ.get("VIEWER4_ID", "viewer3"), os.environ.get("VIEWER4_PASSWORD", "view123"), "viewer", "Factory Viewer 3")
    ]
    con = db()
    for uid, pw, role, name in default_accounts:
        existing = con.execute("SELECT id FROM users WHERE username = ?", (uid,)).fetchone()
        if not existing:
            con.execute(
                "INSERT INTO users (username, password_hash, role, display_name) VALUES (?, ?, ?, ?)",
                (uid, generate_password_hash(pw), role, name)
            )
        else:
            # Update password hash if explicitly provided in environment
            if uid in [os.environ.get("ADMIN_ID"), os.environ.get("VIEWER2_ID"), os.environ.get("VIEWER3_ID"), os.environ.get("VIEWER4_ID")]:
                con.execute("UPDATE users SET password_hash = ? WHERE username = ?", (generate_password_hash(pw), uid))
    con.commit()
    con.close()

def import_workbook_if_empty():
    con = db()
    item_count = con.execute("SELECT COUNT(*) FROM items").fetchone()[0]
    con.close()
    if item_count > 0 or not Path(XLSX_PATH).exists():
        return

    try:
        wb = load_workbook(XLSX_PATH, data_only=True, read_only=True)
    except Exception as e:
        print("Could not load workbook:", e)
        return

    con = db()
    # Import Category items
    for sheet in CATEGORIES:
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        subcat = ""
        for row in ws.iter_rows(min_row=5, values_only=True):
            if not row or not row[0]:
                continue
            code = str(row[0]).strip()
            if not code.startswith("EC-"):
                continue
            desc = row[1]
            if not desc or not str(desc).strip():
                subcat = code
                continue

            vals = list(row) + [None] * max(0, 15 - len(row))
            try:
                opening_val = float(vals[6]) if vals[6] is not None else 0.0
            except (ValueError, TypeError):
                opening_val = 0.0
            try:
                min_val = float(vals[10]) if vals[10] is not None else 0.0
            except (ValueError, TypeError):
                min_val = 0.0

            con.execute("""
                INSERT OR IGNORE INTO items
                (code, description, specification, size_rating, standard, uom, opening, min_level, bin_location, supplier, remarks, category)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                code,
                str(desc).strip(),
                vals[2],
                vals[3],
                vals[4],
                vals[5],
                opening_val,
                min_val,
                vals[12],
                vals[13],
                vals[14],
                sheet
            ))

    # Import Stock Log
    if "Stock Log" in wb.sheetnames:
        ws_log = wb["Stock Log"]
        for row in ws_log.iter_rows(min_row=6, values_only=True):
            if not row or len(row) < 8:
                continue
            vals = list(row) + [None] * max(0, 12 - len(row))
            direction, date_val, doc_no, code, qty = vals[1], vals[2], vals[3], vals[4], vals[7]
            remarks = str(vals[11] or "")
            if direction not in ("IN", "OUT") or not code or not qty:
                continue
            if "example entry" in remarks.lower():
                continue
            code_str = str(code).strip()
            item_exists = con.execute("SELECT 1 FROM items WHERE code = ?", (code_str,)).fetchone()
            if not item_exists:
                continue
            try:
                qty_val = float(qty)
            except (ValueError, TypeError):
                continue
            if qty_val <= 0:
                continue

            if hasattr(date_val, "strftime"):
                dt_str = date_val.strftime("%Y-%m-%d")
            else:
                dt_str = str(date_val or datetime.now().strftime("%Y-%m-%d"))[:10]

            con.execute("""
                INSERT INTO transactions (direction, txn_date, doc_no, item_code, quantity, job_no, party, entered_by, remarks)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (direction, dt_str, doc_no, code_str, qty_val, vals[8], vals[9], vals[10] or "System Import", remarks))

    con.commit()
    con.close()
    wb.close()

def login_required(fn):
    @wraps(fn)
    def wrapped(*args, **kwargs):
        if not session.get("user"):
            return redirect(url_for("login"))
        return fn(*args, **kwargs)
    return wrapped

def admin_required(fn):
    @wraps(fn)
    @login_required
    def wrapped(*args, **kwargs):
        if session.get("role") != "admin":
            abort(403)
        return fn(*args, **kwargs)
    return wrapped

def csrf():
    token = session.get("csrf")
    if not token:
        token = secrets.token_urlsafe(32)
        session["csrf"] = token
    return token

@app.before_request
def check_csrf():
    if request.method == "POST":
        sent = request.form.get("csrf_token", "")
        stored = session.get("csrf", "")
        if not stored or not hmac.compare_digest(sent, stored):
            abort(400, "Invalid or missing CSRF token. Please refresh and try again.")

# Base HTML layout
BASE_TEMPLATE = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{{ title or 'Factory Inventory Portal' }}</title>
  <style>
    :root {
      --primary: #153b60;
      --primary-hover: #0e2a47;
      --accent: #0077cc;
      --success: #1b8754;
      --danger: #dc3545;
      --warning: #ffc107;
      --light-bg: #f4f6f9;
      --card-bg: #ffffff;
      --text: #212529;
      --text-muted: #6c757d;
      --border: #dee2e6;
    }
    * { box-sizing: border-box; }
    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
      background: var(--light-bg);
      color: var(--text);
      margin: 0;
      padding: 0;
      display: flex;
      flex-direction: column;
      min-height: 100vh;
    }
    header {
      background: var(--primary);
      color: white;
      padding: 14px 24px;
      display: flex;
      justify-content: space-between;
      align-items: center;
      flex-wrap: wrap;
      box-shadow: 0 2px 8px rgba(0,0,0,0.15);
    }
    .brand {
      display: flex;
      align-items: center;
      gap: 12px;
    }
    .brand h1 {
      margin: 0;
      font-size: 20px;
      font-weight: 700;
      letter-spacing: 0.5px;
    }
    .brand span {
      font-size: 13px;
      color: #b0c7de;
    }
    nav {
      display: flex;
      align-items: center;
      gap: 16px;
      flex-wrap: wrap;
    }
    nav a {
      color: #e2ecf5;
      text-decoration: none;
      font-size: 14px;
      font-weight: 500;
      padding: 6px 12px;
      border-radius: 6px;
      transition: background 0.2s;
    }
    nav a:hover, nav a.active {
      background: rgba(255,255,255,0.15);
      color: white;
    }
    .user-pill {
      display: flex;
      align-items: center;
      gap: 8px;
      background: rgba(0,0,0,0.25);
      padding: 6px 14px;
      border-radius: 20px;
      font-size: 13px;
    }
    .badge {
      display: inline-block;
      padding: 3px 8px;
      font-size: 11px;
      font-weight: 700;
      border-radius: 4px;
      text-transform: uppercase;
    }
    .badge-admin { background: #e02424; color: white; }
    .badge-viewer { background: #057a55; color: white; }
    .badge-ok { background: #def7ec; color: #03543f; }
    .badge-reorder { background: #fde8e8; color: #9b1c1c; font-weight: bold; }
    .badge-in { background: #def7ec; color: #03543f; }
    .badge-out { background: #feecdc; color: #b43403; }

    main {
      max-width: 1300px;
      width: 100%;
      margin: 24px auto;
      padding: 0 20px;
      flex: 1;
    }
    .card {
      background: var(--card-bg);
      border-radius: 10px;
      padding: 20px;
      margin-bottom: 20px;
      box-shadow: 0 1px 4px rgba(0,0,0,0.06);
      border: 1px solid var(--border);
    }
    .grid-stats {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
      gap: 16px;
      margin-bottom: 20px;
    }
    .stat-card {
      background: white;
      padding: 18px;
      border-radius: 10px;
      border: 1px solid var(--border);
      box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .stat-title {
      font-size: 13px;
      color: var(--text-muted);
      text-transform: uppercase;
      font-weight: 600;
      letter-spacing: 0.5px;
    }
    .stat-value {
      font-size: 28px;
      font-weight: 700;
      margin-top: 8px;
      color: var(--primary);
    }
    .stat-value.danger { color: var(--danger); }
    .stat-value.success { color: var(--success); }

    .btn {
      display: inline-block;
      font-weight: 600;
      text-align: center;
      padding: 9px 18px;
      font-size: 14px;
      border-radius: 6px;
      border: none;
      cursor: pointer;
      text-decoration: none;
      transition: all 0.2s;
    }
    .btn-primary { background: var(--accent); color: white; }
    .btn-primary:hover { background: #005fa3; }
    .btn-success { background: var(--success); color: white; }
    .btn-success:hover { background: #146c43; }
    .btn-outline {
      background: transparent;
      border: 1px solid var(--border);
      color: var(--text);
    }
    .btn-outline:hover { background: #f1f3f5; }
    .btn-sm { padding: 5px 10px; font-size: 12px; }

    input, select, textarea {
      width: 100%;
      padding: 9px 12px;
      font-size: 14px;
      border: 1px solid var(--border);
      border-radius: 6px;
      outline: none;
      background: white;
      color: var(--text);
    }
    input:focus, select:focus, textarea:focus {
      border-color: var(--accent);
      box-shadow: 0 0 0 3px rgba(0,119,204,0.15);
    }
    label {
      display: block;
      font-size: 13px;
      font-weight: 600;
      margin-bottom: 5px;
      color: #495057;
    }
    .form-group { margin-bottom: 15px; }
    .form-row {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
      gap: 15px;
    }

    table {
      width: 100%;
      border-collapse: collapse;
      font-size: 13.5px;
    }
    th, td {
      padding: 10px 14px;
      text-align: left;
      border-bottom: 1px solid var(--border);
    }
    th {
      background: #f8fafc;
      font-weight: 600;
      color: #475569;
      white-space: nowrap;
    }
    tr:hover td { background: #fbfcfd; }
    .table-container {
      overflow-x: auto;
      border: 1px solid var(--border);
      border-radius: 8px;
    }

    .alert {
      padding: 12px 18px;
      border-radius: 6px;
      margin-bottom: 18px;
      font-size: 14px;
    }
    .alert-success { background: #def7ec; color: #03543f; border: 1px solid #bcf0da; }
    .alert-danger { background: #fde8e8; color: #9b1c1c; border: 1px solid #f8b4b4; }
    .alert-info { background: #e1effe; color: #1e429f; border: 1px solid #b4c6fc; }

    footer {
      background: white;
      border-top: 1px solid var(--border);
      padding: 16px 24px;
      text-align: center;
      font-size: 13px;
      color: var(--text-muted);
      margin-top: 40px;
    }
  </style>
</head>
<body>
  <header>
    <div class="brand">
      <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"></path><polyline points="3.27 6.96 12 12.01 20.73 6.96"></polyline><line x1="12" y1="22.08" x2="12" y2="12"></line></svg>
      <div>
        <h1>Factory Inventory Portal</h1>
        <span>Environmental Chamber Manufacturing · Centralized System</span>
      </div>
    </div>
    {% if session.user %}
    <nav>
      <a href="{{ url_for('dashboard') }}" class="{% if active_tab == 'dashboard' %}active{% endif %}">Dashboard</a>
      <a href="{{ url_for('stock_log') }}" class="{% if active_tab == 'log' %}active{% endif %}">Stock Log</a>
      <a href="{{ url_for('users_info') }}" class="{% if active_tab == 'users' %}active{% endif %}">Access Accounts</a>
      <div class="user-pill">
        <span><strong>{{ session.user }}</strong></span>
        <span class="badge {% if session.role == 'admin' %}badge-admin{% else %}badge-viewer{% endif %}">{{ session.role }}</span>
        <a href="{{ url_for('logout') }}" style="color:#ffb3ba; margin-left:6px; padding:0;">Sign out</a>
      </div>
    </nav>
    {% endif %}
  </header>

  <main>
    {% for category, message in get_flashed_messages(with_categories=true) %}
      <div class="alert alert-{{ category if category in ['success', 'danger', 'info'] else 'info' }}">{{ message }}</div>
    {% endfor %}
    {{ body|safe }}
  </main>

  <footer>
    Centralized Factory Inventory &copy; 2026 · Role-based access: 1 Administrator (Edit) + 3 Viewers (Read-only) · Live Synchronized Database
  </footer>
</body>
</html>
"""

def render_page(body_template, active_tab="dashboard", title="Factory Inventory Portal", **context):
    full_ctx = {
        "csrf_token": csrf(),
        "active_tab": active_tab,
        "title": title,
        **context
    }
    rendered_body = render_template_string(body_template, **full_ctx)
    return render_template_string(BASE_TEMPLATE, body=rendered_body, **full_ctx)

@app.route("/")
@login_required
def dashboard():
    con = db()
    query = """
    SELECT
        i.*,
        COALESCE(SUM(CASE WHEN t.direction = 'IN' THEN t.quantity ELSE 0 END), 0) AS total_in,
        COALESCE(SUM(CASE WHEN t.direction = 'OUT' THEN t.quantity ELSE 0 END), 0) AS total_out,
        COALESCE(SUM(CASE WHEN t.direction = 'IN' THEN t.quantity WHEN t.direction = 'OUT' THEN -t.quantity ELSE 0 END), 0) AS net_movement
    FROM items i
    LEFT JOIN transactions t ON t.item_code = i.code
    GROUP BY i.code
    ORDER BY i.category, i.code
    """
    rows = con.execute(query).fetchall()

    # Category summaries
    cat_summary = {}
    for c in CATEGORIES:
        cat_summary[c] = {"name": c, "total": 0, "stocked": 0, "reorder": 0, "in_qty": 0, "out_qty": 0}

    items = []
    total_stocked = 0
    total_reorder = 0
    total_receipts_qty = 0
    total_issues_qty = 0

    for r in rows:
        d = dict(r)
        opening = float(d["opening"] or 0)
        movement = float(d["net_movement"] or 0)
        closing = opening + movement
        min_lvl = float(d["min_level"] or 0)
        status = "REORDER" if closing <= min_lvl else "OK"

        d["closing"] = closing
        d["status"] = status
        items.append(d)

        cat = d.get("category")
        if cat in cat_summary:
            cat_summary[cat]["total"] += 1
            if closing > 0:
                cat_summary[cat]["stocked"] += 1
            if status == "REORDER":
                cat_summary[cat]["reorder"] += 1
            cat_summary[cat]["in_qty"] += float(d["total_in"] or 0)
            cat_summary[cat]["out_qty"] += float(d["total_out"] or 0)

        if closing > 0:
            total_stocked += 1
        if status == "REORDER":
            total_reorder += 1
        total_receipts_qty += float(d["total_in"] or 0)
        total_issues_qty += float(d["total_out"] or 0)

    con.close()

    # Filters
    search_q = request.args.get("q", "").strip().lower()
    selected_cat = request.args.get("category", "").strip()
    selected_status = request.args.get("status", "").strip()

    filtered_items = []
    for item in items:
        if search_q:
            in_code = search_q in item["code"].lower()
            in_desc = search_q in (item["description"] or "").lower()
            in_spec = search_q in (item["specification"] or "").lower()
            in_part = search_q in (item["bin_location"] or "").lower()
            if not (in_code or in_desc or in_spec or in_part):
                continue
        if selected_cat and item["category"] != selected_cat:
            continue
        if selected_status and item["status"] != selected_status:
            continue
        filtered_items.append(item)

    body = """
    <!-- Top Metrics -->
    <div class="grid-stats">
      <div class="stat-card">
        <div class="stat-title">Total Items Cataloged</div>
        <div class="stat-value">{{ items|length }}</div>
      </div>
      <div class="stat-card">
        <div class="stat-title">Items In Stock</div>
        <div class="stat-value success">{{ total_stocked }}</div>
      </div>
      <div class="stat-card">
        <div class="stat-title">Items at REORDER Level</div>
        <div class="stat-value danger">{{ total_reorder }}</div>
      </div>
      <div class="stat-card">
        <div class="stat-title">Total Receipts (Qty)</div>
        <div class="stat-value">{{ "%.1f"|format(total_receipts_qty) }}</div>
      </div>
      <div class="stat-card">
        <div class="stat-title">Total Issues (Qty)</div>
        <div class="stat-value">{{ "%.1f"|format(total_issues_qty) }}</div>
      </div>
    </div>

    <!-- Category Summary Accordion/Table -->
    <div class="card">
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
        <h2 style="font-size:17px; margin:0; color:var(--primary);">Category Summary (Synced with Excel Summary Sheet)</h2>
        <div>
          <a href="{{ url_for('export_excel') }}" class="btn btn-outline btn-sm">Export to Excel (.xlsx)</a>
          <a href="{{ url_for('export_csv') }}" class="btn btn-outline btn-sm">Export CSV</a>
        </div>
      </div>
      <div class="table-container">
        <table>
          <thead>
            <tr>
              <th>Category</th>
              <th>Total Items</th>
              <th>Stocked Items</th>
              <th>Reorder Items</th>
              <th>Total Received (Qty)</th>
              <th>Total Issued (Qty)</th>
            </tr>
          </thead>
          <tbody>
            {% for cat in categories %}
            {% set cs = cat_summary[cat] %}
            <tr>
              <td><strong>{{ cs.name }}</strong></td>
              <td>{{ cs.total }}</td>
              <td>{{ cs.stocked }}</td>
              <td><span class="badge {% if cs.reorder > 0 %}badge-reorder{% else %}badge-ok{% endif %}">{{ cs.reorder }}</span></td>
              <td>{{ "%.1f"|format(cs.in_qty) }}</td>
              <td>{{ "%.1f"|format(cs.out_qty) }}</td>
            </tr>
            {% endfor %}
          </tbody>
        </table>
      </div>
    </div>

    {% if session.role == 'admin' %}
    <!-- Admin Transaction Panel -->
    <div class="card" style="border-left: 4px solid var(--accent);">
      <h2 style="font-size:17px; margin:0 0 14px 0; color:var(--primary);">
        Record Stock Movement (Admin Only)
      </h2>
      <form method="post" action="{{ url_for('add_transaction') }}">
        <input type="hidden" name="csrf_token" value="{{ csrf_token }}">
        <div class="form-row">
          <div class="form-group">
            <label>Movement Type *</label>
            <select name="direction" required>
              <option value="IN">IN (Receipt / GRN)</option>
              <option value="OUT">OUT (Issue to Job/Chamber)</option>
            </select>
          </div>
          <div class="form-group">
            <label>Date *</label>
            <input type="date" name="txn_date" value="{{ today }}" required>
          </div>
          <div class="form-group" style="grid-column: span 2;">
            <label>Item *</label>
            <select name="item_code" required>
              <option value="">-- Choose Item Code --</option>
              {% for itm in items %}
              <option value="{{ itm.code }}">[{{ itm.code }}] {{ itm.description }} (Avail: {{ "%.2f"|format(itm.closing) }} {{ itm.uom or '' }})</option>
              {% endfor %}
            </select>
          </div>
          <div class="form-group">
            <label>Quantity *</label>
            <input type="number" step="any" min="0.001" name="quantity" placeholder="e.g. 5.0" required>
          </div>
          <div class="form-group">
            <label>GRN / Issue Slip No.</label>
            <input type="text" name="doc_no" placeholder="e.g. GRN-002">
          </div>
          <div class="form-group">
            <label>Job / Chamber No.</label>
            <input type="text" name="job_no" placeholder="e.g. EC-JOB-101">
          </div>
          <div class="form-group">
            <label>Supplier / Issued To</label>
            <input type="text" name="party" placeholder="e.g. Fabrication Dept">
          </div>
          <div class="form-group" style="grid-column: span 2;">
            <label>Remarks</label>
            <input type="text" name="remarks" placeholder="Optional notes">
          </div>
        </div>
        <button type="submit" class="btn btn-primary" style="margin-top:5px;">Save Transaction & Update Stock</button>
      </form>
    </div>
    {% else %}
    <div class="alert alert-info" style="display:flex; justify-content:space-between; align-items:center;">
      <span>You are logged in with <strong>View-Only</strong> access ({{ session.user }}). Stock movements are managed by the Administrator.</span>
      <span class="badge badge-viewer">Read-Only</span>
    </div>
    {% endif %}

    <!-- Inventory Item Master Register -->
    <div class="card">
      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px; margin-bottom:14px;">
        <h2 style="font-size:17px; margin:0; color:var(--primary);">Inventory Items Master Register</h2>
        <span style="font-size:13px; color:var(--text-muted);">Showing {{ filtered_items|length }} of {{ items|length }} items</span>
      </div>

      <!-- Filters -->
      <form method="get" class="form-row" style="margin-bottom:16px;">
        <div class="form-group">
          <label>Search</label>
          <input type="text" name="q" value="{{ search_q }}" placeholder="Search Code, Description, Spec, Location...">
        </div>
        <div class="form-group">
          <label>Category</label>
          <select name="category">
            <option value="">All Categories (7)</option>
            {% for c in categories %}
            <option value="{{ c }}" {% if c == selected_cat %}selected{% endif %}>{{ c }}</option>
            {% endfor %}
          </select>
        </div>
        <div class="form-group">
          <label>Stock Status</label>
          <select name="status">
            <option value="">All Statuses</option>
            <option value="OK" {% if selected_status == 'OK' %}selected{% endif %}>OK (Above Min Level)</option>
            <option value="REORDER" {% if selected_status == 'REORDER' %}selected{% endif %}>REORDER (Low Stock)</option>
          </select>
        </div>
        <div class="form-group" style="display:flex; align-items:flex-end; gap:8px;">
          <button type="submit" class="btn btn-primary" style="flex:1;">Filter</button>
          <a href="{{ url_for('dashboard') }}" class="btn btn-outline">Reset</a>
        </div>
      </form>

      <!-- Items Table -->
      <div class="table-container">
        <table>
          <thead>
            <tr>
              <th>Item Code</th>
              <th>Description</th>
              <th>Category</th>
              <th>Specification</th>
              <th>UOM</th>
              <th>Opening</th>
              <th>In</th>
              <th>Out</th>
              <th>Closing Stock</th>
              <th>Min Level</th>
              <th>Status</th>
              <th>Location</th>
              <th>Supplier</th>
            </tr>
          </thead>
          <tbody>
            {% for itm in filtered_items %}
            <tr>
              <td><code><strong>{{ itm.code }}</strong></code></td>
              <td>{{ itm.description }}</td>
              <td><span style="font-size:12px; color:var(--text-muted);">{{ itm.category }}</span></td>
              <td><span style="font-size:12px;">{{ itm.specification or '-' }}</span></td>
              <td>{{ itm.uom or '-' }}</td>
              <td>{{ "%.1f"|format(itm.opening) }}</td>
              <td>{{ "%.1f"|format(itm.total_in) }}</td>
              <td>{{ "%.1f"|format(itm.total_out) }}</td>
              <td><strong>{{ "%.1f"|format(itm.closing) }}</strong></td>
              <td>{{ "%.1f"|format(itm.min_level) }}</td>
              <td>
                <span class="badge {% if itm.status == 'REORDER' %}badge-reorder{% else %}badge-ok{% endif %}">
                  {{ itm.status }}
                </span>
              </td>
              <td>{{ itm.bin_location or '-' }}</td>
              <td>{{ itm.supplier or '-' }}</td>
            </tr>
            {% else %}
            <tr>
              <td colspan="13" style="text-align:center; padding:24px; color:var(--text-muted);">
                No inventory items found matching your filter criteria.
              </td>
            </tr>
            {% endfor %}
          </tbody>
        </table>
      </div>
    </div>
    """

    return render_page(
        body,
        active_tab="dashboard",
        items=items,
        filtered_items=filtered_items,
        categories=CATEGORIES,
        cat_summary=cat_summary,
        total_stocked=total_stocked,
        total_reorder=total_reorder,
        total_receipts_qty=total_receipts_qty,
        total_issues_qty=total_issues_qty,
        search_q=search_q,
        selected_cat=selected_cat,
        selected_status=selected_status,
        today=datetime.now().strftime("%Y-%m-%d")
    )

@app.route("/stock-log")
@login_required
def stock_log():
    con = db()
    txns = con.execute("""
        SELECT t.*, i.description, i.uom, i.category
        FROM transactions t
        LEFT JOIN items i ON i.code = t.item_code
        ORDER BY t.txn_date DESC, t.id DESC
    """).fetchall()
    con.close()

    body = """
    <div class="card">
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:14px;">
        <div>
          <h2 style="font-size:18px; margin:0; color:var(--primary);">Stock Movement Audit Log (Receipts & Issues)</h2>
          <span style="font-size:13px; color:var(--text-muted);">Chronological history of all stock entries</span>
        </div>
        <div>
          <a href="{{ url_for('export_excel') }}" class="btn btn-outline btn-sm">Export Excel</a>
        </div>
      </div>

      <div class="table-container">
        <table>
          <thead>
            <tr>
              <th>ID</th>
              <th>Direction</th>
              <th>Date</th>
              <th>Doc / GRN No.</th>
              <th>Item Code</th>
              <th>Description</th>
              <th>Quantity</th>
              <th>UOM</th>
              <th>Job / Chamber No.</th>
              <th>Party / Issued To</th>
              <th>Entered By</th>
              <th>Remarks</th>
            </tr>
          </thead>
          <tbody>
            {% for tx in txns %}
            <tr>
              <td>#{{ tx.id }}</td>
              <td>
                <span class="badge {% if tx.direction == 'IN' %}badge-in{% else %}badge-out{% endif %}">
                  {{ tx.direction }}
                </span>
              </td>
              <td>{{ tx.txn_date }}</td>
              <td>{{ tx.doc_no or '-' }}</td>
              <td><code><strong>{{ tx.item_code }}</strong></code></td>
              <td>{{ tx.description or '-' }}</td>
              <td><strong>{{ "%.2f"|format(tx.quantity) }}</strong></td>
              <td>{{ tx.uom or '-' }}</td>
              <td>{{ tx.job_no or '-' }}</td>
              <td>{{ tx.party or '-' }}</td>
              <td>{{ tx.entered_by or '-' }}</td>
              <td>{{ tx.remarks or '-' }}</td>
            </tr>
            {% else %}
            <tr>
              <td colspan="12" style="text-align:center; padding:24px; color:var(--text-muted);">
                No stock transactions logged yet.
              </td>
            </tr>
            {% endfor %}
          </tbody>
        </table>
      </div>
    </div>
    """
    return render_page(body, active_tab="log", title="Stock Movement Log - Factory Inventory", txns=txns)

@app.route("/transactions/add", methods=["POST"])
@admin_required
def add_transaction():
    direction = request.form.get("direction")
    code = request.form.get("item_code", "").strip()
    date_str = request.form.get("txn_date", "").strip()
    doc_no = request.form.get("doc_no", "").strip()
    job_no = request.form.get("job_no", "").strip()
    party = request.form.get("party", "").strip()
    remarks = request.form.get("remarks", "").strip()

    if direction not in ("IN", "OUT"):
        flash("Invalid direction selected.", "danger")
        return redirect(url_for("dashboard"))

    try:
        qty = float(request.form.get("quantity", 0))
    except (ValueError, TypeError):
        flash("Quantity must be a valid positive number.", "danger")
        return redirect(url_for("dashboard"))

    if qty <= 0:
        flash("Quantity must be greater than zero.", "danger")
        return redirect(url_for("dashboard"))

    con = db()
    item = con.execute("SELECT code, opening FROM items WHERE code = ?", (code,)).fetchone()
    if not item:
        con.close()
        flash("Item code not found in inventory catalog.", "danger")
        return redirect(url_for("dashboard"))

    # Stock validation for OUT movement
    if direction == "OUT":
        curr_row = con.execute("""
            SELECT opening + COALESCE(SUM(CASE WHEN direction='IN' THEN quantity WHEN direction='OUT' THEN -quantity ELSE 0 END), 0) AS avail
            FROM items i
            LEFT JOIN transactions t ON t.item_code = i.code
            WHERE i.code = ?
            GROUP BY i.code
        """, (code,)).fetchone()
        avail = curr_row["avail"] if curr_row else 0
        if qty > avail:
            con.close()
            flash(f"Issue rejected: Requested {qty:g} units exceeds available stock of {avail:g} units.", "danger")
            return redirect(url_for("dashboard"))

    con.execute("""
        INSERT INTO transactions (direction, txn_date, doc_no, item_code, quantity, job_no, party, entered_by, remarks)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (direction, date_str or datetime.now().strftime("%Y-%m-%d"), doc_no, code, qty, job_no, party, session.get("user"), remarks))
    con.commit()
    con.close()

    flash(f"Successfully recorded {direction} transaction for item {code}. Live stock updated.", "success")
    return redirect(url_for("dashboard"))

@app.route("/users-info")
@login_required
def users_info():
    con = db()
    all_users = con.execute("SELECT id, username, role, display_name FROM users ORDER BY id").fetchall()
    con.close()

    body = """
    <div class="card" style="max-width: 800px; margin: 0 auto;">
      <h2 style="font-size:18px; margin:0 0 10px 0; color:var(--primary);">Configured Access Accounts</h2>
      <p style="font-size:14px; color:var(--text-muted); margin-bottom:20px;">
        The system is configured for 4 authorized people (1 Administrator with edit/movement permissions, and 3 Viewers with view-only access).
      </p>

      <div class="table-container">
        <table>
          <thead>
            <tr>
              <th>ID</th>
              <th>Username</th>
              <th>Display Name</th>
              <th>Role</th>
              <th>Permissions</th>
            </tr>
          </thead>
          <tbody>
            {% for u in all_users %}
            <tr>
              <td>{{ u.id }}</td>
              <td><strong>{{ u.username }}</strong></td>
              <td>{{ u.display_name or '-' }}</td>
              <td>
                <span class="badge {% if u.role == 'admin' %}badge-admin{% else %}badge-viewer{% endif %}">
                  {{ u.role }}
                </span>
              </td>
              <td>
                {% if u.role == 'admin' %}
                <span style="color:var(--success); font-weight:600;">Full Access (Record IN/OUT, manage stock, export)</span>
                {% else %}
                <span style="color:var(--text-muted);">View-Only (Live dashboard, stock log, search, export)</span>
                {% endif %}
              </td>
            </tr>
            {% endfor %}
          </tbody>
        </table>
      </div>

      <div style="background:#f8fafc; border:1px solid var(--border); border-radius:8px; padding:16px; margin-top:20px; font-size:13px;">
        <h4 style="margin:0 0 8px 0; color:var(--primary);">Credential Management & Deployment Security:</h4>
        <ul style="margin:0; padding-left:18px; line-height:1.6; color:#475569;">
          <li>All passwords are encrypted with strong cryptographic salts (PBKDF2/SHA256).</li>
          <li>Default login for testing: <code>admin</code> / <code>admin123</code>, <code>viewer1</code> / <code>view123</code>, <code>viewer2</code> / <code>view123</code>, <code>viewer3</code> / <code>view123</code>.</li>
          <li>For online production deployment on Render, customize credentials using Environment Variables (<code>ADMIN_ID</code>, <code>ADMIN_PASSWORD</code>, <code>VIEWER2_ID</code>, etc.).</li>
        </ul>
      </div>
    </div>
    """
    return render_page(body, active_tab="users", title="Access Accounts - Factory Inventory", all_users=all_users)

@app.route("/login", methods=["GET", "POST"])
def login():
    if session.get("user"):
        return redirect(url_for("dashboard"))

    if request.method == "POST":
        uname = request.form.get("username", "").strip()
        pw = request.form.get("password", "")

        con = db()
        user = con.execute("SELECT * FROM users WHERE username = ?", (uname,)).fetchone()
        con.close()

        if user and check_password_hash(user["password_hash"], pw):
            session.clear()
            session["user"] = user["username"]
            session["role"] = user["role"]
            session["display_name"] = user["display_name"]
            session["csrf"] = secrets.token_urlsafe(32)
            flash(f"Welcome back, {user['display_name'] or user['username']}!", "success")
            return redirect(url_for("dashboard"))
        else:
            flash("Invalid User ID or Password. Please try again.", "danger")

    body = """
    <div style="max-width: 440px; margin: 40px auto;">
      <div class="card" style="box-shadow: 0 4px 14px rgba(0,0,0,0.08);">
        <div style="text-align:center; margin-bottom:20px;">
          <h2 style="margin:0; color:var(--primary); font-size:22px;">Factory Inventory Portal</h2>
          <span style="font-size:13px; color:var(--text-muted);">Sign in to access centralized inventory</span>
        </div>

        <form method="post">
          <input type="hidden" name="csrf_token" value="{{ csrf_token }}">
          <div class="form-group">
            <label>User ID</label>
            <input type="text" name="username" placeholder="e.g. admin or viewer1" required autofocus autocomplete="username">
          </div>
          <div class="form-group">
            <label>Password</label>
            <input type="password" name="password" placeholder="Enter your password" required autocomplete="current-password">
          </div>
          <button type="submit" class="btn btn-primary" style="width:100%; margin-top:8px;">Sign In</button>
        </form>

        <div style="margin-top:20px; padding-top:16px; border-top:1px solid var(--border); font-size:12px; color:var(--text-muted);">
          <strong>Configured User IDs:</strong>
          <div style="margin-top:6px; display:grid; grid-template-columns:1fr 1fr; gap:6px;">
            <div>&bull; <code>admin</code> (Admin)</div>
            <div>&bull; <code>viewer1</code> (View-only)</div>
            <div>&bull; <code>viewer2</code> (View-only)</div>
            <div>&bull; <code>viewer3</code> (View-only)</div>
          </div>
        </div>
      </div>
    </div>
    """
    return render_page(body, active_tab="login", title="Sign In - Factory Inventory Portal")

@app.route("/logout")
def logout():
    session.clear()
    flash("You have successfully signed out.", "info")
    return redirect(url_for("login"))

@app.route("/export/excel")
@login_required
def export_excel():
    con = db()
    items = con.execute("""
        SELECT
            i.*,
            COALESCE(SUM(CASE WHEN t.direction = 'IN' THEN t.quantity ELSE 0 END), 0) AS total_in,
            COALESCE(SUM(CASE WHEN t.direction = 'OUT' THEN t.quantity ELSE 0 END), 0) AS total_out,
            COALESCE(SUM(CASE WHEN t.direction = 'IN' THEN t.quantity WHEN t.direction = 'OUT' THEN -t.quantity ELSE 0 END), 0) AS net_movement
        FROM items i
        LEFT JOIN transactions t ON t.item_code = i.code
        GROUP BY i.code
        ORDER BY i.category, i.code
    """).fetchall()

    wb = Workbook()
    ws = wb.active
    ws.title = "Live Inventory"

    headers = [
        "Item Code", "Item Description", "Category", "Specification",
        "Size / Rating", "Standard", "UOM", "Opening Stock",
        "Total Received (IN)", "Total Issued (OUT)", "Closing Stock",
        "Minimum Level", "Status", "Bin / Location", "Preferred Supplier"
    ]
    ws.append(headers)

    header_fill = PatternFill(start_color="153B60", end_color="153B60", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font

    for item in items:
        opening = float(item["opening"] or 0)
        movement = float(item["net_movement"] or 0)
        closing = opening + movement
        min_lvl = float(item["min_level"] or 0)
        status = "REORDER" if closing <= min_lvl else "OK"

        ws.append([
            item["code"], item["description"], item["category"], item["specification"],
            item["size_rating"], item["standard"], item["uom"], opening,
            float(item["total_in"] or 0), float(item["total_out"] or 0), closing,
            min_lvl, status, item["bin_location"], item["supplier"]
        ])

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = col[0].column_letter
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 40)

    # Second sheet: Stock Log
    ws_log = wb.create_sheet(title="Stock Log")
    log_headers = ["ID", "Direction", "Date", "Doc No", "Item Code", "Quantity", "Job No", "Party", "Entered By", "Remarks"]
    ws_log.append(log_headers)
    for col_idx in range(1, len(log_headers) + 1):
        cell = ws_log.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font

    logs = con.execute("SELECT * FROM transactions ORDER BY txn_date DESC, id DESC").fetchall()
    for l in logs:
        ws_log.append([
            l["id"], l["direction"], l["txn_date"], l["doc_no"], l["item_code"],
            float(l["quantity"]), l["job_no"], l["party"], l["entered_by"], l["remarks"]
        ])

    con.close()

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"Inventory_Export_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    return send_file(output, download_name=filename, as_attachment=True, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/export/csv")
@login_required
def export_csv():
    con = db()
    items = con.execute("""
        SELECT
            i.code, i.description, i.category, i.uom, i.opening,
            COALESCE(SUM(CASE WHEN t.direction = 'IN' THEN t.quantity ELSE 0 END), 0) AS total_in,
            COALESCE(SUM(CASE WHEN t.direction = 'OUT' THEN t.quantity ELSE 0 END), 0) AS total_out,
            (i.opening + COALESCE(SUM(CASE WHEN t.direction = 'IN' THEN t.quantity WHEN t.direction = 'OUT' THEN -t.quantity ELSE 0 END), 0)) AS closing,
            i.min_level, i.bin_location
        FROM items i
        LEFT JOIN transactions t ON t.item_code = i.code
        GROUP BY i.code
        ORDER BY i.category, i.code
    """).fetchall()
    con.close()

    lines = ["Item Code,Description,Category,UOM,Opening,Received,Issued,Closing,Min Level,Status,Location"]
    for it in items:
        closing = float(it["closing"] or 0)
        min_lvl = float(it["min_level"] or 0)
        status = "REORDER" if closing <= min_lvl else "OK"
        desc = (it["description"] or "").replace('"', '""')
        lines.append(f'"{it["code"]}","{desc}","{it["category"]}","{it["uom"] or ""}",{it["opening"]},{it["total_in"]},{it["total_out"]},{closing},{min_lvl},"{status}","{it["bin_location"] or ""}"')

    csv_data = "\n".join(lines)
    return send_file(
        io.BytesIO(csv_data.encode("utf-8")),
        download_name=f"Inventory_{datetime.now().strftime('%Y%m%d')}.csv",
        as_attachment=True,
        mimetype="text/csv"
    )

@app.errorhandler(403)
def forbidden(e):
    return render_page(
        "<div class='card' style='max-width:500px;margin:50px auto;text-align:center;'><h2 style='color:var(--danger);'>403 - Access Denied</h2><p>This action requires Administrator privileges. Your account has View-Only access.</p><a href='/' class='btn btn-primary'>Return to Dashboard</a></div>",
        title="403 Forbidden"
    ), 403

@app.errorhandler(400)
def bad_request(e):
    return render_page(
        f"<div class='card' style='max-width:500px;margin:50px auto;text-align:center;'><h2 style='color:var(--danger);'>400 - Bad Request</h2><p>{e.description}</p><a href='/' class='btn btn-primary'>Return to Dashboard</a></div>",
        title="400 Bad Request"
    ), 400

# Initialize database on startup
init_db()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
