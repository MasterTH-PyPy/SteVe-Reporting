"""
SteVe Ladehistorie – Web-Frontend
==================================
Berechtigungsstufen:
  admin        – Userverwaltung (SteVe-Mitarbeiter zuweisen); kein Zugriff auf Ladevorgänge
  buchhaltung  – Alle Chips, alle Zeiträume, Excel-Export
  userplus     – Eigene Chips, Monatsansicht + Tagesansicht, kein Export
  user         – Eigene Chips, nur Monatsansicht, kein Export
"""

import os
import io
import sqlite3
import subprocess
import hashlib
import secrets
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
from functools import wraps

from flask import (
    Flask, render_template, request, redirect, url_for,
    session, flash, send_file, g, abort, jsonify
)
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────
# Konfiguration
# ─────────────────────────────────────────────

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH  = os.path.join(BASE_DIR, "data", "users.db")

# SteVe-DB-Verbindung aus Umgebungsvariablen (via docker-compose env_file)
DOCKER_CONTAINER = os.environ.get("DOCKER_CONTAINER", "steve-steve-3120-db-1")
DB_NAME          = os.environ.get("DB_NAME",          "steve")
DB_USER          = os.environ.get("DB_USER",          "steve")
DB_PASSWORD      = os.environ.get("DB_PASSWORD",      "")
STEVE_VERSION    = os.environ.get("STEVE_VERSION",    "current")

ROLES = ["admin", "buchhaltung", "userplus", "user"]
APP_VERSION = "1.1.0"

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", secrets.token_hex(32))

@app.context_processor
def inject_globals():
    return {
        "app_version": APP_VERSION,
        "steve_version": STEVE_VERSION,
    }

# ─────────────────────────────────────────────
# SQLite – internes User-Management
# ─────────────────────────────────────────────

def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
    return g.db

@app.teardown_appcontext
def close_db(e=None):
    db = g.pop("db", None)
    if db:
        db.close()

def init_db():
    db = sqlite3.connect(DB_PATH)
    db.row_factory = sqlite3.Row
    db.executescript("""
        CREATE TABLE IF NOT EXISTS web_users (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            username    TEXT    NOT NULL UNIQUE,
            password    TEXT    NOT NULL,
            role        TEXT    NOT NULL DEFAULT 'user',
            steve_user  TEXT    DEFAULT NULL
        );
    """)
    row = db.execute("SELECT COUNT(*) AS cnt FROM web_users").fetchone()
    if row["cnt"] == 0:
        pw_hash = hash_password("admin")
        db.execute(
            "INSERT INTO web_users (username, password, role) VALUES (?, ?, ?)",
            ("admin", pw_hash, "admin")
        )
        db.commit()
    db.close()

def hash_password(pw: str) -> str:
    return hashlib.sha256(pw.encode()).hexdigest()

# ─────────────────────────────────────────────
# SteVe-DB-Abfragen via docker exec
# ─────────────────────────────────────────────

def run_mysql(sql: str) -> list:
    cmd = [
        "docker", "exec", DOCKER_CONTAINER,
        "mysql",
        f"--user={DB_USER}",
        f"--password={DB_PASSWORD}",
        "--batch",
        "--skip-column-names",
        "--default-character-set=utf8mb4",
        DB_NAME,
        "-e", sql,
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")
    rows = []
    for line in result.stdout.splitlines():
        if not line.strip():
            continue
        rows.append(line.split("\t"))
    return rows


def fetch_chip_notes() -> list:
    """Alle eindeutigen, nicht-leeren Notes aus ocpp_tag (HEX-kodiert)."""
    sql = "SELECT DISTINCT HEX(note) AS note_hex FROM ocpp_tag WHERE note IS NOT NULL AND note != '' ORDER BY note;"
    rows = run_mysql(sql)
    notes = []
    for r in rows:
        if not r or not r[0]:
            continue
        try:
            note = bytes.fromhex(r[0]).decode("utf-8")
        except Exception:
            note = r[0]
        note = " - ".join(p.strip() for p in note.replace("\r", "\n").split("\n") if p.strip())
        if note:
            notes.append(note)
    return notes


def fetch_assigned_notes() -> list:
    """Notes, die bereits einem internen Web-User zugewiesen sind."""
    db = get_db()
    rows = db.execute(
        "SELECT steve_user FROM web_users WHERE steve_user IS NOT NULL"
    ).fetchall()
    return [r["steve_user"] for r in rows]


def fetch_chips_for_note(note: str) -> list:
    """Alle RFID-Chips, deren Note mit dem gespeicherten Wert übereinstimmt."""
    # Note HEX-kodiert vergleichen um Sonderzeichen zu umgehen
    sql = f"SELECT id_tag FROM ocpp_tag WHERE HEX(note) != '' ORDER BY id_tag;"
    rows = run_mysql(sql)
    # Alle Chips laden und Note dekodiert vergleichen
    sql_all = "SELECT id_tag, HEX(note) AS note_hex FROM ocpp_tag WHERE note IS NOT NULL AND note != '';"
    rows_all = run_mysql(sql_all)
    chips = []
    for r in rows_all:
        if len(r) < 2:
            continue
        try:
            chip_note = bytes.fromhex(r[1]).decode("utf-8")
        except Exception:
            chip_note = r[1]
        chip_note = " - ".join(p.strip() for p in chip_note.replace("\r", "\n").split("\n") if p.strip())
        if chip_note == note:
            chips.append(r[0])
    return chips


def fetch_all_chips() -> list:
    """Alle Chips mit Note (HEX-kodiert für sichere Übertragung)."""
    sql = "SELECT id_tag, HEX(note) AS note_hex FROM ocpp_tag ORDER BY id_tag;"
    rows = run_mysql(sql)
    result = []
    for r in rows:
        if len(r) < 2:
            continue
        note_raw = r[1]
        try:
            note = bytes.fromhex(note_raw).decode("utf-8") if note_raw else ""
        except Exception:
            note = note_raw
        note = " - ".join(p.strip() for p in note.replace("\r", "\n").split("\n") if p.strip())
        result.append({"id_tag": r[0], "note": note})
    return result


def fetch_transactions(chips: list, date_from: str, date_to: str) -> list:
    """Ladevorgänge für eine Liste von Chips in einem Zeitraum."""
    if not chips:
        return []
    chip_list = ", ".join(f"'{c}'" for c in chips)
    if STEVE_VERSION == "current":
        sql = f"""
            SELECT
                ts.transaction_pk                                   AS id,
                ts.id_tag                                           AS chip,
                COALESCE(wu.username,'')                            AS benutzer,
                c.charge_box_id                                     AS station,
                c.connector_id                                      AS anschluss,
                ts.start_timestamp                                  AS start_time,
                te.stop_timestamp                                   AS stop_time,
                ROUND(TIMESTAMPDIFF(SECOND,ts.start_timestamp,te.stop_timestamp)/60.0,1) AS dauer_min,
                ts.start_value                                      AS zaehler_start,
                te.stop_value                                       AS zaehler_ende,
                ROUND((te.stop_value - ts.start_value)/1000.0,3)   AS energie_kwh,
                te.stop_reason                                      AS stopgrund
            FROM transaction_start ts
            LEFT JOIN transaction_stop te  ON ts.transaction_pk = te.transaction_pk
            LEFT JOIN connector c          ON ts.connector_pk   = c.connector_pk
            LEFT JOIN ocpp_tag ot          ON ts.id_tag         = ot.id_tag
            LEFT JOIN user_ocpp_tag uot    ON ot.ocpp_tag_pk    = uot.ocpp_tag_pk
            LEFT JOIN web_user wu          ON uot.user_pk       = wu.web_user_pk
            WHERE ts.id_tag IN ({chip_list})
              AND ts.start_timestamp BETWEEN '{date_from}' AND '{date_to}'
            ORDER BY ts.start_timestamp;
        """
    else:
        sql = f"""
            SELECT
                t.transaction_pk                                    AS id,
                t.id_tag                                            AS chip,
                COALESCE(CONCAT(u.first_name,' ',u.last_name),'')  AS benutzer,
                t.charge_box_id                                     AS station,
                t.connector_id                                      AS anschluss,
                t.start_timestamp                                   AS start_time,
                t.stop_timestamp                                    AS stop_time,
                ROUND(TIMESTAMPDIFF(SECOND,t.start_timestamp,t.stop_timestamp)/60.0,1) AS dauer_min,
                t.start_value                                       AS zaehler_start,
                t.stop_value                                        AS zaehler_ende,
                ROUND((t.stop_value - t.start_value)/1000.0,3)     AS energie_kwh,
                t.stop_reason                                       AS stopgrund
            FROM `transaction` t
            LEFT JOIN ocpp_tag ot ON t.id_tag = ot.id_tag
            LEFT JOIN `user` u    ON ot.ocpp_tag_pk = u.ocpp_tag_pk
            WHERE t.id_tag IN ({chip_list})
              AND t.start_timestamp BETWEEN '{date_from}' AND '{date_to}'
            ORDER BY t.start_timestamp;
        """
    cols = ["id","chip","benutzer","station","anschluss","start_time","stop_time",
            "dauer_min","zaehler_start","zaehler_ende","energie_kwh","stopgrund"]
    rows = run_mysql(sql)
    result = []
    for r in rows:
        if len(r) < len(cols):
            continue
        result.append(dict(zip(cols, r)))
    return result

# ─────────────────────────────────────────────
# Statistik-Abfragen
# ─────────────────────────────────────────────

def fetch_monthly_stats_for_chips(chips: list, months: list) -> list:
    """
    Gibt für jeden Monat (YYYY-MM) die Summe kWh + Anzahl Ladevorgänge zurück.
    months: Liste von {'value': 'YYYY-MM', 'label': '...'}
    Rückgabe: [{'month': 'YYYY-MM', 'label': '...', 'kwh': float, 'count': int}]
    """
    if not chips:
        return [{"month": m["value"], "label": m["label"], "kwh": 0.0, "count": 0} for m in months]
    chip_list = ", ".join(f"'{c}'" for c in chips)
    # Ältester und neuester Monat – unabhängig von der Reihenfolge der Liste
    values   = sorted(m["value"] for m in months)
    oldest   = values[0] + "-01 00:00:00"
    newest_m = values[-1]
    y, mo = int(newest_m[:4]), int(newest_m[5:7])
    newest_last = (date(y, mo, 1) + relativedelta(months=1) - relativedelta(days=1)).strftime("%Y-%m-%d 23:59:59")

    if STEVE_VERSION == "current":
        sql = f"""
            SELECT DATE_FORMAT(ts.start_timestamp,'%Y-%m') AS ym,
                   COUNT(*) AS cnt,
                   ROUND(SUM((te.stop_value - ts.start_value)/1000.0),3) AS kwh
            FROM transaction_start ts
            LEFT JOIN transaction_stop te ON ts.transaction_pk = te.transaction_pk
            WHERE ts.id_tag IN ({chip_list})
              AND ts.start_timestamp BETWEEN '{oldest}' AND '{newest_last}'
              AND te.stop_value IS NOT NULL
            GROUP BY ym ORDER BY ym;
        """
    else:
        sql = f"""
            SELECT DATE_FORMAT(t.start_timestamp,'%Y-%m') AS ym,
                   COUNT(*) AS cnt,
                   ROUND(SUM((t.stop_value - t.start_value)/1000.0),3) AS kwh
            FROM `transaction` t
            WHERE t.id_tag IN ({chip_list})
              AND t.start_timestamp BETWEEN '{oldest}' AND '{newest_last}'
              AND t.stop_value IS NOT NULL
            GROUP BY ym ORDER BY ym;
        """
    rows = run_mysql(sql)
    db_data = {}
    for r in rows:
        if len(r) >= 3:
            db_data[r[0]] = {"count": int(r[1] or 0), "kwh": float(r[2] or 0)}
    result = []
    for m in months:
        d = db_data.get(m["value"], {"count": 0, "kwh": 0.0})
        result.append({"month": m["value"], "label": m["label"], "kwh": d["kwh"], "count": d["count"]})
    return result


def fetch_monthly_stats_by_note(months: list) -> dict:
    """
    Buchhaltung: kWh pro Monat aufgeteilt nach Chip-Note.
    Rückgabe: {'labels': [...Monate...], 'notes': ['Note A', 'Note B', ...],
               'datasets': [{'note': '...', 'data': [kwh, kwh, ...]}, ...]}
    """
    # Alle Chips mit Note laden
    all_chips_info = fetch_all_chips()
    # Note → [chip_ids]
    note_to_chips = {}
    for c in all_chips_info:
        note = c["note"] or c["id_tag"]
        note_to_chips.setdefault(note, []).append(c["id_tag"])

    labels = [m["label"] for m in months]
    datasets = []
    for note, chips in sorted(note_to_chips.items()):
        stats = fetch_monthly_stats_for_chips(chips, months)
        datasets.append({
            "note": note,
            "data": [s["kwh"] for s in stats],
        })
    return {"labels": labels, "datasets": datasets}


# ─────────────────────────────────────────────
# Excel-Export (nur Buchhaltung)
# ─────────────────────────────────────────────

FONT_NAME = "Segoe UI"

def build_excel(transactions: list, label: str) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = label[:31]

    headers = ["Chip","Benutzer","Station","Anschluss","Start","Ende",
               "Dauer (min)","Zähler Start (Wh)","Zähler Ende (Wh)","Energie (kWh)","Stopgrund"]
    keys    = ["chip","benutzer","station","anschluss","start_time","stop_time",
               "dauer_min","zaehler_start","zaehler_ende","energie_kwh","stopgrund"]

    fill = PatternFill("solid", fgColor="1F3864")
    hdr_font = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=10)

    ws.append(headers)
    for c, _ in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = fill
        cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for tx in transactions:
        row = []
        for k in keys:
            val = tx.get(k, "")
            if val in ("NULL", None):
                val = ""
            row.append(val)
        ws.append(row)

    energy_col = keys.index("energie_kwh") + 1
    for row_idx in range(2, ws.max_row + 1):
        cell = ws.cell(row=row_idx, column=energy_col)
        try:
            cell.value = float(cell.value) if cell.value else 0.0
            cell.number_format = '#,##0.000'
        except (ValueError, TypeError):
            pass
        for col_idx in range(1, len(headers) + 1):
            ws.cell(row=row_idx, column=col_idx).font = Font(name=FONT_NAME, size=10)

    for i, h in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(i)].width = max(12, len(h) + 4)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()

# ─────────────────────────────────────────────
# Auth-Hilfsfunktionen
# ─────────────────────────────────────────────

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if "user_id" not in session:
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated


def role_required(*roles):
    def decorator(f):
        @wraps(f)
        def decorated(*args, **kwargs):
            if "user_id" not in session:
                return redirect(url_for("login"))
            if session.get("role") not in roles:
                abort(403)
            return f(*args, **kwargs)
        return decorated
    return decorator

# ─────────────────────────────────────────────
# Zeitraum-Hilfsfunktionen
# ─────────────────────────────────────────────

def get_months(n=12, include_current=True) -> list:
    """Gibt n Monate zurück, optional mit dem laufenden Monat voran."""
    months = []
    today = date.today().replace(day=1)
    start = 0 if include_current else 1
    for i in range(start, n + start):
        d = today - relativedelta(months=i)
        label = d.strftime("%B %Y")
        months.append({
            "value": d.strftime("%Y-%m"),
            "label": label,
            "label_select": label + (" (aktuell)" if i == 0 else ""),
            "is_current": i == 0,
        })
    return months


def get_past_months(n=12) -> list:
    return get_months(n, include_current=False)


def month_range(ym: str) -> tuple:
    y, m = int(ym[:4]), int(ym[5:7])
    first = date(y, m, 1)
    last  = first + relativedelta(months=1) - relativedelta(days=1)
    return first.strftime("%Y-%m-%d 00:00:00"), last.strftime("%Y-%m-%d 23:59:59")


def day_range(ymd: str) -> tuple:
    return f"{ymd} 00:00:00", f"{ymd} 23:59:59"


def get_days_in_month(ym: str) -> list:
    y, m = int(ym[:4]), int(ym[5:7])
    first = date(y, m, 1)
    last  = first + relativedelta(months=1) - relativedelta(days=1)
    days = []
    cur = first
    while cur <= last:
        days.append({"value": cur.strftime("%Y-%m-%d"), "label": cur.strftime("%d.%m.%Y")})
        cur += relativedelta(days=1)
    return days

# ─────────────────────────────────────────────
# Routen – Auth
# ─────────────────────────────────────────────

@app.route("/")
def index():
    if "user_id" in session:
        if session.get("role") == "admin":
            return redirect(url_for("admin_users"))
        return redirect(url_for("dashboard"))
    return redirect(url_for("login"))


@app.route("/login", methods=["GET", "POST"])
def login():
    if "user_id" in session:
        return redirect(url_for("index"))
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        db = get_db()
        user = db.execute(
            "SELECT * FROM web_users WHERE username = ?", (username,)
        ).fetchone()
        if user and user["password"] == hash_password(password):
            session.clear()
            session["user_id"] = user["id"]
            session["username"] = user["username"]
            session["role"] = user["role"]
            session.permanent = True
            if user["role"] == "admin":
                return redirect(url_for("admin_users"))
            return redirect(url_for("dashboard"))
        flash("Benutzername oder Passwort falsch.", "danger")
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

# ─────────────────────────────────────────────
# Routen – Dashboard & Ladehistorie
# ─────────────────────────────────────────────

@app.route("/dashboard")
@role_required("user", "userplus", "buchhaltung")
def dashboard():
    role   = session["role"]
    # Aktueller Monat + 12 vergangene = 13 Einträge
    months = get_months(12, include_current=True)
    all_chips_info = fetch_all_chips() if role == "buchhaltung" else []

    # Statistik für die Stat-Cards (letzte 12 abgeschlossene Monate)
    if role == "buchhaltung":
        all_chips = [c["id_tag"] for c in all_chips_info]
        stats_months = get_months(12, include_current=True)
        init_stats   = fetch_monthly_stats_for_chips(all_chips, stats_months)
        init_kwh     = round(sum(s["kwh"] for s in init_stats), 3)
        init_count   = sum(s["count"] for s in init_stats)
    else:
        db         = get_db()
        user_row   = db.execute("SELECT steve_user FROM web_users WHERE id=?",
                                (session["user_id"],)).fetchone()
        steve_user = user_row["steve_user"] if user_row else None
        chips      = fetch_chips_for_note(steve_user) if steve_user else []
        stats_months = get_months(12, include_current=True)  # inkl. laufender Monat
        monthly    = fetch_monthly_stats_for_chips(chips, stats_months)
        total_kwh  = sum(s["kwh"] for s in monthly)
        total_cnt  = sum(s["count"] for s in monthly)
        init_kwh, init_count = total_kwh, total_cnt

    return render_template("dashboard.html",
                           months=months,
                           role=role,
                           username=session["username"],
                           all_chips_info=all_chips_info,
                           init_kwh=init_kwh,
                           init_count=init_count)


@app.route("/api/stats")
@role_required("user", "userplus", "buchhaltung")
def api_stats():
    """JSON-Endpunkt für Chart.js-Daten."""
    role = session["role"]
    ym   = request.args.get("month", "")   # leer = Übersicht 12 Monate

    if role == "buchhaltung":
        if ym:
            months = [{"value": ym,
                       "label": datetime.strptime(ym, "%Y-%m").strftime("%B %Y"),
                       "is_current": False}]
        else:
            # Aktueller Monat + 12 vergangene, älteste zuerst für Chart
            months = list(reversed(get_months(12, include_current=True)))
        data = fetch_monthly_stats_by_note(months)
        all_chips = [c["id_tag"] for c in fetch_all_chips()]
        summary_stats = fetch_monthly_stats_for_chips(all_chips, months)
        total_kwh  = round(sum(s["kwh"] for s in summary_stats), 3)
        total_cnt  = sum(s["count"] for s in summary_stats)
        return jsonify({
            "type": "stacked",
            "labels": data["labels"],
            "datasets": data["datasets"],
            "total_kwh": total_kwh,
            "total_count": total_cnt,
        })
    else:
        # user / userplus: einfacher Balken kWh pro Monat
        db       = get_db()
        user_row = db.execute("SELECT steve_user FROM web_users WHERE id=?",
                              (session["user_id"],)).fetchone()
        steve_user = user_row["steve_user"] if user_row else None
        chips    = fetch_chips_for_note(steve_user) if steve_user else []
        if ym:
            months = [{"value": ym,
                       "label": datetime.strptime(ym, "%Y-%m").strftime("%B %Y"),
                       "is_current": False}]
        else:
            # Aktueller Monat + 12 vergangene, älteste zuerst für Chart
            months = list(reversed(get_months(12, include_current=True)))
        stats = fetch_monthly_stats_for_chips(chips, months)
        return jsonify({
            "type": "bar",
            "labels": [s["label"] for s in stats],
            "data":   [s["kwh"]  for s in stats],
            "total_kwh":   round(sum(s["kwh"]   for s in stats), 3),
            "total_count": sum(s["count"] for s in stats),
        })


@app.route("/history")
@role_required("user", "userplus", "buchhaltung")
def history():
    role = session["role"]
    ym   = request.args.get("month", "")
    day  = request.args.get("day", "")
    chip_filter = request.args.getlist("chips")

    if not ym:
        flash("Bitte einen Monat auswählen.", "warning")
        return redirect(url_for("dashboard"))

    # Zeitraum
    if day and role in ("userplus", "buchhaltung"):
        date_from, date_to = day_range(day)
        period_label = datetime.strptime(day, "%Y-%m-%d").strftime("%d.%m.%Y")
    else:
        date_from, date_to = month_range(ym)
        period_label = datetime.strptime(ym, "%Y-%m").strftime("%B %Y")

    # Chips ermitteln
    if role == "buchhaltung":
        all_chips = [c["id_tag"] for c in fetch_all_chips()]
        chips = chip_filter if chip_filter else all_chips
    else:
        db = get_db()
        user = db.execute("SELECT * FROM web_users WHERE id=?", (session["user_id"],)).fetchone()
        steve_user = user["steve_user"] if user else None
        if not steve_user:
            flash("Deinem Konto ist noch kein Chip-Kommentar zugewiesen.", "warning")
            return redirect(url_for("dashboard"))
        chips = fetch_chips_for_note(steve_user)

    transactions = fetch_transactions(chips, date_from, date_to)

    total_energy = 0.0
    for tx in transactions:
        try:
            total_energy += float(tx.get("energie_kwh") or 0)
        except (ValueError, TypeError):
            pass

    days_list = get_days_in_month(ym) if role in ("userplus", "buchhaltung") else []
    all_chips_info = fetch_all_chips() if role == "buchhaltung" else []

    # Kein Export für den laufenden Monat
    today_ym   = date.today().strftime("%Y-%m")
    allow_export = (role == "buchhaltung") and (ym != today_ym)

    return render_template("history.html",
                           transactions=transactions,
                           period_label=period_label,
                           total_energy=round(total_energy, 3),
                           role=role,
                           ym=ym,
                           day=day,
                           days_list=days_list,
                           all_chips_info=all_chips_info,
                           selected_chips=chip_filter,
                           allow_export=allow_export,
                           username=session["username"])


@app.route("/export")
@role_required("buchhaltung")
def export():
    ym = request.args.get("month", "")
    chip_filter = request.args.getlist("chips")
    if not ym:
        flash("Bitte einen Monat auswählen.", "warning")
        return redirect(url_for("dashboard"))

    date_from, date_to = month_range(ym)
    label = datetime.strptime(ym, "%Y-%m").strftime("%B %Y")
    all_chips = [c["id_tag"] for c in fetch_all_chips()]
    chips = chip_filter if chip_filter else all_chips
    transactions = fetch_transactions(chips, date_from, date_to)
    xlsx_bytes = build_excel(transactions, label)

    return send_file(
        io.BytesIO(xlsx_bytes),
        as_attachment=True,
        download_name=f"Ladehistorie_{ym}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ─────────────────────────────────────────────
# Routen – Admin / Benutzerverwaltung
# ─────────────────────────────────────────────

@app.route("/admin/users")
@role_required("admin")
def admin_users():
    db = get_db()
    users = db.execute("SELECT * FROM web_users ORDER BY username").fetchall()
    return render_template("admin_users.html", users=users, username=session["username"])


@app.route("/admin/users/new", methods=["GET", "POST"])
@role_required("admin")
def admin_user_new():
    all_notes = fetch_chip_notes()
    assigned  = fetch_assigned_notes()
    available = [n for n in all_notes if n not in assigned]

    if request.method == "POST":
        uname = request.form.get("username", "").strip()
        pw    = request.form.get("password", "").strip()
        role  = request.form.get("role", "user")
        steve = request.form.get("steve_user", "") or None

        if not uname or not pw:
            flash("Benutzername und Passwort sind Pflichtfelder.", "danger")
        elif role not in ROLES:
            flash("Ungültige Rolle.", "danger")
        else:
            db = get_db()
            try:
                db.execute(
                    "INSERT INTO web_users (username, password, role, steve_user) VALUES (?,?,?,?)",
                    (uname, hash_password(pw), role, steve)
                )
                db.commit()
                flash(f"Benutzer '{uname}' angelegt.", "success")
                return redirect(url_for("admin_users"))
            except sqlite3.IntegrityError:
                flash(f"Benutzername '{uname}' ist bereits vergeben.", "danger")

    return render_template("admin_user_form.html",
                           action="new",
                           available_steve_users=available,
                           roles=ROLES,
                           form={},
                           username=session["username"])


@app.route("/admin/users/<int:uid>/edit", methods=["GET", "POST"])
@role_required("admin")
def admin_user_edit(uid):
    db   = get_db()
    user = db.execute("SELECT * FROM web_users WHERE id=?", (uid,)).fetchone()
    if not user:
        abort(404)

    all_notes = fetch_chip_notes()
    assigned  = fetch_assigned_notes()
    # Aktuell zugewiesene Note bleibt auswählbar
    available = [n for n in all_notes if n not in assigned or n == user["steve_user"]]

    if request.method == "POST":
        uname = request.form.get("username", "").strip()
        pw    = request.form.get("password", "").strip()
        role  = request.form.get("role", "user")
        steve = request.form.get("steve_user", "") or None

        if not uname:
            flash("Benutzername ist ein Pflichtfeld.", "danger")
        elif role not in ROLES:
            flash("Ungültige Rolle.", "danger")
        else:
            try:
                if pw:
                    db.execute(
                        "UPDATE web_users SET username=?, password=?, role=?, steve_user=? WHERE id=?",
                        (uname, hash_password(pw), role, steve, uid)
                    )
                else:
                    db.execute(
                        "UPDATE web_users SET username=?, role=?, steve_user=? WHERE id=?",
                        (uname, role, steve, uid)
                    )
                db.commit()
                flash(f"Benutzer '{uname}' aktualisiert.", "success")
                return redirect(url_for("admin_users"))
            except sqlite3.IntegrityError:
                flash(f"Benutzername '{uname}' ist bereits vergeben.", "danger")

    return render_template("admin_user_form.html",
                           action="edit",
                           uid=uid,
                           available_steve_users=available,
                           roles=ROLES,
                           form=dict(user),
                           username=session["username"])


@app.route("/admin/users/<int:uid>/delete", methods=["POST"])
@role_required("admin")
def admin_user_delete(uid):
    if uid == session["user_id"]:
        flash("Du kannst dich nicht selbst löschen.", "danger")
        return redirect(url_for("admin_users"))
    db = get_db()
    db.execute("DELETE FROM web_users WHERE id=?", (uid,))
    db.commit()
    flash("Benutzer gelöscht.", "success")
    return redirect(url_for("admin_users"))


@app.route("/admin/users/<int:uid>/reset-password", methods=["POST"])
@role_required("admin")
def admin_user_reset_pw(uid):
    new_pw = request.form.get("new_password", "").strip()
    if not new_pw:
        flash("Neues Passwort darf nicht leer sein.", "danger")
        return redirect(url_for("admin_users"))
    db = get_db()
    db.execute("UPDATE web_users SET password=? WHERE id=?", (hash_password(new_pw), uid))
    db.commit()
    flash("Passwort wurde zurückgesetzt.", "success")
    return redirect(url_for("admin_users"))

# ─────────────────────────────────────────────
# Error Handler
# ─────────────────────────────────────────────

@app.route("/favicon.ico")
def favicon():
    return "", 204


@app.errorhandler(403)
def forbidden(e):
    return render_template("error.html", code=403, msg="Zugriff verweigert."), 403

@app.errorhandler(404)
def not_found(e):
    return render_template("error.html", code=404, msg="Seite nicht gefunden."), 404

# ─────────────────────────────────────────────
# Start
# ─────────────────────────────────────────────

if __name__ == "__main__":
    init_db()
    app.run(host="0.0.0.0", port=5000, debug=False)
