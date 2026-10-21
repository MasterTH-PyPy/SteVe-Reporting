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
import logging
import sqlite3
import subprocess
import hashlib
import re
from werkzeug.security import generate_password_hash, check_password_hash as _check_pw
import secrets
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
from functools import wraps

from flask import (
    Flask, render_template, request, redirect, url_for,
    session, flash, send_file, g, abort, jsonify
)
from werkzeug.middleware.proxy_fix import ProxyFix
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address
from flask_wtf.csrf import CSRFProtect
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from collections import defaultdict

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

# ─────────────────────────────────────────────
# Logging
# ─────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger("steve")

def audit(action: str, detail: str = ""):
    ip   = request.remote_addr if request else "-"
    user = session.get("username", "-") if session else "-"
    msg  = f"{action} | user={user} | ip={ip}"
    if detail:
        msg += f" | {detail}"
    log.info(msg)

# ─────────────────────────────────────────────
# App
# ─────────────────────────────────────────────

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", secrets.token_hex(32))
app.wsgi_app = ProxyFix(app.wsgi_app, x_proto=1, x_host=1)
app.config["SESSION_COOKIE_SECURE"]   = True
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["WTF_CSRF_TIME_LIMIT"]     = 3600  # 1h

csrf    = CSRFProtect(app)
limiter = Limiter(get_remote_address, app=app, default_limits=[])

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
            id                  INTEGER PRIMARY KEY AUTOINCREMENT,
            username            TEXT    NOT NULL UNIQUE,
            password            TEXT    NOT NULL,
            role                TEXT    NOT NULL DEFAULT 'user',
            steve_user          TEXT    DEFAULT NULL,
            created_at          TEXT    DEFAULT NULL,
            password_changed_at TEXT    DEFAULT NULL,
            last_login          TEXT    DEFAULT NULL
        );
    """)
    # Migration: Spalten nachrüsten falls sie noch fehlen
    cols = {r[1] for r in db.execute("PRAGMA table_info(web_users)")}
    if "created_at" not in cols:
        db.execute("ALTER TABLE web_users ADD COLUMN created_at TEXT DEFAULT NULL")
    if "password_changed_at" not in cols:
        db.execute("ALTER TABLE web_users ADD COLUMN password_changed_at TEXT DEFAULT NULL")
    if "last_login" not in cols:
        db.execute("ALTER TABLE web_users ADD COLUMN last_login TEXT DEFAULT NULL")
    db.commit()

    row = db.execute("SELECT COUNT(*) AS cnt FROM web_users").fetchone()
    if row["cnt"] == 0:
        pw_hash = hash_password("admin")
        now = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")
        db.execute(
            "INSERT INTO web_users (username, password, role, created_at, password_changed_at) VALUES (?, ?, ?, ?, ?)",
            ("admin", pw_hash, "admin", now, now)
        )
        db.commit()
    db.close()

def hash_password(pw: str) -> str:
    return generate_password_hash(pw)

def verify_password(pw: str, hashed: str) -> bool:
    # Rückwärtskompatibilität: alte SHA-256-Hashes (64 hex Zeichen)
    if len(hashed) == 64 and re.fullmatch(r'[0-9a-f]{64}', hashed):
        import hashlib
        return hashlib.sha256(pw.encode()).hexdigest() == hashed
    return _check_pw(hashed, pw)

def _validate_date(d: str) -> str:
    """Wirft ValueError bei ungültigem Datumsformat."""
    if not re.fullmatch(r'\d{4}-\d{2}-\d{2}( \d{2}:\d{2}:\d{2})?', d):
        raise ValueError(f"Ungültiges Datum: {d!r}")
    return d

def _sanitize_chip(chip: str) -> str:
    """Nur erlaubte Zeichen für RFID-Chip-IDs (Hex, Bindestriche, alphanumerisch)."""
    if not re.fullmatch(r'[0-9A-Za-z:\-]{1,64}', chip):
        raise ValueError(f"Ungültige Chip-ID: {chip!r}")
    return chip

# ─────────────────────────────────────────────
# SteVe-DB-Abfragen via docker exec
# ─────────────────────────────────────────────

def run_mysql(sql: str) -> list:
    cmd = [
        "docker", "exec",
        "--env", f"MYSQL_PWD={DB_PASSWORD}",  # nicht in Prozessliste sichtbar
        DOCKER_CONTAINER,
        "mysql",
        f"--user={DB_USER}",
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
    chip_list = ", ".join(f"'{_sanitize_chip(c)}'" for c in chips)
    date_from = _validate_date(date_from)
    date_to   = _validate_date(date_to)
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
    chip_list = ", ".join(f"'{_sanitize_chip(c)}'" for c in chips)
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


def fetch_daily_stats_for_chips(chips: list, ym: str) -> dict:
    """
    kWh + Anzahl pro Tag für einen Monat.
    Rückgabe: {'YYYY-MM-DD': kwh, 'YYYY-MM-DD_count': int, ...}
    """
    if not chips:
        return {}
    chip_list  = ", ".join(f"'{_sanitize_chip(c)}'" for c in chips)
    date_from, date_to = month_range(ym)
    if STEVE_VERSION == "current":
        sql = f"""
            SELECT DATE(ts.start_timestamp) AS d,
                   COUNT(*) AS cnt,
                   ROUND(SUM((te.stop_value - ts.start_value)/1000.0),3) AS kwh
            FROM transaction_start ts
            LEFT JOIN transaction_stop te ON ts.transaction_pk = te.transaction_pk
            WHERE ts.id_tag IN ({chip_list})
              AND ts.start_timestamp BETWEEN '{date_from}' AND '{date_to}'
              AND te.stop_value IS NOT NULL
            GROUP BY d ORDER BY d;
        """
    else:
        sql = f"""
            SELECT DATE(t.start_timestamp) AS d,
                   COUNT(*) AS cnt,
                   ROUND(SUM((t.stop_value - t.start_value)/1000.0),3) AS kwh
            FROM `transaction` t
            WHERE t.id_tag IN ({chip_list})
              AND t.start_timestamp BETWEEN '{date_from}' AND '{date_to}'
              AND t.stop_value IS NOT NULL
            GROUP BY d ORDER BY d;
        """
    rows   = run_mysql(sql)
    result = {}
    for r in rows:
        if len(r) >= 3:
            result[r[0]]               = float(r[2] or 0)
            result[r[0] + "_count"]    = int(r[1] or 0)
    return result


def filter_chips_by_note(all_chips_info: list, note_filter: str) -> list:
    """
    Filtert Chip-Liste nach Substring im Note-Feld (case-insensitive).
    note_filter leer → alle zurückgeben.
    """
    if not note_filter:
        return all_chips_info
    f = note_filter.lower()
    return [c for c in all_chips_info if f in (c["note"] or c["id_tag"]).lower()]


def build_note_to_chips(chips_info: list) -> dict:
    """Gruppiert Chip-Liste nach Note → {note: [id_tag, ...]}."""
    note_to_chips = {}
    for c in chips_info:
        note = c["note"] or c["id_tag"]
        note_to_chips.setdefault(note, []).append(c["id_tag"])
    return note_to_chips


def fetch_daily_stats_by_note(ym: str, note_filter: str = "") -> dict:
    """
    Buchhaltung: kWh pro Tag und Note für einen Monat.
    note_filter: optionaler Substring-Filter auf das Note-Feld.
    Rückgabe: {'labels': [...], 'datasets': [{'note':..., 'data':[kwh,...]}, ...]}
    """
    chips_info    = filter_chips_by_note(fetch_all_chips(), note_filter)
    note_to_chips = build_note_to_chips(chips_info)

    days = get_days_in_month(ym)
    datasets = []
    for note, chips in sorted(note_to_chips.items()):
        stats = fetch_daily_stats_for_chips(chips, ym)
        datasets.append({
            "note": note,
            "data": [stats.get(d["value"], 0.0) for d in days],
        })
    return {
        "labels":   [d["label"] for d in days],
        "datasets": datasets,
    }


def fetch_monthly_stats_by_note(months: list, note_filter: str = "") -> dict:
    """
    Buchhaltung: kWh pro Monat aufgeteilt nach Chip-Note.
    note_filter: optionaler Substring-Filter auf das Note-Feld.
    Rückgabe: {'labels': [...Monate...], 'datasets': [{'note': '...', 'data': [kwh,...]}, ...]}
    """
    chips_info    = filter_chips_by_note(fetch_all_chips(), note_filter)
    note_to_chips = build_note_to_chips(chips_info)

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

# ── Excel-Stil-Konstanten ────────────────────────────────────
_XL_SUBHEAD_FILL = PatternFill("solid", fgColor="2E75B6")
_XL_ALT_FILL     = PatternFill("solid", fgColor="DEEAF1")
_XL_TOTAL_FILL   = PatternFill("solid", fgColor="BDD7EE")
_XL_WHITE_FILL   = PatternFill("solid", fgColor="FFFFFF")
_XL_THIN         = Side(style="thin", color="B8CCE4")
_XL_BORDER       = Border(left=_XL_THIN, right=_XL_THIN, top=_XL_THIN, bottom=_XL_THIN)
_XL_CENTER       = Alignment(horizontal="center", vertical="center")
_XL_LEFT         = Alignment(horizontal="left",   vertical="center")
_XL_RIGHT        = Alignment(horizontal="right",  vertical="center")

_XL_SUBHEAD_FONT = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=10)
_XL_BODY_FONT    = Font(name=FONT_NAME, size=10)
_XL_TOTAL_FONT   = Font(name=FONT_NAME, bold=True, size=10, color="1F4E79")
_XL_LINK_FONT    = Font(name=FONT_NAME, size=10, color="1F4E79", underline="single")

# Spalten-Definition: (key_in_tx, Spaltenüberschrift, Breite, Ausrichtung)
_XL_COLUMNS = [
    ("id",            "Trans.-ID",          10, _XL_RIGHT),
    ("chip",          "RFID-Chip",          16, _XL_LEFT),
    ("benutzer",      "Benutzer",           18, _XL_LEFT),
    ("station",       "Ladestation",        22, _XL_LEFT),
    ("anschluss",     "Anschluss",          11, _XL_CENTER),
    ("start_time",    "Start",              20, _XL_CENTER),
    ("stop_time",     "Ende",               20, _XL_CENTER),
    ("dauer_min",     "Dauer (min)",        12, _XL_RIGHT),
    ("zaehler_start", "Zähler Start (Wh)", 18, _XL_RIGHT),
    ("zaehler_ende",  "Zähler Ende (Wh)",  18, _XL_RIGHT),
    ("energie_kwh",   "Energie (kWh)",      15, _XL_RIGHT),
    ("stopgrund",     "Stopgrund",          18, _XL_LEFT),
]


def _xl_hdr_cell(ws, row, col, val):
    c = ws.cell(row=row, column=col, value=val)
    c.font = _XL_SUBHEAD_FONT; c.fill = _XL_SUBHEAD_FILL
    c.alignment = _XL_CENTER;  c.border = _XL_BORDER
    ws.row_dimensions[row].height = 20
    return c


def _xl_data_cell(ws, row, col, val, align, fill):
    c = ws.cell(row=row, column=col, value=val)
    c.font = _XL_BODY_FONT; c.fill = fill
    c.alignment = align; c.border = _XL_BORDER
    return c


def _xl_total_cell(ws, row, col, val, align):
    c = ws.cell(row=row, column=col, value=val)
    c.font = _XL_TOTAL_FONT; c.fill = _XL_TOTAL_FILL
    c.alignment = align; c.border = _XL_BORDER
    return c


def _xl_link_cell(ws, row, col, val, target_sheet, fill):
    c = ws.cell(row=row, column=col, value=val)
    c.hyperlink = f"#'{target_sheet}'!A1"
    c.font = _XL_LINK_FONT; c.fill = fill
    c.border = _XL_BORDER; c.alignment = _XL_LEFT
    return c


def _xl_sheet_header(ws, title_val, sub_val, n_cols):
    ws.sheet_view.showGridLines = False
    ws.merge_cells(f"A1:{get_column_letter(n_cols)}1")
    t = ws["A1"]
    t.value = title_val
    t.font  = Font(name=FONT_NAME, bold=True, size=14, color="1F4E79")
    t.alignment = _XL_CENTER
    ws.row_dimensions[1].height = 30
    ws.merge_cells(f"A2:{get_column_letter(n_cols)}2")
    s = ws["A2"]
    s.value = sub_val
    s.font  = Font(name=FONT_NAME, italic=True, size=10, color="595959")
    s.alignment = _XL_CENTER
    ws.row_dimensions[2].height = 18
    ws.row_dimensions[3].height = 8


def build_excel(transactions: list, label: str, date_from: str = "", date_to: str = "",
                chip_notes_map: dict = None) -> bytes:
    """Baut eine Excel-Datei im Format des Original-Skripts.

    Sheets:
      1. Auswertung nach Benutzer  – Zusammenfassung pro Name
      2. Übersicht pro Chip        – alle Chips mit Link zu Detailblatt
      3..N  Ein Detailblatt pro RFID-Chip

    chip_notes_map: optional dict {id_tag -> note_string} für alle bekannten Chips,
                    auch solche ohne Transaktionen im Zeitraum.
    """
    if chip_notes_map is None:
        chip_notes_map = {}

    # ── Daten nach Chip gruppieren ───────────────────────────
    all_data: dict = {}
    for tx in transactions:
        chip = tx.get("chip", "")
        if not chip:
            continue
        all_data.setdefault(chip, []).append(tx)

    # Alle bekannten Chips einschließen (auch ohne Transaktionen)
    for chip in chip_notes_map:
        all_data.setdefault(chip, [])

    # ── Statistiken vorberechnen ─────────────────────────────
    chip_stats: dict = {}
    for chip, rows in all_data.items():
        energy   = sum(float(r.get("energie_kwh") or 0) for r in rows
                       if r.get("energie_kwh") not in (None, "NULL", ""))
        duration = sum(float(r.get("dauer_min") or 0)   for r in rows
                       if r.get("dauer_min") not in (None, "NULL", ""))
        benutzer = (rows[0].get("benutzer") or "").strip() if rows else ""
        # note aus der Map bevorzugen (enthält Chip-Kommentar aus ocpp_tag)
        note = chip_notes_map.get(chip, "") or ""
        chip_stats[chip] = {
            "energy":   round(energy, 3),
            "duration": round(duration, 1),
            "count":    len(rows),
            "benutzer": benutzer,
            "note":     note,
        }

    user_stats: dict = defaultdict(lambda: {"energy": 0.0, "duration": 0.0, "count": 0, "chips": []})
    for chip, st in chip_stats.items():
        name = st["benutzer"] or st["note"] or f"(kein Name – {chip})"
        user_stats[name]["energy"]   += st["energy"]
        user_stats[name]["duration"] += st["duration"]
        user_stats[name]["count"]    += st["count"]
        user_stats[name]["chips"].append(chip)

    date_from_str = date_from[:10] if date_from else ""
    date_to_str   = date_to[:10]   if date_to   else ""
    zeitraum_sub  = (f"Zeitraum: {date_from_str} bis {date_to_str}"
                     if date_from_str else label)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ── Sheet 1: Auswertung nach Benutzer ────────────────────
    ws_user = wb.create_sheet("Auswertung nach Benutzer")
    _xl_sheet_header(ws_user,
        f"SteVe Ladehistorie – Auswertung nach Benutzer – {label}",
        zeitraum_sub, 4)

    for ci, h in enumerate(["Benutzer", "Ladevorgänge", "Ges. Energie (kWh)", "RFID-Chips"], 1):
        _xl_hdr_cell(ws_user, 4, ci, h)

    urow = 5
    for name, st in sorted(user_stats.items()):
        fill = _XL_ALT_FILL if urow % 2 == 0 else _XL_WHITE_FILL
        _xl_data_cell(ws_user, urow, 1, name,          _XL_LEFT,  fill)
        _xl_data_cell(ws_user, urow, 2, st["count"],   _XL_RIGHT, fill)
        c = _xl_data_cell(ws_user, urow, 3, st["energy"], _XL_RIGHT, fill)
        c.number_format = "0.000"
        _xl_data_cell(ws_user, urow, 4, ", ".join(st["chips"]), _XL_LEFT, fill)
        ws_user.row_dimensions[urow].height = 18
        urow += 1

    if urow > 5:
        _xl_total_cell(ws_user, urow, 1, "GESAMT", _XL_LEFT)
        _xl_total_cell(ws_user, urow, 2, f"=SUM(B5:B{urow-1})", _XL_RIGHT)
        c = _xl_total_cell(ws_user, urow, 3, f"=SUM(C5:C{urow-1})", _XL_RIGHT)
        c.number_format = "0.000"
        _xl_total_cell(ws_user, urow, 4, "", _XL_LEFT)

    ws_user.column_dimensions["A"].width = 28
    ws_user.column_dimensions["B"].width = 14
    ws_user.column_dimensions["C"].width = 20
    ws_user.column_dimensions["D"].width = 50

    # ── Sheet 2: Übersicht pro Chip ──────────────────────────
    ws_summary = wb.create_sheet("Übersicht pro Chip")
    _xl_sheet_header(ws_summary,
        f"SteVe Ladehistorie – Übersicht pro Chip – {label}",
        zeitraum_sub, 6)

    for ci, h in enumerate(["RFID-Chip", "Note / Kommentar", "Benutzer",
                              "Ladevorgänge", "Ges. Energie (kWh)", "Blatt"], 1):
        _xl_hdr_cell(ws_summary, 4, ci, h)

    summary_row = 5
    for chip, st in chip_stats.items():
        fill = _XL_ALT_FILL if summary_row % 2 == 0 else _XL_WHITE_FILL
        sheet_name = chip[:31]
        _xl_link_cell(ws_summary, summary_row, 1, chip, sheet_name, fill)
        _xl_data_cell(ws_summary, summary_row, 2, st["note"],     _XL_LEFT,  fill)
        _xl_data_cell(ws_summary, summary_row, 3, st["benutzer"], _XL_LEFT,  fill)
        _xl_data_cell(ws_summary, summary_row, 4, st["count"],    _XL_RIGHT, fill)
        c = _xl_data_cell(ws_summary, summary_row, 5, st["energy"], _XL_RIGHT, fill)
        c.number_format = "0.000"
        _xl_link_cell(ws_summary, summary_row, 6, sheet_name, sheet_name, fill)
        ws_summary.row_dimensions[summary_row].height = 18
        summary_row += 1

    if summary_row > 5:
        _xl_total_cell(ws_summary, summary_row, 1, "GESAMT", _XL_LEFT)
        for col in (2, 3, 6):
            _xl_total_cell(ws_summary, summary_row, col, "", _XL_LEFT)
        _xl_total_cell(ws_summary, summary_row, 4, f"=SUM(D5:D{summary_row-1})", _XL_RIGHT)
        c = _xl_total_cell(ws_summary, summary_row, 5, f"=SUM(E5:E{summary_row-1})", _XL_RIGHT)
        c.number_format = "0.000"

    ws_summary.column_dimensions["A"].width = 20
    ws_summary.column_dimensions["B"].width = 28
    ws_summary.column_dimensions["C"].width = 22
    ws_summary.column_dimensions["D"].width = 14
    ws_summary.column_dimensions["E"].width = 20
    ws_summary.column_dimensions["F"].width = 22

    # ── Ein Detailblatt pro Chip ─────────────────────────────
    for chip, rows in sorted(all_data.items()):
        safe_name = chip[:31]
        note      = chip_notes_map.get(chip, "") or ""
        ws = wb.create_sheet(safe_name)
        ws.sheet_view.showGridLines = False
        ws.freeze_panes = "A4"

        n_cols = len(_XL_COLUMNS)
        ws.merge_cells(f"A1:{get_column_letter(n_cols)}1")
        t = ws["A1"]
        t.value     = (f"Ladehistorie: {chip}"
                       + (f"  |  {note}" if note else "")
                       + f"  |  {label}")
        t.font      = Font(name=FONT_NAME, bold=True, size=12, color="1F4E79")
        t.alignment = _XL_CENTER
        ws.row_dimensions[1].height = 26

        ws.merge_cells(f"A2:{get_column_letter(n_cols)}2")
        ws["A2"].value     = zeitraum_sub
        ws["A2"].font      = Font(name=FONT_NAME, italic=True, size=9, color="595959")
        ws["A2"].alignment = _XL_CENTER
        ws.row_dimensions[2].height = 16
        ws.row_dimensions[3].height = 6

        for col_idx, (_, col_label, width, _align) in enumerate(_XL_COLUMNS, 1):
            _xl_hdr_cell(ws, 3, col_idx, col_label)
            ws.column_dimensions[get_column_letter(col_idx)].width = width
        ws.row_dimensions[3].height = 20

        if not rows:
            ws.merge_cells(f"A4:{get_column_letter(n_cols)}4")
            c = ws["A4"]
            c.value     = "Keine Ladevorgänge im gewählten Zeitraum."
            c.font      = Font(name=FONT_NAME, italic=True, size=10, color="595959")
            c.alignment = _XL_CENTER
            continue

        for row_idx, tx in enumerate(rows, 4):
            fill = _XL_ALT_FILL if row_idx % 2 == 0 else _XL_WHITE_FILL
            ws.row_dimensions[row_idx].height = 18
            for col_idx, (key, _, _, align) in enumerate(_XL_COLUMNS, 1):
                raw = tx.get(key, "")
                val = None if raw in ("NULL", None, "") else raw
                if key in ("anschluss", "zaehler_start", "zaehler_ende"):
                    try: val = int(val)
                    except: pass
                elif key in ("dauer_min", "energie_kwh"):
                    try: val = float(val)
                    except: pass
                c = ws.cell(row=row_idx, column=col_idx, value=val)
                c.font = _XL_BODY_FONT; c.fill = fill
                c.alignment = align; c.border = _XL_BORDER
                if key == "energie_kwh" and val is not None:
                    c.number_format = "0.000"
                elif key == "dauer_min" and val is not None:
                    c.number_format = "0.0"

        total_row = len(rows) + 4
        ws.row_dimensions[total_row].height = 20

        # Zuerst alle Zellen der SUMME-Zeile stylen/befüllen,
        # Spaltennummern für SUM-Spalten ermitteln
        dauer_col  = next(i for i, (k,_,_,_) in enumerate(_XL_COLUMNS, 1) if k == "dauer_min")
        energie_col= next(i for i, (k,_,_,_) in enumerate(_XL_COLUMNS, 1) if k == "energie_kwh")
        merge_end  = dauer_col - 1  # Merge endet eine Spalte vor dauer_min

        # Alle Zellen ZUERST befüllen/stylen, dann erst mergen
        for col_idx, (key, _, _, _align) in enumerate(_XL_COLUMNS, 1):
            if key in ("dauer_min", "energie_kwh"):
                col_letter = get_column_letter(col_idx)
                c = ws.cell(row=total_row, column=col_idx,
                            value=f"=SUM({col_letter}4:{col_letter}{total_row-1})")
                c.font = _XL_TOTAL_FONT; c.fill = _XL_TOTAL_FILL
                c.border = _XL_BORDER; c.alignment = _XL_RIGHT
                c.number_format = "0.000" if key == "energie_kwh" else "0.0"
            else:
                val = "SUMME" if col_idx == 1 else ""
                c = ws.cell(row=total_row, column=col_idx, value=val)
                c.font = _XL_TOTAL_FONT; c.fill = _XL_TOTAL_FILL
                c.border = _XL_BORDER
                if col_idx == 1:
                    c.alignment = _XL_RIGHT

        # Merge: A bis Spalte vor dauer_min; dauer_min + energie_kwh bleiben eigene Zellen
        ws.merge_cells(f"A{total_row}:{get_column_letter(merge_end)}{total_row}")

    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True
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
            # Passwortänderung erzwingen bevor andere Seiten zugänglich sind
            if session.get("must_change_password") and request.endpoint != "change_password":
                return redirect(url_for("change_password"))
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
@limiter.limit("10 per minute")
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
        if user and verify_password(password, user["password"]):
            session.clear()
            session["user_id"] = user["id"]
            session["username"] = user["username"]
            session["role"] = user["role"]
            session.permanent = True
            # Last-Login timestamp aktualisieren
            now = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")
            db.execute("UPDATE web_users SET last_login=? WHERE id=?", (now, user["id"]))
            db.commit()
            audit("LOGIN_OK", f"role={user['role']}")
            # Erstes Login: created_at == password_changed_at → Passwort ändern erzwingen
            ca  = user["created_at"]
            pca = user["password_changed_at"]
            if ca and pca and ca == pca and user["role"] != "admin":
                session["must_change_password"] = True
                return redirect(url_for("change_password"))
            if user["role"] == "admin":
                return redirect(url_for("admin_users"))
            return redirect(url_for("dashboard"))
        log.warning("LOGIN_FAIL | user=%s | ip=%s", username, request.remote_addr)
        flash("Benutzername oder Passwort falsch.", "danger")
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/change-password", methods=["GET", "POST"])
@login_required
def change_password():
    """Eigenes Passwort ändern (für alle Rollen außer admin, erzwungen beim Erst-Login)."""
    must_change = session.get("must_change_password", False)
    error = None
    if request.method == "POST":
        current_pw  = request.form.get("current_password", "")
        new_pw      = request.form.get("new_password", "").strip()
        confirm_pw  = request.form.get("confirm_password", "").strip()
        db = get_db()
        user = db.execute("SELECT * FROM web_users WHERE id=?", (session["user_id"],)).fetchone()
        if not user or not verify_password(current_pw, user["password"]):
            error = "Das aktuelle Passwort ist falsch."
        elif len(new_pw) < 6:
            error = "Das neue Passwort muss mindestens 6 Zeichen lang sein."
        elif new_pw != confirm_pw:
            error = "Die neuen Passwörter stimmen nicht überein."
        else:
            now = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")
            db.execute(
                "UPDATE web_users SET password=?, password_changed_at=? WHERE id=?",
                (hash_password(new_pw), now, session["user_id"])
            )
            db.commit()
            session.pop("must_change_password", None)
            audit("PASSWORD_CHANGE")
            flash("Passwort erfolgreich geändert.", "success")
            role = session.get("role")
            if role == "admin":
                return redirect(url_for("admin_users"))
            return redirect(url_for("dashboard"))
    return render_template("change_password.html", must_change=must_change, error=error)

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

    today_ym = date.today().strftime("%Y-%m")
    return render_template("dashboard.html",
                           months=months,
                           role=role,
                           username=session["username"],
                           all_chips_info=all_chips_info,
                           init_kwh=init_kwh,
                           init_count=init_count,
                           today_ym=today_ym)


@app.route("/api/stats")
@role_required("user", "userplus", "buchhaltung")
def api_stats():
    """JSON-Endpunkt für Chart.js-Daten."""
    role        = session["role"]
    ym          = request.args.get("month", "")        # leer = Übersicht 12 Monate
    note_filter = request.args.get("note_filter", "").strip()

    if role == "buchhaltung":
        # Für Totals nur die gefilterten Chips verwenden
        filtered_chips_info = filter_chips_by_note(fetch_all_chips(), note_filter)
        filtered_chips = [c["id_tag"] for c in filtered_chips_info]
        if ym:
            # Tagesauflösung: gestapelter Balken pro Tag, aufgeteilt nach Note
            data = fetch_daily_stats_by_note(ym, note_filter)
            day_stats = fetch_daily_stats_for_chips(filtered_chips, ym)
            days = get_days_in_month(ym)
            total_kwh = round(sum(day_stats.get(d["value"], 0.0) for d in days), 3)
            total_cnt = sum(day_stats.get(d["value"] + "_count", 0) for d in days)
        else:
            # Monatsauflösung: gestapelter Balken pro Monat
            months = list(reversed(get_months(12, include_current=True)))
            data = fetch_monthly_stats_by_note(months, note_filter)
            summary = fetch_monthly_stats_for_chips(filtered_chips, months)
            total_kwh = round(sum(s["kwh"] for s in summary), 3)
            total_cnt = sum(s["count"] for s in summary)
        return jsonify({
            "type": "stacked",
            "labels": data["labels"],
            "datasets": data["datasets"],
            "total_kwh": total_kwh,
            "total_count": total_cnt,
            "note_filter": note_filter,
        })
    else:
        # user / userplus: einfacher Balken kWh pro Monat oder pro Tag
        db       = get_db()
        user_row = db.execute("SELECT steve_user FROM web_users WHERE id=?",
                              (session["user_id"],)).fetchone()
        steve_user = user_row["steve_user"] if user_row else None
        chips    = fetch_chips_for_note(steve_user) if steve_user else []
        if ym:
            # Tagesauflösung für den gewählten Monat
            days  = get_days_in_month(ym)
            stats = fetch_daily_stats_for_chips(chips, ym)
            labels = [d["label"] for d in days]
            data   = [stats.get(d["value"], 0.0) for d in days]
            counts = [stats.get(d["value"] + "_count", 0) for d in days]
            return jsonify({
                "type": "bar_daily",
                "labels": labels,
                "data":   data,
                "total_kwh":   round(sum(data), 3),
                "total_count": sum(counts),
            })
        else:
            # Aktueller Monat + 12 vergangene, älteste zuerst für Chart
            months = list(reversed(get_months(12, include_current=True)))
            stats  = fetch_monthly_stats_for_chips(chips, months)
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
    role        = session["role"]
    ym          = request.args.get("month", "")
    day         = request.args.get("day", "")
    chip_filter = request.args.getlist("chips")
    note_filter = request.args.get("note_filter", "Mitarbeiter").strip()

    if not ym:
        flash("Bitte wähle einen Monat aus.", "warning")
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
        all_chips_info_full = fetch_all_chips()
        # Note-Filter anwenden (Chip-Auswahl hat Vorrang vor Note-Filter)
        if chip_filter:
            chips = chip_filter
        else:
            filtered = filter_chips_by_note(all_chips_info_full, note_filter)
            chips = [c["id_tag"] for c in filtered]
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
    today_ym     = date.today().strftime("%Y-%m")
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
                           note_filter=note_filter,
                           allow_export=allow_export,
                           username=session["username"])


@app.route("/export")
@role_required("buchhaltung")
def export():
    ym          = request.args.get("month", "")
    chip_filter = request.args.getlist("chips")
    note_filter = request.args.get("note_filter", "").strip()
    if not ym:
        flash("Bitte wähle einen Monat aus.", "warning")
        return redirect(url_for("dashboard"))

    date_from, date_to = month_range(ym)
    label = datetime.strptime(ym, "%Y-%m").strftime("%B %Y")

    # Alle bekannten Chips mit Notes laden (auch ohne Transaktionen im Zeitraum)
    all_chips_info = fetch_all_chips()
    if chip_filter:
        chips = chip_filter
        chip_notes_map = {c["id_tag"]: c["note"] for c in all_chips_info
                          if c["id_tag"] in chip_filter}
    else:
        filtered = filter_chips_by_note(all_chips_info, note_filter)
        chips = [c["id_tag"] for c in filtered]
        chip_notes_map = {c["id_tag"]: c["note"] for c in filtered}

    transactions = fetch_transactions(chips, date_from, date_to)
    xlsx_bytes = build_excel(transactions, label, date_from, date_to,
                             chip_notes_map=chip_notes_map)

    audit("EXPORT", f"month={ym} chips={len(chips)} rows={len(transactions)}")
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

        force_pw = request.form.get("force_pw_change") == "on"
        if not uname or not pw:
            flash("Benutzername und Passwort sind Pflichtfelder.", "danger")
        elif role not in ROLES:
            flash("Ungültige Rolle.", "danger")
        else:
            db = get_db()
            try:
                now = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")
                # force_pw_change: created_at == password_changed_at → Login erzwingt Änderung
                # kein force: password_changed_at auf NULL setzen → kein Zwang
                pw_changed_at = now if force_pw else None
                db.execute(
                    "INSERT INTO web_users (username, password, role, steve_user, created_at, password_changed_at) VALUES (?,?,?,?,?,?)",
                    (uname, hash_password(pw), role, steve, now, pw_changed_at)
                )
                db.commit()
                audit("USER_CREATE", f"new_user={uname} role={role}")
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

        force_pw = request.form.get("force_pw_change") == "on"
        if not uname:
            flash("Benutzername ist ein Pflichtfeld.", "danger")
        elif role not in ROLES:
            flash("Ungültige Rolle.", "danger")
        else:
            try:
                if pw:
                    now = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")
                    # Wenn force_pw_change: created_at auf now setzen, password_changed_at auch auf now
                    # → ca == pca → Zwang beim nächsten Login
                    # Wenn nicht: password_changed_at auf NULL → kein Zwang
                    if force_pw and role != "admin":
                        db.execute(
                            "UPDATE web_users SET username=?, password=?, role=?, steve_user=?, created_at=?, password_changed_at=? WHERE id=?",
                            (uname, hash_password(pw), role, steve, now, now, uid)
                        )
                    else:
                        db.execute(
                            "UPDATE web_users SET username=?, password=?, role=?, steve_user=?, password_changed_at=? WHERE id=?",
                            (uname, hash_password(pw), role, steve, None, uid)
                        )
                else:
                    # Kein neues Passwort — nur force_pw_change Flag anpassen
                    if force_pw and role != "admin":
                        # created_at auf jetzt, password_changed_at auch → Zwang
                        now = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")
                        db.execute(
                            "UPDATE web_users SET username=?, role=?, steve_user=?, created_at=?, password_changed_at=? WHERE id=?",
                            (uname, role, steve, now, now, uid)
                        )
                    else:
                        db.execute(
                            "UPDATE web_users SET username=?, role=?, steve_user=?, password_changed_at=? WHERE id=?",
                            (uname, role, steve, None, uid)
                        )
                db.commit()
                audit("USER_EDIT", f"target_uid={uid} new_name={uname} role={role}")
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
    target = db.execute("SELECT username FROM web_users WHERE id=?", (uid,)).fetchone()
    db.execute("DELETE FROM web_users WHERE id=?", (uid,))
    db.commit()
    audit("USER_DELETE", f"target_uid={uid} target_name={target['username'] if target else '?'}")
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
    target = db.execute("SELECT username FROM web_users WHERE id=?", (uid,)).fetchone()
    db.execute("UPDATE web_users SET password=? WHERE id=?", (hash_password(new_pw), uid))
    db.commit()
    audit("PASSWORD_RESET", f"target_uid={uid} target_name={target['username'] if target else '?'}")
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
