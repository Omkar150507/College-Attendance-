from flask import Flask, render_template, request, redirect, url_for, session, flash, Response, abort
import os
import sqlite3
from urllib.parse import urlparse
from xml.sax.saxutils import escape as xml_escape

try:
    import psycopg2
    from psycopg2.extras import RealDictCursor
    from psycopg2.pool import ThreadedConnectionPool
    POSTGRES_AVAILABLE = True
except ImportError:
    psycopg2 = None
    RealDictCursor = None
    ThreadedConnectionPool = None
    POSTGRES_AVAILABLE = False

from flask import Flask, render_template, request, redirect, url_for, session, flash, Response
from pathlib import Path
from functools import wraps
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo
import csv
import io
import json
import zipfile
import smtplib
import re
import math
import secrets
from markupsafe import Markup
import time
import hmac
import hashlib
import struct
import base64
from calendar import monthrange
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment
    from openpyxl.drawing.image import Image as XLImage
    OPENPYXL_AVAILABLE = True
except ImportError:
    Workbook = None
    XLImage = None
    OPENPYXL_AVAILABLE = False
from email.message import EmailMessage
try:
    import qrcode
    QRCODE_AVAILABLE = True
except ImportError:
    qrcode = None
    QRCODE_AVAILABLE = False
from itsdangerous import URLSafeTimedSerializer, BadSignature, SignatureExpired, BadData
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.middleware.proxy_fix import ProxyFix
from flask_wtf.csrf import CSRFProtect
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address
from config import COLLEGE_NAME, COLLEGE_TAGLINE, LOGO_FILE
from attendance_statement_pdf import build_statement_pdf, build_class_statement_pdf

IST = ZoneInfo("Asia/Kolkata")


def now_ist():
    """Current time in India. Render servers run in UTC, so use these helpers (not the server date) for college dates."""
    return datetime.now(IST)


def today_ist():
    return now_ist().date()


BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "attendance.db"
DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()
RESET_TOKEN_MAX_AGE = 30 * 60

# Small per-worker PostgreSQL pool. Render can occasionally recycle an idle
# PostgreSQL/TLS connection. A pooled connection that has gone stale must never
# be handed back to a request; otherwise pages can intermittently return HTTP 500
# with errors such as: "SSL error: decryption failed or bad record mac".
PG_POOL = None

def _discard_pg_connection(conn):
    """Return a bad PostgreSQL connection to the pool and force it to close."""
    if conn is None or PG_POOL is None:
        return
    try:
        PG_POOL.putconn(conn, close=True)
    except Exception:
        try:
            conn.close()
        except Exception:
            pass

def _new_pg_pool():
    if not POSTGRES_AVAILABLE:
        raise RuntimeError("PostgreSQL driver is not installed. Run: pip install psycopg2-binary")
    if not DATABASE_URL:
        raise RuntimeError("DATABASE_URL is not configured")

    # TCP keepalives reduce the chance that an idle Render/PostgreSQL connection
    # survives in the pool after its network/TLS session has already expired.
    return ThreadedConnectionPool(
        1,
        int(os.environ.get("DB_POOL_MAX", "10")),
        DATABASE_URL,
        sslmode=os.environ.get("DB_SSLMODE", "require"),
        connect_timeout=10,
        keepalives=1,
        keepalives_idle=int(os.environ.get("DB_KEEPALIVES_IDLE", "30")),
        keepalives_interval=int(os.environ.get("DB_KEEPALIVES_INTERVAL", "10")),
        keepalives_count=int(os.environ.get("DB_KEEPALIVES_COUNT", "5")),
    )

def get_pg_pool():
    global PG_POOL
    if PG_POOL is None:
        PG_POOL = _new_pg_pool()
    return PG_POOL

def get_healthy_pg_connection():
    """Get a live connection; discard stale/broken pooled connections."""
    pool = get_pg_pool()
    last_error = None
    for _ in range(3):
        conn = pool.getconn()
        try:
            if getattr(conn, "closed", 0):
                _discard_pg_connection(conn)
                continue
            # A lightweight round-trip detects stale TLS/socket state before the
            # actual application query uses the connection.
            with conn.cursor() as cur:
                cur.execute("SELECT 1")
                cur.fetchone()
            conn.rollback()
            conn.autocommit = False
            return conn
        except Exception as exc:
            last_error = exc
            _discard_pg_connection(conn)
    raise RuntimeError(f"Unable to obtain a healthy PostgreSQL connection: {last_error}")

def get_smtp_settings():
    """Read SMTP settings at request time so Render environment changes are picked up after restart."""
    host = os.environ.get("SMTP_HOST", "").strip()
    username = os.environ.get("SMTP_USERNAME", "").strip()
    password = os.environ.get("SMTP_PASSWORD", "").strip().replace(" ", "")
    from_email = os.environ.get("SMTP_FROM_EMAIL", "").strip() or username
    try:
        port = int(os.environ.get("SMTP_PORT", "587") or 587)
    except ValueError:
        port = 587
    use_tls = os.environ.get("SMTP_USE_TLS", "1").strip().lower() not in ("0", "false", "no")
    return host, port, username, password, from_email, use_tls
YEAR_OPTIONS = ["1st Year", "2nd Year", "3rd Year", "4th Year"]
SEMESTERS_BY_YEAR = {"1st Year": [1, 2], "2nd Year": [3, 4], "3rd Year": [5, 6], "4th Year": [7, 8]}
SECURITY_QUESTIONS = [
    "What was the name of your first school?",
    "What is the name of your hometown?",
    "What was your favorite subject in school?",
    "What is your favorite teacher's name?",
]

app = Flask(__name__)
# Trust only the known reverse-proxy fields used by Render. Do not let arbitrary
# Host headers be treated as proxy-verified hostnames.
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1)
_configured_secret = os.environ.get("SECRET_KEY", "").strip()
if not _configured_secret and (DATABASE_URL or os.environ.get("RENDER", "").lower() == "true" or os.environ.get("FLASK_ENV") == "production"):
    raise RuntimeError("SECRET_KEY must be configured with a strong random value in production.")
app.secret_key = _configured_secret or secrets.token_hex(32)
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=bool(DATABASE_URL or os.environ.get("RENDER", "").lower() == "true" or os.environ.get("FLASK_ENV") == "production"),
    SESSION_REFRESH_EACH_REQUEST=False,
    WTF_CSRF_TIME_LIMIT=3600,
    MAX_CONTENT_LENGTH=16 * 1024 * 1024,
    PERMANENT_SESSION_LIFETIME=timedelta(days=14),
)
csrf = CSRFProtect(app)
# Rate limits protect login/recovery endpoints from automated guessing.
# Set RATELIMIT_STORAGE_URI to a Redis URL (for example a free Upstash database,
# "rediss://default:PASSWORD@HOST:6379") so limits are shared and survive restarts.
# Without it the app falls back to per-process memory, which still works but is weaker.
_rate_limit_storage = os.environ.get("RATELIMIT_STORAGE_URI", "").strip()
if not _rate_limit_storage:
    _rate_limit_storage = "memory://"
    app.logger.warning(
        "RATELIMIT_STORAGE_URI is not set; using in-memory rate limiting. "
        "Configure a Redis-compatible URL for shared limits."
    )
_limiter_options = {}
if _rate_limit_storage.startswith(("redis://", "rediss://")):
    # Never let a slow/unavailable Redis hang or break logins.
    _limiter_options["storage_options"] = {"socket_connect_timeout": 2, "socket_timeout": 2}
limiter = Limiter(
    get_remote_address,
    app=app,
    default_limits=[],
    storage_uri=_rate_limit_storage,
    swallow_errors=True,
    in_memory_fallback_enabled=True,
    **_limiter_options,
)

SESSION_HOURS = 12      # normal login lasts 12 hours
REMEMBER_DAYS = 14      # "Remember Me" login lasts 14 days


LOGIN_LOCK_FAILS = 5            # wrong passwords from one IP for one username before it is locked
LOGIN_LOCK_GLOBAL_FAILS = 25    # wrong passwords for one username from all IPs together
LOGIN_LOCK_MINUTES = 15
LOGIN_WINDOW_MINUTES = 30


def client_ip():
    return (request.remote_addr or "")[:64]


def create_login_session(user_id, ttl_seconds):
    """Server-side record of a login, so a logout or password change really ends it."""
    token = secrets.token_urlsafe(32)
    now = datetime.utcnow()
    db = get_db()
    try:
        db.execute("DELETE FROM login_sessions WHERE expires_at < ?", (now.isoformat(),))
        db.execute("INSERT INTO login_sessions(token,user_id,created_at,expires_at) VALUES(?,?,?,?)",
                   (token, user_id, now.isoformat(), (now + timedelta(seconds=ttl_seconds)).isoformat()))
        db.commit()
    finally:
        db.close()
    return token


def revoke_login_sessions(user_id, keep_token=None):
    db = get_db()
    try:
        if keep_token:
            db.execute("DELETE FROM login_sessions WHERE user_id=? AND token<>?", (user_id, keep_token))
        else:
            db.execute("DELETE FROM login_sessions WHERE user_id=?", (user_id,))
        db.commit()
    finally:
        db.close()


def lockout_minutes_left(db, username):
    """Minutes a username is still locked (0 = not locked). Expired locks are cleaned up."""
    uname = (username or "").strip().lower()[:80]
    now = datetime.utcnow()
    left = 0
    for key in (f"u:{uname}|ip:{client_ip()}", f"u:{uname}"):
        row = db.execute("SELECT locked_until FROM login_attempts WHERE key=?", (key,)).fetchone()
        if row and row["locked_until"]:
            until = datetime.fromisoformat(row["locked_until"])
            if until > now:
                left = max(left, int((until - now).total_seconds() // 60) + 1)
            else:
                db.execute("DELETE FROM login_attempts WHERE key=?", (key,))
                db.commit()
    return left


def record_failed_login(username, user_id=None):
    uname = (username or "").strip().lower()[:80]
    ip = client_ip()
    now = datetime.utcnow()
    newly_locked = False
    db = get_db()
    try:
        for key, limit, minutes in ((f"u:{uname}|ip:{ip}", LOGIN_LOCK_FAILS, LOGIN_LOCK_MINUTES),
                                    (f"u:{uname}", LOGIN_LOCK_GLOBAL_FAILS, LOGIN_LOCK_MINUTES * 2)):
            row = db.execute("SELECT fails, first_at FROM login_attempts WHERE key=?", (key,)).fetchone()
            in_window = bool(row and row["first_at"] and now - datetime.fromisoformat(row["first_at"]) < timedelta(minutes=LOGIN_WINDOW_MINUTES))
            fails = int(row["fails"]) + 1 if in_window else 1
            first_at = row["first_at"] if in_window else now.isoformat()
            locked_until = (now + timedelta(minutes=minutes)).isoformat() if fails >= limit else None
            if fails == limit:
                newly_locked = True
            if row:
                db.execute("UPDATE login_attempts SET fails=?, first_at=?, locked_until=? WHERE key=?", (fails, first_at, locked_until, key))
            else:
                db.execute("INSERT INTO login_attempts(key,fails,first_at,locked_until) VALUES(?,?,?,?)", (key, fails, first_at, locked_until))
        db.commit()
    finally:
        db.close()
    log_activity_for(user_id, "Failed login", f"username={uname} ip={ip}")
    if newly_locked:
        log_activity_for(user_id, "Login locked", f"username={uname} ip={ip} locked for {LOGIN_LOCK_MINUTES} minutes")
        alert_admins("Security alert: login locked",
                     f"Too many wrong passwords for '{uname}' from {ip}. Logging in is blocked for {LOGIN_LOCK_MINUTES} minutes.\n"
                     "If this was not you or a student, check Admin > Activity Log.")


def clear_failed_logins(username):
    uname = (username or "").strip().lower()[:80]
    db = get_db()
    try:
        db.execute("DELETE FROM login_attempts WHERE key=?", (f"u:{uname}|ip:{client_ip()}",))
        db.commit()
    finally:
        db.close()


def alert_admins(subject, body):
    """Best-effort security email to every admin that has an address."""
    try:
        if not email_is_configured():
            return
        db = get_db()
        try:
            admins = db.execute("SELECT email FROM users WHERE role='admin' AND email<>''").fetchall()
        finally:
            db.close()
        for a in admins:
            send_email_message(a["email"], "Admin", f"{COLLEGE_NAME} - {subject}", body)
    except Exception as exc:  # an alert must never break a login
        app.logger.warning("Security alert email failed: %s", exc)


def safe_cell(value):
    """Stops spreadsheet formula injection: text that starts with = + - @ is stored as plain text."""
    if isinstance(value, str) and value[:1] in ("=", "+", "-", "@", "\t", "\r"):
        return "'" + value
    return value


class SafeCsvWriter:
    def __init__(self, writer):
        self._writer = writer

    def writerow(self, row):
        self._writer.writerow([safe_cell(c) for c in row])


def sanitize_workbook(wb):
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str) and v[:1] in ("=", "+", "-", "@", "\t", "\r"):
                    cell.data_type = "s"   # literal text, Excel never evaluates it


@app.before_request
def enforce_session_lifetime():
    """Log users out when their time limit is over, after logout/password change, or when the account is blocked."""
    if "user_id" not in session or request.endpoint == "static":
        return
    try:
        valid = float(session.get("exp", 0) or 0) >= time.time()
    except (TypeError, ValueError):
        valid = False
    if valid:
        valid = False
        sid = session.get("sid")
        if sid:
            db = get_db()
            try:
                row = db.execute("""SELECT s.expires_at, u.approved, u.role, u.student_id FROM login_sessions s
                                    JOIN users u ON u.id=s.user_id WHERE s.token=? AND s.user_id=?""", (sid, session["user_id"])).fetchone()
                if row and row["expires_at"] >= datetime.utcnow().isoformat() and int(row["approved"] or 0) == 1:
                    valid = True
                    if row["role"] == "student" and row["student_id"]:
                        b = db.execute("SELECT blocked FROM students WHERE id=?", (row["student_id"],)).fetchone()
                        valid = bool(b) and int(b["blocked"] or 0) == 0
            finally:
                db.close()
    if not valid:
        session.clear()
        flash("Your session has expired. Please log in again.", "error")


ICON_PATHS = {
    "home": '<path d="M4 11l8-7 8 7v9a1 1 0 01-1 1h-4v-6H9v6H5a1 1 0 01-1-1z"/>',
    "cal": '<rect x="4" y="5" width="16" height="15" rx="2"/><path d="M8 3v4M16 3v4M4 10h16"/>',
    "check": '<path d="M5 12l5 5 9-11"/>', "x": '<path d="M6 6l12 12M18 6L6 18"/>',
    "qr": '<path d="M4 8V5a1 1 0 011-1h3M16 4h3a1 1 0 011 1v3M20 16v3a1 1 0 01-1 1h-3M8 20H5a1 1 0 01-1-1v-3M8 12h8"/>',
    "chart": '<path d="M4 20V4M4 20h16"/><path d="M8 16v-5M12 16V8M16 16v-8"/>',
    "alert": '<path d="M12 3l10 18H2z"/><path d="M12 10v5M12 18h.01"/>',
    "clock": '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
    "history": '<path d="M3 12a9 9 0 109-9M3 4v5h5"/><path d="M12 7v5l3 2"/>',
    "leave": '<path d="M7 3h8l4 4v14H7z"/><path d="M15 3v4h4M10 12h6M10 16h6"/>',
    "users": '<circle cx="9" cy="8" r="3.5"/><path d="M2 20c0-3.5 3-5.5 7-5.5s7 2 7 5.5"/><path d="M16 4.5a3.5 3.5 0 010 7M18 14.8c2.2.6 4 2.2 4 5.2"/>',
    "user": '<circle cx="12" cy="8" r="4"/><path d="M4 21c0-4 4-6 8-6s8 2 8 6"/>',
    "user-plus": '<circle cx="9" cy="8" r="4"/><path d="M2 20c0-4 3-6 7-6M18 8v6M15 11h6"/>',
    "book": '<path d="M4 5a2 2 0 012-2h13v16H6a2 2 0 00-2 2z"/><path d="M4 19V5"/>',
    "gear": '<circle cx="12" cy="12" r="3"/><path d="M12 2v3M12 19v3M4.2 4.2l2.1 2.1M17.7 17.7l2.1 2.1M2 12h3M19 12h3M4.2 19.8l2.1-2.1M17.7 6.3l2.1-2.1"/>',
    "shield": '<path d="M12 3l8 3v6c0 5-3.5 8-8 9-4.5-1-8-4-8-9V6z"/><path d="M9 12l2 2 4-4"/>',
    "bell": '<path d="M6 9a6 6 0 0112 0c0 6 2 7 2 7H4s2-1 2-7"/><path d="M10 20a2 2 0 004 0"/>',
    "search": '<circle cx="11" cy="11" r="7"/><path d="M20 20l-4-4"/>',
    "menu": '<path d="M4 7h16M4 12h16M4 17h16"/>', "chev": '<path d="M6 9l6 6 6-6"/>',
    "logout": '<path d="M9 4H5a1 1 0 00-1 1v14a1 1 0 001 1h4M16 8l4 4-4 4M20 12H9"/>',
    "download": '<path d="M12 4v11M7 11l5 5 5-5M5 20h14"/>', "upload": '<path d="M12 20V9M7 13l5-5 5 5M5 4h14"/>',
    "up": '<path d="M12 20V6M6 12l6-6 6 6"/>', "restore": '<path d="M4 12a8 8 0 108-8M4 4v5h5"/>',
    "list": '<path d="M8 6h12M8 12h12M8 18h12M4 6h.01M4 12h.01M4 18h.01"/>', "edit": '<path d="M4 20h4L19 9l-4-4L4 16z"/>',
    "percent": '<path d="M19 5L5 19"/><circle cx="7" cy="7" r="2.5"/><circle cx="17" cy="17" r="2.5"/>',
    "plus": '<path d="M12 5v14M5 12h14"/>', "eye": '<path d="M2 12s4-7 10-7 10 7 10 7-4 7-10 7S2 12 2 12z"/><circle cx="12" cy="12" r="3"/>', "dot": '<circle cx="12" cy="12" r="4"/>',
}


def icon(name, size=20):
    """Inline SVG icon (stroke style) for templates: {{ icon('home') }}. Static markup only, never user input."""
    return Markup(f'<svg class="ic" width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
                  f'stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">{ICON_PATHS.get(name, ICON_PATHS["dot"])}</svg>')


app.jinja_env.globals["icon"] = icon


def reset_serializer():
    return URLSafeTimedSerializer(app.secret_key, salt="password-reset")


def email_is_configured():
    if os.environ.get("BREVO_API_KEY", "").strip() and os.environ.get("BREVO_SENDER_EMAIL", "").strip():
        return True
    host, port, username, password, from_email, use_tls = get_smtp_settings()
    return bool(host and username and password and from_email)


def send_email_message(to_email, to_name, subject, text_body, html_body=None, attachments=None):
    """Send an email. Uses the Brevo HTTPS API when configured (works on Render free, where
    SMTP ports are blocked); otherwise falls back to SMTP. attachments = [(filename, bytes)]."""
    attachments = attachments or []
    api_key = os.environ.get("BREVO_API_KEY", "").strip()
    sender_email = os.environ.get("BREVO_SENDER_EMAIL", "").strip()
    if api_key and sender_email:
        import urllib.request, urllib.error
        payload = {
            "sender": {"name": os.environ.get("BREVO_SENDER_NAME", "").strip() or COLLEGE_NAME, "email": sender_email},
            "to": [{"email": to_email, "name": to_name or to_email}],
            "subject": subject,
            "textContent": text_body,
        }
        if html_body:
            payload["htmlContent"] = html_body
        if attachments:
            payload["attachment"] = [{"name": n, "content": base64.b64encode(c).decode("ascii")} for n, c in attachments]
        req = urllib.request.Request(
            "https://api.brevo.com/v3/smtp/email",
            data=json.dumps(payload).encode("utf-8"),
            headers={"api-key": api_key, "content-type": "application/json", "accept": "application/json"},
            method="POST",
        )
        try:
            with urllib.request.urlopen(req, timeout=25) as resp:
                resp.read()
            return
        except urllib.error.HTTPError as exc:
            detail = ""
            try:
                detail = exc.read().decode("utf-8", "ignore")[:200]
            except Exception:
                pass
            raise RuntimeError(f"Email service error {exc.code}. {detail}")
    host, port, username, password, from_email, use_tls = get_smtp_settings()
    if not (host and username and password and from_email):
        raise RuntimeError("Email is not configured. Set BREVO_API_KEY and BREVO_SENDER_EMAIL.")
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = from_email
    msg["To"] = to_email
    msg.set_content(text_body)
    if html_body:
        msg.add_alternative(html_body, subtype="html")
    for name, content in attachments:
        msg.add_attachment(content, maintype="application", subtype="octet-stream", filename=name)
    if use_tls and port == 465:
        with smtplib.SMTP_SSL(host, port, timeout=20) as server:
            server.login(username, password)
            server.send_message(msg)
    else:
        with smtplib.SMTP(host, port, timeout=20) as server:
            if use_tls:
                server.starttls()
            server.login(username, password)
            server.send_message(msg)


def send_password_reset_email(to_email, full_name, reset_link):
    text = (
        f"Hello {full_name or 'User'},\n\n"
        f"A password reset was requested for your {COLLEGE_NAME} account.\n\n"
        f"Open this link within 30 minutes to set a new password:\n{reset_link}\n\n"
        "The link works only once. If you did not request this, you can ignore this email.\n"
    )
    html = (
        f"<p>Hello {xml_escape(full_name or 'User')},</p>"
        f"<p>A password reset was requested for your {xml_escape(COLLEGE_NAME)} account.</p>"
        f"<p><a href=\"{xml_escape(reset_link)}\">Click here to set a new password</a> (valid for 30 minutes, works once).</p>"
        "<p>If you did not request this, you can ignore this email.</p>"
    )
    send_email_message(to_email, full_name, f"{COLLEGE_NAME} - Password Reset", text, html)


class CompatRow(dict):
    def __getitem__(self, key):
        if isinstance(key, int):
            return list(self.values())[key]
        return super().__getitem__(key)

class CompatCursor:
    def __init__(self, cursor, postgres=False):
        self.cursor = cursor
        self.postgres = postgres
        self.lastrowid = None

    def fetchone(self):
        row = self.cursor.fetchone()
        return CompatRow(row) if row is not None else None

    def fetchall(self):
        return [CompatRow(r) for r in self.cursor.fetchall()]

class CompatDB:
    def __init__(self):
        self.postgres = bool(DATABASE_URL)
        self.conn = None
        self.pool = None
        self._closed = False
        if self.postgres:
            self.pool = get_pg_pool()
            self.conn = get_healthy_pg_connection()
            self.conn.autocommit = False
        else:
            self.conn = sqlite3.connect(DB_PATH)
            self.conn.row_factory = sqlite3.Row
            self.conn.execute("PRAGMA foreign_keys = ON")

    def _prepare_sql(self, sql):
        sql = sql.replace("?", "%s")
        # SQLite-specific INSERT OR IGNORE -> PostgreSQL equivalent.
        sql = sql.replace("INSERT OR IGNORE INTO", "INSERT INTO")
        if "INSERT INTO subject_teachers" in sql and "ON CONFLICT" not in sql:
            sql = sql.rstrip().rstrip(";") + " ON CONFLICT (teacher_id, subject_id) DO NOTHING"
        return sql

    def _replace_broken_connection(self):
        old = self.conn
        _discard_pg_connection(old)
        self.conn = get_healthy_pg_connection()
        self.conn.autocommit = False

    def execute(self, sql, params=()):
        if self.postgres:
            sql = self._prepare_sql(sql)
            try:
                cur = self.conn.cursor(cursor_factory=RealDictCursor)
                cur.execute(sql, params)
                return CompatCursor(cur, True)
            except (psycopg2.OperationalError, psycopg2.InterfaceError) as exc:
                # Retry only when the connection is still at the beginning of a
                # transaction. Retrying inside a partially completed transaction
                # could silently lose earlier writes. This safely fixes the common
                # stale pooled-connection case seen on read-only page loads.
                try:
                    can_retry = self.conn.get_transaction_status() == psycopg2.extensions.TRANSACTION_STATUS_IDLE
                except Exception:
                    can_retry = False
                if not can_retry:
                    raise exc
                self._replace_broken_connection()
                cur = self.conn.cursor(cursor_factory=RealDictCursor)
                cur.execute(sql, params)
                return CompatCursor(cur, True)
        cur = self.conn.execute(sql, params)
        wrapped = CompatCursor(cur, False)
        wrapped.lastrowid = cur.lastrowid
        return wrapped

    def executemany(self, sql, seq_of_params):
        """Execute a parameterized statement for each parameter tuple."""
        if self.postgres:
            sql = self._prepare_sql(sql)
            try:
                cur = self.conn.cursor(cursor_factory=RealDictCursor)
                cur.executemany(sql, seq_of_params)
                return CompatCursor(cur, True)
            except (psycopg2.OperationalError, psycopg2.InterfaceError) as exc:
                try:
                    can_retry = self.conn.get_transaction_status() == psycopg2.extensions.TRANSACTION_STATUS_IDLE
                except Exception:
                    can_retry = False
                if not can_retry:
                    raise exc
                self._replace_broken_connection()
                cur = self.conn.cursor(cursor_factory=RealDictCursor)
                cur.executemany(sql, seq_of_params)
                return CompatCursor(cur, True)
        cur = self.conn.executemany(sql, seq_of_params)
        wrapped = CompatCursor(cur, False)
        wrapped.lastrowid = cur.lastrowid
        return wrapped

    def executescript(self, script):
        if self.postgres:
            # PostgreSQL schema used for hosted deployment.
            statements = [x.strip() for x in script.split(";") if x.strip()]
            for stmt in statements:
                stmt = stmt.replace("INTEGER PRIMARY KEY AUTOINCREMENT", "SERIAL PRIMARY KEY")
                stmt = stmt.replace("CURRENT_TIMESTAMP", "CURRENT_TIMESTAMP")
                stmt = stmt.replace("TEXT DEFAULT ''", "TEXT DEFAULT ''")
                self.execute(stmt)
        else:
            self.conn.executescript(script)

    def commit(self):
        self.conn.commit()

    def rollback(self):
        try:
            self.conn.rollback()
        except Exception:
            pass

    def close(self):
        if self._closed:
            return
        self._closed = True
        if self.postgres:
            conn = self.conn
            self.conn = None
            if conn is None:
                return
            try:
                # Always rollback before returning a connection to the pool so a
                # later request never inherits an unfinished transaction.
                conn.rollback()
            except Exception:
                pass
            try:
                if getattr(conn, "closed", 0):
                    self.pool.putconn(conn, close=True)
                else:
                    self.pool.putconn(conn)
            except Exception:
                try:
                    conn.close()
                except Exception:
                    pass
        else:
            self.conn.close()


def get_db():
    return CompatDB()


def ensure_column(db, table, column, definition):
    if db.postgres:
        exists = db.execute("SELECT 1 FROM information_schema.columns WHERE table_name=? AND column_name=?", (table, column)).fetchone()
        if not exists:
            db.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")
    else:
        cols = {r[1] for r in db.execute(f"PRAGMA table_info({table})").fetchall()}
        if column not in cols:
            db.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")


def migrate_attendance_for_lectures(db):
    """Upgrade old attendance schema to support lecture 1/2/3 per subject per day."""
    if db.postgres:
        cols = db.execute("SELECT column_name FROM information_schema.columns WHERE table_name='attendance'").fetchall()
        names = {r["column_name"] for r in cols}
        if "lecture_no" not in names:
            db.execute("ALTER TABLE attendance ADD COLUMN lecture_no INTEGER NOT NULL DEFAULT 1")
        if "lecture_time" not in names:
            db.execute("ALTER TABLE attendance ADD COLUMN lecture_time TEXT DEFAULT ''")
        # Replace the old 3-column unique constraint with the new 4-column constraint.
        constraints = db.execute("""SELECT conname FROM pg_constraint
            WHERE conrelid='attendance'::regclass AND contype='u'""").fetchall()
        for c in constraints:
            db.execute(f'ALTER TABLE attendance DROP CONSTRAINT IF EXISTS "{c["conname"]}"')
        db.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_attendance_student_subject_date_lecture ON attendance(student_id,subject_id,attendance_date,lecture_no)")
    else:
        cols = db.execute("PRAGMA table_info(attendance)").fetchall()
        names = {r[1] for r in cols}
        if "lecture_no" not in names:
            db.execute("ALTER TABLE attendance RENAME TO attendance_old")
            db.execute("""CREATE TABLE attendance (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                student_id INTEGER NOT NULL,
                subject_id INTEGER NOT NULL,
                attendance_date TEXT NOT NULL,
                lecture_no INTEGER NOT NULL DEFAULT 1,
                lecture_time TEXT DEFAULT '',
                status TEXT NOT NULL CHECK(status IN ('Present','Absent')),
                marked_by INTEGER,
                UNIQUE(student_id, subject_id, attendance_date, lecture_no),
                FOREIGN KEY(student_id) REFERENCES students(id) ON DELETE CASCADE,
                FOREIGN KEY(subject_id) REFERENCES subjects(id) ON DELETE CASCADE,
                FOREIGN KEY(marked_by) REFERENCES users(id) ON DELETE SET NULL
            )""")
            db.execute("""INSERT INTO attendance(id,student_id,subject_id,attendance_date,lecture_no,lecture_time,status,marked_by)
                SELECT id,student_id,subject_id,attendance_date,1,'',status,marked_by FROM attendance_old""")
            db.execute("DROP TABLE attendance_old")
        else:
            db.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_attendance_student_subject_date_lecture ON attendance(student_id,subject_id,attendance_date,lecture_no)")


def init_db():
    db = get_db()
    db.executescript("""
    CREATE TABLE IF NOT EXISTS students (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        roll_no TEXT UNIQUE NOT NULL,
        prn TEXT DEFAULT '',
        name TEXT NOT NULL,
        course TEXT NOT NULL DEFAULT 'B.Pharm',
        year TEXT NOT NULL DEFAULT '1st Year',
        semester INTEGER NOT NULL DEFAULT 1,
        division TEXT NOT NULL DEFAULT 'A',
        batch TEXT DEFAULT '',
        blocked INTEGER NOT NULL DEFAULT 0
    );

    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password TEXT NOT NULL,
        role TEXT NOT NULL DEFAULT 'teacher',
        full_name TEXT DEFAULT '',
        email TEXT DEFAULT '',
        mobile TEXT DEFAULT '',
        security_question TEXT DEFAULT '',
        security_answer TEXT DEFAULT '',
        approved INTEGER NOT NULL DEFAULT 1,
        student_id INTEGER,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(student_id) REFERENCES students(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS subjects (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        code TEXT UNIQUE NOT NULL,
        name TEXT NOT NULL,
        year TEXT NOT NULL DEFAULT '1st Year'
    );

    CREATE TABLE IF NOT EXISTS subject_teachers (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        teacher_id INTEGER NOT NULL,
        subject_id INTEGER NOT NULL,
        UNIQUE(teacher_id, subject_id),
        FOREIGN KEY(teacher_id) REFERENCES users(id) ON DELETE CASCADE,
        FOREIGN KEY(subject_id) REFERENCES subjects(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS attendance (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        student_id INTEGER NOT NULL,
        subject_id INTEGER NOT NULL,
        attendance_date TEXT NOT NULL,
        lecture_no INTEGER NOT NULL DEFAULT 1,
        lecture_time TEXT DEFAULT '',
        session_type TEXT NOT NULL DEFAULT 'Lecture',
        batch TEXT DEFAULT '',
        status TEXT NOT NULL CHECK(status IN ('Present','Absent')),
        marked_by INTEGER,
        UNIQUE(student_id, subject_id, attendance_date, lecture_no),
        FOREIGN KEY(student_id) REFERENCES students(id) ON DELETE CASCADE,
        FOREIGN KEY(subject_id) REFERENCES subjects(id) ON DELETE CASCADE,
        FOREIGN KEY(marked_by) REFERENCES users(id) ON DELETE SET NULL
    );

    CREATE TABLE IF NOT EXISTS activity_logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER,
        action TEXT NOT NULL,
        details TEXT DEFAULT '',
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE SET NULL
    );
    CREATE TABLE IF NOT EXISTS attendance_corrections (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        attendance_id INTEGER,
        student_id INTEGER NOT NULL,
        subject_id INTEGER NOT NULL,
        attendance_date TEXT NOT NULL,
        lecture_no INTEGER NOT NULL DEFAULT 1,
        old_status TEXT,
        new_status TEXT NOT NULL,
        reason TEXT NOT NULL,
        corrected_by INTEGER,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(attendance_id) REFERENCES attendance(id) ON DELETE SET NULL,
        FOREIGN KEY(student_id) REFERENCES students(id) ON DELETE CASCADE,
        FOREIGN KEY(subject_id) REFERENCES subjects(id) ON DELETE CASCADE,
        FOREIGN KEY(corrected_by) REFERENCES users(id) ON DELETE SET NULL
    );
    CREATE TABLE IF NOT EXISTS app_settings (
        key TEXT PRIMARY KEY,
        value TEXT DEFAULT ''
    );
    CREATE TABLE IF NOT EXISTS departments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        code TEXT UNIQUE NOT NULL,
        name TEXT UNIQUE NOT NULL,
        active INTEGER NOT NULL DEFAULT 1
    );
    CREATE TABLE IF NOT EXISTS timetable (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        department_id INTEGER,
        year TEXT NOT NULL,
        day_of_week TEXT NOT NULL,
        lecture_no INTEGER NOT NULL,
        start_time TEXT NOT NULL,
        end_time TEXT NOT NULL,
        subject_id INTEGER,
        teacher_id INTEGER,
        room TEXT DEFAULT '',
        FOREIGN KEY(subject_id) REFERENCES subjects(id) ON DELETE SET NULL,
        FOREIGN KEY(teacher_id) REFERENCES users(id) ON DELETE SET NULL
    );
    CREATE TABLE IF NOT EXISTS notifications (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER NOT NULL,
        title TEXT NOT NULL,
        message TEXT NOT NULL,
        link TEXT DEFAULT '',
        is_read INTEGER NOT NULL DEFAULT 0,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE
    );
    CREATE TABLE IF NOT EXISTS leave_applications (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        student_id INTEGER NOT NULL,
        from_date TEXT NOT NULL,
        to_date TEXT NOT NULL,
        reason TEXT NOT NULL,
        status TEXT NOT NULL DEFAULT 'Pending',
        reviewed_by INTEGER,
        review_note TEXT DEFAULT '',
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(student_id) REFERENCES students(id) ON DELETE CASCADE,
        FOREIGN KEY(reviewed_by) REFERENCES users(id) ON DELETE SET NULL
    );
    CREATE TABLE IF NOT EXISTS qr_sessions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        token TEXT UNIQUE NOT NULL,
        subject_id INTEGER NOT NULL,
        year TEXT NOT NULL,
        attendance_date TEXT NOT NULL,
        lecture_no INTEGER NOT NULL,
        lecture_time TEXT DEFAULT '',
        session_type TEXT NOT NULL DEFAULT 'Lecture',
        batch TEXT DEFAULT '',
        teacher_id INTEGER NOT NULL,
        expires_at TEXT NOT NULL,
        active INTEGER NOT NULL DEFAULT 1,
        FOREIGN KEY(subject_id) REFERENCES subjects(id) ON DELETE CASCADE,
        FOREIGN KEY(teacher_id) REFERENCES users(id) ON DELETE CASCADE
    );
    """)

    # Migrations for databases created by earlier versions.
    ensure_column(db, "users", "full_name", "TEXT DEFAULT ''")
    ensure_column(db, "users", "email", "TEXT DEFAULT ''")
    ensure_column(db, "users", "mobile", "TEXT DEFAULT ''")
    ensure_column(db, "users", "security_question", "TEXT DEFAULT ''")
    ensure_column(db, "users", "security_answer", "TEXT DEFAULT ''")
    ensure_column(db, "users", "approved", "INTEGER NOT NULL DEFAULT 1")
    ensure_column(db, "users", "student_id", "INTEGER")
    ensure_column(db, "users", "created_at", "TEXT DEFAULT CURRENT_TIMESTAMP")
    ensure_column(db, "students", "prn", "TEXT DEFAULT ''")
    ensure_column(db, "students", "semester", "INTEGER DEFAULT 1")
    db.execute("UPDATE students SET semester=CASE year WHEN '1st Year' THEN 1 WHEN '2nd Year' THEN 3 WHEN '3rd Year' THEN 5 WHEN '4th Year' THEN 7 ELSE 1 END WHERE semester IS NULL OR semester<1 OR semester>8 OR (semester=1 AND year<>'1st Year')")
    ensure_column(db, "students", "department_id", "INTEGER")
    ensure_column(db, "users", "department_id", "INTEGER")
    ensure_column(db, "subjects", "department_id", "INTEGER")
    ensure_column(db, "subjects", "year", "TEXT DEFAULT '1st Year'" )
    ensure_column(db, "subjects", "semester", "INTEGER DEFAULT 1")
    ensure_column(db, "subjects", "session_type", "TEXT NOT NULL DEFAULT 'Lecture'")
    db.execute("UPDATE subjects SET semester=CASE year WHEN '1st Year' THEN 1 WHEN '2nd Year' THEN 3 WHEN '3rd Year' THEN 5 WHEN '4th Year' THEN 7 ELSE 1 END WHERE semester IS NULL OR semester<1 OR (semester=1 AND year<>'1st Year')")
    ensure_column(db, "attendance", "marked_by", "INTEGER")
    ensure_column(db, "attendance", "academic_year", "TEXT DEFAULT ''")
    ensure_column(db, "attendance", "session_type", "TEXT NOT NULL DEFAULT 'Lecture'")
    ensure_column(db, "attendance", "batch", "TEXT DEFAULT ''")
    ensure_column(db, "attendance", "scheduled_teacher_id", "INTEGER")
    ensure_column(db, "attendance", "conducted_by_teacher_id", "INTEGER")
    ensure_column(db, "attendance", "teacher_status", "TEXT DEFAULT 'Scheduled teacher present'")
    ensure_column(db, "attendance", "start_time", "TEXT DEFAULT ''")
    ensure_column(db, "attendance", "end_time", "TEXT DEFAULT ''")
    ensure_column(db, "qr_sessions", "session_type", "TEXT NOT NULL DEFAULT 'Lecture'")
    ensure_column(db, "qr_sessions", "batch", "TEXT DEFAULT ''")
    ensure_column(db, "qr_sessions", "scheduled_teacher_id", "INTEGER")
    ensure_column(db, "qr_sessions", "conducted_by_teacher_id", "INTEGER")
    ensure_column(db, "qr_sessions", "teacher_status", "TEXT DEFAULT 'Scheduled teacher present'")
    ensure_column(db, "qr_sessions", "start_time", "TEXT DEFAULT ''")
    ensure_column(db, "qr_sessions", "end_time", "TEXT DEFAULT ''")
    ensure_column(db, "timetable", "session_type", "TEXT NOT NULL DEFAULT 'Lecture'")
    ensure_column(db, "timetable", "batch", "TEXT DEFAULT ''")
    db.execute("CREATE INDEX IF NOT EXISTS idx_timetable_week ON timetable(year,day_of_week,start_time)")
    ensure_column(db, "students", "academic_year", "TEXT DEFAULT ''")
    ensure_column(db, "students", "batch", "TEXT DEFAULT ''")
    ensure_column(db, "students", "blocked", "INTEGER NOT NULL DEFAULT 0")
    ensure_column(db, "users", "totp_secret", "TEXT DEFAULT ''")
    ensure_column(db, "users", "totp_enabled", "INTEGER NOT NULL DEFAULT 0")
    ensure_column(db, "users", "totp_recovery", "TEXT DEFAULT ''")
    archive_id_def = "SERIAL PRIMARY KEY" if db.postgres else "INTEGER PRIMARY KEY AUTOINCREMENT"
    db.execute("""CREATE TABLE IF NOT EXISTS login_sessions (
        token TEXT PRIMARY KEY,
        user_id INTEGER NOT NULL,
        created_at TEXT NOT NULL,
        expires_at TEXT NOT NULL
    )""")
    db.execute("""CREATE TABLE IF NOT EXISTS login_attempts (
        key TEXT PRIMARY KEY,
        fails INTEGER NOT NULL DEFAULT 0,
        first_at TEXT,
        locked_until TEXT
    )""")
    db.execute(f"""CREATE TABLE IF NOT EXISTS email_log (
        id {archive_id_def},
        student_id INTEGER,
        kind TEXT NOT NULL,
        sent_at TEXT NOT NULL
    )""")
    db.execute(f"""CREATE TABLE IF NOT EXISTS deleted_students (
        id {archive_id_def},
        student_id INTEGER,
        name TEXT,
        prn TEXT DEFAULT '',
        roll_no TEXT DEFAULT '',
        year TEXT DEFAULT '',
        deleted_at TEXT,
        deleted_by INTEGER,
        snapshot TEXT NOT NULL
    )""")
    history_id_def = "SERIAL PRIMARY KEY" if db.postgres else "INTEGER PRIMARY KEY AUTOINCREMENT"
    db.execute(f"""CREATE TABLE IF NOT EXISTS student_academic_history (
        id {history_id_def},
        student_id INTEGER NOT NULL,
        academic_year TEXT NOT NULL,
        year TEXT NOT NULL,
        promoted_at TEXT DEFAULT CURRENT_TIMESTAMP,
        promoted_by INTEGER,
        FOREIGN KEY(student_id) REFERENCES students(id) ON DELETE CASCADE,
        FOREIGN KEY(promoted_by) REFERENCES users(id) ON DELETE SET NULL,
        UNIQUE(student_id, academic_year)
    )""")
    current_ay = db.execute("SELECT value FROM app_settings WHERE key='academic_year'").fetchone()
    current_ay_value = (current_ay["value"] if current_ay and current_ay["value"] else "2026-27")
    db.execute("INSERT INTO app_settings(key,value) VALUES(?,?) ON CONFLICT(key) DO NOTHING", ("academic_year", current_ay_value))
    db.execute("UPDATE students SET academic_year=? WHERE academic_year IS NULL OR academic_year=''", (current_ay_value,))
    db.execute("UPDATE attendance SET academic_year=? WHERE academic_year IS NULL OR academic_year=''", (current_ay_value,))
    # Multi-lecture attendance migration: old versions had one record per student/subject/day.
    # Rebuild the attendance table once so a subject can have up to 3 lectures per day.
    migrate_attendance_for_lectures(db)

    # Indexes for fast filtering/reporting when many users access the system together.
    db.execute("CREATE INDEX IF NOT EXISTS idx_students_year ON students(year)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_students_name ON students(name)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_students_batch ON students(year,batch)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_attendance_date ON attendance(attendance_date)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_attendance_subject_date ON attendance(subject_id,attendance_date)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_attendance_student_date ON attendance(student_id,attendance_date)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_attendance_academic_year ON attendance(academic_year)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_student_academic_year ON students(academic_year)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_activity_created ON activity_logs(created_at)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_notifications_user ON notifications(user_id,is_read,created_at)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_leave_status ON leave_applications(status,created_at)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_qr_token ON qr_sessions(token,active)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_timetable_day ON timetable(year,day_of_week)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_correction_date ON attendance_corrections(attendance_date)")
    # One PRN can belong to only one student when it is not blank.
    db.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_students_prn_unique ON students(prn) WHERE prn IS NOT NULL AND prn <> ''")

    if db.execute("SELECT COUNT(*) FROM departments").fetchone()[0] == 0:
        db.execute("INSERT INTO departments(code,name,active) VALUES(?,?,1)", ("PHARM", "Pharmacy"))
    default_dept = db.execute("SELECT id FROM departments ORDER BY id LIMIT 1").fetchone()[0]
    db.execute("UPDATE students SET department_id=? WHERE department_id IS NULL", (default_dept,))
    db.execute("UPDATE subjects SET department_id=? WHERE department_id IS NULL", (default_dept,))
    db.execute("UPDATE users SET department_id=? WHERE department_id IS NULL AND role IN ('teacher','hod')", (default_dept,))

    admin_username = os.environ.get("ADMIN_USERNAME", "admin").strip()
    admin_password = os.environ.get("ADMIN_PASSWORD", "")
    admin_name = os.environ.get("ADMIN_NAME", "College Administrator").strip() or "College Administrator"
    admin_email = os.environ.get("ADMIN_EMAIL", "").strip().lower()
    if db.execute("SELECT COUNT(*) FROM users WHERE role='admin'").fetchone()[0] == 0:
        if not admin_password or len(admin_password) < 12:
            db.close()
            raise RuntimeError("Set ADMIN_PASSWORD to a strong password of at least 12 characters before first startup. No default admin password is created.")
        db.execute("INSERT INTO users(username,password,role,full_name,email,approved) VALUES(?,?,?,?,?,?)",
                   (admin_username, generate_password_hash(admin_password), "admin", admin_name, admin_email, 1))
    else:
        db.execute("UPDATE users SET approved=1 WHERE role='admin'")
        if admin_email:
            db.execute("UPDATE users SET email=? WHERE role='admin' AND username=?", (admin_email, admin_username))

    # Department/HOD functionality has been removed. Convert any legacy HOD account to a regular teacher.
    db.execute("UPDATE users SET role='teacher' WHERE role='hod'")
    # Old demo teacher is intentionally removed; teachers now register.
    db.execute("DELETE FROM users WHERE role='teacher' AND username='teacher'")

    if db.execute("SELECT COUNT(*) FROM students").fetchone()[0] == 0:
        students = [
            ("01", "01", "Rahul Sharma", "B.Pharm", "1st Year", "A"),
            ("02", "02", "Priya Patel", "B.Pharm", "1st Year", "A"),
            ("03", "03", "Ankit Yadav", "B.Pharm", "1st Year", "A"),
            ("04", "04", "Sneha Gupta", "B.Pharm", "1st Year", "A"),
            ("05", "05", "Rohan Singh", "B.Pharm", "1st Year", "A"),
        ]
        db.executemany("INSERT INTO students(roll_no,prn,name,course,year,division) VALUES(?,?,?,?,?,?)", students)

    if db.execute("SELECT COUNT(*) FROM subjects").fetchone()[0] == 0:
        subjects = [
            ("BP503T", "Pharmacology-II"),
            ("BP502T", "Pharmacognosy-II"),
            ("BP501T", "Medicinal Chemistry-II"),
        ]
        db.executemany("INSERT INTO subjects(code,name,year) VALUES(?,?,?)", [(c,n,"1st Year") for c,n in subjects])

    db.commit()
    db.close()


@app.after_request
def add_security_headers(response):
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "SAMEORIGIN"
    response.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"
    response.headers["Permissions-Policy"] = "camera=(self), microphone=(), geolocation=()"
    response.headers.setdefault("Cache-Control", "no-store" if session.get("user_id") else "no-cache")
    response.headers["Cross-Origin-Opener-Policy"] = "same-origin"
    response.headers["X-Permitted-Cross-Domain-Policies"] = "none"
    if request.is_secure:
        response.headers["Strict-Transport-Security"] = "max-age=31536000; includeSubDomains"
    return response


def validate_password(password):
    if len(password) < 8:
        return "Password must be at least 8 characters."
    if not any(c.isupper() for c in password):
        return "Password must contain at least one uppercase letter."
    if not any(c.islower() for c in password):
        return "Password must contain at least one lowercase letter."
    if not any(c.isdigit() for c in password):
        return "Password must contain at least one number."
    return None


def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if "user_id" not in session:
            return redirect(url_for("login"))
        return view(*args, **kwargs)
    return wrapped


def staff_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if session.get("role") not in ("admin", "teacher"):
            flash("This page is available only to teachers and administrators.", "error")
            return redirect(url_for("student_dashboard"))
        return view(*args, **kwargs)
    return wrapped


def admin_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if session.get("role") != "admin":
            flash("Admin access required.", "error")
            return redirect(url_for("dashboard"))
        return view(*args, **kwargs)
    return wrapped


def student_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if session.get("role") != "student":
            flash("Student access required.", "error")
            return redirect(url_for("dashboard"))
        try:
            db = get_db()
            row = db.execute("SELECT blocked FROM students WHERE id=?", (session.get("student_id"),)).fetchone()
            db.close()
            if row and int(row["blocked"] or 0) == 1:
                session.clear()
                flash("You are not registered. Please contact the college administration.", "error")
                return redirect(url_for("login"))
        except Exception:
            try: db.close()
            except Exception: pass
        return view(*args, **kwargs)
    return wrapped


def teacher_or_admin_subjects(db):
    if session.get("role") == "admin":
        return db.execute("SELECT * FROM subjects ORDER BY code").fetchall()
    return db.execute("""
        SELECT s.* FROM subjects s
        JOIN subject_teachers st ON st.subject_id=s.id
        WHERE st.teacher_id=? ORDER BY s.code
    """, (session["user_id"],)).fetchall()


def can_use_subject(db, subject_id):
    if session.get("role") == "admin":
        return True
    row = db.execute("SELECT 1 FROM subject_teachers WHERE teacher_id=? AND subject_id=?",
                     (session["user_id"], subject_id)).fetchone()
    return bool(row)


def current_student(db):
    return db.execute("""
        SELECT s.* FROM students s
        JOIN users u ON u.student_id=s.id
        WHERE u.id=? AND u.role='student'
    """, (session["user_id"],)).fetchone()

def log_activity(action, details=""):
    try:
        db = get_db()
        db.execute("INSERT INTO activity_logs(user_id,action,details,created_at) VALUES(?,?,?,?)",
                   (session.get("user_id"), action, details, datetime.utcnow().isoformat()))
        db.commit(); db.close()
    except Exception:
        pass


@app.context_processor
def inject_globals():
    unread_notifications = 0
    if session.get("user_id"):
        try:
            db = get_db()
            unread_notifications = db.execute("SELECT COUNT(*) c FROM notifications WHERE user_id=? AND is_read=0", (session["user_id"],)).fetchone()["c"]
            db.close()
        except Exception:
            pass
    today_timetable=[]
    try:
        db=get_db()
        if session.get("role")=="student":
            st=current_student(db); today_timetable=get_timetable_rows(db, role="student", user_id=session.get("user_id"), student=st) if st else []
        elif session.get("role")=="teacher": today_timetable=get_timetable_rows(db, role="teacher", user_id=session.get("user_id"))
        elif session.get("role")=="admin": today_timetable=get_timetable_rows(db, role="admin", user_id=session.get("user_id"))
        db.close()
    except Exception:
        try: db.close()
        except Exception: pass
    return {
        "today_timetable": today_timetable,
        "today": today_ist().isoformat(),
        "current_user": session.get("username"),
        "college_name": COLLEGE_NAME,
        "college_tagline": COLLEGE_TAGLINE,
        "logo_file": LOGO_FILE,
        "unread_notifications": unread_notifications,
    }


@app.errorhandler(500)
def internal_error(error):
    original = getattr(error, "original_exception", None) or error
    app.logger.error("Unhandled error on %s %s: %s: %s", request.method, request.path,
                     type(original).__name__, original, exc_info=original if isinstance(original, BaseException) else None)
    page = (
        "<!doctype html><html lang='en'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'>"
        "<title>Something went wrong</title></head><body style='font-family:Arial,sans-serif;max-width:560px;margin:12vh auto;padding:0 18px;color:#173f63'>"
        "<h2>Something went wrong</h2><p>The page could not be loaded. Please try again in a minute. "
        "If it keeps happening, tell the administrator this error code:</p>"
        f"<p style='background:#f5f8fb;border:1px solid #dbe4ee;border-radius:8px;padding:10px;font-family:monospace'>{xml_escape(type(original).__name__)}</p>"
        "<p><a href='/'>Back to login</a></p></body></html>"
    )
    return Response(page, status=500, mimetype="text/html")


@app.get("/health")
def health():
    return {"status": "ok", "service": "Personal Attendance Management"}


def complete_login(user, remember=False):
    """Create the logged-in session (with a login time limit) and send the user to their dashboard."""
    role = "teacher" if user["role"] == "hod" else user["role"]  # legacy HOD accounts act as teachers
    session.clear()
    session.update(user_id=user["id"], username=user["username"], role=role,
                   full_name=user["full_name"] or user["username"], student_id=user["student_id"])
    session.permanent = bool(remember)
    ttl = REMEMBER_DAYS * 86400 if remember else SESSION_HOURS * 3600
    session["exp"] = time.time() + ttl
    session["sid"] = create_login_session(user["id"], ttl)
    if role in ("admin", "teacher"):
        log_activity_for(user["id"], "Login", f"{role} login from {client_ip()}")
    if role == "student":
        return redirect(url_for("student_dashboard"))
    return redirect(url_for("dashboard"))


@app.route("/", methods=["GET", "POST"])
@limiter.limit("10 per minute", methods=["POST"])
def login():
    if "user_id" in session:
        return redirect(url_for("dashboard"))
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        remember = request.form.get("remember") is not None
        selected_role = request.form.get("login_role", "admin").strip().lower()
        if selected_role not in ("admin", "teacher", "student"):
            selected_role = "admin"
        db = get_db()
        wait = lockout_minutes_left(db, username)
        if wait:
            db.close()
            flash(f"Too many failed attempts. Please try again in {wait} minute(s).", "error")
            return render_template("login.html", selected_role=selected_role)
        user = db.execute("SELECT * FROM users WHERE username=?", (username,)).fetchone()
        if user and check_password_hash(user["password"], password):
            account_role = "teacher" if user["role"] == "hod" else user["role"]
            if account_role != selected_role:
                db.close()
                flash(f"This is not a {selected_role.capitalize()} account. Please select the correct login type "
                      f"({account_role.capitalize()}) and try again.", "error")
                return render_template("login.html", selected_role=selected_role)
            if user["role"] == "student":
                blocked = db.execute("SELECT blocked FROM students WHERE id=?", (user["student_id"],)).fetchone() if user["student_id"] else None
                if blocked and int(blocked["blocked"] or 0) == 1:
                    db.close()
                    flash("You are not registered. Please contact the college administration.", "error")
                    return redirect(url_for("login"))
            db.close()
            if user["approved"] == 0:
                message = "Your teacher account is waiting for admin approval." if user["role"] == "teacher" else "Your student account is waiting for approval."
                flash(message, "error")
                return redirect(url_for("login"))
            if account_role == "admin" and int(user["totp_enabled"] or 0) == 1 and user["totp_secret"]:
                # Second step: ask for the authenticator code before creating the session.
                session.clear()
                session["pending_2fa"] = {"uid": int(user["id"]), "remember": remember, "at": time.time(), "tries": 0}
                return redirect(url_for("login_2fa"))
            clear_failed_logins(username)
            return complete_login(user, remember)
        db.close()
        record_failed_login(username, user["id"] if user else None)
        flash("Invalid username or password.", "error")
        return render_template("login.html", selected_role=selected_role)
    return render_template("login.html", selected_role="admin")


# ---------------- Two-step verification (TOTP, works with Google Authenticator etc.) ----------------
def totp_code(secret, for_time=None, step=30, digits=6):
    key = base64.b32decode(secret.upper() + "=" * (-len(secret) % 8))
    counter = int((for_time if for_time is not None else time.time()) // step)
    digest = hmac.new(key, struct.pack(">Q", counter), hashlib.sha1).digest()
    offset = digest[-1] & 0x0F
    value = (struct.unpack(">I", digest[offset:offset + 4])[0] & 0x7FFFFFFF) % (10 ** digits)
    return f"{value:0{digits}d}"


def totp_verify(secret, code, window=1):
    code = (code or "").strip().replace(" ", "")
    if len(code) != 6 or not code.isdigit():
        return False
    now = time.time()
    ok = False
    for offset in range(-window, window + 1):
        if hmac.compare_digest(totp_code(secret, now + offset * 30), code):
            ok = True
    return ok


def new_totp_secret():
    return base64.b32encode(os.urandom(20)).decode("ascii").rstrip("=")


def _recovery_hash(code):
    return hashlib.sha256((code or "").strip().lower().replace("-", "").replace(" ", "").encode("utf-8")).hexdigest()


def generate_recovery_codes(count=8):
    return [f"{secrets.token_hex(3)}-{secrets.token_hex(3)}" for _ in range(count)]


def use_recovery_code(db, user, code):
    """Accept a one-time recovery code (and burn it). Returns True if it was valid."""
    try:
        hashes = json.loads(user["totp_recovery"] or "[]")
    except Exception:
        hashes = []
    wanted = _recovery_hash(code)
    for h in hashes:
        if hmac.compare_digest(h, wanted):
            hashes.remove(h)
            db.execute("UPDATE users SET totp_recovery=? WHERE id=?", (json.dumps(hashes), user["id"]))
            db.commit()
            return True
    return False


@app.route("/login/2fa", methods=["GET", "POST"])
@limiter.limit("10 per minute", methods=["POST"])
def login_2fa():
    pending = session.get("pending_2fa")
    if not pending or time.time() - float(pending.get("at", 0)) > 300:
        session.pop("pending_2fa", None)
        flash("Please log in again.", "error")
        return redirect(url_for("login"))
    if request.method == "POST":
        code = request.form.get("code", "").strip()
        db = get_db()
        user = db.execute("SELECT * FROM users WHERE id=? AND role='admin'", (pending["uid"],)).fetchone()
        ok = False
        if user and int(user["totp_enabled"] or 0) == 1 and user["totp_secret"]:
            ok = totp_verify(user["totp_secret"], code) or use_recovery_code(db, user, code)
        db.close()
        if ok:
            remember = bool(pending.get("remember"))
            session.pop("pending_2fa", None)
            log_activity_for(user["id"], "Admin login", "Logged in with 2-step verification")
            return complete_login(user, remember)
        log_activity_for(pending.get("uid"), "Failed 2-step code", f"ip={client_ip()}")
        pending["tries"] = int(pending.get("tries", 0)) + 1
        session["pending_2fa"] = pending
        if pending["tries"] >= 5:
            session.pop("pending_2fa", None)
            flash("Too many wrong codes. Please log in again.", "error")
            return redirect(url_for("login"))
        flash("Invalid code. Check your authenticator app and try again.", "error")
    return render_template("login_2fa.html")


def log_activity_for(user_id, action, details=""):
    try:
        db = get_db()
        db.execute("INSERT INTO activity_logs(user_id,action,details,created_at) VALUES(?,?,?,?)",
                   (user_id, action, details, datetime.utcnow().isoformat()))
        db.commit(); db.close()
    except Exception:
        pass


@app.route("/admin/security", methods=["GET", "POST"])
@login_required
@admin_required
@limiter.limit("20 per minute", methods=["POST"])
def admin_security():
    db = get_db()
    user = db.execute("SELECT * FROM users WHERE id=?", (session["user_id"],)).fetchone()
    enabled = bool(int(user["totp_enabled"] or 0) and user["totp_secret"])
    new_codes = None
    if request.method == "POST":
        action = request.form.get("action", "")
        code = request.form.get("code", "").strip()
        if action == "start" and not enabled:
            session["totp_setup_secret"] = new_totp_secret()
        elif action == "enable" and not enabled:
            secret = session.get("totp_setup_secret")
            if secret and totp_verify(secret, code):
                new_codes = generate_recovery_codes()
                db.execute("UPDATE users SET totp_secret=?, totp_enabled=1, totp_recovery=? WHERE id=?",
                           (secret, json.dumps([_recovery_hash(c) for c in new_codes]), user["id"]))
                db.commit()
                session.pop("totp_setup_secret", None)
                enabled = True
                log_activity("2-step verification enabled", "Admin turned on 2-step verification")
            else:
                flash("That code is not correct. Scan the QR again and enter the current 6-digit code.", "error")
        elif action in ("disable", "regen") and enabled:
            password = request.form.get("password", "")
            if check_password_hash(user["password"], password) and totp_verify(user["totp_secret"], code):
                if action == "disable":
                    db.execute("UPDATE users SET totp_secret='', totp_enabled=0, totp_recovery='' WHERE id=?", (user["id"],))
                    db.commit()
                    enabled = False
                    log_activity("2-step verification disabled", "Admin turned off 2-step verification")
                    flash("2-step verification is now off.", "success")
                else:
                    new_codes = generate_recovery_codes()
                    db.execute("UPDATE users SET totp_recovery=? WHERE id=?",
                               (json.dumps([_recovery_hash(c) for c in new_codes]), user["id"]))
                    db.commit()
                    log_activity("Recovery codes regenerated", "Admin created new recovery codes")
            else:
                flash("Password or code is not correct.", "error")
    db.close()
    setup_secret = session.get("totp_setup_secret") if not enabled else None
    qr_image = None
    if setup_secret and QRCODE_AVAILABLE:
        from urllib.parse import quote
        uri = (f"otpauth://totp/{quote(COLLEGE_NAME)}:{quote(user['username'])}"
               f"?secret={setup_secret}&issuer={quote(COLLEGE_NAME)}")
        img = qrcode.make(uri)
        bio = io.BytesIO(); img.save(bio, format="PNG")
        qr_image = "data:image/png;base64," + base64.b64encode(bio.getvalue()).decode()
    return render_template("admin_security.html", enabled=enabled, setup_secret=setup_secret,
                           qr_image=qr_image, new_codes=new_codes)


@app.route("/register", methods=["GET", "POST"])
@limiter.limit("5 per minute", methods=["POST"])
def register():
    db0 = get_db(); db0.close()
    if request.method == "POST":
        full_name = request.form.get("full_name", "").strip()
        username = request.form.get("username", "").strip()
        email = request.form.get("email", "").strip().lower()
        mobile = request.form.get("mobile", "").strip()
        security_question = request.form.get("security_question", "").strip()
        security_answer = request.form.get("security_answer", "").strip()
        password = request.form.get("password", "")
        confirm = request.form.get("confirm_password", "")
        if not all([full_name, username, email, mobile, security_question, security_answer, password, confirm]):
            flash("All fields are required. Please fill every field.", "error")
            return render_template("register.html", security_questions=SECURITY_QUESTIONS)
        if security_question not in SECURITY_QUESTIONS:
            flash("Please select a valid security question.", "error")
            return render_template("register.html", security_questions=SECURITY_QUESTIONS)
        if password != confirm:
            flash("Passwords do not match.", "error")
            return render_template("register.html", security_questions=SECURITY_QUESTIONS)
        password_error = validate_password(password)
        if password_error:
            flash(password_error, "error")
            return render_template("register.html", security_questions=SECURITY_QUESTIONS)
        db = get_db()
        try:
            if db.execute("SELECT id FROM users WHERE lower(email)=lower(?)", (email,)).fetchone():
                db.close(); flash("This email is already linked to another account. Use a different email.", "error")
                return render_template("register.html", security_questions=SECURITY_QUESTIONS)
            dept_row = db.execute("SELECT id FROM departments ORDER BY id LIMIT 1").fetchone()
            department_id = dept_row["id"] if dept_row else None
            db.execute("""INSERT INTO users(username,password,role,full_name,email,mobile,security_question,security_answer,approved,department_id)
                          VALUES(?,?,?,?,?,?,?,?,0,?)""",
                       (username, generate_password_hash(password), "teacher", full_name, email, mobile,
                        security_question, generate_password_hash(security_answer.casefold()), department_id))
            db.commit(); db.close()
            flash("Teacher registration submitted. Ask the college admin to approve your account and assign subjects.", "success")
            return redirect(url_for("login"))
        except Exception as exc:
            if not isinstance(exc, sqlite3.IntegrityError) and not (POSTGRES_AVAILABLE and isinstance(exc, psycopg2.IntegrityError)):
                raise
            db.close()
            flash("Username already exists. Choose another username.", "error")
    return render_template("register.html", security_questions=SECURITY_QUESTIONS)


@app.route("/student/register", methods=["GET", "POST"])
@limiter.limit("5 per minute", methods=["POST"])
def student_register():
    """Create a pending student account; admin approval is required before login."""
    batch_options = ["Batch A", "Batch B", "Batch C", "Batch D"]
    context = dict(year_options=YEAR_OPTIONS, selected_year="", selected_semester=None,
                   semesters_by_year=SEMESTERS_BY_YEAR, batch_options=batch_options,
                   selected_batch="", security_questions=SECURITY_QUESTIONS, departments=[])
    if request.method == "POST":
        prn = request.form.get("prn", "").strip()
        name = " ".join(request.form.get("name", "").strip().split())
        username = request.form.get("username", "").strip()
        email = request.form.get("email", "").strip().lower()
        mobile = request.form.get("mobile", "").strip()
        security_question = request.form.get("security_question", "").strip()
        security_answer = request.form.get("security_answer", "").strip()
        password = request.form.get("password", "")
        confirm = request.form.get("confirm_password", "")
        year = request.form.get("year", "").strip()
        semester = request.form.get("semester", type=int)
        batch = request.form.get("batch", "").strip()
        context.update(selected_year=year, selected_semester=semester, selected_batch=batch)
        if not all([prn, name, email, username, mobile, security_question, security_answer, password, confirm, year, semester, batch]):
            flash("All fields are required. Please fill every field and select your year and batch.", "error")
            return render_template("student_register.html", **context)
        if year not in YEAR_OPTIONS or semester not in SEMESTERS_BY_YEAR.get(year, []):
            flash("Please select a valid year and semester.", "error")
            return render_template("student_register.html", **context)
        if batch not in batch_options or security_question not in SECURITY_QUESTIONS:
            flash("Please select a valid batch and security question.", "error")
            return render_template("student_register.html", **context)
        if password != confirm:
            flash("Passwords do not match.", "error")
            return render_template("student_register.html", **context)
        password_error = validate_password(password)
        if password_error:
            flash(password_error, "error")
            return render_template("student_register.html", **context)
        db = get_db()
        created_student_id = None
        try:
            if db.execute("SELECT id FROM users WHERE lower(email)=lower(?)", (email,)).fetchone():
                db.close(); flash("This email is already linked to an account. Use a different email.", "error")
                return render_template("student_register.html", **context)
            if db.execute("SELECT id FROM users WHERE lower(username)=lower(?)", (username,)).fetchone():
                db.close(); flash("This username is already in use. Choose another username.", "error")
                return render_template("student_register.html", **context)
            student = db.execute("SELECT * FROM students WHERE prn=? LIMIT 1", (prn,)).fetchone()
            if student:
                linked = db.execute("SELECT id FROM users WHERE student_id=? LIMIT 1", (student["id"],)).fetchone()
                if linked:
                    db.close(); flash("An account for this PRN already exists or is awaiting approval. Please contact the admin.", "error")
                    return render_template("student_register.html", **context)
                existing_name = " ".join(str(student["name"] or "").strip().split())
                if (existing_name.casefold() != name.casefold() or str(student["year"] or "").strip() != year
                    or int(student["semester"] or 1) != semester or str(student["batch"] or "").strip() != batch):
                    db.close(); flash("These details do not match the college record. Please contact the admin to verify your PRN and details.", "error")
                    return render_template("student_register.html", **context)
            else:
                # Unknown PRNs may submit a request, but the resulting student record is blocked
                # and cannot be used until an administrator verifies and approves it.
                roll_no = "PENDING-" + prn
                db.execute("""INSERT INTO students(roll_no,prn,name,course,year,semester,division,batch,blocked)
                              VALUES(?,?,?,'B.Pharm',?,?, 'A', ?,1)""",
                           (roll_no, prn, name, year, semester, batch))
                created_student_id = db.execute("SELECT id FROM students WHERE prn=? ORDER BY id DESC LIMIT 1", (prn,)).fetchone()["id"]
                student = db.execute("SELECT * FROM students WHERE id=?", (created_student_id,)).fetchone()
            department_id = student["department_id"] if student["department_id"] is not None else None
            db.execute("""INSERT INTO users(username,password,role,full_name,email,mobile,security_question,security_answer,approved,student_id,department_id)
                          VALUES(?,?,?,?,?,?,?,?,0,?,?)""",
                       (username, generate_password_hash(password), "student", name, email, mobile,
                        security_question, generate_password_hash(security_answer.casefold()), student["id"], department_id))
            db.commit(); db.close()
            flash("Registration submitted. You cannot log in until the admin verifies and approves your account.", "success")
            return redirect(url_for("login"))
        except Exception as exc:
            try:
                db.rollback()
                if created_student_id:
                    db.execute("DELETE FROM students WHERE id=? AND blocked=1", (created_student_id,))
                    db.commit()
                db.close()
            except Exception:
                pass
            if not isinstance(exc, sqlite3.IntegrityError) and not (POSTGRES_AVAILABLE and isinstance(exc, psycopg2.IntegrityError)):
                raise
            flash("PRN, username, email or account details already exist. Please check your details and try again.", "error")
    return render_template("student_register.html", **context)

def _password_fingerprint(password_hash):
    return hashlib.sha256((password_hash or "").encode("utf-8")).hexdigest()[:16]


def _user_from_reset_token(token):
    """Return the user for a valid, unused reset token, otherwise None (bad/expired/used link).
    Real server problems (for example the database) are NOT hidden: they are logged and raised."""
    try:
        data = reset_serializer().loads(token, max_age=RESET_TOKEN_MAX_AGE)
        uid = int(data["uid"])
        fingerprint = str(data["fp"])
    except (BadData, KeyError, TypeError, ValueError, AttributeError):
        return None
    db = get_db()
    try:
        user = db.execute("SELECT id,password,approved FROM users WHERE id=?", (uid,)).fetchone()
    finally:
        db.close()
    if not user or not user["approved"]:
        return None
    if not hmac.compare_digest(_password_fingerprint(user["password"]), fingerprint):
        return None  # password already changed: link was used before
    return user


@app.route("/forgot-password", methods=["GET", "POST"])
@limiter.limit("5 per minute", methods=["POST"])
@limiter.limit("20 per hour", methods=["POST"])
def forgot_password():
    if request.method == "GET":
        session.pop("_flashes", None)
        return render_template("forgot_password.html")
    identifier = request.form.get("identifier", "").strip()
    email = request.form.get("email", "").strip().lower()
    if not identifier or not email or "@" not in email:
        flash("Please enter your PRN / username and your registered email address.", "error")
        return render_template("forgot_password.html")
    if not email_is_configured():
        flash("Password reset by email is not available right now. Please contact the administrator.", "error")
        return render_template("forgot_password.html")
    db = get_db()
    user = db.execute("""SELECT id,full_name,email,password,approved FROM users
                         WHERE lower(username)=lower(?) AND lower(email)=? LIMIT 1""", (identifier, email)).fetchone()
    if not user:
        user = db.execute("""SELECT u.id,u.full_name,u.email,u.password,u.approved FROM users u
                             JOIN students s ON s.id=u.student_id
                             WHERE s.prn=? AND lower(u.email)=? AND u.role='student' LIMIT 1""", (identifier, email)).fetchone()
    db.close()
    if user and user["approved"]:
        token = reset_serializer().dumps({"uid": int(user["id"]), "fp": _password_fingerprint(user["password"])})
        link = url_for("reset_password", token=token, _external=True)
        try:
            send_password_reset_email(user["email"], user["full_name"], link)
        except Exception as exc:
            app.logger.warning("Password reset email failed: %s", exc)
    # Same answer whether or not the account exists, so nobody can test which accounts exist.
    flash("If these details match an account, a password reset link has been sent to the registered email. "
          "The link is valid for 30 minutes. Please also check your spam folder.", "success")
    return render_template("forgot_password.html", sent=True)


@app.route("/reset-password/<token>", methods=["GET", "POST"])
@limiter.limit("10 per minute", methods=["POST"])
def reset_password(token):
    user = _user_from_reset_token(token)
    if not user:
        flash("This reset link is invalid or has expired. Please request a new one.", "error")
        return redirect(url_for("forgot_password"))
    if request.method == "POST":
        password = request.form.get("password", "")
        confirm = request.form.get("confirm_password", "")
        password_error = validate_password(password)
        if password_error:
            flash(password_error, "error")
            return render_template("reset_password.html", token=token)
        if password != confirm:
            flash("Passwords do not match.", "error")
            return render_template("reset_password.html", token=token)
        db = get_db()
        db.execute("UPDATE users SET password=? WHERE id=?", (generate_password_hash(password), user["id"]))
        db.commit(); db.close()
        revoke_login_sessions(user["id"])
        log_activity_for(user["id"], "Password reset", "Password changed with an email reset link")
        flash("Password changed successfully. You can now log in with your new password.", "success")
        return redirect(url_for("login"))
    return render_template("reset_password.html", token=token)


@app.route("/logout", methods=["POST"])
def logout():
    """Log out only through a CSRF-protected POST request. The server-side login record is deleted too."""
    sid = session.get("sid")
    if sid:
        db = get_db()
        try:
            db.execute("DELETE FROM login_sessions WHERE token=?", (sid,))
            db.commit()
        finally:
            db.close()
    session.clear()
    return redirect(url_for("login"))


def get_timetable_rows(db, role=None, user_id=None, student=None, day_name=None):
    day_name = day_name or today_ist().strftime("%A")
    where=["t.day_of_week=?"]; params=[day_name]
    if student is not None:
        where.append("t.year=?"); params.append(student["year"])
        where.append("(COALESCE(t.session_type,'Lecture')='Lecture' OR COALESCE(t.batch,'')=?)"); params.append(student["batch"] or "")
    elif role == "teacher":
        where.append("t.teacher_id=?"); params.append(user_id)
    return db.execute(f"""SELECT t.*, COALESCE(t.session_type,'Lecture') session_type, COALESCE(t.batch,'') batch,
        COALESCE(sub.code,'') subject_code, COALESCE(sub.name,'') subject_name,
        COALESCE(u.full_name,u.username,'') teacher_name
        FROM timetable t LEFT JOIN subjects sub ON sub.id=t.subject_id LEFT JOIN users u ON u.id=t.teacher_id
        WHERE {' AND '.join(where)} ORDER BY t.start_time,t.lecture_no,t.id""", tuple(params)).fetchall()


def timetable_candidates(db, year, session_type, batch, subject_id, day_name):
    return db.execute("""SELECT t.*, COALESCE(u.full_name,u.username,'') teacher_name,
        COALESCE(sub.code,'') subject_code, COALESCE(sub.name,'') subject_name
        FROM timetable t LEFT JOIN users u ON u.id=t.teacher_id LEFT JOIN subjects sub ON sub.id=t.subject_id
        WHERE t.year=? AND t.day_of_week=? AND t.subject_id=?
          AND COALESCE(t.session_type,'Lecture')=?
          AND (COALESCE(t.batch,'')=? OR ?='Lecture')
        ORDER BY t.start_time,t.lecture_no,t.id""",
        (year,day_name,subject_id,session_type,batch if session_type=='Practical' else '',session_type)).fetchall()


def current_academic_year_value(db):
    row=db.execute("SELECT value FROM app_settings WHERE key='academic_year'").fetchone()
    return row['value'] if row and row['value'] else '2026-27'


def _time_minutes(value):
    try:
        h, m = [int(x) for x in str(value or '').split(':', 1)]
        return h * 60 + m
    except Exception:
        return 0

def _format_time_12(value):
    try:
        h, m = [int(x) for x in str(value or '').split(':', 1)]
        suffix = 'am' if h < 12 else 'pm'
        hh = h % 12 or 12
        return f'{hh:02d}:{m:02d} {suffix}'
    except Exception:
        return str(value or '')

def _year_heading(year):
    return {
        '1st Year':'FIRST YEAR', '2nd Year':'SECOND YEAR',
        '3rd Year':'THIRD YEAR', '4th Year':'FOURTH YEAR'
    }.get(year, str(year or 'ALL YEARS').upper())

def all_teacher_rows(db):
    return db.execute("SELECT id,COALESCE(full_name,username) name FROM users WHERE role='teacher' AND approved=1 ORDER BY name").fetchall()


def timetable_scope_sql(role, user_id, student):
    where=[]; params=[]
    if role=='teacher':
        where.append('t.teacher_id=?'); params.append(user_id)
    elif role=='student':
        where += ['t.year=?', "(COALESCE(t.session_type,'Lecture')='Lecture' OR COALESCE(t.batch,'')=?)"]
        params += [student['year'], student['batch'] or '']
    return where, params


def _timetable_scope(role, user_id, student, year, session_type, batch, semester=0):
    where=[]; params=[]
    if role == 'teacher':
        where.append('t.teacher_id=?'); params.append(user_id)
    elif role == 'student':
        if not student:
            where.append('1=0')
        else:
            where += ['t.year=?', "(COALESCE(t.session_type,'Lecture')='Lecture' OR COALESCE(t.batch,'')=?)"]
            params += [student['year'], student['batch'] or '']
    if role in ('admin','teacher') and year:
        where.append('t.year=?'); params.append(year)
    if session_type in ('Lecture','Practical'):
        where.append("COALESCE(t.session_type,'Lecture')=?"); params.append(session_type)
    if semester:
        where.append('t.subject_id IN (SELECT id FROM subjects WHERE semester=?)'); params.append(semester)
    if session_type == 'Practical' and batch in ('Batch A','Batch B','Batch C','Batch D'):
        where.append("COALESCE(t.batch,'')=?"); params.append(batch)
    return where, params

def _load_timetable_rows(db, role, user_id, student, year='', day='', view='week', session_type='', batch='', semester=0):
    where,params=_timetable_scope(role,user_id,student,year,session_type,batch,semester)
    if view == 'day':
        where.append('t.day_of_week=?'); params.append(day)
    condition=' AND '.join(where) if where else '1=1'
    return db.execute(f"""SELECT t.*, COALESCE(t.session_type,'Lecture') session_type,
        COALESCE(t.batch,'') batch, COALESCE(sub.code,'') subject_code,
        COALESCE(sub.name,'') subject_name, COALESCE(u.full_name,u.username,'') teacher_name
        FROM timetable t
        LEFT JOIN subjects sub ON sub.id=t.subject_id
        LEFT JOIN users u ON u.id=t.teacher_id
        WHERE {condition}
        ORDER BY CASE t.day_of_week WHEN 'Monday' THEN 1 WHEN 'Tuesday' THEN 2 WHEN 'Wednesday' THEN 3 WHEN 'Thursday' THEN 4 WHEN 'Friday' THEN 5 WHEN 'Saturday' THEN 6 ELSE 7 END,
                 t.start_time,t.end_time,t.year,t.id""", tuple(params)).fetchall()

def _timetable_slot_rows(rows, days):
    """Return the timetable's elementary time intervals.

    Every distinct start/end boundary becomes part of one shared timeline.
    This lets a 09:30-12:30 practical span across several 45-minute lecture
    rows on another day, while each lecture keeps its exact duration.
    """
    boundaries=set()
    for r in rows:
        boundaries.add(r['start_time'])
        boundaries.add(r['end_time'])
    ordered=sorted(boundaries, key=_time_minutes)
    return [(ordered[i], ordered[i+1]) for i in range(len(ordered)-1)
            if _time_minutes(ordered[i]) < _time_minutes(ordered[i+1])]

def _build_timetable_grid(rows, days):
    """Build merged cells for long practicals and exact-duration lectures."""
    intervals=_timetable_slot_rows(rows, days)
    # Merge practical entries for the same slot/subject/teacher/room.
    grouped={}
    for r in rows:
        key=(r['day_of_week'],r['start_time'],r['end_time'],r['subject_id'],r['teacher_id'],r['room'] or '',r['session_type'])
        grouped.setdefault(key,[]).append(r)

    entries_by_day={d:[] for d in days}
    for key,items in grouped.items():
        day,st,en,subject_id,teacher_id,room,stype=key
        first=items[0]
        # sqlite Row is immutable, so copy the fields needed by the template.
        data=dict(first)
        if stype=='Practical':
            batches=sorted({(x['batch'] or '').strip() for x in items if (x['batch'] or '').strip()})
            data['batch_display']='All Batches' if set(batches)=={'Batch A','Batch B','Batch C','Batch D'} else (', '.join(batches) if batches else 'All Batches')
        else:
            data['batch_display']=''
        entries_by_day[day].append(data)

    grid=[]
    for st,en in intervals:
        row={'start_time':st,'end_time':en,
             'display_start':_format_time_12(st),'display_end':_format_time_12(en),
             'lunch_before':False,'cells':{}}
        for d in days:
            candidates=[]
            for e in entries_by_day[d]:
                if _time_minutes(e['start_time']) <= _time_minutes(st) and _time_minutes(e['end_time']) >= _time_minutes(en):
                    candidates.append(e)
            row['cells'][d]={'skip':False,'rowspan':1,'entries':[]}
            # Only render an entry in the first interval it occupies.
            for e in candidates:
                if e['start_time']==st:
                    row['cells'][d]['entries'].append(e)
                    span=sum(1 for a,b in intervals
                             if _time_minutes(a)>=_time_minutes(st) and _time_minutes(b)<=_time_minutes(e['end_time']))
                    row['cells'][d]['rowspan']=max(row['cells'][d]['rowspan'],span)
                elif e['start_time'] < st:
                    row['cells'][d]['skip']=True
        grid.append(row)

    # Mark lunch gap between 12:30 and 13:00 when both sides of the gap exist.
    for i,row in enumerate(grid):
        if i and _time_minutes(row['start_time'])-_time_minutes(grid[i-1]['end_time'])>=30:
            row['lunch_before']=True
    return grid

@app.route('/timetable', methods=['GET','POST'])
@login_required
def timetable():
    db=get_db(); role=session.get('role')
    days=['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday']
    batches=['Batch A','Batch B','Batch C','Batch D']
    if role not in ('admin','teacher','student'):
        db.close(); flash('Timetable is available only for Admin, Teacher and Student accounts.','error'); return redirect(url_for('dashboard'))

    if request.method=='POST':
        if role!='admin':
            db.close(); flash('Only administrators can manage the timetable.','error'); return redirect(url_for('timetable'))
        year=request.form.get('year','').strip(); day=request.form.get('day_of_week','').strip()
        start=request.form.get('start_time','').strip(); end=request.form.get('end_time','').strip()
        subject_id=request.form.get('subject_id',type=int); teacher_id=request.form.get('teacher_id',type=int)
        session_type=request.form.get('session_type','Lecture').strip(); batch=request.form.get('batch','').strip(); room=request.form.get('room','').strip()
        subject_row=db.execute("SELECT id,year,COALESCE(session_type,'Lecture') session_type FROM subjects WHERE id=?",(subject_id,)).fetchone() if subject_id else None
        if year not in YEAR_OPTIONS or day not in days or not start or not end or start>=end or not subject_id or not teacher_id or session_type not in ('Lecture','Practical'):
            flash('Please fill all timetable fields correctly. Start time must be before end time.','error')
        elif not subject_row or subject_row['year']!=year or subject_row['session_type']!=session_type:
            flash('Please select a subject belonging to the selected year and session type.','error')
        elif session_type=='Practical' and batch not in batches:
            flash('Select a valid practical batch.','error')
        else:
            # Department is intentionally not used by the timetable UI or logic.
            db.execute("INSERT INTO timetable(department_id,year,day_of_week,lecture_no,start_time,end_time,subject_id,teacher_id,room,session_type,batch) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                       (None,year,day,1,start,end,subject_id,teacher_id,room,session_type,batch if session_type=='Practical' else ''))
            db.commit(); log_activity('Timetable entry added',f'{year} | {day} | {start}-{end} | {subject_id} | {session_type} {batch}')
            flash('Timetable entry added successfully.','success')

    selected_year=request.args.get('year','').strip()
    selected_semester=request.args.get('semester',type=int) or 0
    if selected_semester not in SEMESTERS_BY_YEAR.get(selected_year, []): selected_semester=0
    selected_day=request.args.get('day','').strip()
    view=request.args.get('view','week').strip().lower()
    session_type=request.args.get('session_type','').strip()
    batch=request.args.get('batch','').strip()
    if view not in ('day','week'): view='week'
    if selected_day not in days: selected_day=today_ist().strftime('%A')
    if role=='admin' and not selected_year: selected_year='1st Year'
    student=current_student(db) if role=='student' else None
    if role=='student' and student: selected_year=student['year']; batch=student['batch'] or ''
    rows=_load_timetable_rows(db,role,session.get('user_id'),student,selected_year,selected_day,view,session_type,batch,selected_semester)
    subjects=db.execute("SELECT id,code,name,year,COALESCE(semester,1) semester,COALESCE(session_type,'Lecture') session_type FROM subjects ORDER BY CASE year WHEN '1st Year' THEN 1 WHEN '2nd Year' THEN 2 WHEN '3rd Year' THEN 3 ELSE 4 END,semester,session_type,code").fetchall()
    teachers=all_teacher_rows(db)
    slots=_timetable_slot_rows(rows,days)
    grid_rows=_build_timetable_grid(rows,days)
    subject_details=[]
    seen_subjects=set()
    for r in rows:
        key=(r['subject_code'],r['subject_name'],r['teacher_name'])
        if key not in seen_subjects:
            seen_subjects.add(key); subject_details.append({'code':key[0],'name':key[1],'teacher':key[2]})
    db.close()
    return render_template('timetable.html',rows=rows,subjects=subjects,teachers=teachers,years=YEAR_OPTIONS,selected_year=selected_year,selected_semester=selected_semester,semesters_by_year=SEMESTERS_BY_YEAR,selected_day=selected_day,days=days,batch_options=batches,view=view,session_type=session_type,batch=batch,slots=slots,grid_rows=grid_rows,subject_details=subject_details)

@app.route('/timetable/export.pdf')
@login_required
def export_timetable_pdf():
    db=get_db(); role=session.get('role')
    days=['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday']
    selected_day=request.args.get('day','').strip()
    if selected_day not in days: selected_day=today_ist().strftime('%A')
    selected_year=request.args.get('year','').strip()
    selected_semester=request.args.get('semester',type=int) or 0
    if selected_semester not in SEMESTERS_BY_YEAR.get(selected_year, []): selected_semester=0
    view=request.args.get('view','week').strip().lower()
    if view not in ('day','week'): view='week'
    session_type=request.args.get('session_type','').strip(); batch=request.args.get('batch','').strip()
    if role=='admin' and not selected_year: selected_year='1st Year'
    student=current_student(db) if role=='student' else None
    if role=='student':
        if not student:
            db.close(); return Response('Student record not found.',status=404)
        selected_year=student['year']; batch=student['batch'] or ''
    if role not in ('admin','teacher','student'):
        db.close(); return Response('Not authorized.',status=403)
    rows=_load_timetable_rows(db,role,session.get('user_id'),student,selected_year,selected_day,view,session_type,batch,selected_semester)
    ayrow=db.execute("SELECT value FROM app_settings WHERE key='academic_year'").fetchone()
    academic_year=(ayrow['value'] if ayrow and ayrow['value'] else '2026-27')
    # Unique subject/teacher/room legend.
    legend=[]; seen=set()
    for r in rows:
        key=(r['subject_code'],r['subject_name'],r['teacher_name'],r['room'] or '')
        if key not in seen:
            seen.add(key); legend.append(key)
    db.close()

    buffer=io.BytesIO()
    doc=SimpleDocTemplate(buffer,pagesize=landscape(A4),rightMargin=7*mm,leftMargin=7*mm,topMargin=7*mm,bottomMargin=10*mm)
    styles=getSampleStyleSheet()
    system=ParagraphStyle('TTSystem',parent=styles['Title'],alignment=TA_CENTER,fontName='Helvetica-Bold',fontSize=18,leading=20,textColor=colors.HexColor('#123f73'),spaceAfter=1)
    timetable_title=ParagraphStyle('TTTitle',parent=styles['Title'],alignment=TA_CENTER,fontName='Helvetica-Bold',fontSize=14,leading=16,textColor=colors.HexColor('#07579a'),spaceAfter=0)
    meta=ParagraphStyle('TTMeta',parent=styles['Normal'],alignment=TA_CENTER,fontName='Helvetica-Bold',fontSize=8.5,leading=10,textColor=colors.HexColor('#244f73'))
    cell=ParagraphStyle('TTCell',parent=styles['Normal'],alignment=TA_CENTER,fontSize=7.2,leading=8.4,textColor=colors.HexColor('#172f49'))
    head=ParagraphStyle('TTHead',parent=cell,fontName='Helvetica-Bold',fontSize=8.3,textColor=colors.white)
    left=ParagraphStyle('TTLeft',parent=cell,alignment=TA_LEFT)
    note=ParagraphStyle('TTNote',parent=styles['Normal'],fontSize=7.5,leading=9,textColor=colors.HexColor('#263b52'))

    story=[]
    logo_path=BASE_DIR/'static'/LOGO_FILE
    logo=Image(str(logo_path),width=21*mm,height=21*mm) if logo_path.exists() else ''
    title_row=Table([[logo,Paragraph('PERSONAL ATTENDANCE MANAGEMENT SYSTEM',system),Paragraph('TIMETABLE',timetable_title)]],colWidths=[28*mm,176*mm,60*mm])
    title_row.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'MIDDLE'),('ALIGN',(0,0),(-1,-1),'CENTER'),('BOX',(0,0),(-1,-1),0.8,colors.HexColor('#9dbbd1')),('BACKGROUND',(0,0),(-1,-1),colors.white),('LEFTPADDING',(0,0),(-1,-1),4),('RIGHTPADDING',(0,0),(-1,-1),4),('TOPPADDING',(0,0),(-1,-1),4),('BOTTOMPADDING',(0,0),(-1,-1),4)]))
    story += [title_row,Spacer(1,2*mm)]
    scope_year=_year_heading(selected_year or 'All Years')
    scope=f'{scope_year} B. PHARMACY  |  A.Y.: {academic_year}'
    if view=='day': scope += f'  |  {selected_day.upper()}'
    scope_box=Table([[Paragraph(scope,meta)]],colWidths=[264*mm])
    scope_box.setStyle(TableStyle([('BOX',(0,0),(-1,-1),0.6,colors.HexColor('#b7cfdf')),('BACKGROUND',(0,0),(-1,-1),colors.HexColor('#eef7fc')),('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),5)]))
    story += [scope_box,Spacer(1,1.8*mm)]

    grid=_build_timetable_grid(rows,days)
    day_list=[selected_day] if view=='day' else days
    span_commands=[]
    if grid:
        data=[[Paragraph('DAY & TIME',head)]+[Paragraph(d[:4].upper() if d=='Thursday' else d[:3].upper(),head) for d in day_list]]
        tr=1
        for g in grid:
            if g['lunch_before']:
                data.append([Paragraph('LUNCH BREAK',meta)]+[Paragraph('LUNCH BREAK',meta) for _ in day_list])
                tr += 1
            cells=[Paragraph(f'<b>{g["display_start"]}</b> – <b>{g["display_end"]}</b>',left)]
            for ci,d in enumerate(day_list, start=1):
                c=g['cells'][d]
                if c['skip']:
                    cells.append('')
                    continue
                bits=[]
                for r in c['entries']:
                    # PDF timetable intentionally shows only subject name.
                    # Practical sessions additionally show their batch; teacher, subject code, and room are omitted.
                    txt=f'<b>{r["subject_name"]}</b>'
                    if r['session_type']=='Practical':
                        txt += f'<br/><b>Batch: {r.get("batch_display") or r.get("batch") or "All Batches"}</b>'
                    bits.append(txt)
                cells.append(Paragraph('<br/><br/>'.join(bits) if bits else '—',cell))
                if c['rowspan']>1:
                    span_commands.append(('SPAN',(ci,tr),(ci,tr+c['rowspan']-1)))
            data.append(cells)
            tr += 1
        table=Table(data,repeatRows=1,colWidths=[42*mm]+[(222/len(day_list))*mm]*len(day_list))
    else:
        table=Table([[Paragraph('No timetable entries found for the selected filters.',cell)]],colWidths=[264*mm])
    # Rebuild rowspan-aware table cleanly because Table accepts tuple cell values.
    # If tuple values were injected above, normalize them as (flowables, span).
    table.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#0876c9')),('BOX',(0,0),(-1,-1),0.65,colors.HexColor('#7ea8c9')),('INNERGRID',(0,0),(-1,-1),0.35,colors.HexColor('#9ab9d2')),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('ALIGN',(0,0),(-1,-1),'CENTER'),('TOPPADDING',(0,0),(-1,-1),4.2),('BOTTOMPADDING',(0,0),(-1,-1),4.2),('LEFTPADDING',(0,0),(-1,-1),3.5),('RIGHTPADDING',(0,0),(-1,-1),3.5)] + span_commands))
    story += [table,Spacer(1,3*mm)]
    def page_frame(canvas,doc_obj):
        canvas.saveState(); w,h=landscape(A4)
        canvas.setStrokeColor(colors.HexColor('#1a4f87')); canvas.setLineWidth(1.0); canvas.rect(5*mm,5*mm,w-10*mm,h-10*mm,fill=0,stroke=1)
        canvas.setFillColor(colors.HexColor('#084b83')); canvas.setFont('Helvetica',7.5); canvas.drawRightString(w-9*mm,7*mm,f'Page {doc_obj.page}')
        canvas.restoreState()
    doc.build(story,onFirstPage=page_frame,onLaterPages=page_frame)
    pdf=buffer.getvalue(); buffer.close()
    fname=f"timetable_{(selected_year or 'all').replace(' ','_')}_{'week' if view=='week' else selected_day.lower()}.pdf"
    return Response(pdf,mimetype='application/pdf',headers={'Content-Disposition':f'attachment; filename={fname}'})

@app.post('/timetable/delete/<int:timetable_id>')
@login_required
@admin_required
def delete_timetable(timetable_id):
    db=get_db(); db.execute('DELETE FROM timetable WHERE id=?',(timetable_id,)); db.commit(); db.close(); flash('Timetable entry deleted.','success'); return redirect(url_for('timetable'))

@app.route("/dashboard")
@login_required
def dashboard():
    if session.get("role") == "student":
        return redirect(url_for("student_dashboard"))
    db = get_db()
    today = today_ist().isoformat()
    total = db.execute("SELECT COUNT(*) c FROM students").fetchone()["c"]
    pending = db.execute("SELECT COUNT(*) c FROM users WHERE role='teacher' AND approved=0").fetchone()["c"] if session.get("role") == "admin" else 0

    subject_filter = "" if session.get("role") == "admin" else " AND a.subject_id IN (SELECT subject_id FROM subject_teachers WHERE teacher_id=?)"
    base_params = [] if session.get("role") == "admin" else [session["user_id"]]
    today_rows = db.execute(f"""SELECT a.student_id,a.status,a.lecture_no,s.year
        FROM attendance a JOIN students s ON s.id=a.student_id
        WHERE a.attendance_date=?{subject_filter}""", [today] + base_params).fetchall()
    present_marks = sum(1 for r in today_rows if r["status"] == "Present")
    total_marks = len(today_rows)
    absent_marks = total_marks - present_marks
    today_pct = round(100.0 * present_marks / total_marks, 1) if total_marks else 0

    year_stats=[]
    for y in YEAR_OPTIONS:
        yr_students = db.execute("SELECT COUNT(*) c FROM students WHERE year=?", (y,)).fetchone()["c"]
        rows = [r for r in today_rows if r["year"] == y]
        p = sum(1 for r in rows if r["status"] == "Present")
        a = sum(1 for r in rows if r["status"] == "Absent")
        marked_students = {r["student_id"] for r in rows}
        # A student is counted as present if present for every marked lecture; otherwise absent.
        per_student = {}
        for r in rows:
            per_student.setdefault(r["student_id"], []).append(r["status"])
        present_students = sum(1 for sid, sts in per_student.items() if sts and all(x == "Present" for x in sts))
        absent_students = sum(1 for sid, sts in per_student.items() if any(x == "Absent" for x in sts))
        pct = round(100.0 * p / (p+a), 1) if (p+a) else 0
        year_stats.append({"year":y,"total":yr_students,"present":present_students,"absent":absent_students,"pct":pct,"marked":len(marked_students),"present_marks":p,"absent_marks":a})

    threshold_row = db.execute("SELECT value FROM app_settings WHERE key='attendance_threshold'").fetchone()
    try: threshold = float(threshold_row["value"] if threshold_row and threshold_row["value"] else 75)
    except Exception: threshold = 75
    low = db.execute("""SELECT s.prn,s.name,s.course,s.year,
        CASE WHEN COUNT(a.id)=0 THEN 0 ELSE ROUND(100.0*SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END)/COUNT(a.id),1) END pct
        FROM students s LEFT JOIN attendance a ON a.student_id=s.id
        GROUP BY s.id HAVING (CASE WHEN COUNT(a.id)=0 THEN 0 ELSE ROUND(100.0*SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END)/COUNT(a.id),1) END) < ?
        ORDER BY pct ASC, s.name LIMIT 10""", (threshold,)).fetchall()
    # Overall attendance uses lecture marks, so multiple lectures are counted correctly.
    avg = db.execute("""SELECT COALESCE(AVG(pct),0) avg_pct FROM (
        SELECT s.id, CASE WHEN COUNT(a.id)=0 THEN 0 ELSE 100.0*SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END)/COUNT(a.id) END pct
        FROM students s LEFT JOIN attendance a ON a.student_id=s.id GROUP BY s.id)""").fetchone()["avg_pct"]
    if session.get("role") == "admin":
        subject_count = db.execute("SELECT COUNT(*) c FROM subjects").fetchone()["c"]
        teacher_count = db.execute("SELECT COUNT(*) c FROM users WHERE role='teacher' AND approved=1").fetchone()["c"]
    else:
        subject_count = db.execute("SELECT COUNT(*) c FROM subject_teachers WHERE teacher_id=?", (session["user_id"],)).fetchone()["c"]
        teacher_count = 0

    recent_attendance = db.execute("""SELECT a.attendance_date, sub.code, sub.name subject_name, s.year,
        COUNT(a.id) total,
        SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END) present,
        SUM(CASE WHEN a.status='Absent' THEN 1 ELSE 0 END) absent
        FROM attendance a JOIN subjects sub ON sub.id=a.subject_id JOIN students s ON s.id=a.student_id
        WHERE 1=1
        """ + (" AND a.subject_id IN (SELECT subject_id FROM subject_teachers WHERE teacher_id=?) " if session.get('role')=='teacher' else "") + """GROUP BY a.attendance_date, sub.id, sub.code, sub.name, s.year
        ORDER BY a.attendance_date DESC LIMIT 6""", ([session["user_id"]] if session.get('role')=='teacher' else [])).fetchall()
    db.close()
    return render_template("dashboard.html", total=total, present=present_marks, absent=absent_marks,
        avg=round(avg or 0,1), low=low, pending=pending, year_stats=year_stats, today=today, today_pct=today_pct, today_marks=total_marks, threshold=threshold, subject_count=subject_count, teacher_count=teacher_count, recent_attendance=recent_attendance)


@app.route("/student/dashboard")
@student_required
def student_dashboard():
    db = get_db()
    student = current_student(db)
    if not student:
        db.close()
        session.clear()
        flash("Student account is not linked correctly. Contact the administrator.", "error")
        return redirect(url_for("login"))

    today = today_ist().isoformat()
    selected_date = request.args.get("date", "").strip()
    selected_subject = request.args.get("subject", "").strip()

    today_rows = db.execute("""SELECT sub.code, sub.name subject_name, COALESCE(sub.session_type,'Lecture') session_type, COALESCE(sub.semester,1) semester, a.status,
        COALESCE(a.lecture_no,1) lecture_no, COALESCE(a.lecture_time,'') lecture_time
        FROM subjects sub LEFT JOIN attendance a
        ON a.subject_id=sub.id AND a.student_id=? AND a.attendance_date=?
        WHERE sub.department_id=? AND sub.year=? AND sub.semester=?
        ORDER BY sub.code, a.lecture_no""", (student["id"], today, student["department_id"], student["year"], student["semester"])).fetchall()

    subjects = db.execute("SELECT id, code, name, COALESCE(session_type,'Lecture') session_type, COALESCE(semester,1) semester FROM subjects WHERE department_id=? AND year=? AND semester=? ORDER BY session_type, code", (student["department_id"], student["year"], student["semester"])).fetchall()

    history_sql = """SELECT a.attendance_date, sub.code, sub.name subject_name,
        a.status, COALESCE(a.lecture_no,1) lecture_no, COALESCE(a.lecture_time,'') lecture_time
        FROM attendance a JOIN subjects sub ON sub.id=a.subject_id
        WHERE a.student_id=? AND sub.year=? AND sub.semester=?"""
    history_params = [student["id"], student["year"], student["semester"]]
    if selected_date:
        history_sql += " AND a.attendance_date=?"
        history_params.append(selected_date)
    if selected_subject:
        history_sql += " AND a.subject_id=?"
        history_params.append(selected_subject)
    history_sql += " ORDER BY a.attendance_date DESC, sub.code, a.lecture_no"
    history = db.execute(history_sql, tuple(history_params)).fetchall()

    summary = db.execute("""SELECT sub.code, sub.name subject_name, COALESCE(sub.session_type,'Lecture') session_type, COALESCE(sub.semester,1) semester,
        COUNT(a.id) total,
        COALESCE(SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END),0) present,
        COALESCE(SUM(CASE WHEN a.status='Absent' THEN 1 ELSE 0 END),0) absent,
        CASE WHEN COUNT(a.id)=0 THEN 0 ELSE ROUND(100.0*SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END)/COUNT(a.id),1) END pct
        FROM subjects sub LEFT JOIN attendance a ON a.subject_id=sub.id AND a.student_id=?
        WHERE sub.department_id=? AND sub.year=? AND sub.semester=? GROUP BY sub.id ORDER BY sub.code""", (student["id"], student["department_id"], student["year"], student["semester"])).fetchall()

    total = db.execute("SELECT COUNT(*) FROM attendance WHERE student_id=?", (student["id"],)).fetchone()[0]
    present = db.execute("SELECT COUNT(*) FROM attendance WHERE student_id=? AND status='Present'", (student["id"],)).fetchone()[0]
    overall = round((present / total * 100), 1) if total else 0
    threshold = attendance_threshold_value(db)
    db.close()
    return render_template("student_dashboard.html", student=student, today_rows=today_rows,
                           summary=summary, overall=overall, history=history, subjects=subjects,
                           selected_date=selected_date, selected_subject=selected_subject, today=today,
                           threshold=threshold)


@app.route("/student/qr-scanner")
@student_required
def student_qr_scanner():
    """Open the in-app camera QR scanner for the logged-in student."""
    return render_template("student_qr_scanner.html")


@app.route("/student/my-report")
@student_required
def student_my_report():
    """Show the logged-in student's own attendance report with optional session/subject filters."""
    db = get_db()
    student = current_student(db)
    if not student:
        db.close()
        session.clear()
        flash("Student account is not linked correctly. Contact the administrator.", "error")
        return redirect(url_for("login"))

    selected_session = (request.args.get("session") or "").strip()
    selected_semester = int(student["semester"] or 1)
    if selected_session not in ("", "Lecture", "Practical"):
        selected_session = ""
    try:
        selected_subject = int(request.args.get("subject_id", "0")) if request.args.get("subject_id") else 0
    except ValueError:
        selected_subject = 0

    subjects = db.execute("""SELECT id, code, name, COALESCE(session_type,'Lecture') session_type, COALESCE(semester,1) semester FROM subjects
        WHERE department_id=? AND year=? AND semester=? AND (?='' OR COALESCE(session_type,'Lecture')=?) ORDER BY session_type, code""",
        (student["department_id"], student["year"], selected_semester, selected_session, selected_session)).fetchall()
    allowed_subject_ids = {int(r["id"]) for r in subjects}
    if selected_subject and selected_subject not in allowed_subject_ids:
        selected_subject = 0

    summary_where = ["sub.department_id=?", "sub.year=?"]
    summary_params = [student["department_id"], student["year"]]
    if selected_semester:
        summary_where.append("sub.semester=?"); summary_params.append(selected_semester)
    if selected_session:
        summary_where.append("COALESCE(sub.session_type,'Lecture')=?"); summary_params.append(selected_session)
    attendance_join = "a.subject_id=sub.id AND a.student_id=?"
    join_params = [student["id"]]
    if selected_session:
        attendance_join += " AND COALESCE(a.session_type,'Lecture')=?"
        join_params.append(selected_session)
    if selected_subject:
        summary_where.append("sub.id=?")
        summary_params.append(selected_subject)

    summary = db.execute(f"""SELECT sub.code, sub.name subject_name,
        COUNT(a.id) total,
        COALESCE(SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END),0) present,
        COALESCE(SUM(CASE WHEN a.status='Absent' THEN 1 ELSE 0 END),0) absent,
        CASE WHEN COUNT(a.id)=0 THEN 0
             ELSE ROUND(100.0*SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END)/COUNT(a.id),1) END pct
        FROM subjects sub
        LEFT JOIN attendance a ON {attendance_join}
        WHERE {' AND '.join(summary_where)}
        GROUP BY sub.id ORDER BY sub.code""",
        tuple(join_params + summary_params)).fetchall()

    history_where = ["a.student_id=?"]
    history_params = [student["id"]]
    if selected_semester:
        history_where.append("sub.semester=?"); history_params.append(selected_semester)
    if selected_session:
        history_where.append("COALESCE(sub.session_type,'Lecture')=?"); history_params.append(selected_session)
    if selected_subject:
        history_where.append("a.subject_id=?")
        history_params.append(selected_subject)
    if selected_session:
        history_where.append("COALESCE(a.session_type,'Lecture')=?")
        history_params.append(selected_session)

    history = db.execute(f"""SELECT a.attendance_date, sub.code, sub.name subject_name,
        COALESCE(a.lecture_no,1) lecture_no, COALESCE(a.lecture_time,'') lecture_time,
        COALESCE(a.session_type,'Lecture') session_type, COALESCE(a.batch,'') batch, a.status
        FROM attendance a JOIN subjects sub ON sub.id=a.subject_id
        WHERE {' AND '.join(history_where)}
        ORDER BY a.attendance_date DESC, sub.code, a.lecture_no""",
        tuple(history_params)).fetchall()

    total = sum(int(r["total"] or 0) for r in summary)
    present = sum(int(r["present"] or 0) for r in summary)
    absent = sum(int(r["absent"] or 0) for r in summary)
    overall = round((present / total * 100), 1) if total else 0
    selected_subject_name = ""
    if selected_subject:
        for sub in subjects:
            if int(sub["id"]) == selected_subject:
                selected_subject_name = sub["name"]
                break

    db.close()
    return render_template("student_report.html", student=student, summary=summary, history=history,
                           subjects=subjects, selected_session=selected_session, selected_semester=selected_semester, semesters=SEMESTERS_BY_YEAR.get(student["year"], []),
                           selected_subject=selected_subject, selected_subject_name=selected_subject_name,
                           total=total, present=present, absent=absent, overall=overall)


@app.route("/student/my-report.pdf")
@student_required
def student_my_report_pdf():
    """Export the logged-in student's filtered attendance report as PDF."""
    db = get_db()
    student = current_student(db)
    if not student:
        db.close()
        return Response("Student record not found.", status=404)

    selected_session = (request.args.get("session") or "").strip()
    if selected_session not in ("", "Lecture", "Practical"):
        selected_session = ""
    selected_semester = int(student["semester"] or 1)
    try:
        selected_subject = int(request.args.get("subject_id", "0")) if request.args.get("subject_id") else 0
    except ValueError:
        selected_subject = 0

    subjects = db.execute("""SELECT id, code, name, COALESCE(session_type,'Lecture') session_type, COALESCE(semester,1) semester FROM subjects
        WHERE department_id=? AND year=? AND semester=? AND (?='' OR COALESCE(session_type,'Lecture')=?) ORDER BY session_type, code""",
        (student["department_id"], student["year"], selected_semester, selected_session, selected_session)).fetchall()
    allowed_subject_ids = {int(r["id"]) for r in subjects}
    if selected_subject not in allowed_subject_ids:
        selected_subject = 0

    summary_where = ["sub.department_id=?", "sub.year=?"]
    summary_params = [student["department_id"], student["year"]]
    if selected_semester:
        summary_where.append("sub.semester=?"); summary_params.append(selected_semester)
    if selected_session:
        summary_where.append("COALESCE(sub.session_type,'Lecture')=?"); summary_params.append(selected_session)
    attendance_join = "a.subject_id=sub.id AND a.student_id=?"
    join_params = [student["id"]]
    if selected_subject:
        summary_where.append("sub.id=?"); summary_params.append(selected_subject)
    if selected_session:
        attendance_join += " AND COALESCE(a.session_type,'Lecture')=?"
        join_params.append(selected_session)
    summary = db.execute(f"""SELECT sub.code, sub.name subject_name,
        COUNT(a.id) total,
        COALESCE(SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END),0) present,
        COALESCE(SUM(CASE WHEN a.status='Absent' THEN 1 ELSE 0 END),0) absent,
        CASE WHEN COUNT(a.id)=0 THEN 0
             ELSE ROUND(100.0*SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END)/COUNT(a.id),1) END pct
        FROM subjects sub
        LEFT JOIN attendance a ON {attendance_join}
        WHERE {' AND '.join(summary_where)}
        GROUP BY sub.id ORDER BY sub.code""", tuple(join_params + summary_params)).fetchall()

    history_where = ["a.student_id=?"]
    history_params = [student["id"]]
    if selected_semester:
        history_where.append("sub.semester=?"); history_params.append(selected_semester)
    if selected_session:
        history_where.append("COALESCE(sub.session_type,'Lecture')=?"); history_params.append(selected_session)
    if selected_subject:
        history_where.append("a.subject_id=?"); history_params.append(selected_subject)
    if selected_session:
        history_where.append("COALESCE(a.session_type,'Lecture')=?"); history_params.append(selected_session)
    history = db.execute(f"""SELECT a.attendance_date, sub.code, sub.name subject_name,
        COALESCE(a.lecture_no,1) lecture_no, COALESCE(a.lecture_time,'') lecture_time,
        COALESCE(a.session_type,'Lecture') session_type, COALESCE(a.batch,'') batch, a.status
        FROM attendance a JOIN subjects sub ON sub.id=a.subject_id
        WHERE {' AND '.join(history_where)}
        ORDER BY a.attendance_date DESC, sub.code, a.lecture_no""", tuple(history_params)).fetchall()

    total = sum(int(r["total"] or 0) for r in summary)
    present = sum(int(r["present"] or 0) for r in summary)
    absent = sum(int(r["absent"] or 0) for r in summary)
    overall = round((present / total * 100), 1) if total else 0
    selected_subject_name = ""
    selected_subject_code = ""
    if selected_subject:
        for sub in subjects:
            if int(sub["id"]) == selected_subject:
                selected_subject_name = sub["name"]
                selected_subject_code = sub["code"] or ""
                break
    threshold = attendance_threshold_value(db)
    db.close()

    by_key = {}
    for r in history:
        by_key.setdefault(r["code"], []).append(r)
    report_items = []
    for r in summary:
        if not int(r["total"] or 0) and not selected_subject_name:
            continue
        label = r["subject_name"]
        report_items.append({"semester": selected_semester or student["semester"] or "", "code": r["code"], "name": label,
                             "total": int(r["total"] or 0), "present": int(r["present"] or 0), "absent": int(r["absent"] or 0),
                             "pct": float(r["pct"] or 0), "history": by_key.get(r["code"], [])})
    if not report_items:
        return Response("No attendance records found for the selected filters.", status=404, mimetype="text/plain")
    pdf_bytes = build_statement_pdf(dict(student), report_items, college=COLLEGE_NAME,
                                    logo_path=BASE_DIR / "static" / LOGO_FILE, threshold=threshold)
    safe_prn = re.sub(r"[^A-Za-z0-9_-]+", "_", str(student["prn"] or student["id"])).strip("_") or "student"
    suffix = ""
    if selected_subject_name:
        suffix += "_" + re.sub(r"[^A-Za-z0-9_-]+", "_", selected_subject_name).strip("_")
    if selected_session:
        suffix += "_" + selected_session.lower()
    return Response(pdf_bytes, mimetype="application/pdf",
                    headers={"Content-Disposition": f"attachment; filename=student_attendance_report_{safe_prn}{suffix}.pdf"})


@app.route("/student/calendar")
@student_required
def student_calendar():
    db = get_db(); student = current_student(db)
    month = request.args.get("month", today_ist().strftime("%Y-%m"))
    if not re.match(r"^\d{4}-\d{2}$", month):
        month = today_ist().strftime("%Y-%m")
    start = f"{month}-01"
    y, m = map(int, month.split("-"))
    end = f"{y+1:04d}-01-01" if m == 12 else f"{y:04d}-{m+1:02d}-01"

    rows = db.execute("""SELECT a.attendance_date,
        SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END) present,
        SUM(CASE WHEN a.status='Absent' THEN 1 ELSE 0 END) absent,
        COUNT(a.id) total
        FROM attendance a
        WHERE a.student_id=? AND a.attendance_date>=? AND a.attendance_date<?
        GROUP BY a.attendance_date ORDER BY a.attendance_date""",
        (student["id"], start, end)).fetchall()

    detail_rows = db.execute("""SELECT a.attendance_date, a.status, a.lecture_no,
        a.lecture_time, a.start_time, a.end_time,
        COALESCE(sub.code,'') subject_code, COALESCE(sub.name,'') subject_name
        FROM attendance a
        LEFT JOIN subjects sub ON sub.id=a.subject_id
        WHERE a.student_id=? AND a.attendance_date>=? AND a.attendance_date<?
        ORDER BY a.attendance_date, a.start_time, a.lecture_no, a.id""",
        (student["id"], start, end)).fetchall()
    db.close()

    details_by_date = {}
    for row in detail_rows:
        key = str(row["attendance_date"])
        details_by_date.setdefault(key, []).append(row)

    # Build the exact number of calendar cells for the selected month.
    first_weekday, total_days = monthrange(y, m)  # Monday=0
    calendar_days = []
    for _ in range(first_weekday):
        calendar_days.append(None)
    for day_number in range(1, total_days + 1):
        key = f"{month}-{day_number:02d}"
        calendar_days.append({"number": day_number, "date": key,
                              "details": details_by_date.get(key, [])})
    while len(calendar_days) % 7:
        calendar_days.append(None)

    return render_template("student_calendar.html", student=student, month=month,
                           days=rows, details_by_date=details_by_date,
                           calendar_days=calendar_days)


@app.route("/students")
@login_required
@admin_required
def students():
    q = request.args.get("q", "").strip()
    selected_year = request.args.get("year", "").strip()
    selected_semester = request.args.get("semester", type=int) or 0
    db = get_db()

    # Four year-wise student cards. Counts come directly from the students table.
    year_counts = {}
    for year in YEAR_OPTIONS:
        year_counts[year] = db.execute(
            "SELECT COUNT(*) FROM students WHERE year=?", (year,)
        ).fetchone()[0]

    conditions = []
    params = []
    if selected_year in YEAR_OPTIONS:
        conditions.append("s.year=?")
        params.append(selected_year)
    if selected_semester in range(1, 9):
        conditions.append("s.semester=?")
        params.append(selected_semester)
    if q:
        conditions.append("(s.prn LIKE ? OR s.name LIKE ? OR s.course LIKE ?)")
        params.extend([f"%{q}%", f"%{q}%", f"%{q}%"])

    where = (" WHERE " + " AND ".join(conditions)) if conditions else ""
    rows = db.execute(f"""SELECT s.*, u.username student_username
        FROM students s LEFT JOIN users u ON u.student_id=s.id AND u.role='student'
        {where}
        ORDER BY s.prn, s.name""", tuple(params)).fetchall()

    # Arrange students by complete PRN in ascending numerical order.
    # Numeric PRNs are compared as numbers, so ...3009, ...3010, ...3011
    # comes before ...3020, ...3024, etc. Non-numeric PRNs are kept after
    # numeric PRNs and sorted alphabetically.
    def _prn_sort_key(row):
        value = str(row["prn"] or "").strip()
        if value.isdigit():
            return (0, int(value), "")
        return (1, 0, value.lower())

    rows = sorted(rows, key=_prn_sort_key)
    db.close()
    return render_template("students.html", students=rows, q=q, selected_year=selected_year, selected_semester=selected_semester,
                           year_options=YEAR_OPTIONS, semesters_by_year=SEMESTERS_BY_YEAR, year_counts=year_counts)


@app.route("/students/add", methods=["GET", "POST"])
@login_required
@admin_required
def add_student():
    db = get_db()
    departments = db.execute("SELECT * FROM departments WHERE active=1 ORDER BY name").fetchall()
    if request.method == "POST":
        data = [request.form.get(k, "").strip() for k in ("prn", "name", "course", "year")]
        semester = request.form.get("semester", type=int)
        batch = request.form.get("batch", "").strip()
        dept_id = departments[0]["id"] if departments else None
        if not data[0] or not data[1] or data[3] not in YEAR_OPTIONS or semester not in SEMESTERS_BY_YEAR.get(data[3], []) or not dept_id:
            db.close(); flash("PRN, name and year are required.", "error")
            return render_template("student_form.html", student=None, title="Add Student", year_options=YEAR_OPTIONS, semesters_by_year=SEMESTERS_BY_YEAR, batch_options=["Batch A","Batch B","Batch C","Batch D"])
        try:
            db.execute("INSERT INTO students(roll_no,prn,name,course,year,semester,division,department_id,batch) VALUES(?,?,?,?,?,?,?,?,?)", (data[0], data[0], data[1], data[2] or "B.Pharm", data[3], semester, "A", dept_id, batch))
            db.commit(); flash("Student added successfully.", "success")
        except Exception as exc:
            if not isinstance(exc, sqlite3.IntegrityError) and not (POSTGRES_AVAILABLE and isinstance(exc, psycopg2.IntegrityError)): raise
            flash("PRN already exists.", "error")
        db.close(); return redirect(url_for("students"))
    db.close(); return render_template("student_form.html", student=None, title="Add Student", year_options=YEAR_OPTIONS, semesters_by_year=SEMESTERS_BY_YEAR, batch_options=["Batch A","Batch B","Batch C","Batch D"])


@app.route("/students/edit/<int:student_id>", methods=["GET", "POST"])
@login_required
@admin_required
def edit_student(student_id):
    db = get_db(); student = db.execute("SELECT * FROM students WHERE id=?", (student_id,)).fetchone()
    if not student:
        db.close(); flash("Student not found.", "error"); return redirect(url_for("students"))
    if request.method == "POST":
        data = [request.form.get(k, "").strip() for k in ("prn", "name", "course", "year")]
        semester = request.form.get("semester", type=int)
        batch = request.form.get("batch", "").strip()
        try:
            dept_id = student["department_id"]
            if not dept_id:
                dept_row = db.execute("SELECT id FROM departments ORDER BY id LIMIT 1").fetchone()
                dept_id = dept_row["id"] if dept_row else None
            if not data[0] or not data[1] or data[3] not in YEAR_OPTIONS or semester not in SEMESTERS_BY_YEAR.get(data[3], []):
                raise ValueError("PRN, name, valid year and semester are required.")
            db.execute("UPDATE students SET roll_no=?,prn=?,name=?,course=?,year=?,semester=?,department_id=?,batch=? WHERE id=?", (data[0], data[0], data[1], data[2] or "B.Pharm", data[3], semester, dept_id, batch, student_id))
            db.execute("UPDATE users SET full_name=? WHERE student_id=? AND role='student'", (data[1], student_id))
            db.commit(); flash("Student updated successfully.", "success"); db.close(); return redirect(url_for("students"))
        except ValueError as exc:
            db.close(); flash(str(exc), "error")
        except Exception as exc:
            if not isinstance(exc, sqlite3.IntegrityError) and not (POSTGRES_AVAILABLE and isinstance(exc, psycopg2.IntegrityError)):
                raise
            flash("PRN already exists.", "error")
    departments=db.execute("SELECT * FROM departments WHERE active=1 ORDER BY name").fetchall(); db.close(); return render_template("student_form.html", student=student, title="Edit Student", year_options=YEAR_OPTIONS, semesters_by_year=SEMESTERS_BY_YEAR, batch_options=["Batch A","Batch B","Batch C","Batch D"])


def archive_student(db, student_id, deleted_by):
    """Save a full copy of the student (account, attendance, leaves) so an admin can restore it later."""
    student = db.execute("SELECT * FROM students WHERE id=?", (student_id,)).fetchone()
    if not student:
        return False
    def rows(table):
        return [dict(r) for r in db.execute(f"SELECT * FROM {table} WHERE student_id=?", (student_id,)).fetchall()]
    snapshot = {"student": dict(student), "users": rows("users"), "attendance": rows("attendance"),
                "leave_applications": rows("leave_applications"),
                "student_academic_history": rows("student_academic_history")}
    db.execute("""INSERT INTO deleted_students(student_id,name,prn,roll_no,year,deleted_at,deleted_by,snapshot)
                  VALUES(?,?,?,?,?,?,?,?)""",
               (student_id, student["name"], student["prn"] or "", student["roll_no"] or "", student["year"] or "",
                datetime.utcnow().isoformat(), deleted_by, json.dumps(snapshot, default=str)))
    return True


@app.route("/students/delete/<int:student_id>", methods=["POST"])
@login_required
@admin_required
def delete_student(student_id):
    db = get_db()
    try:
        archived = archive_student(db, student_id, session.get("user_id"))
        db.execute("DELETE FROM students WHERE id=?", (student_id,))
        db.commit()
        if archived:
            log_activity("Student deleted", f"Student id {student_id} moved to Deleted Students")
            flash("Student deleted. You can restore this student from Admin > Deleted Students.", "success")
        else:
            flash("Student not found.", "error")
    except Exception:
        db.rollback()
        app.logger.exception("Student delete failed")
        flash("Student could not be deleted.", "error")
    finally:
        db.close()
    return redirect(url_for("students"))


_SAFE_COLUMN = re.compile(r"^[a-z_][a-z0-9_]*$")


def _insert_snapshot_row(db, table, row):
    cols = [c for c in row.keys() if _SAFE_COLUMN.match(c)]
    marks = ",".join("?" for _ in cols)
    db.execute(f"INSERT INTO {table}({','.join(cols)}) VALUES({marks})", [row[c] for c in cols])


def restore_deleted_student(db, archive_row):
    """Put a deleted student back exactly as it was. Raises ValueError with a readable message on conflict."""
    snap = json.loads(archive_row["snapshot"])
    st = snap["student"]
    if db.execute("SELECT 1 FROM students WHERE id=? OR roll_no=? OR (prn<>'' AND prn=?)",
                  (st["id"], st["roll_no"], st.get("prn") or "")).fetchone():
        raise ValueError("A student with the same ID, roll number or PRN already exists, so this one cannot be restored.")
    for u in snap.get("users", []):
        if db.execute("SELECT 1 FROM users WHERE id=? OR username=?", (u["id"], u["username"])).fetchone():
            raise ValueError("The login username of this student is already in use.")
    _insert_snapshot_row(db, "students", st)
    for u in snap.get("users", []):
        _insert_snapshot_row(db, "users", u)
    subject_ids = {int(r["id"]) for r in db.execute("SELECT id FROM subjects").fetchall()}
    user_ids = {int(r["id"]) for r in db.execute("SELECT id FROM users").fetchall()}
    skipped = 0
    for a in snap.get("attendance", []):
        if a.get("subject_id") is None or int(a["subject_id"]) not in subject_ids:
            skipped += 1
            continue
        for key in ("marked_by", "scheduled_teacher_id", "conducted_by_teacher_id"):
            if a.get(key) is not None and int(a[key]) not in user_ids:
                a[key] = None
        _insert_snapshot_row(db, "attendance", a)
    for l in snap.get("leave_applications", []):
        if l.get("reviewed_by") is not None and int(l["reviewed_by"]) not in user_ids:
            l["reviewed_by"] = None
        _insert_snapshot_row(db, "leave_applications", l)
    for h in snap.get("student_academic_history", []):
        if h.get("promoted_by") is not None and int(h["promoted_by"]) not in user_ids:
            h["promoted_by"] = None
        _insert_snapshot_row(db, "student_academic_history", h)
    return skipped


@app.route("/admin/deleted-students")
@login_required
@admin_required
def deleted_students():
    db = get_db()
    rows = db.execute("SELECT id,student_id,name,prn,roll_no,year,deleted_at FROM deleted_students ORDER BY id DESC").fetchall()
    db.close()
    return render_template("deleted_students.html", rows=rows)


@app.route("/admin/deleted-students/<int:archive_id>/restore", methods=["POST"])
@login_required
@admin_required
def restore_student(archive_id):
    db = get_db()
    try:
        row = db.execute("SELECT * FROM deleted_students WHERE id=?", (archive_id,)).fetchone()
        if not row:
            flash("Deleted student record not found.", "error")
        else:
            skipped = restore_deleted_student(db, row)
            db.execute("DELETE FROM deleted_students WHERE id=?", (archive_id,))
            db.commit()
            log_activity("Student restored", f"{row['name']} ({row['roll_no']}) restored")
            extra = f" {skipped} attendance record(s) were skipped because their subject no longer exists." if skipped else ""
            flash(f"{row['name']} was restored with the account and attendance.{extra}", "success")
    except ValueError as exc:
        db.rollback()
        flash(str(exc), "error")
    except Exception:
        db.rollback()
        app.logger.exception("Student restore failed")
        flash("Student could not be restored.", "error")
    finally:
        db.close()
    return redirect(url_for("deleted_students"))


@app.route("/admin/deleted-students/<int:archive_id>/purge", methods=["POST"])
@login_required
@admin_required
def purge_deleted_student(archive_id):
    db = get_db()
    db.execute("DELETE FROM deleted_students WHERE id=?", (archive_id,))
    db.commit(); db.close()
    flash("Deleted student record removed permanently.", "success")
    return redirect(url_for("deleted_students"))


@app.route("/students/<int:student_id>/toggle-block", methods=["POST"])
@login_required
@admin_required
def toggle_block_student(student_id):
    db = get_db()
    student = db.execute("SELECT id, name, blocked FROM students WHERE id=?", (student_id,)).fetchone()
    if not student:
        db.close()
        flash("Student not found.", "error")
        return redirect(url_for("students"))
    new_status = 0 if int(student["blocked"] or 0) else 1
    db.execute("UPDATE students SET blocked=? WHERE id=?", (new_status, student_id))
    db.commit()
    db.close()
    if new_status:
        flash(f"{student['name']} has been blocked. Student login and registration are disabled.", "success")
    else:
        flash(f"{student['name']} has been unblocked. Student can log in and register again.", "success")
    return redirect(url_for("students"))


@app.route("/teachers")
@login_required
@admin_required
def teachers():
    db = get_db()
    rows = db.execute("SELECT * FROM users WHERE role='teacher' ORDER BY approved ASC, full_name, username").fetchall()
    subjects = db.execute("SELECT * FROM subjects ORDER BY year, code").fetchall()
    assignments = db.execute("SELECT teacher_id, subject_id FROM subject_teachers").fetchall()
    assigned = {(r["teacher_id"], r["subject_id"]) for r in assignments}
    db.close()
    return render_template("teachers.html", teachers=rows, subjects=subjects, assigned=assigned, years=YEAR_OPTIONS)


@app.route("/teachers/<int:teacher_id>/approve", methods=["POST"])
@login_required
@admin_required
def approve_teacher(teacher_id):
    db = get_db(); db.execute("UPDATE users SET approved=1 WHERE id=? AND role='teacher'", (teacher_id,)); db.commit(); db.close()
    flash("Teacher account approved.", "success"); return redirect(url_for("teachers"))


@app.route("/teachers/<int:teacher_id>/toggle", methods=["POST"])
@login_required
@admin_required
def toggle_teacher(teacher_id):
    db = get_db(); db.execute("UPDATE users SET approved=CASE approved WHEN 1 THEN 0 ELSE 1 END WHERE id=? AND role='teacher'", (teacher_id,)); db.commit(); db.close()
    flash("Teacher account status updated.", "success"); return redirect(url_for("teachers"))


@app.route("/teachers/<int:teacher_id>/delete", methods=["POST"])
@login_required
@admin_required
def delete_teacher(teacher_id):
    db = get_db()
    teacher = db.execute("SELECT full_name,username FROM users WHERE id=? AND role='teacher'", (teacher_id,)).fetchone()
    if not teacher:
        db.close(); flash("Teacher not found.", "error"); return redirect(url_for("teachers"))
    db.execute("DELETE FROM users WHERE id=? AND role='teacher'", (teacher_id,))
    db.commit(); db.close()
    flash(f"Teacher {teacher['full_name'] or teacher['username']} removed successfully.", "success")
    return redirect(url_for("teachers"))


@app.route("/teachers/<int:teacher_id>/assign", methods=["POST"])
@login_required
@admin_required
def assign_subjects(teacher_id):
    db = get_db()
    db.execute("DELETE FROM subject_teachers WHERE teacher_id=?", (teacher_id,))
    subject_ids = request.form.getlist("subject_ids")
    valid = db.execute("SELECT id FROM subjects").fetchall()
    valid_ids = {str(r["id"]) for r in valid}
    for sid in subject_ids:
        if sid in valid_ids:
            db.execute("INSERT OR IGNORE INTO subject_teachers(teacher_id,subject_id) VALUES(?,?)", (teacher_id, int(sid)))
    db.commit(); db.close()
    flash("Teacher subject assignments saved year-wise.", "success"); return redirect(url_for("teachers"))


@app.route("/subjects")
@login_required
@admin_required
def subjects():
    db = get_db()
    rows = db.execute("SELECT s.*, d.name department_name FROM subjects s LEFT JOIN departments d ON d.id=s.department_id ORDER BY s.year, s.semester, s.session_type, s.code").fetchall()
    db.close()
    grouped = {y: [] for y in YEAR_OPTIONS}
    for row in rows:
        grouped.setdefault(row["year"] or "1st Year", []).append(row)
    return render_template("subjects.html", subjects=rows, grouped=grouped, years=YEAR_OPTIONS, semesters_by_year=SEMESTERS_BY_YEAR)


@app.route("/subjects/add", methods=["GET", "POST"])
@login_required
@admin_required
def add_subject():
    db = get_db()
    departments = db.execute("SELECT * FROM departments WHERE active=1 ORDER BY name").fetchall()
    if request.method == "POST":
        code = request.form.get("code", "").strip().upper(); name = request.form.get("name", "").strip()
        year = request.form.get("year", "").strip()
        semester = request.form.get("semester", type=int)
        session_type = request.form.get("session_type", "Lecture").strip()
        # Department is intentionally hidden from the UI. Keep the existing
        # internal department relationship by assigning the default department.
        default_dept_row = db.execute("SELECT id FROM departments ORDER BY id LIMIT 1").fetchone()
        dept_id = default_dept_row["id"] if default_dept_row else None
        if year not in YEAR_OPTIONS: year = ""
        if semester not in SEMESTERS_BY_YEAR.get(year, []) or session_type not in ("Lecture", "Practical"): semester = None
        if not code or not name or not year or not semester or not dept_id:
            db.close(); flash("Subject code, name and year are required.", "error")
            return render_template("subject_form.html", subject=None, title="Add Subject", years=YEAR_OPTIONS, semesters_by_year=SEMESTERS_BY_YEAR, selected_year=request.args.get("year", "1st Year"), selected_semester=request.args.get("semester", type=int) or 1, selected_session_type=request.args.get("session_type", "Lecture"))
        try:
            db.execute("INSERT INTO subjects(code,name,year,semester,session_type,department_id) VALUES(?,?,?,?,?,?)", (code, name, year, semester, session_type, dept_id)); db.commit(); flash("Subject added.", "success")
        except Exception as exc:
            if not isinstance(exc, sqlite3.IntegrityError) and not (POSTGRES_AVAILABLE and isinstance(exc, psycopg2.IntegrityError)): raise
            flash("Subject code already exists.", "error")
        db.close(); return redirect(url_for("subjects"))
    db.close(); return render_template("subject_form.html", subject=None, title="Add Subject", years=YEAR_OPTIONS, semesters_by_year=SEMESTERS_BY_YEAR, selected_year=request.args.get("year", "1st Year"), selected_semester=request.args.get("semester", type=int) or 1, selected_session_type=request.args.get("session_type", "Lecture"))


@app.route("/subjects/edit/<int:subject_id>", methods=["GET", "POST"])
@login_required
@admin_required
def edit_subject(subject_id):
    db = get_db(); subject = db.execute("SELECT * FROM subjects WHERE id=?", (subject_id,)).fetchone()
    if not subject:
        db.close(); flash("Subject not found.", "error"); return redirect(url_for("subjects"))
    if request.method == "POST":
        code = request.form.get("code", "").strip().upper()
        name = request.form.get("name", "").strip()
        year = request.form.get("year", "").strip()
        semester = request.form.get("semester", type=int)
        session_type = request.form.get("session_type", "Lecture").strip()
        # Preserve the existing internal department relationship; users no longer select it.
        dept_id = subject["department_id"]
        if not code or not name or year not in YEAR_OPTIONS or semester not in SEMESTERS_BY_YEAR.get(year, []) or session_type not in ("Lecture", "Practical"):
            flash("Subject code, name, year, semester and session are required.", "error")
        else:
            try:
                db.execute("UPDATE subjects SET code=?,name=?,year=?,semester=?,session_type=?,department_id=? WHERE id=?", (code, name, year, semester, session_type, dept_id, subject_id)); db.commit(); flash("Subject updated.", "success"); db.close(); return redirect(url_for("subjects"))
            except Exception as exc:
                if not isinstance(exc, sqlite3.IntegrityError) and not (POSTGRES_AVAILABLE and isinstance(exc, psycopg2.IntegrityError)): raise
                flash("Subject code already exists.", "error")
    departments=db.execute("SELECT * FROM departments WHERE active=1 ORDER BY name").fetchall(); db.close(); return render_template("subject_form.html", subject=subject, title="Edit Subject", years=YEAR_OPTIONS, semesters_by_year=SEMESTERS_BY_YEAR, selected_year=subject["year"], selected_semester=subject["semester"], selected_session_type=subject["session_type"])


@app.route("/subjects/delete/<int:subject_id>", methods=["POST"])
@login_required
@admin_required
def delete_subject(subject_id):
    db = get_db(); db.execute("DELETE FROM subjects WHERE id=?", (subject_id,)); db.commit(); db.close()
    flash("Subject deleted.", "success"); return redirect(url_for("subjects"))


def _prn_sort_key(row):
    value = str(row["prn"] or "").strip()
    if value.isdigit():
        return (0, int(value), "")
    return (1, 0, value.lower())


@app.route("/attendance", methods=["GET", "POST"])
@login_required
@staff_required
def take_attendance():
    db=get_db(); subjects=teacher_or_admin_subjects(db); years=YEAR_OPTIONS[:]
    selected_year=request.args.get('year') or request.form.get('year') or '1st Year'
    selected_semester=request.args.get('semester', type=int) or request.form.get('semester', type=int)
    if selected_semester not in SEMESTERS_BY_YEAR.get(selected_year, []): selected_semester=SEMESTERS_BY_YEAR.get(selected_year, [1])[0]
    session_type=(request.args.get('session_type') or request.form.get('session_type') or 'Lecture').strip()
    if session_type not in ('Lecture','Practical'): session_type='Lecture'
    selected_batch=(request.args.get('batch') or request.form.get('batch') or '').strip()
    selected_subject=int(request.args.get('subject_id') or request.form.get('subject_id') or 0)
    selected_date=request.args.get('attendance_date') or request.form.get('attendance_date') or today_ist().isoformat()
    try:
        lecture_no=int(request.args.get('lecture_no') or request.form.get('lecture_no') or 1)
    except ValueError:
        lecture_no=1
    if lecture_no not in range(1,6): lecture_no=1
    # Backward compatibility with older links/forms that used lecture_count.
    if not request.args.get('lecture_no') and not request.form.get('lecture_no'):
        try:
            old_count=int(request.args.get('lecture_count') or request.form.get('lecture_count') or 1)
            if old_count in range(1,6): lecture_no=old_count
        except ValueError: pass
    if session_type=='Practical' and selected_batch not in ('Batch A','Batch B','Batch C','Batch D'): selected_batch=''
    subjects=[s for s in subjects if s['year']==selected_year and int(s['semester'] or SEMESTERS_BY_YEAR.get(selected_year,[1])[0])==selected_semester and (s['session_type'] or 'Lecture')==session_type]
    if selected_subject and not any(s['id']==selected_subject for s in subjects): selected_subject=0
    if selected_subject and not can_use_subject(db,selected_subject): selected_subject=0
    if session_type=='Practical':
        students=db.execute('SELECT * FROM students WHERE year=? AND batch=? ORDER BY prn,name',(selected_year,selected_batch)).fetchall() if selected_subject and selected_batch else []
    else:
        students=db.execute('SELECT * FROM students WHERE year=? ORDER BY prn,name',(selected_year,)).fetchall() if selected_subject else []
    students=sorted(students,key=_prn_sort_key)
    day_name=date.fromisoformat(selected_date).strftime('%A') if selected_date else today_ist().strftime('%A')
    candidates=timetable_candidates(db,selected_year,session_type,selected_batch,selected_subject,day_name) if selected_subject else []
    teachers=all_teacher_rows(db)
    lecture_time=''
    existing={}
    if request.method=='POST' and selected_subject and students:
        tid=request.form.get('timetable_id',type=int)
        tr=db.execute('SELECT * FROM timetable WHERE id=?',(tid,)).fetchone() if tid else None
        start_time=request.form.get('start_time','').strip(); end_time=request.form.get('end_time','').strip()
        teacher_status=request.form.get('teacher_status','Scheduled teacher present').strip()
        conducted=request.form.get('conducted_by',type=int)
        scheduled=tr['teacher_id'] if tr else None
        if tr and (tr['year']!=selected_year or tr['subject_id']!=selected_subject or tr['day_of_week']!=day_name): scheduled=None
        if teacher_status not in ('Scheduled teacher present','Scheduled teacher absent - substitute teacher','Another teacher conducted the session'): teacher_status='Scheduled teacher present'
        if teacher_status=='Scheduled teacher present': conducted=scheduled or session['user_id']
        elif not conducted: conducted=session['user_id']
        if not start_time or not end_time or start_time>=end_time:
            flash(f'Please select valid start and end time for Lecture {lecture_no}.','error')
        else:
            # Detect status changes before writing so an edit cannot bypass the audit trail.
            submitted = []
            changed_statuses = []
            for st in students:
                att_status = request.form.get(f'status_{st["id"]}')
                if att_status not in ('Present', 'Absent'):
                    continue
                prior = db.execute(
                    """SELECT id,status FROM attendance
                       WHERE student_id=? AND subject_id=? AND attendance_date=? AND lecture_no=?""",
                    (st['id'], selected_subject, selected_date, lecture_no)
                ).fetchone()
                submitted.append((st, att_status, prior))
                if prior and prior['status'] != att_status:
                    changed_statuses.append((st, att_status, prior))
            correction_reason = request.form.get('correction_reason', '').strip()
            if changed_statuses and not correction_reason:
                flash('Attendance already exists for some students. Enter a correction reason before changing their status; no changes were saved.', 'error')
            else:
                for st, att_status, prior in submitted:
                    db.execute("""INSERT INTO attendance(student_id,subject_id,attendance_date,lecture_no,lecture_time,start_time,end_time,session_type,batch,status,marked_by,academic_year,scheduled_teacher_id,conducted_by_teacher_id,teacher_status)
                        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                        ON CONFLICT(student_id,subject_id,attendance_date,lecture_no) DO UPDATE SET lecture_time=excluded.lecture_time,start_time=excluded.start_time,end_time=excluded.end_time,session_type=excluded.session_type,batch=excluded.batch,status=excluded.status,marked_by=excluded.marked_by,academic_year=excluded.academic_year,scheduled_teacher_id=excluded.scheduled_teacher_id,conducted_by_teacher_id=excluded.conducted_by_teacher_id,teacher_status=excluded.teacher_status""",
                        (st['id'], selected_subject, selected_date, lecture_no, f'{start_time}-{end_time}', start_time, end_time, session_type, selected_batch if session_type=='Practical' else '', att_status, session['user_id'], current_academic_year_value(db), scheduled, conducted, teacher_status))
                    if prior and prior['status'] != att_status:
                        db.execute("""INSERT INTO attendance_corrections
                            (attendance_id,student_id,subject_id,attendance_date,lecture_no,old_status,new_status,reason,corrected_by)
                            VALUES(?,?,?,?,?,?,?,?,?)""",
                            (prior['id'], st['id'], selected_subject, selected_date, lecture_no,
                             prior['status'], att_status, correction_reason, session['user_id']))
                db.commit()
                if changed_statuses:
                    flash(f'{session_type} attendance updated; {len(changed_statuses)} correction(s) were recorded in the audit trail.', 'success')
                else:
                    flash(f'{session_type} attendance for Lecture {lecture_no} saved with teacher status.', 'success')
    if selected_subject:
        rows=db.execute("SELECT student_id,lecture_no,lecture_time,status FROM attendance WHERE subject_id=? AND attendance_date=? AND lecture_no=? AND COALESCE(session_type,'Lecture')=? AND (COALESCE(batch,'')=? OR ?='Lecture')",(selected_subject,selected_date,lecture_no,session_type,selected_batch,session_type)).fetchall()
        for r in rows: existing[r['student_id']]=r['status']; lecture_time=r['lecture_time'] or lecture_time
    batch_counts={b:db.execute('SELECT COUNT(*) FROM students WHERE year=? AND batch=?',(selected_year,b)).fetchone()[0] for b in ('Batch A','Batch B','Batch C','Batch D')}
    db.close()
    return render_template('attendance.html',subjects=subjects,students=students,years=years,selected_year=selected_year,selected_semester=selected_semester,semesters_by_year=SEMESTERS_BY_YEAR,selected_subject=selected_subject,selected_date=selected_date,existing=existing,lecture_no=lecture_no,lecture_time=lecture_time,session_type=session_type,selected_batch=selected_batch,batch_options=['Batch A','Batch B','Batch C','Batch D'],batch_counts=batch_counts,candidates=candidates,teachers=teachers,day_name=day_name)


@app.route("/attendance/history/add")
@login_required
@admin_required
def add_attendance_history():
    """Compatibility endpoint for older attendance-history templates.

    Adding attendance is handled by the normal attendance-taking screen; this
    route exists so older templates using url_for('add_attendance_history')
    do not raise a Flask BuildError.
    """
    return redirect(url_for("take_attendance"))


@app.route("/attendance/history")
@app.route("/attendance-history")
@login_required
@staff_required
def attendance_history():
    db = get_db()
    where = ["1=1"]; params = []
    if session.get("role") == "teacher":
        where.append("a.subject_id IN (SELECT subject_id FROM subject_teachers WHERE teacher_id=?)")
        params.append(session["user_id"])
    year = request.args.get("year", "").strip()
    session_type = request.args.get("session_type", "").strip()
    semester = request.args.get("semester", type=int) or 0
    batch = request.args.get("batch", "").strip()
    subject_id = request.args.get("subject_id", type=int)
    from_date = request.args.get("from_date", "").strip(); to_date = request.args.get("to_date", "").strip()
    if year:
        where.append("s.year=?"); params.append(year)
    if semester:
        where.append("sub.semester=?"); params.append(semester)
    if session_type in ("Lecture","Practical"):
        where.append("COALESCE(a.session_type,'Lecture')=?"); params.append(session_type)
    if batch and session_type == "Practical":
        where.append("COALESCE(a.batch,'')=?"); params.append(batch)
    if subject_id:
        where.append("a.subject_id=?"); params.append(subject_id)
    if from_date:
        where.append("a.attendance_date>=?"); params.append(from_date)
    if to_date:
        where.append("a.attendance_date<=?"); params.append(to_date)
    condition=" AND ".join(where)
    rows=db.execute(f"""SELECT a.attendance_date,a.lecture_no,a.lecture_time,COALESCE(a.session_type,'Lecture') session_type,COALESCE(a.batch,'') batch,sub.id subject_id,sub.code,sub.name subject_name,s.year,
        COUNT(a.id) total, SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END) present,
        SUM(CASE WHEN a.status='Absent' THEN 1 ELSE 0 END) absent
        FROM attendance a JOIN subjects sub ON sub.id=a.subject_id JOIN students s ON s.id=a.student_id
        WHERE {condition}
        GROUP BY a.attendance_date,a.lecture_no,a.lecture_time,COALESCE(a.session_type,'Lecture'),COALESCE(a.batch,''),sub.id,sub.code,sub.name,s.year
        ORDER BY a.attendance_date DESC,a.lecture_no DESC,sub.code,s.year""", params).fetchall()
    subjects=teacher_or_admin_subjects(db); years=YEAR_OPTIONS
    subjects=[x for x in subjects if (not year or x["year"]==year) and (not semester or int(x["semester"] or 1)==semester) and (not session_type or (x["session_type"] or "Lecture")==session_type)]
    db.close()
    return render_template("attendance_history.html", rows=rows, subjects=subjects, years=years, year=year, semester=semester, semesters_by_year=SEMESTERS_BY_YEAR, subject_id=subject_id, from_date=from_date, to_date=to_date, session_type=session_type, batch=batch, batch_options=["Batch A","Batch B","Batch C","Batch D"])


@app.route("/attendance/history/delete", methods=["POST"])
@app.route("/attendance-history/delete", methods=["POST"])
@login_required
@admin_required
def delete_attendance_session():
    """Delete one complete attendance session safely on SQLite or PostgreSQL."""
    attendance_date=(request.form.get("attendance_date") or "").strip()
    session_type=(request.form.get("session_type") or "Lecture").strip()
    batch=(request.form.get("batch") or "").strip()
    lecture_time=(request.form.get("lecture_time") or "").strip()
    try:
        subject_id=int(request.form.get("subject_id") or 0)
        lecture_no=int(request.form.get("lecture_no") or 0)
    except (TypeError, ValueError):
        subject_id=0; lecture_no=0

    if not attendance_date or not subject_id or not lecture_no or session_type not in ("Lecture", "Practical"):
        flash("Invalid attendance session. Please try again.", "error")
        return redirect(url_for("attendance_history"))

    db=get_db()
    try:
        subject=db.execute("SELECT code,name FROM subjects WHERE id=?", (subject_id,)).fetchone()
        if not subject:
            db.rollback(); flash("Subject not found.", "error")
            return redirect(url_for("attendance_history"))

        where="subject_id=? AND attendance_date=? AND lecture_no=? AND COALESCE(session_type,'Lecture')=?"
        params=[subject_id,attendance_date,lecture_no,session_type]
        if session_type == "Practical":
            where += " AND COALESCE(batch,'')=?"
            params.append(batch)
        if lecture_time:
            where += " AND COALESCE(lecture_time,'')=?"
            params.append(lecture_time)

        # Preserve every deleted row's identifying fields and previous status in
        # the correction ledger before deleting attendance. The correction table
        # keeps attendance_id nullable, so its audit record survives the deletion.
        rows_to_delete=db.execute(
            f"SELECT id,student_id,subject_id,attendance_date,lecture_no,status FROM attendance WHERE {where}",
            tuple(params),
        ).fetchall()
        deleted=len(rows_to_delete)
        if deleted == 0:
            db.rollback()
            flash("No matching attendance session was found. It may already be deleted.", "error")
            return redirect(url_for("attendance_history"))

        deletion_reason=(request.form.get("deletion_reason") or "Admin deleted attendance session").strip()[:500]
        if not deletion_reason:
            deletion_reason="Admin deleted attendance session"
        for old_row in rows_to_delete:
            db.execute("""INSERT INTO attendance_corrections
                (attendance_id,student_id,subject_id,attendance_date,lecture_no,old_status,new_status,reason,corrected_by)
                VALUES(?,?,?,?,?,?,?,?,?)""",
                (old_row["id"], old_row["student_id"], old_row["subject_id"], old_row["attendance_date"],
                 old_row["lecture_no"], old_row["status"], "DELETED: " + str(old_row["status"]),
                 deletion_reason, session.get("user_id")))
        db.execute(f"DELETE FROM attendance WHERE {where}", tuple(params))
        db.commit()
        code=subject["code"] if subject else str(subject_id)
        log_activity("Attendance session deleted", f"{attendance_date} | {code} | {session_type} {batch} | Lecture {lecture_no} | {deleted} records")
        flash(f"Attendance session deleted successfully ({deleted} records).", "success")
    except Exception as exc:
        try: db.rollback()
        except Exception: pass
        app.logger.exception("Attendance session deletion failed")
        flash("Attendance could not be deleted. Please check the server log and try again.", "error")
    finally:
        try: db.close()
        except Exception: pass
    return redirect(url_for("attendance_history"))


@app.route("/students/promotion", methods=["GET","POST"])
@login_required
@admin_required
def student_promotion():
    db=get_db()
    current_row=db.execute("SELECT value FROM app_settings WHERE key='academic_year'").fetchone()
    current_ay=current_row["value"] if current_row and current_row["value"] else "2026-27"
    if request.method=="POST":
        from_year=request.form.get("from_year", "").strip(); from_semester=request.form.get("from_semester", type=int)
        to_year=request.form.get("to_year", "").strip(); to_semester=request.form.get("to_semester", type=int)
        new_ay=request.form.get("new_academic_year", "").strip()
        valid = (from_year in YEAR_OPTIONS and to_year in YEAR_OPTIONS and
                 from_semester in SEMESTERS_BY_YEAR.get(from_year, []) and to_semester in SEMESTERS_BY_YEAR.get(to_year, []) and bool(new_ay))
        if not valid:
            flash("Select valid current year/semester and destination year/semester, then enter the academic year.", "error")
        else:
            students=db.execute("SELECT id FROM students WHERE year=? AND semester=?",(from_year,from_semester)).fetchall(); changed=0
            for st in students:
                db.execute("INSERT INTO student_academic_history(student_id,academic_year,year,promoted_by) VALUES(?,?,?,?) ON CONFLICT(student_id,academic_year) DO NOTHING", (st["id"],new_ay,to_year,session["user_id"]))
                db.execute("UPDATE students SET year=?,semester=?,academic_year=? WHERE id=?",(to_year,to_semester,new_ay,st["id"])); changed+=1
            db.execute("INSERT INTO app_settings(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",("academic_year",new_ay))
            db.commit(); log_activity("Students promoted",f"{from_year} semester {from_semester} to {to_year} semester {to_semester} | {new_ay} | {changed} students")
            flash(f"{changed} student(s) promoted to {to_year}, Semester {to_semester} for {new_ay}.","success")
            current_ay=new_ay
    counts={y:db.execute("SELECT COUNT(*) FROM students WHERE year=?",(y,)).fetchone()[0] for y in YEAR_OPTIONS}
    history=db.execute("""SELECT h.academic_year,h.year,h.promoted_at,s.prn,s.name,s.semester FROM student_academic_history h JOIN students s ON s.id=h.student_id ORDER BY h.promoted_at DESC LIMIT 300""").fetchall()
    db.close(); return render_template("promotion.html",years=YEAR_OPTIONS,current_ay=current_ay,counts=counts,history=history,semesters_by_year=SEMESTERS_BY_YEAR)


@app.route("/reports")
@login_required
@staff_required
def reports():
    db = get_db(); subjects = teacher_or_admin_subjects(db); students = db.execute("SELECT * FROM students").fetchall()
    subject_id = request.args.get("subject_id", type=int); student_id = request.args.get("student_id", type=int)
    year = request.args.get("year", "")
    semester = request.args.get("semester", type=int) or 0
    session_type = request.args.get("session_type", "")
    batch = request.args.get("batch", "")
    if session_type not in ("", "Lecture", "Practical"):
        session_type = ""
    if batch not in ("", "Batch A", "Batch B", "Batch C", "Batch D"):
        batch = ""
    if session_type != "Practical":
        batch = ""
    from_date = request.args.get("from_date", ""); to_date = request.args.get("to_date", "")
    subjects = [sub for sub in subjects if (not year or sub["year"]==year) and (not semester or int(sub["semester"] or 0)==semester) and (not session_type or (sub["session_type"] or "Lecture")==session_type)]
    if subject_id and (not can_use_subject(db, subject_id) or not any(int(sub["id"]) == int(subject_id) for sub in subjects)):
        subject_id = None; flash("Please select a subject matching the current filters and your assignment.", "error")
    on_where = ["1=1"]; on_params = []
    student_where = ["1=1"]; student_params = []
    if session.get("role") == "teacher":
        on_where.append("a.subject_id IN (SELECT subject_id FROM subject_teachers WHERE teacher_id=?)"); on_params.append(session["user_id"])
    if subject_id: on_where.append("a.subject_id=?"); on_params.append(subject_id)
    if semester: on_where.append("a.subject_id IN (SELECT id FROM subjects WHERE semester=?)"); on_params.append(semester)
    if session_type: on_where.append("COALESCE(a.session_type,'Lecture')=?"); on_params.append(session_type)
    if session_type == "Practical" and batch: on_where.append("COALESCE(a.batch,'')=?"); on_params.append(batch)
    if from_date: on_where.append("a.attendance_date>=?"); on_params.append(from_date)
    if to_date: on_where.append("a.attendance_date<=?"); on_params.append(to_date)
    if student_id: student_where.append("s.id=?"); student_params.append(student_id)
    if year: student_where.append("s.year=?"); student_params.append(year)
    if session_type == "Practical" and batch: student_where.append("COALESCE(s.batch,'')=?"); student_params.append(batch)
    on_condition = " AND ".join(on_where)
    student_condition = " AND ".join(student_where)
    summary = db.execute(f"""SELECT s.id,s.prn,s.name,s.year,s.course,COUNT(a.id) total,
        COALESCE(SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END),0) present,
        COALESCE(SUM(CASE WHEN a.status='Absent' THEN 1 ELSE 0 END),0) absent,
        CASE WHEN COUNT(a.id)=0 THEN 0 ELSE ROUND(100.0*SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END)/COUNT(a.id),1) END pct
        FROM students s LEFT JOIN attendance a ON a.student_id=s.id AND {on_condition}
        WHERE {student_condition}
        GROUP BY s.id""", on_params + student_params).fetchall()
    summary = sorted(summary, key=_prn_sort_key)
    students = sorted(students, key=_prn_sort_key)
    detail_where = ["1=1"]; detail_params = []
    if session.get("role") == "teacher":
        detail_where.append("a.subject_id IN (SELECT subject_id FROM subject_teachers WHERE teacher_id=?)"); detail_params.append(session["user_id"])
    if subject_id: detail_where.append("a.subject_id=?"); detail_params.append(subject_id)
    if semester: detail_where.append("sub.semester=?"); detail_params.append(semester)
    if session_type: detail_where.append("COALESCE(a.session_type,'Lecture')=?"); detail_params.append(session_type)
    if session_type == "Practical" and batch: detail_where.append("COALESCE(a.batch,'')=?"); detail_params.append(batch)
    if student_id: detail_where.append("a.student_id=?"); detail_params.append(student_id)
    if year: detail_where.append("s.year=?"); detail_params.append(year)
    if from_date: detail_where.append("a.attendance_date>=?"); detail_params.append(from_date)
    if to_date: detail_where.append("a.attendance_date<=?"); detail_params.append(to_date)
    detail_condition = " AND ".join(detail_where)
    detail = db.execute(f"""SELECT a.attendance_date,a.lecture_no,a.lecture_time,COALESCE(a.session_type,'Lecture') session_type,COALESCE(a.batch,'') batch,s.prn,s.name,sub.code,sub.name subject,a.status
        FROM attendance a JOIN students s ON s.id=a.student_id JOIN subjects sub ON sub.id=a.subject_id
        WHERE {detail_condition} ORDER BY a.attendance_date DESC,a.lecture_no""", detail_params).fetchall()
    detail = sorted(detail, key=_prn_sort_key)
    detail = sorted(detail, key=lambda r: (str(r["attendance_date"] or ""), int(r["lecture_no"] or 0)), reverse=True)
    years = [r["year"] for r in db.execute("SELECT DISTINCT year FROM students WHERE year IS NOT NULL AND year<>'' ORDER BY year").fetchall()]
    subject_scope_where = "1=1"
    subject_scope_params = []
    if session.get("role") == "teacher":
        subject_scope_where = "sub.id IN (SELECT subject_id FROM subject_teachers WHERE teacher_id=?)"
        subject_scope_params = [session["user_id"]]
    subject_summary = db.execute(f"""SELECT sub.code,sub.name,COUNT(a.id) total,
        COALESCE(SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END),0) present,
        COALESCE(SUM(CASE WHEN a.status='Absent' THEN 1 ELSE 0 END),0) absent,
        CASE WHEN COUNT(a.id)=0 THEN 0 ELSE ROUND(100.0*SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END)/COUNT(a.id),1) END pct
        FROM subjects sub LEFT JOIN attendance a ON a.subject_id=sub.id AND {on_condition}
        WHERE {subject_scope_where}
        GROUP BY sub.id ORDER BY sub.code""", on_params + subject_scope_params).fetchall()
    for y in YEAR_OPTIONS:
        if y not in years: years.append(y)
    db.close()
    return render_template("reports.html", subjects=subjects, students=students, years=years, summary=summary, detail=detail, subject_summary=subject_summary,
                           subject_id=subject_id, student_id=student_id, year=year, semester=semester, semesters_by_year=SEMESTERS_BY_YEAR, session_type=session_type, batch=batch, batch_options=["Batch A","Batch B","Batch C","Batch D"], from_date=from_date, to_date=to_date)


@app.route("/reports/export.csv")
@login_required
@staff_required
def export_csv():
    log_activity("Report exported", "CSV attendance report downloaded")
    db = get_db(); subject_id = request.args.get("subject_id", type=int); student_id = request.args.get("student_id", type=int)
    year = request.args.get("year", "")
    semester = request.args.get("semester", type=int) or 0
    session_type = request.args.get("session_type", "")
    batch = request.args.get("batch", "")
    if session_type not in ("", "Lecture", "Practical"):
        session_type = ""
    if batch not in ("", "Batch A", "Batch B", "Batch C", "Batch D"):
        batch = ""
    if session_type != "Practical":
        batch = ""
    from_date = request.args.get("from_date", ""); to_date = request.args.get("to_date", "")
    on_where = ["1=1"]; on_params = []
    student_where = ["1=1"]; student_params = []
    if session.get("role") != "admin":
        on_where.append("a.subject_id IN (SELECT subject_id FROM subject_teachers WHERE teacher_id=?)"); on_params.append(session["user_id"])
    if subject_id and can_use_subject(db, subject_id): on_where.append("a.subject_id=?"); on_params.append(subject_id)
    if semester: on_where.append("a.subject_id IN (SELECT id FROM subjects WHERE semester=?)"); on_params.append(semester)
    if session_type: on_where.append("COALESCE(a.session_type,'Lecture')=?"); on_params.append(session_type)
    if session_type == "Practical" and batch: on_where.append("COALESCE(a.batch,'')=?"); on_params.append(batch)
    if from_date: on_where.append("a.attendance_date>=?"); on_params.append(from_date)
    if to_date: on_where.append("a.attendance_date<=?"); on_params.append(to_date)
    if student_id: student_where.append("s.id=?"); student_params.append(student_id)
    if year: student_where.append("s.year=?"); student_params.append(year)
    if session_type == "Practical" and batch: student_where.append("COALESCE(s.batch,'')=?"); student_params.append(batch)
    on_condition = " AND ".join(on_where)
    student_condition = " AND ".join(student_where)
    rows = db.execute(f"""SELECT s.prn,s.name,s.course,s.year,
        COUNT(a.id) total,
        COALESCE(SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END),0) present,
        COALESCE(SUM(CASE WHEN a.status='Absent' THEN 1 ELSE 0 END),0) absent,
        CASE WHEN COUNT(a.id)=0 THEN 0 ELSE ROUND(100.0*SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END)/COUNT(a.id),1) END pct
        FROM students s LEFT JOIN attendance a ON a.student_id=s.id AND {on_condition}
        WHERE {student_condition}
        GROUP BY s.id""", on_params + student_params).fetchall()
    subject_summary = db.execute(f"""SELECT sub.code,sub.name,COUNT(a.id) total, COALESCE(SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END),0) present, COALESCE(SUM(CASE WHEN a.status='Absent' THEN 1 ELSE 0 END),0) absent, CASE WHEN COUNT(a.id)=0 THEN 0 ELSE ROUND(100.0*SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END)/COUNT(a.id),1) END pct FROM subjects sub LEFT JOIN attendance a ON a.subject_id=sub.id AND {on_condition} GROUP BY sub.id ORDER BY sub.code""", on_params).fetchall()
    db.close(); out = io.StringIO(); writer = SafeCsvWriter(csv.writer(out))
    writer.writerow(["SUBJECT-WISE SUMMARY"])
    writer.writerow(["Subject","Total Attendance","Present","Absent","Attendance Percentage"])
    for r in subject_summary: writer.writerow([f"{r['code']} - {r['name']}",r['total'],r['present'],r['absent'],f"{r['pct']}%"] )
    writer.writerow([])
    writer.writerow(["STUDENT SUMMARY"])
    writer.writerow(["PRN","Name","Year","Course","Total Attendance","Present Days","Absent Days","Attendance Percentage"])
    for r in rows:
        writer.writerow([r["prn"], r["name"], r["year"], r["course"], r["total"], r["present"], r["absent"], f'{r["pct"]}%'])
    return Response(out.getvalue(), mimetype="text/csv", headers={"Content-Disposition":"attachment; filename=attendance_summary.csv"})


@app.route("/reports/export.pdf")
@login_required
@staff_required
def export_pdf():
    log_activity("Report exported", "PDF attendance report downloaded")
    db = get_db()
    subject_id = request.args.get("subject_id", type=int)
    student_id = request.args.get("student_id", type=int)
    year = request.args.get("year", "")
    semester = request.args.get("semester", type=int) or 0
    session_type = request.args.get("session_type", "")
    batch = request.args.get("batch", "")
    if session_type not in ("", "Lecture", "Practical"):
        session_type = ""
    if batch not in ("", "Batch A", "Batch B", "Batch C", "Batch D"):
        batch = ""
    if session_type != "Practical":
        batch = ""
    from_date = request.args.get("from_date", "")
    to_date = request.args.get("to_date", "")

    on_where = ["1=1"]; on_params = []
    student_where = ["1=1"]; student_params = []
    if session.get("role") != "admin":
        on_where.append("a.subject_id IN (SELECT subject_id FROM subject_teachers WHERE teacher_id=?)")
        on_params.append(session["user_id"])
    if subject_id and can_use_subject(db, subject_id):
        on_where.append("a.subject_id=?"); on_params.append(subject_id)
    elif subject_id:
        subject_id = None
    if semester:
        on_where.append("a.subject_id IN (SELECT id FROM subjects WHERE semester=?)"); on_params.append(semester)
    if session_type:
        on_where.append("COALESCE(a.session_type,'Lecture')=?"); on_params.append(session_type)
    if session_type == "Practical" and batch:
        on_where.append("COALESCE(a.batch,'')=?"); on_params.append(batch)
    if from_date:
        on_where.append("a.attendance_date>=?"); on_params.append(from_date)
    if to_date:
        on_where.append("a.attendance_date<=?"); on_params.append(to_date)
    if student_id:
        student_where.append("s.id=?"); student_params.append(student_id)
    if year:
        student_where.append("s.year=?"); student_params.append(year)
    if semester:
        student_where.append("s.semester=?"); student_params.append(semester)
    if session_type == "Practical" and batch:
        student_where.append("COALESCE(s.batch,'')=?"); student_params.append(batch)

    on_condition = " AND ".join(on_where)
    student_condition = " AND ".join(student_where)
    rows = db.execute(f"""SELECT s.prn,s.name,s.course,s.year,
        COUNT(a.id) total,
        COALESCE(SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END),0) present,
        COALESCE(SUM(CASE WHEN a.status='Absent' THEN 1 ELSE 0 END),0) absent,
        CASE WHEN COUNT(a.id)=0 THEN 0 ELSE ROUND(100.0*SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END)/COUNT(a.id),1) END pct
        FROM students s LEFT JOIN attendance a ON a.student_id=s.id AND {on_condition}
        WHERE {student_condition}
        GROUP BY s.id""", on_params + student_params).fetchall()
    rows = sorted(rows, key=_prn_sort_key)

    # When one specific student + one specific subject is selected, also
    # prepare that student's complete attendance history for the PDF.
    student_history = []
    selected_student_name = ""
    if student_id and subject_id:
        hist_where = ["a.student_id=?", "a.subject_id=?"]
        hist_params = [student_id, subject_id]
        if session.get("role") != "admin":
            hist_where.append("a.subject_id IN (SELECT subject_id FROM subject_teachers WHERE teacher_id=?)")
            hist_params.append(session["user_id"])
        if session_type:
            hist_where.append("COALESCE(a.session_type,'Lecture')=?"); hist_params.append(session_type)
        if session_type == "Practical" and batch:
            hist_where.append("COALESCE(a.batch,'')=?"); hist_params.append(batch)
        if from_date:
            hist_where.append("a.attendance_date>=?"); hist_params.append(from_date)
        if to_date:
            hist_where.append("a.attendance_date<=?"); hist_params.append(to_date)
        hist_condition = " AND ".join(hist_where)
        student_row = db.execute("SELECT name FROM students WHERE id=?", (student_id,)).fetchone()
        selected_student_name = student_row["name"] if student_row else ""
        student_history = db.execute(f"""SELECT a.attendance_date,a.lecture_no,a.lecture_time,
            COALESCE(a.session_type,'Lecture') session_type,COALESCE(a.batch,'') batch,a.status
            FROM attendance a WHERE {hist_condition}
            ORDER BY a.attendance_date ASC, COALESCE(a.lecture_no,0) ASC, COALESCE(a.lecture_time,'') ASC""", hist_params).fetchall()

    subject_name = "All Subjects"
    if subject_id:
        subject_row = db.execute("SELECT code,name FROM subjects WHERE id=?", (subject_id,)).fetchone()
        if subject_row:
            subject_name = f"{subject_row['code']} - {subject_row['name']}"

    teacher_rows = db.execute(f"SELECT DISTINCT u.full_name,u.username FROM attendance a LEFT JOIN users u ON u.id=a.scheduled_teacher_id WHERE {on_condition} AND a.scheduled_teacher_id IS NOT NULL ORDER BY u.full_name,u.username", on_params).fetchall()
    conducted_rows = db.execute(f"SELECT DISTINCT u.full_name,u.username FROM attendance a LEFT JOIN users u ON u.id=a.conducted_by_teacher_id WHERE {on_condition} AND a.conducted_by_teacher_id IS NOT NULL ORDER BY u.full_name,u.username", on_params).fetchall()
    scheduled_teacher_text = ", ".join((r["full_name"] or r["username"]) for r in teacher_rows) or "—"
    conducted_teacher_text = ", ".join((r["full_name"] or r["username"]) for r in conducted_rows) or "—"
    status_rows = db.execute(f"SELECT DISTINCT COALESCE(a.teacher_status,'Scheduled teacher present') teacher_status FROM attendance a WHERE {on_condition} ORDER BY teacher_status", on_params).fetchall()
    teacher_status_text = ", ".join(r["teacher_status"] for r in status_rows) or "—"
    threshold = attendance_threshold_value(db)
    db.close()
    class_name = year if year else "All Classes"
    meta = [("Course", "B.Pharm"), ("Year", class_name), ("Semester", f"Semester {semester}" if semester else "All Semesters"),
            ("Subject", subject_name), ("Session", session_type or "All"),
            ("Date range", f"{from_date or '—'} to {to_date or '—'}")]
    if subject_id:
        meta += [("Scheduled teacher", scheduled_teacher_text), ("Conducted by", conducted_teacher_text), ("Teacher status", teacher_status_text)]
    history_rows = None
    history_title = ""
    if student_id and subject_id:
        history_rows = [dict(h) for h in student_history]
        history_title = f"Lecture-wise history - {selected_student_name or 'Student'}"
    pdf = build_class_statement_pdf([dict(r) for r in rows], college=COLLEGE_NAME, logo_path=BASE_DIR / "static" / LOGO_FILE,
                                 threshold=threshold, meta=meta,
                                 student_history=history_rows, history_title=history_title)
    return Response(pdf, mimetype="application/pdf", headers={"Content-Disposition": "attachment; filename=attendance_report.pdf"})


@app.route("/reports/export-subject-pdfs.zip")
@login_required
@staff_required
def export_subject_pdfs():
    """Download one attendance-statement PDF per subject in a ZIP archive."""
    db = get_db()
    try:
        year = request.args.get("year", "").strip()
        semester = request.args.get("semester", type=int) or 0
        session_type = request.args.get("session_type", "")
        batch = request.args.get("batch", "")
        student_id = request.args.get("student_id", type=int)
        from_date = request.args.get("from_date", "").strip()
        to_date = request.args.get("to_date", "").strip()
        if session_type not in ("", "Lecture", "Practical"):
            session_type = ""
        if batch not in ("", "Batch A", "Batch B", "Batch C", "Batch D") or session_type != "Practical":
            batch = ""

        subject_where, subject_params = ["1=1"], []
        if session.get("role") == "teacher":
            subject_where.append("sub.id IN (SELECT subject_id FROM subject_teachers WHERE teacher_id=?)")
            subject_params.append(session["user_id"])
        if year:
            subject_where.append("sub.year=?"); subject_params.append(year)
        if semester:
            subject_where.append("sub.semester=?"); subject_params.append(semester)
        if session_type:
            subject_where.append("COALESCE(sub.session_type,'Lecture')=?"); subject_params.append(session_type)
        subjects_for_pdf = db.execute(
            "SELECT sub.id,sub.code,sub.name,sub.year,sub.semester FROM subjects sub WHERE " +
            " AND ".join(subject_where) + " ORDER BY sub.code,sub.name", subject_params
        ).fetchall()

        student_where, student_params = ["1=1"], []
        if student_id:
            student_where.append("s.id=?"); student_params.append(student_id)
        if year:
            student_where.append("s.year=?"); student_params.append(year)
        if semester:
            student_where.append("s.semester=?"); student_params.append(semester)
        if session_type == "Practical" and batch:
            student_where.append("COALESCE(s.batch,'')=?"); student_params.append(batch)
        threshold = attendance_threshold_value(db)
        archive = io.BytesIO()
        created = 0
        with zipfile.ZipFile(archive, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for sub in subjects_for_pdf:
                attendance_on, attendance_params = ["a.subject_id=?", "a.student_id=s.id"], [sub["id"]]
                if session_type:
                    attendance_on.append("COALESCE(a.session_type,'Lecture')=?"); attendance_params.append(session_type)
                if session_type == "Practical" and batch:
                    attendance_on.append("COALESCE(a.batch,'')=?"); attendance_params.append(batch)
                if from_date:
                    attendance_on.append("a.attendance_date>=?"); attendance_params.append(from_date)
                if to_date:
                    attendance_on.append("a.attendance_date<=?"); attendance_params.append(to_date)
                query = ("SELECT s.prn,s.name,s.course,s.year, COUNT(a.id) total, "
                    "COALESCE(SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END),0) present, "
                    "COALESCE(SUM(CASE WHEN a.status='Absent' THEN 1 ELSE 0 END),0) absent, "
                    "CASE WHEN COUNT(a.id)=0 THEN 0 ELSE ROUND(100.0*SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END)/COUNT(a.id),1) END pct "
                    + "FROM students s LEFT JOIN attendance a ON " + " AND ".join(attendance_on)
                    + " WHERE " + " AND ".join(student_where) + " GROUP BY s.id")
                rows = db.execute(query, attendance_params + student_params).fetchall()
                rows = sorted(rows, key=_prn_sort_key)
                meta = [("Course", "B.Pharm"), ("Year", year or "All Classes"),
                        ("Semester", f"Semester {semester}" if semester else "All Semesters"),
                        ("Subject", f"{sub['code']} - {sub['name']}"),
                        ("Session", session_type or "All"),
                        ("Date range", f"{from_date or '—'} to {to_date or '—'}")]
                pdf_bytes = build_class_statement_pdf([dict(r) for r in rows], college=COLLEGE_NAME,
                    logo_path=BASE_DIR / "static" / LOGO_FILE, threshold=threshold, meta=meta)
                safe_code = re.sub(r"[^A-Za-z0-9_-]+", "_", str(sub["code"] or sub["id"])).strip("_") or str(sub["id"])
                safe_name = re.sub(r"[^A-Za-z0-9_-]+", "_", str(sub["name"] or "subject")).strip("_")[:55]
                zf.writestr(f"{safe_code}_{safe_name or 'subject'}_attendance.pdf", pdf_bytes)
                created += 1
        if created == 0:
            return Response("No subjects found for the selected filters.", status=404, mimetype="text/plain")
        archive.seek(0)
        log_activity("Report exported", f"Subject-wise PDF reports downloaded ({created} files)")
        return Response(archive.getvalue(), mimetype="application/zip",
            headers={"Content-Disposition": "attachment; filename=subject_wise_attendance_reports.zip"})
    finally:
        db.close()

@app.post("/activity-log/clear")
@login_required
@admin_required
def clear_activity_log():
    mode = request.form.get("mode", "all")
    db = get_db()
    if mode == "30":
        cutoff = (datetime.utcnow() - timedelta(days=30)).isoformat()
        db.execute("DELETE FROM activity_logs WHERE created_at < ?", (cutoff,))
        message = "Activity log older than 30 days cleared."
    else:
        db.execute("DELETE FROM activity_logs")
        message = "Activity log cleared."
    db.commit(); db.close()
    flash(message, "success")
    return redirect(url_for("activity_log"))


@app.post("/notifications/clear")
@login_required
def clear_notifications():
    mode = request.form.get("mode", "all")
    db = get_db()
    if mode == "30":
        cutoff = (datetime.utcnow() - timedelta(days=30)).isoformat()
        db.execute("DELETE FROM notifications WHERE user_id=? AND created_at < ?", (session["user_id"], cutoff))
        message = "Notifications older than 30 days cleared."
    else:
        db.execute("DELETE FROM notifications WHERE user_id=?", (session["user_id"],))
        message = "Notifications cleared."
    db.commit(); db.close()
    flash(message, "success")
    return redirect(url_for("notifications"))


@app.post("/leaves/clear")
@login_required
def clear_leaves():
    mode = request.form.get("mode", "all")
    db = get_db()
    if session.get("role") == "student":
        student = current_student(db)
        if not student:
            db.close(); flash("Student record not found.", "error"); return redirect(url_for("leaves"))
        if mode == "30":
            cutoff = (datetime.utcnow() - timedelta(days=30)).isoformat()
            db.execute("DELETE FROM leave_applications WHERE student_id=? AND created_at < ?", (student["id"], cutoff))
            message = "Leave applications older than 30 days cleared."
        else:
            db.execute("DELETE FROM leave_applications WHERE student_id=?", (student["id"],))
            message = "Your leave applications cleared."
    else:
        if session.get("role") != "admin":
            db.close(); flash("Only the administrator can clear leave applications.", "error"); return redirect(url_for("leaves"))
        if mode == "30":
            cutoff = (datetime.utcnow() - timedelta(days=30)).isoformat()
            db.execute("DELETE FROM leave_applications WHERE created_at < ?", (cutoff,))
            message = "Leave applications older than 30 days cleared."
        else:
            db.execute("DELETE FROM leave_applications")
            message = "Leave applications cleared."
    db.commit(); db.close()
    flash(message, "success")
    return redirect(url_for("leaves"))


@app.route("/activity-log")
@login_required
@admin_required
def activity_log():
    db=get_db(); q=request.args.get("q","").strip(); params=[]
    where="1=1"
    if q: where="(l.action LIKE ? OR l.details LIKE ? OR u.username LIKE ?)"; params=[f"%{q}%"]*3
    rows=db.execute(f"SELECT l.created_at,l.action,l.details,COALESCE(u.username,'System') username FROM activity_logs l LEFT JOIN users u ON u.id=l.user_id WHERE {where} ORDER BY l.created_at DESC LIMIT 300",params).fetchall(); db.close()
    return render_template("activity_log.html",rows=rows,q=q)

@app.route("/profile", methods=["GET","POST"])
@login_required
def profile():
    db=get_db(); user=db.execute("SELECT * FROM users WHERE id=?",(session["user_id"],)).fetchone()
    if request.method=="POST":
        full_name=request.form.get("full_name","").strip(); mobile=request.form.get("mobile","").strip(); new_password=request.form.get("new_password","")
        if new_password:
            err=validate_password(new_password)
            if not check_password_hash(user["password"], request.form.get("current_password","")):
                err="Your current password is not correct."
            if err: flash(err,"error")
            else:
                db.execute("UPDATE users SET full_name=?,mobile=?,password=? WHERE id=?",(full_name,mobile,generate_password_hash(new_password),session["user_id"])); db.commit(); session["full_name"]=full_name; log_activity("Password changed","Password changed from profile"); revoke_login_sessions(session["user_id"], session.get("sid")); flash("Profile updated.","success")
        else:
            db.execute("UPDATE users SET full_name=?,mobile=? WHERE id=?",(full_name,mobile,session["user_id"])); db.commit(); session["full_name"]=full_name; log_activity("Profile updated","Name/mobile updated"); flash("Profile updated.","success")
        user=db.execute("SELECT * FROM users WHERE id=?",(session["user_id"],)).fetchone()
    db.close(); return render_template("profile.html",user=user)

@app.route("/settings", methods=["GET","POST"])
@login_required
@admin_required
def settings():
    db=get_db()
    if request.method=="POST":
        threshold=request.form.get("threshold","75").strip()
        academic_year=request.form.get("academic_year","").strip()
        if threshold.isdigit() and 0 <= int(threshold) <= 100:
            for k,v in (("attendance_threshold",threshold),("academic_year",academic_year)):
                db.execute("INSERT INTO app_settings(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",(k,v))
            db.commit(); log_activity("Settings updated",f"Threshold {threshold}, academic year {academic_year}"); flash("Settings saved.","success")
        else: flash("Threshold must be between 0 and 100.","error")
    vals={r["key"]:r["value"] for r in db.execute("SELECT key,value FROM app_settings").fetchall()}; db.close()
    return render_template("settings.html",settings=vals)

@app.route("/reports/export.xlsx")
@login_required
@staff_required
def export_xlsx():
    log_activity("Report exported", "Excel attendance report downloaded")
    if not OPENPYXL_AVAILABLE:
        return Response("Excel export is not installed on this server.", status=500, mimetype="text/plain")
    db = get_db(); subject_id=request.args.get("subject_id", type=int); student_id=request.args.get("student_id", type=int)
    year=request.args.get("year", "")
    semester=request.args.get("semester", type=int) or 0
    session_type=request.args.get("session_type", "")
    batch=request.args.get("batch", "")
    if session_type not in ("", "Lecture", "Practical"):
        session_type = ""
    if batch not in ("", "Batch A", "Batch B", "Batch C", "Batch D"):
        batch = ""
    if session_type != "Practical":
        batch = ""; from_date=request.args.get("from_date", ""); to_date=request.args.get("to_date", "")
    on_where=["1=1"]; on_params=[]; student_where=["1=1"]; student_params=[]
    if session.get("role") != "admin": on_where.append("a.subject_id IN (SELECT subject_id FROM subject_teachers WHERE teacher_id=?)"); on_params.append(session["user_id"])
    if subject_id and can_use_subject(db, subject_id): on_where.append("a.subject_id=?"); on_params.append(subject_id)
    elif subject_id: subject_id=None
    if semester: on_where.append("a.subject_id IN (SELECT id FROM subjects WHERE semester=?)"); on_params.append(semester)
    if session_type: on_where.append("COALESCE(a.session_type,'Lecture')=?"); on_params.append(session_type)
    if session_type == "Practical" and batch: on_where.append("COALESCE(a.batch,'')=?"); on_params.append(batch)
    if from_date: on_where.append("a.attendance_date>=?"); on_params.append(from_date)
    if to_date: on_where.append("a.attendance_date<=?"); on_params.append(to_date)
    if student_id: student_where.append("s.id=?"); student_params.append(student_id)
    if year: student_where.append("s.year=?"); student_params.append(year)
    if session_type == "Practical" and batch: student_where.append("COALESCE(s.batch,'')=?"); student_params.append(batch)
    subject_name = "All Subjects"
    if subject_id:
        sr = db.execute("SELECT code,name FROM subjects WHERE id=?", (subject_id,)).fetchone()
        if sr: subject_name = f"{sr['code']} - {sr['name']}"
    rows=db.execute(f"""SELECT s.prn,s.name,s.year,s.semester,s.course,COUNT(a.id) total,
        COALESCE(SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END),0) present,
        COALESCE(SUM(CASE WHEN a.status='Absent' THEN 1 ELSE 0 END),0) absent,
        CASE WHEN COUNT(a.id)=0 THEN 0 ELSE ROUND(100.0*SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END)/COUNT(a.id),1) END pct
        FROM students s LEFT JOIN attendance a ON a.student_id=s.id AND {' AND '.join(on_where)}
        WHERE {' AND '.join(student_where)} GROUP BY s.id""", on_params+student_params).fetchall()
    rows = sorted(rows, key=_prn_sort_key)
    subject_summary = db.execute(f"""SELECT sub.code,sub.name,COUNT(a.id) total, COALESCE(SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END),0) present, COALESCE(SUM(CASE WHEN a.status='Absent' THEN 1 ELSE 0 END),0) absent, CASE WHEN COUNT(a.id)=0 THEN 0 ELSE ROUND(100.0*SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END)/COUNT(a.id),1) END pct FROM subjects sub LEFT JOIN attendance a ON a.subject_id=sub.id AND {' AND '.join(on_where)} GROUP BY sub.id ORDER BY sub.code""", on_params).fetchall()
    db.close()
    wb=Workbook(); ws=wb.active; ws.title="Attendance Summary"
    sws=wb.create_sheet("Subject Summary")
    sws.append(["Subject","Total Attendance","Present","Absent","Attendance Percentage"])
    for r in subject_summary: sws.append([f"{r['code']} - {r['name']}",r['total'],r['present'],r['absent'],f"{r['pct']}%"] )
    for c in sws[1]: c.font=Font(bold=True)
    for col,width in zip("ABCDE",[35,20,15,15,25]): sws.column_dimensions[col].width=width

    ws.row_dimensions[1].height = 72
    ws.merge_cells("B1:I1")
    ws["B1"] = COLLEGE_NAME
    ws["B1"].font = Font(bold=True, size=22)
    ws["B1"].alignment = Alignment(horizontal="left", vertical="center")
    ws.merge_cells("B2:I2")
    ws["B2"] = "Attendance Report"
    ws["B2"].font = Font(bold=True, size=13)
    ws["B2"].alignment = Alignment(horizontal="left", vertical="center")
    ws.merge_cells("B3:I3")
    ws["B3"] = f"Subject: {subject_name}"
    ws["B3"].font = Font(size=12)
    ws["B3"].alignment = Alignment(horizontal="left", vertical="center")
    ws.merge_cells("B4:I4")
    ws["B4"] = f"Class: {year if year else 'All Classes'} | Semester: {('Semester '+str(semester)) if semester else 'All Semesters'}"
    ws["B4"].font = Font(size=12)
    ws["B4"].alignment = Alignment(horizontal="left", vertical="center")
    filter_text = [f"Session: {session_type or 'All'}"]
    if from_date or to_date: filter_text.append(f"Date Range: {from_date or '—'} to {to_date or '—'}")
    if filter_text:
        ws.merge_cells("B5:I5")
        ws["B5"] = " | ".join(filter_text)
        ws["B5"].font = Font(size=10)
        ws["B5"].alignment = Alignment(horizontal="left", vertical="center")
    header_row = 7
    if XLImage is not None:
        logo_path = BASE_DIR / "static" / LOGO_FILE
        if logo_path.exists():
            xl_logo = XLImage(str(logo_path))
            xl_logo.width = 95
            xl_logo.height = 95
            ws.add_image(xl_logo, "A1")
    headers=["PRN","Name","Year","Semester","Course","Total Attendance","Present Days","Absent Days","Attendance Percentage"]
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(header_row, col_idx, header)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for r_idx, r in enumerate(rows, header_row + 1):
        values=[r["prn"],r["name"],r["year"],r["semester"] if "semester" in r.keys() else semester,r["course"],r["total"],r["present"],r["absent"],f'{r["pct"]}%']
        for c_idx, value in enumerate(values, 1):
            ws.cell(r_idx, c_idx, value)
    for col,width in zip("ABCDEFGHI",[15,28,16,13,16,20,15,15,25]): ws.column_dimensions[col].width=width
    ws.row_dimensions[header_row].height = 30
    for row in ws.iter_rows(min_row=header_row):
        for cell in row: cell.alignment=Alignment(vertical="center")
    ws.freeze_panes = f"A{header_row+1}"
    ws.print_title_rows = f"1:{header_row}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    sanitize_workbook(wb)
    out=io.BytesIO(); wb.save(out); out.seek(0)
    return Response(out.getvalue(),mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",headers={"Content-Disposition":"attachment; filename=attendance_summary.xlsx"})


@app.route("/reports/print")
@login_required
@staff_required
def print_report():
    return redirect(url_for("reports", **request.args.to_dict()))


_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def attendance_threshold_value(db):
    row = db.execute("SELECT value FROM app_settings WHERE key='attendance_threshold'").fetchone()
    try:
        return float(row["value"]) if row and row["value"] else 75.0
    except (TypeError, ValueError):
        return 75.0


def lectures_needed(present, total, threshold):
    """How many lectures in a row a student must attend (no absence) to reach the required percentage."""
    t = threshold / 100.0
    if total <= 0 or t >= 1 or present / total >= t:
        return 0
    return int(math.ceil((t * total - present) / (1 - t) - 1e-9))


def attendance_email_content(student, report, kind, threshold):
    """Returns (subject, plain_text, html) for the weekly report ('weekly') or the low attendance alert ('low')."""
    overall = report_overall(report)
    total = sum(r["total"] for r in report)
    present = sum(r["present"] for r in report)
    as_on = now_ist().strftime("%d %B %Y")
    below = overall < threshold
    if kind == "low":
        subject = f"Low Attendance Alert - {COLLEGE_NAME}"
        intro = f"Your overall attendance is {overall:.1f}%, which is below the required {threshold:.0f}% (as on {as_on})."
    else:
        subject = f"Weekly Attendance Report - {COLLEGE_NAME}"
        intro = f"Here is your attendance report from {COLLEGE_NAME} (as on {as_on})."
    lines = []
    for r in report:
        flag = f"  <-- below {threshold:.0f}%" if r["pct"] < threshold else ""
        lines.append(f"{r['code']} - {r['name']}: {r['present']}/{r['total']} ({r['pct']:.1f}%){flag}")
    if kind == "low":
        need = lectures_needed(present, total, threshold)
        advice = (f"To reach {threshold:.0f}%, you need to attend about {need} more lectures in a row without any absence."
                  if need else "")
        closing = "Please attend your lectures regularly."
    else:
        advice = ("Your overall attendance is below the required limit. Please attend your lectures regularly."
                  if below else "Your overall attendance is above the required limit. Keep it up!")
        closing = "Your subject-wise PDF report is attached to this email."
    footer = ("If any record looks wrong, please contact your subject teacher.\n"
              "This is an automatic email from the Attendance Office. Please do not reply.")
    text = "\n".join([f"Dear {student['name']},", "", intro, "", "Subject-wise attendance (present/total):", *lines, "",
                      f"Overall: {present}/{total} ({overall:.1f}%)  |  Required: {threshold:.0f}%"]
                      + ([advice] if advice else []) + ["", closing, "", footer, "", f"- Attendance Office, {COLLEGE_NAME}"])
    rows_html = "".join(
        f"<tr><td style='padding:7px 10px;border:1px solid #dbe4ee'>{xml_escape(str(r['code']))} - {xml_escape(str(r['name']))}</td>"
        f"<td style='padding:7px 10px;border:1px solid #dbe4ee;text-align:center'>{r['present']}/{r['total']}</td>"
        f"<td style='padding:7px 10px;border:1px solid #dbe4ee;text-align:center;font-weight:bold;color:{'#b4233b' if r['pct'] < threshold else '#087a53'}'>{r['pct']:.1f}%</td></tr>"
        for r in report)
    html = (f"<div style='font-family:Arial,sans-serif;font-size:14px;color:#183b64;max-width:620px'>"
            f"<p>Dear {xml_escape(str(student['name']))},</p><p>{xml_escape(intro)}</p>"
            "<table style='border-collapse:collapse;width:100%;font-size:13px'><tr style='background:#07579a;color:#fff'>"
            "<th style='padding:8px 10px;text-align:left'>Subject</th><th style='padding:8px 10px'>Present/Total</th><th style='padding:8px 10px'>%</th></tr>"
            f"{rows_html}</table>"
            f"<p><b>Overall: {present}/{total} ({overall:.1f}%)</b> &nbsp;|&nbsp; Required: {threshold:.0f}%</p>"
            + (f"<p>{xml_escape(advice)}</p>" if advice else "")
            + f"<p>{xml_escape(closing)}</p>"
            f"<p style='color:#5b7087;font-size:12px'>If any record looks wrong, please contact your subject teacher.<br>"
            f"This is an automatic email from the Attendance Office. Please do not reply.</p>"
            f"<p>- Attendance Office, {xml_escape(COLLEGE_NAME)}</p></div>")
    return subject, text, html


def send_attendance_email(db, student_id, kind):
    """Email one student their report ('weekly', with the subject-wise PDF) or the low attendance alert ('low').
    Returns (ok, message). Teachers only ever see numbers for their own subjects, same as the pages."""
    data = load_student_for_report(db, student_id)
    if not data:
        return False, "Student not found or you are not assigned to this student's year."
    student, contact, report = data
    email = ((contact["email"] if contact else "") or "").strip()
    if not email or not _EMAIL_RE.match(email):
        return False, "This student has no valid email address."
    if not report:
        return False, "There is no attendance record to send yet."
    threshold = attendance_threshold_value(db)
    if kind == "low" and report_overall(report) >= threshold:
        return False, f"{student['name']} is not below the {threshold:.0f}% limit."
    if not email_is_configured():
        return False, "Email is not set up yet. Add BREVO_API_KEY and BREVO_SENDER_EMAIL in Render."
    subject, text, html = attendance_email_content(student, report, kind, threshold)
    attachments = None
    if kind == "weekly":
        safe = re.sub(r"[^A-Za-z0-9_-]+", "_", str(student["prn"] or student["roll_no"])).strip("_") or "student"
        attachments = [(f"Attendance_Report_{safe}.pdf", build_subject_wise_pdf(student, report, threshold))]
    try:
        send_email_message(email, student["name"], subject, text, html, attachments)
    except Exception as exc:
        app.logger.warning("Attendance email failed for student %s: %s", student_id, exc)
        return False, "Email could not be sent. Check the Brevo key, the verified sender and the daily limit (300)."
    db.execute("INSERT INTO email_log(student_id,kind,sent_at) VALUES(?,?,?)", (student_id, kind, datetime.utcnow().isoformat()))
    db.commit()
    log_activity("Attendance email sent", f"{'Report' if kind == 'weekly' else 'Low attendance alert'} emailed to student {student['roll_no']}")
    return True, f"Email sent to {email}."


def email_sent_label(sent_at):
    """'Emailed today' / 'Emailed 3 days ago' from a stored UTC timestamp, or ''."""
    if not sent_at:
        return ""
    try:
        days = (datetime.utcnow() - datetime.fromisoformat(str(sent_at))).days
    except ValueError:
        return ""
    return "Emailed today" if days <= 0 else (f"Emailed {days} day ago" if days == 1 else f"Emailed {days} days ago")


@app.route("/send-report/<int:student_id>/email", methods=["POST"])
@login_required
@staff_required
@limiter.limit("100 per hour")
def send_report_email(student_id):
    db = get_db()
    ok, message = send_attendance_email(db, student_id, "weekly")
    db.close()
    flash(message, "success" if ok else "error")
    return redirect(url_for("send_report", student_id=student_id))


@app.route("/low-attendance/<int:student_id>/email", methods=["POST"])
@login_required
@staff_required
@limiter.limit("100 per hour")
def low_attendance_email(student_id):
    db = get_db()
    ok, message = send_attendance_email(db, student_id, "low")
    db.close()
    flash(message, "success" if ok else "error")
    return redirect(url_for("low_attendance", year=request.form.get("year", "").strip()))


# Tables are restored parents first, so foreign keys always find their parent row.
RESTORE_ORDER = ["departments", "students", "users", "subjects", "subject_teachers", "attendance",
                 "attendance_corrections", "timetable", "notifications", "leave_applications", "qr_sessions",
                 "app_settings", "student_academic_history", "deleted_students", "activity_logs"]


def _table_columns(db, table):
    if db.postgres:
        rows = db.execute("SELECT column_name FROM information_schema.columns WHERE table_name=?", (table,)).fetchall()
        return {r["column_name"] for r in rows}
    return {r[1] for r in db.execute(f"PRAGMA table_info({table})").fetchall()}


# A brand-new database is seeded by init_db() with these sample rows; they do not count as real data.
_DEMO_STUDENTS = {("01", "Rahul Sharma"), ("02", "Priya Patel"), ("03", "Ankit Yadav"), ("04", "Sneha Gupta"), ("05", "Rohan Singh")}
_DEMO_SUBJECTS = {("BP503T", "Pharmacology-II"), ("BP502T", "Pharmacognosy-II"), ("BP501T", "Medicinal Chemistry-II")}


def restore_database_is_fresh(db):
    """A restore may only run on a fresh database (nothing but the built-in sample rows and the admin),
    so it can never mix with or erase live data."""
    for table in ("attendance", "timetable"):
        if db.execute(f"SELECT COUNT(*) c FROM {table}").fetchone()["c"]:
            return False
    if db.execute("SELECT COUNT(*) c FROM users WHERE role<>'admin'").fetchone()["c"]:
        return False
    if {(r["roll_no"], r["name"]) for r in db.execute("SELECT roll_no,name FROM students").fetchall()} - _DEMO_STUDENTS:
        return False
    if {(r["code"], r["name"]) for r in db.execute("SELECT code,name FROM subjects").fetchall()} - _DEMO_SUBJECTS:
        return False
    return True


def restore_from_backup(db, payload):
    """Replace the (fresh) database content with the backup. The caller commits or rolls back."""
    tables = payload.get("tables") if isinstance(payload, dict) else None
    if not isinstance(tables, dict):
        raise ValueError("This is not a valid backup file.")
    if not any(isinstance(u, dict) and u.get("role") == "admin" for u in tables.get("users", [])):
        raise ValueError("The backup has no admin account, so nobody could log in after restoring.")
    columns = {t: _table_columns(db, t) for t in RESTORE_ORDER}
    for table in reversed(RESTORE_ORDER):
        db.execute(f"DELETE FROM {table}")
    counts = {}
    for table in RESTORE_ORDER:
        rows = tables.get(table) or []
        groups = {}
        for row in rows:
            if not isinstance(row, dict):
                raise ValueError(f"Invalid data in table {table}.")
            if table == "users" and "password" not in row:
                row["password"] = generate_password_hash(secrets.token_urlsafe(24))  # emailed backups have no passwords: use Forgot Password
            cols = tuple(c for c in row.keys() if c in columns[table] and _SAFE_COLUMN.match(c))
            if not cols:
                continue
            values = []
            for c in cols:
                v = row[c]
                if isinstance(v, bool):
                    v = int(v)
                elif isinstance(v, (dict, list)):
                    v = json.dumps(v)
                values.append(v)
            groups.setdefault(cols, []).append(tuple(values))
        for cols, batch in groups.items():
            marks = ",".join("?" for _ in cols)
            db.executemany(f"INSERT INTO {table}({','.join(cols)}) VALUES({marks})", batch)
        counts[table] = len(rows)
    if db.postgres:
        # Rows were inserted with their old ids, so move each id counter past the highest one.
        for table in RESTORE_ORDER:
            if "id" in columns[table]:
                db.execute(f"SELECT setval(pg_get_serial_sequence('{table}','id'), "
                           f"COALESCE((SELECT MAX(id) FROM {table}),0)+1, false)")
    return counts


@app.route("/admin/restore", methods=["GET", "POST"])
@login_required
@admin_required
@limiter.limit("5 per hour", methods=["POST"])
def restore_backup():
    db = get_db()
    fresh = restore_database_is_fresh(db)
    if request.method == "POST":
        admin = db.execute("SELECT password FROM users WHERE id=?", (session["user_id"],)).fetchone()
        upload = request.files.get("backup")
        if not admin or not check_password_hash(admin["password"], request.form.get("password", "")):
            flash("Your admin password is not correct.", "error")
        elif request.form.get("confirm") != "yes":
            flash("Please tick the confirmation box.", "error")
        elif not fresh:
            flash("Restore works only on a fresh, empty database. This one already has students, subjects or attendance.", "error")
        elif not upload or not upload.filename:
            flash("Please choose the backup .zip file.", "error")
        else:
            try:
                with zipfile.ZipFile(upload.stream) as z:
                    info = z.getinfo("attendance_backup.json")
                    if info.file_size > 60 * 1024 * 1024:
                        raise ValueError("The backup file is too large.")
                    payload = json.loads(z.read(info).decode("utf-8"))
                counts = restore_from_backup(db, payload)
                db.commit()
                db.close()
                session.clear()
                flash("Restore complete: " + ", ".join(f"{counts[t]} {t.replace('_', ' ')}" for t in
                      ("students", "users", "subjects", "attendance") if t in counts) +
                      ". Please log in with the admin username and password from the backup.", "success")
                return redirect(url_for("login"))
            except (ValueError, KeyError, zipfile.BadZipFile, json.JSONDecodeError) as exc:
                db.rollback()
                flash(f"Restore failed: {exc}" if isinstance(exc, ValueError) else "Restore failed: this is not a valid backup zip.", "error")
            except Exception:
                db.rollback()
                app.logger.exception("Restore failed")
                flash("Restore failed because of a database error. Nothing was changed.", "error")
    db.close()
    return render_template("restore_backup.html", fresh=fresh)


BACKUP_TABLES = ["departments", "students", "users", "subjects", "subject_teachers", "attendance", "timetable",
                 "notifications", "leave_applications", "qr_sessions", "app_settings",
                 "student_academic_history", "attendance_corrections", "deleted_students", "activity_logs"]


def build_backup_zip(strip_credentials=False):
    db = get_db()
    payload = {"college": COLLEGE_NAME, "generated_at": datetime.utcnow().isoformat() + "Z", "tables": {}}
    for table in BACKUP_TABLES:
        rows = [dict(r) for r in db.execute(f"SELECT * FROM {table}").fetchall()]
        if table == "users":
            for r in rows:  # never put 2-step secrets in a backup file
                r.pop("totp_secret", None); r.pop("totp_recovery", None)
                if strip_credentials:  # the emailed copy never contains password hashes
                    r.pop("password", None); r["totp_enabled"] = 0
        payload["tables"][table] = rows
    db.close()
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("attendance_backup.json", json.dumps(payload, indent=2, default=str))
        z.writestr("README.txt", f"{COLLEGE_NAME} Attendance backup\nGenerated: {payload['generated_at']}\n\n"
                   "This backup contains application data from all attendance tables. Keep it securely because it "
                   "includes account and attendance information.\n")
    return out.getvalue(), payload["generated_at"]


@app.route("/admin/backup")
@login_required
@admin_required
def backup_database():
    log_activity("Backup downloaded", "Full backup zip downloaded")
    data, _ = build_backup_zip()
    return Response(data, mimetype="application/zip", headers={"Content-Disposition": "attachment; filename=personal_attendance_backup.zip"})


@app.route("/cron/backup", methods=["GET", "POST"])
@csrf.exempt
@limiter.limit("6 per hour")
def cron_backup():
    """Called by a free scheduler (cron-job.org). Emails a fresh backup zip to the admin."""
    expected = os.environ.get("CRON_SECRET", "").strip()
    given = request.headers.get("X-Cron-Key") or request.args.get("key") or request.form.get("key") or ""
    if len(expected) < 16 or not hmac.compare_digest(given.encode("utf-8"), expected.encode("utf-8")):
        return Response("Not allowed.", status=403, mimetype="text/plain")
    recipient = os.environ.get("BACKUP_EMAIL", "").strip() or os.environ.get("ADMIN_EMAIL", "").strip()
    if not recipient:
        db = get_db()
        row = db.execute("SELECT email FROM users WHERE role='admin' AND email<>'' ORDER BY id LIMIT 1").fetchone()
        db.close()
        recipient = row["email"] if row else ""
    if not recipient:
        return Response("No backup email address. Set BACKUP_EMAIL.", status=400, mimetype="text/plain")
    if not email_is_configured():
        return Response("Email is not configured.", status=503, mimetype="text/plain")
    data, stamp = build_backup_zip(strip_credentials=True)
    if len(data) > 20 * 1024 * 1024:
        return Response("Backup is larger than 20 MB, too big to email.", status=413, mimetype="text/plain")
    day = now_ist().strftime("%Y-%m-%d")
    try:
        send_email_message(recipient, "Admin", f"{COLLEGE_NAME} - Attendance Backup {day}",
                           f"Automatic attendance backup generated at {stamp}.\nKeep this file safe: it contains student data. It does not contain passwords (use Admin > Backup for a full copy).",
                           None, [(f"personal_attendance_backup_{day}.zip", data)])
    except Exception as exc:
        app.logger.warning("Backup email failed: %s", exc)
        return Response("Backup email could not be sent.", status=502, mimetype="text/plain")
    return {"status": "sent", "bytes": len(data)}


@app.route("/student-approvals")
@login_required
@admin_required
def student_approvals():
    db = get_db()
    pending = db.execute("""SELECT u.id user_id,u.username,u.full_name,u.email,u.mobile,
                          u.security_question,u.created_at,
                          s.id student_id,s.prn,s.name,s.course,s.roll_no,s.year,s.semester,s.division,s.batch
                          FROM users u LEFT JOIN students s ON s.id=u.student_id
                          WHERE u.role='student' AND u.approved=0
                          ORDER BY u.created_at,u.id""").fetchall()
    db.close()
    return render_template("student_approvals.html", pending_students=pending)

@app.route("/student-approvals/<int:user_id>/approve", methods=["POST"])
@login_required
@admin_required
def approve_student(user_id):
    db = get_db()
    user = db.execute("SELECT id,student_id FROM users WHERE id=? AND role='student' AND approved=0", (user_id,)).fetchone()
    if not user:
        db.close(); flash("Pending student account not found.", "error"); return redirect(url_for("student_approvals"))
    db.execute("UPDATE users SET approved=1 WHERE id=? AND role='student'", (user_id,))
    if user["student_id"]:
        db.execute("UPDATE students SET blocked=0 WHERE id=?", (user["student_id"],))
    db.commit(); db.close()
    flash("Student account approved. The student can now log in.", "success")
    return redirect(url_for("student_approvals"))

@app.route("/student-approvals/<int:user_id>/reject", methods=["POST"])
@login_required
@admin_required
def reject_student(user_id):
    db = get_db()
    user = db.execute("SELECT id,student_id FROM users WHERE id=? AND role='student' AND approved=0", (user_id,)).fetchone()
    if not user:
        db.close(); flash("Pending student account not found.", "error"); return redirect(url_for("student_approvals"))
    student_id = user["student_id"]
    db.execute("DELETE FROM users WHERE id=?", (user_id,))
    if student_id:
        attendance = db.execute("SELECT 1 FROM attendance WHERE student_id=? LIMIT 1", (student_id,)).fetchone()
        student = db.execute("SELECT blocked,roll_no FROM students WHERE id=?", (student_id,)).fetchone()
        if student and int(student["blocked"] or 0) == 1 and str(student["roll_no"] or "").startswith("PENDING-") and not attendance:
            db.execute("DELETE FROM students WHERE id=?", (student_id,))
    db.commit(); db.close()
    flash("Student registration request rejected.", "success")
    return redirect(url_for("student_approvals"))


@app.route("/student-report/<int:student_id>")
@login_required
@staff_required
def student_report(student_id):
    db = get_db(); student = db.execute("SELECT * FROM students WHERE id=?", (student_id,)).fetchone()
    if not student:
        db.close(); flash("Student not found.", "error"); return redirect(url_for("students"))
    if session.get("role") == "teacher":
        assigned = db.execute("SELECT 1 FROM subject_teachers st JOIN subjects sub ON sub.id=st.subject_id WHERE st.teacher_id=? AND sub.year=? LIMIT 1", (session["user_id"], student["year"])).fetchone()
        if not assigned:
            db.close(); abort(403)
        rows = db.execute("""SELECT sub.code,sub.name,COUNT(a.id) total,
            COALESCE(SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END),0) present,
            COALESCE(SUM(CASE WHEN a.status='Absent' THEN 1 ELSE 0 END),0) absent,
            CASE WHEN COUNT(a.id)=0 THEN 0 ELSE ROUND(100.0*SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END)/COUNT(a.id),1) END pct
            FROM subjects sub LEFT JOIN attendance a ON a.subject_id=sub.id AND a.student_id=?
            JOIN subject_teachers st ON st.subject_id=sub.id AND st.teacher_id=?
            WHERE sub.year=? GROUP BY sub.id ORDER BY sub.code""", (student_id, session["user_id"], student["year"])).fetchall()
    else:
        rows = db.execute("""SELECT sub.code,sub.name,COUNT(a.id) total,
            COALESCE(SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END),0) present,
            COALESCE(SUM(CASE WHEN a.status='Absent' THEN 1 ELSE 0 END),0) absent,
            CASE WHEN COUNT(a.id)=0 THEN 0 ELSE ROUND(100.0*SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END)/COUNT(a.id),1) END pct
            FROM subjects sub LEFT JOIN attendance a ON a.subject_id=sub.id AND a.student_id=?
            GROUP BY sub.id ORDER BY sub.code""", (student_id,)).fetchall()
    db.close(); return render_template("student_report.html", student=student, rows=rows)




def notify_user(db, user_id, title, message, link=""):
    if user_id:
        db.execute("INSERT INTO notifications(user_id,title,message,link,created_at) VALUES(?,?,?,?,?)",
                   (user_id, title, message, link, datetime.utcnow().isoformat()))


def notify_students_for_attendance(db, subject_id, year, attendance_date, lecture_no):
    rows = db.execute("""SELECT u.id,u.student_id,s.name,sub.name subject FROM users u
        JOIN students s ON s.id=u.student_id JOIN subjects sub ON sub.id=?
        WHERE u.role='student' AND u.approved=1 AND s.year=?""", (subject_id,year)).fetchall()
    for r in rows:
        notify_user(db, r["id"], "Attendance Updated", f"Attendance was recorded for {r['subject']} (Lecture {lecture_no}) on {attendance_date}.", "/student/dashboard")


def create_low_attendance_notifications(db):
    threshold_row = db.execute("SELECT value FROM app_settings WHERE key='attendance_threshold'").fetchone()
    try: threshold = float(threshold_row["value"] if threshold_row and threshold_row["value"] else 75)
    except Exception: threshold = 75
    rows = db.execute("""SELECT u.id, s.name, COUNT(a.id) total,
        COALESCE(SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END),0) present
        FROM users u JOIN students s ON s.id=u.student_id LEFT JOIN attendance a ON a.student_id=s.id
        WHERE u.role='student' GROUP BY u.id,s.name HAVING COUNT(a.id)>0""").fetchall()
    for r in rows:
        pct = 100.0*r["present"]/r["total"] if r["total"] else 0
        if pct < threshold:
            recent = db.execute("SELECT id FROM notifications WHERE user_id=? AND title='Low Attendance Alert' AND created_at>=? LIMIT 1",
                                (r["id"], now_ist().strftime('%Y-%m-01T00:00:00'))).fetchone()
            if not recent:
                notify_user(db, r["id"], "Low Attendance Alert", f"Your attendance is {pct:.1f}%, below the {threshold:.0f}% threshold.", "/student/dashboard")


def whatsapp_number(raw):
    """Return an international WhatsApp number (India default) or '' if invalid."""
    digits = re.sub(r"\D", "", raw or "")
    if len(digits) == 10:
        return "91" + digits
    if len(digits) == 11 and digits.startswith("0"):
        return "91" + digits[1:]
    if len(digits) == 12 and digits.startswith("91"):
        return digits
    return ""


def load_student_for_report(db, student_id):
    """Return (student, contact, rows) or None when the staff user may not access this student."""
    student = db.execute("SELECT * FROM students WHERE id=?", (student_id,)).fetchone()
    if not student:
        return None
    params = [student_id]
    teacher_join, teacher_where = "", ""
    if session.get("role") == "teacher":
        assigned = db.execute("""SELECT 1 FROM subject_teachers st JOIN subjects sub ON sub.id=st.subject_id
                                 WHERE st.teacher_id=? AND sub.year=? LIMIT 1""",
                              (session["user_id"], student["year"])).fetchone()
        if not assigned:
            return None
        teacher_join = " JOIN subject_teachers st ON st.subject_id=sub.id AND st.teacher_id=?"
        params.append(session["user_id"])
        teacher_where = " AND sub.year=?"
        params.append(student["year"])
    rows = db.execute(f"""SELECT sub.id, sub.code, sub.name, COALESCE(sub.semester,1) semester,
            COUNT(a.id) total,
            COALESCE(SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END),0) present
        FROM subjects sub
        LEFT JOIN attendance a ON a.subject_id=sub.id AND a.student_id=?{teacher_join}
        WHERE 1=1{teacher_where}
        GROUP BY sub.id, sub.code, sub.name, sub.semester
        HAVING COUNT(a.id)>0
        ORDER BY sub.code""", params).fetchall()
    history_by_subject = {}
    if rows:
        ids = [int(r["id"]) for r in rows]
        marks = ",".join("?" for _ in ids)
        hist = db.execute(f"""SELECT a.subject_id, a.attendance_date, COALESCE(a.lecture_no,1) lecture_no,
                COALESCE(a.lecture_time,'') lecture_time, COALESCE(a.session_type,'Lecture') session_type, a.status
            FROM attendance a WHERE a.student_id=? AND a.subject_id IN ({marks})
            ORDER BY a.attendance_date DESC, a.lecture_no""", [student_id] + ids).fetchall()
        for h in hist:
            history_by_subject.setdefault(int(h["subject_id"]), []).append(h)
    contact = db.execute("SELECT email, mobile FROM users WHERE student_id=? AND role='student' LIMIT 1",
                         (student_id,)).fetchone()
    report = []
    for r in rows:
        total, present = int(r["total"] or 0), int(r["present"] or 0)
        report.append({"id": int(r["id"]), "code": r["code"], "name": r["name"], "semester": int(r["semester"] or 1),
                       "total": total, "present": present, "absent": total - present,
                       "pct": (100.0 * present / total) if total else 0.0,
                       "history": history_by_subject.get(int(r["id"]), [])})
    return student, contact, report


def report_overall(report):
    total = sum(r["total"] for r in report)
    present = sum(r["present"] for r in report)
    return (100.0 * present / total) if total else 0.0


def report_text(student, report, with_names=True):
    lines = [f"{COLLEGE_NAME}", f"Attendance Report - {student['name']} (Roll No {student['roll_no']})", ""]
    for r in report:
        label = f"{r['code']} {r['name']}" if with_names else f"{r['code']}"
        lines.append(f"{label}: {r['present']}/{r['total']} ({r['pct']:.1f}%)")
    lines += ["", f"Overall: {report_overall(report):.1f}%"]
    return "\n".join(lines)


def build_subject_wise_pdf(student, report, threshold=75.0):
    """Share / email / public-link PDF: the website design, every subject on its own page."""
    return build_statement_pdf(dict(student), report, college=COLLEGE_NAME, logo_path=BASE_DIR / "static" / LOGO_FILE,
                               threshold=threshold)


@app.route("/send-report/<int:student_id>/pdf")
@login_required
@staff_required
def send_report_pdf(student_id):
    db = get_db()
    data = load_student_for_report(db, student_id)
    threshold = attendance_threshold_value(db)
    db.close()
    if not data:
        return Response("Student not found or not allowed.", status=404, mimetype="text/plain")
    student, contact, report = data
    if not report:
        return Response("No attendance records for this student yet.", status=404, mimetype="text/plain")
    pdf = build_subject_wise_pdf(student, report, threshold)
    safe = re.sub(r"[^A-Za-z0-9_-]+", "_", str(student["prn"] or student["roll_no"] or student["id"])).strip("_") or "student"
    log_activity("Report PDF generated", f"Attendance PDF generated for student {student['roll_no']}")
    return Response(pdf, mimetype="application/pdf",
                    headers={"Content-Disposition": f"inline; filename=Attendance_Report_{safe}.pdf"})


def report_link_serializer():
    return URLSafeTimedSerializer(app.secret_key, salt="attendance-report-link")


REPORT_LINK_MAX_AGE = 3 * 24 * 3600  # links expire after 3 days


@app.route("/r/<token>")
@limiter.limit("30 per minute")
def public_report_pdf(token):
    """Signed, expiring link so a student can open their own PDF from WhatsApp without logging in."""
    try:
        student_id = int(report_link_serializer().loads(token, max_age=REPORT_LINK_MAX_AGE))
    except (BadSignature, SignatureExpired, ValueError, TypeError):
        return Response("This report link is invalid or has expired. Please ask your teacher to send it again.",
                        status=410, mimetype="text/plain")
    db = get_db()
    data = load_student_for_report(db, student_id)
    threshold = attendance_threshold_value(db)
    db.close()
    if not data or not data[2]:
        return Response("Report not available.", status=404, mimetype="text/plain")
    student, contact, report = data
    pdf = build_subject_wise_pdf(student, report, threshold)
    safe = re.sub(r"[^A-Za-z0-9_-]+", "_", str(student["prn"] or student["roll_no"] or student["id"])).strip("_") or "student"
    return Response(pdf, mimetype="application/pdf",
                    headers={"Content-Disposition": f"inline; filename=Attendance_Report_{safe}.pdf",
                             "X-Robots-Tag": "noindex, nofollow", "Cache-Control": "private, no-store"})


@app.route("/send-report/<int:student_id>")
@login_required
@staff_required
def send_report(student_id):
    from urllib.parse import quote
    db = get_db()
    data = load_student_for_report(db, student_id)
    last = db.execute("SELECT MAX(sent_at) m FROM email_log WHERE student_id=? AND kind='weekly'", (student_id,)).fetchone()
    db.close()
    if not data:
        flash("Student not found or you are not assigned to this student's year.", "error")
        return redirect(url_for("students"))
    student, contact, report = data
    mobile = (contact["mobile"] if contact else "") or ""
    email = ((contact["email"] if contact else "") or "").strip()
    num = whatsapp_number(mobile)
    wa_link = ""
    if num and report:
        token = report_link_serializer().dumps(int(student_id))
        link = url_for("public_report_pdf", token=token, _external=True)
        msg = (f"Dear {student['name']}, your attendance report (Roll No {student['roll_no']}). "
               f"Overall: {report_overall(report):.1f}%.\n"
               f"Download your subject-wise PDF report (link valid for 3 days):\n{link}")
        wa_link = f"https://wa.me/{num}?text={quote(msg)}"
    return render_template("send_report.html", student=student, report=report, overall=report_overall(report),
                           mobile=mobile, wa_link=wa_link, email=email,
                           email_ok=bool(email and _EMAIL_RE.match(email)), email_ready=email_is_configured(),
                           emailed=email_sent_label(last["m"] if last else None),
                           pdf_url=url_for("send_report_pdf", student_id=student_id),
                           short_text=(f"Attendance report of {student['name']} (Roll No {student['roll_no']}). "
                                       f"Overall: {report_overall(report):.1f}%."))


@app.route("/low-attendance")
@login_required
@staff_required
def low_attendance():
    from urllib.parse import quote
    db = get_db()
    row = db.execute("SELECT value FROM app_settings WHERE key='attendance_threshold'").fetchone()
    try:
        threshold = float(row["value"]) if row and row["value"] else 75.0
    except Exception:
        threshold = 75.0
    year = request.args.get("year", "").strip()
    params, join_extra, where = [], "", ["s.blocked=0"]
    if session.get("role") == "teacher":
        join_extra = " AND a.subject_id IN (SELECT subject_id FROM subject_teachers WHERE teacher_id=?)"
        params.append(session["user_id"])
    if year:
        where.append("s.year=?")
        params.append(year)
    rows = db.execute(f"""SELECT s.id, s.name, s.roll_no, s.year, s.division, u.mobile, u.email,
            COUNT(a.id) total,
            COALESCE(SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END),0) present
        FROM students s
        LEFT JOIN attendance a ON a.student_id=s.id{join_extra}
        LEFT JOIN users u ON u.student_id=s.id AND u.role='student'
        WHERE {' AND '.join(where)}
        GROUP BY s.id, s.name, s.roll_no, s.year, s.division, u.mobile, u.email
        HAVING COUNT(a.id)>0""", params).fetchall()
    emailed = {int(r["student_id"]): r["m"] for r in db.execute("SELECT student_id, MAX(sent_at) m FROM email_log WHERE kind='low' GROUP BY student_id").fetchall()}
    years = [r["year"] for r in db.execute("SELECT DISTINCT year FROM students ORDER BY year").fetchall()]
    db.close()
    students = []
    for r in rows:
        pct = 100.0 * r["present"] / r["total"] if r["total"] else 0.0
        if pct >= threshold:
            continue
        num = whatsapp_number(r["mobile"])
        msg = (f"Dear Parent/Student, attendance of {r['name']} (Roll No {r['roll_no']}) is "
               f"{pct:.1f}%, which is below the required {threshold:.0f}%. "
               f"Please attend lectures regularly. - Attendance Office")
        students.append({"name": r["name"], "roll_no": r["roll_no"], "year": r["year"],
                         "division": r["division"], "present": r["present"], "total": r["total"],
                         "pct": pct, "mobile": r["mobile"] or "", "email": (r["email"] or "").strip(),
                         "id": r["id"], "can_email": bool(r["email"] and _EMAIL_RE.match(r["email"].strip())),
                         "emailed": email_sent_label(emailed.get(int(r["id"]))),
                         "wa_link": f"https://wa.me/{num}?text={quote(msg)}" if num else ""})
    students.sort(key=lambda x: x["pct"])
    return render_template("low_attendance.html", students=students, threshold=threshold,
                           years=years, year=year, email_ready=email_is_configured())


@app.route("/notifications")
@login_required
def notifications():
    db=get_db()
    rows=db.execute("SELECT * FROM notifications WHERE user_id=? ORDER BY created_at DESC LIMIT 100", (session["user_id"],)).fetchall()
    db.execute("UPDATE notifications SET is_read=1 WHERE user_id=?", (session["user_id"],)); db.commit(); db.close()
    return render_template("notifications.html", notifications=rows)


@app.route("/notifications/read/<int:notification_id>")
@login_required
def read_notification(notification_id):
    db=get_db(); row=db.execute("SELECT link FROM notifications WHERE id=? AND user_id=?", (notification_id,session["user_id"])).fetchone()
    db.execute("UPDATE notifications SET is_read=1 WHERE id=? AND user_id=?", (notification_id,session["user_id"])); db.commit(); db.close()
    return redirect(row["link"] if row and row["link"] else url_for("notifications"))


@app.route("/leaves", methods=["GET","POST"])
@login_required
def leaves():
    db=get_db()
    if session.get("role")=="student":
        st=current_student(db)
        if request.method=="POST":
            fd=request.form.get("from_date",""); td=request.form.get("to_date",""); reason=request.form.get("reason","").strip()
            if not fd or not td or not reason or fd>td: flash("Enter valid leave dates and reason.","error")
            else:
                db.execute("INSERT INTO leave_applications(student_id,from_date,to_date,reason,status,created_at) VALUES(?,?,?,?,?,?)",(st["id"],fd,td,reason,"Pending",datetime.utcnow().isoformat()))
                staff=db.execute("SELECT id FROM users WHERE role IN ('admin','teacher') AND approved=1").fetchall()
                for x in staff: notify_user(db,x["id"],"New Leave Application",f"{st['name']} submitted a leave application.","/leaves")
                db.commit(); flash("Leave application submitted.","success")
        rows=db.execute("SELECT * FROM leave_applications WHERE student_id=? ORDER BY created_at DESC",(st["id"],)).fetchall(); db.close()
        return render_template("leaves.html", student_mode=True, leaves=rows)
    rows=db.execute("SELECT l.*,s.name,s.prn,s.year FROM leave_applications l JOIN students s ON s.id=l.student_id ORDER BY l.created_at DESC").fetchall()
    db.close(); return render_template("leaves.html", student_mode=False, leaves=rows)


@app.post("/leaves/<int:leave_id>/review")
@login_required
def review_leave(leave_id):
    if session.get("role") != 'admin':
        flash("Only Admin can review leave applications.","error"); return redirect(url_for("leaves"))
    status=request.form.get("status"); note=request.form.get("review_note","").strip()
    if status not in ('Approved','Rejected'): return redirect(url_for("leaves"))
    db=get_db(); row=db.execute("SELECT l.*,u.id user_id,s.name FROM leave_applications l JOIN students s ON s.id=l.student_id JOIN users u ON u.student_id=s.id WHERE l.id=?",(leave_id,)).fetchone()
    if row:
        db.execute("UPDATE leave_applications SET status=?,reviewed_by=?,review_note=? WHERE id=?",(status,session["user_id"],note,leave_id))
        notify_user(db,row["user_id"],"Leave Application Updated",f"Your leave application has been {status.lower()}.","/leaves")
        db.commit(); flash(f"Leave {status.lower()}.","success")
    db.close(); return redirect(url_for("leaves"))


@app.route('/qr-attendance', methods=['GET','POST'])
@login_required
@staff_required
def qr_attendance():
    db = get_db()
    try:
        subjects_all = teacher_or_admin_subjects(db)
        years = YEAR_OPTIONS[:]
        selected_year = (request.args.get('year') or request.form.get('year') or '1st Year').strip()
        if selected_year not in YEAR_OPTIONS:
            selected_year = '1st Year'
        session_type = (request.args.get('session_type') or request.form.get('session_type') or 'Lecture').strip()
        if session_type not in ('Lecture', 'Practical'):
            session_type = 'Lecture'
        # Academic year maps to its two valid semesters (B.Pharm 4-year structure).
        valid_semesters = SEMESTERS_BY_YEAR.get(selected_year, [1, 2])
        semester_raw = request.args.get('semester') or request.form.get('semester')
        try:
            selected_semester = int(semester_raw) if semester_raw else valid_semesters[0]
        except (TypeError, ValueError):
            selected_semester = valid_semesters[0]
        if selected_semester not in valid_semesters:
            selected_semester = valid_semesters[0]

        batch = (request.args.get('batch') or request.form.get('batch') or '').strip()
        if session_type != 'Practical':
            batch = ''
        elif batch not in ('Batch A', 'Batch B', 'Batch C', 'Batch D'):
            batch = ''

        subject_id = request.args.get('subject_id', type=int) or request.form.get('subject_id', type=int)
        ad = request.args.get('attendance_date') or request.form.get('attendance_date') or today_ist().isoformat()
        try:
            date.fromisoformat(ad)
        except ValueError:
            ad = today_ist().isoformat()
        try:
            lecture_no = int(request.args.get('lecture_no') or request.form.get('lecture_no') or 1)
        except ValueError:
            lecture_no = 1
        if lecture_no not in range(1,6):
            lecture_no = 1
        start_time = (request.args.get('start_time') or request.form.get('start_time') or '').strip()
        end_time = (request.args.get('end_time') or request.form.get('end_time') or '').strip()
        teacher_status = (request.args.get('teacher_status') or request.form.get('teacher_status') or 'Scheduled teacher present').strip()
        allowed_status = (
            'Scheduled teacher present',
            'Scheduled teacher absent - substitute teacher',
            'Another teacher conducted the session'
        )
        if teacher_status not in allowed_status:
            teacher_status = 'Scheduled teacher present'
        conducted_by = request.args.get('conducted_by', type=int) or request.form.get('conducted_by', type=int)

        # Show only subjects matching the chosen year, semester and session type.
        subjects = [x for x in subjects_all
                    if x['year'] == selected_year
                    and int(x['semester'] or 1) == selected_semester
                    and (x['session_type'] or 'Lecture').strip().lower() == session_type.lower()]
        if subject_id and not any(int(x['id']) == int(subject_id) for x in subjects):
            subject_id = None
        teachers = all_teacher_rows(db)

        # Find a timetable teacher when a matching slot exists. Timetable is optional
        # for QR creation so an unassigned timetable teacher cannot block QR generation.
        scheduled = None
        timetable_id = request.args.get('timetable_id', type=int) or request.form.get('timetable_id', type=int)
        if subject_id and start_time and end_time:
            day_name = date.fromisoformat(ad).strftime('%A')
            q = """SELECT t.*, COALESCE(u.full_name,u.username,'') teacher_name
                   FROM timetable t LEFT JOIN users u ON u.id=t.teacher_id
                   WHERE t.year=? AND t.subject_id=? AND t.day_of_week=?
                     AND COALESCE(t.session_type,'Lecture')=?
                     AND (COALESCE(t.batch,'')=? OR ?='Lecture')
                     AND t.start_time=? AND t.end_time=?
                   ORDER BY CASE WHEN t.teacher_id IS NULL THEN 1 ELSE 0 END, t.id LIMIT 1"""
            scheduled = db.execute(q, (selected_year, subject_id, day_name, session_type,
                                       batch if session_type == 'Practical' else '', session_type,
                                       start_time, end_time)).fetchone()
        if timetable_id and not scheduled:
            try:
                scheduled = db.execute("SELECT t.*, COALESCE(u.full_name,u.username,'') teacher_name FROM timetable t LEFT JOIN users u ON u.id=t.teacher_id WHERE t.id=? AND t.year=? AND t.subject_id=?", (timetable_id, selected_year, subject_id)).fetchone()
            except Exception:
                scheduled = None

        if request.method == 'POST':
            valid_subject = any(int(x['id']) == int(subject_id or 0) for x in subjects)
            if not valid_subject:
                flash('Please select a subject belonging to the selected year.', 'error')
            elif not start_time or not end_time or start_time >= end_time:
                flash('Please select a valid start and end time.', 'error')
            elif session_type == 'Practical' and batch not in ('Batch A', 'Batch B', 'Batch C', 'Batch D'):
                flash('Select a valid practical batch.', 'error')
            elif teacher_status != 'Scheduled teacher present' and not conducted_by:
                flash('Please select the actual / conducted-by teacher.', 'error')
            else:
                scheduled_teacher_id = scheduled['teacher_id'] if scheduled and scheduled['teacher_id'] else None
                if teacher_status == 'Scheduled teacher present':
                    if scheduled_teacher_id:
                        conducted_by = scheduled_teacher_id
                    elif not conducted_by:
                        # If timetable has no assigned teacher, the current staff member
                        # becomes the conducting teacher instead of blocking QR creation.
                        conducted_by = session['user_id']
                elif not conducted_by:
                    conducted_by = session['user_id']

                token = secrets.token_urlsafe(24)
                exp = datetime.utcnow().replace(microsecond=0) + timedelta(minutes=10)
                db.execute("""INSERT INTO qr_sessions(
                    token,subject_id,year,attendance_date,lecture_no,lecture_time,
                    start_time,end_time,session_type,batch,teacher_id,
                    scheduled_teacher_id,conducted_by_teacher_id,teacher_status,expires_at,active)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,1)""",
                    (token, subject_id, selected_year, ad, lecture_no, f'{start_time}-{end_time}',
                     start_time, end_time, session_type, batch, session['user_id'],
                     scheduled_teacher_id, conducted_by, teacher_status, exp.isoformat()))
                db.commit()
                db.close()
                return redirect(url_for('qr_attendance', token=token))

        token = request.args.get('token', '')
        qr = None
        qr_session = None
        if token:
            qr_session = db.execute("""SELECT q.*,s.code,s.name subject_name,s.semester subject_semester,
                COALESCE(st.full_name,st.username,'') scheduled_teacher_name,
                COALESCE(ct.full_name,ct.username,'') conducted_teacher_name
                FROM qr_sessions q JOIN subjects s ON s.id=q.subject_id
                LEFT JOIN users st ON st.id=q.scheduled_teacher_id
                LEFT JOIN users ct ON ct.id=q.conducted_by_teacher_id
                WHERE q.token=? AND q.active=1""", (token,)).fetchone()
            if qr_session:
                try:
                    if datetime.fromisoformat(qr_session['expires_at']) < datetime.utcnow():
                        qr_session = None
                except Exception:
                    qr_session = None
            if qr_session and QRCODE_AVAILABLE:
                import base64
                img = qrcode.make(qr_scan_full_url(token))
                bio = io.BytesIO(); img.save(bio, format='PNG')
                qr = 'data:image/png;base64,' + base64.b64encode(bio.getvalue()).decode()

        db.close()
        return render_template('qr_attendance.html', subjects=subjects, years=years,
                               attendance_date=ad, qr=qr, qr_session=qr_session,
                               token=token, batch=batch, session_type=session_type,
                               teachers=teachers, selected_year=selected_year,
                               semesters=valid_semesters, selected_semester=selected_semester,
                               selected_subject=subject_id, start_time=start_time,
                               end_time=end_time, teacher_status=teacher_status,
                               conducted_by=conducted_by, scheduled_teacher=scheduled, lecture_no=lecture_no)
    except Exception:
        try: db.rollback(); db.close()
        except Exception: pass
        app.logger.exception('QR attendance page error')
        flash('QR Attendance could not be loaded. Please check the selected details.', 'error')
        return redirect(url_for('qr_attendance'))


QR_WINDOW_SECONDS = 20  # the QR on the teacher's screen changes every 20 seconds (a scan is valid for 20-40 s)


def qr_rotating_code(token, window=None):
    if window is None:
        window = int(time.time() // QR_WINDOW_SECONDS)
    return hmac.new(str(app.secret_key).encode("utf-8"), f"qr:{token}:{window}".encode("utf-8"), hashlib.sha256).hexdigest()[:12]


def qr_code_is_valid(token, code):
    """A scanned QR is valid for the current 30-second window and the one before it (30-60 seconds in total)."""
    code = (code or "").strip()
    if not code:
        return False
    now = int(time.time() // QR_WINDOW_SECONDS)
    ok = False
    for window in (now, now - 1):
        if hmac.compare_digest(qr_rotating_code(token, window), code):
            ok = True
    return ok


def qr_scan_full_url(token):
    return request.url_root.rstrip('/') + url_for('qr_scan', token=token, k=qr_rotating_code(token))


@app.route('/qr-attendance/image')
@login_required
@staff_required
def qr_image():
    """Fresh QR picture for the teacher's screen. The page asks for a new one every 20 seconds."""
    token = request.args.get('token', '')
    db = get_db()
    row = db.execute("SELECT expires_at FROM qr_sessions WHERE token=? AND active=1", (token,)).fetchone()
    db.close()
    try:
        alive = bool(row) and datetime.fromisoformat(row['expires_at']) >= datetime.utcnow()
    except Exception:
        alive = False
    if not alive or not QRCODE_AVAILABLE:
        return Response("QR session is closed or expired.", status=404, mimetype="text/plain")
    img = qrcode.make(qr_scan_full_url(token))
    bio = io.BytesIO(); img.save(bio, format='PNG')
    return Response(bio.getvalue(), mimetype='image/png', headers={'Cache-Control': 'no-store'})


@app.route('/qr/scan',methods=['GET','POST'])
@limiter.limit("20 per minute", methods=["POST"])
@login_required
def qr_scan():
    if session.get('role')!='student':
        flash('Only students can scan attendance QR codes.','error')
        return redirect(url_for('dashboard'))

    token=request.form.get('token', '') if request.method == 'POST' else request.args.get('token', '')
    auto_mark=request.method == 'POST' and request.form.get('auto') == '1'
    db=get_db()
    row=db.execute("""SELECT q.*,s.code,s.name subject_name FROM qr_sessions q
        JOIN subjects s ON s.id=q.subject_id
        WHERE q.token=? AND q.active=1""",(token,)).fetchone()
    if not row or datetime.fromisoformat(row['expires_at'])<datetime.utcnow():
        db.close()
        return render_template('qr_scan.html',valid=False)

    st=current_student(db)
    if not st:
        db.close()
        return render_template('qr_scan.html',valid=False,reason='Student account could not be found.')
    if st['year']!=row['year'] or (row['session_type']=='Practical' and st['batch']!=row['batch']):
        db.close()
        return render_template('qr_scan.html',valid=False,reason='This QR is not for your class or practical batch.')

    scan_code = request.form.get('k', '') if request.method == 'POST' else request.args.get('k', '')
    if not qr_code_is_valid(token, scan_code):
        db.close()
        return render_template('qr_scan.html', valid=False, reason='This QR code has expired. Please scan the QR that is currently shown on your teacher\'s screen.')
    already = db.execute("SELECT status FROM attendance WHERE student_id=? AND subject_id=? AND attendance_date=? AND lecture_no=?",
                         (st['id'], row['subject_id'], row['attendance_date'], row['lecture_no'])).fetchone()
    if already:
        db.close()
        if already['status'] == 'Present':
            return render_template('qr_scan.html', valid=True, done=True, session_info=row)
        return render_template('qr_scan.html', valid=False, reason='Your attendance for this lecture is already recorded as Absent. Please contact your teacher.')

    # The in-app Student QR Scanner uses auto=1, so a valid scan immediately
    # marks the student Present. The existing normal /qr/scan flow remains
    # available without auto=1.
    if request.method=='POST':
        db.execute("""INSERT INTO attendance(student_id,subject_id,attendance_date,lecture_no,lecture_time,start_time,end_time,session_type,batch,status,marked_by,academic_year,scheduled_teacher_id,conducted_by_teacher_id,teacher_status)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            ON CONFLICT(student_id,subject_id,attendance_date,lecture_no) DO NOTHING""",
                   (st['id'],row['subject_id'],row['attendance_date'],row['lecture_no'],row['lecture_time'],row['start_time'],row['end_time'],row['session_type'],row['batch'],'Present',row['conducted_by_teacher_id'] or row['teacher_id'],current_academic_year_value(db),row['scheduled_teacher_id'],row['conducted_by_teacher_id'] or row['teacher_id'],row['teacher_status']))
        notify_user(db, session['user_id'], "Attendance Marked", f"You were marked Present for {row['subject_name']} (Lecture {row['lecture_no']}).", "/student/dashboard")
        create_low_attendance_notifications(db)
        db.commit()
        db.close()
        return render_template('qr_scan.html',valid=True,done=True,session_info=row)

    db.close()
    return render_template('qr_scan.html',valid=True,session_info=row)


def finalize_qr_session(db, token, teacher_user_id=None):
    """Close a QR session and create Absent records for eligible students who
    did not already have an attendance record for that exact lecture.
    Existing Present/Absent records are never overwritten.
    """
    q = db.execute("SELECT * FROM qr_sessions WHERE token=?", (token,)).fetchone()
    if not q:
        return 0, None
    if teacher_user_id is not None and q['teacher_id'] != teacher_user_id:
        return 0, q

    db.execute("UPDATE qr_sessions SET active=0 WHERE token=?", (token,))

    if q['session_type'] == 'Practical':
        students = db.execute("""SELECT s.id student_id, u.id user_id
            FROM students s JOIN users u ON u.student_id=s.id
            WHERE u.role='student' AND u.approved=1 AND s.year=? AND COALESCE(s.batch,'')=?""",
            (q['year'], q['batch'] or '')).fetchall()
    else:
        students = db.execute("""SELECT s.id student_id, u.id user_id
            FROM students s JOIN users u ON u.student_id=s.id
            WHERE u.role='student' AND u.approved=1 AND s.year=?""",
            (q['year'],)).fetchall()

    absent_count = 0
    for st in students:
        cur = db.execute("""INSERT INTO attendance(
                student_id,subject_id,attendance_date,lecture_no,lecture_time,
                start_time,end_time,session_type,batch,status,marked_by,academic_year,
                scheduled_teacher_id,conducted_by_teacher_id,teacher_status)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            ON CONFLICT(student_id,subject_id,attendance_date,lecture_no) DO NOTHING""",
            (st['student_id'], q['subject_id'], q['attendance_date'], q['lecture_no'],
             q['lecture_time'], q['start_time'], q['end_time'], q['session_type'],
             q['batch'], 'Absent', q['conducted_by_teacher_id'] or q['teacher_id'],
             current_academic_year_value(db), q['scheduled_teacher_id'],
             q['conducted_by_teacher_id'] or q['teacher_id'], q['teacher_status']))
        try:
            absent_count += int(cur.rowcount or 0)
        except Exception:
            pass
    return absent_count, q


@app.post('/qr/close/<token>')
@login_required
@staff_required
def close_qr(token):
    db=get_db()
    try:
        absent_count, q = finalize_qr_session(db, token, session['user_id'])
        if not q or q['teacher_id'] != session['user_id']:
            db.rollback(); db.close(); flash('QR attendance session not found.', 'error')
            return redirect(url_for('qr_attendance'))
        db.commit()
        db.close()
        flash(f'QR attendance session closed. {absent_count} unmarked student(s) were marked Absent.', 'success')
        return redirect(url_for('qr_attendance'))
    except Exception:
        try: db.rollback(); db.close()
        except Exception: pass
        app.logger.exception('QR attendance finalization error')
        flash('QR attendance could not be finalized.', 'error')
        return redirect(url_for('qr_attendance'))


@app.get('/manifest.json')
def manifest():
    return {"name":COLLEGE_NAME,"short_name":"Attendance","start_url":"/","display":"standalone","background_color":"#ffffff","theme_color":"#1f4e79","icons":[{"src":url_for('static',filename=LOGO_FILE),"sizes":"192x192","type":"image/png"},{"src":url_for('static',filename=LOGO_FILE),"sizes":"512x512","type":"image/png"}]}


@app.get('/service-worker.js')
def service_worker():
    js="""self.addEventListener('install',e=>self.skipWaiting());self.addEventListener('activate',e=>self.clients.claim());self.addEventListener('notificationclick',e=>{e.notification.close();e.waitUntil(clients.openWindow('/notifications'))});"""
    return Response(js,mimetype='application/javascript')

# Initialize the database when the application is imported by Gunicorn/Render.
init_db()

if __name__ == "__main__":
    app.run(debug=os.environ.get("FLASK_DEBUG", "0") == "1",
            host=os.environ.get("HOST", "0.0.0.0"),
            port=int(os.environ.get("PORT", "5000")))
