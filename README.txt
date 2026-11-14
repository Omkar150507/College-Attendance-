Y.N.P. College of Pharmacy - Attendance Management System
==========================================================

Standard Flask project: app.py, config.py, templates/, static/ (static/images/college_logo.png,
static/images/attendance-background.png). The folder name must be exactly "static" (no space).

ENVIRONMENT VARIABLES (set them in Render > Environment)
---------------------------------------------------------
DATABASE_URL            Neon connection string (postgresql://...?sslmode=require). Without it the app uses SQLite.
SECRET_KEY              Long random value (Render can generate it).
ADMIN_USERNAME          Admin login name (default "admin").
ADMIN_PASSWORD          Admin password, minimum 12 characters (used only when the admin is first created).
ADMIN_NAME, ADMIN_EMAIL Optional admin details.

Email (password reset + backups) - Brevo HTTPS API (SMTP is blocked on Render free):
BREVO_API_KEY           API key from Brevo > SMTP & API.
BREVO_SENDER_EMAIL      A sender email verified in Brevo > Senders.
BREVO_SENDER_NAME       Optional display name.

Rate limiting (optional but recommended):
RATELIMIT_STORAGE_URI   Free Upstash Redis URL, e.g. rediss://default:PASSWORD@HOST:6379

Automatic backup by email (optional):
CRON_SECRET             Random secret of 16+ characters.
BACKUP_EMAIL            Where the backup zip is emailed (defaults to ADMIN_EMAIL).
Schedule a weekly call on cron-job.org: POST (or GET) https://YOUR-SITE.onrender.com/cron/backup
with header  X-Cron-Key: <CRON_SECRET>   (or ?key=<CRON_SECRET>).
The same scheduler can also call /health every 5 minutes to keep the free service awake.

SECURITY FEATURES
-----------------
- Login type (Admin / Teacher / Student) is enforced on the server.
- Login time limit: 12 hours (14 days with "Remember Me").
- Forgot password: one-time email link valid for 30 minutes.
- Admin 2-step verification (authenticator app) at Admin > 2-Step Security, with recovery codes.
- QR attendance: the QR changes every 30 seconds; a scan never overwrites an existing attendance record.
- Only the administrator can clear leave applications.
- Deleting a student keeps a restorable copy: Admin > Deleted Students.
- Attendance dates use India time (Asia/Kolkata).

FEATURES
--------
Timetable, QR attendance, manual attendance, reports (PDF/Excel/CSV), leave applications, student
promotion, Low Attendance page with WhatsApp reminder button, Send Report (subject-wise PDF) per student.
