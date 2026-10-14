YNP College Attendance - Timetable & Teacher Status + Professional PDF Export Update v14

This package is an update for the existing Y.N.P. College of Pharmacy Attendance Management System.

Included:
- app.py: updated Flask backend + professional timetable PDF export
- timetable.html: timetable management/view page
- attendance.html: timetable-linked manual attendance UI
- qr_attendance.html: timetable-linked QR attendance UI
- README_TIMETABLE_V13.txt: installation notes

Features added:
1. Timetable per academic year and day.
2. Lecture sessions are common for the whole year.
3. Practical sessions are separate for Batch A, Batch B, Batch C and Batch D.
4. Start Time + End Time supports long practicals such as 09:30-12:30 and afternoon sessions.
5. Timetable stores Subject + Teacher + Room.
6. Student timetable automatically filters by the student's year; practicals also filter by the student's batch.
7. Teacher timetable automatically shows only that teacher's assigned sessions.
8. Take Attendance can select the scheduled timetable slot.
9. Teacher status can be Scheduled teacher present, Scheduled teacher absent/substitute, or Another teacher conducted the session.
10. Actual/conducted-by teacher is saved when a substitute/other teacher conducts the session.
11. QR attendance saves the same teacher information.
12. Attendance records store scheduled teacher, conducted-by teacher, teacher status, start time and end time.
13. PDF report header includes scheduled teacher/conducted-by teacher information when matching attendance exists.

Installation:
- Back up the current app.py and database first.
- Replace app.py with the included app.py.
- Copy the three included templates into your existing templates/ folder, replacing attendance.html and qr_attendance.html and adding timetable.html.
- Keep all existing static files and other templates unchanged.
- In base.html, add a sidebar/navigation link to {{ url_for('timetable') }} with label “Timetable” if it is not already present.
- The Timetable page now has an “Export Timetable PDF” button. The PDF respects the selected year/day and the logged-in role.
- Open /timetable after login. Admin manages entries; teachers/students see their relevant timetable.
- Add a navigation link in base.html to /timetable if the existing navigation does not already have one.

Important:
- This is an update package, not a replacement of your database. The app automatically adds the new timetable/teacher-status columns when it starts.
- Do not delete the existing database.
- Practical batch values are exactly: Batch A, Batch B, Batch C, Batch D.


Project structure note
----------------------
The project uses the standard Flask `templates/` and `static/` directories. Static assets are under `static/images/`, including `college_logo.png` and `login-campus-bg.jpg`.


SECURITY CONFIGURATION (UPDATED — NO REDIS REQUIRED)
--------------------------------
- Set SECRET_KEY to a long, randomly generated secret in production (Render already supports generated values).
- Before first database initialization, set ADMIN_PASSWORD to a strong password of at least 12 characters. The application no longer creates an admin with a default password.
- Public student registration now requires the student's PRN/name/year/semester/batch to match a student record entered by an administrator.
- POST forms require CSRF tokens. Keep Flask-WTF installed from requirements.txt.
- Back up the database before deploying security updates to an existing production database. Existing user accounts and attendance records are not intentionally changed by these code updates.


Option B deployment note (no external Redis)
--------------------------------------------
If RATELIMIT_STORAGE_URI is not configured, the app uses in-memory rate limiting.
This avoids a separate Redis service, but counters are per worker/instance and reset
on restart. It is not equivalent to shared rate limiting. Configure shared Redis later
if stronger cross-worker protection is required. Back up the database before deployment.
Test registration, pending approval, admin approve/reject, and login before relying on it.
