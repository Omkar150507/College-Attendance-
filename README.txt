Y.N.P. College of Pharmacy - College Attendance Management System

UPGRADED FEATURES
- College name and logo branding
- Admin login
- Teacher registration, admin approval and subject assignment
- Teacher can take attendance only for assigned subjects
- Admin can add/edit/delete students and subjects
- Student registration using PRN, name, username and password
- Student login
- Student sees ONLY their own attendance
- Student dashboard shows today's attendance for every subject
- Student dashboard shows subject-wise total/present/absent/percentage
- Teacher/admin attendance reports and CSV export
- SQLite database

IMPORTANT LOGIN
Admin: admin / admin123
Teachers: register from the login page; admin must approve and assign subjects.
Students: register from the login page using PRN, name, username and password.

HOW TO RUN ON WINDOWS
1. Open Command Prompt.
2. Go to this folder, for example:
   cd /d "C:\Users\sy050\OneDrive\Desktop\Attendance\CollegeAttendance"
3. Create a virtual environment:
   python -m venv venv
4. Activate it:
   venv\Scripts\activate
5. Install requirements:
   pip install -r requirements.txt
6. Start the website:
   python app.py
7. Open Chrome and visit:
   http://127.0.0.1:5000

COLLEGE NAME / LOGO
- College name: config.py -> COLLEGE_NAME
- Tagline: config.py -> COLLEGE_TAGLINE
- Logo path: config.py -> LOGO_FILE
- Current logo file: static/images/college_logo.png

STUDENT REGISTRATION CODE
- Backend: app.py -> /student/register route
- Page: templates/student_register.html
- Student dashboard: templates/student_dashboard.html
- Student navigation/login behavior: templates/base.html and templates/login.html

DATABASE
The file attendance.db is created automatically in the project folder when the app first runs.
The upgraded database adds a PRN field to students and links student login accounts to their student record.

SECURITY NOTE
Before real college deployment, change app.secret_key, use HTTPS, and use a production database/server setup.
