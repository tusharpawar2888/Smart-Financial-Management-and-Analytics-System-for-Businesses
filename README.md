# Business Finance Desk

A simple Flask application with a public overview page, client signup/login, a separate admin login, and the original Excel-connected financial dashboard.

## Features
- Public home page with project overview and feature list
- Client signup using company name, email, and password
- Separate login route for clients and admins; role-based redirect after login
- Admin account overview listing company name, email, and signup date only
- Admin dashboard does not query or display client financial data
- Each client gets a separate Excel workbook under `data/clients/`
- Existing financial dashboard, transaction entry, Excel import/export, filters, payment follow-ups and charts are retained
- SQLite stores account metadata and password hashes; passwords are never stored as plain text

## Run locally (Windows PowerShell)

```powershell
py -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
$env:SECRET_KEY = "replace-with-a-long-random-secret"
$env:ADMIN_EMAIL = "admin@example.com"
$env:ADMIN_PASSWORD = "replace-with-a-strong-admin-password"
python app.py
```

Open `http://127.0.0.1:5000`. The public overview is `/`, signup is `/signup`, and login is `/login`. Admin users are created at startup from `ADMIN_EMAIL` and `ADMIN_PASSWORD`; use those credentials at the same login page. Do not use the sample values for a public deployment.

If you change `ADMIN_EMAIL` after an admin account already exists, it will create a new admin account; it does not reset an existing account's password.

## Excel workflow

Clients can enter transactions in the dashboard, download the template, upload a completed `.xlsx` workbook, or download their current workbook. The workbook is stored per client in `data/clients/`, so one client's dashboard does not load another client's file. Importing a changed workbook is how edits made offline in Excel are brought back into the app.

## Security notes

- Configure a long random `SECRET_KEY` and strong admin credentials.
- Do not commit `.env`, account databases, client workbooks, or real customer financial data.
- Admin pages show account metadata only. This is application-level separation, not encryption against a server/database administrator with filesystem access.
- For production, use HTTPS, secure session-cookie settings, CSRF protection, backups and a production WSGI server. SQLite and local workbook storage need persistent disk storage on the host.
