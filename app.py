"""Excel-connected financial dashboard for small businesses.

The workbook in data/financial_records.xlsx is the single source of truth.
Changes made directly in Excel appear in the dashboard on its next refresh.
"""

from __future__ import annotations

import io
import os
import tempfile
import uuid
from datetime import date, datetime, timedelta
from pathlib import Path
from functools import wraps
import sqlite3
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from typing import Any

import pandas as pd
from flask import Flask, jsonify, render_template, request, send_file, redirect, url_for, session, flash, g, has_request_context
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DEFAULT_WORKBOOK_PATH = DATA_DIR / "financial_records.xlsx"
DB_PATH = DATA_DIR / "accounts.sqlite3"
SHEET_NAME = "Transactions"

COLUMNS = [
    "Transaction ID",
    "Date",
    "Transaction Type",
    "Category",
    "Description",
    "Amount",
    "Party",
    "Due Date",
    "Payment Status",
]

REQUIRED_COLUMNS = ["Date", "Transaction Type", "Category", "Description", "Amount"]
HEADER_ALIASES = {
    "transaction id": "Transaction ID",
    "id": "Transaction ID",
    "date": "Date",
    "transaction date": "Date",
    "transaction type": "Transaction Type",
    "type": "Transaction Type",
    "category": "Category",
    "description": "Description",
    "details": "Description",
    "amount": "Amount",
    "party": "Party",
    "customer": "Party",
    "supplier": "Party",
    "vendor": "Party",
    "due date": "Due Date",
    "payment status": "Payment Status",
    "status": "Payment Status",
}


class WorkbookError(Exception):
    """Raised when the shared workbook cannot be read or safely updated."""


app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-only-change-this-secret-before-deploying")
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024


def get_workbook_path() -> Path:
    """Give each client a private Excel workbook; retain the legacy default outside a client session."""
    user_id = session.get("user_id") if has_request_context() else None
    role = session.get("role") if has_request_context() else None
    if user_id and role == "client":
        client_dir = DATA_DIR / "clients"
        client_dir.mkdir(parents=True, exist_ok=True)
        return client_dir / f"client_{int(user_id)}_financial_records.xlsx"
    return DEFAULT_WORKBOOK_PATH


def get_db():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    return connection


def init_db():
    with get_db() as db:
        db.execute("""CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT NOT NULL UNIQUE COLLATE NOCASE,
            password_hash TEXT NOT NULL,
            company_name TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'client',
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        )""")
        admin_email = os.environ.get("ADMIN_EMAIL", "").strip().lower()
        admin_password = os.environ.get("ADMIN_PASSWORD", "")
        if admin_email and admin_password:
            existing = db.execute("SELECT id FROM users WHERE email = ?", (admin_email,)).fetchone()
            if not existing:
                db.execute("INSERT INTO users(email,password_hash,company_name,role) VALUES(?,?,?,?)",
                           (admin_email, generate_password_hash(admin_password), "Ledgerly Administration", "admin"))


# Initialise the account database when run directly or imported by a WSGI server.
init_db()


def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get("user_id"):
            if request.path.startswith("/api/"):
                return jsonify({"error": "Please log in to continue."}), 401
            return redirect(url_for("login"))
        return view(*args, **kwargs)
    return wrapped


def client_required(view):
    @wraps(view)
    @login_required
    def wrapped(*args, **kwargs):
        if session.get("role") != "client":
            if request.path.startswith("/api/"):
                return jsonify({"error": "This feature is available to client accounts only."}), 403
            return redirect(url_for("admin_dashboard"))
        return view(*args, **kwargs)
    return wrapped


def admin_required(view):
    @wraps(view)
    @login_required
    def wrapped(*args, **kwargs):
        if session.get("role") != "admin":
            if request.path.startswith("/api/"):
                return jsonify({"error": "Administrator access required."}), 403
            return redirect(url_for("index"))
        return view(*args, **kwargs)
    return wrapped



def empty_transactions() -> pd.DataFrame:
    return pd.DataFrame(columns=COLUMNS)


def create_workbook_if_missing() -> None:
    """Create a styled, empty template only on the first run."""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    if get_workbook_path().exists():
        return

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = SHEET_NAME
    sheet.append(COLUMNS)
    header_fill = PatternFill("solid", fgColor="0F766E")
    for cell in sheet[1]:
        cell.font = Font(color="FFFFFF", bold=True)
        cell.fill = header_fill
    widths = [18, 14, 18, 20, 36, 14, 22, 14, 18]
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[chr(64 + index)].width = width
    sheet.freeze_panes = "A2"
    workbook.save(get_workbook_path())


def clean_header(value: Any) -> str:
    return " ".join(str(value).strip().lower().replace("_", " ").split())


def standardise_headers(frame: pd.DataFrame) -> pd.DataFrame:
    renamed = {}
    for column in frame.columns:
        mapped = HEADER_ALIASES.get(clean_header(column))
        if mapped:
            renamed[column] = mapped
    return frame.rename(columns=renamed)


def normalise_transactions(frame: pd.DataFrame, generate_ids: bool = False) -> pd.DataFrame:
    """Validate the Excel data and convert it to the app's canonical columns."""
    frame = standardise_headers(frame.copy())
    frame = frame.dropna(how="all")

    for column in COLUMNS:
        if column not in frame.columns:
            frame[column] = ""

    frame = frame[COLUMNS].copy()
    if frame.empty:
        return empty_transactions()

    frame["Date"] = pd.to_datetime(frame["Date"], errors="coerce")
    frame["Due Date"] = pd.to_datetime(frame["Due Date"], errors="coerce")
    frame["Amount"] = pd.to_numeric(
        frame["Amount"].astype(str).str.replace("₹", "", regex=False).str.replace(",", "", regex=False),
        errors="coerce",
    )
    frame["Transaction Type"] = frame["Transaction Type"].astype(str).str.strip().str.title()
    frame["Category"] = frame["Category"].fillna("").astype(str).str.strip()
    frame["Description"] = frame["Description"].fillna("").astype(str).str.strip()
    frame["Party"] = frame["Party"].fillna("").astype(str).str.strip()
    frame["Payment Status"] = frame["Payment Status"].fillna("Paid").astype(str).str.strip().str.title()
    frame["Transaction ID"] = frame["Transaction ID"].fillna("").astype(str).str.strip()

    # A workbook using a Category named Income is accepted for easier migration.
    missing_type = frame["Transaction Type"].isin(["", "Nan", "None"])
    frame.loc[missing_type & (frame["Category"].str.lower() == "income"), "Transaction Type"] = "Income"

    errors: list[str] = []
    for index, row in frame.iterrows():
        row_number = index + 2
        if pd.isna(row["Date"]):
            errors.append(f"row {row_number}: Date is invalid")
        if row["Transaction Type"] not in {"Income", "Expense"}:
            errors.append(f"row {row_number}: Transaction Type must be Income or Expense")
        if not row["Category"]:
            errors.append(f"row {row_number}: Category is required")
        if not row["Description"]:
            errors.append(f"row {row_number}: Description is required")
        if pd.isna(row["Amount"]) or row["Amount"] <= 0:
            errors.append(f"row {row_number}: Amount must be greater than zero")

    if errors:
        details = "; ".join(errors[:8])
        remainder = "" if len(errors) <= 8 else f"; and {len(errors) - 8} more issue(s)"
        raise WorkbookError(f"The workbook needs correction: {details}{remainder}.")

    if generate_ids:
        blank_ids = frame["Transaction ID"].eq("")
        frame.loc[blank_ids, "Transaction ID"] = [f"TXN-{uuid.uuid4().hex[:8].upper()}" for _ in range(blank_ids.sum())]

    frame["Date"] = frame["Date"].dt.date
    frame["Due Date"] = frame["Due Date"].apply(lambda value: "" if pd.isna(value) else value.date())
    frame["Amount"] = frame["Amount"].round(2)
    return frame


def read_transactions() -> pd.DataFrame:
    create_workbook_if_missing()
    try:
        frame = pd.read_excel(get_workbook_path(), sheet_name=SHEET_NAME, engine="openpyxl")
    except PermissionError as error:
        raise WorkbookError("Close financial_records.xlsx in Excel, then refresh the dashboard.") from error
    except ValueError as error:
        raise WorkbookError(f"The workbook must contain a sheet named '{SHEET_NAME}'.") from error
    except Exception as error:
        raise WorkbookError("The financial workbook could not be read. Restore a valid .xlsx file and try again.") from error
    return normalise_transactions(frame, generate_ids=False)


def write_transactions(frame: pd.DataFrame) -> None:
    """Write through a temporary file so the workbook is never half-saved."""
    frame = normalise_transactions(frame, generate_ids=True)
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    file_descriptor, temporary_name = tempfile.mkstemp(suffix=".xlsx", dir=DATA_DIR)
    os.close(file_descriptor)
    temporary_path = Path(temporary_name)

    try:
        with pd.ExcelWriter(temporary_path, engine="openpyxl") as writer:
            frame.to_excel(writer, sheet_name=SHEET_NAME, index=False)
        workbook = load_workbook(temporary_path)
        sheet = workbook[SHEET_NAME]
        header_fill = PatternFill("solid", fgColor="0F766E")
        for cell in sheet[1]:
            cell.font = Font(color="FFFFFF", bold=True)
            cell.fill = header_fill
        sheet.freeze_panes = "A2"
        for row in sheet.iter_rows(min_row=2, min_col=2, max_col=2):
            row[0].number_format = "yyyy-mm-dd"
        for row in sheet.iter_rows(min_row=2, min_col=6, max_col=6):
            row[0].number_format = '₹#,##0.00'
        for row in sheet.iter_rows(min_row=2, min_col=8, max_col=8):
            row[0].number_format = "yyyy-mm-dd"
        workbook.save(temporary_path)
        os.replace(temporary_path, get_workbook_path())
    except PermissionError as error:
        raise WorkbookError("Close financial_records.xlsx in Excel before saving from the dashboard.") from error
    finally:
        if temporary_path.exists():
            temporary_path.unlink(missing_ok=True)


def frame_for_json(frame: pd.DataFrame) -> list[dict[str, Any]]:
    if frame.empty:
        return []
    result = frame.copy()
    result["Date"] = pd.to_datetime(result["Date"]).dt.strftime("%Y-%m-%d")
    result["Due Date"] = result["Due Date"].apply(
        lambda value: "" if value in ("", None) or pd.isna(value) else pd.Timestamp(value).strftime("%Y-%m-%d")
    )
    result["Amount"] = result["Amount"].round(2)
    return result.where(pd.notnull(result), "").to_dict(orient="records")


def filter_transactions(frame: pd.DataFrame, start: str | None, end: str | None, transaction_type: str | None, category: str | None) -> pd.DataFrame:
    if frame.empty:
        return frame
    result = frame.copy()
    result["Date"] = pd.to_datetime(result["Date"])
    if start:
        result = result[result["Date"] >= pd.Timestamp(start)]
    if end:
        result = result[result["Date"] <= pd.Timestamp(end)]
    if transaction_type and transaction_type in {"Income", "Expense"}:
        result = result[result["Transaction Type"] == transaction_type]
    if category:
        result = result[result["Category"] == category]
    return result


def money(value: float) -> float:
    return round(float(value), 2)


def build_analysis(frame: pd.DataFrame, all_transactions: pd.DataFrame) -> dict[str, Any]:
    income = frame.loc[frame["Transaction Type"] == "Income", "Amount"].sum()
    expenses = frame.loc[frame["Transaction Type"] == "Expense", "Amount"].sum()
    profit = income - expenses
    margin = (profit / income * 100) if income else 0

    outstanding_mask = (
        (frame["Transaction Type"] == "Income")
        & frame["Payment Status"].isin(["Pending", "Overdue"])
    )
    outstanding = frame.loc[outstanding_mask, "Amount"].sum()
    today = pd.Timestamp(date.today())
    due_dates = pd.to_datetime(frame["Due Date"], errors="coerce")
    overdue_mask = outstanding_mask & due_dates.notna() & (due_dates < today)
    overdue = frame.loc[overdue_mask, "Amount"].sum()
    due_soon_mask = (
        outstanding_mask
        & due_dates.notna()
        & (due_dates >= today)
        & (due_dates <= today + timedelta(days=7))
    )
    due_soon = frame.loc[due_soon_mask, "Amount"].sum()

    payment_followups: list[dict[str, Any]] = []
    pending_payments = frame.loc[outstanding_mask].copy()
    if not pending_payments.empty:
        pending_payments["Due Date Sort"] = pd.to_datetime(pending_payments["Due Date"], errors="coerce")
        pending_payments["Priority"] = 2
        pending_payments.loc[pending_payments["Due Date Sort"].notna(), "Priority"] = 1
        pending_payments.loc[pending_payments["Due Date Sort"] < today, "Priority"] = 0
        pending_payments = pending_payments.sort_values(["Priority", "Due Date Sort", "Amount"], ascending=[True, True, False])
        for _, payment in pending_payments.iterrows():
            due_date = payment["Due Date Sort"]
            if pd.isna(due_date):
                due_message = "Due date not recorded"
                display_status = "Pending"
            elif due_date < today:
                days = (today - due_date).days
                due_message = f"{days} day{'s' if days != 1 else ''} overdue"
                display_status = "Overdue"
            elif due_date == today:
                due_message = "Due today"
                display_status = "Pending"
            else:
                days = (due_date - today).days
                due_message = f"Due in {days} day{'s' if days != 1 else ''}"
                display_status = "Pending"
            payment_followups.append({
                "party": payment["Party"] or "Customer not recorded",
                "description": payment["Description"],
                "amount": money(payment["Amount"]),
                "due_date": "" if pd.isna(due_date) else due_date.strftime("%Y-%m-%d"),
                "due_message": due_message,
                "status": display_status,
            })

    company_payment_mask = (
        (frame["Transaction Type"] == "Expense")
        & frame["Payment Status"].isin(["Pending", "Overdue"])
    )
    company_payable = frame.loc[company_payment_mask, "Amount"].sum()
    company_overdue_mask = company_payment_mask & due_dates.notna() & (due_dates < today)
    company_overdue = frame.loc[company_overdue_mask, "Amount"].sum()
    company_due_soon_mask = (
        company_payment_mask
        & due_dates.notna()
        & (due_dates >= today)
        & (due_dates <= today + timedelta(days=7))
    )
    company_due_soon = frame.loc[company_due_soon_mask, "Amount"].sum()

    company_payment_followups: list[dict[str, Any]] = []
    pending_company_payments = frame.loc[company_payment_mask].copy()
    if not pending_company_payments.empty:
        pending_company_payments["Due Date Sort"] = pd.to_datetime(pending_company_payments["Due Date"], errors="coerce")
        pending_company_payments["Priority"] = 2
        pending_company_payments.loc[pending_company_payments["Due Date Sort"].notna(), "Priority"] = 1
        pending_company_payments.loc[pending_company_payments["Due Date Sort"] < today, "Priority"] = 0
        pending_company_payments = pending_company_payments.sort_values(
            ["Priority", "Due Date Sort", "Amount"], ascending=[True, True, False]
        )
        for _, payment in pending_company_payments.iterrows():
            due_date = payment["Due Date Sort"]
            if pd.isna(due_date):
                due_message = "Due date not recorded"
                display_status = "Pending"
            elif due_date < today:
                days = (today - due_date).days
                due_message = f"{days} day{'s' if days != 1 else ''} overdue"
                display_status = "Overdue"
            elif due_date == today:
                due_message = "Due today"
                display_status = "Pending"
            else:
                days = (due_date - today).days
                due_message = f"Due in {days} day{'s' if days != 1 else ''}"
                display_status = "Pending"
            company_payment_followups.append({
                "party": payment["Party"] or "Supplier not recorded",
                "description": payment["Description"],
                "amount": money(payment["Amount"]),
                "due_date": "" if pd.isna(due_date) else due_date.strftime("%Y-%m-%d"),
                "due_message": due_message,
                "status": display_status,
            })

    monthly = frame.copy()
    if not monthly.empty:
        monthly["Month"] = pd.to_datetime(monthly["Date"]).dt.to_period("M").astype(str)
        monthly_data = (
            monthly.pivot_table(index="Month", columns="Transaction Type", values="Amount", aggfunc="sum", fill_value=0)
            .reset_index()
            .rename_axis(None, axis=1)
        )
        if "Income" not in monthly_data:
            monthly_data["Income"] = 0.0
        if "Expense" not in monthly_data:
            monthly_data["Expense"] = 0.0
        monthly_data["Profit"] = monthly_data["Income"] - monthly_data["Expense"]
        monthly_trend = [
            {"month": row["Month"], "income": money(row["Income"]), "expense": money(row["Expense"]), "profit": money(row["Profit"])}
            for _, row in monthly_data.iterrows()
        ]
    else:
        monthly_trend = []

    expense_categories = frame[frame["Transaction Type"] == "Expense"].groupby("Category")["Amount"].sum().sort_values(ascending=False)
    category_breakdown = [{"category": category, "amount": money(amount)} for category, amount in expense_categories.items()]
    top_revenue = frame[frame["Transaction Type"] == "Income"].groupby("Description")["Amount"].sum().sort_values(ascending=False).head(5)
    top_revenue_items = [{"description": description, "amount": money(amount)} for description, amount in top_revenue.items()]

    last_90_days = all_transactions.copy()
    if not last_90_days.empty:
        last_90_days["Date"] = pd.to_datetime(last_90_days["Date"])
        cutoff = max(last_90_days["Date"].max() - timedelta(days=89), last_90_days["Date"].min())
        recent = last_90_days[last_90_days["Date"] >= cutoff]
        recent_income = recent.loc[recent["Transaction Type"] == "Income", "Amount"].sum()
        recent_expense = recent.loc[recent["Transaction Type"] == "Expense", "Amount"].sum()
        days_observed = max((recent["Date"].max() - recent["Date"].min()).days + 1, 1) if not recent.empty else 1
        thirty_day_cash_flow = (recent_income - recent_expense) / days_observed * 30
    else:
        thirty_day_cash_flow = 0

    insights: list[dict[str, str]] = []
    if frame.empty:
        insights.append({"level": "info", "text": "Add a transaction or upload your Excel workbook to begin financial analysis."})
    else:
        if income == 0:
            insights.append({"level": "warning", "text": "No income is recorded for the selected period. Review sales entries before making decisions from the profit figure."})
        elif profit < 0:
            insights.append({"level": "warning", "text": f"Expenses are ₹{abs(profit):,.2f} higher than income in the selected period."})
        else:
            insights.append({"level": "success", "text": f"The selected period has a positive profit of ₹{profit:,.2f} and a {margin:.1f}% profit margin."})
        if not expense_categories.empty:
            top_category = expense_categories.index[0]
            top_amount = expense_categories.iloc[0]
            share = top_amount / expenses * 100 if expenses else 0
            insights.append({"level": "info", "text": f"{top_category} is the largest expense category at ₹{top_amount:,.2f} ({share:.1f}% of expenses)."})
        if overdue > 0:
            insights.append({"level": "warning", "text": f"₹{overdue:,.2f} is overdue from customers. Follow up on these payments to protect cash flow."})
        if due_soon > 0:
            insights.append({"level": "warning", "text": f"₹{due_soon:,.2f} in customer payments is due within the next 7 days."})
        elif outstanding > 0:
            insights.append({"level": "info", "text": f"₹{outstanding:,.2f} is awaiting customer payment."})
        if company_overdue > 0:
            insights.append({"level": "warning", "text": f"₹{company_overdue:,.2f} in supplier bills is overdue and needs to be paid."})
        if company_due_soon > 0:
            insights.append({"level": "warning", "text": f"₹{company_due_soon:,.2f} in company payments is due within the next 7 days."})
        elif company_payable > 0:
            insights.append({"level": "info", "text": f"₹{company_payable:,.2f} in supplier bills is still pending payment."})
        if thirty_day_cash_flow < 0:
            insights.append({"level": "warning", "text": f"The recent transaction pattern indicates an estimated ₹{abs(thirty_day_cash_flow):,.2f} cash outflow over the next 30 days."})

    return {
        "kpis": {
            "income": money(income),
            "expenses": money(expenses),
            "profit": money(profit),
            "profit_margin": money(margin),
            "outstanding": money(outstanding),
            "overdue": money(overdue),
            "due_soon": money(due_soon),
            "pending_payment_count": int(outstanding_mask.sum()),
            "due_soon_count": int(due_soon_mask.sum()),
            "company_payable": money(company_payable),
            "company_overdue": money(company_overdue),
            "company_due_soon": money(company_due_soon),
            "company_payment_count": int(company_payment_mask.sum()),
            "forecast_30_days": money(thirty_day_cash_flow),
            "transaction_count": int(len(frame)),
        },
        "monthly_trend": monthly_trend,
        "expense_categories": category_breakdown,
        "top_revenue": top_revenue_items,
        "payment_followups": payment_followups,
        "company_payment_followups": company_payment_followups,
        "insights": insights,
    }


def workbook_modified() -> str:
    return datetime.fromtimestamp(get_workbook_path().stat().st_mtime).isoformat(timespec="seconds")


@app.errorhandler(413)
def request_too_large(_: Exception):
    return jsonify({"error": "The upload is too large. Use a file smaller than 10 MB."}), 413


@app.route("/")
def index():
    if session.get("role") == "admin":
        return redirect(url_for("admin_dashboard"))
    if session.get("role") == "client":
        create_workbook_if_missing()
        return render_template("index.html", company_name=session.get("company_name", "Your business"), email=session.get("email", ""))
    return render_template("home.html")


@app.route("/signup", methods=["GET", "POST"])
def signup():
    if request.method == "POST":
        email = request.form.get("email", "").strip().lower()
        password = request.form.get("password", "")
        company_name = request.form.get("company_name", "").strip()
        if not email or "@" not in email or not company_name or len(password) < 8:
            flash("Enter a valid email, company name, and password with at least 8 characters.", "error")
            return render_template("signup.html")
        try:
            with get_db() as db:
                cursor = db.execute("INSERT INTO users(email,password_hash,company_name,role) VALUES(?,?,?,'client')",
                                    (email, generate_password_hash(password), company_name))
                user_id = cursor.lastrowid
        except sqlite3.IntegrityError:
            flash("An account with that email already exists. Please log in.", "error")
            return render_template("signup.html")
        session.clear()
        session.update(user_id=user_id, email=email, company_name=company_name, role="client")
        create_workbook_if_missing()
        return redirect(url_for("index"))
    return render_template("signup.html")


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        email = request.form.get("email", "").strip().lower()
        password = request.form.get("password", "")
        with get_db() as db:
            user = db.execute("SELECT * FROM users WHERE email = ?", (email,)).fetchone()
        if not user or not check_password_hash(user["password_hash"], password):
            flash("Email or password is incorrect.", "error")
            return render_template("login.html")
        session.clear()
        session.update(user_id=user["id"], email=user["email"], company_name=user["company_name"], role=user["role"])
        return redirect(url_for("admin_dashboard" if user["role"] == "admin" else "index"))
    return render_template("login.html")


@app.route("/logout", methods=["POST"])
def logout():
    session.clear()
    return redirect(url_for("home"))


@app.route("/home")
def home():
    return render_template("home.html")


@app.route("/admin")
@admin_required
def admin_dashboard():
    with get_db() as db:
        users = db.execute("SELECT id,email,company_name,created_at FROM users WHERE role='client' ORDER BY created_at DESC").fetchall()
        total = db.execute("SELECT COUNT(*) AS n FROM users WHERE role='client'").fetchone()["n"]
    # Admin sees account/company metadata only; never read clients' workbooks here.
    return render_template("admin.html", users=users, total=total, email=session.get("email"))


@app.route("/api/dashboard")
@client_required
def dashboard():
    try:
        all_transactions = read_transactions()
        selected = filter_transactions(
            all_transactions,
            request.args.get("start"),
            request.args.get("end"),
            request.args.get("type"),
            request.args.get("category"),
        )
        analysis = build_analysis(selected, all_transactions)
        records = selected.sort_values("Date", ascending=False).head(250) if not selected.empty else selected
        return jsonify({
            "analysis": analysis,
            "transactions": frame_for_json(records),
            "categories": sorted(all_transactions["Category"].dropna().unique().tolist()),
            "modified_at": workbook_modified(),
            "workbook": get_workbook_path().name,
        })
    except WorkbookError as error:
        return jsonify({"error": str(error)}), 422


@app.route("/api/transactions", methods=["POST"])
@client_required
def add_transaction():
    payload = request.get_json(silent=True) or {}
    row = pd.DataFrame([{
        "Transaction ID": "", "Date": payload.get("date", ""),
        "Transaction Type": payload.get("transaction_type", ""), "Category": payload.get("category", ""),
        "Description": payload.get("description", ""), "Amount": payload.get("amount", ""),
        "Party": payload.get("party", ""), "Due Date": payload.get("due_date", ""),
        "Payment Status": payload.get("payment_status", "Paid"),
    }])
    try:
        new_row = normalise_transactions(row, generate_ids=True)
        existing = read_transactions()
        write_transactions(pd.concat([existing, new_row], ignore_index=True))
        return jsonify({"message": "Transaction saved to your Excel workbook.", "modified_at": workbook_modified()}), 201
    except WorkbookError as error:
        return jsonify({"error": str(error)}), 400


@app.route("/api/upload", methods=["POST"])
@client_required
def upload_workbook():
    uploaded = request.files.get("file")
    mode = request.form.get("mode", "append")
    if not uploaded or not uploaded.filename:
        return jsonify({"error": "Choose an Excel file first."}), 400
    if not uploaded.filename.lower().endswith(".xlsx"):
        return jsonify({"error": "Upload an Excel workbook in .xlsx format."}), 400
    try:
        source = pd.read_excel(io.BytesIO(uploaded.read()), engine="openpyxl")
        incoming = normalise_transactions(source, generate_ids=True)
        existing = read_transactions()
        combined = incoming if mode == "replace" else pd.concat([existing, incoming], ignore_index=True)
        write_transactions(combined.drop_duplicates(subset=["Transaction ID"], keep="last"))
        return jsonify({"message": f"{len(incoming)} transaction(s) imported into your workbook.", "modified_at": workbook_modified()})
    except WorkbookError as error:
        return jsonify({"error": str(error)}), 400
    except Exception:
        return jsonify({"error": "The upload could not be read. Download the template and use its column headings."}), 400


@app.route("/api/template")
@client_required
def download_template():
    create_workbook_if_missing()
    return send_file(get_workbook_path(), as_attachment=True, download_name="financial_records_template.xlsx")


@app.route("/api/workbook")
@client_required
def download_workbook():
    create_workbook_if_missing()
    return send_file(get_workbook_path(), as_attachment=True, download_name="financial_records.xlsx")


if __name__ == "__main__":
    init_db()
    app.run(debug=True, host="127.0.0.1", port=5000)
