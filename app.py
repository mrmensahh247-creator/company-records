
from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    session,
    flash,
    send_file
)

import sqlite3
import csv
import io
import os
import shutil

from datetime import datetime, date, timedelta

from werkzeug.security import (
    generate_password_hash,
    check_password_hash
from werkzeug.utils import secure_filename
)

from openpyxl import Workbook

from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer
)
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet


# =========================================================
# APP
# =========================================================

app = Flask(__name__)

# Configure a strong SECRET_KEY in your environment before deployment.
# The fallback is for local development only and must not be used online.
app.config["SECRET_KEY"] = os.environ.get(
    "SECRET_KEY",
    "local-development-only-change-this-key"
)

app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_COOKIE_SECURE"] = (
    os.environ.get("COOKIE_SECURE", "0") == "1"
)

# Resolve the database relative to this file, not the current working directory.
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DATABASE = os.path.join(BASE_DIR, "company_records.db")

UPLOAD_FOLDER = os.path.join(BASE_DIR, "static", "uploads")
ALLOWED_PROFILE_EXTENSIONS = {"png", "jpg", "jpeg", "webp"}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

app.config["MAX_CONTENT_LENGTH"] = 2 * 1024 * 1024


# =========================================================
# DATABASE CONNECTION
# =========================================================

def get_db():

    conn = sqlite3.connect(DATABASE)

    conn.row_factory = sqlite3.Row

    return conn


# =========================================================
# INITIALIZE DATABASE
# =========================================================

def init_db():

    conn = get_db()

    # -----------------------------------------------------
    # RECORDS
    # -----------------------------------------------------

    conn.execute("""
        CREATE TABLE IF NOT EXISTS records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_id TEXT,
            staff_name TEXT NOT NULL,
            record_date TEXT NOT NULL,
            department TEXT NOT NULL,
            work_done TEXT NOT NULL,
            remarks TEXT,
            created_at TEXT NOT NULL
        )
    """)

    record_columns = [
        row["name"]
        for row in conn.execute(
            "PRAGMA table_info(records)"
        ).fetchall()
    ]

    if "employee_id" not in record_columns:

        conn.execute("""
            ALTER TABLE records
            ADD COLUMN employee_id TEXT
        """)

    # -----------------------------------------------------
    # USERS
    # -----------------------------------------------------
id INTEGER PRIMARY KEY AUTOINCREMENT,
username TEXT UNIQUE NOT NULL,
password TEXT NOT NULL,
role TEXT NOT NULL DEFAULT 'staff',
conn.execute("""     CREATE TABLE IF NOT EXISTS users (
    conn.execute("""
        CREATE TABLE IF NOT EXISTS users (
           
        )
    """)
        user_columns = [
        row["name"]
        for row in conn.execute(
            "PRAGMA table_info(users)"
        ).fetchall()
    ]

    if "profile_picture" not in user_columns:
        conn.execute(
            "ALTER TABLE users ADD COLUMN profile_picture TEXT"
        )

    # -----------------------------------------------------
    # COMPANY SETTINGS
    # -----------------------------------------------------

    conn.execute("""
        CREATE TABLE IF NOT EXISTS company_settings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_name TEXT NOT NULL,
            address TEXT,
            phone TEXT,
            email TEXT,
            website TEXT
        )
    """)

    # -----------------------------------------------------
    # EMPLOYEES
    # -----------------------------------------------------

    conn.execute("""
        CREATE TABLE IF NOT EXISTS employees (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_id TEXT UNIQUE NOT NULL,
            full_name TEXT NOT NULL,
            department TEXT NOT NULL,
            position TEXT,
            phone TEXT,
            email TEXT,
            status TEXT NOT NULL DEFAULT 'Active',
            created_at TEXT NOT NULL
        )
    """)

    # -----------------------------------------------------
    # ATTENDANCE
    # -----------------------------------------------------

    conn.execute("""
        CREATE TABLE IF NOT EXISTS attendance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_id TEXT NOT NULL,
            staff_name TEXT NOT NULL,
            attendance_date TEXT NOT NULL,
            check_in TEXT,
            check_out TEXT,
            status TEXT NOT NULL,
            remarks TEXT,
            created_at TEXT NOT NULL
        )
    """)

    # -----------------------------------------------------
    # DEFAULT ADMIN
    # -----------------------------------------------------

    admin = conn.execute("""
        SELECT id
        FROM users
        WHERE username = ?
    """, ("admin",)).fetchone()

    if not admin:

        initial_password = os.environ.get("INITIAL_ADMIN_PASSWORD", "")

        if len(initial_password) < 12:
            conn.close()
            raise RuntimeError(
                "A strong INITIAL_ADMIN_PASSWORD environment variable "
                "of at least 12 characters is required to create the "
                "initial admin account. Set it before starting a fresh "
                "installation. Existing databases with an admin account "
                "are not affected."
            )

        conn.execute("""
            INSERT INTO users (username, password, role)
            VALUES (?, ?, ?)
        """, (
            "admin",
            generate_password_hash(initial_password),
            "admin"
        ))

    # -----------------------------------------------------
    # DEFAULT COMPANY
    # -----------------------------------------------------

    company = conn.execute("""
        SELECT *
        FROM company_settings
        LIMIT 1
    """).fetchone()

    if not company:

        conn.execute("""
            INSERT INTO company_settings
            (
                company_name,
                address,
                phone,
                email,
                website
            )
            VALUES (?, ?, ?, ?, ?)
        """, (
            "My Company",
            "",
            "",
            "",
            ""
        ))

    conn.commit()

    conn.close()


# =========================================================
# HELPERS
# =========================================================

def login_required():

    return "user_id" in session


def admin_required():

    return session.get("role") == "admin"


def validate_password(password):

    if len(password) < 8:
        return False

    if not any(char.isalpha() for char in password):
        return False

    if not any(char.isdigit() for char in password):
        return False

    return True


def get_company_settings():

    conn = get_db()

    settings = conn.execute("""
        SELECT *
        FROM company_settings
        WHERE id = 1
    """).fetchone()

    conn.close()

    return settings


def get_active_employees():

    conn = get_db()

    employees = conn.execute("""
        SELECT *
        FROM employees
        WHERE status = 'Active'
        ORDER BY full_name ASC
    """).fetchall()

    conn.close()

    return employees


# =========================================================
# FILTERED RECORDS
# =========================================================

def get_filtered_records(
    start_date=None,
    end_date=None,
    staff_name=None,
    department=None,
    search=None
):

    conn = get_db()

    query = """
        SELECT *
        FROM records
        WHERE 1 = 1
    """

    params = []

    if start_date:

        query += """
            AND record_date >= ?
        """

        params.append(start_date)

    if end_date:

        query += """
            AND record_date <= ?
        """

        params.append(end_date)

    if staff_name:

        query += """
            AND staff_name = ?
        """

        params.append(staff_name)

    if department:

        query += """
            AND department = ?
        """

        params.append(department)

    if search:

        query += """
            AND (
                staff_name LIKE ?
                OR employee_id LIKE ?
                OR department LIKE ?
                OR work_done LIKE ?
                OR remarks LIKE ?
            )
        """

        search_value = "%" + search + "%"

        params.extend([
            search_value,
            search_value,
            search_value,
            search_value,
            search_value
        ])

    query += """
        ORDER BY record_date DESC, id DESC
    """

    records = conn.execute(
        query,
        params
    ).fetchall()

    conn.close()

    return records


# =========================================================
# LOGIN
# =========================================================

@app.route(
    "/login",
    methods=["GET", "POST"]
)
def login():

    if request.method == "POST":

        username = request.form.get(
            "username",
            ""
        ).strip()

        password = request.form.get(
            "password",
            ""
        )

        conn = get_db()

        user = conn.execute("""
            SELECT *
            FROM users
            WHERE username = ?
        """, (
            username,
        )).fetchone()

        conn.close()

        if user and check_password_hash(
            user["password"],
            password
        ):

            session.clear()

            session["user_id"] = user["id"]
            session["username"] = user["username"]
            session["role"] = user["role"]
session["profile_picture"] = user["profile_picture"]
            return redirect(
                url_for("dashboard")
            )

        flash(
            "Invalid username or password.",
            "error"
        )

    return render_template(
        "login.html"
    )


# =========================================================
# LOGOUT
# =========================================================

@app.route("/logout")
def logout():

    session.clear()

    return redirect(
        url_for("login")
    )


# =========================================================
# DASHBOARD
# =========================================================

@app.route(
    "/",
    methods=["GET", "POST"]
)
def dashboard():

    if not login_required():

        return redirect(
            url_for("login")
        )

    conn = get_db()

    # -----------------------------------------------------
    # ADD RECORD FROM DASHBOARD
    # -----------------------------------------------------

    if request.method == "POST":

        employee_id = request.form.get(
            "employee_id",
            ""
        ).strip()

        staff_name = request.form.get(
            "staff_name",
            ""
        ).strip()

        record_date = request.form.get(
            "record_date",
            ""
        ).strip()

        department = request.form.get(
            "department",
            ""
        ).strip()

        work_done = request.form.get(
            "work_done",
            ""
        ).strip()

        remarks = request.form.get(
            "remarks",
            ""
        ).strip()

        employee = None

        if employee_id:

            employee = conn.execute("""
                SELECT
                    employee_id,
                    full_name,
                    department
                FROM employees
                WHERE employee_id = ?
                AND status = 'Active'
                LIMIT 1
            """, (
                employee_id,
            )).fetchone()

        elif staff_name:

            employee = conn.execute("""
                SELECT
                    employee_id,
                    full_name,
                    department
                FROM employees
                WHERE full_name = ?
                AND status = 'Active'
                ORDER BY id ASC
                LIMIT 1
            """, (
                staff_name,
            )).fetchone()

        if employee:

            employee_id = employee["employee_id"]

            staff_name = employee["full_name"]

            department = employee["department"]

        else:

            employee_id = ""

        if (
            not staff_name
            or not record_date
            or not department
            or not work_done
        ):

            flash(
                "Please fill in all required fields.",
                "error"
            )

        elif not employee:

            flash(
                "Please select a valid employee.",
                "error"
            )

        else:

            conn.execute("""
                INSERT INTO records
                (
                    employee_id,
                    staff_name,
                    record_date,
                    department,
                    work_done,
                    remarks,
                    created_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (
                employee_id,
                staff_name,
                record_date,
                department,
                work_done,
                remarks,
                datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S"
                )
            ))

            conn.commit()

            flash(
                "Record added successfully.",
                "success"
            )

    # -----------------------------------------------------
    # DASHBOARD DATA
    # -----------------------------------------------------

    records = conn.execute("""
        SELECT *
        FROM records
        ORDER BY record_date DESC, id DESC
    """).fetchall()

    total_records = conn.execute(
        "SELECT COUNT(*) FROM records"
    ).fetchone()[0]

    today = date.today().isoformat()

    today_records = conn.execute("""
        SELECT COUNT(*)
        FROM records
        WHERE record_date = ?
    """, (
        today,
    )).fetchone()[0]

    departments = conn.execute("""
        SELECT COUNT(DISTINCT department)
        FROM records
    """).fetchone()[0]

    department_breakdown = conn.execute("""
        SELECT
            department,
            COUNT(*) AS total
        FROM records
        GROUP BY department
        ORDER BY total DESC
    """).fetchall()

    current_month = date.today().strftime(
        "%Y-%m"
    )

    this_month = conn.execute("""
        SELECT COUNT(*)
        FROM records
        WHERE record_date LIKE ?
    """, (
        current_month + "%",
    )).fetchone()[0]

    employees = conn.execute("""
        SELECT *
        FROM employees
        WHERE status = 'Active'
        ORDER BY full_name ASC
    """).fetchall()

    conn.close()

    return render_template(
        "index.html",
        records=records,
        total_records=total_records,
        today_records=today_records,
        departments=departments,
        department_breakdown=department_breakdown,
        this_month=this_month,
        employees=employees,
        today=today,
        company_settings=get_company_settings()
    )


# =========================================================
# ADD RECORD BUTTON
# =========================================================
# This route makes a dedicated "Add Record" button work.
# It simply takes the user to the dashboard form.

@app.route(
    "/add-record",
    methods=["GET", "POST"]
)
def add_record():

    if not login_required():

        return redirect(
            url_for("login")
        )

    if request.method == "POST":

        return dashboard()

    return redirect(
        url_for("dashboard")
    )


# =========================================================
# RECORDS BUTTON
# =========================================================
# This route makes a dedicated "Records" menu button work.
# Reports already provides the full searchable record list.

@app.route("/records")
def records_page():

    if not login_required():

        return redirect(
            url_for("login")
        )

    return redirect(
        url_for("reports")
    )


# =========================================================
# CHART DATA
# =========================================================

@app.route("/chart-data")
def chart_data():

    if not login_required():

        return redirect(
            url_for("login")
        )

    conn = get_db()

    department_rows = conn.execute("""
        SELECT
            department,
            COUNT(*) AS total
        FROM records
        GROUP BY department
        ORDER BY total DESC
    """).fetchall()

    last_7_days = []

    for i in range(6, -1, -1):

        day = date.today() - timedelta(
            days=i
        )

        total = conn.execute("""
            SELECT COUNT(*)
            FROM records
            WHERE record_date = ?
        """, (
            day.isoformat(),
        )).fetchone()[0]

        last_7_days.append({
            "date": day.strftime("%d %b"),
            "total": total
        })

    conn.close()

    return {
        "departments": [
            {
                "department": row["department"],
                "total": row["total"]
            }
            for row in department_rows
        ],
        "last_7_days": last_7_days
    }


# =========================================================
# EDIT RECORD
# =========================================================

@app.route(
    "/edit/<int:record_id>",
    methods=["GET", "POST"]
)
def edit_record(record_id):

    if not login_required():

        return redirect(
            url_for("login")
        )

    conn = get_db()

    record = conn.execute("""
        SELECT *
        FROM records
        WHERE id = ?
    """, (
        record_id,
    )).fetchone()

    if not record:

        conn.close()

        flash(
            "Record not found.",
            "error"
        )

        return redirect(
            url_for("dashboard")
        )

    if request.method == "POST":

        employee_id = request.form.get(
            "employee_id",
            ""
        ).strip()

        staff_name = request.form.get(
            "staff_name",
            ""
        ).strip()

        record_date = request.form.get(
            "record_date",
            ""
        ).strip()

        work_done = request.form.get(
            "work_done",
            ""
        ).strip()

        remarks = request.form.get(
            "remarks",
            ""
        ).strip()

        employee = None

        if employee_id:

            employee = conn.execute("""
                SELECT
                    employee_id,
                    full_name,
                    department
                FROM employees
                WHERE employee_id = ?
                AND status = 'Active'
                LIMIT 1
            """, (
                employee_id,
            )).fetchone()

        elif staff_name:

            employee = conn.execute("""
                SELECT
                    employee_id,
                    full_name,
                    department
                FROM employees
                WHERE full_name = ?
                AND status = 'Active'
                ORDER BY id ASC
                LIMIT 1
            """, (
                staff_name,
            )).fetchone()

        if employee:

            employee_id = employee["employee_id"]
            staff_name = employee["full_name"]
            department = employee["department"]

        else:

            employee_id = ""
            department = ""

        if (
            not staff_name
            or not record_date
            or not work_done
        ):

            flash(
                "Please fill in all required fields.",
                "error"
            )

        elif not employee:

            flash(
                "Please select a valid employee.",
                "error"
            )

        else:

            conn.execute("""
                UPDATE records
                SET
                    employee_id = ?,
                    staff_name = ?,
                    record_date = ?,
                    department = ?,
                    work_done = ?,
                    remarks = ?
                WHERE id = ?
            """, (
                employee_id,
                staff_name,
                record_date,
                department,
                work_done,
                remarks,
                record_id
            ))

            conn.commit()

            conn.close()

            flash(
                "Record updated successfully.",
                "success"
            )

            return redirect(
                url_for("dashboard")
            )

    employees = conn.execute("""
        SELECT *
        FROM employees
        WHERE status = 'Active'
        ORDER BY full_name
    """).fetchall()

    conn.close()

    return render_template(
        "edit.html",
        record=record,
        employees=employees,
        company_settings=get_company_settings()
    )


# =========================================================
# DELETE RECORD
# =========================================================

@app.route(
    "/delete/<int:record_id>",
    methods=["GET", "POST"]
)
def delete_record(record_id):

    if not login_required():

        return redirect(
            url_for("login")
        )

    if not admin_required():

        flash(
            "Only administrators can delete records.",
            "error"
        )

        return redirect(
            url_for("dashboard")
        )

    conn = get_db()

    record = conn.execute("""
        SELECT id
        FROM records
        WHERE id = ?
    """, (
        record_id,
    )).fetchone()

    if not record:

        conn.close()

        flash(
            "Record not found.",
            "error"
        )

        return redirect(
            url_for("dashboard")
        )

    conn.execute("""
        DELETE FROM records
        WHERE id = ?
    """, (
        record_id,
    ))

    conn.commit()

    conn.close()

    flash(
        "Record deleted successfully.",
        "success"
    )

    return redirect(
        url_for("dashboard")
    )


# =========================================================
# EMPLOYEES
# =========================================================

@app.route(
    "/employees",
    methods=["GET", "POST"]
)
def employees():

    if not login_required():

        return redirect(
            url_for("login")
        )

    if not admin_required():

        flash(
            "Only administrators can manage employees.",
            "error"
        )

        return redirect(
            url_for("dashboard")
        )

    conn = get_db()

    if request.method == "POST":

        employee_id = request.form.get(
            "employee_id",
            ""
        ).strip()

        full_name = request.form.get(
            "full_name",
            ""
        ).strip()

        department = request.form.get(
            "department",
            ""
        ).strip()

        position = request.form.get(
            "position",
            ""
        ).strip()

        phone = request.form.get(
            "phone",
            ""
        ).strip()

        email = request.form.get(
            "email",
            ""
        ).strip()

        status = request.form.get(
            "status",
            "Active"
        ).strip()

        if (
            not employee_id
            or not full_name
            or not department
        ):

            flash(
                "Employee ID, Full Name and Department are required.",
                "error"
            )

        else:

            try:

                conn.execute("""
                    INSERT INTO employees
                    (
                        employee_id,
                        full_name,
                        department,
                        position,
                        phone,
                        email,
                        status,
                        created_at
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    employee_id,
                    full_name,
                    department,
                    position,
                    phone,
                    email,
                    status,
                    datetime.now().strftime(
                        "%Y-%m-%d %H:%M:%S"
                    )
                ))

                conn.commit()

                flash(
                    "Employee added successfully.",
                    "success"
                )

            except sqlite3.IntegrityError:

                flash(
                    "Employee ID already exists.",
                    "error"
                )

    employee_list = conn.execute("""
        SELECT *
        FROM employees
        ORDER BY id DESC
    """).fetchall()

    conn.close()

    return render_template(
        "employees.html",
        employees=employee_list,
        company_settings=get_company_settings()
    )


# =========================================================
# EDIT EMPLOYEE
# =========================================================

@app.route(
    "/edit-employee/<int:employee_id>",
    methods=["GET", "POST"]
)
def edit_employee(employee_id):

    if not login_required():

        return redirect(
            url_for("login")
        )

    if not admin_required():

        flash(
            "Only administrators can edit employees.",
            "error"
        )

        return redirect(
            url_for("dashboard")
        )

    conn = get_db()

    employee = conn.execute("""
        SELECT *
        FROM employees
        WHERE id = ?
    """, (
        employee_id,
    )).fetchone()

    if not employee:

        conn.close()

        flash(
            "Employee not found.",
            "error"
        )

        return redirect(
            url_for("employees")
        )

    if request.method == "POST":

        employee_code = request.form.get(
            "employee_id",
            ""
        ).strip()

        full_name = request.form.get(
            "full_name",
            ""
        ).strip()

        department = request.form.get(
            "department",
            ""
        ).strip()

        position = request.form.get(
            "position",
            ""
        ).strip()

        phone = request.form.get(
            "phone",
            ""
        ).strip()

        email = request.form.get(
            "email",
            ""
        ).strip()

        status = request.form.get(
            "status",
            "Active"
        ).strip()

        if (
            not employee_code
            or not full_name
            or not department
        ):

            flash(
                "Employee ID, Full Name and Department are required.",
                "error"
            )

        else:

            try:

                conn.execute("""
                    UPDATE employees
                    SET
                        employee_id = ?,
                        full_name = ?,
                        department = ?,
                        position = ?,
                        phone = ?,
                        email = ?,
                        status = ?
                    WHERE id = ?
                """, (
                    employee_code,
                    full_name,
                    department,
                    position,
                    phone,
                    email,
                    status,
                    employee_id
                ))

                conn.commit()

                conn.close()

                flash(
                    "Employee updated successfully.",
                    "success"
                )

                return redirect(
                    url_for("employees")
                )

            except sqlite3.IntegrityError:

                flash(
                    "That Employee ID already exists.",
                    "error"
                )

    conn.close()

    return render_template(
        "edit_employee.html",
        employee=employee,
        company_settings=get_company_settings()
    )


# =========================================================
# DELETE EMPLOYEE
# =========================================================

@app.route(
    "/delete-employee/<int:employee_id>",
    methods=["GET", "POST"]
)
def delete_employee(employee_id):

    if not login_required():

        return redirect(
            url_for("login")
        )

    if not admin_required():

        flash(
            "Only administrators can delete employees.",
            "error"
        )

        return redirect(
            url_for("dashboard")
        )

    conn = get_db()

    employee = conn.execute("""
        SELECT *
        FROM employees
        WHERE id = ?
    """, (
        employee_id,
    )).fetchone()

    if not employee:

        conn.close()

        flash(
            "Employee not found.",
            "error"
        )

        return redirect(
            url_for("employees")
        )

    conn.execute("""
        DELETE FROM employees
        WHERE id = ?
    """, (
        employee_id,
    ))

    conn.commit()

    conn.close()

    flash(
        "Employee deleted successfully.",
        "success"
    )

    return redirect(
        url_for("employees")
    )


# =========================================================
# USERS
# =========================================================

@app.route(
    "/users",
    methods=["GET", "POST"]
)
def users():

    if not login_required():

        return redirect(
            url_for("login")
        )

    if not admin_required():

        flash(
            "Only administrators can manage users.",
            "error"
        )

        return redirect(
            url_for("dashboard")
        )

    conn = get_db()

    if request.method == "POST":

        username = request.form.get(
            "username",
            ""
        ).strip()

        password = request.form.get(
            "password",
            ""
        )

        role = request.form.get(
            "role",
            "staff"
        ).strip()

        if not username or not password:

            flash(
                "Username and password are required.",
                "error"
            )

        elif not validate_password(password):

            flash(
                "Password must be at least 8 characters and contain a letter and number.",
                "error"
            )

        else:

            try:

                conn.execute("""
                    INSERT INTO users
                    (
                        username,
                        password,
                        role
                    )
                    VALUES (?, ?, ?)
                """, (
                    username,
                    generate_password_hash(password),
                    role
                ))

                conn.commit()

                flash(
                    "User added successfully.",
                    "success"
                )

            except sqlite3.IntegrityError:

                flash(
                    "Username already exists.",
                    "error"
                )

    user_list = conn.execute("""
        SELECT
            id,
            username,
            role
        FROM users
        ORDER BY id DESC
    """).fetchall()

    conn.close()

    return render_template(
        "users.html",
        users=user_list,
        company_settings=get_company_settings()
    )


# =========================================================
# EDIT USER
# =========================================================

@app.route(
    "/edit-user/<int:user_id>",
    methods=["GET", "POST"]
)
def edit_user(user_id):

    if not login_required():

        return redirect(
            url_for("login")
        )

    if not admin_required():

        flash(
            "Only administrators can edit users.",
            "error"
        )

        return redirect(
            url_for("dashboard")
        )

    conn = get_db()

    user = conn.execute("""
        SELECT *
        FROM users
        WHERE id = ?
    """, (
        user_id,
    )).fetchone()

    if not user:

        conn.close()

        flash(
            "User not found.",
            "error"
        )

        return redirect(
            url_for("users")
        )

    if request.method == "POST":

        username = request.form.get(
            "username",
            ""
        ).strip()

        password = request.form.get(
            "password",
            ""
        )

        role = request.form.get(
            "role",
            "staff"
        ).strip()

        if not username:

            flash(
                "Username is required.",
                "error"
            )

        elif password and not validate_password(password):

            flash(
                "Password must be at least 8 characters and contain a letter and number.",
                "error"
            )

        else:

            try:

                if password:

                    conn.execute("""
                        UPDATE users
                        SET
                            username = ?,
                            password = ?,
                            role = ?
                        WHERE id = ?
                    """, (
                        username,
                        generate_password_hash(password),
                        role,
                        user_id
                    ))

                else:

                    conn.execute("""
                        UPDATE users
                        SET
                            username = ?,
                            role = ?
                        WHERE id = ?
                    """, (
                        username,
                        role,
                        user_id
                    ))

                conn.commit()

                conn.close()

                flash(
                    "User updated successfully.",
                    "success"
                )

                return redirect(
                    url_for("users")
                )

            except sqlite3.IntegrityError:

                flash(
                    "Username already exists.",
                    "error"
                )

    conn.close()

    return render_template(
        "edit_user.html",
        user=user,
        company_settings=get_company_settings()
    )


# =========================================================
# DELETE USER
# =========================================================

@app.route(
    "/delete-user/<int:user_id>",
    methods=["GET", "POST"]
)
def delete_user(user_id):

    if not login_required():

        return redirect(
            url_for("login")
        )

    if not admin_required():

        flash(
            "Only administrators can delete users.",
            "error"
        )

        return redirect(
            url_for("dashboard")
        )

    if user_id == session.get("user_id"):

        flash(
            "You cannot delete your own account.",
            "error"
        )

        return redirect(
            url_for("users")
        )

    conn = get_db()

    user = conn.execute("""
        SELECT id
        FROM users
        WHERE id = ?
    """, (
        user_id,
    )).fetchone()

    if not user:

        conn.close()

        flash(
            "User not found.",
            "error"
        )

        return redirect(
            url_for("users")
        )

    conn.execute("""
        DELETE FROM users
        WHERE id = ?
    """, (
        user_id,
    ))

    conn.commit()

    conn.close()

    flash(
        "User deleted successfully.",
        "success"
    )

    return redirect(
        url_for("users")
    )


# =========================================================
# CHANGE PASSWORD
# =========================================================

@app.route(
    "/change-password",
    methods=["GET", "POST"]
)
def change_password():

    if not login_required():

        return redirect(
            url_for("login")
        )

    conn = get_db()

    user = conn.execute("""
        SELECT *
        FROM users
        WHERE id = ?
    """, (
        session["user_id"],
    )).fetchone()

    company_settings = conn.execute("""
        SELECT *
        FROM company_settings
        WHERE id = 1
    """).fetchone()

    if not user:

        conn.close()

        session.clear()

        return redirect(
            url_for("login")
        )

    if request.method == "POST":

        current_password = request.form.get(
            "current_password",
            ""
        )

        new_password = request.form.get(
            "new_password",
            ""
        )

        confirm_password = request.form.get(
            "confirm_password",
            ""
        )

        if not check_password_hash(
            user["password"],
            current_password
        ):

            conn.close()

            flash(
                "Current password is incorrect.",
                "error"
            )

            return redirect(
                url_for("change_password")
            )

        if new_password != confirm_password:

            conn.close()

            flash(
                "New passwords do not match.",
                "error"
            )

            return redirect(
                url_for("change_password")
            )

        if not validate_password(new_password):

            conn.close()

            flash(
                "Password must be at least 8 characters and contain a letter and number.",
                "error"
            )

            return redirect(
                url_for("change_password")
            )

        conn.execute("""
            UPDATE users
            SET password = ?
            WHERE id = ?
        """, (
            generate_password_hash(new_password),
            session["user_id"]
        ))

        conn.commit()

        conn.close()

        flash(
            "Password changed successfully.",
            "success"
        )

        return redirect(
            url_for("dashboard")
        )

    conn.close()

    return render_template(
        "change_password.html",
        company_settings=company_settings
    )


# =========================================================
# COMPANY SETTINGS
# =========================================================

# =========================================================
# USER PROFILE AND PROFILE PICTURE
# =========================================================

@app.route("/profile", methods=["GET", "POST"])
def profile():

    if not login_required():
        return redirect(url_for("login"))

    conn = get_db()

    user = conn.execute("""
        SELECT id, username, role, profile_picture
        FROM users
        WHERE id = ?
    """, (session["user_id"],)).fetchone()

    if not user:
        conn.close()
        session.clear()
        return redirect(url_for("login"))

    if request.method == "POST":

        uploaded_file = request.files.get("profile_picture")

        if not uploaded_file or not uploaded_file.filename:
            conn.close()
            flash("Please choose a picture first.", "error")
            return redirect(url_for("profile"))

        original_name = secure_filename(uploaded_file.filename)
        extension = os.path.splitext(original_name)[1].lower().lstrip(".")

        if extension not in ALLOWED_PROFILE_EXTENSIONS:
            conn.close()
            flash("Use a PNG, JPG, JPEG, or WEBP image.", "error")
            return redirect(url_for("profile"))

        try:
            os.makedirs(UPLOAD_FOLDER, exist_ok=True)

            filename = (
                f"user_{session['user_id']}_"
                f"{int(datetime.now().timestamp())}.{extension}"
            )

            uploaded_file.save(
                os.path.join(UPLOAD_FOLDER, filename)
            )

            conn.execute("""
                UPDATE users
                SET profile_picture = ?
                WHERE id = ?
            """, (filename, session["user_id"]))

            conn.commit()

            session["profile_picture"] = filename

            flash("Profile picture updated successfully.", "success")

        except Exception:
            flash("The picture could not be saved. Please try again.", "error")

        finally:
            conn.close()

        return redirect(url_for("profile"))

    conn.close()

    return render_template(
        "profile.html",
        user=user,
        company_settings=get_company_settings()
    )

@app.route(
    "/settings",
    methods=["GET", "POST"]
)
def settings():

    if not login_required():

        return redirect(
            url_for("login")
        )

    if not admin_required():

        flash(
            "Only administrators can change company settings.",
            "error"
        )

        return redirect(
            url_for("dashboard")
        )

    conn = get_db()

    company_settings = conn.execute("""
        SELECT *
        FROM company_settings
        WHERE id = 1
    """).fetchone()

    if request.method == "POST":

        company_name = request.form.get(
            "company_name",
            ""
        ).strip()

        address = request.form.get(
            "address",
            request.form.get(
                "company_address",
                ""
            )
        ).strip()

        phone = request.form.get(
            "phone",
            request.form.get(
                "company_phone",
                ""
            )
        ).strip()

        email = request.form.get(
            "email",
            request.form.get(
                "company_email",
                ""
            )
        ).strip()

        website = request.form.get(
            "website",
            ""
        ).strip()

        if not company_name:

            flash(
                "Company name is required.",
                "error"
            )

        else:

            conn.execute("""
                UPDATE company_settings
                SET
                    company_name = ?,
                    address = ?,
                    phone = ?,
                    email = ?,
                    website = ?
                WHERE id = 1
            """, (
                company_name,
                address,
                phone,
                email,
                website
            ))

            conn.commit()

            company_settings = conn.execute("""
                SELECT *
                FROM company_settings
                WHERE id = 1
            """).fetchone()

            flash(
                "Company settings updated successfully.",
                "success"
            )

    conn.close()

    return render_template(
        "settings.html",
        company_settings=company_settings
    )


# =========================================================
# CSV EXPORT
# =========================================================

@app.route("/export-csv")
def export_csv():

    if not login_required():

        return redirect(
            url_for("login")
        )

    records = get_filtered_records()

    output = io.StringIO()

    writer = csv.writer(output)

    writer.writerow([
        "Employee ID",
        "Staff Name",
        "Date",
        "Department",
        "Work Done",
        "Remarks",
        "Created At"
    ])

    for record in records:

        writer.writerow([
            record["employee_id"] or "",
            record["staff_name"],
            record["record_date"],
            record["department"],
            record["work_done"],
            record["remarks"] or "",
            record["created_at"]
        ])

    output.seek(0)

    return send_file(
        io.BytesIO(
            output.getvalue().encode("utf-8")
        ),
        mimetype="text/csv",
        as_attachment=True,
        download_name="company_records.csv"
    )


# =========================================================
# EXCEL EXPORT
# =========================================================

@app.route("/export-excel")
def export_excel():

    if not login_required():

        return redirect(
            url_for("login")
        )

    records = get_filtered_records()

    workbook = Workbook()

    sheet = workbook.active

    sheet.title = "Company Records"

    sheet.append([
        "Employee ID",
        "Staff Name",
        "Date",
        "Department",
        "Work Done",
        "Remarks",
        "Created At"
    ])

    for record in records:

        sheet.append([
            record["employee_id"] or "",
            record["staff_name"],
            record["record_date"],
            record["department"],
            record["work_done"],
            record["remarks"] or "",
            record["created_at"]
        ])

    output = io.BytesIO()

    workbook.save(output)

    output.seek(0)

    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="company_records.xlsx"
    )


# =========================================================
# CREATE PDF
# =========================================================

def create_pdf(records):

    company = get_company_settings()

    buffer = io.BytesIO()

    document = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        rightMargin=25,
        leftMargin=25,
        topMargin=25,
        bottomMargin=25
    )

    styles = getSampleStyleSheet()

    elements = []

    elements.append(
        Paragraph(
            company["company_name"],
            styles["Title"]
        )
    )

    if company["address"]:

        elements.append(
            Paragraph(
                company["address"],
                styles["Normal"]
            )
        )

    contact = []

    if company["phone"]:
        contact.append(company["phone"])

    if company["email"]:
        contact.append(company["email"])

    if company["website"]:
        contact.append(company["website"])

    if contact:

        elements.append(
            Paragraph(
                " | ".join(contact),
                styles["Normal"]
            )
        )

    elements.append(
        Spacer(1, 15)
    )

    elements.append(
        Paragraph(
            "Company Records Report",
            styles["Heading2"]
        )
    )

    elements.append(
        Spacer(1, 10)
    )

    data = [[
        "Employee ID",
        "Staff Name",
        "Date",
        "Department",
        "Work Done",
        "Remarks",
        "Created At"
    ]]

    for record in records:

        data.append([
            record["employee_id"] or "",
            record["staff_name"],
            record["record_date"],
            record["department"],
            record["work_done"],
            record["remarks"] or "",
            record["created_at"]
        ])

    table = Table(
        data,
        repeatRows=1
    )

    table.setStyle(
        TableStyle([
            (
                "BACKGROUND",
                (0, 0),
                (-1, 0),
                colors.HexColor("#1f2937")
            ),
            (
                "TEXTCOLOR",
                (0, 0),
                (-1, 0),
                colors.white
            ),
            (
                "FONTNAME",
                (0, 0),
                (-1, 0),
                "Helvetica-Bold"
            ),
            (
                "GRID",
                (0, 0),
                (-1, -1),
                0.5,
                colors.grey
            ),
            (
                "VALIGN",
                (0, 0),
                (-1, -1),
                "TOP"
            ),
            (
                "FONTSIZE",
                (0, 0),
                (-1, -1),
                8
            ),
            (
                "ROWBACKGROUNDS",
                (0, 1),
                (-1, -1),
                [
                    colors.white,
                    colors.HexColor("#f3f4f6")
                ]
            )
        ])
    )

    elements.append(table)

    document.build(elements)

    buffer.seek(0)

    return buffer


# =========================================================
# PDF EXPORT
# =========================================================

@app.route("/export-pdf")
def export_pdf():

    if not login_required():

        return redirect(
            url_for("login")
        )

    records = get_filtered_records()

    pdf = create_pdf(records)

    return send_file(
        pdf,
        mimetype="application/pdf",
        as_attachment=True,
        download_name="company_records.pdf"
    )


# =========================================================
# REPORTS
# =========================================================

@app.route("/reports")
def reports():

    if not login_required():

        return redirect(
            url_for("login")
        )

    records = get_filtered_records()

    conn = get_db()

    staff_list = conn.execute("""
        SELECT DISTINCT staff_name
        FROM records
        WHERE staff_name != ''
        ORDER BY staff_name
    """).fetchall()

    department_list = conn.execute("""
        SELECT DISTINCT department
        FROM records
        WHERE department != ''
        ORDER BY department
    """).fetchall()

    conn.close()

    return render_template(
        "reports.html",
        records=records,
        staff_list=staff_list,
        department_list=department_list,
        staff_names=[
            row["staff_name"]
            for row in staff_list
        ],
        department_names=[
            row["department"]
            for row in department_list
        ],
        total_records=len(records),
        search="",
        start_date="",
        end_date="",
        selected_staff="",
        selected_department="",
        company_settings=get_company_settings()
    )


# =========================================================
# REPORT FILTER
# =========================================================

@app.route("/reports/filter")
def reports_filter():

    if not login_required():

        return redirect(
            url_for("login")
        )

    search = request.args.get(
        "search",
        ""
    ).strip()

    start_date = request.args.get(
        "start_date",
        ""
    ).strip()

    end_date = request.args.get(
        "end_date",
        ""
    ).strip()

    selected_staff = request.args.get(
        "staff_name",
        ""
    ).strip()

    selected_department = request.args.get(
        "department",
        ""
    ).strip()

    records = get_filtered_records(
        start_date=start_date or None,
        end_date=end_date or None,
        staff_name=selected_staff or None,
        department=selected_department or None,
        search=search or None
    )

    conn = get_db()

    staff_list = conn.execute("""
        SELECT DISTINCT staff_name
        FROM records
        WHERE staff_name != ''
        ORDER BY staff_name
    """).fetchall()

    department_list = conn.execute("""
        SELECT DISTINCT department
        FROM records
        WHERE department != ''
        ORDER BY department
    """).fetchall()

    conn.close()

    return render_template(
        "reports.html",
        records=records,
        staff_list=staff_list,
        department_list=department_list,
        staff_names=[
            row["staff_name"]
            for row in staff_list
        ],
        department_names=[
            row["department"]
            for row in department_list
        ],
        total_records=len(records),
        search=search,
        start_date=start_date,
        end_date=end_date,
        selected_staff=selected_staff,
        selected_department=selected_department,
        company_settings=get_company_settings()
    )


# =========================================================
# REPORT EXCEL
# =========================================================

@app.route("/reports/export-excel")
def reports_export_excel():

    if not login_required():

        return redirect(
            url_for("login")
        )

    search = request.args.get(
        "search",
        ""
    ).strip()

    start_date = request.args.get(
        "start_date",
        ""
    ).strip()

    end_date = request.args.get(
        "end_date",
        ""
    ).strip()

    staff_name = request.args.get(
        "staff_name",
        ""
    ).strip()

    department = request.args.get(
        "department",
        ""
    ).strip()

    records = get_filtered_records(
        start_date=start_date or None,
        end_date=end_date or None,
        staff_name=staff_name or None,
        department=department or None,
        search=search or None
    )

    workbook = Workbook()

    sheet = workbook.active

    sheet.title = "Filtered Report"

    sheet.append([
        "Employee ID",
        "Staff Name",
        "Date",
        "Department",
        "Work Done",
        "Remarks",
        "Created At"
    ])

    for record in records:

        sheet.append([
            record["employee_id"] or "",
            record["staff_name"],
            record["record_date"],
            record["department"],
            record["work_done"],
            record["remarks"] or "",
            record["created_at"]
        ])

    output = io.BytesIO()

    workbook.save(output)

    output.seek(0)

    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="filtered_company_report.xlsx"
    )


# =========================================================
# REPORT PDF
# =========================================================

@app.route("/reports/export-pdf")
def reports_export_pdf():

    if not login_required():

        return redirect(
            url_for("login")
        )

    search = request.args.get(
        "search",
        ""
    ).strip()

    start_date = request.args.get(
        "start_date",
        ""
    ).strip()

    end_date = request.args.get(
        "end_date",
        ""
    ).strip()

    staff_name = request.args.get(
        "staff_name",
        ""
    ).strip()

    department = request.args.get(
        "department",
        ""
    ).strip()

    records = get_filtered_records(
        start_date=start_date or None,
        end_date=end_date or None,
        staff_name=staff_name or None,
        department=department or None,
        search=search or None
    )

    pdf = create_pdf(records)

    return send_file(
        pdf,
        mimetype="application/pdf",
        as_attachment=True,
        download_name="filtered_company_report.pdf"
    )


# =========================================================
# ATTENDANCE
# =========================================================

@app.route(
    "/attendance",
    methods=["GET", "POST"]
)
def attendance():

    if not login_required():

        return redirect(
            url_for("login")
        )

    conn = get_db()

    if request.method == "POST":

        employee_id = request.form.get(
            "employee_id",
            ""
        ).strip()

        attendance_action = request.form.get(
            "attendance_action",
            ""
        ).strip()

        remarks = request.form.get(
            "remarks",
            ""
        ).strip()

        employee = conn.execute("""
            SELECT
                employee_id,
                full_name
            FROM employees
            WHERE employee_id = ?
            AND status = 'Active'
            LIMIT 1
        """, (
            employee_id,
        )).fetchone()

        if not employee:

            conn.close()

            flash(
                "Please select a valid active employee.",
                "error"
            )

            return redirect(
                url_for("attendance")
            )

        today = date.today().isoformat()

        current_time = datetime.now().strftime(
            "%H:%M"
        )

        existing = conn.execute("""
            SELECT *
            FROM attendance
            WHERE employee_id = ?
            AND attendance_date = ?
            ORDER BY id DESC
            LIMIT 1
        """, (
            employee_id,
            today
        )).fetchone()

        # -------------------------------------------------
        # CHECK IN
        # -------------------------------------------------

        if attendance_action == "check_in":

            if existing and existing["check_in"]:

                conn.close()

                flash(
                    f"{employee['full_name']} has already checked in today at {existing['check_in']}.",
                    "error"
                )

                return redirect(
                    url_for("attendance")
                )

            if existing:

                conn.execute("""
                    UPDATE attendance
                    SET
                        check_in = ?,
                        status = ?,
                        remarks = ?
                    WHERE id = ?
                """, (
                    current_time,
                    "Present",
                    remarks,
                    existing["id"]
                ))

            else:

                conn.execute("""
                    INSERT INTO attendance
                    (
                        employee_id,
                        staff_name,
                        attendance_date,
                        check_in,
                        check_out,
                        status,
                        remarks,
                        created_at
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    employee["employee_id"],
                    employee["full_name"],
                    today,
                    current_time,
                    None,
                    "Present",
                    remarks,
                    datetime.now().strftime(
                        "%Y-%m-%d %H:%M:%S"
                    )
                ))

            conn.commit()

            conn.close()

            flash(
                f"{employee['full_name']} checked in at {current_time}.",
                "success"
            )

            return redirect(
                url_for("attendance")
            )

        # -------------------------------------------------
        # CHECK OUT
        # -------------------------------------------------

        elif attendance_action == "check_out":

            if not existing:

                conn.close()

                flash(
                    f"{employee['full_name']} has not checked in today.",
                    "error"
                )

                return redirect(
                    url_for("attendance")
                )

            if not existing["check_in"]:

                conn.close()

                flash(
                    f"{employee['full_name']} must check in before checking out.",
                    "error"
                )

                return redirect(
                    url_for("attendance")
                )

            if existing["check_out"]:

                conn.close()

                flash(
                    f"{employee['full_name']} has already checked out today at {existing['check_out']}.",
                    "error"
                )

                return redirect(
                    url_for("attendance")
                )

            conn.execute("""
                UPDATE attendance
                SET
                    check_out = ?,
                    remarks = ?
                WHERE id = ?
            """, (
                current_time,
                remarks or existing["remarks"],
                existing["id"]
            ))

            conn.commit()

            conn.close()

            flash(
                f"{employee['full_name']} checked out at {current_time}.",
                "success"
            )

            return redirect(
                url_for("attendance")
            )

        # -------------------------------------------------
        # ABSENT
        # -------------------------------------------------

        elif attendance_action == "absent":

            if existing:

                conn.close()

                flash(
                    f"Attendance already exists for {employee['full_name']} today.",
                    "error"
                )

                return redirect(
                    url_for("attendance")
                )

            conn.execute("""
                INSERT INTO attendance
                (
                    employee_id,
                    staff_name,
                    attendance_date,
                    check_in,
                    check_out,
                    status,
                    remarks,
                    created_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                employee["employee_id"],
                employee["full_name"],
                today,
                None,
                None,
                "Absent",
                remarks,
                datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S"
                )
            ))

            conn.commit()

            conn.close()

            flash(
                f"{employee['full_name']} has been marked absent.",
                "success"
            )

            return redirect(
                url_for("attendance")
            )

        else:

            conn.close()

            flash(
                "Please choose Check In, Check Out or Absent.",
                "error"
            )

            return redirect(
                url_for("attendance")
            )

    employees = conn.execute("""
        SELECT
            employee_id,
            full_name
        FROM employees
        WHERE status = 'Active'
        ORDER BY full_name ASC
    """).fetchall()

    attendance_records = conn.execute("""
        SELECT *
        FROM attendance
        ORDER BY attendance_date DESC, id DESC
        LIMIT 100
    """).fetchall()

    conn.close()

    return render_template(
        "attendance.html",
        employees=employees,
        attendance_records=attendance_records,
        company_settings=get_company_settings(),
        today=date.today().isoformat()
    )


# =========================================================
# ATTENDANCE REPORTS
# =========================================================

@app.route("/attendance-reports")
def attendance_reports():

    if not login_required():

        return redirect(
            url_for("login")
        )

    conn = get_db()

    start_date = request.args.get(
        "start_date",
        ""
    ).strip()

    end_date = request.args.get(
        "end_date",
        ""
    ).strip()

    employee_id = request.args.get(
        "employee_id",
        ""
    ).strip()

    selected_department = request.args.get(
        "department",
        ""
    ).strip()

    selected_status = request.args.get(
        "status",
        ""
    ).strip()

    query = """
        SELECT
            attendance.*,
            employees.department
        FROM attendance
        LEFT JOIN employees
            ON attendance.employee_id = employees.employee_id
        WHERE 1 = 1
    """

    params = []

    if start_date:

        query += """
            AND attendance.attendance_date >= ?
        """

        params.append(start_date)

    if end_date:

        query += """
            AND attendance.attendance_date <= ?
        """

        params.append(end_date)

    if employee_id:

        query += """
            AND attendance.employee_id = ?
        """

        params.append(employee_id)

    if selected_department:

        query += """
            AND employees.department = ?
        """

        params.append(selected_department)

    if selected_status:

        query += """
            AND attendance.status = ?
        """

        params.append(selected_status)

    query += """
        ORDER BY
            attendance.attendance_date DESC,
            attendance.id DESC
    """

    attendance_records = conn.execute(
        query,
        params
    ).fetchall()

    employees = conn.execute("""
        SELECT
            employee_id,
            full_name
        FROM employees
        ORDER BY full_name ASC
    """).fetchall()

    departments = conn.execute("""
        SELECT DISTINCT department
        FROM employees
        WHERE department IS NOT NULL
        AND department != ''
        ORDER BY department ASC
    """).fetchall()

    total_attendance = len(
        attendance_records
    )

    present_count = 0
    absent_count = 0
    late_count = 0

    for record in attendance_records:

        status = str(
            record["status"]
        ).lower().strip()

        if status == "present":

            present_count += 1

        elif status == "absent":

            absent_count += 1

        elif status == "late":

            late_count += 1

    conn.close()

    return render_template(
        "attendance_reports.html",
        attendance_records=attendance_records,
        records=attendance_records,
        employees=employees,
        departments=departments,
        total_attendance=total_attendance,
        total=total_attendance,
        present_count=present_count,
        present=present_count,
        absent_count=absent_count,
        absent=absent_count,
        late_count=late_count,
        late=late_count,
        start_date=start_date,
        end_date=end_date,
        employee_id=employee_id,
        employee=employee_id,
        selected_department=selected_department,
        selected_status=selected_status,
        company_settings=get_company_settings()
    )


# =========================================================
# BACKUP
# =========================================================

@app.route("/backup")
def backup():

    if not login_required():

        return redirect(
            url_for("login")
        )

    if not admin_required():

        flash(
            "Only administrators can create backups.",
            "error"
        )

        return redirect(
            url_for("dashboard")
        )

    if not os.path.exists(DATABASE):

        flash(
            "Database file not found.",
            "error"
        )

        return redirect(
            url_for("dashboard")
        )

    filename = (
        "company_records_backup_"
        + datetime.now().strftime(
            "%Y%m%d_%H%M%S"
        )
        + ".db"
    )

    return send_file(
        DATABASE,
        as_attachment=True,
        download_name=filename
    )


# =========================================================
# RESTORE
# =========================================================

@app.route(
    "/restore",
    methods=["GET", "POST"]
)
def restore():

    if not login_required():

        return redirect(
            url_for("login")
        )

    if not admin_required():

        flash(
            "Only administrators can restore backups.",
            "error"
        )

        return redirect(
            url_for("dashboard")
        )

    if request.method == "POST":

        uploaded_file = request.files.get(
            "backup_file"
        )

        if (
            not uploaded_file
            or uploaded_file.filename == ""
        ):

            flash(
                "Please select a backup file.",
                "error"
            )

            return redirect(
                url_for("restore")
            )

        temp_file = "restore_temp.db"

        uploaded_file.save(temp_file)

        try:

            test_conn = sqlite3.connect(
                temp_file
            )

            integrity = test_conn.execute(
                "PRAGMA integrity_check"
            ).fetchone()[0]

            if integrity != "ok":

                test_conn.close()

                if os.path.exists(temp_file):
                    os.remove(temp_file)

                flash(
                    "The backup file is not a valid SQLite database.",
                    "error"
                )

                return redirect(
                    url_for("restore")
                )

            tables = test_conn.execute("""
                SELECT name
                FROM sqlite_master
                WHERE type = 'table'
            """).fetchall()

            table_names = {
                row[0]
                for row in tables
            }

            required_tables = {
                "records",
                "users",
                "company_settings"
            }

            if not required_tables.issubset(
                table_names
            ):

                test_conn.close()

                if os.path.exists(temp_file):
                    os.remove(temp_file)

                flash(
                    "The backup is missing required database tables.",
                    "error"
                )

                return redirect(
                    url_for("restore")
                )

            test_conn.close()

            if os.path.exists(DATABASE):

                safety_backup = (
                    "company_records_before_restore_"
                    + datetime.now().strftime(
                        "%Y%m%d_%H%M%S"
                    )
                    + ".db"
                )

                shutil.copy2(
                    DATABASE,
                    safety_backup
                )

            os.replace(
                temp_file,
                DATABASE
            )

            flash(
                "Database restored successfully.",
                "success"
            )

        except Exception as e:

            if os.path.exists(temp_file):

                os.remove(temp_file)

            flash(
                "Restore failed: " + str(e),
                "error"
            )

        return redirect(
            url_for("restore")
        )

    return render_template(
        "restore.html",
        company_settings=get_company_settings()
    )


# =========================================================
# ERROR HANDLERS
# =========================================================

@app.errorhandler(404)
def page_not_found(error):

    return render_template(
        "login.html"
    ) if not login_required() else redirect(
        url_for("dashboard")
    )


@app.errorhandler(500)
def internal_server_error(error):

    if login_required():

        flash(
            "An internal server error occurred. Please try again.",
            "error"
        )

        return redirect(
            url_for("dashboard")
        )

    return redirect(
        url_for("login")
    )


# =========================================================
# START DATABASE
# =========================================================

init_db()


# =========================================================
# START APPLICATION
# =========================================================

if __name__ == "__main__":

    app.run(
        host="0.0.0.0",
        port=5000,
        debug=False
    )
