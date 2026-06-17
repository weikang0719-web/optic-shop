from flask import send_file
from io import BytesIO
from openpyxl import Workbook, load_workbook
from reportlab.pdfgen import canvas
from flask import Flask, request, session, redirect
from datetime import datetime
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Spacer,
    Paragraph
)
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
import calendar
import sqlite3
import os
import psycopg2

app = Flask(__name__)
app.secret_key = "optic_shop_secret_key"

DATABASE_URL = os.environ.get("DATABASE_URL")

def get_conn():
    return psycopg2.connect(DATABASE_URL)

def init_db():
    conn = get_conn()
    c = conn.cursor()

    c.execute("""
        CREATE TABLE IF NOT EXISTS sales (
            id SERIAL PRIMARY KEY,
            customer TEXT,
            amount REAL,
            staff TEXT,
            date TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)

    c.execute("""
        ALTER TABLE sales
        ADD COLUMN IF NOT EXISTS company_code TEXT
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS expenses (
            id SERIAL PRIMARY KEY,
            category TEXT,
            amount REAL,
            note TEXT,
            date TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)

    c.execute("""
        ALTER TABLE expenses
        ADD COLUMN IF NOT EXISTS company_code TEXT
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS salaries (
            id SERIAL PRIMARY KEY,
            staff TEXT,
            amount REAL,
            month TEXT,
            date TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)

    c.execute("""
        ALTER TABLE salaries
        ADD COLUMN IF NOT EXISTS company_code TEXT
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id SERIAL PRIMARY KEY,
            username TEXT UNIQUE,
            password TEXT,
            role TEXT DEFAULT 'staff',

            can_add_sales BOOLEAN DEFAULT FALSE,
            can_edit_sales BOOLEAN DEFAULT FALSE,
            can_delete_sales BOOLEAN DEFAULT FALSE,

            can_add_expenses BOOLEAN DEFAULT FALSE,
            can_edit_expenses BOOLEAN DEFAULT FALSE,
            can_delete_expenses BOOLEAN DEFAULT FALSE,

            can_add_salary BOOLEAN DEFAULT FALSE,
            can_edit_salary BOOLEAN DEFAULT FALSE,
            can_delete_salary BOOLEAN DEFAULT FALSE,

            can_view_reports BOOLEAN DEFAULT FALSE,
            can_export BOOLEAN DEFAULT FALSE,
            can_backup BOOLEAN DEFAULT FALSE,
            can_restore BOOLEAN DEFAULT FALSE,

        is_active BOOLEAN DEFAULT TRUE
        )
    """)

    permission_columns = [
    "can_add_sales",
    "can_edit_sales",
    "can_delete_sales",
    "can_add_expenses",
    "can_edit_expenses",
    "can_delete_expenses",
    "can_add_salary",
    "can_edit_salary",
    "can_delete_salary",
    "can_view_reports",
    "can_export",
    "can_backup",
    "can_restore",
    "is_active"
    ]

    for col in permission_columns:
        c.execute(f"""
        ALTER TABLE users
        ADD COLUMN IF NOT EXISTS {col} BOOLEAN DEFAULT TRUE
    """)

    c.execute("""
    INSERT INTO users (
        username, password, role,
        can_add_sales, can_edit_sales, can_delete_sales,
        can_add_expenses, can_edit_expenses, can_delete_expenses,
        can_add_salary, can_edit_salary, can_delete_salary,
        can_view_reports, can_export, can_backup, can_restore,
        is_active
    )
    VALUES (
        'admin', 'admin123', 'admin',
        TRUE, TRUE, TRUE,
        TRUE, TRUE, TRUE,
        TRUE, TRUE, TRUE,
        TRUE, TRUE, TRUE, TRUE,
        TRUE
    )
    ON CONFLICT (username) DO UPDATE SET
        password='admin123',
        role='admin',
        can_add_sales=TRUE,
        can_edit_sales=TRUE,
        can_delete_sales=TRUE,
        can_add_expenses=TRUE,
        can_edit_expenses=TRUE,
        can_delete_expenses=TRUE,
        can_add_salary=TRUE,
        can_edit_salary=TRUE,
        can_delete_salary=TRUE,
        can_view_reports=TRUE,
        can_export=TRUE,
        can_backup=TRUE,
        can_restore=TRUE,
        is_active=TRUE
""")

    c.execute("""
    CREATE TABLE IF NOT EXISTS companies (
        id SERIAL PRIMARY KEY,
        company_code TEXT UNIQUE,
        company_name TEXT,
        address TEXT,
        phone TEXT,
        is_active BOOLEAN DEFAULT TRUE,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
    """)

    c.execute("""
    ALTER TABLE users
    ADD COLUMN IF NOT EXISTS company_code TEXT
    """)

    c.execute("""
    ALTER TABLE sales
    ADD COLUMN IF NOT EXISTS company_code TEXT
    """)

    c.execute("""
    ALTER TABLE expenses
    ADD COLUMN IF NOT EXISTS company_code TEXT
    """)

    c.execute("""
    ALTER TABLE salaries
    ADD COLUMN IF NOT EXISTS company_code TEXT
    """)

    conn.commit()
    conn.close()

init_db()

@app.route("/register-owner", methods=["GET", "POST"])
def register_owner():

    if request.method == "POST":
        company_code = request.form["company_code"]
        username = request.form["username"]
        password = request.form["password"]

        conn = get_conn()
        c = conn.cursor()

        c.execute("""
            SELECT company_code 
            FROM companies 
            WHERE company_code=%s AND is_active=TRUE
        """, (company_code,))

        company = c.fetchone()

        if not company:
            conn.close()
            return "Invalid Company Code"

        c.execute("""
            INSERT INTO users (username, password, role, company_code, is_active)
            VALUES (%s, %s, 'owner', %s, TRUE)
        """, (username, password, company_code))

        conn.commit()
        conn.close()

        return redirect("/login")

    return """
    <h1>Register Owner Account</h1>

    <form method="POST">
        Company Code:<br>
        <input name="company_code" required><br><br>

        Username:<br>
        <input name="username" required><br><br>

        Password:<br>
        <input type="password" name="password" required><br><br>

        <button type="submit">Create Owner Account</button>
    </form>

    <br>
    <a href="/login">Login</a>
    """

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form["username"]
        password = request.form["password"]

        conn = get_conn()
        c = conn.cursor()

        c.execute("""
            SELECT
                username, role,
                can_add_sales, can_edit_sales, can_delete_sales,
                can_add_expenses, can_edit_expenses, can_delete_expenses,
                can_add_salary, can_edit_salary, can_delete_salary,
                can_view_reports, can_export, can_backup, can_restore, company_code
            FROM users
            WHERE username=%s AND password=%s AND is_active=TRUE
        """, (username, password))

        user = c.fetchone()
        conn.close()

        if user:
            session["logged_in"] = True
            session["username"] = user[0]
            session["role"] = user[1]
            session["can_add_sales"] = user[2]
            session["can_edit_sales"] = user[3]
            session["can_delete_sales"] = user[4]
            session["can_add_expenses"] = user[5]
            session["can_edit_expenses"] = user[6]
            session["can_delete_expenses"] = user[7]
            session["can_add_salary"] = user[8]
            session["can_edit_salary"] = user[9]
            session["can_delete_salary"] = user[10]
            session["can_view_reports"] = user[11]
            session["can_export"] = user[12]
            session["can_backup"] = user[13]
            session["can_restore"] = user[14]
            session["company_code"] = user[15]

            return redirect("/")
        else:
            return """
            <h1>Login Failed</h1>
            <p>Wrong username or password.</p>
            <a href="/login">Try Again</a>
            """

    return """
    <!DOCTYPE html>
    <html>
<head>
<title>Login</title>
<style>
body{
    font-family: Arial;
    background:#f4f6f9;
    display:flex;
    justify-content:center;
    align-items:center;
    height:100vh;
}

.login-box{
    background:white;
    padding:30px;
    border-radius:10px;
    width:350px;
    box-shadow:0 0 15px rgba(0,0,0,0.1);
}

h2{
    text-align:center;
    margin-bottom:20px;
}

input{
    width:100%;
    padding:10px;
    margin-top:5px;
    margin-bottom:15px;
    border:1px solid #ccc;
    border-radius:5px;
}

button{
    width:100%;
    padding:10px;
    background:#007bff;
    color:white;
    border:none;
    border-radius:5px;
    cursor:pointer;
}

button:hover{
    background:#0056b3;
}
</style>
</head>

<body>

<div class="login-box">
    <h2>Optic Shop Login</h2>

    <form method="POST">

        <label>Username</label>
        <input type="text" name="username" required>

        <label>Password</label>
        <input type="password" name="password" required>

        <button type="submit">Login</button>

    </form>
</div>

</body>
</html>
"""

def has_permission(permission):
    return session.get(permission, False) or session.get("role") == "admin"

@app.route("/")
def home():
    if not session.get("logged_in"):
        return redirect("/login")

    conn = get_conn()
    c = conn.cursor()

    from datetime import datetime

    current_month = datetime.now().strftime("%Y-%m")

    # Monthly Sales
    if session.get("role") == "admin":
        c.execute("""
        SELECT COALESCE(SUM(amount),0)
        FROM sales
        WHERE date LIKE %s
        """, (current_month + "%",))
    else:
        c.execute("""
        SELECT COALESCE(SUM(amount),0)
        FROM sales
        WHERE date LIKE %s AND company_code=%s
        """, (current_month + "%", session["company_code"]))

    total_sales = c.fetchone()[0]

    # Monthly Expenses
    if session.get("role") == "admin":
        c.execute("""
        SELECT COALESCE(SUM(amount),0)
        FROM expenses
        WHERE date LIKE %s
        """, (current_month + "%",))
    else:
        c.execute("""
        SELECT COALESCE(SUM(amount),0)
        FROM expenses
        WHERE date LIKE %s AND company_code=%s
        """, (current_month + "%", session["company_code"]))

    normal_expenses = c.fetchone()[0]

    # Monthly Salaries
    if session.get("role") == "admin":
        c.execute("""
        SELECT COALESCE(SUM(amount),0)
        FROM salaries
        WHERE date LIKE %s
        """, (current_month + "%",))
    else:
        c.execute("""
        SELECT COALESCE(SUM(amount),0)
        FROM salaries
        WHERE date LIKE %s AND company_code=%s
        """, (current_month + "%", session["company_code"]))

    total_salaries = c.fetchone()[0]

    total_expenses = normal_expenses + total_salaries
    profit = total_sales - total_expenses
    profit_color = "green"
    if profit < 0:
        profit_color = "red"

    conn.close()

    menu_html = ""

    if session.get("can_add_sales"):
        menu_html += '<a href="/sales"><button>Add Sales</button></a>'

    if session.get("can_add_expenses"):
        menu_html += '<a href="/expenses"><button>Add Expense</button></a>'

    if session.get("can_add_salary"):
        menu_html += '<a href="/salary"><button>Add Staff Salary</button></a>'

    if session.get("can_view_reports"):
        menu_html += '<a href="/reports"><button>Reports</button></a>'

    if session.get("role") in ["admin", "owner"]:
        menu_html += '<a href="/permissions"><button>Permissions</button></a>'
        menu_html += '<a href="/companies"><button>Companies</button></a>'

    menu_html += '<a href="/logout"><button>Logout</button></a>'

    return f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>Optic Shop System</title>
        <style>
body {{
    font-family: Arial;
    background: #f4f6f8;
    padding: 30px;
}}

.header {{
    background: #111827;
    color: white;
    padding: 20px;
    border-radius: 10px;
    margin-bottom: 25px;
}}

.cards {{
    display: flex;
    gap: 20px;
    flex-wrap: wrap;
    margin-bottom: 30px;
}}

.card {{
    background: white;
    padding: 25px;
    border-radius: 10px;
    width: 220px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.1);
}}

.card h3 {{
    margin: 0;
    color: #555;
}}

.amount {{
    font-size: 28px;
    font-weight: bold;
    margin-top: 10px;
}}

.menu {{
    display: flex;
    flex-wrap: wrap;
    gap: 10px;
}}

button {{
    background: #111827;
    color: white;
    border: none;
    padding: 12px 18px;
    border-radius: 8px;
    cursor: pointer;
}}

button:hover {{
    background: #2563eb;
}}
</style>
    </head>

    <body>
        <div class="header">
            <h1>OPTIC SHOP MANAGEMENT SYSTEM</h1>
            <p>Business Dashboard</p>
            <div id="datetime" style="
                color:white;
                font-size:18px;
                font-weight:bold;
                margin-top:10px;
        "></div>
        </div>

        <div class="cards">
            <div class="card">
                <h3>This Month Sales</h3>
                <div class="amount">RM {total_sales:,.2f}</div>
            </div>

            <div class="card">
                <h3>This Month Expenses</h3>
                <div class="amount">RM {total_expenses:,.2f}</div>
            </div>

            <div class="card">
                <h3>This Month P&L</h3>
                <div class="amount" style="color:{profit_color}">RM {profit:,.2f}
            </div>
        </div>

        <div class="menu">
            {menu_html}
        </div>

        <script>
            function updateDateTime() {{
            const now = new Date();

            const date = now.toLocaleDateString('en-GB', {{
                weekday: 'long',
                year: 'numeric',
                month: 'long',
                day: 'numeric'
            }});

    const time = now.toLocaleTimeString('en-GB');

    document.getElementById("datetime").innerHTML =
        date + "<br>" + time;
}}

updateDateTime();
setInterval(updateDateTime, 1000);
</script>

</body>
</html>
"""

@app.route("/sales", methods=["GET", "POST"])
def sales():
    if not session.get("logged_in"):
        return redirect("/login")
    
    if not has_permission("can_add_sales"):
        return "Access Denied"

    if request.method == "POST":
        sale_date = request.form["sale_date"]
        customer = request.form["customer"]
        amount = float(request.form["amount"])
        staff = request.form["staff"]

        conn = get_conn()
        c = conn.cursor()

        c.execute("""
            INSERT INTO sales (
                date,
                customer,
                amount,
                staff,
                company_code
            )
            VALUES (%s, %s, %s, %s, %s)
        """, (
        sale_date,
        customer,
        amount,
        staff,
        session["company_code"]
    ))

        conn.commit()
        conn.close()

        return redirect("/sales-list")

    return f"""
    <h1>Add Sales</h1>

    <form method="POST">

        <label>Date:</label><br>
        <input type="date" name="sale_date" value="{datetime.now().strftime('%Y-%m-%d')}" required><br><br>

        <label>Customer:</label><br>
        <input type="text" name="customer" required><br><br>

        <label>Amount (RM):</label><br>
        <input type="number" step="0.01" name="amount" required><br><br>

        <label>Staff:</label><br>
        <input type="text" name="staff" required><br><br>

        <button type="submit">Save</button>

        <button type="button"
        onclick="window.location.href='/'">
           Cancel
        </button>

    </form>

    <br>
    <a href="/">Back to Dashboard</a>
    """


@app.route("/sales-list")
def sales_list():
    if not session.get("logged_in"):
        return redirect("/login")

    today = datetime.now()
    first_day = today.replace(day=1).strftime("%Y-%m-%d")
    last_day = today.replace(day=calendar.monthrange(today.year, today.month)[1]).strftime("%Y-%m-%d")

    from_date = request.args.get("from_date", first_day)
    to_date = request.args.get("to_date", last_day)

    conn = get_conn()
    c = conn.cursor()

    c.execute("""
        SELECT id, customer, amount, staff, date
        FROM sales
        WHERE date BETWEEN %s AND %s
        ORDER BY date DESC, id DESC
    """, (from_date, to_date))

    records = c.fetchall()
    conn.close()

    rows = ""
    for r in records:
        rows += f"""
        <tr>
            <td>{str(r[4])[:10]}</td>
            <td>{r[1]}</td>
            <td>RM {float(r[2]):,.2f}</td>
            <td>{r[3]}</td>
            <td>
                <a href="/edit-sale/{r[0]}">Edit</a> |
                <a href="/delete-sale/{r[0]}" onclick="return confirm('Delete this sale?')">Delete</a>
            </td>
        </tr>
        """

    return f"""
    <h1>Sales List</h1>

    <form method="GET">
        <label>From Date:</label>
        <input type="date" name="from_date" value="{from_date}" required>

        <label>To Date:</label>
        <input type="date" name="to_date" value="{to_date}" required>

        <button type="submit">Search</button>
    </form>

    <br>

    <table border="1" cellpadding="10">
        <tr>
            <th>Date</th>
            <th>Customer</th>
            <th>Amount</th>
            <th>Staff</th>
            <th>Action</th>
        </tr>
        {rows}
    </table>

    <br>
    <a href="/">Back to Dashboard</a>
    """

@app.route("/delete-sale/<int:sale_id>")
def delete_sale(sale_id):
    if not session.get("logged_in"):
        return redirect("/login")

    conn = get_conn()
    c = conn.cursor()
    c.execute("DELETE FROM sales WHERE id=%s", (sale_id,))
    conn.commit()
    conn.close()

    return redirect("/sales-list")

@app.route("/edit-sale/<int:sale_id>", methods=["GET", "POST"])
def edit_sale(sale_id):
    if not session.get("logged_in"):
        return redirect("/login")

    conn = get_conn()
    c = conn.cursor()

    if request.method == "POST":
        sale_date = request.form["sale_date"]
        customer = request.form["customer"]
        amount = float(request.form["amount"].replace(",", ""))
        staff = request.form["staff"]

        c.execute("""
            UPDATE sales
            SET date=%s, customer=%s, amount=%s, staff=%s
            WHERE id=%s
        """, (sale_date, customer, amount, staff, sale_id))

        conn.commit()
        conn.close()
        return redirect("/sales-list")

    c.execute("SELECT id, customer, amount, staff, date FROM sales WHERE id=%s", (sale_id,))
    sale = c.fetchone()
    conn.close()

    return f"""
    <h1>Edit Sale</h1>

    <form method="POST">
        <label>Date:</label><br>
        <input type="date" name="sale_date" value="{str(sale[4])[:10]}" required><br><br>

        <label>Customer Name:</label><br>
        <input type="text" name="customer" value="{sale[1]}" required><br><br>

        <label>Sales Amount (RM):</label><br>
        <input type="text" name="amount" value="{float(sale[2]):,.2f}" required><br><br>

        <label>Staff Name:</label><br>
        <input type="text" name="staff" value="{sale[3]}" required><br><br>

        <button type="submit">Update Sale</button>
    </form>

    <br>
    <a href="/sales-list">Back to Sales List</a>
    """

@app.route("/salary", methods=["GET", "POST"])
def salary():
    if not session.get("logged_in"):
        return redirect("/login")
    
    if not has_permission("can_add_salary"):
        return "Access Denied"

    if request.method == "POST":
        salary_date = request.form["salary_date"]
        staff = request.form["staff"]
        amount = float(request.form["amount"].replace(",", ""))
        month = request.form["month"]

        conn = get_conn()
        c = conn.cursor()

        c.execute("""
            INSERT INTO salaries (date, staff, amount, month)
            VALUES (%s, %s, %s, %s)
        """, (salary_date, staff, amount, month))

        conn.commit()
        conn.close()

        return redirect("/salary-list")

    return f"""
    <h1>Add Staff Salary</h1>

    <form method="POST">

        <label>Date:</label><br>
        <input type="date" name="salary_date" value="{datetime.now().strftime('%Y-%m-%d')}" required><br><br>

        <label>Staff Name:</label><br>
        <input type="text" name="staff" required><br><br>

        <label>Salary Amount (RM):</label><br>
        <input type="text" name="amount" required><br><br>

        <label>Month:</label><br>
        <input type="month"
               name="month"
               value="{datetime.now().strftime('%Y-%m')}"
               required><br><br>

        <button type="submit">Save</button>

<button type="button"
        onclick="window.location.href='/'">
    Cancel
</button>

    </form>

    <br>
    <a href="/">Back to Dashboard</a>
    """

@app.route("/salary-list")
def salary_list():
    if not session.get("logged_in"):
        return redirect("/login")

    today = datetime.now()
    first_day = today.replace(day=1).strftime("%Y-%m-%d")
    last_day = today.replace(day=calendar.monthrange(today.year, today.month)[1]).strftime("%Y-%m-%d")

    from_date = request.args.get("from_date", first_day)
    to_date = request.args.get("to_date", last_day)

    conn = get_conn()
    c = conn.cursor()

    c.execute("""
        SELECT id, staff, amount, month, date
        FROM salaries
        WHERE date BETWEEN %s AND %s
        ORDER BY date DESC, id DESC
    """, (from_date, to_date))

    records = c.fetchall()
    conn.close()

    rows = ""
    for r in records:
        rows += f"""
        <tr>
            <td>{str(r[4])[:10]}</td>
            <td>{r[1]}</td>
            <td>RM {float(r[2]):,.2f}</td>
            <td>{r[3]}</td>
            <td>
                <a href="/edit-salary/{r[0]}">Edit</a> |
                <a href="/delete-salary/{r[0]}" onclick="return confirm('Delete this salary?')">Delete</a>
            </td>
        </tr>
        """

    return f"""
    <h1>Salary List</h1>

    <form method="GET">
        <label>From Date:</label>
        <input type="date" name="from_date" value="{from_date}" required>

        <label>To Date:</label>
        <input type="date" name="to_date" value="{to_date}" required>

        <button type="submit">Search</button>
    </form>

    <br>

    <table border="1" cellpadding="10">
        <tr>
            <th>Date</th>
            <th>Staff</th>
            <th>Amount</th>
            <th>Month</th>
            <th>Action</th>
        </tr>
        {rows}
    </table>

    <br>
    <a href="/">Back to Dashboard</a>
    """

@app.route("/edit-salary/<int:salary_id>", methods=["GET", "POST"])
def edit_salary(salary_id):
    if not session.get("logged_in"):
        return redirect("/login")

    conn = get_conn()
    c = conn.cursor()

    if request.method == "POST":
        salary_date = request.form["salary_date"]
        staff = request.form["staff"]
        amount = float(request.form["amount"].replace(",", ""))
        month = request.form["month"]

        c.execute("""
            UPDATE salaries
            SET date=%s, staff=%s, amount=%s, month=%s
            WHERE id=%s
        """, (salary_date, staff, amount, month, salary_id))

        conn.commit()
        conn.close()
        return redirect("/salary-list")

    c.execute("SELECT id, staff, amount, month, date FROM salaries WHERE id=%s", (salary_id,))
    salary = c.fetchone()
    conn.close()

    return f"""
    <h1>Edit Staff Salary</h1>

    <form method="POST">
        <label>Date:</label><br>
        <input type="date" name="salary_date" value="{str(salary[4])[:10]}" required><br><br>

        <label>Staff Name:</label><br>
        <input type="text" name="staff" value="{salary[1]}" required><br><br>

        <label>Salary Amount (RM):</label><br>
        <input type="text" name="amount" value="{float(salary[2]):,.2f}" required><br><br>

        <label>Month:</label><br>
        <input type="month" name="month" value="{salary[3]}" required><br><br>

        <button type="submit">Update Salary</button>
    </form>

    <br>
    <a href="/salary-list">Back to Salary List</a>
    """

@app.route("/delete-salary/<int:salary_id>")
def delete_salary(salary_id):
    if not session.get("logged_in"):
        return redirect("/login")

    conn = get_conn()
    c = conn.cursor()
    c.execute("DELETE FROM salaries WHERE id=%s", (salary_id,))
    conn.commit()
    conn.close()

    return redirect("/salary-list")

@app.route("/expenses", methods=["GET", "POST"])
def expense():
    if not session.get("logged_in"):
        return redirect("/login")
    
    if not has_permission("can_add_expenses"):
        return "Access Denied"

    if request.method == "POST":
        expense_date = request.form["expense_date"]
        category = request.form["category"]
        amount = float(request.form["amount"])
        note = request.form["note"]

        conn = get_conn()
        c = conn.cursor()

        c.execute("""
            INSERT INTO expenses (date, category, amount, note)
            VALUES (%s, %s, %s, %s)
        """, (expense_date, category, amount, note))

        conn.commit()
        conn.close()

        return redirect("/expenses-list")

    return f"""
    <h1>Add Expense</h1>

    <form method="POST">

        <label>Date:</label><br>
        <input type="date" name="expense_date" value="{datetime.now().strftime('%Y-%m-%d')}" required><br><br>

        <label>Category:</label><br>
        <input type="text" name="category" required><br><br>

        <label>Amount (RM):</label><br>
        <input type="number" step="0.01" name="amount" required><br><br>

        <label>Note:</label><br>
        <input type="text" name="note"><br><br>

        <button type="submit">Save</button>

<button type="button"
        onclick="window.location.href='/'">
    Cancel
</button>

    </form>

    <br>
    <a href="/">Back to Dashboard</a>
    """


@app.route("/expenses-list")
def expenses_list():
    if not session.get("logged_in"):
        return redirect("/login")

    today = datetime.now()
    first_day = today.replace(day=1).strftime("%Y-%m-%d")
    last_day = today.replace(day=calendar.monthrange(today.year, today.month)[1]).strftime("%Y-%m-%d")

    from_date = request.args.get("from_date", first_day)
    to_date = request.args.get("to_date", last_day)

    conn = get_conn()
    c = conn.cursor()

    c.execute("""
        SELECT id, category, amount, note, date
        FROM expenses
        WHERE date BETWEEN %s AND %s
        ORDER BY date DESC, id DESC
    """, (from_date, to_date))

    records = c.fetchall()
    conn.close()

    rows = ""
    for r in records:
        rows += f"""
        <tr>
            <td>{str(r[4])[:10]}</td>
            <td>{r[1]}</td>
            <td>RM {float(r[2]):,.2f}</td>
            <td>{r[3]}</td>
            <td>
                <a href="/edit-expense/{r[0]}">Edit</a> |
                <a href="/delete-expense/{r[0]}" onclick="return confirm('Delete this expense?')">Delete</a>
            </td>
        </tr>
        """

    return f"""
    <h1>Expenses List</h1>

    <form method="GET">
        <label>From Date:</label>
        <input type="date" name="from_date" value="{from_date}" required>

        <label>To Date:</label>
        <input type="date" name="to_date" value="{to_date}" required>

        <button type="submit">Search</button>
    </form>

    <br>

    <table border="1" cellpadding="10">
        <tr>
            <th>Date</th>
            <th>Category</th>
            <th>Amount</th>
            <th>Note</th>
            <th>Action</th>
        </tr>
        {rows}
    </table>

    <br>
    <a href="/">Back to Dashboard</a>
    """

@app.route("/edit-expense/<int:expense_id>", methods=["GET", "POST"])
def edit_expense(expense_id):
    if not session.get("logged_in"):
        return redirect("/login")

    conn = get_conn()
    c = conn.cursor()

    if request.method == "POST":
        expense_date = request.form["expense_date"]
        category = request.form["category"]
        amount = float(request.form["amount"].replace(",", ""))
        note = request.form["note"]

        c.execute("""
            UPDATE expenses
            SET date=%s, category=%s, amount=%s, note=%s
            WHERE id=%s
        """, (expense_date, category, amount, note, expense_id))

        conn.commit()
        conn.close()
        return redirect("/expenses-list")

    c.execute("SELECT id, category, amount, note, date FROM expenses WHERE id=%s", (expense_id,))
    expense = c.fetchone()
    conn.close()

    return f"""
    <h1>Edit Expense</h1>

    <form method="POST">
        <label>Date:</label><br>
        <input type="date" name="expense_date" value="{str(expense[4])[:10]}" required><br><br>

        <label>Category:</label><br>
        <input type="text" name="category" value="{expense[1]}" required><br><br>

        <label>Amount (RM):</label><br>
        <input type="text" name="amount" value="{float(expense[2]):,.2f}" required><br><br>

        <label>Note:</label><br>
        <input type="text" name="note" value="{expense[3]}"><br><br>

        <button type="submit">Update Expense</button>
    </form>

    <br>
    <a href="/expenses-list">Back to Expenses List</a>
    """

@app.route("/delete-expense/<int:expense_id>")
def delete_expense(expense_id):
    if not session.get("logged_in"):
        return redirect("/login")

    conn = get_conn()
    c = conn.cursor()
    c.execute("DELETE FROM expenses WHERE id=%s", (expense_id,))
    conn.commit()
    conn.close()

    return redirect("/expenses-list")


@app.route("/report", methods=["GET", "POST"])
def report():
    if not session.get("logged_in"):
        return redirect("/login")

    sales_total = 0
    expenses_total = 0
    salary_total = 0
    total_expenses = 0
    profit = 0
    from_date = ""
    to_date = ""
    report_generated = False

    if request.method == "POST":
        from_date = request.form["from_date"]
        to_date = request.form["to_date"]
        report_generated = True

        conn = get_conn()
        c = conn.cursor()

    if session.get("role") == "admin":
        c.execute("""
            SELECT COALESCE(SUM(amount),0)
            FROM sales
            WHERE date BETWEEN %s AND %s
        """, (from_date, to_date))
    else:
        c.execute("""
            SELECT COALESCE(SUM(amount),0)
            FROM sales
            WHERE date BETWEEN %s AND %s
            AND company_code=%s
        """, (from_date, to_date, session["company_code"]))

    sales_total = c.fetchone()[0]

    if session.get("role") == "admin":
        c.execute("""
            SELECT COALESCE(SUM(amount),0)
            FROM expenses
            WHERE date BETWEEN %s AND %s
        """, (from_date, to_date))
    else:
        c.execute("""
            SELECT COALESCE(SUM(amount),0)
            FROM expenses
            WHERE date BETWEEN %s AND %s
            AND company_code=%s
        """, (from_date, to_date, session["company_code"]))

    expenses_total = c.fetchone()[0]

    if session.get("role") == "admin":
        c.execute("""
            SELECT COALESCE(SUM(amount),0)
            FROM salaries
            WHERE date BETWEEN %s AND %s
        """, (from_date, to_date))
    else:
        c.execute("""
            SELECT COALESCE(SUM(amount),0)
            FROM salaries
            WHERE date BETWEEN %s AND %s
            AND company_code=%s
        """, (from_date, to_date, session["company_code"]))

        salary_total = c.fetchone()[0]

        conn.close()

        total_expenses = expenses_total + salary_total
        profit = sales_total - total_expenses

        return f"""
    <h1>Profit & Loss Report</h1>

    <form method="POST">
        <label>From Date:</label><br>
        <input type="date" name="from_date" value="{from_date}" required><br><br>

        <label>To Date:</label><br>
        <input type="date" name="to_date" value="{to_date}" required><br><br>

        <button type="submit">Generate Report</button>
    </form>

    <br>

    {f'''
    <h2>Report From {from_date} To {to_date}</h2>

    <table border="1" cellpadding="10">
        <tr>
            <th>Total Sales</th>
            <td>RM {sales_total:,.2f}</td>
        </tr>
        <tr>
            <th>Expenses</th>
            <td>RM {expenses_total:,.2f}</td>
        </tr>
        <tr>
            <th>Staff Salaries</th>
            <td>RM {salary_total:,.2f}</td>
        </tr>
        <tr>
            <th>Total Expenses</th>
            <td>RM {total_expenses:,.2f}</td>
        </tr>
        <tr>
            <th>Profit / Loss</th>
            <td>RM {profit:,.2f}</td>
        </tr>
    </table>
    ''' if report_generated else ''}

    <br>
    <a href="/">Back to Dashboard</a>
    """

@app.route("/reports")
def reports():
    if not session.get("logged_in"):
        return redirect("/login")

    return """
    <!DOCTYPE html>
    <html>
    <head>
        <title>Reports</title>
        <style>
            body {
                font-family: Arial, sans-serif;
                background: #f4f6f8;
                margin: 0;
                padding: 40px;
            }

            .container {
                max-width: 850px;
                margin: auto;
            }

            .header {
                background: #2f2f2f;
                color: white;
                padding: 30px;
                border-radius: 12px;
                margin-bottom: 30px;
            }

            .header h1 {
                margin: 0;
                font-size: 36px;
            }

            .header p {
                margin-top: 8px;
                color: #ddd;
            }

            .report-grid {
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(230px, 1fr));
                gap: 20px;
            }

            .report-card {
                background: white;
                padding: 25px;
                border-radius: 12px;
                box-shadow: 0 4px 12px rgba(0,0,0,0.08);
                text-decoration: none;
                color: #222;
                transition: 0.2s;
            }

            .report-card:hover {
                transform: translateY(-4px);
                box-shadow: 0 6px 18px rgba(0,0,0,0.15);
            }

            .report-card h2 {
                margin: 0 0 10px;
                font-size: 22px;
            }

            .report-card p {
                margin: 0;
                color: #666;
                font-size: 14px;
            }

            .back {
                display: inline-block;
                margin-top: 30px;
                background: #333;
                color: white;
                padding: 12px 18px;
                border-radius: 8px;
                text-decoration: none;
            }

            .back:hover {
                background: #111;
            }
        </style>
    </head>

    <body>
        <div class="container">

            <div class="header">
                <h1>Reports</h1>
                <p>View sales, expenses, salary and financial reports</p>
            </div>

            <div class="report-grid">

                <a href="/sales-list" class="report-card">
                    <h2>Sales List</h2>
                    <p>View sales records by date range</p>
                </a>

                <a href="/expenses-list" class="report-card">
                    <h2>Expenses List</h2>
                    <p>View business expenses by date range</p>
                </a>

                <a href="/salary-list" class="report-card">
                    <h2>Salary List</h2>
                    <p>View staff salary records by date range</p>
                </a>

                <a href="/report" class="report-card">
                    <h2>Monthly P&L</h2>
                    <p>Check profit and loss summary</p>
                </a>

                <a href="/search" class="report-card">
                    <h2>Detailed Report</h2>
                    <p>View sales, expenses and salary details together</p>
                </a>

                <a href="/backup-excel" class="report-card">
                    <h2>Backup Data</h2>
                    <p>Download all sales, expenses and salary data</p>
                </a>

                <form action="/restore-backup"
                      method="POST"
                      enctype="multipart/form-data"
                      class="report-card">

                    <h2>Restore Backup</h2>

                    <p>Upload backup Excel and restore data</p>

                    <input type="file"
                           name="backup_file"
                           accept=".xlsx"
                           required>

                    <br><br>

                    <button type="submit">
                        Restore Data
                    </button>

                </form>

            </div>

            <a href="/" class="back">Back to Dashboard</a>

        </div>
    
    <script>
    document.addEventListener("DOMContentLoaded", function() {

        const urlParams = new URLSearchParams(window.location.search);

        if (urlParams.get("msg") === "restore_success") {

            const toast = document.createElement("div");

            toast.innerHTML = "✅ Backup restored successfully!";

            toast.style.position = "fixed";
            toast.style.top = "50%";
            toast.style.left = "50%";
            toast.style.transform = "translate(-50%, -50%)";

            toast.style.background = "#28a745";
            toast.style.color = "#fff";
            toast.style.padding = "30px 60px";
            toast.style.borderRadius = "15px";
            toast.style.fontWeight = "bold";
            toast.style.fontSize = "24px";
            toast.style.boxShadow = "0 10px 30px rgba(0,0,0,0.4)";
            toast.style.zIndex = "99999";

            document.body.appendChild(toast);

            setTimeout(() => {
                toast.remove();
            }, 3000);

            window.history.replaceState({}, document.title, "/reports");
        }

    });
    </script>

    </body>
    </html>
    """

@app.route("/search", methods=["GET", "POST"])
def search():
    if not session.get("logged_in"):
        return redirect("/login")

    keyword = ""
    sales_results = []
    expenses_results = []
    salary_results = []

    today = datetime.now()
    from_date = today.replace(day=1).strftime("%Y-%m-%d")
    to_date = today.strftime("%Y-%m-%d")

    if request.method == "POST":

        from_date = request.form["from_date"]
        to_date = request.form["to_date"]

    conn = get_conn()
    c = conn.cursor()

    # Sales
    if session.get("role") == "admin":
        c.execute("""
            SELECT date, customer, amount, staff
            FROM sales
            WHERE date BETWEEN %s AND %s
            ORDER BY date DESC, id DESC
        """, (from_date, to_date))
    else:
        c.execute("""
            SELECT date, customer, amount, staff
            FROM sales
            WHERE date BETWEEN %s AND %s
            AND company_code=%s
            ORDER BY date DESC, id DESC
        """, (from_date, to_date, session["company_code"]))

    sales_results = c.fetchall()

    # Expenses
    if session.get("role") == "admin":
        c.execute("""
            SELECT date, category, amount, note
            FROM expenses
            WHERE date BETWEEN %s AND %s
            ORDER BY date DESC, id DESC
        """, (from_date, to_date))
    else:
        c.execute("""
            SELECT date, category, amount, note
            FROM expenses
            WHERE date BETWEEN %s AND %s
            AND company_code=%s
            ORDER BY date DESC, id DESC
        """, (from_date, to_date, session["company_code"]))

    expenses_results = c.fetchall()

    # Salaries
    if session.get("role") == "admin":
        c.execute("""
            SELECT date, staff, amount, month
            FROM salaries
            WHERE date BETWEEN %s AND %s
            ORDER BY date DESC, id DESC
        """, (from_date, to_date))
    else:
        c.execute("""
            SELECT date, staff, amount, month
            FROM salaries
            WHERE date BETWEEN %s AND %s
            AND company_code=%s
            ORDER BY date DESC, id DESC
        """, (from_date, to_date, session["company_code"]))

    salary_results = c.fetchall()

    conn.close()

    sales_rows = ""
    for r in sales_results:
        sales_rows += f"""
        <tr>
            <td>{str(r[0])[:10]}</td>
            <td>{r[1]}</td>
            <td>RM {float(r[2]):,.2f}</td>
            <td>{r[3]}</td>
        </tr>
        """

    expenses_rows = ""
    for r in expenses_results:
        expenses_rows += f"""
        <tr>
            <td>{str(r[0])[:10]}</td>
            <td>{r[1]}</td>
            <td>RM {float(r[2]):,.2f}</td>
            <td>{r[3]}</td>
        </tr>
        """

    salary_rows = ""
    for r in salary_results:
        salary_rows += f"""
        <tr>
            <td>{str(r[0])[:10]}</td>
            <td>{r[1]}</td>
            <td>RM {float(r[2]):,.2f}</td>
            <td>{r[3]}</td>
        </tr>
        """

    return f"""
    <h1>Detailed Report</h1>

    <form method="POST">
        <label>From Date:</label>
        <input type="date" name="from_date" value="{from_date}" required>

        <label>To Date:</label>
        <input type="date" name="to_date" value="{to_date}" required>

        <br><br>

        <button type="submit">Search</button>

        <a href="/export-excel">
            <button type="button">Export Excel</button>
        </a>

        <a href="/export-pdf">
            <button type="button">Export PDF</button>
        </a>


    </form>

    <hr>

    <h2>Sales Results</h2>
    <table border="1" cellpadding="10">
        <tr>
            <th>Date</th>
            <th>Customer</th>
            <th>Amount</th>
            <th>Staff</th>
        </tr>
        {sales_rows}
    </table>

    <h2>Expenses Results</h2>
    <table border="1" cellpadding="10">
        <tr>
            <th>Date</th>
            <th>Category</th>
            <th>Amount</th>
            <th>Note</th>
        </tr>
        {expenses_rows}
    </table>

    <h2>Salary Results</h2>
    <table border="1" cellpadding="10">
        <tr>
            <th>Date</th>
            <th>Staff</th>
            <th>Amount</th>
            <th>Month</th>
        </tr>
        {salary_rows}
    </table>

    <br>
    <a href="/">Back to Dashboard</a>
    """

@app.route("/export-excel")
def export_excel():

    wb = Workbook()

    ws = wb.active
    ws.title = "Detailed Report"

    ws.append(["Date", "Type", "Description", "Amount"])

    conn = get_conn()
    c = conn.cursor()

    # Sales
    c.execute("""
    SELECT date, customer, amount, staff
    FROM sales
    ORDER BY date DESC
    """)
    sales = c.fetchall()

    for r in sales:
        ws.append([str(r[0])[:10], "Sale", f"{r[1]} / Staff: {r[3]}", float(r[2])])

    # Expenses
    c.execute("""
    SELECT date, category, amount, note
    FROM expenses
    ORDER BY date DESC
    """)
    expenses = c.fetchall()

    for r in expenses:
        ws.append([str(r[0])[:10], "Expense", f"{r[1]} / {r[3]}", float(r[2])])

    # Salary
    c.execute("""
    SELECT date, staff, amount, month
    FROM salaries
    ORDER BY date DESC
    """)
    salaries = c.fetchall()

    for r in salaries:
        ws.append([str(r[0])[:10], "Salary", f"{r[1]} / Month: {r[3]}", float(r[2])])

    conn.close()

    for row in ws.iter_rows(min_row=2, min_col=4, max_col=4):
        for cell in row:
            cell.number_format = '#,##0.00'

    output = BytesIO()

    wb.save(output)

    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="Detailed_Report.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route("/backup-excel")
def backup_excel():
    if not session.get("logged_in"):
        return redirect("/login")

    conn = get_conn()
    c = conn.cursor()

    wb = Workbook()

    # Sales Sheet
    ws = wb.active
    ws.title = "Sales"
    ws.append(["ID", "Date", "Customer", "Amount", "Staff"])

    c.execute("SELECT id, date, customer, amount, staff FROM sales ORDER BY date DESC")
    for r in c.fetchall():
        ws.append([r[0], str(r[1])[:10], r[2], float(r[3]), r[4]])

    # Expenses Sheet
    ws2 = wb.create_sheet("Expenses")
    ws2.append(["ID", "Date", "Category", "Amount", "Note"])

    c.execute("SELECT id, date, category, amount, note FROM expenses ORDER BY date DESC")
    for r in c.fetchall():
        ws2.append([r[0], str(r[1])[:10], r[2], float(r[3]), r[4]])

    # Salaries Sheet
    ws3 = wb.create_sheet("Salaries")
    ws3.append(["ID", "Date", "Staff", "Amount", "Month"])

    c.execute("SELECT id, date, staff, amount, month FROM salaries ORDER BY date DESC")
    for r in c.fetchall():
        ws3.append([r[0], str(r[1])[:10], r[2], float(r[3]), r[4]])

    conn.close()

    # Format amount columns
    for sheet in [ws, ws2, ws3]:
        for row in sheet.iter_rows(min_row=2, min_col=4, max_col=4):
            for cell in row:
                cell.number_format = '#,##0.00'

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="optic_shop_backup.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route("/restore-backup", methods=["POST"])
def restore_backup():
    if not session.get("logged_in"):
        return redirect("/login")

    file = request.files.get("backup_file")

    if not file:
        return "No backup file uploaded"

    wb = load_workbook(file)

    conn = get_conn()
    c = conn.cursor()

    try:
        # Clear old data
        c.execute("DELETE FROM sales")
        c.execute("DELETE FROM expenses")
        c.execute("DELETE FROM salaries")

        # Restore Sales
        if "Sales" in wb.sheetnames:
            ws = wb["Sales"]

            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or not row[1]:
                    continue

                date = row[1]
                customer = row[2]
                amount = row[3]
                staff = row[4]

                c.execute("""
                    INSERT INTO sales (date, customer, amount, staff)
                    VALUES (%s, %s, %s, %s)
                """, (date, customer, amount, staff))

        # Restore Expenses
        if "Expenses" in wb.sheetnames:
            ws = wb["Expenses"]

            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or not row[1]:
                    continue

                date = row[1]
                category = row[2]
                amount = row[3]
                note = row[4]

                c.execute("""
                    INSERT INTO expenses (date, category, amount, note)
                    VALUES (%s, %s, %s, %s)
                """, (date, category, amount, note))

        # Restore Salaries
        if "Salaries" in wb.sheetnames:
            ws = wb["Salaries"]

            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or not row[1]:
                    continue

                date = row[1]
                staff = row[2]
                amount = row[3]
                month = row[4]

                c.execute("""
                    INSERT INTO salaries (date, staff, amount, month)
                    VALUES (%s, %s, %s, %s)
                """, (date, staff, amount, month))

        conn.commit()

    except Exception as e:
        conn.rollback()
        conn.close()
        return f"Restore failed: {e}"

    conn.close()

    return redirect("/reports?msg=restore_success")

@app.route("/export-pdf")
def export_pdf():

    conn = get_conn()
    c = conn.cursor()

    buffer = BytesIO()

    doc = SimpleDocTemplate(buffer)

    styles = getSampleStyleSheet()

    elements = []

    title = Paragraph(
        "<b>OPTIC SHOP REPORT</b>",
        styles["Title"]
    )

    elements.append(title)
    elements.append(Spacer(1,20))
    elements.append(
        Paragraph("<b>SALES</b>", styles["Heading2"])
    )

    c.execute("""
        SELECT date, customer, staff, amount
        FROM sales
        ORDER BY date DESC
    """)

    sales_data = [
        ["Date","Customer","Staff","Amount"]
    ]

    for r in c.fetchall():

        sales_data.append([
            str(r[0])[:10],
            r[1],
            r[2],
            f"{float(r[3]):,.2f}"
        ])

    sales_table = Table(
        sales_data,
        colWidths=[90,220,80,100]
    )

    sales_table.setStyle(TableStyle([
        ('GRID',(0,0),(-1,-1),1,colors.black),
        ('BACKGROUND',(0,0),(-1,0),colors.lightgrey),
        ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
        ('ALIGN',(3,1),(3,-1),'RIGHT')
    ]))

    elements.append(sales_table)
    elements.append(Spacer(1,20))
    elements.append(
        Paragraph("<b>EXPENSES</b>", styles["Heading2"])
    )

    c.execute("""
        SELECT date, category, note, amount
        FROM expenses
        ORDER BY date DESC
    """)

    expense_data = [
        ["Date","Category","Note","Amount"]
    ]

    for r in c.fetchall():

        expense_data.append([
            str(r[0])[:10],
            r[1],
            r[2],
            f"{float(r[3]):,.2f}"
        ])

    expense_table = Table(
        expense_data,
        colWidths=[90,140,180,80]
    )

    expense_table.setStyle(TableStyle([
        ('GRID',(0,0),(-1,-1),1,colors.black),
        ('BACKGROUND',(0,0),(-1,0),colors.lightgrey),
        ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
        ('ALIGN',(3,1),(3,-1),'RIGHT')
    ]))

    elements.append(expense_table)
    elements.append(Spacer(1,20))
    elements.append(
        Paragraph("<b>SALARIES</b>", styles["Heading2"])
    )

    c.execute("""
        SELECT date, staff, month, amount
        FROM salaries
        ORDER BY date DESC
    """)

    salary_data = [
        ["Date","Staff","Month","Amount"]
    ]

    for r in c.fetchall():

        salary_data.append([
            str(r[0])[:10],
            r[1],
            r[2],
            f"{float(r[3]):,.2f}"
        ])

    salary_table = Table(
        salary_data,
        colWidths=[90,220,100,80]
    )

    salary_table.setStyle(TableStyle([
        ('GRID',(0,0),(-1,-1),1,colors.black),
        ('BACKGROUND',(0,0),(-1,0),colors.lightgrey),
        ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
        ('ALIGN',(3,1),(3,-1),'RIGHT')
    ]))

    elements.append(salary_table)

    c.execute("SELECT COALESCE(SUM(amount),0) FROM sales")
    total_sales = c.fetchone()[0]

    c.execute("SELECT COALESCE(SUM(amount),0) FROM expenses")
    total_expenses = c.fetchone()[0]

    c.execute("SELECT COALESCE(SUM(amount),0) FROM salaries")
    total_salary = c.fetchone()[0]

    profit = total_sales - total_expenses - total_salary

    elements.append(Spacer(1,25))

    summary = [
        ["Total Sales", f"RM {total_sales:,.2f}"],
        ["Total Expenses", f"RM {total_expenses:,.2f}"],
        ["Total Salary", f"RM {total_salary:,.2f}"],
        ["Net Profit", f"RM {profit:,.2f}"]
    ]

    summary_table = Table(
        summary,
        colWidths=[180,180]
    )

    summary_table.setStyle(TableStyle([
        ('GRID',(0,0),(-1,-1),1,colors.black),
        ('FONTNAME',(0,0),(-1,-1),'Helvetica-Bold')
    ]))

    elements.append(summary_table)

    doc.build(elements)

    conn.close()

    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name="Detailed_Report.pdf",
        mimetype="application/pdf"
    )

@app.route("/permissions")
def permissions():

    if session.get("role") not in ["admin", "owner"]:
        return "Access Denied"

    conn = get_conn()
    c = conn.cursor()

    if session.get("role") == "admin":
        c.execute("""
            SELECT id, username, role
            FROM users
            ORDER BY username
    """)
    else:
        c.execute("""
            SELECT id, username, role
            FROM users
            WHERE role != 'admin'
            AND company_code%s      
            ORDER BY username
        """, (session["company_code"],))

    users = c.fetchall()
    conn.close()

    rows = ""

    for u in users:
        rows += f"""
        <tr>
            <td>{u[0]}</td>
            <td>{u[1]}</td>
            <td>{u[2]}</td>
            <td>
                <a href="/edit-user/{u[0]}">Edit Permission</a>
                 |
                <a href="/reset-password/{u[0]}">Reset Password</a>
                 |
                <a href="/toggle-user/{u[0]}">Enable / Disable</a>
                 |
                <a href="/delete-user/{u[0]}" onclick="return confirm('Delete this user?')">Delete</a>
            </td>
        </tr>
        """

    return f"""
    <h1>Permission Management</h1>

    <a href="/add-user">
        <button>Add User</button>
    </a>
    <br><br>

    <table border="1" cellpadding="10">
        <tr>
            <th>ID</th>
            <th>Username</th>
            <th>Role</th>
            <th>Action</th>
        </tr>

        {rows}
    </table>

    <br>
    <a href="/">Back Dashboard</a>
    """

@app.route("/companies")
def companies():

    if session.get("role") != "admin":
        return "Access Denied"

    conn = get_conn()
    c = conn.cursor()

    c.execute("""
        SELECT id, company_code, company_name, address, phone, is_active
        FROM companies
        ORDER BY company_name
    """)

    companies = c.fetchall()
    conn.close()

    rows = ""

    for co in companies:
        status = "Active" if co[5] else "Disabled"

        rows += f"""
        <tr>
            <td>{co[0]}</td>
            <td>{co[1]}</td>
            <td>{co[2]}</td>
            <td>{co[3]}</td>
            <td>{co[4]}</td>
            <td>{status}</td>
        </tr>
        """

    return f"""
    <h1>Company Management</h1>

    <a href="/add-company">
        <button>Add Company</button>
    </a>

    <br><br>

    <table border="1" cellpadding="10">
        <tr>
            <th>ID</th>
            <th>Company Code</th>
            <th>Company Name</th>
            <th>Address</th>
            <th>Phone</th>
            <th>Status</th>
        </tr>

        {rows}
    </table>

    <br>
    <a href="/">Back Dashboard</a>
    """

@app.route("/add-company", methods=["GET", "POST"])
def add_company():

    if session.get("role") != "admin":
        return "Access Denied"

    if request.method == "POST":
        company_code = request.form["company_code"]
        company_name = request.form["company_name"]
        address = request.form["address"]
        phone = request.form["phone"]

        conn = get_conn()
        c = conn.cursor()

        c.execute("""
            INSERT INTO companies
            (company_code, company_name, address, phone, is_active)
            VALUES (%s, %s, %s, %s, TRUE)
        """, (company_code, company_name, address, phone))

        conn.commit()
        conn.close()

        return redirect("/companies")

    return """
    <h1>Add Company</h1>

    <form method="POST">
        Company Code:<br>
        <input type="text" name="company_code" required><br><br>

        Company Name:<br>
        <input type="text" name="company_name" required><br><br>

        Address:<br>
        <input type="text" name="address"><br><br>

        Phone:<br>
        <input type="text" name="phone"><br><br>

        <button type="submit">Create Company</button>
    </form>

    <br>
    <a href="/companies">Back</a>
    """

@app.route("/add-user", methods=["GET", "POST"])
def add_user():

    if session.get("role") not in ["admin", "owner"]:
        return "Access Denied"

    if request.method == "POST":
        username = request.form["username"]
        password = request.form["password"]
        role = request.form["role"]

        if session.get("role") == "owner" and role in ["admin", "owner"]:
            return "Access Denied"

        conn = get_conn()
        c = conn.cursor()

        c.execute("""
            INSERT INTO users (username, password, role, company_code, is_active)
            VALUES (%s, %s, %s, %s, TRUE)
        """, (username, password, role, session["company_code"]))

        conn.commit()
        conn.close()

        return redirect("/permissions")

    return """
    <h1>Add User</h1>

    <form method="POST">
        Username:<br>
        <input type="text" name="username" required><br><br>

        Password:<br>
        <input type="text" name="password" required><br><br>

        Role:<br>
        <select name="role">
            <option value="staff">Staff</option>
            <option value="manager">Manager</option>
            <option value="owner">Owner</option>
        </select><br><br>

        <button type="submit">Create User</button>
    </form>

    <br>
    <a href="/permissions">Back</a>
    """

@app.route("/edit-user/<int:user_id>", methods=["GET", "POST"])
def edit_user(user_id):

    if session.get("role") not in ["admin", "owner"]:
        return "Access Denied"

    conn = get_conn()
    c = conn.cursor()

    if session.get("role") == "owner":
        c.execute(
            "SELECT role FROM users WHERE id=%s",
            (user_id,)
        )
        target = c.fetchone()

        if not target or target[0] == "admin":
            conn.close()
            return "Access Denied"

    if request.method == "POST":
        role = request.form["role"]

        permissions = [
            "can_add_sales", "can_edit_sales", "can_delete_sales",
            "can_add_expenses", "can_edit_expenses", "can_delete_expenses",
            "can_add_salary", "can_edit_salary", "can_delete_salary",
            "can_view_reports", "can_export", "can_backup", "can_restore",
            "is_active"
        ]

        values = [role]
        for p in permissions:
            values.append(p in request.form)

        values.append(user_id)

        c.execute("""
            UPDATE users SET
                role=%s,
                can_add_sales=%s,
                can_edit_sales=%s,
                can_delete_sales=%s,
                can_add_expenses=%s,
                can_edit_expenses=%s,
                can_delete_expenses=%s,
                can_add_salary=%s,
                can_edit_salary=%s,
                can_delete_salary=%s,
                can_view_reports=%s,
                can_export=%s,
                can_backup=%s,
                can_restore=%s,
                is_active=%s
            WHERE id=%s
        """, values)

        conn.commit()
        conn.close()

        return redirect("/permissions")

    c.execute("""
        SELECT id, username, role,
               can_add_sales, can_edit_sales, can_delete_sales,
               can_add_expenses, can_edit_expenses, can_delete_expenses,
               can_add_salary, can_edit_salary, can_delete_salary,
               can_view_reports, can_export, can_backup, can_restore,
               is_active
        FROM users
        WHERE id=%s
    """, (user_id,))

    u = c.fetchone()
    conn.close()

    def checked(value):
        return "checked" if value else ""

    return f"""
    <h1>Edit Permission: {u[1]}</h1>

    <form method="POST">
        <label>Role:</label>
        <select name="role">
            <option value="admin" {"selected" if u[2]=="admin" else ""}>Admin</option>
            <option value="manager" {"selected" if u[2]=="manager" else ""}>Manager</option>
            <option value="staff" {"selected" if u[2]=="staff" else ""}>Staff</option>
        </select>

        <h3>Sales</h3>
        <label><input type="checkbox" name="can_add_sales" {checked(u[3])}> Add Sales</label><br>
        <label><input type="checkbox" name="can_edit_sales" {checked(u[4])}> Edit Sales</label><br>
        <label><input type="checkbox" name="can_delete_sales" {checked(u[5])}> Delete Sales</label><br>

        <h3>Expenses</h3>
        <label><input type="checkbox" name="can_add_expenses" {checked(u[6])}> Add Expenses</label><br>
        <label><input type="checkbox" name="can_edit_expenses" {checked(u[7])}> Edit Expenses</label><br>
        <label><input type="checkbox" name="can_delete_expenses" {checked(u[8])}> Delete Expenses</label><br>

        <h3>Salary</h3>
        <label><input type="checkbox" name="can_add_salary" {checked(u[9])}> Add Salary</label><br>
        <label><input type="checkbox" name="can_edit_salary" {checked(u[10])}> Edit Salary</label><br>
        <label><input type="checkbox" name="can_delete_salary" {checked(u[11])}> Delete Salary</label><br>

        <h3>Reports</h3>
        <label><input type="checkbox" name="can_view_reports" {checked(u[12])}> View Reports</label><br>
        <label><input type="checkbox" name="can_export" {checked(u[13])}> Export</label><br>
        <label><input type="checkbox" name="can_backup" {checked(u[14])}> Backup</label><br>
        <label><input type="checkbox" name="can_restore" {checked(u[15])}> Restore</label><br>

        <h3>Status</h3>
        <label><input type="checkbox" name="is_active" {checked(u[16])}> Active</label><br><br>

        <button type="submit">Save Permission</button>
    </form>

    <br>
    <a href="/permissions">Back</a>
    """

@app.route("/delete-user/<int:user_id>")
def delete_user(user_id):

    if session.get("role") not in ["admin", "owner"]:
        return "Access Denied"

    conn = get_conn()
    c = conn.cursor()

    if session.get("role") == "owner":
        c.execute(
            "SELECT role FROM users WHERE id=%s",
            (user_id,)
        )
        target = c.fetchone()

        if target and target[0] == "admin":
            conn.close()
            return "Access Denied"

    c.execute(
        "DELETE FROM users WHERE id=%s",
        (user_id,)
    )

    conn.commit()
    conn.close()

    return redirect("/permissions")

@app.route("/reset-password/<int:user_id>", methods=["GET", "POST"])
def reset_password(user_id):

    if session.get("role") not in ["admin", "owner"]:
        return "Access Denied"

    conn = get_conn()
    c = conn.cursor()

    if session.get("role") == "owner":
        c.execute("SELECT role FROM users WHERE id=%s", (user_id,))
        target = c.fetchone()

        if target and target[0] == "admin":
            conn.close()
            return "Access Denied"

    if request.method == "POST":
        new_password = request.form["password"]

        c.execute(
            "UPDATE users SET password=%s WHERE id=%s",
            (new_password, user_id)
        )

        conn.commit()
        conn.close()

        return redirect("/permissions")

    conn.close()

    return """
    <h1>Reset Password</h1>

    <form method="POST">
        New Password:<br>
        <input type="text" name="password" required><br><br>

        <button type="submit">Save Password</button>
    </form>

    <br>
    <a href="/permissions">Back</a>
    """

@app.route("/toggle-user/<int:user_id>")
def toggle_user(user_id):

    if session.get("role") not in ["admin", "owner"]:
        return "Access Denied"

    conn = get_conn()
    c = conn.cursor()

    # owner不能动admin
    if session.get("role") == "owner":
        c.execute(
            "SELECT role FROM users WHERE id=%s",
            (user_id,)
        )
        target = c.fetchone()

        if target and target[0] == "admin":
            conn.close()
            return "Access Denied"

    c.execute(
        "SELECT is_active FROM users WHERE id=%s",
        (user_id,)
    )

    user = c.fetchone()

    if user:
        new_status = not user[0]

        c.execute(
            "UPDATE users SET is_active=%s WHERE id=%s",
            (new_status, user_id)
        )

        conn.commit()

    conn.close()

    return redirect("/permissions")

@app.route("/logout")
def logout():
    session.pop("logged_in", None)
    return redirect("/login")

if __name__ == "__main__":
    init_db()
    app.run(debug=True)