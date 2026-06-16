from flask import send_file
from io import BytesIO
from openpyxl import Workbook
from reportlab.pdfgen import canvas

from flask import Flask, request, session, redirect
from datetime import datetime
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
        CREATE TABLE IF NOT EXISTS expenses (
            id SERIAL PRIMARY KEY,
            category TEXT,
            amount REAL,
            note TEXT,
            date TEXT DEFAULT CURRENT_TIMESTAMP
        )
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

    conn.commit()
    conn.close()

init_db()

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form["username"]
        password = request.form["password"]

        if username == "admin" and password == "1234":
            session["logged_in"] = True
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


@app.route("/")
def home():
    if not session.get("logged_in"):
        return redirect("/login")

    conn = get_conn()
    c = conn.cursor()

    from datetime import datetime

    current_month = datetime.now().strftime("%Y-%m")

    # Monthly Sales
    c.execute("""
    SELECT COALESCE(SUM(amount),0)
    FROM sales
    WHERE date LIKE %s
    """, (current_month + "%",))
    total_sales = c.fetchone()[0]

    # Monthly Expenses
    c.execute("""
    SELECT COALESCE(SUM(amount),0)
    FROM expenses
    WHERE date LIKE %s
    """, (current_month + "%",))
    normal_expenses = c.fetchone()[0]

    # Monthly Salaries
    c.execute("""
    SELECT COALESCE(SUM(amount),0)
    FROM salaries
    WHERE month = %s
    """, (current_month,))

    total_salaries = c.fetchone()[0]

    total_expenses = normal_expenses + total_salaries
    profit = total_sales - total_expenses
    profit_color = "green"
    if profit < 0:
        profit_color = "red"

    conn.close()

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
            <a href="/sales"><button>Add Sales</button></a>
            <a href="/expenses"><button>Add Expense</button></a>
            <a href="/salary"><button>Add Staff Salary</button></a>
            <a href="/reports"><button>Reports</button></a>
            <a href="/logout"><button>Logout</button></a>
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

    if request.method == "POST":
        sale_date = request.form["sale_date"]
        customer = request.form["customer"]
        amount = float(request.form["amount"])
        staff = request.form["staff"]

        conn = get_conn()
        c = conn.cursor()

        c.execute("""
            INSERT INTO sales (date, customer, amount, staff)
            VALUES (%s, %s, %s, %s)
        """, (sale_date, customer, amount, staff))

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

        c.execute("""
            SELECT COALESCE(SUM(amount), 0)
            FROM sales
            WHERE date BETWEEN %s AND %s
        """, (from_date, to_date))
        sales_total = c.fetchone()[0]

        c.execute("""
            SELECT COALESCE(SUM(amount), 0)
            FROM expenses
            WHERE date BETWEEN %s AND %s
        """, (from_date, to_date))
        expenses_total = c.fetchone()[0]

        c.execute("""
            SELECT COALESCE(SUM(amount), 0)
            FROM salaries
            WHERE date BETWEEN %s AND %s
        """, (from_date, to_date))
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

            </div>

            <a href="/" class="back">Back to Dashboard</a>

        </div>
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
    c.execute("""
        SELECT date, customer, amount, staff
        FROM sales
        WHERE date BETWEEN %s AND %s
        ORDER BY date DESC, id DESC
    """, (from_date, to_date))
    sales_results = c.fetchall()

    # Expenses
    c.execute("""
        SELECT date, category, amount, note
        FROM expenses
        WHERE date BETWEEN %s AND %s
        ORDER BY date DESC, id DESC
    """, (from_date, to_date))
    expenses_results = c.fetchall()

    # Salaries
    c.execute("""
        SELECT date, staff, amount, month
        FROM salaries
        WHERE date BETWEEN %s AND %s
        ORDER BY date DESC, id DESC
    """, (from_date, to_date))
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

@app.route("/export-pdf")
def export_pdf():

    conn = get_conn()
    c = conn.cursor()

    buffer = BytesIO()
    p = canvas.Canvas(buffer)

    y = 800

    p.setFont("Helvetica-Bold", 18)
    p.drawString(50, y, "OPTIC SHOP REPORT")

    y -= 40

    # SALES
    p.setFont("Helvetica-Bold", 12)
    p.drawString(50, y, "SALES")

    y -= 20

    p.setFont("Courier", 10)

    p.drawString(
        50,
        y,
        f"{'Date':<12}{'Customer':<25}{'Staff':<15}{'Amount':>15}"
    )

    y -= 15

    p.drawString(50, y, "-" * 90)

    y -= 15

    c.execute("""
        SELECT date, customer, amount, staff
        FROM sales
        ORDER BY date DESC
    """)

    for r in c.fetchall():

        line = (
            f"{str(r[0])[:10]:<12}"
            f"{str(r[1])[:25]}"
            f"{str(r[3]):<15}"
            f"{float(r[2]):>15,.2f}"
        )

        p.drawString(50, y, line)

        y -= 15

    y -= 25

    # EXPENSES
    p.setFont("Helvetica-Bold", 12)
    p.drawString(50, y, "EXPENSES")

    y -= 20

    p.setFont("Courier", 10)

    p.drawString(
        50,
        y,
        f"{'Date':<12}{'Category':<25}{'Note':<25}{'Amount':>15}"
    )

    y -= 15

    p.drawString(50, y, "-" * 90)

    y -= 15

    c.execute("""
        SELECT date, category, amount, note
        FROM expenses
        ORDER BY date DESC
    """)

    for r in c.fetchall():

        line = (
            f"{str(r[0])[:10]:<12}"
            f"{str(r[1])[:25]}"
            f"{str(r[3]):<25}"
            f"{float(r[2]):>15,.2f}"
        )

        p.drawString(50, y, line)

        y -= 15

    y -= 25

    # SALARIES
    p.setFont("Helvetica-Bold", 12)
    p.drawString(50, y, "SALARIES")

    y -= 20

    p.setFont("Courier", 10)

    p.drawString(
        50,
        y,
        f"{'Date':<12}{'Staff':<25}{'Month':<15}{'Amount':>15}"
    )

    y -= 15

    p.drawString(50, y, "-" * 90)

    y -= 15

    c.execute("""
        SELECT date, staff, amount, month
        FROM salaries
        ORDER BY date DESC
    """)

    for r in c.fetchall():

        line = (
            f"{str(r[0])[:10]:<12}"
            f"{str(r[1])[:25]}"
            f"{str(r[3]):<15}"
            f"{float(r[2]):>15,.2f}"
        )

        p.drawString(50, y, line)

        y -= 15

    # SUMMARY
    c.execute("SELECT COALESCE(SUM(amount),0) FROM sales")
    total_sales = c.fetchone()[0]

    c.execute("SELECT COALESCE(SUM(amount),0) FROM expenses")
    total_expenses = c.fetchone()[0]

    c.execute("SELECT COALESCE(SUM(amount),0) FROM salaries")
    total_salary = c.fetchone()[0]

    profit = total_sales - total_expenses - total_salary

    y -= 40

    p.setFont("Helvetica-Bold", 12)
    p.drawString(50, y, "SUMMARY")

    y -= 20

    p.setFont("Helvetica", 10)

    p.drawString(50, y, f"Total Sales     : RM {total_sales:,.2f}")
    y -= 15

    p.drawString(50, y, f"Total Expenses  : RM {total_expenses:,.2f}")
    y -= 15

    p.drawString(50, y, f"Total Salaries  : RM {total_salary:,.2f}")
    y -= 15

    p.drawString(50, y, f"Net Profit      : RM {profit:,.2f}")

    conn.close()

    p.save()

    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name="Detailed_Report.pdf",
        mimetype="application/pdf"
    )

@app.route("/logout")
def logout():
    session.pop("logged_in", None)
    return redirect("/login")

if __name__ == "__main__":
    init_db()
    app.run(debug=True)