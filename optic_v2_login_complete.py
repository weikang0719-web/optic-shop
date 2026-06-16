from flask import Flask, request, session, redirect
import sqlite3

app = Flask(__name__)
app.secret_key = "optic_shop_secret_key"

DB = "optic_shop.db"


def init_db():
    conn = sqlite3.connect(DB)
    c = conn.cursor()

    c.execute("""
        CREATE TABLE IF NOT EXISTS sales (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            customer TEXT,
            amount REAL,
            staff TEXT,
            date TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS expenses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            category TEXT,
            amount REAL,
            note TEXT,
            date TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS salaries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            staff TEXT,
            amount REAL,
            month TEXT,
            date TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)

    conn.commit()
    conn.close()


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

    conn = sqlite3.connect(DB)
    c = conn.cursor()

    total_sales = c.execute("SELECT COALESCE(SUM(amount), 0) FROM sales").fetchone()[0]
    normal_expenses = c.execute("SELECT COALESCE(SUM(amount), 0) FROM expenses").fetchone()[0]
    total_salaries = c.execute("SELECT COALESCE(SUM(amount), 0) FROM salaries").fetchone()[0]

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
                margin-top: 30px;
            }}

            button {{
                padding: 12px 18px;
                margin: 5px;
                border: none;
                border-radius: 6px;
                cursor: pointer;
                background: #2563eb;
                color: white;
                font-size: 15px;
            }}

            button:hover {{
                background: #1d4ed8;
            }}
        </style>
    </head>

    <body>
        <div class="header">
            <h1>OPTIC SHOP MANAGEMENT SYSTEM</h1>
            <p>Business Dashboard</p>
        </div>

        <div class="cards">
            <div class="card">
                <h3>Total Sales</h3>
                <div class="amount">RM {total_sales:.2f}</div>
            </div>

            <div class="card">
                <h3>Total Expenses</h3>
                <div class="amount">RM {total_expenses:.2f}</div>
            </div>

            <div class="card">
                <h3>Profit / Loss</h3>
                <div class="amount" style="color:{profit_color}">RM {profit:.2f}
            </div>
        </div>

        <div class="menu">
            <a href="/sales"><button>Add Sales</button></a>
            <a href="/sales-list"><button>View Sales List</button></a>
            <a href="/expenses"><button>Add Expense</button></a>
            <a href="/expenses-list"><button>View Expenses List</button></a>
            <a href="/salary"><button>Add Staff Salary</button></a>
            <a href="/salary-list"><button>View Salary List</button></a>
            <a href="/report"><button>Monthly P&L</button></a>
            <a href="/logout"><button>Logout</button></a>
        </div>
    </body>
    </html>
    """

@app.route("/sales", methods=["GET", "POST"])
def sales():
    if not session.get("logged_in"):
        return redirect("/login")

    if request.method == "POST":
        customer = request.form["customer"]
        amount = float(request.form["amount"])
        staff = request.form["staff"]

        conn = sqlite3.connect(DB)
        c = conn.cursor()
        c.execute(
            "INSERT INTO sales (customer, amount, staff) VALUES (?, ?, ?)",
            (customer, amount, staff)
        )
        conn.commit()
        conn.close()

        return """
        <h1>Sale Saved</h1>
        <a href="/sales">Add Another Sale</a><br>
        <a href="/">Back to Dashboard</a>
        """

    return """
    <h1>Sales Entry</h1>

    <form method="POST">
        Customer Name:<br>
        <input type="text" name="customer"><br><br>

        Sales Amount (RM):<br>
        <input type="number" name="amount"><br><br>

        Staff Name:<br>
        <input type="text" name="staff"><br><br>

        <button type="submit">Save Sale</button>
    </form>
    """


@app.route("/sales-list")
def sales_list():
    if not session.get("logged_in"):
        return redirect("/login")

    conn = sqlite3.connect(DB)
    c = conn.cursor()
    records = c.execute(
        "SELECT customer, amount, staff, date FROM sales ORDER BY id DESC"
    ).fetchall()
    conn.close()

    rows = ""
    for r in records:
        rows += f"""
        <tr>
            <td>{r[3]}</td>
            <td>{r[0]}</td>
            <td>RM {r[1]:.2f}</td>
            <td>{r[2]}</td>
        </tr>
        """

    return f"""
    <h1>Sales List</h1>

    <table border="1" cellpadding="8">
        <tr>
            <th>Date</th>
            <th>Customer</th>
            <th>Amount</th>
            <th>Staff</th>
        </tr>
        {rows}
    </table>

    <br>
    <a href="/">Back to Dashboard</a>
    """

@app.route("/salary", methods=["GET", "POST"])
def salary():
    if not session.get("logged_in"):
        return redirect("/login")
    if request.method == "POST":

        staff = request.form["staff"]
        amount = float(request.form["amount"])
        month = request.form["month"]

        conn = sqlite3.connect(DB)
        c = conn.cursor()

        c.execute(
            "INSERT INTO salaries (staff, amount, month) VALUES (?, ?, ?)",
            (staff, amount, month)
        )

        conn.commit()
        conn.close()

        return """
        <h1>Salary Saved</h1>

        <a href="/salary">Add Another Salary</a><br><br>

        <a href="/">Back to Dashboard</a>
        """

    return """
    <h1>Staff Salary Entry</h1>

    <form method="POST">

        Staff Name:<br>
        <input type="text" name="staff"><br><br>

        Salary Amount (RM):<br>
        <input type="number" name="amount"><br><br>

        Month:<br>
        <input type="text" name="month"><br><br>

        <button type="submit">
            Save Salary
        </button>

    </form>
    """


@app.route("/salary-list")
def salary_list():
    if not session.get("logged_in"):
        return redirect("/login")
    conn = sqlite3.connect(DB)
    c = conn.cursor()

    records = c.execute(
        "SELECT staff, amount, month, date FROM salaries ORDER BY id DESC"
    ).fetchall()

    conn.close()

    rows = ""

    for r in records:
        rows += f"""
        <tr>
            <td>{r[3]}</td>
            <td>{r[0]}</td>
            <td>RM {r[1]:.2f}</td>
            <td>{r[2]}</td>
        </tr>
        """

    return f"""
    <h1>Salary List</h1>

    <table border="1" cellpadding="8">
        <tr>
            <th>Date</th>
            <th>Staff</th>
            <th>Amount</th>
            <th>Month</th>
        </tr>
        {rows}
    </table>

    <br>
    <a href="/">Back to Dashboard</a>
    """


@app.route("/expenses", methods=["GET", "POST"])
def expenses():
    if not session.get("logged_in"):
        return redirect("/login")
    if request.method == "POST":
        category = request.form["category"]
        amount = float(request.form["amount"])
        note = request.form["note"]

        conn = sqlite3.connect(DB)
        c = conn.cursor()
        c.execute(
            "INSERT INTO expenses (category, amount, note) VALUES (?, ?, ?)",
            (category, amount, note)
        )
        conn.commit()
        conn.close()

        return """
        <h1>Expense Saved</h1>
        <a href="/expenses">Add Another Expense</a><br>
        <a href="/">Back to Dashboard</a>
        """

    return """
    <h1>Expense Entry</h1>

    <form method="POST">
        Category:<br>
        <input type="text" name="category"><br><br>

        Expense Amount (RM):<br>
        <input type="number" name="amount"><br><br>

        Note:<br>
        <input type="text" name="note"><br><br>

        <button type="submit">Save Expense</button>
    </form>
    """


@app.route("/expenses-list")
def expenses_list():
    if not session.get("logged_in"):
        return redirect("/login")
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    records = c.execute("SELECT category, amount, note, date FROM expenses ORDER BY id DESC").fetchall()
    conn.close()

    rows = ""
    for r in records:
        rows += f"""
        <tr>
            <td>{r[3]}</td>
            <td>{r[0]}</td>
            <td>RM {r[1]}</td>
            <td>{r[2]}</td>
        </tr>
        """

    return f"""
    <h1>Expenses List</h1>

    <table border="1" cellpadding="8">
        <tr>
            <th>Date</th>
            <th>Category</th>
            <th>Amount</th>
            <th>Note</th>
        </tr>
        {rows}
    </table>

    <br>
    <a href="/">Back to Dashboard</a>
    """


@app.route("/report")
def report():
    if not session.get("logged_in"):
        return redirect("/login")
    conn = sqlite3.connect(DB)
    c = conn.cursor()

    total_sales = c.execute(
        "SELECT COALESCE(SUM(amount), 0) FROM sales"
    ).fetchone()[0]

    normal_expenses = c.execute(
        "SELECT COALESCE(SUM(amount), 0) FROM expenses"
    ).fetchone()[0]

    total_salaries = c.execute(
        "SELECT COALESCE(SUM(amount), 0) FROM salaries"
    ).fetchone()[0]

    total_expenses = normal_expenses + total_salaries
    profit = total_sales - total_expenses

    conn.close()

    profit_color = "green"

    if profit < 0:
        profit_color = "red"

    return f"""
    <h1>Monthly Profit & Loss Report</h1>

    <table border="1" cellpadding="10">
        <tr>
            <th>Item</th>
            <th>Amount</th>
        </tr>

        <tr>
            <td>Total Sales</td>
            <td>RM {total_sales:.2f}</td>
        </tr>

        <tr>
            <td>Normal Expenses</td>
            <td>RM {normal_expenses:.2f}</td>
        </tr>

        <tr>
            <td>Staff Salary</td>
            <td>RM {total_salaries:.2f}</td>
        </tr>

        <tr>
            <td>Total Expenses</td>
            <td>RM {total_expenses:.2f}</td>
        </tr>

        <tr>
            <td><strong>Net Profit / Loss</strong></td>
            <td style="color:{profit_color}">
                <strong>RM {profit:.2f}</strong>
            </td>
        </tr>
    </table>

    <br>
    <a href="/">Back to Dashboard</a>
    """

@app.route("/logout")
def logout():
    session.pop("logged_in", None)
    return redirect("/login")

if __name__ == "__main__":
    init_db()
    app.run(debug=True)