total_sales = 0
total_expenses = 0

while True:

    print("\n========================")
    print("OPTIC SHOP SYSTEM")
    print("========================")
    print("1. Add Sales")
    print("2. Add Expense")
    print("3. View P&L")
    print("4. Exit")

    choice = input("Choose option: ")

    if choice == "1":
        sales = float(input("Sales Amount (RM): "))
        total_sales += sales
        print("Sales Added!")

    elif choice == "2":
        expense = float(input("Expense Amount (RM): "))
        total_expenses += expense
        print("Expense Added!")

    elif choice == "3":
        profit = total_sales - total_expenses

        print("\n----- P&L REPORT -----")
        print("Total Sales: RM", total_sales)
        print("Total Expenses: RM", total_expenses)
        print("Profit: RM", profit)

    elif choice == "4":
        print("Goodbye!")
        break

    else:
        print("Invalid Option")