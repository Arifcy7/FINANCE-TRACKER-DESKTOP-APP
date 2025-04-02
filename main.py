import csv
import datetime
import os
import sqlite3
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox as mb
from tkinter import ttk
import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from tkcalendar import DateEntry
from openpyxl import Workbook
from openpyxl.utils import get_column_letter


def adapt_date(date):
    return date.strftime("%Y-%m-%d")


def convert_date(s):
    return datetime.datetime.strptime(s, "%Y-%m-%d")


sqlite3.register_adapter(datetime.date, adapt_date)
sqlite3.register_converter("DATE", convert_date)


def listAllExpenses():
    global dbconnector, data_table
    data_table.delete(*data_table.get_children())
    all_data = dbconnector.execute('SELECT * FROM ExpenseTracker')
    data = all_data.fetchall()
    for val in data:
        data_table.insert('', END, values=val)


def viewExpenseInfo():
    global data_table
    global dateField, payee, description, amount, modeOfPayment, category
    if not data_table.selection():
        mb.showerror('No expense selected', 'Please select an expense from the table to view its details')
        return
    currentSelectedExpense = data_table.item(data_table.focus())
    val = currentSelectedExpense['values']
    expenditureDate = datetime.date(int(val[1][:4]), int(val[1][5:7]), int(val[1][8:]))
    dateField.set_date(expenditureDate)
    payee.set(val[2])
    description.set(val[3])
    amount.set(val[4])
    modeOfPayment.set(val[5])
    category.set(val[6])


def clearFields():
    global description, payee, amount, modeOfPayment, category, dateField, data_table
    todayDate = datetime.datetime.now().date()
    description.set('')
    payee.set('')
    amount.set(0.0)
    modeOfPayment.set('Cash')
    category.set('Food')
    dateField.set_date(todayDate)
    data_table.selection_remove(*data_table.selection())


def removeExpense():
    if not data_table.selection():
        mb.showerror('No record selected!', 'Please select a record to delete!')
        return
    currentSelectedExpense = data_table.item(data_table.focus())
    valuesSelected = currentSelectedExpense['values']
    confirmation = mb.askyesno('Are you sure?',
                               f'Are you sure that you want to delete the record of {valuesSelected[2]}')
    if confirmation:
        dbconnector.execute('DELETE FROM ExpenseTracker WHERE ID=%d' % valuesSelected[0])
        dbconnector.commit()
        listAllExpenses()
        mb.showinfo('Record deleted successfully!', 'The record you wanted to delete has been deleted successfully')


def removeAllExpenses():
    confirmation = mb.askyesno('Are you sure?',
                               'Are you sure that you want to delete all the expense items from the database?',
                               icon='warning')
    if confirmation:
        data_table.delete(*data_table.get_children())
        dbconnector.execute('DELETE FROM ExpenseTracker')
        dbconnector.commit()
        clearFields()
        listAllExpenses()
        mb.showinfo('All Expenses deleted', 'All the expenses were successfully deleted')
    else:
        mb.showinfo('Ok then', 'The task was aborted and no expense was deleted!')


def addAnotherExpense():
    global dateField, payee, description, amount, modeOfPayment, category, dbconnector
    if not dateField.get() or not payee.get() or not description.get() or not amount.get() or not modeOfPayment.get() or not category.get():
        mb.showerror('Fields empty!', "Please fill all the missing fields before pressing the add button!")
        return
    try:
        float(amount.get())
    except ValueError:
        mb.showerror('Invalid Amount', 'Please enter a valid amount')
        return
    if float(amount.get()) <= 0:
        mb.showerror('Invalid Amount', 'Please enter a positive amount')
        return
    dbconnector.execute(
        'INSERT INTO ExpenseTracker (Date, Payee, Description, Amount, ModeOfPayment, Category) VALUES (?, ?, ?, ?, ?, ?)',
        (dateField.get_date(), payee.get(), description.get(), amount.get(), modeOfPayment.get(), category.get())
    )
    dbconnector.commit()
    clearFields()
    listAllExpenses()
    mb.showinfo('Expense added', 'The expense whose details you just entered has been added to the database')


def editExpense():
    def editExistingExpense():
        global dateField, amount, description, payee, modeOfPayment, category
        global dbconnector, data_table
        currentSelectedExpense = data_table.item(data_table.focus())
        content = currentSelectedExpense['values']
        dbconnector.execute(
            'UPDATE ExpenseTracker SET Date = ?, Payee = ?, Description = ?, Amount = ?, ModeOfPayment = ?, Category = ? WHERE ID = ?',
            (dateField.get_date(), payee.get(), description.get(), amount.get(), modeOfPayment.get(), category.get(), content[0])
        )
        dbconnector.commit()
        clearFields()
        listAllExpenses()
        mb.showinfo('Data edited', 'We have updated the data and stored in the database as you wanted')
        editSelectedButton.destroy()

    if not data_table.selection():
        mb.showerror('No expense selected!',
                     'You have not selected any expense in the table for us to edit; please do that!')
        return
    viewExpenseInfo()
    editSelectedButton = Button(
        frameL3,
        text="Edit Expense",
        font=("Bahnschrift Condensed", "13"),
        width=30,
        bg="#90EE90",
        fg="#000000",
        relief=GROOVE,
        activebackground="#008000",
        activeforeground="#FF0000",
        command=editExistingExpense
    )
    editSelectedButton.grid(row=0, column=0, sticky=W, padx=50, pady=10)


def selectedExpenseToWords():
    global data_table
    if not data_table.selection():
        mb.showerror('No expense selected!', 'Please select an expense from the table for us to read')
        return
    currentSelectedExpense = data_table.item(data_table.focus())
    val = currentSelectedExpense['values']
    msg = f'Your expense can be read like: \n"You paid {val[4]} to {val[2]} for {val[3]} on {val[1]} via {val[5]} in category {val[6]}"'
    mb.showinfo('Here\'s how to read your expense', msg)


def expenseToWordsBeforeAdding():
    global dateField, description, amount, payee, modeOfPayment, category

    msg = (f'Your expense can be read like: \n"You paid {amount.get()} to {payee.get()} '
           f'for {description.get()} on '
           f'{dateField.get()} via {modeOfPayment.get()} in category {category.get()}"')
    mb.showinfo('Here\'s how to read your expense', msg)


def exportExpenses():
    global dbconnector
    all_data = dbconnector.execute('SELECT * FROM ExpenseTracker')
    data = all_data.fetchall()

    df = pd.DataFrame(data, columns=['ID', 'Date', 'Payee', 'Description', 'Amount', 'ModeOfPayment', 'Category'])
    
    file_path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel files', '*.xlsx')])
    
    if file_path:
        try:
            df.to_excel(file_path, index=False)
            mb.showinfo('Success!', f'Your expense data has been exported to {file_path}')
        except Exception as e:
            mb.showerror('Export Failed', f'Could not export to Excel: {str(e)}')


def displayGraph():
    global dbconnector
    all_data = dbconnector.execute('SELECT * FROM ExpenseTracker')
    data = all_data.fetchall()

    graphWindow = Toplevel()
    graphWindow.title("Amount Spent Graph")

    fig, ax = plt.subplots(figsize=(10, 6))

    if graphOption.get() == "Total Amount Spent per Mode of Payment":
        modeOfPayment_amount = {}
        for row in data:
            if row[5] not in modeOfPayment_amount:
                modeOfPayment_amount[row[5]] = 0
            modeOfPayment_amount[row[5]] += row[4]
        labels = modeOfPayment_amount.keys()
        values = modeOfPayment_amount.values()
        ax.bar(labels, values, color='skyblue')
        ax.set_xlabel('Mode of Payment')
        ax.set_ylabel('Total Amount Spent')

    elif graphOption.get() == "Total Amount Spent per Payee":
        payee_amount = {}
        for row in data:
            if row[2] not in payee_amount:
                payee_amount[row[2]] = 0
            payee_amount[row[2]] += row[4]
        labels = payee_amount.keys()
        values = payee_amount.values()
        ax.bar(labels, values, color='lightgreen')
        ax.set_xlabel('Payee')
        ax.set_ylabel('Total Amount Spent')

    elif graphOption.get() == "Total Amount Spent per Month":
        month_amount = {}
        for row in data:
            month = row[1].split('-')[1]
            if month not in month_amount:
                month_amount[month] = 0
            month_amount[month] += row[4]
        labels = month_amount.keys()
        values = month_amount.values()
        ax.bar(labels, values, color='salmon')
        ax.set_xlabel('Month')
        ax.set_ylabel('Total Amount Spent')
    
    elif graphOption.get() == "Total Amount Spent per Category":
        category_amount = {}
        for row in data:
            if row[6] not in category_amount:
                category_amount[row[6]] = 0
            category_amount[row[6]] += row[4]
        labels = category_amount.keys()
        values = category_amount.values()
        ax.bar(labels, values, color='purple')
        ax.set_xlabel('Category')
        ax.set_ylabel('Total Amount Spent')

    ax.set_title('Amount Spent Graph')
    plt.xticks(rotation=45)
    plt.tight_layout()

    canvas = FigureCanvasTkAgg(fig, master=graphWindow)
    canvas.draw()
    canvas.get_tk_widget().pack()

    Button(graphWindow, text="Save Graph", font=("Bahnschrift Condensed", "13"), width=30, bg="#90EE90",
           fg="#000000", relief=GROOVE, activebackground="#008000", activeforeground="#FF0000",
           command=lambda: saveGraph(fig)).pack(pady=10)


def saveGraph(fig):
    file_path = filedialog.asksaveasfilename(defaultextension='.png', filetypes=[('PNG files', '*.png')])
    if file_path:
        fig.savefig(file_path)
        mb.showinfo('Graph Saved', f'The graph has been saved to {file_path}')


def searchExpenses():
    keyword = searchEntry.get()
    query = ("SELECT * FROM ExpenseTracker WHERE Date LIKE ? OR Payee LIKE ? OR Description LIKE ? OR Amount LIKE ? OR "
             "ModeOfPayment LIKE ? OR Category LIKE ?")
    data = dbconnector.execute(query, (
        '%' + keyword + '%', '%' + keyword + '%', '%' + keyword + '%', '%' + keyword + '%', '%' + keyword + '%', '%' + keyword + '%'))
    data_table.delete(*data_table.get_children())
    for val in data:
        data_table.insert('', END, values=val)


def showTotalExpense():
    global dbconnector
    total = dbconnector.execute('SELECT SUM(Amount) FROM ExpenseTracker').fetchone()[0]
    total = total if total else 0
    mb.showinfo('Total Expense', f'The total expense is: {total}')


def showMonthlyExpense():
    global dbconnector
    current_month = datetime.datetime.now().strftime("%Y-%m")
    total = dbconnector.execute('SELECT SUM(Amount) FROM ExpenseTracker WHERE Date LIKE ?', (f'{current_month}%',)).fetchone()[0]
    total = total if total else 0
    mb.showinfo('Monthly Expense', f'The total expense for this month is: {total}')


def showYearlyExpense():
    global dbconnector
    current_year = datetime.datetime.now().strftime("%Y")
    total = dbconnector.execute('SELECT SUM(Amount) FROM ExpenseTracker WHERE Date LIKE ?', (f'{current_year}%',)).fetchone()[0]
    total = total if total else 0
    mb.showinfo('Yearly Expense', f'The total expense for this year is: {total}')


mainWindow = Tk()
mainWindow.geometry("1920x1080")
mainWindow.title("Expense Tracker")

style = ttk.Style()
style.configure('TButton', font=('Bahnschrift Condensed', 13), background='#FFA07A', foreground='#FFFFFF',
                relief=GROOVE, padding=5)
style.configure('TLabel', font=('Bahnschrift Condensed', 15), foreground='#333333')
style.configure('Treeview.Heading', font=('Bahnschrift Condensed', 15), background='#FFA07A', foreground='#000000')
style.configure('Treeview', font=('Bahnschrift Condensed', 13), rowheight=25)

titleLabel = Label(mainWindow, text="Expense Tracker", font=("Bahnschrift Condensed", 24), bg="#FFA07A", fg="#FFFFFF")
titleLabel.pack(pady=10, fill=X)

searchFrame = Frame(mainWindow, bg="#F5F5F5")
searchFrame.pack(fill=X)

searchLabel = Label(searchFrame, text="Search Expense:", font=("Bahnschrift Condensed", "13"), bg="#F5F5F5", fg="#333333")
searchLabel.pack(side=LEFT, padx=10, pady=10)

searchEntry = Entry(searchFrame, font=("Bahnschrift Condensed", "13"), bg="#FFFFFF", fg="#333333")
searchEntry.pack(side=LEFT, padx=10, pady=10)

searchButton = Button(searchFrame, text="Search", font=("Bahnschrift Condensed", "13"), width=10, bg="#FFA07A",
                      fg="#FFFFFF", relief=GROOVE, activebackground="#FF7F50", activeforeground="#FFFFFF",
                      command=searchExpenses)
searchButton.pack(side=LEFT, padx=10, pady=10)

mainFrame = Frame(mainWindow, bg="#FFFFFF")
mainFrame.pack(fill=BOTH, expand=True)

root_dir = os.path.dirname(os.path.abspath(__file__))
db_path = os.path.join(root_dir, 'ExpenseTracker.db')

dbconnector = sqlite3.connect(db_path, detect_types=sqlite3.PARSE_DECLTYPES)
dbconnector.execute('''CREATE TABLE IF NOT EXISTS ExpenseTracker (
                       ID INTEGER PRIMARY KEY AUTOINCREMENT,
                       Date TEXT NOT NULL,
                       Payee TEXT NOT NULL,
                       Description TEXT NOT NULL,
                       Amount REAL NOT NULL,
                       ModeOfPayment TEXT NOT NULL,
                       Category TEXT NOT NULL
                   )''')
dbconnector.commit()

data_table = ttk.Treeview(mainFrame, columns=('ID', 'Date', 'Payee', 'Description', 'Amount', 'ModeOfPayment', 'Category'),
                          show='headings', style='Treeview')
for col in ('ID', 'Date', 'Payee', 'Description', 'Amount', 'ModeOfPayment', 'Category'):
    data_table.heading(col, text=col)
data_table.pack(fill=BOTH, expand=True)

frameL1 = Frame(mainWindow, padx=10, pady=10, bg="#F5F5F5")
frameL1.pack(side=LEFT, anchor=NW)

Label(frameL1, text="Date:", font=("Bahnschrift Condensed", "15"), bg="#F5F5F5", fg="#000000").grid(row=1, column=0, sticky=W, padx=50, pady=10)
dateField = DateEntry(frameL1, width=20, background='#4CAF50', foreground='black', borderwidth=2)
dateField.grid(row=1, column=1, sticky=W, padx=10, pady=10)

Label(frameL1, text="Payee:", font=("Bahnschrift Condensed", "15"), bg="#F5F5F5", fg="#333333").grid(row=2, column=0, sticky=W, padx=50, pady=10)
payee = StringVar()
Entry(frameL1, textvariable=payee, font=("Bahnschrift Condensed", "15"), bg="#FFFFFF", fg="#333333").grid(row=2, column=1, sticky=W, padx=10,
                                                                              pady=10)

Label(frameL1, text="Description:", font=("Bahnschrift Condensed", "15"), bg="#F5F5F5", fg="#333333").grid(row=3, column=0, sticky=W, padx=50,
                                                                               pady=10)
description = StringVar()
Entry(frameL1, textvariable=description, font=("Bahnschrift Condensed", "15"), bg="#FFFFFF", fg="#333333").grid(row=3, column=1, sticky=W, padx=10,
                                                                                    pady=10)

Label(frameL1, text="Amount:", font=("Bahnschrift Condensed", "15"), bg="#F5F5F5", fg="#333333").grid(row=4, column=0, sticky=W, padx=50, pady=10)
amount = DoubleVar()
Entry(frameL1, textvariable=amount, font=("Bahnschrift Condensed", "15"), bg="#FFFFFF", fg="#333333").grid(row=4, column=1, sticky=W, padx=10,
                                                                               pady=10)

Label(frameL1, text="Mode of Payment:", font=("Bahnschrift Condensed", "15"), bg="#F5F5F5", fg="#333333").grid(row=5, column=0, sticky=W, padx=50,
                                                                                   pady=10)
modeOfPayment = StringVar()
paymentOptions = ['Cash', 'Credit Card', 'Debit Card', 'Net Banking', 'UPI', 'Others']
modeOfPayment.set('Cash')
option_menu = OptionMenu(frameL1, modeOfPayment, *paymentOptions)
option_menu.grid(row=5, column=1, sticky=W, padx=10, pady=10)
option_menu.config(width=18)

Label(frameL1, text="Category:", font=("Bahnschrift Condensed", "15"), bg="#F5F5F5", fg="#333333").grid(row=6, column=0, sticky=W, padx=50, pady=10)
category = StringVar()
categoryOptions = ['Food', 'Groceries', 'Bills', 'Transportation', 'Entertainment', 'Education', 'Health', 'Shopping', 'Housing', 'Others']
category.set('Food')
category_menu = OptionMenu(frameL1, category, *categoryOptions)
category_menu.grid(row=6, column=1, sticky=W, padx=10, pady=10)
category_menu.config(width=18)

Button(frameL1, text="Clear Fields", font=("Bahnschrift Condensed", "13"), width=20, bg="#FFA07A", fg="#FFFFFF",
       relief=GROOVE, activebackground="#FF7F50", activeforeground="#FFFFFF", command=clearFields).grid(row=7,
                                                                                                        column=0,
                                                                                                        sticky=W,
                                                                                                        padx=50,
                                                                                                        pady=10)

Button(frameL1, text="Add", font=("Bahnschrift Condensed", "13"), width=20, bg="#FFA07A", fg="#FFFFFF", relief=GROOVE,
       activebackground="#FF7F50", activeforeground="#FFFFFF", command=addAnotherExpense).grid(row=7, column=1,
                                                                                               sticky=W, padx=10,
                                                                                               pady=10)
frameL2 = Frame(mainWindow, padx=10, pady=10, bg="#F5F5F5")
frameL2.pack(side=LEFT, anchor=NW)

listAllExpenses()

Button(frameL2, text="View Expense Info", font=("Bahnschrift Condensed", "13"), width=30, bg="#FFA07A", fg="#FFFFFF",
       relief=GROOVE, activebackground="#FF7F50", activeforeground="#FFFFFF", command=viewExpenseInfo).grid(row=1,
                                                                                                            column=0,
                                                                                                            sticky=W,
                                                                                                            padx=50,
                                                                                                            pady=10)

Button(frameL2, text="Edit Expense", font=("Bahnschrift Condensed", "13"), width=30, bg="#FFA07A", fg="#FFFFFF",
       relief=GROOVE, activebackground="#FF7F50", activeforeground="#FFFFFF", command=editExpense).grid(row=2,
                                                                                                        column=0,
                                                                                                        sticky=W,
                                                                                                        padx=50,
                                                                                                        pady=10)

Button(frameL2, text="Delete Expense", font=("Bahnschrift Condensed", "13"), width=30, bg="#FFA07A", fg="#FFFFFF",
       relief=GROOVE, activebackground="#FF7F50", activeforeground="#FFFFFF", command=removeExpense).grid(row=3,
                                                                                                          column=0,
                                                                                                          sticky=W,
                                                                                                          padx=50,
                                                                                                          pady=10)

Button(frameL2, text="Delete All Expenses", font=("Bahnschrift Condensed", "13"), width=30, bg="#FFA07A", fg="#FFFFFF",
       relief=GROOVE, activebackground="#FF7F50", activeforeground="#FFFFFF", command=removeAllExpenses).grid(row=4,
                                                                                                              column=0,
                                                                                                              sticky=W,
                                                                                                              padx=50,
                                                                                                              pady=10)

Button(frameL2, text="Read Selected Expense", font=("Bahnschrift Condensed", "13"), width=30, bg="#FFA07A",
       fg="#FFFFFF", relief=GROOVE, activebackground="#FF7F50", activeforeground="#FFFFFF",
       command=selectedExpenseToWords).grid(row=5, column=0, sticky=W, padx=50, pady=10)

Button(frameL2, text="Read Expense before Adding", font=("Bahnschrift Condensed", "13"), width=30, bg="#FFA07A",
       fg="#FFFFFF", relief=GROOVE, activebackground="#FF7F50", activeforeground="#FFFFFF",
       command=expenseToWordsBeforeAdding).grid(row=6, column=0, sticky=W, padx=50, pady=10)

frameL3 = Frame(mainWindow, padx=10, pady=10, bg="#F5F5F5")
frameL3.pack(side=LEFT, anchor=NW)

Button(frameL3, text="Export to Excel", font=("Bahnschrift Condensed", "13"), width=30, bg="#FFA07A", fg="#FFFFFF",
       relief=GROOVE, activebackground="#FF7F50", activeforeground="#FFFFFF", command=exportExpenses).grid(row=1,
                                                                                                           column=0,
                                                                                                           sticky=W,
                                                                                                           padx=50,
                                                                                                           pady=10)

graphFrame = Frame(frameL3, bg="#F5F5F5")
graphFrame.grid(row=2, column=0, sticky=W, padx=50, pady=10)

Label(graphFrame, text="Select Graph Type:", font=("Bahnschrift Condensed", "13"), bg="#F5F5F5", fg="#333333").pack(side=LEFT, padx=5)

graphOption = StringVar()
graphOption.set("Total Amount Spent per Mode of Payment")
graph_option_menu = OptionMenu(graphFrame, graphOption,
                               "Total Amount Spent per Mode of Payment",
                               "Total Amount Spent per Payee",
                               "Total Amount Spent per Month",
                               "Total Amount Spent per Category")
graph_option_menu.pack(side=LEFT)
graph_option_menu.config(width=25)

Button(frameL3, text="View Graph", font=("Bahnschrift Condensed", "13"), width=30, bg="#FFA07A", fg="#FFFFFF",
       relief=GROOVE, activebackground="#FF7F50", activeforeground="#FFFFFF", command=displayGraph).grid(row=3,
                                                                                                         column=0,
                                                                                                         sticky=W,
                                                                                                         padx=50,
                                                                                                         pady=10)

Button(frameL3, text="Show Total Expense", font=("Bahnschrift Condensed", "13"), width=30, bg="#FFA07A", fg="#FFFFFF",
       relief=GROOVE, activebackground="#FF7F50", activeforeground="#FFFFFF", command=showTotalExpense).grid(row=4,
                                                                                                             column=0,
                                                                                                             sticky=W,
                                                                                                             padx=50,
                                                                                                             pady=10)

Button(frameL3, text="Show Monthly Expense", font=("Bahnschrift Condensed", "13"), width=30, bg="#FFA07A", fg="#FFFFFF",
       relief=GROOVE, activebackground="#FF7F50", activeforeground="#FFFFFF", command=showMonthlyExpense).grid(row=5,
                                                                                                              column=0,
                                                                                                              sticky=W,
                                                                                                              padx=50,
                                                                                                              pady=10)

Button(frameL3, text="Show Yearly Expense", font=("Bahnschrift Condensed", "13"), width=30, bg="#FFA07A", fg="#FFFFFF",
       relief=GROOVE, activebackground="#FF7F50", activeforeground="#FFFFFF", command=showYearlyExpense).grid(row=6,
                                                                                                             column=0,
                                                                                                             sticky=W,
                                                                                                             padx=50,
                                                                                                             pady=10)

Button(frameL3, text="Exit", font=("Bahnschrift Condensed", "13"), width=30, bg="#FFA07A", fg="#FFFFFF", relief=GROOVE,
       activebackground="#FF7F50", activeforeground="#FFFFFF", command=mainWindow.quit).grid(row=7,
                                                                                             column=0,
                                                                                             sticky=W,
                                                                                             padx=50,
                                                                                             pady=10)

mainWindow.configure(bg="#F5F5F5")
mainWindow.mainloop()
