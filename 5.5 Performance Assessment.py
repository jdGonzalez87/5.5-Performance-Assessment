import csv
import datetime
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.chart import PieChart, Reference

# Absolute path to your folder
FILE_PATH = r"C:\FinalExam\final.csv"
EXCEL_PATH = r"C:\FinalExam\final.xlsx"

STUDENT_ID = "jasdan3949" 
TODAY = datetime.date.today().strftime("%B %d, %Y")


def askUser():
    """
    Loop logic:
    This loop runs exactly 5 times, each time asking the user for a number.
    Each number is added to a running total, which is printed at the end.
    """
    total = 0
    for i in range(5):
        num = int(input("Please enter a number: "))
        total += num
    print(f"The sum for the 5 numbers entered is: {total}")


def askIncome():
    """
    Loop logic:
    This loop asks for 5 names and incomes.
    Each entry is appended to final.csv using newline characters.
    """
    with open(FILE_PATH, "a", newline="") as file:
        writer = csv.writer(file)

        for i in range(5):
            name = input("Please enter a name: ")
            income = int(input("Please enter their income: "))
            writer.writerow([name, income])


def excelPie():
    """
    Creates an Excel file with a pie chart of the CSV data.
    Each line below is commented to explain its purpose.
    """

    # Read CSV data
    names = []
    incomes = []
    with open(FILE_PATH, "r") as file:
        reader = csv.reader(file)
        for row in reader:
            names.append(row[0])
            incomes.append(int(row[1]))  # Cast to int so the chart works

    # Create a new Excel workbook
    wb = Workbook()
    ws = wb.active

    # Write data into Excel sheet
    ws.append(["Name", "Income"])
    for i in range(len(names)):
        ws.append([names[i], incomes[i]])

    # Create a pie chart object
    pie = PieChart()

    # Select the income column for the chart
    data = Reference(ws, min_col=2, min_row=1, max_row=len(incomes) + 1)

    # Select the names column for labels
    labels = Reference(ws, min_col=1, min_row=2, max_row=len(names) + 1)

    # Add data and labels to the chart
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(labels)

    # Set the chart title
    pie.title = f"{STUDENT_ID} {TODAY}"

    # Add chart to the sheet
    ws.add_chart(pie, "E5")

    # Save Excel file
    wb.save(EXCEL_PATH)


def verticalBar():
    """
    Creates a vertical bar graph using matplotlib.
    """

    names = []
    incomes = []

    with open(FILE_PATH, "r") as file:
        reader = csv.reader(file)
        for row in reader:
            names.append(row[0])
            incomes.append(int(row[1]))

    plt.figure(figsize=(10, 6))
    plt.bar(names, incomes, color="green")
    plt.title(f"{STUDENT_ID} {TODAY}")
    plt.xlabel("Name")
    plt.ylabel("Income")
    plt.tight_layout()
    plt.show()


# -------------------------
# MAIN PROGRAM EXECUTION
# -------------------------

askUser()
askIncome()
excelPie()
verticalBar()
