import openpyxl
from datetime import datetime

def initialize_workbook():
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Assistance Data"
    sheet.append(["Date", "Unikey", "Counter Location", "Assistance Category", "Time Spent (minutes)", "Notes"])
    return workbook, sheet

def save_to_excel(sheet, unikey, counter_location, assistance_category, elapsed_time):
    date = datetime.now().strftime("%Y-%m-%d")
    notes = ""  # Add notes if needed
    sheet.append([date, unikey, counter_location, assistance_category, elapsed_time, notes])