import os
from openpyxl import Workbook, load_workbook

file_path = "/projects/records.xlsx"

sno = os.environ.get("S_No", "")
name = os.environ.get("Name", "")
emp_id = os.environ.get("EmpID", "")

try:
    wb = load_workbook(file_path)
    ws = wb.active
except FileNotFoundError:
    wb = Workbook()
    ws = wb.active
    ws.append(["S_No", "Name", "EmpID"])

ws.append([sno, name, emp_id])
wb.save(file_path)
print(f"Added {sno} - {name} ({emp_id}) to {file_path}")
