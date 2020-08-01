import openpyxl
from openpyxl.styles import Font, Color, Alignment, Border, Side, colors

work_book = openpyxl.load_workbook("student_data.xlsx")
print(work_book.sheetnames)

sheet = work_book["Sheet"]
