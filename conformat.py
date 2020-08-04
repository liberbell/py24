import openpyxl
from openpyxl.styles import PatternFill, colors
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import Rule

work_book = openpyxl.load_workbook("sales_record.xlsx")
sheet = work_book.active

print(sheet.max_row)
print(sheet.max_column)

for row in sheet["L2:N101"]:
    for cell in row:
        cell.number_format = "#,##0"

# work_book.save("sales_basic_conditional.xlsx")

yellow_background = PatternFill(bgColor="00FFFF00")