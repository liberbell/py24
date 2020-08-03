import openpyxl

work_book = openpyxl.load_workbook("sales_record.xlsx")
sheet = work_book.active

print(sheet.max_row)
print(sheet.max_column)

for row in sheet["K2:M101"]:
    for cell in row:
        cell.number_format = "#,##0"