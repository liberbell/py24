import openpyxl

work_book = openpyxl.load_workbook("sales_record.xlsx")
sheet = work_book.active

print(sheet.max_row)