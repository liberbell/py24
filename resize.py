import openpyxl

work_book = openpyxl.Workbook()
sheet_obj = work_book.active

print(sheet_obj)