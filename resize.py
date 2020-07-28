import openpyxl

work_book = openpyxl.Workbook()
sheet_obj = work_book.active

print(sheet_obj)
sheet_obj.title = "FirstSheet"
print(sheet_obj)