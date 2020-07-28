import openpyxl

work_book = openpyxl.Workbook()
sheet_obj = work_book.active

print(sheet_obj)
sheet_obj.title = "FirstSheet"
print(sheet_obj)

sheet_obj["C1"] = "A high row"
sheet_obj["D4"] = "A wide column"

sheet_obj.row_dimensions[1].height = 70