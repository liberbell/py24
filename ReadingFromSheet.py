import openpyxl

work_book = openpyxl.load_workbook("company_revenue.xlsx")
print(work_book.sheetnames)

sheet_obj = work_book["Revenue"]
print(sheet_obj.title)
print(sheet_obj["A1"])
print(sheet_obj["A1"].value)

cell = sheet_obj["B1"]
print(type(cell))
print(dir(cell))

print(cell.row)
print(cell.column)

print(cell.number_format)
print(cell.coordinate)
print(cell.data_type)