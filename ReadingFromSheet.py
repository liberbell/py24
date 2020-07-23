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

print(cell.value)

print(sheet_obj["A2"].value + ", based in " + sheet_obj["B2"].value + " has a revenue of $" + str (sheet_obj["C2"].value) + " billion.")
print(sheet_obj.cell(row=1, column=2))
print(sheet_obj.cell(row=1, column=2).value)

print(sheet_obj.max_row)
print(sheet_obj.max_column)