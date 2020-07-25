import openpyxl

work_book = openpyxl.Workbook()
print(work_book.active)

sheet = work_book["Sheet"]
sheet["A1"] = "Hello"
sheet["B1"] = "Excel"
sheet["C1"] = "Users!"

print(sheet["A1"].value)
print(sheet["B1"].value)
print(sheet["C1"].value)

# work_book.save("brand_new_workbook.xlsx")

sheet["A1"] = "Goodbye"
sheet["B1"] = "Excel Users"
sheet["C1"] = "!"

print(sheet["A1"].value)
print(sheet["B1"].value)
print(sheet["C1"].value)

# work_book.save("brand_new_workbook.xlsx")

print(sheet.calculate_dimension())

sheet.append(["One", "row", "of", "text"])