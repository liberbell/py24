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
print(sheet.calculate_dimension())

# work_book.save("brand_new_workbook.xlsx")

sheet.insert_rows(idx=2, amount=3)
sheet.insert_cols(idx=3)
print(sheet.calculate_dimension())

# work_book.save("brand_new_workbook.xlsx")

sheet.delete_rows(idx=2, amount=3)
sheet.delete_cols(3)
print(sheet.calculate_dimension())
# work_book.save("brand_new_workbook.xlsx")

sheet.title = "FirstSheet"
# work_book.save("brand_new_workbook.xlsx")

data = [["Planet", "Radius(km)", "Distance from the Sun (M km)"],
        ["Earth", 6371, 150],
        ["Mars", 2289, 228],
        ["Mercury", 2440, 58]]