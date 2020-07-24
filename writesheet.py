import openpyxl

work_book = openpyxl.Workbook()
print(work_book.active)

sheet = work_book["Sheet"]
sheet["A1"] = "Hello"
sheet["B1"] = "Excel"
sheet["C1"] = "Users!"