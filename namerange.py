import openpyxl

work_book = openpyxl.load_workbook("products.xlsx")
sheet = work_book["Products"]