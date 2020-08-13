import openpyxl

work_book = openpyxl.load_workbook("products.xlsx")
sheet = work_book["Products"]

fx_range = work_book.defined_names["fx_rates"]
print(fx_range)
print(fx_range.destinations)