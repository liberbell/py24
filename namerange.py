import openpyxl

work_book = openpyxl.load_workbook("products.xlsx")
sheet = work_book["Products"]

fx_range = work_book.defined_names["fx_rates"]
print(fx_range)
print(fx_range.destinations)

cells = []

for title, coord in fx_range.destinations:
    ws = work_book[title]
    cells.append(ws[coord])

print(cells)

max_row_str = str(sheet.max_row)
print(max_row_str)