import openpyxl

work_book = openpyxl.Workbook()
print(work_book.active)

sheet = work_book["Sheet"]