import openpyxl

work_book = openpyxl.load_workbook("company_revenue.xlsx")
print(work_book.sheetnames)

sheet_obj = work_book["Revenue"]