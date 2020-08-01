import openpyxl

work_book = openpyxl.load_workbook("student_data.xlsx")
print(work_book.sheetnames)