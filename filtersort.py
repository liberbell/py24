import openpyxl

work_book = openpyxl.Workbook()
work_sheet = work_book.active

data = [["Champion", "Year"],
        ["France", 2018],
        ["Spain", 2010],
        ["Italy", 2006]]