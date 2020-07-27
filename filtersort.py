import openpyxl

work_book = openpyxl.Workbook()
work_sheet = work_book.active

data = [["Champion", "Year"],
        ["France", 2018],
        ["Spain", 2010],
        ["Italy", 2006],
        ["France", 1998],
        ["Brazil", 1994],
        ["Argentina", 1986],
        ["Italy", 1982],
        ["Argentina", 1978]]