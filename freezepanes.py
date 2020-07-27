import openpyxl

work_book = openpyxl.load_workbook("zomato-reviews.xlsx")
sheet = work_book.active

sheet.freeze_panes = "A2"
# work_book.save("zomato-reviews.xlsx")

sheet.freeze_panes = "B1"
# work_book.save("zomato-reviews.xlsx")

sheet.freeze_panes = "C2"
work_book.save("zomato-reviews.xlsx")