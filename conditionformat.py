import openpyxl
from openpyxl.formatting.rule import IconSetRule

# work_book = openpyxl.load_workbook("zomato_reviews.xlsx")
work_book = openpyxl.load_workbook("zomato-reviews.xlsx")

sheet = work_book.active