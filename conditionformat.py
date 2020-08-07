import openpyxl
from openpyxl.formatting.rule import IconSetRule

# work_book = openpyxl.load_workbook("zomato_reviews.xlsx")
work_book = openpyxl.load_workbook("zomato-reviews.xlsx")

sheet = work_book.active

icon_set_rule = IconSetRule(icon_style="4Arrows", type="num", values=[1, 2, 3, 4])
print(sheet.max_row)

sheet.conditional_formatting.add("G2:G9558", icon_set_rule)
work_book.save("zomato_iconset.xlsx")