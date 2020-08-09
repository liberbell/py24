from openpyxl.utils import FORMULAE
import openpyxl

print(FORMULAE)

work_book = openpyxl.Workbook()
sheet = work_book.active

sheet["A1"] = 21
sheet["A2"] = 11
sheet["A3"] = 7
sheet["A4"] = 9
sheet["A5"] = 6

sheet["C2"] = "SUM:"
sheet["D2"] = "=SUM(A1:A5)"

# work_book.save("formurae.xlsx")

sheet["C3"] = "PRODUCT:"
sheet["D3"] = "=PRODUCT(A1:A5)"
# work_book.save("formurae.xlsx")

sheet["C4"] = "COUNT:"
sheet["D4"] = "=COUNT(A1:A9)"

sheet["C5"] = "MEAN:"
sheet["D5"] = "=AVERAGE(A1:A5)"
# work_book.save("formurae.xlsx")