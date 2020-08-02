import openpyxl
from openpyxl.styles import Font, Color, Alignment, Border, Side, colors
from openpyxl.styles import NamedStyle

work_book = openpyxl.load_workbook("student_data.xlsx")
print(work_book.sheetnames)

sheet = work_book["Sheet"]
bold_font = Font(bold=True)
big_red_text = Font(color="FFFF0000", size=20)

center_aligned_text = Alignment(horizontal="center")
doubule_border_side = Side(border_style="double")

square_border = Border(top=doubule_border_side, right=doubule_border_side, bottom=doubule_border_side, left=doubule_border_side)

sheet["B2"].font = bold_font
sheet["B3"].font = big_red_text
sheet["C4"].alignment = center_aligned_text
sheet["C5"].border = square_border

# work_book.save(filename="styled.xlsx")

sheet["B7"].alignment = center_aligned_text
sheet["B7"].font = big_red_text
sheet["B7"].border = square_border
# work_book.save(filename="styled.xlsx")

custom_style = NamedStyle(name="header")