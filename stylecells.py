import openpyxl
from openpyxl.styles import Font, Color, Alignment, Border, Side, colors

work_book = openpyxl.load_workbook("student_data.xlsx")
print(work_book.sheetnames)

sheet = work_book["Sheet"]
bold_font = Font(bold=True)
big_red_text = Font(color="#FF0000", size=20)

center_aligned_text = Alignment(horizontal="center")