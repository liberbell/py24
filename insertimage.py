from openpyxl import Workbook
from openpyxl.drawing.image import Image

Work_book = Workbook()
work_sheet = Work_book.active

img = Image("small-image.jpg")

work_sheet.add_image(img, "C11")
Work_book.save("images.xlsx")