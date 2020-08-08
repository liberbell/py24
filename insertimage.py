from openpyxl import Workbook
from openpyxl.drawing.image import Image

Work_book = Workbook()
work_sheet = Work_book.active

img = Image("small-image.jpg")

work_sheet.add_image(img, "C11")
# Work_book.save("images.xlsx")

print(img.width, img.height)

large_img = Image("small-image.jpg")
large_img.width = 600
large_img.height = 400
# print(img.width, img.height)

work_sheet.add_image(large_img, "L11")
Work_book.save("images.xlsx")