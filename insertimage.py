from openpyxl import Workbook
from openpyxl.drawing.image import Image

Work_book = Workbook()
work_sheet = Work_book.active

img = Image("small-image.jpg")