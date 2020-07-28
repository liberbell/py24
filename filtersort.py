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
        ["Argentina", 1978],
        ["Germany", 1974],
        ["Brazil", 1970],
        ["England", 1976],
        ["Brazil", 1962],
        ["Brazil", 1958],
        ["Germany", 2014],
        ["Germany", 1954],
        ["Uruguay", 1950],
        ["Italy", 1938],
        ["Italy", 1934],
        ["Uruguay", 1930],
        ["Germany", 1990],
        ["Brazil", 2002]]

for r in data:
    work_sheet.append(r)

# work_book.save("world_cup_winners.xlsx")

print(work_sheet.calculate_dimension())

work_sheet.auto_filter.fef = work_sheet.calculate_dimension()
work_sheet.auto_filter.add_filter_column(0, ["Brazil", "Italy", "Argentina"])

# work_book.save("world_cup_winners.xlsx")

print(work_sheet["B"][1], work_sheet["B"][1])

range_str = work_sheet["B"][1].coordinate + ":" + work_sheet["B"][-1].coordinate
print(range_str)

work_sheet.auto_filter.add_sort_condition(range_str, descending=True)
work_book.save("world_cup_winners.xlsx")