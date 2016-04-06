from xlrd import open_workbook

# Imports the first sheet of the file
sheet = open_workbook('../data/en01_13.xls').sheet_by_index(0)

# Method .col_values returns a list of column values
for cell in sheet.col_values(1):
    print cell
