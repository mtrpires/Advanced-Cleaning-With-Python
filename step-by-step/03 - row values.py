from xlrd import open_workbook

# Imports the first sheet of the file
sheet = open_workbook('../data/en01_13.xls').sheet_by_index(0)

# Method .row_slice returns a list of values in a row
for cell in sheet.row_slice(4, start_colx=2, end_colx=7):
    print cell.value
