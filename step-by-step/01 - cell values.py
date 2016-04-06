from xlrd import open_workbook

# Imports the first sheet of the file
sheet = open_workbook('../data/en01_13.xls').sheet_by_index(0)

# Prints row 4, column 1 (A4)
print sheet.cell_value(3,0)
