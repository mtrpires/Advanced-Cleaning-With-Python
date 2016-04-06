from xlrd import open_workbook
from pprint import pprint

# Imports the first sheet of the file
sheet = open_workbook('../data/en01_13.xls').sheet_by_index(0)

# Method .col_values returns a list of column values
pprint(sheet.col_values(1))
