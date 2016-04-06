from xlrd import open_workbook

# Imports the first sheet of the file
sheet = open_workbook('../data/en01_13.xls').sheet_by_index(0)

row = {}

# Method .row_slice returns a list of values in a row
row_values = sheet.row_slice(6, start_colx=2, end_colx=7)

row['ADC'] = int(row_values[0].value)
row['SNA'] = int(row_values[1].value)
row['SSI'] = int(row_values[3].value)
row['NYSoH'] = ''
row['TOTAL'] = int(row_values[4].value)

print row
