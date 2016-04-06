from xlrd import open_workbook
import csv

# Imports the first sheet of the file
sheet = open_workbook('../data/en01_13.xls').sheet_by_index(0)

row = {}

# Method .row_slice returns a list of values in a row
values = sheet.row_slice(6, start_colx=2, end_colx=7)

row['ADC'] = int(values[0].value)
row['SNA'] = int(values[1].value)
row['SSI'] = int(values[3].value)
row['NYSoH'] = ''
row['TOTAL'] = int(values[4].value)

with open('csv/table.csv', 'w') as f:
    w = csv.DictWriter(f, row.keys())
    w.writeheader()
    w.writerow(row)
