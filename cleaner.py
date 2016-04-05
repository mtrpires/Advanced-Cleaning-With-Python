import re
import csv
from xlrd import open_workbook
from os import listdir


# Gets the date from the specific cell in the documents
# and use regular expression to separate month from year.
# returns a dictionary with YEAR and MONTH
def getDate(worksheet):
    '''
    :param worksheet:
    :return: dictionary
    '''
    date = worksheet.cell_value(3, 0)
    match = re.match(r'NYS[ ,]+(\w+)[ ,](\d+)', date).groups()
    date = {'YEAR': match[1],
            'MONTH': match[0]
            }
    return date

# Empty dictionary where we will store each row, before parsing it to the CSV
row = {}

# This is where the xls files are
basedir = 'data/'

# This is an os function that returns a list of filenames in a folder
files = listdir('data')

# Empty list to store only XLS files found in the folder
sheets = []
[sheets.append(filename) for filename in files if filename.endswith("xls")]

header_is_written = False
# Iterating over the files in folder
for filename in sheets:
    print('Parsing {0}{1}\r'.format(basedir, filename)),
    # Opens the xls file
    worksheet = open_workbook(basedir + filename).sheet_by_index(0)
    # Get the date from it
    date = getDate(worksheet)
    row['YEAR'] = date['YEAR']
    row['MONTH'] = date['MONTH']
    # We're going to iterate over all the values in column[1]
    col1_values = worksheet.col_values(1)
    for i in range(len(col1_values)):
        if worksheet.cell_value(i, 1) == u'TOTALS:':
            row['COUNTY'] = worksheet.cell_value(i, 0).strip()
            county = True
            counter = 1
            # While we're in that county, let's save all the data it has for different plan names
            while county is True:
                row['PLAN NAME'] = worksheet.cell_value(i + counter, 1).strip()
                values = worksheet.row_slice(i + counter, start_colx=2, end_colx=7)
                if worksheet.cell_value(4,5) == u'NYSoH':
                    row['ADC'] = int(values[0].value)
                    row['SNA'] = int(values[1].value)
                    row['SSI'] = int(values[2].value)
                    row['NYSoH'] = int(values[3].value)
                    row['TOTAL'] = int(values[4].value)
                else:
                    row['ADC'] = int(values[0].value)
                    row['SNA'] = int(values[1].value)
                    row['SSI'] = int(values[3].value)
                    row['NYSoH'] = ""
                    row['TOTAL'] = int(values[4].value)
                counter += 1
                # If the robot finds another cell with the value "TOTALS:", it means we reached another county
                # Time to break and start again
                if worksheet.cell_value(i + counter, 1) == u'TOTALS:':
                    county = False
                # If we reach a point in this column that the value is blank, it means our search is over!
                elif row['PLAN NAME'] == '':
                    break
                # Saving the row in the CSV file...
                with open('table.csv', 'ab') as f:
                    w = csv.DictWriter(f, row.keys())
                    if header_is_written is False:
                        w.writeheader()
                        header_is_written = True
                    w.writerow(row)
