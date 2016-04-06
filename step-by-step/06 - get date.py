from xlrd import open_workbook
import re


# Imports the first sheet of the file
sheet = open_workbook('../data/en01_13.xls').sheet_by_index(0)

# Gets cell with the date
date = sheet.cell_value(3, 0)

# NYS[ ,]+ -> NYS followed by either " " or ","
# (\w+)[ ,] -> capture[0]: any word followed by " " or ","
# (\d+) -> capture[1]: any number of digits
match = re.match(r'NYS[ ,]+(\w+)[ ,](\d+)', date).groups()

date = {'YEAR': match[1],
        'MONTH': match[0]
        }
print("This is the date dictionary:")
print date
print

# Gets the date from the specific cell in the documents
# and use regular expression to separate month from year.
# returns a dictionary with YEAR and MONTH
def getDate(worksheet):
    '''
    sheet: open_workbook.sheet_by_index() object
    return: dictionary with YEAR and MONTH
    '''
    date = worksheet.cell_value(3, 0)
    match = re.match(r'NYS[ ,]+(\w+)[ ,](\d+)', date).groups()
    date = {'YEAR': match[1],
            'MONTH': match[0]
            }
    return date

print("This is the date dictionary using the function getDate()")
print getDate(sheet)
print
