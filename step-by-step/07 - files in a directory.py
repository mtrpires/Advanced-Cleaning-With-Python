from os import listdir
from pprint import pprint

# List of files in a dir
files = listdir('../data')

sheets_expanded = []
for filename in files:
    if filename.endswith("xls"):
        sheets_expanded.append(filename)

pprint("This is the first list:")
pprint(sheets_expanded)
print

# Empty list to store only XLS files found in the folder
sheets = [filename for filename in files if filename.endswith("xls")]
pprint("This is the same list, but with less steps:")
pprint(sheets)
print
