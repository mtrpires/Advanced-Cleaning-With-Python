from xlrd import open_workbook
import csv

worksheet = open_workbook('../data/en01_13.xls').sheet_by_index(0)

row = {}

header_is_written = False
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

            with open('csv/table.csv', 'a') as f:
                w = csv.DictWriter(f, row.keys())
                if header_is_written is False:
                    w.writeheader()
                    header_is_written = True
                w.writerow(row)
