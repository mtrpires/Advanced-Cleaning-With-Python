# Advanced Data Cleaning with Python
## International Journalism Festival - Perugia 2016
### School of Data (http://schoolofdata.org)
#### Sildes used in the presentations: [view](https://docs.google.com/presentation/d/10Vwk4hfgkmnAcgNlxmvxYT3LBvXb_j4N26LAe0R-w3Q/edit?usp=sharing)

Welcome! In this session we will show you how you can use Python (2.7) clean data! You will get information from a series of spreadsheets that have more or less the same structure and create a single clean CSV file. Pretty rad, huh?

This tutorial was inspired by a script I created for [J++ São Paulo](http://jmaismais.com) to clean some Brazil's government data and [an excellent OpenRefine workshop](https://github.com/sarahcnyt/data-journalism/tree/master/openrefine) that [Sarah Cohen](https://github.com/sarahcnyt) gave during the [Computer Assisted Reporting Conference 2016](http://www.ire.org/conferences/nicar2016/) (CAR Conference 2016) in Denver, USA.

I will assume:

- You know what Python, ```CSV``` and ```XLS``` files are;
- You know what [regular expressions](https://en.wikipedia.org/wiki/Regular_expression) are;
- You know what modules/libraries are and how to import them;
- You have some basic knowledge of how computer algorithms work (assign values to variables, create if/else tests, control flows - while/for - etc);
- You are not afraid of learning a bunch of cool tricks :)

Even if you don't know all of the above you should be able to follow!

## What will we be using?

- Python 2.7 (should work on 3.5 as well, though I haven't tested. Let me know how it goes!);
- Built-in Regular Expression (```re```) module to find specific text;
- Built-in ```CSV``` module to create and manipulate ```CSV``` files;
- ```listdir()``` function from built in ```os``` module, to get the list of file names in a folder;
- ```xlrd``` library to manipulate ```XLS``` files. You can install it using ```pip install xlrd``` on your terminal.

Let's go ahead and do our imports. Create a python file and include the following at the top:


```python
import re # Regular Expression library
import csv # The CSV library
from os import listdir # The listdir() function from the OS library
from xlrd import open_workbook # The open_workbook class from the xlrd library
```

Good. We should now be able to start cleaning, but a bit of context first.

## So what are we doing anyway?

I'll show you how to solve an annoying problem that might pop up anytime during your data journalism career: the data you want is spread across multiple files that **share the same structure between themselves**. This could be the case in a number of situations:

- A government agency publishes a report every month on their website with updated information;
- A company publishes, every quarter, reports about their financial status;
- A civil society organisation releases an ```.xls``` report every week about what's going on in the area they care about;
- Any situation where there are a number of files that preserve the same structure between themselves and you want to consolidate all the data together in a single ```CSV``` file for further analysis.

We will explore a specific case and learn the techniques to work around the issue. Keep in mind the specific case here is not important. You can extrapolate the approach presented here to apply anywhere else where the same conditions apply. Let's get started!

## Our data

The original data is here: https://www.health.ny.gov/health_care/managed_care/reports/enrollment/monthly/

The xls files in the ```data/``` folder are Medicaid long-term managed care reports from New York State in the United States. This data can be used as a way to determine which company would make a good subject based on its growth and size.

The published data is in the XLS format. While these reports are useful for human inspection, they can't really be processed by a computer as they are. We need to find a way to put all the information in every single file together and clean it. Sarah Cohen [shows how to do it in OpenRefine](https://github.com/sarahcnyt/data-journalism/tree/master/openrefine). You will learn how to do it with Python.

## Look for patterns

Computers are really good at doing the same thing over and over. What we want is to find common patterns and assign our script to look for them and repeat the procedure until it's finished. Also, we have to make sure the files share the same pattern -- they have the same number of columns, the same categories, they have the same keywords or empty spaces in similar place, etc.

You may have noticed in the source files that the series go far back 2013. To make things easier, I excluded from this excersise files before 2013. Why? Because the Yankee government changed the structure of the files slightly, with a different number of columns and the type of data. Including them would give an extra layer of complexity to this tutorial. I may include how to extract data from them in the future.

Since the files after 2013 share the same basic structure, we will figure out how get the information we want from one of them and assign our script to do the same thing for all the other files.

## Design a cleaning strategy

### Thinking about columns
Think about which columns your final dataset will have. What are you looking for? In our case, it makes sense to think about having the following columns:

- ```YEAR```
- ```MONTH```
- ```COUNTY```
- ```PLAN NAME```
- ```ADC``` (TANF ADC & MA-ADC)
- ```SNA``` (SNA HR & MA-HR)
- ```SSI``` (SSI & MA-SSI)
- ```NYSoH```
- ```TOTAL ENROLLED```

The other columns are aggregated numbers that we don't really need. We can get the same numbers by adding the values in the columns above.

### Are there and repetitions?
Yes! Several! In every file:

- ...the information we need is always in the first sheet of the ```XLS``` file
- ...dates (month & year) are always in the same cell (```A4```)
- ...names of counties and plans are always in the same columns (```A``` & ```B```)
- ...the list of plan names always start with the string ```TOTALS:```
- ...the county name is always immediately one cell to the left of ```TOTALS:```
- ...the list of plan names always end in an empty space (after that we have data for New York city, which we won't capture)

### But there is a catch...
Notice that in some files the name of the columns are different. There are two cases:

- Files with the ```NYSoH``` column
- Files without the ```NYSoH``` column

No problem. We will take care of each one of those cases populating our ```CSV``` file with the correct data.

### Final thoughts about strategy
Before starting your cleaning procedure take your time evaluating the files. Write down patterns that are more evident. Watch out for differencies between the files. They might be similar, but have slight differences, such as different column names or aggregated data.


# ENOUGH! Show me the code!
Alright. Let's get to it.

### Getting the date
When you open the ```XLS``` files, one of the first things you'll notice is that cell **```A4```** holds the information about the month and year of the current file. That is always the case. We will store this information in two separate columns in our final ```CSV``` file: ```MONTH``` and ```YEAR```. To do that, let's create a function that will give us a dictionary with key/value pairs for the ```MONTH``` and for the ```YEAR```. We will use this dictionary later.


```python
# Gets the date from the specific cell in the documents
# and use regular expression to separate month from year.
# returns a dictionary with YEAR and MONTH
def getDate(worksheet):
    '''
    worksheet: an xlrd book object
    return: dictionary
    '''
    date = worksheet.cell_value(3, 0)
    match = re.match(r'NYS[ ,]+(\w+)[ ,](\d+)', date).groups()
    date = {'YEAR': match[1],
            'MONTH': match[0]
            }
    return date
```

Let's see how this works.

This function needs an ```xlrd``` book object (essentially, the sheet where will extract the information from) that we will create later. It uses the ```cell_value``` method from ```xlrd``` to get the contents of a specific cell. The first parameter (3) is the row and the second is the column. Since we want the contents of cell ```A4``` we use the pair 3 (the count starts at zero!) and 0. We store the contents of cell ```A4``` in the variable ```date```.

After that we do a regex (Regular Expression) search in our ```date``` variable to extract the month and the year. The government sometimes uses a comma to separate ```NYS``` from the month, sometimes it doesn't, which is frustrating. That also happens between month and year. But no matter! Regex to the rescue.

```NYS[ ,]+``` = matches a string that starts with NYS and is followed by either an empty space or a comma and any number of characters after that.

```(\w+)[ ,] =``` will do our first capture (that's why the parethesis are there): any word after the above followed by an empty space or a comma.

```(\d+)``` = will do our second capture, any number of digits after the above.

With our two values captured, we create our dictionary and return it.

### Prep code & storing filenames
We will create our ```CSV``` file row by row. So let's create an empty dictionary to store the information about the current row we will be scraping.


```python
row = {}
```

Let's also assign the path to the folder where the ```XLS``` files are.


```python
basedir = 'data/'
```

And let's use the ```listdir()``` function to get the list of filenames in that folder.


```python
files = listdir(basedir)
```

We will go over each and every ```XLS``` file. We need to tell the script which file it will be working on at any given time. Let's make sure our list has only files with the ```XLS``` extension.


```python
sheets = [filename for filename in files if filename.endswith("xls")]
```

Whoa! Wait a minute. I just did what is called a [list comprehension](https://docs.python.org/2/tutorial/datastructures.html#list-comprehensions): get the filenames in the files list but only those which end with ```'xls'```. Pretty cool, right?

### Thinking ahead
The first line of a ```CSV``` file is usually the name of the columns, you know, the header. Since we will be writing row by row in our ```CSV``` file using a series of repetitions, we need a device to know if we have written the header or not. Let's create a flag that will help us with that.


```python
header_is_written = False
```

We will set ```header_is_written``` to True after we write the header.

## Iterating over the files
Ok. Now we're going to scrape the information from all the files in the sheets list we just created. To do that we will start a simple for loop.

We will also open the worksheet using ```open_workbook``` from ```xlrd```. Since the information we need is in the first sheet of the file, let's use the ```sheet_by_index()``` function to open that sheet:


```python
for filename in sheets:
    worksheet = open_workbook(basedir + filename).sheet_by_index(0)
```

Now let's get the ```date ``` from this file using the function we created in the beginning:


```python
    date = getDate(worksheet)
    row['YEAR'] = date['YEAR']
    row['MONTH'] = date['MONTH']
```

Notice we're still inside the loop, so the code needs to be idented. This will be the case moving forward.

## Our reference column

**Column B** is the most important. It has the name of the companies and all the names of the counties are right next to it. Let's get the values in that column and store them in a list. After that we will iterate over them to get what we need.


```python
    col1_values = worksheet.col_values(1) # column A == 0, column B == 1...
```

We're iterating over ```range``` instead of the values because we need to keep track of the coordinate of the cells. With the coordinates we can get the actual content, so we should be fine.

## Getting the data
We're all set to get the data we need. Remember we need to get the names of the counties, the name of the companies and the values associated with them. Our reference is the value ```TOTALS:``` in **Column B**. If we find it we will be one cell away from the county name. Let's iterate over the values of **Column B**. If we find ```TOTALS:``` let's save the value immediately to the left in our row dictionary, assigning it to the key ```COUNTY```. We will be using the ```cell_value``` method from ```xlrd```. I'm also using the strip function. It will trim trailing and leading spaces from the cell value, just in case.


```python
    for i in range(len(col1_values)):
        if worksheet.cell_value(i, 1) == u'TOTALS:':
            row['COUNTY'] = worksheet.cell_value(i, 0).strip()
            county = True
            counter = 1
```

Now we need to get the plan names that are associated with that county. Our robot is still looking for the ```TOTALS:``` value and when it finds it, we will have reached the next county. So let's create a flag called county. ```True``` if the robot doesn't find the next ```TOTALS:```, ```False``` if it hits the next county. Also, since the plan names are the plan names right after ```TOTALS:```, we create a counter that will basically increment the row number by one to get the values down the list until we find another ```TOTALS:```.


```python
            while county is True:
                row['PLAN NAME'] = worksheet.cell_value(i + counter, 1).strip()
                values = worksheet.row_slice(i + counter, start_colx=2, end_colx=7)
```

So, while the county flag is ```True```, we will store the ```PLAN NAME``` in the dictionary adding the counter to the position of the row. We're also using the nifty ```row_slice``` method from ```xlrd``` to generate a list of all the associated values for that ```PLAN NAME```, from **Columns C** (2) to **G** (6 - the parameter ```end_colx``` gets up to the number before the one you set).

Since we have a list of values in the same order of the columns, we will assign each one of them to the appropriate key in the ```row``` dictionary. The list won't have the actual values, but ```xlrd``` objects with the values. To get the values we need to use the getter ```.value```. We're also converting them to integers with the ```int()```  function.

# BUT WAIT!

Remmeber that some files have the column ```NYSoH``` and others don't. Depending on the case, the values on the list will change places. A simple ```if/else``` statement will take care of that. The ```SSI``` column becomes the 4th number in the list whenever ```NYSoH``` is not present. When ```NYSoH``` is not present, let's give it an empty value.

We have to give ```NYSoH``` an empty value because we will be writing rows to the ```CSV``` file. All the rows must have the same amount of values, otherwise things will break.

We're also incrementing ```counter``` by 1 so that whenever it goes back to the beginning of the loop, it will look for the row after the one we just stored.


```python
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
```

Everything is good. We have the info we need to write the row in the ```CSV``` file. Before we do that, let's set the county flag to ```False``` if it finds ```TOTALS:``` once it increments the counter. Also, if the ```PLAN NAME``` in that row was assigned an empty space, it means we reached the end of the counties column. The empty space is what separates the data from the counties from the data for New York city. If that's the case, we will ```break``` to get outside of this loop and go to the next file without writing the row with the empty space.


```python
                if worksheet.cell_value(i + counter, 1) == u'TOTALS:':
                    county = False
                # If we reach a point in this column that the value is blank, it means our search in this file is over!
                elif row['PLAN NAME'] == '':
                    break
```

Alright. Time to write the row dictionary to the ```CSV``` file. Easy. We will use the ```CSV``` module we imported earlier. Remember the flag we created for the header? The handy thing about having stored all the row information in a dictionary is that we can use the ```writeheader()``` function to get the name of the keys and write the header. Once we do that, we update the ```header_is_written``` flag to ```True``` and never use it again, since we need to do this only once, at the very first iteration :)


```python
                # Saving the row in the CSV file...
                with open('table.csv', 'ab') as f:
                    w = csv.DictWriter(f, row.keys())
                    if header_is_written is False:
                        w.writeheader()
                        header_is_written = True
                    w.writerow(row)
```

# That's it?

Yeah. That's it.

There's no formula that will fit all files out there. Every situation will bring new challenges and you will have to figure out what the patterns are and the best way to approach them. There are many ways to clean the same file. We could have used a column other than **B** as our reference, for example. Use the approach that makes more sense to you.

Let's take a look at the whole code, shall we? :)


```python
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
files = listdir(basedir)

# Empty list to store only XLS files found in the folder
sheets = [filename for filename in files if filename.endswith("xls")]

header_is_written = False
# Iterating over the files in folder
for filename in sheets:
    # Opens the xls file
    worksheet = open_workbook(basedir+filename).sheet_by_index(0)
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
```
