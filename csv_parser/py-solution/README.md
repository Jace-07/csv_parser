# python implementation for simple xml spreadsheet formatting


## source: "https://www.marsja.se/your-guide-to-reading-excel-xlsx-files-in-python/"
`
What is the use of Openpyxl in Python?
Openpyxl is a Python module that can be used for reading and writing Excel (with extension xlsx/xlsm/xltx/xltm) files. Furthermore, this module enables a Python script to modify Excel files. For instance, if we want togo through thousands of rows but just read certain data points and make small changes to these points, we can do this based on some criteria with openpyxl.

How do I read an Excel (xlsx) File in Python?
Now, the general method for reading xlsx files in Python (with openpyxl) is to import openpyxl (import openpyxl) and then read the workbook: wb = openpyxl.load_workbook(PATH_TO_EXCEL_FILE). In this post, we will learn more about this, of course.

How to Read an Excel (xlsx) File in Python
Now, in this section, we will be reading a xlsx file in Python using openpyxl. In a previous section, we have already been familiarized with the general template (syntax) for reading an Excel file using openpyxl and we will now get into this module in more detail. Note, we will also work with the Path method from the Pathlib module.
`

### 1. Import the Needed Modules
In the first step, to reading a xlsx file in Python, we need to import the modules we need. That is, we will import Path and openpyxl:

```python
import openpyxl
from pathlib import Path
```

[!https://www.marsja.se/wp-content/uploads/2020/02/step1-to-read-xlsx-files-import-the-modules.jpg]

### 2. Setting the Path to the Excel (xlsx) File
In the second step, we will create a variable using Path. Furthermore, this variable will point at the location and filename of the Excel file we want to import with Python:

# Setting the path to the xlsx file:
```python
xlsx_file = Path('SimData', 'play_data.xlsx')</code></pre>
print(xlsx_file)
```

Note, “SimData” is a subdirectory to that of the Python script (or notebook). That is, if we were to store the Excel file in a completely different directory, we need to put in the full path. For example, xlsx_file = Path(Path.home(), 'Documents', 'SimData', 'play_data.xlsx')if the data is stored in the Documents in our home directory.

[https://www.marsja.se/wp-content/uploads/2020/02/step2-to-reading-a-xlsx-file-using-Python-set-the-path.jpg]


### 3. Read the Excel File (Workbook)
In the third step, we are actually going to use Python to read the xlsx file. Now, we are using the load_workbook() method:

```python
wb_obj = openpyxl.load_workbook(xlsx_file)
print(wb_obj)
```

### 4. Read the Active Sheet from the Excel file
Now, in the fourth step, we are going to read the active sheet using the active method:

```python
sheet = wb_obj.active
print(sheet)
```

Note, if we know the sheet name we can also use this to read the sheet we want: play_data = wb_obj['play_data']

### 5. Work, or Manipulate, the Excel Sheet
In the final, and fifth step, we can work, or manipulate, the Excel sheet we have imported with Python. For example, if we want to get the value from a specific cell we can do as follows:

```python
print(sheet["C2"].value)
```

Code language: Python (python)
Another example, on what we can do with the spreadsheet in Python, is that we can iterate through the rows and print them:

```python
for row in sheet.iter_rows(max_row=6):
    for cell in row:
        print(cell.value, end=" ")
    print()
```

Note, that we used the max_row and set it to 6 to print the 6 first row from the Excel file.

### 6. Bonus: Determining the Number of Rows and Columns in the Excel File
In the sixth, and bonus step, we are going to find out how many rows and columns we have in the example Excel file we have imported with Python:

```python
print('Number of rows: ' + str(sheet.max_row) + '\nNumber of sheets: ' + str(sheet.max_column))
```

Reading an Excel (xlsx) File to a Python Dictionary
Now, before we learn how to read multiple xlsx files we are going to import data from Excel into a Python dictionary. It’s quite simple, but for the example below, we need to know the column names before we start. If we want to find out the column names we can run the following code (or just open the Excel file):

```python
import openpyxl
from pathlib import Path

xlsx_file = Path('SimData', 'play_data.xlsx')
wb_obj = openpyxl.load_workbook(xlsx_file)
sheet = wb_obj.active

col_names = []
for column in sheet.iter_cols(1, sheet.max_column):
    col_names.append(column[0].value)
   
    
print(col_names)
Code language: Python (python)
Creating a Dictionary from an Excel File
In this section, we will finally read the Excel file using Python and create a dictionary.

data = {}

for i, row in enumerate(sheet.iter_rows(values_only=True)):
    if i == 0:
        data[row[1]] = []
        data[row[2]] = []
        data[row[3]] = []
        data[row[4]] = []
        data[row[5]] = []
        data[row[6]] = []

    else:
        data['Subject ID'].append(row[1])
        data['First Name'].append(row[2])
        data['Day'].append(row[3])
        data['Age'].append(row[4])
        data['RT'].append(row[5])
        data['Gender'].append(row[6])
```
Code language: Python (python)
Now, let’s walk through the code example above. First, we create a Python dictionary (data). Second, we loop through each row (using iter_rows) and we only go through the rows where there are values. Second, we have an if statement where we check if it’s the first row and we add the keys to the dictionary. That is, we set the column names as keys. Third, we append the data to each key (column name) in the else statement.

How to Read Multiple Excel (xlsx) Files in Python
In this section, we will learn how to read multiple xlsx files in Python using openpyxl. Additionally to openpyxl and Path, we are also going to work with the os module.

### 1. Import the Modules
In the first step, we are going to import the modules Path, glob, and openpyxl:

```python
import glob
import openpyxl
from pathlib import Path
Code language: Python (python)
```

### 2. Read all xlsx Files in the Directory to a List
Second, we are going to read all the .xlsx files in a subdirectory into a list. Now, we use the glob module together with Path:

```python
xlsx_files = [path for path in Path('XLSX_FILES').rglob('*.xlsx')]
Code language: Python (python)
```

### 3. Create Workbook Objects (i.e., read the xlsx files)
Third, we can now read all the xlsx files using Python. Again, we will use the load_workbook method. However, this time we will loop through each file we found in the subdirectory,

```python
wbs = [openpyxl.load_workbook(wb) for wb in xlsx_files]
Code language: Python (python)
```

Now, in the code examples above, we are using Python list comprehension (twice, in both step 2 and 3). First, we create a list of all the xlsx files in the “XLSX_FILES” directory. Second, we loop through this list and create a list of workbooks. Of course, we could add this to the first line of code above.

### 4. Work with the Imported Excel Files
In the fourth step, we can now work with the imported excel files. For example, we can get the first file by adding “[0]” to the list. If we want to know the sheet names of this file we do like this:wbs[0].sheetnames .That is, many of the things we can do, and have done in the previous example on reading xlsx files in Python, can be done when we’ve read multiple Excel files.

Notice, this is one great example of how to use this programming language. Other examples are, for instance, to use it for renaming files in Python.

Conclusion: Reading Excel (xlsx) Files in Python
In this post, we have learned how to:

Read an Excel file in Python using openpyxl
Read a xlsx file to a Python dictionary
Read multiple Excel fils in Python