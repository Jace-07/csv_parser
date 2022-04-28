[! python implementation for simple xml spreadsheet formatting]


[_source: "https://www.marsja.se/your-guide-to-reading-excel-xlsx-files-in-python/"]

What is the use of Openpyxl in Python?
Openpyxl is a Python module that can be used for reading and writing Excel (with extension xlsx/xlsm/xltx/xltm) files. Furthermore, this module enables a Python script to modify Excel files. For instance, if we want togo through thousands of rows but just read certain data points and make small changes to these points, we can do this based on some criteria with openpyxl.

How do I read an Excel (xlsx) File in Python?
Now, the general method for reading xlsx files in Python (with openpyxl) is to import openpyxl (import openpyxl) and then read the workbook: wb = openpyxl.load_workbook(PATH_TO_EXCEL_FILE). In this post, we will learn more about this, of course.

How to Read an Excel (xlsx) File in Python
Now, in this section, we will be reading a xlsx file in Python using openpyxl. In a previous section, we have already been familiarized with the general template (syntax) for reading an Excel file using openpyxl and we will now get into this module in more detail. Note, we will also work with the Path method from the Pathlib module.

