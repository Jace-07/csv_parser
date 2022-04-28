#!/usr/bin/python3
# by jace mitnick


import openpyxl

# new module openpyxl testing
#
ex_file = "../TUK-TUK-04.2022.xlsx"

# create new read_excel/csv pandas dataframe in variable and select worksheet
wb = openpyxl.load_workbook(ex_file)
ws = wb['Sheet1']


rows = ws.iter_rows(min_row=1, min_col=1)
# print(rows)

# d_sales = ['Package', 'Value'])
# try converting to json
#new_json = names.to_json()


# names of the dsa
names = []
# package = []
amount = []

for a,b in rows:
    names.append(a.value)
    # package.append(s.value)
    amount.append(b.value)

# check dataframe content
print(names, amount)

