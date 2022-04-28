#!/usr/bin/python3
# by jace mitnick

# we start by importing the module/library we are going to use to interact with our files
import pandas as pd

#

# sales = pd.read_excel('TUK-TUK-04.2022.xlsx')

# here we load the data from the excel file we want to work with
ex_file = "../TUK-TUK-04.2022.xlsx"


# create new read_excel/csv pandas dataframe in variable from the data in the file we loaded
# names is for the names of the dsa
names = pd.read_excel(ex_file, usecols=['DSA'])

# d_sales is for both the packages and their amount
d_sales = pd.read_excel(ex_file, usecols=['Package', 'Value'])


# try converting to json
#new_json = names.to_json()


# here we start the heavy lifting of the algorithm

# names of the dsa
names = []
# package = []
amount = []

for a,b in rows:
    names.append(a.value)
    # package.append(s.value)
    amount.append(b.value)


for i in names:
	for j in d_sales:
		print(f'name:{i} \n')
		print(f'sale:{j} \n')



# function to sort the names into unique occurrence
# memo pad used for the sorting
# def sort_names(dnames, memo={}, sorted_names=[]):
	# # iterator to crawl the file length
	# for i in range(len(dnames):
	# 	# name variable to fetch names from list
	# 	for name in dnames:
	# 		if name in memo:
	# 			sorted_names += memo[name]
	# 			i = i + 1
	# 		else:
	# 			sorted_names = memo[name]
	# 			i = i + 1
	# return sorted_names
	



# check dataframe content
print(names, d_sales)
