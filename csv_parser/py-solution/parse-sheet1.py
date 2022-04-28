#!/usr/bin/python3 env

import numpy as np
import pandas as pd

# start with the list of names you want to work with
sales = pd.read_excel('TUK-TUK-04.2022.xlsx')


# create a list object for the names in our documnent and return it in chunks of 10
dsa_sales = list(sales['value'])


# create an empty list to populate with the string objects from the group of sales above
sales_strings = []

# iterate through the dsa_sales list and append to the sales_strings list by joining them at comma seperations
for i in range(len(sales)):
	sales_strings.append(','.join(dsa_sales[i]))
	print(sales_strings)


# reinitialise the pandas final data-frame object with new values/parameters
final_dataframe = pd.DataFrame(columns = my_columns)


# loop through the list of stocks and fetch data for every symbol in that list
for sales_string in sales_strings:

	# put response in data variable/requests object and parse it using the .json() function
	data = open('TUK-TUK-04.2022.xlsx').json()

	# iterate through  the sales_strings list and append the strings to the final data-frame at comma seperations
	for sale in sales_string.split(','):
		final_dataframe = final_dataframe.append(
			# inside the pandas data-frame add a pandas series object with the data to append

			pd.Series([
				symbol,
				data[symbol]['value']['package'], #
				data[symbol]['value']['DSA'],
				'N/A'

			], # this is needed for the objects we are working with in pandas otherwise it won't work
			index = my_columns),

			# if this is left out the pandas objects will throw an error
			ignore_index = True
		)



# iterate through a range from 0 to the total length of the index and calculate how many sales each dsa made
for i in range(0, len(final_dataframe.index)):
	final_dataframe.loc[i, 'DSA'] = math.floor(final_dataframe.loc[i, 'value'])  #<-needs an equation


# the final number of sales by adding the total number of sales in the document
#number_of_sales = position_size/500
