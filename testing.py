import pandas as pd
import recordlinkage as rl

#read the data files, for this example .csv
#read_csv may be switched for read_excel, but openpyxl should be used.
data_a = pd.read_csv("testa.csv",",")
data_b = pd.read_csv("testb.csv",",")
slicd = pd.DataFrame(data_a,columns=["Last"])

print(type(data_a))

#creating empty lists for cycling
names_a = []
names_b = []

for x,y in zip(data_a['First '], data_a['Last']):
	names_a.append(x + ' ' + y)
for x,y in zip(data_b['First '], data_b['Last']):
	names_b.append(x + ' ' + y)

#replacing the name list with a merged name list.
#This will create a cell with no repeated valued.
data_a['Name'] = names_a
data_b['Name'] = names_b

#checking data
print(data_a.head())
print(data_b.head())

#using record linkage to compare the two .csv files
#values are either compared exactly or by a string method
indexer = rl.Index()
indexer.block('Name')
full_pairs = indexer.index(data_a, data_b)
ccl = rl.Compare()
ccl.exact('Name', 'Name', label='Name')
ccl.string('city', 'city', label='city')
ccl.exact('Phone','Phone',label='Phone')
potential_matches=ccl.compute(full_pairs,data_a,data_b)
matches = potential_matches[potential_matches.sum(axis=1) >= 3]
matches.index

dup = matches.index.get_level_values(1)
dup_b = data_b[~data_b.index.isin(dup)]

#create empty counts for accurate cycling
count_1 = 0
count_2 = 0

#cycle through values and replace them. The amount of repeats will change depending on columns.
#this method is preferred personally as it allows count to be reset each time
#rather than creating a longer loop with more count variables
for x in data_a['Name']:
	if x in list(data_b['Name']):
		if data_a.iloc[count_2,2] != data_b.iloc[count_1,2]:
			data_a.iloc[count_2,2] = data_b.iloc[count_1,2]
		count_1 += 1
	count_2 += 1

count_1 = 0
count_2 = 0

for x in data_a['Name']:
	if x in list(data_b['Name']):
		if data_a.iloc[count_2,3] != data_b.iloc[count_1,3]:
			data_a.iloc[count_2,3] = data_b.iloc[count_1,3]
		count_1 += 1
	count_2 += 1
	
count_1 = 0
count_2 = 0

for x in data_a['Name']:
	if x in list(data_b['Name']):
		if data_a.iloc[count_2,4] != data_b.iloc[count_1,4]:
			data_a.iloc[count_2,4] = data_b.iloc[count_1,4]
		count_1 += 1
	count_2 += 1

#append the data
full_data = data_a.append(dup_b)

print(full_data)

#export to (existing) excel file
full_data.to_excel("final.xlsx")
