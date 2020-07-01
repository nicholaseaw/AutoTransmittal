# Load the Python Standard and DesignScript Libraries
import sys
import clr
import os
import os.path
import csv
clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *

#variables
revdates = IN[0]
currentdate = IN[1]
dir = IN[2]
PATH = dir + "\\" + "data.csv"
lst = []
datejoin = []


#filter out empty list
if currentdate in revdates:
	pass
else:
	revdates.append(currentdate)

filtered = [x for x in revdates if len(x.strip()) > 0]

#read data from csv
if os.path.isfile(PATH) and os.access(PATH, os.R_OK):
	f = open(PATH , 'rb')
	reader = csv.reader(f)
	for row in reader:
		lst.append(row)
	datelst = lst[:3]

#join back dates
for i in range(len(datelst)):
	for j in range(len(datelst[i])):
		if i == j:
			datejoin.append(datelst[i][j] + '/' + datelst[i+1][j] + '/' + datelst[i+2][j])

#check if dates from csv match all issuance dates 
datebool = [x in datejoin for x in filtered]

#get index of date not matched
idx = [i for i, x in enumerate(datebool) if not x]

#append date not matched into list
for i in idx:
	datejoin.append(filtered[i])
# Assign your output to the OUT variable.
OUT = datejoin