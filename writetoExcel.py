#libraries
import sys
import clr
import csv
clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *
clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
from Microsoft.Office.Interop import Excel
from System import Array

#declare inputs
dict = IN[0]
date = IN[1]
dir = IN[2]
rev = IN[3]
stakeholders = IN[4]
paststatus = IN[5]
updatedemail = IN[6]
email = IN[7]
datesplit = []
issuecount = []
startrow = 27
startcol = 6

#file path
fp_xls = dir + "\\" + "Master Drawing Issue Sheet.xls"
fp_csv = dir + "\\" + "data.csv" 

#open Excel doc
excel = Excel.ApplicationClass()
excel.Visible = True
excel.DisplayAlerts = False

workbook = excel.Workbooks.Open(fp_xls)
ws = workbook.Worksheets[1]

#combine email from csv and updated email from user input
email_bool = [x in email for x in updatedemail]

if any(email_bool):
	combinedemail = updatedemail
elif not any(email_bool):
	combinedemail = email + [updatedemail]

#transpose back revision list
rev_transpose = map(list,zip(*rev))

#transpose back status list
status_transpose = map(list,zip(*paststatus))

#split date into DD/MM/YY format
for i in date:
	datesplit.append(i.split('/'))

#count number of sheets issued
count = 0
for i in rev[-1]:
	if i != '':
		count +=1
		issuecount.append(count)
		
#create dates arrays
for i in range(len(datesplit)):
	datesarr = Array.CreateInstance(object,3, len(date))

#assign values to dates arrays
for i in range(len(datesplit)):
	for j in range(len(datesplit[i])):
		datesarr[j,i] = datesplit[i][j]

#get project code
splitsheetnum = dict[1][0].split('/')

#create stakeholders arrays
for i in range(len(stakeholders)):
	stakeholdersarr = Array.CreateInstance(object, len(stakeholders), 1)

#define range of cells
#fixed columns
xlrangesheetnames = ws.Range[ws.Cells(startrow,1), ws.Cells(len(issuecount)+startrow-1,1)]
xlrangesheetnumbers = ws.Range[ws.Cells(startrow,4), ws.Cells(len(issuecount)+startrow-1,4)]
xlrangestakeholders = ws.Range[ws.Cells(startrow-16,startcol-4), ws.Cells(len(stakeholders)+startrow-17,startcol-4)]
xlrangeemail = ws.Range[ws.Cells(startrow-16,startcol), ws.Cells(len(updatedemail)+startrow-17,startcol+ len(date)-1)]

#fixed rows
xlrangesheetstatus = ws.Range[ws.Cells(startrow-2,startcol), ws.Cells(startrow-2,len(date)+startcol-1)]
xlrangerecord = ws.Range[ws.Cells(startrow-4,startcol),ws.Cells(startrow-4,len(date)+startcol-1)]

#range of rows and columns
xlrangedate = ws.Range[ws.Cells(startrow-21,startcol), ws.Cells(startrow-19,len(date)+startcol-1)]
xlrangerevision = ws.Range[ws.Cells(startrow,startcol), ws.Cells(len(rev[0])+startrow-1, len(date)+startcol-1)]

#create arrays
sheetnamearr = Array.CreateInstance(object,len(dict[0]) , 1)
sheetnumberarr = Array.CreateInstance(object,len(dict[1]), len(date))
revisionarr = Array.CreateInstance(object,len(rev[0]), len(date))
recordarr = Array.CreateInstance(object,1, len(date))
emailarr = Array.CreateInstance(object,len(updatedemail),len(date))


#assign values to record array
for i in range(len(date)):
	recordarr[0,i] = 1

#create status array
statusarr = Array.CreateInstance(object,1, len(paststatus))

#assign values to status array
for i in range(len(paststatus)):
	statusarr[0,i] = paststatus[i][0]

#assign values to arrays
for i in range(len(dict)):
	for j in range(len(dict[i])):
		sheetnamearr[j,0] = dict[0][j]
		sheetnumberarr[j,0] = dict[1][j]

for i in range(len(date)):
	for j in range(len(rev[i])):	
		revisionarr[j,i] = rev[i][j]

for i in range(len(stakeholders)):
	stakeholdersarr[i,0] = stakeholders[i]

#assign values to email array
if any(email_bool):
	for i in range(len(combinedemail)):
		emailarr[i,0] = combinedemail[i]
elif not any(email_bool):
	for i in range(len(combinedemail)):
		for j in range(len(combinedemail[i])):
			emailarr[j,i] = combinedemail[i][j]


#transpose email arr
if any(email_bool):
	email_transpose = combinedemail
elif not any(email_bool):
	email_transpose = map(list,zip(*combinedemail))

#transpose date split
date_transpose = map(list,zip(*datesplit))

#write to csv
with open(fp_csv, 'wb') as csvfile:
	csvwriter = csv.writer(csvfile)
	csvwriter.writerows(date_transpose)
	csvwriter.writerows(email_transpose)
	csvwriter.writerows(status_transpose)
	csvwriter.writerows(rev_transpose)

#write arrays to excel		
xlrangerevision.Value2 = revisionarr
xlrangesheetnames.Value2 = sheetnamearr
xlrangesheetnumbers.Value2 = sheetnumberarr 
xlrangesheetstatus.Value2 = statusarr
xlrangerecord.Value2 = recordarr 
xlrangedate.Value2 = datesarr
xlrangestakeholders.Value2 = stakeholdersarr
xlrangeemail.Value2 = emailarr

OUT = 'Success!'