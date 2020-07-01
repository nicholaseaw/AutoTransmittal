# Load the Python Standard and DesignScript Libraries
import sys
import clr
import os
import os.path
import csv
clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *

#inputs
allrev = IN[0]
allrevdates = IN[1]
issueindex = IN[2]
issuance = IN[3]
dir = IN[4]
sheetstatus = IN[5]
consultants = IN[6]
updatedemail = IN[7]
lst = []
newlst = []
newindex = []
status_key = []

#check if data backup is created
PATH = dir + "\\" + "data.csv"

def revarray():
#cross check temp array with current data
	for i in range(0, len(lst)):
		if lst[i][-1] == allrev[i][-1] and issuance[i] == True:
#no changes to rev and append to existing rev
			newindex = len(lst[i])
			for rev in reversed(lst[i]):
				if rev != '':
					lastrev = lst[i][-1]
					break
			lst[i].insert(newindex,lastrev)
		elif lst[i][-1] != allrev[i][-1] and issuance[i] == True:
#get new revision and append to existing rev
			newindex = len(lst[i])
			lst[i].insert(newindex, allrev[i][-1])
#get last revision before non-issuance of sheets
		elif lst[i][-1] == '' and issuance[i] == True:
			newindex = len(lst[i])
			for rev in reversed(lst[i]):
				if rev == '':
					pass
				elif rev != '':
					lastrev = allrev[i][-1]
					break
			lst[i].insert(newindex,lastrev)
#blank issuance if not issue now
		elif lst[i][-1] == '' and issuance[i] == False:
			newindex = len(lst[i])
			lst[i].insert(newindex,'')
		elif lst[i][-1] == '*' and issuance[i] == False:
			newindex = len(lst[i])
			lst[i].insert(newindex,'')
		elif lst[i][-1] == allrev[i][-1] and issuance[i] == False:
			newindex = len(lst[i])
			lst[i].insert(newindex,'')
				
def statusarray():
	len_sheetstatus = len('Sheet Status : ')
	for i in range(len(allrevdates)):
		status_key.append(sheetstatus[i][len_sheetstatus:])
	for j in range(len(status_key)):
		status_key[j] = status_key[j][:1]
	statuslst.append(status_key[-1])

if os.path.isfile(PATH) and os.access(PATH, os.R_OK):
	f = open(PATH , 'rb')
	reader = csv.reader(f)
#create temp array of data
	for row in reader:
		lst.append(row)
	datelst = lst[:3]
	consultantlst = lst[3:len(consultants)+3]
	statuslst = lst[len(consultants)+3]
	lst = lst[len(consultants)+3+1:]

#create sublist for consultantlst if empty list
	newconsultantlst = []
	for i in range(len(consultantlst)):
		if len(consultantlst[i]) > 0:
			pass
		elif len(consultantlst[i]) == 0:
			consultantlst[i] = [consultantlst[i:+1]]
			
#check if len of csv match len of rev lst
	if len(lst) < len(allrev):
		newlstcount = len(allrev) - len(lst)
		for i in range(newlstcount+1,len(allrev)):
			lst.append([[]]*len(statuslst))
	elif len(lst) == len(allrev):
		pass
#replace empty list with  null value in lst
	for i in range(len(lst)):
		for j in range(len(lst[i])):
			if len(lst[i][j]) == 0:
				lst[i][j] = ''
#replace empty list with null value in consultantlst
	for i in range(len(consultantlst)):
		for j in range(len(consultantlst[i])):
			if len(consultantlst[i][j]) == 0:
				consultantlst[i][j] = ''

#count number of sheets from data csv and all revisions list
	lstcount = []
	allrevcount = []
	count = 0
	for i in range(len(lst)):
		if len(lst[i][0]) != 0:
			count = count + 1
			lstcount.append(count)
	count = 0
	for i in range(len(allrev)):
		if len(allrev[i][0]) != 0:
			count = count + 1
			allrevcount.append(count)

#check if new sheets are added and count number of new sheets
	if len(lstcount) == len(allrevcount):
#run function
		revarray()
		statusarray()	
	elif len(allrevcount) > len(lstcount):
		newsheetcount = len(allrevcount)
#get new list of new sheets and get indexes of new sheets
		if len(allrevcount) - len(lstcount) == 1:
			newsheetindex = len(allrevcount)
			newsheetlist = allrev[len(lstcount) : newsheetindex]
		elif len(allrevcount) - len(lstcount) > 1:
			newsheetindex = range(len(lstcount), len(allrevcount))
			newsheetlist = allrev[len(lstcount) : len(lstcount) + len(newsheetindex)]

#run function		
		revarray()
		statusarray()

#if no data csv, then use current data
else:
	for rev in allrev:
		for i in rev:
			if i == -1:
				revindex = rev.index(i)
				if revindex != 0:
					currentrev = rev[revindex-1]
					rev[revindex] = currentrev
				elif revindex == 0:
					rev[revindex] = ''
	for i in range(len(allrev)):
		if i not in issueindex:
			for j in range(len(allrevdates)):
				if j == (len(allrevdates) - 1):
					allrev[i][j] = ''
	len_sheetstatus = len('Sheet Status : ')
	for i in range(len(allrevdates)):
		status_key.append(sheetstatus[i][len_sheetstatus:])
	for j in range(len(status_key)):
		status_key[j] = status_key[j][:1]
	

# outputs
OUT = [lst,statuslst,consultantlst] if os.path.isfile(PATH) and os.access(PATH, os.R_OK) else [allrev,[status_key[-1]],updatedemail]