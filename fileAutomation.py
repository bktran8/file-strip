from openpyxl import load_workbook
from collections import defaultdict
import time
import os.path

# initilize and decalre key and values with file attributes
def FileInfoDict(dictionary, dataPath):
	for root, dirs, files in os.walk(dataPath):
		path, deptCategory = os.path.split(root)
		for file in files:
			filePath = os.path.join(root, file)
			modification = time.ctime(os.path.getmtime(filePath))
			created = time.ctime(os.path.getctime(filePath))
			dictionary[file].append(modification)
			dictionary[file].append(created)
			dictionary[file].append(deptCategory)
	return(dictionary)

# create and load dictionary into appropriate columns with department infomartion/data
def DeptInfo(dictionary, excel):
	nextRow = 4
	workBook = load_workbook(excel)
	sheet = workBook.active
	deptName = input("What is the department name?\n")
	sheet['B2'] = deptName
	meetingDate = input("When is the meeting date?\n")
	sheet['F2'] = meetingDate
	contact = input("Who is the point of contact?\n")
	sheet['H2'] = contact 
	contactEmail = input("What is {}'s email?\n".format(contact))
	sheet['K2'] = contactEmail
	for key, values in dictionary.items():
		sheet['D' + str(nextRow)] = key
		sheet['B' + str(nextRow)] = values[0]
		sheet['A' + str(nextRow)] = values[1]
		sheet['H' + str(nextRow)] = values[2]
		nextRow += 1
	newFile = deptName+'.xlsx'
	workBook.save(newFile)
	

		
	
 


