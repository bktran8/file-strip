import glob 
import xlrd
import os.path
import zipfile
from xlrd import open_workbook

#Initialize dictionary for hashing
def InitDict(dictionary, excel):
	stuff = {':', '/'}
	workBook = xlrd.open_workbook(excel)
	sheet = workBook.sheet_by_index(0)
	for row in range(1, sheet.nrows):
		fileKey = sheet.cell(row, 4).value
		fileValue = sheet.cell(row, 0).value
		for char in fileValue:
			if char in stuff:
				fileValue = fileValue.replace(char, ' -')
			elif char is '"':
				fileValue = fileValue.replace('"', '')
		dictionary[fileKey] = fileValue
	workBook.close()
	return(dictionary)

#Rename files with extensions 
def RenameFile(dictionary, path):
	for file in os.listdir(path):
		fileName = os.path.splitext(file)
		if file not in dictionary:
			print("No friendly name for:", file)
			print(" ")
		else:
			newFileName = dictionary[file]
			oldPath = os.path.join(path+file)
			newPath = os.path.join(path, newFileName+fileName[1])
			print(" ")
			print("Old Path ... ")
			print(oldPath)
			print(" ")
			print("Now changing paths ... \n")
			os.replace(oldPath, newPath)
			print("New Path... ")
			print(newPath)
			print(" ")
	#Double check what files are in path 
	print("Double checking whats in directory...")
	for files in os.listdir(path):
		print(files)
		print(" ")
					
#Compress files over 20MB 
def CompressFile(path):
	for filePath in glob.glob(os.path.join(path, '*.*')):
		file = (os.path.basename(filePath))
		fileName, fileExt = os.path.splitext(file)
		size = os.path.getsize(filePath)
		if size > 20971520:
			fileZip = zipfile.ZipFile(os.path.join(path,fileName+'.zip'), 'w', allowZip64 =True)
			fileZip.write(filePath, os.path.relpath(filePath, path),compress_type = zipfile.ZIP_DEFLATED)
			fileZip.close()
			os.remove(filePath)
					
#Delete/replace special characters
def Strip(path):
	stuff = {'(', ')', '#', '\'', ',', '&', '.'}
	morestuff = {'--', ' '}
	for filePath in glob.glob(os.path.join(path, '*.*')):
		file = (os.path.basename(filePath))
		fileName, fileExt = os.path.splitext(file)
		for char in fileName:
			if char in stuff:
				fileName = fileName.replace(char, '')
			if char in morestuff:
				fileName = fileName.replace(char, '-')	
		newPath = os.path.join(path, fileName+fileExt)
		os.replace(filePath, newPath)
	print(" ")
	print("Double checking whats in directory...")
	for files in os.listdir(path):
		print(files)
		print(" ")