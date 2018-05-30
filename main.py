import fileStrip
import fileAutomation
from datetime import datetime
from collections import defaultdict

def runStrip():
	print("Running Strip Automation...\n")
	stripSearch = {}
	excel = input("What is the excel sheet?\n")
	path = input("What is the path?\n")
	fileStrip.InitDict(stripSearch, excel)
	print(" ")
	print("Renaming files ...")
	fileStrip.RenameFile(stripSearch, path)
	print("Stripping special characters... ")
	fileStrip.Strip(path)
	fileStrip.CompressFile(path)

def runAutomation():
	print("Running Form Automation...\n")
	dataDict = defaultdict(list)
	excel = input("What is the excel sheet?\n")
	dataPath = input("What is the path to department data?\n")
	fileAutomation.FileInfoDict(dataDict, dataPath)	
	fileAutomation.DeptInfo(dataDict, excel)

def main():
	userInput = None
	choices = {"a", "b"}
	while userInput not in choices:
		userInput = input("Do you want to\nA. Strip files\nB. Fill a form out [A/B]?\n")
	if userInput.lower() == "a":
		runStrip()
	elif userInput == "b":
		runAutomation()
	


    

if __name__=='__main__': main()

