import pandas as pd
import numpy as np
import pyodbc
import math
import os
import shutil
import openpyxl
from pathlib import Path
import re
import msvcrt
import traceback
import sys
import fileParse as fp

workSheetTabs = '5.1 Targeting CMS - Content','5.2 Targeting CMS - Banners'
workSheetCols = 'Message name', 'Name'
mCodes = []
bitmarks = []

def getNewWeek():
	#Check downloads folder for myCD targeting script. Can use this to add new week. Will need to create list of PELWeek to not make F2153+
	print('yes')

def initCurrentWeekData():
	global CurrentWeek, fileCurrentWeek, scriptFile, pPath
	pPath = Path(os.getcwd()).parent #Get Parent Folder and define here
	
	CurrentWeek = detectLatestCurrentWeek()
	fileCurrentWeek = CurrentWeek % 100
	scriptFile = f"{pPath}\F{CurrentWeek}\W{fileCurrentWeek} MyCD Scripts.xlsx" 

def StartSheet():
	print('Parsing Worksheet')

	parseTargeting()
	checkScript()
	
	input("Press Enter to continue...")	

def detectLatestCurrentWeek():
	with os.scandir(pPath) as local:
		i = 0
		wkval=0
		for path in local:
			if not path.is_file():
				i += 1
				if(len(path.name) > 1 and path.name.startswith('F') and path.name[1:].isdecimal()):
					if(int(path.name[1:])) > wkval:
						wkval = int(path.name[1:])
		return wkval

def parseTargeting():
	file = detectLatestWorksheet()
	location = f"{pPath}\F{CurrentWeek}\{file}" 
	
	#Get targeting sheets (can't hardcode sheet indexes as it changes sometimes)
	df = pd.read_excel(location, None);
		
	indices = [i for i, elem in enumerate(df.keys()) if 'Targeting CMS -' in elem]
	
	for i in indices:
		parseWorkSheet(location, i)	

def detectLatestWorksheet():
	reg = re.compile(f"(copy of )?mycd worksheet( )+(- +?)?wk{fileCurrentWeek}(.)*?v.+xlsx") #regex to match filenames, takes into account clone of files
	
	print(os.path.join(pPath,f"f{CurrentWeek}"))	
	
	with os.scandir(os.path.join(pPath,f"f{CurrentWeek}")) as local:
		i = 0
		fileval=0
		toReturn = ""
		
		for file in local:
			if file.is_file():
				i += 1
				if(bool(reg.match(file.name.lower()))):
					culled = file.name[file.name.lower().index("v")+1:file.name.lower().index(".")]
					try:
						if(culled.isdecimal()):
							
							temp = int(culled)
							
							if temp > fileval:
								fileval = temp
								toReturn = file.name
					except ValueError:
						fileval = fileval
		if fileval == 0:
			print(f"no valid files found in folder: F{CurrentWeek}")
			exit()
		else:
			print(f"File: {toReturn}")
			return toReturn
	
def parseWorkSheet(location, sheet):
	xls = pd.ExcelFile(location)
	print(xls.sheet_names[sheet])
	df = pd.read_excel(location, sheet_name=xls.sheet_names[sheet], header=None)
	
	rows = df.iterrows()

	for index, row in rows:
		if not pd.isna(row[0]): #Is there a message code
			if not pd.isna(row[1]) or not pd.isna(row[2]) or not pd.isna(row[3]): #Multicolumn headers won't have data for [1-3]
				mCodes.append(row[0])	
				print(row[0])
	

def checkScript():
	if not checkScriptExists():
		createScript()
	else:
		validateScript()

def createScript():
	print(f"Script file for Current Week {fileCurrentWeek} don't exist.")
	
	source=F"{pPath}\MyCD Scripts.xlsx"
	dest=scriptFile
	
	
	shutil.copy(source,dest)
	print(f"Created: W{CurrentWeek} MyCD Scripts.xlsx")
	
	df = pd.read_excel(scriptFile, sheet_name='CI Targeting')
	wb = openpyxl.load_workbook(dest)
	
	i = 0
	while i < len(mCodes):
		wb['CI Targeting'].cell(column=1, row=i+2, value=mCodes[i])
		i += 1
	wb.save(dest)
		
def validateScript():
	df = pd.read_excel(scriptFile, sheet_name='CI Targeting')
	
	codes = [x for x in df['Message Code'].to_numpy() if x == x]
	
	valid = True
	for val in mCodes:
		if val == val:
			if val not in codes:
				print(val+' not in MyCD Scripts')
				valid = False
				
	for val in codes:
		if val not in mCodes:
			print(val+' in MyCD Scripts but not targetting sheet')
			valid =False
	
	print('No issues found.') if valid else print('Issues found, please resolve before continuing with myCD targetting.')
	
def checkScriptExists():
	return os.path.exists(scriptFile)

def startBitmark():
	print('Checking for bitmarks')
	
	if not checkScriptExists():
		print('myCD Script file missing')
	else:
		df = pd.read_excel(location, sheet_name='CI Targeting')
		codes = [x for x in df('Logic').to_numpy() if x == x]
		
		if not codes:
			print('No bitmark logic inserted.')
		else :
			for val in codes:
				print(val)

#Functions done, start here
initCurrentWeekData()
fp.test()

while True:
	try:		
		data = int(input("Enter choice. \n1: Parse myCD Worksheet\n2: Parse Bitmarks from Targeting\n3: Quit\n>: "), 10)
		if data not in [1, 2, 3]:
			print("Enter valid choice (1-3)")
			print(data)
		else:
			if data == 1:
				StartSheet()
				#ParseSheet()
			elif data == 2:
				startBitmark()
			else:
				print("Quitting")
				break
	except Exception:
		print("Enter valid choice (1-3)")

