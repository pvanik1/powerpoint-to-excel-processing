######################## EXTRACT DATA FROM PPTX TO XLSX #########################
# Created by Peter Vaník, April 2018
'''
Extracts data from pre-defined TAT template and outputs them into an excel in specified format.
Usable for TRP and GSTP activities only.
'''

import textract
import xlsxwriter
from tkinter.filedialog import askopenfilenames
import unicodedata
import tkinter
import re
from tkinter import messagebox

############################ PREPARE WORKBOOK ###################################
# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('web_excel_input.xlsx')
w = workbook.add_worksheet()
# Widen the first column to make the text clearer.
w.set_column('Q:Q', 30)
# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True})
# Add a yellow background colour format.
yellow = workbook.add_format()
yellow.set_pattern(1)
yellow.set_bg_color('yellow')

############################ PROCESS POWERPOINT FILE ############################
thereAreUnindentifiedTATs = False
filenames = askopenfilenames()

for i in range(len(filenames)):
	print (filenames[i])
	problemParsingAttribute = False

	text = textract.process(filenames[i], extension="pptx", encoding="utf_8")
	with open('output.txt', 'wb') as f:
	    f.write(text)
	with open('output.txt','r', encoding = "utf-8") as f:
		text = f.read()

	text = str(text).replace("\n","")
	text = str(text).replace("Contractor(s):","Contractors:")
	text = str(text).replace("Current:","Achieved:")
	text = str(text).replace(")",") ")
	text = str(text).replace("Co-funded Budget:","")
	text = str(text).replace("ESA Budget"," ESA Budget")

	#for debugging
	with open('output.txt', 'w', encoding = "utf-8") as f:
	    f.write(text)


	# PROGRAM REFERENCE
	isTRP = False
	isGSTP = False
	gstp_pattern = re.compile('G\d\d.-.*') # Regular expression matching GSTP reference
	gstp_pattern2 = re.compile('A.*-\d\d\.*') # Regular expression matching GSTP reference
	trp_pattern = re.compile('T\d\d\d-.*') # Regular expression matching TRP reference
	for word in text.split():
		if ((gstp_pattern.match(word) is not None) or (gstp_pattern2.match(word) is not None)):
			progRef = word
			print("Programme Reference: " + progRef)
			isGSTP = True
			break
		elif trp_pattern.match(word) is not None:
			progRef = word
			print("Programme Reference: " + progRef)
			isTRP = True
			break

	if isTRP:
		trpIndex = text.find(progRef)
		if "TRP" in text[trpIndex-10:trpIndex]:
			trpIndex = text.find("TRP")
	elif isGSTP:
		gstpIndex = text.find(progRef)
		if "GSTP" in text[gstpIndex-10:gstpIndex]:
			gstpIndex = text.find("GSTP")
	else:
		progRef = "UNPARSED"
		print("---------- Program reference could not be parsed ----------")
		problemParsingAttribute = True
	

	# TARGET TRL
	startTargetTRL = text.find("Target TRL:")
	if (startTargetTRL < 0):
		print("---------- Target TRL could not be parsed ----------")
		targetTRL = "UNPARSED"
		problemParsingAttribute = True
	else:
		indexTargetTRL = startTargetTRL+len("Target TRL:")
		endTargetTRL = text.find("Date:")
		targetTRL = text[indexTargetTRL:endTargetTRL].lstrip(' ')
		targetTRL = targetTRL.strip(' \t\n\r')
		print("Target TRL: " + targetTRL)


	# CONTRACTORS
	contractorsParsed = True
	if (text.find("Contractors:") < 0):
		contractorsParsed = False
	else:
		indexContractors = text.find("Contractors:") + len("Contractors:")
		if isTRP:
			endContractors = trpIndex
		elif isGSTP:
			endContractors = gstpIndex
		elif text[:300].find("GSTP") > -1:
			endContractors = text[:300].find("GSTP")
		elif text[:300].find("TRP") > -1:
			endContractors = text[:300].find("TRP")
		else:
			contractorsParsed = False
	
	if contractorsParsed == False:
		print("---------- Contractors could not be parsed ----------")
		contractors = "UNPARSED"
		problemParsingAttribute = True
	else:
		contractors = text[indexContractors : endContractors]
		contractors = contractors.replace('\n','').strip(' ')
		print ("Contractors: " + contractors)


	# BUDGET
	indexBudget = text.find("ESA Budget:")
	if (indexBudget < 0):
		print("---------- Budget could not be parsed ----------")
		budget = "UNPARSED"
		problemParsingAttribute = True
	else:
		indexBudget += len("ESA Budget:")
		endBudget = text.find("k",indexBudget)
		budget = text[indexBudget : endBudget]
		budget = budget.strip(' \t\n\r')
		print("Budget (k€): " + budget)


	# YEAR OF COMPLETION
	indexYoc = text.find("YoC")
	indexInitialTRL = text.find("Initial:")
	if(indexYoc < 0):
		print("---------- YoC could not be parsed ----------")
		yoc = "UNPARSED"
		problemParsingAttribute = True
	elif(indexInitialTRL < 0):
		print("---------- Initial TRL could not be parsed ----------")
		initialTRL = "UNPARSED"
		problemParsingAttribute = True
	else:
		yoc = text[indexYoc  + len("Yoc:"): indexInitialTRL].replace(":","")
		yoc = yoc.strip(' \t\n\r')
		print("YoC: " + yoc)


	# INITIAL TRL
	indexTo = text.find("TO:")
	if (indexTo < 0):
		print("---------- TO could not be parsed ----------")
		TO = "UNPARSED"
		problemParsingAttribute = True
	elif (indexInitialTRL > -1):
		initialTRL = text[indexInitialTRL + len("Initial:") : indexTo].lstrip(' ')
		initialTRL = initialTRL.strip(' \t\n\r')
		print("Initial TRL: " + initialTRL)


	# ACHIEVED TRL
	indexAchievedTRL = text.find("Achieved:")
	if(indexAchievedTRL < 0):
		print("---------- Achieved TRL could not be parsed ----------")
		achievedTRL = "UNPARSED"
		problemParsingAttribute = True
	else:
		achievedTRL = text[indexAchievedTRL + len("Achieved:") : indexYoc].lstrip(' ')
		achievedTRL = achievedTRL.strip(' \t\n\r')
		print("Achieved TRL: " + achievedTRL)


	# DATE
	indexDate = text.find("Date:")
	if (indexDate < 0):
		print("---------- Date could not be parsed ----------")
		achievedTRL = "UNPARSED"
		problemParsingAttribute = True
	else:
		endIndexDate = text[indexDate:].find("TRL")
		date = text[indexDate + len("Date:") : indexDate + endIndexDate].lstrip(' ')
		date = date.strip(' \t\n\r')
		print ("Date: " + date)


	# TO WITH SECTION
	endIndexToWithSection = text.find( ")" , indexTo)
	if (endIndexToWithSection < 0):
		endIndexToWithSection = len(text)-1
	toWithSection = text[indexTo + len("TO:") : endIndexToWithSection+1].lstrip(' ')

	try:
		toUnwrap = toWithSection.split()
		if (len(toUnwrap) == 2):
			toSection = toUnwrap [-1]
			temp = toUnwrap[0].split('.')
			toName = temp[0] + '.'
			toSurname = temp[1]

		else:
			toSection = toUnwrap [-1]
			del toUnwrap [-1]
			toSurname = toUnwrap [-1]
			del toUnwrap [-1]
			if (toUnwrap [-1] in ["de", "da", "di"]):
				temp = toSurname
				toSurname = toUnwrap[-1] + " " + temp
				del toUnwrap [-1]
			toName = ""
			for item in toUnwrap:
				toName = toName + " " + item
	except:
		toName = toWithSection
		toSurname = ""
		toSection = ""

	toName = toName.strip(' \t\n\r')
	toSurname = toSurname.strip(' \t\n\r')
	toSection = toSection.strip(' \t\n\r')
	print("TO name: " + toName)
	print("TO surname: " + toSurname)
	print("TO section: " + toSection)


	# COUNTRY ORIGIN
	def getCountryFromContractors(input, delimiter):
		'''Outputs the origin countries separated by delimiter.'''
		countryList = []
		for word in input.split():
			if word[0]==("("):
				country = word.lstrip('(').rstrip('),')
				if country not in countryList:
					countryList.append(country)
		output = ""
		for country in countryList:
			output += country + " " + delimiter + " "
		output = output.rstrip(" " + delimiter + " ")
		return output

	# TD,SD
	TD = ""
	SD = ""
	if isTRP:
		TD = int(progRef[2:4])
		SD_num = int(progRef[1])
		switcher = {
			1:"EO",
			2:"SCI",
			3:"EXP",
			4:"ST",
			5:"TEL",
			6:"NAV",
			7:"GEN"
		}
		SD = switcher.get(SD_num)
		print("SD: " + str(SD))
		print("TD: " + str(TD))
	print('\n')


	# Formatting for web statistics excel file
	def formatIfUnparsed(parameter):
		if (parameter == "UNPARSED" or parameter == ""):
			return yellow

	w.write('I1', "Programme reference", bold)
	w.write('I' + str(2+i), progRef, formatIfUnparsed(progRef))
	w.write('K1', "Initial TRL", bold)
	w.write('K' + str(2+i), initialTRL, formatIfUnparsed(initialTRL))
	w.write('L1', "Achieved TRL", bold)
	w.write('L' + str(2+i), achievedTRL, formatIfUnparsed(achievedTRL))
	w.write('M1', "Target TRL", bold)
	w.write('M' + str(2+i), targetTRL, formatIfUnparsed(targetTRL))
	w.write('N1', "Budget (k€)", bold)
	w.write('N' + str(2+i), budget, formatIfUnparsed(budget))
	w.write('P1', "YoC", bold)
	w.write('P' + str(2+i), yoc, formatIfUnparsed(yoc))
	w.write('Q1', "Contractors", bold)
	w.write('Q' + str(2+i), contractors, formatIfUnparsed(contractors))
	w.write('R1', "Country origin", bold)
	w.write('R' + str(2+i), getCountryFromContractors(contractors, "|"), formatIfUnparsed(contractors))
	w.write('S1', "SD", bold)
	w.write('S' + str(2+i), SD)
	w.write('T1', "TD", bold)
	w.write('T' + str(2+i), TD)
	w.write('V1', "TO", bold)
	w.write('V' + str(2+i), toWithSection, formatIfUnparsed(toWithSection))
	w.write('X1', "SETFPDs", bold)
	w.write('X' + str(2+i), 'SETFPDs November 2017')
	w.write('Y1', "Video published", bold)
	w.write('Y' + str(2+i), '1')
	w.write('Z1', "TAT saved", bold)
	w.write('Z' + str(2+i), '1')

	if (problemParsingAttribute == True):
		indexTATname = filenames[i].rfind('/')
		TATname = filenames[i][indexTATname+1:]
		if (thereAreUnindentifiedTATs == False):
			with open('unparsedTATs.txt', 'w') as g:
				g.write(TATname + '\n')
		else:
			with open('unparsedTATs.txt', 'a') as g:
				g.write(TATname + '\n')
		thereAreUnindentifiedTATs = True

workbook.close()

if thereAreUnindentifiedTATs:
	messagebox.showinfo(
		"Error parsing TATs", 
		"One or more TATs were either not belonging to TRP/GSTP or filled out in a non-standard way and could not be parsed. Please check them manually. You can find the list of these TATs in 'unparsedTATs.txt'")