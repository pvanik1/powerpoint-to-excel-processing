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
import os

############################ PREPARE WORKBOOK ###################################
# Create a new folder and change directory to it
dirpath = os.getcwd()
newdir = dirpath + "\\Outputs" 
os.makedirs(newdir,exist_ok=True)
os.chdir(newdir)

# Create a new Excel file and add a worksheet and check the file isn't already open so that the program can run properly.
workbook = xlsxwriter.Workbook('TAT_data.xlsx')
try:
	workbook.close()
except PermissionError:
	messagebox.showinfo("Error: Close the workbook and try again", 
						"The program cannot run while TAT_data.xlsx is open." + '\n' + '\n' + "Please close the file and re-run the program.")
	print("The program cannot run while the TAT_output.xlsx file is open. Please close the file and re-run the program.")
	quit()

workbook = xlsxwriter.Workbook('TAT_data.xlsx')
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

	text = textract.process(filenames[i], encoding="utf_8")
	with open('output.txt', 'wb') as f:
	    f.write(text)
	with open('output.txt','r', encoding = "utf-8") as f:
		text = f.read()

	# Pre-processing text
	text = str(text).replace("\n","")
	text = str(text).replace("Contractor(s):","Contractors:")
	text = str(text).replace("Prime:","Contractors:")
	text = str(text).replace("Contractor:","Contractors:")
	text = str(text).replace("Contractor :","Contractors:")
	text = str(text).replace("Current:","Achieved:")
	text = str(text).replace(")",") ")
	text = str(text).replace("Co-funded Budget:","")
	text = str(text).replace("ESA Budget"," ESA Budget")
	text = str(text).replace("Contractors :","Contractors:")
	text = str(text).replace("Programme & Reference :","")
	text = str(text).replace("Objective:","Objective(s) : ")
	text = str(text).replace("Objectives:","Objective(s) : ")
	text = str(text).replace("Next Steps:","Next steps:")
	text = str(text).replace("Background:","Background and justification:")
	# Next steps: replace capital	Next Steps:

	#for debugging
	with open('output.txt', 'w', encoding = "utf-8") as f:
	    f.write(text)

	# PROGRAM REFERENCE
	# First checks in-file, then in file name because there is no guarantee filename is correct
	isTRP = False
	isGSTP = False
	trpIndex = -1
	gstpIndex = -1
	progRef = "UNPARSED"
	gstp_pattern = re.compile('G\d\d.-.*') # Regular expression matching GSTP reference
	gstp_pattern2 = re.compile('A.*-\d\d\.*') # Regular expression matching GSTP reference
	trp_pattern = re.compile('T\d\d\d-.*') # Regular expression matching TRP reference
	for word in text.split():
		if ((gstp_pattern.match(word) is not None) or (gstp_pattern2.match(word) is not None)):
			progRef = word
			print("Programme Reference: " + progRef)
			isGSTP = True
			break
		elif (trp_pattern.match(word) is not None):
			progRef = word
			print("Programme Reference: " + progRef)
			isTRP = True
			print(" Y OF OF FO JFOI AFI AEOIF NAOENF AE FNEOIN AOF NEAO FIO ANOIE F")
			break
		else:
			indexFilenameTRP = filenames[i].find("TRP")
			indexFilenameGSTP = filenames[i].find("GSTP")
			if (indexFilenameTRP >= 0):
				temp = filenames[i][indexFilenameTRP + len("TRP_"):]
				temp = temp.replace("–","-")
				temp = temp.replace("_","-")
				temp2 = temp.split("-",2)[:2]
				progRef = "-".join(temp2).rstrip(' ')
				print("Programme Reference: " + progRef)
				isTRP = True
				break				
				
			elif (indexFilenameGSTP >= 0):
				temp = filenames[i][indexFilenameGSTP + len("GSTP_"):]
				temp = temp.replace("–","-")
				temp = temp.replace("_","-")
				temp2 = temp.split("-",2)[:2]
				progRef = "-".join(temp2).rstrip(' ')
				print("Programme Reference: " + progRef)
				isGSTP = True
				break
				
			else:
				progRef = "UNPARSED"
				print("---------- Programme reference could not be parsed ----------")
				problemParsingAttribute = True
				break

	if (progRef.find("XXX") >= 0):		# check for if someone forgot a bad reference in the file name / missing reference
		progRef = "UNPARSED"

	if isTRP:							# contractor parsing code requires the index of "TRP" or "GSTP", if they exist with progRef
		trpIndex = text.find(progRef)
		if "TRP" in text[trpIndex-50:trpIndex]:	
			trpIndex = text.find("TRP")
	elif isGSTP:
		gstpIndex = text.find(progRef)
		if "GSTP" in text[gstpIndex-10:gstpIndex]:
			gstpIndex = text.find("GSTP")

	# TARGET TRL
	startTargetTRL = text.find("Target TRL:")
	indexTargetTRL = startTargetTRL+len("Target TRL:")
	if (startTargetTRL < 0):
		print("---------- Target TRL could not be parsed ----------")
		targetTRL = "UNPARSED"
		problemParsingAttribute = True
	# if Date cannot be found, only year
	else:
		if text.find("Date:") < 0:
			reNumber = re.compile('\d') # Regular expression matching a single digit
			for character in text[indexTargetTRL:]:
				if reNumber.match(character) is not None:
					endTargetTRL = indexTargetTRL + text[indexTargetTRL:].find(character)
					break
			targetTRL = character
			print("Target TRL: " + targetTRL)	
		else:
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
		if (trpIndex > -1):
			endContractors = trpIndex
		elif (gstpIndex > -1):
			endContractors = gstpIndex
		elif text[:300].find("GSTP") > -1:
			endContractors = text[:300].find("GSTP")
		elif text[:300].find("TRP") > -1:
			endContractors = text[:300].find("TRP")
		# elif text[:300].find("TEC-") > -1:
		# 	endContractors = text[:300].find("TEC-")
		# elif text[:300].find("ESA/ITT") > -1:
		# 	endContractors = text[:300].find("ESA/ITT")
		# elif text[:400].find("4000") > -1:
		# 	endContractors = text[:400].find("4000")
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


	# BACKGROUND AND JUSTIFICATION
	backgroundAndJustification = "UNPARSED"
	indexBackground = text.find("Background and justification:")
	indexObjectives = text.find("Objective(s) :")
	if (indexObjectives < 0):
		indexObjectives = text.find("Objective(s)")
	if ((indexBackground < 0) or (indexObjectives < 0)):
		print("---------- Background could not be parsed ----------") #xxxx
		print (" INDEX BACKGROUND " + str(indexBackground))
		problemParsingAttribute = True
	else:
		backgroundAndJustification = text[indexBackground + len("Background and justification:") : indexObjectives]


	# BUDGET
	indexBudget = text.find("ESA Budget:")
	if (indexBudget < 0):
		print("---------- Budget could not be parsed ----------")
		budget = "UNPARSED"
		problemParsingAttribute = True
	else:
		indexBudget += len("ESA Budget:")
		if (indexBackground >= 0):
			endBudget = indexBackground
		else:
			endBudget = text.find("k",indexBudget)
		# If the euro symbol is in front of the number
		budgetWords = text[indexBudget : endBudget].split(" ")
		if (budgetWords[0] == "€"):
			budget = budgetWords[1]
		else:
			budget = budgetWords[0]
		# Process the string for input errors
		budget = budget.strip(' \t\n\r')
		budget = budget.strip('€')
		wrong_pattern = re.compile('.*\d\d,\d\d\d')
		if wrong_pattern.match(budget) is not None:
			budget = budget[:-4]
		if (str(budget).find("k") >= 0):
			budget=budget[:str(budget).find("k")]
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
	initialTRL = "UNPARSED"
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
	if isTRP and trp_pattern.match(progRef) is not None:
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


	# FOLLOW-UP AND NEXT STEPS
	followupKeyWords = ("followup", "follow", "follow-up", "Follow-up", "Follow", "Followup", "Follow-on","follow-on","qualification", "ITT" )
	programsAndMissions = ("TRP","GSTP","GSP","NAVISP","ARTES", "HERA","ARIEL","MSR", "Mars", "Sample", "JUICE","Juice","Athena","Plato","Flex", "Galileo","MTG",
		"national", "Horizon2020", "H2020", "Horizon", "EOEP")
	followup = []
	hasFollowUp = False
	nextSteps = "UNPARSED"

	indexNextSteps = text.find("Next steps:")
	if ((startTargetTRL < 0) or (startTargetTRL < indexNextSteps)): 
		endIndexNextSteps = len(text)
	else:
		endIndexNextSteps = startTargetTRL
	if (indexNextSteps < 0):
		print("---------- Next Steps could not be parsed ----------")
		problemParsingAttribute = True
	else:
		nextSteps = text[indexNextSteps + len("Next steps: ") : endIndexNextSteps]
		for word in nextSteps.split():
			if (word in followupKeyWords) or (word in programsAndMissions):
				hasFollowUp = True
				if word in programsAndMissions:
					followup.append(word)


	# PRINT BACKGROUND & JUSTIFICATION
	print("Background and justification: " + backgroundAndJustification + '\n')


	# OBJECTIVES
	objectives = "UNPARSED"
	indexAchievements = text.find("Achievements and status:")
	if ((indexObjectives < 0) or (indexAchievements < 0)):
		print("---------- Objectives could not be parsed ----------")
		problemParsingAttribute = True
	else:
		objectives = text[indexObjectives + len("Objective(s) :") : indexAchievements]
		print("Objectives: " + objectives + '\n')


	# ACHIEVEMENTS AND STATUS
	achievementsAndStatus = "UNPARSED"
	indexBenefits = text.find("Benefits:")
	if ((indexAchievements < 0) or (indexBenefits < 0)):
		print("---------- Achievements could not be parsed ----------")
		problemParsingAttribute = True
	else:
		achievementsAndStatus = text[indexAchievements + len("Achievements and status: ") : indexBenefits]
		print("Achievements and status: " + achievementsAndStatus + '\n')


	# BENEFITS
	benefits = "UNPARSED"
	if ((indexBenefits < 0) or (indexNextSteps < 0)):
		print("---------- Benefits could not be parsed ----------")
		problemParsingAttribute = True
	else:
		benefits = text[indexBenefits + len("Benefits:") : indexNextSteps]
		print("INDEX BENEFITS " + str(indexBenefits))
		print("INDEX NEXT " + str(indexNextSteps))
		print("Benefits: " + benefits + '\n')


	# NEXT STEPS PRINT
	if (nextSteps != "UNPARSED"):
		print("Next steps: " + nextSteps + '\n')


	# Formatting for web statistics excel file
	def formatIfUnparsed(parameter):
		if (parameter == "UNPARSED" or parameter == ""):
			return yellow

	indexTATname = filenames[i].rfind('/')
	TATname = filenames[i][indexTATname+1:]

	w.write('A1', "File name", bold)
	w.write('A' + str(2+i), TATname)
	w.write('B1', "Follow-up?", bold)
	w.write('B' + str(2+i), str(hasFollowUp))
	w.write('C1', "Follow-up type", bold)
	w.write('C' + str(2+i), str(followup))

	w.write('D1', "Background and justification", bold)
	w.write('D' + str(2+i), backgroundAndJustification, formatIfUnparsed(backgroundAndJustification))
	w.write('E1', "Objectives", bold)
	w.write('E' + str(2+i), objectives, formatIfUnparsed(objectives))
	w.write('F1', "Achievements and status", bold)
	w.write('F' + str(2+i), achievementsAndStatus, formatIfUnparsed(achievementsAndStatus))
	w.write('G1', "Benefits", bold)
	w.write('G' + str(2+i), benefits, formatIfUnparsed(benefits))
	w.write('H1', "Next steps: ", bold)
	w.write('H' + str(2+i), nextSteps, formatIfUnparsed(nextSteps))

	w.write('I1', "Programme reference", bold)
	w.write('I' + str(2+i), progRef, formatIfUnparsed(progRef))
	w.write('K1', "Initial TRL", bold)
	w.write('K' + str(2+i), initialTRL, formatIfUnparsed(initialTRL))
	w.write('L1', "Achieved TRL", bold)
	w.write('L' + str(2+i), achievedTRL, formatIfUnparsed(achievedTRL))
	w.write('M1', "Target TRL", bold)
	w.write('M' + str(2+i), targetTRL, formatIfUnparsed(targetTRL))
	w.write('N1', "ESA Budget (k€)", bold)
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
		if (thereAreUnindentifiedTATs == False):
			with open('unparsedTATs.txt', 'w') as g:
				g.write(TATname + '\n')
		else:
			with open('unparsedTATs.txt', 'a') as g:
				g.write(TATname + '\n')
		thereAreUnindentifiedTATs = True

workbook.close()

# Warn the user that some TATs have not been parsed fully
if thereAreUnindentifiedTATs:
	messagebox.showinfo(
		"Error parsing TATs", 
		"One or more TATs were either not belonging to TRP/GSTP or filled out in a non-standard way and could not be parsed. Please check them manually." + '\n' + '\n' + "You can find the list of these TATs in 'unparsedTATs.txt' as well as in the Excel output file.")

# Open the Excel output file
os.startfile(newdir + '\\'+ 'TAT_data.xlsx')