##################### CONVERT POWERPOINT TO TEXT #####################
import textract
import xlsxwriter
from tkinter.filedialog import askopenfilename
import unicodedata

filename = askopenfilename()
print(filename)

text = textract.process(filename, extension="pptx", encoding="utf_8")
with open('output.txt', 'wb') as f:
    f.write(text)
with open('output.txt','r', encoding = "utf-8") as f:
	text = f.read()

text = str(text).replace("Contractor(s):","Contractors:")
text = str(text).replace("Current:","Achieved:")


# TARGET TRL
startTargetTRL = text.find("Target TRL:")
indexTargetTRL = startTargetTRL+len("Target TRL:")
targetTRL = text[indexTargetTRL:indexTargetTRL+2].lstrip(' ')
targetTRL = targetTRL.strip(' \t\n\r')
print("Target TRL: " + targetTRL)

# Identify if TRP or GSTP
trpIndex = text[:1000].find("TRP")
if trpIndex < 0:
	gstpIndex = text[:1000].find("GSTP")
	isTRP = False
	isGSTP = True
else:
	isTRP = True
	isGSTP = False


# CONTRACTORS
indexContractors = text.find("Contractors:") + len("Contractors:")
if isTRP:
	endContractors = trpIndex
else: 
	endContractors = gstpIndex
contractors = text[indexContractors : endContractors]
contractors = contractors.replace('\n','').strip(' ')
print ("Contractors: " + contractors)


# BUDGET
indexBudget = text.find("ESA Budget:") + len("ESA Budget:")
endBudget = text.find("k",indexBudget)
budget = text[indexBudget : endBudget]
budget = budget.strip(' \t\n\r')
print("Budget: " + budget)


# PROGRAMME REFERENCE
endProgRef = text.find("ESA Budget:")
if isTRP:
	progRef = text[trpIndex + len("TRP ") : endProgRef]
elif isGSTP:
	progRef = text[gstpIndex + len("GSTP ") : endProgRef]
progRef = progRef.strip(' \t\n\r')
print("Programme Reference: " + progRef)


# YEAR OF COMPLETION
indexYoc = text.find("YoC")
indexInitialTRL = text.find("Initial:")
yoc = text[indexYoc  + len("Yoc:"): indexInitialTRL].replace(":","")
yoc = yoc.strip(' \t\n\r')
print("YoC: " + yoc)


# INITIAL TRL
indexTo = text.find("TO:")
initialTRL = text[indexInitialTRL + len("Initial:") : indexTo].lstrip(' ')
initialTRL = initialTRL.strip(' \t\n\r')
print("Initial TRL: " + initialTRL)


# ACHIEVED TRL
indexAchievedTRL = text.find("Achieved:")
achievedTRL = text[indexAchievedTRL + len("Achieved:") : indexYoc].lstrip(' ')
achievedTRL = achievedTRL.strip(' \t\n\r')
print("Achieved TRL: " + achievedTRL)


# DATE
indexDate = text.find("Date:")
endIndexDate = text[indexDate:].find("TRL")
date = text[indexDate + len("Date:") : indexDate + endIndexDate].lstrip(' ')
date = date.strip(' \t\n\r')
print ("Date: " + date)


# TO WITH SECTION
endIndexToWithSection = text.find( ")" , indexTo)
toWithSection = text[indexTo + len("TO:") : endIndexToWithSection+1].lstrip(' ')

toName,toSurname,toSection = toWithSection.split(" ")
toName = toName.strip(' \t\n\r')
toSurname = toSurname.strip(' \t\n\r')
toSection = toSection.strip(' \t\n\r')
print("TO name: " + toName)
print("TO surname: " + toSurname)
print("TO section: " + toSection)


# COUNTRY ORIGIN
def getCountryFromContractors(input, delimiter):
	"Outputs the origin countries separated by delimiter."
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


############################ WRITE IT TO EXCEL ###################################
# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('web_excel_input.xlsx')
w = workbook.add_worksheet()
# Widen the first column to make the text clearer.
w.set_column('Q:Q', 30)
# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True})

# Formatting for web statistics excel file
w.write('K1', "Initial TRL", bold)
w.write('K2', initialTRL)
w.write('L1', "Achieved TRL", bold)
w.write('L2', achievedTRL)
w.write('M1', "Target TRL", bold)
w.write('M2', targetTRL)
w.write('N1', "Budget (k€)", bold)
w.write('N2', budget)
w.write('P1', "YoC", bold)
w.write('P2', yoc)
w.write('Q1', "Contractors", bold)
w.write('Q2', contractors)
w.write('R1', "Country origin", bold)
w.write('R2', getCountryFromContractors(contractors, "|"))
w.write('V1', "TO", bold)
w.write('V2', toSurname + ", " + toName[0] + ".")

# TODO TD SD

workbook.close()