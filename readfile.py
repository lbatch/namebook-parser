import xlwt
from xlwt import Workbook
import xlrd
import io

# define dictionary of states to compare text against
states = ["ALABAMA", "ALASKA", "ARIZONA", "ARKANSAS", "CALIFORNIA", "COLORADO", "CONNECTICUT", "DELAWARE", "FLORIDA",
    "GEORGIA", "HAWAII", "IDAHO", "ILLINOIS", "INDIANA", "IOWA", "KANSAS", "KENTUCKY", "LOUISIANA", "MAINE", "MARYLAND",
    "MASSACHUSETTS" "MICHIGAN", "MINNESOTA", "MISSISSIPPI", "MISSOURI", "MONTANA", "NEBRASKA", "NEVADA", "NEW HAMPSHIRE",
    "NEW JERSEY", "NEW MEXICO", "NEW YORK", "NORTH CAROLINA", "NORTH DAKOTA", "OHIO", "OKLAHOMA", "OREGON", "PENNSYLVANIA",
    "RHODE ISLAND", "SOUTH CAROLINA", "SOUTH DAKOTA", "TENNESSEE", "TEXAS", "UTAH", "VERMONT", "VIRGINIA", "WASHINGTON",
    "WEST VIRGINIA", "WISCONSIN", "WYOMING", "DISTRICT OF COLUMBIA", "GUAM", "NORTHERN MARIANA ISLANDS", "PUERTO RICO", "VIRGIN ISLANDS",
    "AUSTRALIA", "BELGIUM", "CANADA", "CHINA", "ENGLAND", "ETHIOPIA", "FRANCE", "GERMANY", "INDIA", "ITALY", "JAPAN", "JORDAN", "KOREA, SOUTH (ROK)",
    "MALAYSIA", "NIGERIA", "PHILLIPINES", "QATAR", "RUSSIA", "SINGAPORE", "SWITZERLAND", "TAIWAN", "THAILAND", "TURKEY", "UNITED ARAB EMIRATES",
    "VIETNAM", "WALES"]

# create a dictionary of schools with ceeb codes to check text against
ceebdict = {}
ceebfile = xlrd.open_workbook("ceeb-lookup.xlsx") # file populated with high school names, cities, states, and corresponding CEEB codes
sheet = ceebfile.sheet_by_index(0) # select first sheet in file
inLineIdx = 0

# for each row in the file of schools, populate the dictionary at state -> city -> school name with the ceeb code of that school
for row in range(sheet.nrows):
    if(sheet.cell(inLineIdx,3).value in states):
        ceebState = sheet.cell(inLineIdx,3).value
        ceebCity = sheet.cell(inLineIdx,2).value
        ceebName = sheet.cell(inLineIdx,0).value
        ceebCeeb = sheet.cell(inLineIdx,1).value
        try: ceebdict[ceebState]
        except: ceebdict[ceebState] = {}
        try: ceebdict[ceebState][ceebCity]
        except: ceebdict[ceebState][ceebCity] = {}
        ceebdict[ceebState][ceebCity][ceebName] = ceebCeeb
    inLineIdx += 1

# function to write the name of a student, parsed appropriately from given text, into the final spreadsheet
def writeName(name, lineno):
    nameTokens = name.split(' ')
    number = ""
    last = ""
    first = ""
    middle = ""
    ceeb = ""
    phase = 0
    for token in nameTokens:
        if(phase == 0): # pull number from beginning of line; each name preceded by a 3-digit nuber
            number = token
            phase = 1
        elif(phase == 1): # add tokens to last name until comma encountered
            if(token.endswith(',')): # handle typical case when OCR read correctly; last name is followed by comma, then a space before first name
                last = last + token[:-1]
                phase = 2
            elif(',' in token): # handle case when OCR did not read space between comma and beginning of first name
                lastFirst = token.split(",")
                last = last + lastFirst[0]
                first = first + lastFirst[1]
                phase = 2
            else: # handle tokens which are in but do not end the last name (that is, non-terminal tokens in a multi-word last name)
                last = last + token
        elif(phase == 2): # add to first name until middle initial, characterized by '.' at end of single-character token, is found
            if(token.endswith('.') and len(token) == 2):
                middle = middle + token[:-1]
            else:
                first = first + token
    sheet1.write(lineno,0,number) # write number and name to file; writing number even though it is not ultimately used to catch cases where whitespace between number and last name was not read by OCR
    sheet1.write(lineno,1,last)
    sheet1.write(lineno,2,first)
    sheet1.write(lineno,3,middle)

# function to determine a school's ceeb code and write to file; uses ceeb dictionary to identify school iff state, city, and name are a match
def writeCeeb(lookupState, lookupCity, lookupSchool, lineno):
    try:
        toWrite = statedict[lookupState][lookupCity][lookupSchool]
        sheet1.write(lineno,7,toWrite)
    except: # identify unidentified schools for manual lookup
        sheet1.write(lineno,7,"N/A")


# read text file and write to spreadsheet

# variable setup
rowLine = 1
cityLine = 0
citySchool = ""
city = ""
school = ""
state = ""

# add sheet and column headers to workbook file
wb = Workbook()
sheet1 = wb.add_sheet('Sheet 1')
sheet1.write(0,0, "Number")
sheet1.write(0,1, "Last")
sheet1.write(0,2, "First")
sheet1.write(0,3, "MI")
sheet1.write(0,4, "State")
sheet1.write(0,5, "City")
sheet1.write(0,6, "School")
sheet1.write(0,7, "CEEB")


# open text file read via OCR and iterate through rows
with io.open("nm-book-text.txt", "r", encoding="utf-8") as bookText:
    for line in bookText:
        line = line.strip()
        if(line and ("(continued)" not in line)): # eliminate extra "<statename> (continued)" rows
            if(line in states): # if the line identifies a state, identify the state for subsequent students as that state
                cityLine = 0
                state = line 
            elif(line[0].isdigit()): # if the line begins with a number, identify as a student and write to file
		# identifies that the previous line was a school name, since current line is a student name
                if(cityLine == 1): 
                    school = citySchool
                    cityLine = 0
                writeName(line, rowLine) # call function to parse name row and write to file
                sheet1.write(rowLine,4,state) # write state and city to file
                sheet1.write(rowLine,5,city)
		
		# replace common permutations of "HIGH SCHOOL" in file for consistent lookup in the ceeb dictionary
                if("HS" in school):
                    school = school.replace("HS", "HIGH SCHOOL")
                elif("H. S." in school):
                    school = school.replace("H. S.", "HIGH SCHOOL")
                elif("H.S." in school):
                    school = school.replace("H.S.", "HIGH SCHOOL")
                sheet1.write(rowLine,6,school)

		# call function to determine CEEB and write to file
                writeCeeb(state, city, school, rowLine)
                rowLine += 1

	    # identifies that the previous line was a city name, current line is a school name
            elif(cityLine == 1):
                city = citySchool
                school = line
                cityLine = 0   
	    
	    # identifies that current line is either a city or a school, to be determined on next iteration of loop (checked for per cityLine flag)
            else:
                cityLine = 1
                citySchool = line

# save workbook to file
wb.save("Semifinalists2019.xls")
