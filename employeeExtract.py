from xml.dom import minidom #For parsing the XML (ElementTree had issues parsing this type of doc)
from datetime import datetime #For formatting dates
from dateutil import relativedelta #for handling difference in dates
from itertools import tee, islice, chain
import csv #for writing to CSV

mydoc = minidom.parse('WOP-AllWithCourses.xml')

section = mydoc.getElementsByTagName('Detail')

# print(table[0].attribute['detail_description'].value)

# deatilCollection = section[0].getElementsByTagName('Detail')[0].attributes['DetailField_FullName_Section_1'].value

# deatilCollection = type(section[0].getElementsByTagName('Detail'))


#object methods needed
# - to deal with date formats
# 	- Write funciton to calculate years of expierence date
# - Calculate the right licnese to use
# - to accept multiple courses and certs
# - Improvement: Take list of their projects and determine what title has been used most and apply that one
# implement an optional y/n boolean for a prefferred name or a middle name situation. IMplement logic to resolve to the right name
# Split name into first last and middle - concat based on the previous todo (first + Last + Suffix for most) 


# with open('employeeInfo.csv', 'w', newline='') as csvfile:
# 	fieldnames = ["name", "nameSuffix", "title", "hireDate", "priorYearsFirm", "priorYearsOther", "hireDate", "priorYearsFirm", "priorYearsOther", "licenses", "courses", "certifications", "education", "resumeIntro", "totalYearsExp"]
# 	writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
# 	writer.writeheader()


def formatDate(date):
	return datetime.strptime(date.rstrip("T00:00:00"), '%Y-%m-%d').date() #May add to Employee

def representsInt(input):
	if input == None:
		return False
	try:
		int(input)
		return True
	except ValueError:
		pass
	try:
		int(input[:-1])
		return True
	except ValueError:
		return False


def previous_and_next(some_iterable): #https://stackoverflow.com/questions/1011938/python-previous-and-next-values-inside-a-loop
	prevs, items, nexts = tee(some_iterable, 3)
	prevs = chain([None], prevs)
	nexts = chain(islice(nexts, 1, None), [None])
	return zip(prevs, items, nexts)

#base class Employee
class Employee:
	"""Base Class for all employees, used to write each csv row"""
	def __init__(self, name):
		self.data = {
			"name": name,
			"nameSuffix": "",							
			"title": "",
			"hireDate": 0,
			"priorYearsFirm": 0,
			"priorYearsOther": 0,
			"licenses": {}, # Write a function for employee to roll up all the displays for licenses into one string to put in csv
			"licenseDisplay": "",
			"eduDisplay": "",
			"courseDisplay": "",
			"courses": [],
			"courseObjects": {},
			"certifications": [], 
			"education": {},
			"resumeIntro": [], 
			"PELicenseCount": 0, 
			"courseCount": 0,
			"certCount": 0, 
			"degreeCount": 0
		}

	detailKey = {
		"DetailField_FullName_Section_1": "hook",
		"DetailField_Title_Section_1": "bio",
		"DetailField_HireDate_Section_1": "bio",
		"DetailField_PriorYearsFirm_Section_1": "bio", 
		"DetailField_YearsOtherFirms_Section_1": "bio",
		"DetailField_Suffix_Section_1": "bio",
		"detail_Licenses_License": "lic",
		"detail_Licenses_Earned": "lic",
		"detail_Licenses_State": "lic",
		"detail_Licenses_Number": "lic",
		"detail_Licenses_Expires":  "lic", 
		"detail_gridUDCol_Employees_Courses_custAgency": "cor", 
		"detail_gridUDCol_Employees_Courses_custCourseName": "cor",
		"detail_gridUDCol_Employees_Courses_custDate": "cor",
		"detail_gridUDCol_Employees_Courses_custCourseNumber": "cor",
		"detail_gridUDCol_Employees_Certifications_custCertAgency": "cert",
		"detail_gridUDCol_Employees_Certifications_custCertNumber": "cert",
		"detail_gridUDCol_Employees_Certifications_custTitle": "cert",
		"detail_gridUDCol_Employees_Certifications_custExpirationDate": "cert",
		"detail_gridUDCol_Employees_Certifications_CustNoExpiration": "cert",
		"detail_Education_Degree": "edu",
		"detail_Education_Specialty": "edu",
		"detail_Education_Institution": "edu",
		"detail_Education_Year": "edu",
		# "Design and Inspection Resume": "",
		# "detail_level": "res" #Shows up first but serves as a poor hook
	}

	dataKey = { #need to add other fields
		"DetailField_FullName_Section_1": ["name", "str"],
		"DetailField_Title_Section_1": ["title", "str"],
		"DetailField_HireDate_Section_1": ["hireDate", "date"],
		"DetailField_PriorYearsFirm_Section_1": ["priorYearsFirm", "int"], 
		"DetailField_YearsOtherFirms_Section_1": ["priorYearsOther", "int"], 
		"DetailField_Suffix_Section_1": ["nameSuffix", "str"] 
	}

	objectKey = {
		"detail_Licenses_License": "type",
		"detail_Licenses_Earned": "dateEarned",
		"detail_Licenses_State": "state",
		"detail_Licenses_Number": "number",
		"detail_Licenses_Expires":  "dateExpire",
		"detail_gridUDCol_Employees_Courses_custAgency": "agency", 
		"detail_gridUDCol_Employees_Courses_custCourseName": "name",
		"detail_gridUDCol_Employees_Courses_custDate": "dateTaken",
		"detail_gridUDCol_Employees_Courses_custCourseNumber": "number",
		"detail_Education_Degree": "degree",
		"detail_Education_Specialty": "specialty",
		"detail_Education_Institution": "school",
		"detail_Education_Year": "gradYear",
	}

	def formatData(self, data, dataType):
		if dataType == "int":
			return int(data)
		elif dataType == "date":
			return formatDate(data)
		return data


	def bioDataProcessor(self, val, tagName):
		dataType = self.dataKey[tagName][1]
		formattedData = self.formatData(val, dataType)  #Functioned derrived from base class Employee
		self.data[self.dataKey[tagName][0]] = formattedData


	def calculateYearsExp(self):
		today = datetime.now()
		difference = relativedelta.relativedelta(today, self.data["hireDate"]).years
		self.data["totalYearsExp"] = (difference + self.data["priorYearsFirm"] + self.data["priorYearsOther"])

	def rollupLicenses(self):
		for license in self.data["licenses"]:
			if self.data["licenses"][license].displayString == None:
				return
			self.data["licenseDisplay"] += self.data["licenses"][license].displayString

	def removeTrailingComma(self):
		if self.data["licenseDisplay"].endswith(", "):
			self.data["licenseDisplay"] = self.data["licenseDisplay"][: -2]

	def rollupEducation(self):
		for edu in self.data["education"]:
			self.data["eduDisplay"] += self.data["education"][edu].displayString + "   " #Need to make this a GREP searchabel expression to replace with a return line

	def parseCourses(self):
		stagingArray = {}

		for item in self.data["courses"]:
			if item[0] in stagingArray:
				name = None
				dateTaken = None
				agency = None
				number = None
				self.data["courseCount"] += 1
				if "detail_gridUDCol_Employees_Courses_custCourseName" in stagingArray:
					name = stagingArray["detail_gridUDCol_Employees_Courses_custCourseName"] 
				if "detail_gridUDCol_Employees_Courses_custAgency" in stagingArray:
					agency = stagingArray["detail_gridUDCol_Employees_Courses_custAgency"]
				if "detail_gridUDCol_Employees_Courses_custDate" in stagingArray:
					dateTaken = stagingArray["detail_gridUDCol_Employees_Courses_custDate"]
				if "detail_gridUDCol_Employees_Courses_custCourseNumber" in stagingArray:
					number = stagingArray["detail_gridUDCol_Employees_Courses_custCourseNumber"]
				self.data["courseObjects"][self.data["courseCount"]] = Course(name, dateTaken, agency, number)
				stagingArray.clear()
				stagingArray[item[0]] = item[1]
			else:
				stagingArray[item[0]] = item[1]
		name = None
		dateTaken = None
		agency = None
		number = None
		self.data["courseCount"] += 1
		if "detail_gridUDCol_Employees_Courses_custCourseName" in stagingArray:
			name = stagingArray["detail_gridUDCol_Employees_Courses_custCourseName"] 
		if "detail_gridUDCol_Employees_Courses_custAgency" in stagingArray:
			agency = stagingArray["detail_gridUDCol_Employees_Courses_custAgency"]
		if "detail_gridUDCol_Employees_Courses_custDate" in stagingArray:
			dateTaken = stagingArray["detail_gridUDCol_Employees_Courses_custDate"]
		if "detail_gridUDCol_Employees_Courses_custCourseNumber" in stagingArray:
			number = stagingArray["detail_gridUDCol_Employees_Courses_custCourseNumber"]
		self.data["courseObjects"][self.data["courseCount"]] = Course(name, dateTaken, agency, number)



allEmployees = dict()

class License(Employee): 
	"""Class for each individual license - will format itself properly"""
	def __init__(self):
		self.data = {
			"type": "",
			"dateEarned": 0, 
			"state": "",
			"number": 0,
			"dateExpire": 0
		}
		self.isMain = False
		self.isExpired = False
		self.displayString = None 

	def formatType(self, licType):
		if licType == "Professional Engineer":
			return "PE"
		elif licType == "Engineer Intern":
			return "EI"
		elif licType == "Engineer In Training":
			return "EIT"

	def getLicenseResumeFormat(self): #This will omit licenses with no expiration... Consider implications 
		# 1. check for expiration
		# 2. check if it is main (in the state where they live)
		# 3. return State (+ number if main (or only one))   - TO GET HOME STATE - need to grab tag "MainTable_CRM" and get all meta date. Could do the object creation for each this way and then match the first hook (full name) with the full name from the MainTable_CRM as I believe they will always be the same. Or accept User input for main license 

		if self.data["dateExpire"] == 0:
			return
		today = datetime.now().date()
		if today > formatDate(self.data["dateExpire"]):
			return self.data["state"] + " License is expired for " #Get employee name to make this a more useful error
		elif today < formatDate(self.data["dateExpire"]):
			if self.isMain == True:
				self.displayString = self.data["state"] + " #" + self.data["number"]
			self.displayString = self.data["state"] + ", "

	
		
class Course(Employee):

	def __init__(self, name=None, dateTaken=None, agency=None, number=0):
		self.data = {
			"agency": agency,
			"name": name,
			"dateTaken": dateTaken,
			"number": number
		}
		self.isExpired: False
		self.displayString: None

class Degree(Employee):

	def __init__(self):
		self.data = {
			"degree": "",
			"specialty": "",
			"school": "",
			"gradYear": 0
		}

		self.displayString: None

	def getEducationResumeFormat(self):
		self.displayString = self.data["degree"] + ", " + self.data["specialty"] + ", " + self.data["school"] + ", " + self.data["gradYear"]


newCert = {
	"agency": "",
	"number": 0, 
	"expDate": 0, 
	"expires": True #optional
}

newDegree = {
	"degree": "",
	"specialty": "",
	"school": "",
	"gradYear": 0
}

# To do - make this switch nonsense more sane 
for element in section[:]:
	for key in Employee.detailKey:
		if element.getAttribute(key):
			if Employee.detailKey[key] == "hook":
				objectName = element.getAttribute(key) 
				allEmployees[objectName] = Employee(objectName)
			if Employee.detailKey[key] == "bio":
				allEmployees[objectName].bioDataProcessor(element.getAttribute(key), key)
			if Employee.detailKey[key] == "lic":
				if Employee.objectKey[key] == "type": # Creates new license each time "type" is encountered after "lic" creating new license based on the counter. Yes this is messy and needs to be refactored
					allEmployees[objectName].data["PELicenseCount"] += 1
					allEmployees[objectName].data["licenses"][allEmployees[objectName].data["PELicenseCount"]] = License()
					allEmployees[objectName].data["licenses"][allEmployees[objectName].data["PELicenseCount"]].data[Employee.objectKey[key]] = element.getAttribute(key)	

				allEmployees[objectName].data["licenses"][allEmployees[objectName].data["PELicenseCount"]].data[Employee.objectKey[key]] = element.getAttribute(key)	
			if Employee.detailKey[key] == "cor":
				allEmployees[objectName].data["courses"].append([key, element.getAttribute(key)]) 
			if Employee.detailKey[key] == "cert":
				allEmployees[objectName].data["certifications"].append(element.getAttribute(key)) 
			if Employee.detailKey[key] == "edu":
				if Employee.objectKey[key] == "degree":
					allEmployees[objectName].data["degreeCount"] += 1
					allEmployees[objectName].data["education"][allEmployees[objectName].data["degreeCount"]] = Degree()
					allEmployees[objectName].data["education"][allEmployees[objectName].data["degreeCount"]].data[Employee.objectKey[key]] = element.getAttribute(key)

				allEmployees[objectName].data["education"][allEmployees[objectName].data["degreeCount"]].data[Employee.objectKey[key]] = element.getAttribute(key)
			if Employee.detailKey[key] == "res":
				allEmployees[objectName].data["resumeIntro"].append(element.getAttribute(key)) 

# Loop to execute for each license
for emp in allEmployees:
	for edu in allEmployees[emp].data["education"]:
		allEmployees[emp].data["education"][edu].getEducationResumeFormat()

employeeCount = 0 

# Rollup licenses for employee
for emp in allEmployees:
	employeeCount += 1
	allEmployees[emp].rollupLicenses()
	allEmployees[emp].calculateYearsExp()
	allEmployees[emp].removeTrailingComma()
	allEmployees[emp].rollupEducation()
	allEmployees[emp].parseCourses()
	for cor in allEmployees[emp].data["courseObjects"]:
		print(allEmployees[emp].data["name"], allEmployees[emp].data["courseObjects"][cor].data)
	# print(employeeCount, allEmployees[emp].data["name"], allEmployees[emp].data["courseObjects"])



# for emp in allEmployees:
# 	print(allEmployees[emp].data)

# 	for course in allEmployees[emp].data["courses"]:
# 		print(allEmployees[emp].data["courses"][course].data)



# employeeData = Worker.data

# print(employeeData)


# Write to iterate over the dictionary of workers 

with open('employeeInfo.csv', 'w', newline='') as csvfile:
	fieldnames = ["name", "nameSuffix", "title", "hireDate", "priorYearsFirm", "priorYearsOther", "licenseDisplay", "eduDisplay", "courseDisplay", "certifications", "resumeIntro", "totalYearsExp", "degreeCount", "courseCount", "PELicenseCount", "certCount", "licenses", "education", "courses", "courseObjects"]
	writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
	writer.writeheader()
	for emp in allEmployees:
		writer.writerow(allEmployees[emp].data)

