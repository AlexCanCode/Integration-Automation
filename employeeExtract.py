from xml.dom import minidom
from datetime import datetime
from dateutil import relativedelta
import csv

# SEE LINE 99 for to do item

mydoc = minidom.parse('combinedTest.xml')

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


with open('employeeInfo.csv', 'w', newline='') as csvfile:
	fieldnames = ["name", "nameSuffix", "title", "hireDate", "priorYearsFirm", "priorYearsOther", "hireDate", "priorYearsFirm", "priorYearsOther", "licenses", "courses", "certifications", "education", "resumeIntro", "totalYearsExp"]
	writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
	writer.writeheader()


def formatDate(date):
	return datetime.strptime(date.rstrip("T00:00:00"), '%Y-%m-%d').date() #May add to Employee

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
			"courses": [],
			"certifications": [], 
			"education": [], 
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

	def formatLicense(self):
		print(self.data)

	def formatType(self, licType):
		if licType == "Professional Engineer":
			return "PE"
		elif licType == "Engineer Intern":
			return "EI"
		elif licType == "Engineer In Training":
			return "EIT"

	def getLicenseResumeFormat(self): #This will omit license with no expiration... Consider implications 
		# 1. check for expiration
		# 2. check if it is main (in the state where they live)
		# 3. return State (+ number if main (or only one))   - TO GET HOME STATE - need to grab tag "MainTable_CRM" and get all meta date. Could do the object creation for each this way and then match the first hook (full name) with the full name from the MainTable_CRM as I believe they will always be the same. Or accept User input for main license 

		if self.data["dateExpire"] == 0:
			return
		today = datetime.now().date()
		if today > formatDate(self.data["dateExpire"]):
			print(self.data["state"] + " License is expired for ") #Get employee name to make this a more useful error
		elif today < formatDate(self.data["dateExpire"]):
			# if self.isMain == True:
			# 	self.displayString = self.data["state"] + " #" + self.data["number"]
			self.displayString = self.data["state"] + " "
		


newCourse = {
	"agency": "",
	"name": "",
	"dateTaken": 0,
	"number": 0
}

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
				allEmployees[objectName].data["courses"].append(element.getAttribute(key))
			if Employee.detailKey[key] == "cert":
				allEmployees[objectName].data["certifications"].append(element.getAttribute(key)) 
			if Employee.detailKey[key] == "edu":
				allEmployees[objectName].data["education"].append(element.getAttribute(key))
			if Employee.detailKey[key] == "res":
				allEmployees[objectName].data["resumeIntro"].append(element.getAttribute(key)) 

x = allEmployees["James Ikaika Kincaid PE"].data["licenses"]

for j in x: 
	x[j].getLicenseResumeFormat()
	print(x[j].data)
	print(x[j].displayString)






# employeeData = Worker.data

# print(employeeData)


# Write to iterate over the dictionary of workers 

# with open('employeeInfo.csv', 'w', newline='') as csvfile:
# 	fieldnames = ["name", "nameSuffix", "title", "hireDate", "priorYearsFirm", "priorYearsOther", "licenses", "courses", "certifications", "education", "resumeIntro", "totalYearsExp"]
# 	writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
# 	writer.writerow(Worker.data)

