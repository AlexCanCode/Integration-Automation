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
			"licenses": [], # TO -DO SHOULD MAKE EACH OF THESE A DICTIONARY THAT ASSIGNS THE DATA TYPE (even if its just in the xml tag format) to each so that it can be easily processed and seperated into easier to use formats. 
			"courses": [],
			"certifications": [], 
			"education": [], 
			"resumeIntro": []
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


# Worker = Employee()

# UES INPUT LIKE THIS? TO FORMAT DATA THAT HAS MULTIPLE ITERATIONS OF Each (e.g. License) 
    # def add_license(self, info1, info2, info3, info4):
    #     self.info1 = info1
    #     self.zip = info2
    #     self.info3 = info3
    #     self.info4 = info4

    #But remember that scope will be a factor and you need to deal with potential non-inputs - look into best way to perform these kinds of nests

allEmployees = dict()

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
				allEmployees[objectName].bioDataProcessor(element.getAttribute(key), key) # LEFT OFF - this function needs to be dynamic and needs to be able to recognize each object that it needs to be delegated too. 
			if Employee.detailKey[key] == "lic":
				allEmployees[objectName].data["licenses"].append(element.getAttribute(key)) 
			if Employee.detailKey[key] == "cor":
				allEmployees[objectName].data["courses"].append(element.getAttribute(key))
			if Employee.detailKey[key] == "cert":
				allEmployees[objectName].data["certifications"].append(element.getAttribute(key)) 
			if Employee.detailKey[key] == "edu":
				allEmployees[objectName].data["education"].append(element.getAttribute(key))
			if Employee.detailKey[key] == "res":
				allEmployees[objectName].data["resumeIntro"].append(element.getAttribute(key)) 

for worker in allEmployees:
	print(allEmployees[worker].data)

# employeeData = Worker.data

# print(employeeData)


# Write to iterate over the dictionary of workers 

# with open('employeeInfo.csv', 'w', newline='') as csvfile:
# 	fieldnames = ["name", "nameSuffix", "title", "hireDate", "priorYearsFirm", "priorYearsOther", "licenses", "courses", "certifications", "education", "resumeIntro", "totalYearsExp"]
# 	writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
# 	writer.writerow(Worker.data)

