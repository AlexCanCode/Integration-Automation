from xml.dom import minidom
from datetime import datetime
from dateutil import relativedelta
import csv

mydoc = minidom.parse('kincaid.xml')

section = mydoc.getElementsByTagName('Detail')

# print(table[0].attribute['detail_description'].value)

# deatilCollection = section[0].getElementsByTagName('Detail')[0].attributes['DetailField_FullName_Section_1'].value

# deatilCollection = type(section[0].getElementsByTagName('Detail'))


#object methods needed
# - to deal with date formats
# 	- Write funciton to calculate years of expierence date
# - Calculate the right licnese to use
# - to accept multiple courses and certs
# - NOTE - IN ORDER TO MAKE LICENSE AND COURSES WORK - you gotta learn how to define classes so that you can have empty variables of them. Pointers are the issue here

def appendStr(val, field, dataType):
	if dataType == "int":
		bio[field] = int(val)
	elif dataType == "hireDate":
		bio[field] = formatDate(val) #Calculates date when called
	elif dataType == "str":
		bio[field].append(val)
	elif dataType == "lic":
		if field == "dateEarned":
			newLicense[field] = formatDate(val)
		elif field == "number":
			newLicense[field] = int(val)
		elif field == "type" or field == "state":
			newLicense[field] = val
		elif field == "dateExpired":
			newLicense[field] == formatDate(val)
			bio["licenses"].update({ newLicense["state"] + str(newLicense["number"]) : newLicense  })


def formatDate(date):
	return datetime.strptime(date.rstrip("T00:00:00"), '%Y-%m-%d').date() #May add to Employee

#base class Employee
class Employee:
	"""Base Class for all employees, used to write each csv row"""
	def __init__(self):
		self.bioData = None

	detailKey = {
		"DetailField_FullName_Section_1": "bio",
		"DetailField_Title_Section_1": "bio",
		"DetailField_HireDate_Section_1": "bio",
		"DetailField_PriorYearsFirm_Section_1": "bio", 
		"DetailField_YearsOtherFirms_Section_1": "bio",
		"DetailField_Suffix_Section_1": "bio",
		"detail_Licenses_License": "lic",
		"detail_Licenses_Earned": "lic",
		"detail_Licenses_State": "lic",
		"detail_Licenses_Number": "lic",
		"detail_Licenses_Expires":  "lic", #Hook to signal end of newLicense object
		"detail_gridUDCol_Employees_Courses_custAgency": "cor", #Need to rewrite to accept multiple courses and certs
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
		# "detail_level": ""
	}

	def formatData(self, data, dataType):
		if dataType == "int":
			return int(data)
		elif dataType == "date":
			return formatDate(data)
		return data

class Bio(Employee): 
	"""Class for handling bio information for each employee. Calculates tota; years exp"""
	def __init__(self):
		self.data = {
			"name": "", #could split between first last and include preffered too
			"nameSuffix": "",							
			"title": "",
			"hireDate": 0,
			"priorYearsFirm": 0,
			"priorYearsOther": 0,
			"licenses": [],
			"courses": [],
			"certifications": [], 
			"education": []
		}

	bioKey = {
		"DetailField_FullName_Section_1": ["name", "str"],
		"DetailField_Title_Section_1": ["title", "str"],
		"DetailField_HireDate_Section_1": ["hireDate", "date"],
		"DetailField_PriorYearsFirm_Section_1": ["priorYearsFirm", "int"], 
		"DetailField_YearsOtherFirms_Section_1": ["priorYearsOther", "int"], 
		"DetailField_Suffix_Section_1": ["nameSuffix", "str"] 
	}

	def bioDataProcessor(self, val, tagName):
		dataType = self.bioKey[tagName][1]
		formattedData = self.formatData(val, dataType)  #Functioned derrived from base class Employee
		self.data[self.bioKey[tagName][0]] = formattedData



	def calculateYearsExp(self):
		today = datetime.now()
		difference = relativedelta.relativedelta(today, self.data["hireDate"]).years
		self.data["totalYearsExp"] = (difference + self.data["priorYearsFirm"] + self.data["priorYearsOther"])

class License(Bio):

	def __init__(self):
		self.licenseData = {
			"number": 0, 
			"type": "",
			"dateEarned": 0,
			"dateExpired": 0,
			"state": "",
			"proposalUse": False
		}

	licneseKey = {
		"detail_Licenses_License": ["type", "lic"],
		"detail_Licenses_Earned": ["dateEarned", "lic"],
		"detail_Licenses_State": ["state", "lic"],
		"detail_Licenses_Number": ["number", "lic"],
		"detail_Licenses_Expires": ["dateExpired", "lic"], 
	}

	def licenseDataProcessor(self, val, tagName):
		dataType = self.licneseKey[tagName][1]
		formattedData = self.formatData(val, dataType)  #Functioned derrived from base class Employee
		self.licenseData[self.licneseKey[tagName][0]] = formattedData

Sasher = Employee()
SasherBio = Bio()

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



for element in section[:]:
	for key in Employee.detailKey:
		if element.getAttribute(key):
			if Employee.detailKey[key] == "bio":
				SasherBio.bioDataProcessor(element.getAttribute(key), key) # LEFT OFF - this function needs to be dynamic and needs to be able to recognize each object that it needs to be delegated too. 
			if Employee.detailKey[key] == "lic":
				SasherBio.data["licenses"].append(element.getAttribute(key)) 
			if Employee.detailKey[key] == "cor":
				SasherBio.data["courses"].append(element.getAttribute(key))
			if Employee.detailKey[key] == "cert":
				SasherBio.data["certifications"].append(element.getAttribute(key)) 
			if Employee.detailKey[key] == "edu":
				SasherBio.data["education"].append(element.getAttribute(key)) 

SasherBio.calculateYearsExp()

Sasher.bioData = SasherBio.data

print(Sasher.bioData)


# with open('employeeInfo.csv', 'w') as csvfile:
# 	fieldnames = ["DetailField_FullName_Section_1", "DetailField_Title_Section_1", "DetailField_HireDate_Section_1", "DetailField_PriorYearsFirm_Section_1", "DetailField_YearsOtherFirms_Section_1", "detail_Licenses_License", "detail_Licenses_Earned", "detail_Licenses_State", "detail_Licenses_Number", "detail_Licenses_Expires", "detail_gridUDCol_Employees_Courses_custAgency", "detail_gridUDCol_Employees_Courses_custCourseName", "detail_gridUDCol_Employees_Courses_custDate", "detail_gridUDCol_Employees_Courses_custCourseNumber", "detail_gridUDCol_Employees_Certifications_custCertAgency", "detail_gridUDCol_Employees_Certifications_custCertNumber", "detail_gridUDCol_Employees_Certifications_custExpirationDate", "detail_gridUDCol_Employees_Certifications_CustNoExpiration", "detail_Education_Degree", "detail_Education_Specialty", "detail_Education_Institution", "detail_Education_Year", "Design and Inspection Resume", "detail_level"]
# 	writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
# 	writer.writeheader()
# 	writer.writerow(employeeInfo)