from xml.dom import minidom
from datetime import datetime
from dateutil import relativedelta
import csv

mydoc = minidom.parse('sasher.xml')

section = mydoc.getElementsByTagName('Detail')

# print(table[0].attribute['detail_description'].value)

# deatilCollection = section[0].getElementsByTagName('Detail')[0].attributes['DetailField_FullName_Section_1'].value

# deatilCollection = type(section[0].getElementsByTagName('Detail'))


#object methods needed
# - to deal with date formats
# 	- Write funciton to calculate years of expierence date
# - Calculate the right licnese to use
# - to accept multiple courses and certs
# - 

def appendStr(val, field, dataType):
	if dataType == "int":
		employee[field] = int(val)
	elif dataType == "hireDate":
		employee[field] = datetime.strptime(val.rstrip("T00:00:00"), '%Y-%m-%d') #Calculates date when called
	elif dataType == "str":
		employee[field].append(val)

def calculateYearsBetween(startDate):
	today = datetime.now()
	difference = relativedelta.relativedelta(today, startDate).years
	return difference

employee = {
	"name": [],
	"title": [],
	"hireDate": 0,
	"priorYearsFirm": 0,
	"priorYearsOther": 0,
}

employeeInfo = {
	"DetailField_FullName_Section_1": ["name", "str"],
	"DetailField_Title_Section_1": ["title", "str"],
	"DetailField_HireDate_Section_1": ["hireDate", "hireDate"],
	"DetailField_PriorYearsFirm_Section_1": ["priorYearsFirm", "int"], 
	"DetailField_YearsOtherFirms_Section_1": ["priorYearsOther", "int"]
	# "detail_Licenses_License": "",
	# "detail_Licenses_Earned": "",
	# "detail_Licenses_State": "",
	# "detail_Licenses_Number": 0,
	# "detail_Licenses_Expires": 0,
	# "detail_gridUDCol_Employees_Courses_custAgency": "", #Need to rewrite to accept multiple courses and certs
	# "detail_gridUDCol_Employees_Courses_custCourseName": "",
	# "detail_gridUDCol_Employees_Courses_custDate": "",
	# "detail_gridUDCol_Employees_Courses_custCourseNumber": "",
	# "detail_gridUDCol_Employees_Certifications_custCertAgency": "",
	# "detail_gridUDCol_Employees_Certifications_custCertNumber": "",
	# "detail_gridUDCol_Employees_Certifications_custExpirationDate": "",
	# "detail_gridUDCol_Employees_Certifications_CustNoExpiration": "",
	# "detail_Education_Degree": "",
	# "detail_Education_Specialty": "",
	# "detail_Education_Institution": "",
	# "detail_Education_Year": "",
	# "Design and Inspection Resume": "",
	# "detail_level": ""
}


for element in section[:]:
	for key in employeeInfo:
		if element.getAttribute(key):
			appendStr(element.getAttribute(key), employeeInfo[key][0], employeeInfo[key][1])

# yearsAtConsor = calculateYearsBetween(employee["hireDate"])



employee.update({"yearsAtConsor": calculateYearsBetween(employee["hireDate"])}) #update employee with years at consor number

employee.update({"totalYearsExp": employee["yearsAtConsor"] + employee["priorYearsOther"] + employee["priorYearsFirm"]})

print(employee)


# with open('employeeInfo.csv', 'w') as csvfile:
# 	fieldnames = ["DetailField_FullName_Section_1", "DetailField_Title_Section_1", "DetailField_HireDate_Section_1", "DetailField_PriorYearsFirm_Section_1", "DetailField_YearsOtherFirms_Section_1", "detail_Licenses_License", "detail_Licenses_Earned", "detail_Licenses_State", "detail_Licenses_Number", "detail_Licenses_Expires", "detail_gridUDCol_Employees_Courses_custAgency", "detail_gridUDCol_Employees_Courses_custCourseName", "detail_gridUDCol_Employees_Courses_custDate", "detail_gridUDCol_Employees_Courses_custCourseNumber", "detail_gridUDCol_Employees_Certifications_custCertAgency", "detail_gridUDCol_Employees_Certifications_custCertNumber", "detail_gridUDCol_Employees_Certifications_custExpirationDate", "detail_gridUDCol_Employees_Certifications_CustNoExpiration", "detail_Education_Degree", "detail_Education_Specialty", "detail_Education_Institution", "detail_Education_Year", "Design and Inspection Resume", "detail_level"]
# 	writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
# 	writer.writeheader()
# 	writer.writerow(employeeInfo)