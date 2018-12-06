from xml.dom import minidom
import csv

mydoc = minidom.parse('sasher.xml')

section = mydoc.getElementsByTagName('Detail')

# print(table[0].attribute['detail_description'].value)

# deatilCollection = section[0].getElementsByTagName('Detail')[0].attributes['DetailField_FullName_Section_1'].value

# deatilCollection = type(section[0].getElementsByTagName('Detail'))


#object methods needed
# - to deal with date formats
# - to calculate employee expereience number
# - to accept multiple courses and certs
# - 

# x = 4

employeeInfo = {
	"DetailField_FullName_Section_1": 0,
	"DetailField_Title_Section_1": 0,
	"DetailField_HireDate_Section_1": 0,
	"DetailField_PriorYearsFirm_Section_1": 0, 
	"DetailField_YearsOtherFirms_Section_1": 0,
	"detail_Licenses_License": 0,
	"detail_Licenses_Earned": 0,
	"detail_Licenses_State": 0,
	"detail_Licenses_Number": 0,
	"detail_Licenses_Expires": 0,
	"detail_gridUDCol_Employees_Courses_custAgency": 0, #Need to rewrite to accept multiple courses and certs
	"detail_gridUDCol_Employees_Courses_custCourseName": 0,
	"detail_gridUDCol_Employees_Courses_custDate": 0,
	"detail_gridUDCol_Employees_Courses_custCourseNumber": 0,
	"detail_gridUDCol_Employees_Certifications_custCertAgency": 0,
	"detail_gridUDCol_Employees_Certifications_custCertNumber": 0,
	"detail_gridUDCol_Employees_Certifications_custExpirationDate": 0,
	"detail_gridUDCol_Employees_Certifications_CustNoExpiration": 0,
	"detail_Education_Degree": 0,
	"detail_Education_Specialty": 0,
	"detail_Education_Institution": 0,
	"detail_Education_Year": 0,
	"Design and Inspection Resume": 0,
	"detail_level": 0
}


for element in section[:]:
	for key in employeeInfo:
		if element.getAttribute(key):
			employeeInfo[key] = element.getAttribute(key)


with open('employeeInfo.csv', 'w') as csvfile:
	fieldnames = ["DetailField_FullName_Section_1", "DetailField_Title_Section_1", "DetailField_HireDate_Section_1", "DetailField_PriorYearsFirm_Section_1", "DetailField_YearsOtherFirms_Section_1", "detail_Licenses_License", "detail_Licenses_Earned", "detail_Licenses_State", "detail_Licenses_Number", "detail_Licenses_Expires", "detail_gridUDCol_Employees_Courses_custAgency", "detail_gridUDCol_Employees_Courses_custCourseName", "detail_gridUDCol_Employees_Courses_custDate", "detail_gridUDCol_Employees_Courses_custCourseNumber", "detail_gridUDCol_Employees_Certifications_custCertAgency", "detail_gridUDCol_Employees_Certifications_custCertNumber", "detail_gridUDCol_Employees_Certifications_custExpirationDate", "detail_gridUDCol_Employees_Certifications_CustNoExpiration", "detail_Education_Degree", "detail_Education_Specialty", "detail_Education_Institution", "detail_Education_Year", "Design and Inspection Resume", "detail_level"]
	writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
	writer.writeheader()
	writer.writerow(employeeInfo)