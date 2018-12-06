from xml.dom import minidom
import csv

mydoc = minidom.parse('sasher.xml')

section = mydoc.getElementsByTagName('Detail')

# print(table[0].attribute['detail_description'].value)

# deatilCollection = section[0].getElementsByTagName('Detail')[0].attributes['DetailField_FullName_Section_1'].value

# deatilCollection = type(section[0].getElementsByTagName('Detail'))


# x = 4

bio = {
	"DetailField_FullName_Section_1": 0,
	"DetailField_Title_Section_1": 0,
	"DetailField_HireDate_Section_1": 0,
	"DetailField_PriorYearsFirm_Section_1": 0,
	"DetailField_YearsOtherFirms_Section_1": 0
}


for element in section[:]:
	for key in bio:
		if element.getAttribute(key):
			bio[key] = element.getAttribute(key)


with open('bio.csv', 'w') as csvfile:
	fieldnames = ["DetailField_FullName_Section_1", "DetailField_Title_Section_1", "DetailField_HireDate_Section_1", "DetailField_PriorYearsFirm_Section_1", "DetailField_YearsOtherFirms_Section_1"]
	writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
	writer.writeheader()
	writer.writerow(bio)



