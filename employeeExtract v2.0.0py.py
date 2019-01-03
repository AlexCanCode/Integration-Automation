import xmltodict
import json 

with open('sasher.xml', encoding="utf8") as fd:
	doc = xmltodict.parse(fd.read())
	print(json.dumps(doc, indent=4, sort_keys=True))