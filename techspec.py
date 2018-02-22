# -*- coding: utf-8 -*-
from xml.sax.saxutils import unescape
from bs4 import BeautifulSoup
import pyexcel as pe
import os
from os.path import basename

# excel spread sheet with NID, CID, PID
# create dictionary from the NID and PID column
records = pe.get_sheet(file_name="id.xlsx")
nid = []
cid = []
for row in records:
	nid.append(str(row[0]))
	if not row[1]:
		cid.append('none')
	else:
		cid.append(str(row[1]))
dictionary = dict(zip(cid, nid))

# Open ECM export folder
ECMPath = '/Users/ramarit/desktop/python scripts/TechSpec/Clickability/ECM exports'

# loop through directory of ECM exports and open each file
for ECMfilename in os.listdir(ECMPath):
	if not ECMfilename.endswith('.xml'):continue
	fullname = os.path.join(ECMPath, ECMfilename)
	with open(fullname) as export:
		export = BeautifulSoup(export, 'html.parser')
		folder = ECMfilename[31:36].upper()
		path = f'/Users/ramarit/desktop/python scripts/TechSpec/Clickability/Translated_Spec_Tables/{folder}'

		# find correspoding clickability exports in clickability directory
		for filename in os.listdir(path):
			if not filename.endswith('.xml'):continue
			fullname = os.path.join(path, filename)
			print(fullname)
			with open(fullname) as clickability:
				clickability = BeautifulSoup(clickability, 'html.parser')
				for product in export.findAll('entity'):
					nodeID = product.find('id').contents[0]
					# get tech specs for that node
					if product.find('field_specs') is not None:
						for content in clickability.findAll('product'):
							CID = content.find('productid').contents[0]
							CID = str(CID)
							# lookup corresponding clickability node from dictionary and find grab its tech specs
							getNode = dictionary.get(CID, None)
							if (nodeID == getNode):
								exportTechSpecs = product.find('field_specs').contents[0]
								if content.find('techspecs') is not None:
									pcatTechSpecs = content.find('techspecs').contents[0]
									replace = exportTechSpecs.replace_with(pcatTechSpecs)
									# print(replace)

		# save new file and unescape characters to keep CDATA tags
		with open(ECMfilename, 'w') as f:
			f.write(unescape(str(export)))
