#!/usr/bin/python3
#coding: utf8

import os
import sys
import re
from xlrd import open_workbook

xlFile = r'/path/to/excelFile.xls'
wb = open_workbook(xlFile)

def showName(param):
	paramNoCase = re.compile(param,re.IGNORECASE)
	for s in wb.sheets():
		#print ("Sheet: %s " % s.name)
		for row in range(s.nrows):
			values=[]
			for col in range(s.ncols):
				values.append(s.cell(row,col).value)
			rowString = ",".join(str(v) for v in values)
			match = re.search(paramNoCase, rowString)
			if(match):
				#print(match.group())
				print(rowString)
		print()
	return

def main():

	if len(sys.argv) > 1:
		searchParam = sys.argv[1]
		regExpParam = searchParam
	else:
		print ("Enter a search parameter")
		exit()
	
	if not os.path.isfile(xlFile):
		print ("\n Can\'t find file %s - Connect to server" % xlFile)
		exit()
	else:
		showName(regExpParam)

if __name__ == '__main__': 
    main()



