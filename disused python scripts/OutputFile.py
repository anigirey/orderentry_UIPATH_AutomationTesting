"""
Author: Jeremy Wheeler
Project: UIPath PCMT/CPlus Automation
Date: 02/21/2020

The purpose of this file is to create a directory at the root
called '/Output' and to create a file within that directory
called 'outputvars.xlsx'.
"""


import xlsxwriter
import os
from datetime import datetime
from os import path


def main():

    # variables for environments and imports
    now = datetime.now()
    currDir = os.getcwd()
    currDat = now.strftime("%Y-%m-%d")
    currTim = now.strftime("%H.%M.%S")
    dbDat = now.strftime("%d-%b-%y")
    rDatStr = str(currDat)
    rTimStr = str(currTim)
	
    createOutputDirectory(currDir)
    createOutputFile(currDir, rDatStr, rTimStr, dbDat)
	

def createOutputDirectory(currDir):

    if not path.exists((currDir) + "/Output/"):
        outputDirectory = ((currDir) + "/Output/")
        os.makedirs(outputDirectory)
		
def createOutputFile(currDir, rDatStr, rTimStr, dbDat):

    outputWorkbook = xlsxwriter.Workbook((currDir) + "/Output/outputvars.xlsx")
    outputWorksheet = outputWorkbook.add_worksheet("Output")
    outputWorksheet.set_column("A:A", 10)
    outputWorksheet.set_column("B:B", 10)
    outputWorksheet.set_column("C:C", 10)
    outputWorksheet.write("A1", rDatStr)
    outputWorksheet.write("B1", rTimStr)
    outputWorksheet.write("C1", dbDat)
    outputWorkbook.close()
	

main()
