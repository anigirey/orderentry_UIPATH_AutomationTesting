"""

"""


import openpyxl
import xlrd
import os
from os import path


def main():

    currDir = os.getcwd()
    loc = currDir + ("/Output/outputvars.xlsx")
    wBook = xlrd.open_workbook(loc)
    wSheet = wBook.sheet_by_index(0)
    testEnv = wSheet.cell_value(0,3)
    testTim = wSheet.cell_value(0,1)

    sanityWorkbook = openpyxl.Workbook()
    sanityWorkbook.save(currDir+"/"+testEnv+"_SR.xlsx")
    

main()
