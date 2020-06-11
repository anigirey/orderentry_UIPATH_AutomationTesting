"""

"""


import openpyxl
import xlrd
import os
from datetime import datetime
from os import path


def main():

    now = datetime.now()
    currDir = os.getcwd()
    loc = currDir + ("/Output/outputvars.xlsx")
    wBook = xlrd.open_workbook(loc)
    wSheet = wBook.sheet_by_index(0)
    testEnv = wSheet.cell_value(0,3)
    testTim = wSheet.cell_value(0,1)

    sanityWorkbook = xlsxwriter.Workbook(str(currDir) + "/" + testEnv + "_SR.xlsx")
    sanityWorksheet = sanityWorkbook.add_worksheet(testTim)
    sanityWorksheet.set_column("A:A", 15)
    sanityWorksheet.set_column("B:B", 15)
    sanityWorksheet.set_column("C:C", 15)
    sanityWorksheet.set_column("D:D", 15)                                         
    sanityWorksheet.set_column("E:E", 15)
    sanityWorksheet.set_column("F:F", 15)
    sanityWorksheet.set_column("G:G", 15)
    sanityWorksheet.set_column("H:H", 15)
    sanityWorksheet.set_column("I:I", 15)
    sanityWorksheet.set_column("J:J", 15)
    
    bold = sanityWorkbook.add_format({"bold": True})
    cent = sanityWorkbook.add_format({"center_across": True})
    bold.set_bold()
    cent.set_center_across()
    
    sanityWorksheet.write("A1", "Order Type", cent)
    sanityWorksheet.write("B1", "Order Number", cent)
    sanityWorksheet.write("C1", "TN", cent)
    sanityWorksheet.write("D1", "CPlus Result", cent)
    sanityWorksheet.write("E1", "WebSOP Result", cent)
    sanityWorksheet.write("F1", "IOM Data", cent)
    sanityWorksheet.write("G1", "Destination System", cent)
    sanityWorksheet.write("H1", "Order Status", cent)
    sanityWorksheet.write("I1", "OBAN Result", cent)
    sanityWorksheet.write("J1", "Time Stamp", cent)
    
    sanityWorkbook.close()


main()
