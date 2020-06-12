"""

"""

import xlsxwriter
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, Alignment
import xlrd
import os
from os import path


def main():

    currDir = os.getcwd()
    loc = currDir + ("/Output/outputvars.xlsx")
    wBook = xlrd.open_workbook(loc)
    wSheet = wBook.sheet_by_index(0)
    testEnv = 'SFC'
    testTim = str(wSheet.cell_value(0,1))
    
    sfcWorkbook = xlsxwriter.Workbook(str(testEnv+'_SR.xlsx'))
    sfcWorksheet = sfcWorkbook.add_worksheet(testTim)
    sfcWorkbook.close()
    
    sfcWorkbook = openpyxl.load_workbook(str(testEnv+'_SR.xlsx'))
    sfcWorksheet = sfcWorkbook.get_sheet_by_name(testTim)
    celBold = Font(size=11,bold=True,underline='single')
    celCent = Alignment(horizontal='center')
    
    sfcWorksheet.column_dimensions['A'].width = 14
    sfcWorksheet['A1'].font = celBold
    sfcWorksheet['A1'].alignment = celCent
    sfcWorksheet['A1'] = 'CSM/SMB'

    sfcWorksheet.column_dimensions['B'].width = 14
    sfcWorksheet['B1'].font = celBold
    sfcWorksheet['B1'].alignment = celCent
    sfcWorksheet['B1'] = 'Order Type'

    sfcWorksheet.column_dimensions['C'].width = 14
    sfcWorksheet['C1'].font = celBold
    sfcWorksheet['C1'].alignment = celCent
    sfcWorksheet['C1'] = 'Name'

    sfcWorksheet.column_dimensions['D'].width = 14
    sfcWorksheet['D1'].font = celBold
    sfcWorksheet['D1'].alignment = celCent
    sfcWorksheet['D1'] = 'Order Num'

    sfcWorksheet.column_dimensions['E'].width = 16
    sfcWorksheet['E1'].font = celBold
    sfcWorksheet['E1'].alignment = celCent
    sfcWorksheet['E1'] = 'TN'

    sfcWorksheet.column_dimensions['F'].width = 15
    sfcWorksheet['F1'].font = celBold
    sfcWorksheet['F1'].alignment = celCent
    sfcWorksheet['F1'] = 'Order Status'

    sfcWorkbook.save(testEnv +'_SR.xlsx')
    sfcWorkbook.close()    


main()
