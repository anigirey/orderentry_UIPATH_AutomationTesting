"""

"""


import xlsxwriter
import xlrd
import os
from os import path


def main():

    # variables for environments and imports
    #now = datetime.now()
    #currentdate = now.strftime("%Y-%m-%d")
    #timestamp = now.strftime("%H.%M.%S")
    #currDir = os.getcwd()
    #rDateStr = (str(currentdate))
    #rTimeStr = (str(timestamp))
    #directory = ''

    currDir = os.getcwd()
    loc = currDir + ("/Output/outputvars.xlsx")
    wBook = xlrd.open_workbook(loc)
    wSheet = wBook.sheet_by_index(0)
    testEnv = str(wSheet.cell_value(0,3))
    testTim = str(wSheet.cell_value(0,1))
    
    sanityWorkbook = xlsxwriter.Workbook(currDir+"/"+testEnv+"_SR.xlsx")
    sanityWorksheet = sanityWorkbook.add_worksheet("DUST")
    sanityWorkbook.close()

    # loop to create folders for each environment and each test case within
    # those environments
    #for env in testEnv:
     #   for case in testCase:
      #      directory = directoryCreate(env, case, currentdate, rTimeStr,
       #                                 currDir)
        #    sanityResultCreate(env, testCase, currentdate, rTimeStr, currDir)
         #   pageResultCreate(currentdate, env, case, rTimeStr, currDir)

# function for creating the directory structure
def directoryCreate(env, case, currentdate, rTimeStr, currDir):

    date = currentdate

    directory = (str(currDir) + "/Results/" + date + "/" + env + "/" + case +
                 "_" + rTimeStr + "/")
    os.makedirs(directory)

    # <directory> variable returned as a string for later use
    return directory

# function for creating the SanityResult excel file for use in each environment
# folder


def sanityResultCreate(env, testCase, currentdate, rTimeStr, currDir):

    # variables created to use the currentdate and timestamp as strings
    date = currentdate
    
    # <if> statement used to check if a file already exists. if no file exists,
    # it will create a new one
    if not path.exists(str(currDir) + "/Results/" + date + "/" + env + "/" +
                       env + "_SanityResult_" + date + ".xlsx"):
        # SanityResult excel file creation
        sanityResultWorkbook = xlsxwriter.Workbook(str(currDir) + "/Results/" +
                                                   date + "/" + env + "/" +
                                                   env + "_SanityResult_" +
                                                   date + ".xlsx")

        # <for> loop used to create a worksheet for each test case in a single
        # environment SanityResults excel file
        for i in testCase:

            # worksheet creation inside of SanityResults excel file
            sanityResultWorksheet = sanityResultWorkbook.add_worksheet(i + " Results")

            # column formatting
            sanityResultWorksheet.set_column("A:A", 12)
            sanityResultWorksheet.set_column("B:B", 12)
            sanityResultWorksheet.set_column("C:C", 12)
            sanityResultWorksheet.set_column("D:D", 12)
            sanityResultWorksheet.set_column("E:E", 12)
            sanityResultWorksheet.set_column("F:F", 12)
            sanityResultWorksheet.set_column("G:G", 12)
            sanityResultWorksheet.set_column("H:H", 12)
            sanityResultWorksheet.set_column("I:I", 12)
            bold = sanityResultWorkbook.add_format({"bold": True})

            # initial cell population with header information formatted for
            # Row 1
            sanityResultWorksheet.write("A1", "Order Number", bold)
            sanityResultWorksheet.write("B1", "TN", bold)
            sanityResultWorksheet.write("C1", "CPlus Result", bold)
            sanityResultWorksheet.write("D1", "WebSOP Result", bold)
            sanityResultWorksheet.write("E1", "IOM Data", bold)
            sanityResultWorksheet.write("F1", "Destination System", bold)
            sanityResultWorksheet.write("G1", "Order Status", bold)
            sanityResultWorksheet.write("H1", "OBAN Result", bold)
            sanityResultWorksheet.write("I1", "Time Stamp", bold)

            # # initial cell population with test environment information
            # # formatted for Column A
            # sanityResultWorksheet.write("A2", "No Data")
            # sanityResultWorksheet.write("B2", "No Data")
            # sanityResultWorksheet.write("C2", "No Data")
            # sanityResultWorksheet.write("D2", "No Data")
            # sanityResultWorksheet.write("E2", "No Data")
            # sanityResultWorksheet.write("F2", "No Data")
            # sanityResultWorksheet.write("G2", "No Data")
            # sanityResultWorksheet.write("H2", "None")
            # sanityResultWorksheet.write("I2", rTimeStr)

        # File must be closed here so that it can be created
        sanityResultWorkbook.close()

# function for creating the PageResult excel file for use in each test case
# folder


def pageResultCreate(currentdate, env, case, rTimeStr, currDir):

    date = currentdate

    # PageResults excel file creation
    pageResultWorkbook = xlsxwriter.Workbook(str(currDir) + "/Results/" +
                                             date + "/" + env + "/" + case +
                                             "_" + rTimeStr+ "/" +
                                             "PageResult.xlsx")
    # worksheet creation inside of PageResults excel file
    pageResultWorksheet = pageResultWorkbook.add_worksheet(case + " Results")

    # column formatting
    pageResultWorksheet.set_column("A:A", 40)
    pageResultWorksheet.set_column("B:B", 15)
    bold = pageResultWorkbook.add_format({"bold": True})

    # initial cell population with header information formatted for Column A
    pageResultWorksheet.write("A1", "Step Detail", bold)
    pageResultWorksheet.write("A2", "Sign On Page opened")
    pageResultWorksheet.write("A3", "Home Page opened")
    pageResultWorksheet.write("A4", "Customer Information Page opened")
    pageResultWorksheet.write("A5", "Win Back/Win Over Page opened")
    pageResultWorksheet.write("A6", "Service Address Validation Page opened")
    pageResultWorksheet.write("A7", "Facility Check and Results Page opened")
    pageResultWorksheet.write("A8", "Primary Listing Page opened")
    pageResultWorksheet.write("A9", "Product Pricing Page opened")
    pageResultWorksheet.write("A10", "Billing Information Page opened")
    pageResultWorksheet.write("A11", "Business Credit Application Page opened")
    pageResultWorksheet.write("A12", "Credit Information Page opened")
    pageResultWorksheet.write("A13", "Credit Decision Page opened")
    pageResultWorksheet.write("A14", "Service and Equipment Page opened")
    pageResultWorksheet.write("A15", "Product Summary Page opened")
    pageResultWorksheet.write("A16", "Configure Product Page opened")
    pageResultWorksheet.write("A17", "Configure Order Page opened")
    pageResultWorksheet.write("A18", "Configure OLFIDs Page opened")
    pageResultWorksheet.write("A19", "Appointment Scheduler Page opened")
    pageResultWorksheet.write("A20", "Deposit/Advance Payment Page opened")
    pageResultWorksheet.write("A21", "Deposit/Advance Payment Information Page opened")
    pageResultWorksheet.write("A22", "Deposit/Advance Payment Success Page opened")
    pageResultWorksheet.write("A23", "Order Detail Page opened")
    pageResultWorksheet.write("A24", "Order Validation Confirmation")
    pageResultWorksheet.write("A26", "Reason for Failure")

    # initial cell population with default test case information formatted for
    # Column A
    pageResultWorksheet.write("B1", "Step Status", bold)
    pageResultWorksheet.write("B2", "Not Tested")
    pageResultWorksheet.write("B3", "Not Tested")
    pageResultWorksheet.write("B4", "Not Tested")
    pageResultWorksheet.write("B5", "Not Tested")
    pageResultWorksheet.write("B6", "1FB Only")
    pageResultWorksheet.write("B7", "1FB Only")
    pageResultWorksheet.write("B8", "1FB Only")
    pageResultWorksheet.write("B9", "Not Tested")
    pageResultWorksheet.write("B10", "Not Tested")
    pageResultWorksheet.write("B11", "1FB Only")
    pageResultWorksheet.write("B12", "Not Tested")
    pageResultWorksheet.write("B13", "Not Tested")
    pageResultWorksheet.write("B14", "1FB Only")
    pageResultWorksheet.write("B15", "1FB Only")
    pageResultWorksheet.write("B16", "1FB Only")
    pageResultWorksheet.write("B17", "Not Tested")
    pageResultWorksheet.write("B18", "1FB Only")
    pageResultWorksheet.write("B19", "Not Tested")
    pageResultWorksheet.write("B20", "Not Tested")
    pageResultWorksheet.write("B21", "Not Tested")
    pageResultWorksheet.write("B22", "Not Tested")
    pageResultWorksheet.write("B23", "Not Tested")
    pageResultWorksheet.write("B24", "Not Tested")
    pageResultWorksheet.write("B26", "PASSED")

    # File must be closed here so that it can be created
    pageResultWorkbook.close()


main()
