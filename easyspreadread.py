#-------------------------------------------------------------------------------
# Name:        easyspreadread.py
# Purpose:
#
# Author:      Calum Leslie
#
# Created:     15/12/2012
# Copyright:   (c) Calum Leslie 2012
# Licence:     <your licence>
#-------------------------------------------------------------------------------

import xlrd

def read_sheet(filetoread):
    ''' Takes a .xls file to read and outputs cell data in the format
    of a list of lists where the innermost lists represent a row of
    information, for a 2x2 spreadsheet this would be in the format:
    [[column1/row1 value, column2/row1 value], [colum1/row2 value, column2 row2 value]]
    '''
    cellValues=[]
    updatedList=[]
    completedList=[]
    wb = xlrd.open_workbook(filetoread)
    wb.sheet_names()
    sh = wb.sheet_by_index(0)

    for rownum in range(sh.nrows):
        cellValues.append(sh.row_values(rownum))

    # encoding unicode to utf ('getting rid of u')
    for item in cellValues:
        for entry in item:
            if isinstance(entry, unicode):
                updatedList.append(entry.encode('utf-8'))
            else:
                updatedList.append(entry)
        # Putting the list in the format mentioned above
        completedList.append(updatedList)
        updatedList=[]  # clearing temporary list for each new row

    return completedList

def write_info(text_file, input_list):
    ''' Writes information taken from spreadshet to
     a file in the format of a list of lists
    '''
    f = open(text_file, 'w')
    for item in input_list:
        f.write("%s\n" % item)
    f.close()

def main():
    sheet = raw_input("Please enter the name of the spreadsheet you want to take the data from (format: testspreadsheet.xls)")
    writetxt = raw_input("Please enter the name of the text file you want to write the information to (format: test.txt)")
    write_info(writetxt, read_sheet(sheet))

if __name__ == '__main__':
    main()

##print readSheet('testsales.xls')
##writeToTxt('newtext.txt', readSheet('testsales.xls'))

