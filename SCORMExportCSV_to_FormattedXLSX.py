import xlsxwriter
import csv
import os

startDir = os.getcwd()


def folderexists():
    import os
    existsBool = False
    with os.scandir('.') as entries:
        for entry in entries:
            entryName = str(entry).split("'")[1]
            if entryName == 'SCORM Files':
                existsBool = True
    return existsBool


def isCSV(checkFile):
    if '.csv' in checkFile:
        return True
    else:
        return False


# Returns a list of CSV files (including .csv extension)
def GetFileNames(path):
    # Establish empty dictionary for function output
    FileNameList = []
    import os
    # Establish empty dictionary for Objects returned from os.scandir()
    DirRet = []
    with os.scandir(path) as entries:
        for entry in entries:
            DirEnt = str(entry)
            if isCSV(DirEnt):
                DirEnt = DirEnt.split("'")[1]
                FileNameList.append(DirEnt)
    return FileNameList


def writesheets():
    # open each csv in turn
    for fileName in GetFileNames('./'):
        # declares variable and assigns the file name without '.csv' and without underscores.
        sheetName = fileName.replace("_", " ")[:-4]
        # creates a new worksheet and assigns to currWorksheet
        currWorksheet = newWorkbook.add_worksheet(sheetName.split('-')[1])
        # ensures files are found by adding path info ('./SCORM/' or './')
        fileName = os.getcwd() + '/' + fileName

        # open the currently selected csvfile for reading
        with open(fileName, newline='') as csvfile:
            # assign contents to reader dictionary
            reader = csv.DictReader(csvfile)
            # Prepare headings
            currWorksheet.write(0, 0, 'Learner Id')
            currWorksheet.write(0, 1, 'Learner')
            currWorksheet.write(0, 2, 'Complete')
            currWorksheet.write(0, 3, 'Avg Time(Mins)*')
            # iterate target rows, writing data from iterating the reader
            longestStr = [10, 7, 8, 14]  # number is set at the title row length [3] is 15 because of the asterisk
            # prepare the percentage format object
            percentFormat = newWorkbook.add_format({'num_format': '0%'})
            noticeFormat = newWorkbook.add_format({
                'text_wrap': True,
                #'bold': True,
                # 'border': 2,
                'align': 'center',
                # 'valign': 'vcenter',

            })



            for i, readerdata in enumerate(reader):
                if len(readerdata['Learner Id']) > longestStr[0]:
                    longestStr[0] = len(readerdata['Learner Id'])
                if len(readerdata['Learner']) > longestStr[1]:
                    longestStr[1] = len(readerdata['Learner'])
                if len(readerdata['Complete']) > longestStr[3]:
                    longestStr[2] = len(readerdata['Complete'])
                # write to worksheet
                currWorksheet.write(i + 1, 0, readerdata['Learner Id'])
                currWorksheet.write(i + 1, 1, readerdata['Learner'])
                row2data = str(readerdata['Complete'])[:-1]
                # prepare for percentage format
                if row2data == '100':
                    row2data = '1'
                currWorksheet.write(i + 1, 2, int(row2data), percentFormat)
                currWorksheet.write(i + 1, 3, float(readerdata['Avg Time(Mins)']))

                # currWorksheet.write(2,5, '*Please note that Average Time does not just represent the \n amount of time '
                #                         'spent reading and interacting, but includes time \n spent with the course '
                #                         'idle and open in the background. ')
                currWorksheet.merge_range('F2:J5', '*Please note that Average Time does not just represent the amount of time '
                                                    'spent reading and interacting, but includes time spent with the course '
                                                    'idle and open in the background.', noticeFormat)
            # set widths
            for column in range(4):
                currWorksheet.set_column(column, column, longestStr[column])
           # currWorksheet.set_column(5, 5, 54)
if folderexists():
    if GetFileNames(startDir + '/SCORM Files/'):
        os.chdir(startDir + '/SCORM Files/')

if not GetFileNames('./'):
    import sys

    print("No CSV files found")
    sys.exit()
else:
    OrgName = GetFileNames('./')[0].split('-')[0]

# Start a new workbook
newWorkbook = xlsxwriter.Workbook(OrgName + '.xlsx')

writesheets()
# return to starting directory to place xlsx file on same level as script
os.chdir(startDir)
newWorkbook.close()
