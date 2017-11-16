import csv
import getopt
import openpyxl #ignored error for unidentified reference here (error happened after importing this script to git)
import os
import sys
import argparse

#testing commit

#input: path to excel workbook file. output: workbook object
#!What if path doesn't exist?
def getWorkbook(workbookName):
    workbook = openpyxl.load_workbook(workbookName)
    return workbook

#input: workbook object, name of sheet in workbook object. output: sheet object
#!What if sheet doesn't exist?
def getSheet(workbook, sheetName):
    sheet = workbook.get_sheet_by_name(sheetName)
    return sheet

#input: input file path, sheet name, output file path, data to be written. csv created or overwritten to output file path
#!what if output file path doesn't exist?
def makeCSV(inputFileName, sheetName, outputFileName, table):
    with open(outputFileName, 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerow([inputFileName]) #add a header with the path used to create this csv
        writer.writerow([sheetName]) #add header with the sheet name used to create this csv
        writer.writerow([''])
        for row in table: #for each row of data, write it to the csv
            writer.writerow(row)
        f.close()

#input a list of [registers:startbit-stopbit]. output list of lists [register, bit, length]
def splitRegisterContent(listOfReg):
    registerAddress = []
    startBit = []
    numOfBits = []
    registerAddress.append("Register Address")
    startBit.append("Start Bit")
    numOfBits.append("Bitwise Length")

    #need to skip first row of list
    for i in range(1, len(listOfReg)):
        val = listOfReg[i]
        registerAddress.append(val[:5])
        if ':' in val:
            if '-' in val:  # multiple bits
                startBit.append(str(val[(val.index('-') - 2):(val.index('-'))]))
                numOfBits.append(str((int(val[(val.index('-') + 1):(val.index('-') + 3)]) - int(
                    val[(val.index('-') - 2):(val.index('-'))]))))
            else:  # single bit
                startBit.append(str(val[(val.index(':') + 1):(val.index(':') + 3)]))
                numOfBits.append("1")
        else:
            startBit.append("0")
            numOfBits.append("All")
    return [registerAddress, startBit, numOfBits]

def main(argv):
    helpString = """
    Usage: parser.py [INPUT OPTION] <inputFilePath> [SHEET OPTION] <sheetName> [OUTPUT OPTION] <outputFileName>
    Mandatory arguments to long options are manadatory for short options too.
    -i, --iFile     Specifies the name of the workbook you wish to import for parsing
                    Looks in current directory by default but will also accept a path
    -s, --sName     Specifies the name of the sheet in the workbook you wish to parse
    -o, --oFile     Specifies the name of the output file you wish to write the parsed data to (overwrites or creates new)
                    Looks in current directory by default but will also accept a path
    -h, --help      Display this help and exit
    
    """

    inputFileName = ''
    sheetName = ''
    outputFileName = ''
    try:
        opts, args = getopt.getopt(argv, 'hi:s:o:', ['help', 'iFile=', 'sName=', 'oName='])
    except getopt.GetoptError:
        print(getopt.GetoptError.msg)
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h' or opt == '--help':
            print(helpString)
            sys.exit()
        elif opt in ('-i', '--iFile'):
            inputFileName = arg
        elif opt in ('-s', '--sName'):
            sheetName = arg
        elif opt in ('-o', '--oName'):
            outputFileName = arg

    #get the sheet of interest from the pds
    pds = getWorkbook(inputFileName)
    sheet = getSheet(pds, sheetName)
    sheetAsTuple = tuple(sheet.columns)

    #need to create a list with columns full of values, not cell objects
    sheetValues = []
    sheetValuesAsString = []
    for column in sheetAsTuple:
        columnValues = []
        columnValuesAsString = []
        for cell in column:
            columnValues.append(cell.value)
            columnValuesAsString.append(str(cell.value))
        sheetValues.append(columnValues)
        sheetValuesAsString.append(columnValuesAsString)

    #prepare to store content extracted from the 'Register Address' column of the pds (register, start bit, bit length)
    splitRegisterContents = splitRegisterContent(sheetValuesAsString[0])
    data = zip(splitRegisterContents[0], splitRegisterContents[1], splitRegisterContents[2],
       sheetValues[3], sheetValues[5], sheetValues[7], sheetValues[8],
                    sheetValues[9], sheetValues[10], sheetValues[11], sheetValues[13])

    makeCSV(inputFileName, sheetName, outputFileName, data)

#this only executes if modbusMapConversion.py is executed, not imported
if __name__ == "__main__":
    main(sys.argv[1:])