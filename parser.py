import csv
#**************************************************
#Expected to be removed after argparse implemented
import getopt
#**************************************************
import openpyxl #ignored error for unidentified reference here (error happened after importing this script to git)
import os
import sys
import argparse

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
    parser = argparse.ArgumentParser()
    parser.add_argument("-i", "--iFile", help="Specifies the name of the workbook you wish to import for parsing")
    parser.add_argument("-s", "--sName", help="Specifies the name of the sheet in the workbook you wish to parse")
    parser.add_argument("-o", "--oName", help="Specifies the name of the output file you wish to write the parsed data to")
    args = parser.parse_args()

    inputFileName = args.iFile
    sheetName = args.sName
    outputFileName = args.oName

    pds = []
    sheet = []
    try:
        pds = getWorkbook(inputFileName)
    except FileNotFoundError as err:
        #print(err)
        exit(err)
    except PermissionError as err:
        #print(err)
        exit(err)
    except TypeError as err:
        #print(err)
        exit(err)
    try:
        sheet = getSheet(pds, sheetName)
    except KeyError as err:
        #print(err)
        exit(err)

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

    try:
        makeCSV(inputFileName, sheetName, outputFileName, data)
    except FileNotFoundError as err:
        #print(err)
        exit(err)
    except PermissionError as err:
        #print(err)
        exit(err)
    except TypeError as err:
        #print(err)
        exit(err)

    print("OK")
    exit()
#this only executes if modbusMapConversion.py is executed, not imported
if __name__ == "__main__":
    main(sys.argv[1:])