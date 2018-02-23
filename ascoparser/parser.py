#######################################################################################################################
#parser.py
#Script for parsing .xlsx PDS and reformatting it for further use by other scripts or users
#
#A module of M3XG Communications Regression suite
#
#Copyright []
#
#Version 1.0
#   1. Initial Release (JAB 2/2/2018)
#
#######################################################################################################################

#----------------------------------------------------------------------------------------------------------------------
#Standard library imports
import sys
import argparse
import collections
import string
from datetime import date
from csv import writer
#----------------------------------------------------------------------------------------------------------------------

#----------------------------------------------------------------------------------------------------------------------
#3rd party library imports
from openpyxl import load_workbook #absolute import specified as per PEP8
#----------------------------------------------------------------------------------------------------------------------

#----------------------------------------------------------------------------------------------------------------------
#Fuction: read in text file. Expected usage is for user specified strings that correlate to which columns to keep from pds
def readTxt(txtFile):
    with open(txtFile, 'r') as f:
        headerList = []
        for line in f:
            lineClean = line.replace("\n", '')
            headerList.append(lineClean)
        f.close()
    return headerList
#----------------------------------------------------------------------------------------------------------------------

#----------------------------------------------------------------------------------------------------------------------
#Function: writes a zipped list to a csv file
def writeCSV(inputFileName, sheetName, outputFileName, table):
    with open(outputFileName, 'w', newline='') as f:
        w = writer(f)
        w.writerow([inputFileName])  # add a header with the path used to create this csv
        w.writerow([sheetName])  # add header with the sheet name used to create this csv
        w.writerow([date.today()])
        w.writerow([''])
        for row in table:  # for each row of data, write it to the csv
            w.writerow(row)
        f.close()
#----------------------------------------------------------------------------------------------------------------------

#----------------------------------------------------------------------------------------------------------------------
#Function: Writes a dictionary<Parameter Name, dictionary<Header, Cell Value>> to a .txt file
def writeTxt(inputFileName, device, outputFileName, outputData):
    with open(outputFileName, 'w+') as f:
        f.writelines(inputFileName + "\n")
        f.writelines(str(device) + "\n")
        f.writelines(str(date.today())+ "\n")
        f.writelines("\n")
        for x in outputData:
            f.writelines(str(x) + "\n")
            for y in outputData[x]:
                f.writelines(str(y) + ':' + str(outputData[x][y]) + "\n")
            f.writelines("\n")
        f.close()
#----------------------------------------------------------------------------------------------------------------------

#----------------------------------------------------------------------------------------------------------------------
#Function: gets specified excel workbook, gets a specified sheet in that workbook, returns that sheet
def getSheet(inputFileName, sheetName):
    try:
        pds = load_workbook(inputFileName, read_only=True,data_only=True, keep_vba= False, keep_links=False)
        #print("Got workbook. Attempting sheet extraction.")
        try:
            sheet = pds.get_sheet_by_name(sheetName)
            return sheet
        except KeyError as err:
            exit(err)
    except FileNotFoundError as err:
        exit(err)
    except PermissionError as err:
        exit(err)
    except TypeError as err:
        exit(err)
#----------------------------------------------------------------------------------------------------------------------

#----------------------------------------------------------------------------------------------------------------------
#Function: convert a list of rows into a dictionary of dictionaries
def listOfRowsToDictionary(table):
    dictHeaders = table[0]
    parameterNames = []
    pNameIndex = -1
    for i in range(0, len(dictHeaders)):
        header = dictHeaders[i]
        if header == 'Parameter Name':
            pNameIndex = i
    rows = []
    for i in range(1, len(table)):
        row = table[i]
        #parameterName = row[pNameIndex]
        parameterNames.append(row[pNameIndex])
        rowData = collections.OrderedDict(zip(dictHeaders, table[i]))
        rows.append(rowData)
    dictionary = collections.OrderedDict(zip(parameterNames, rows))
    return dictionary
#----------------------------------------------------------------------------------------------------------------------

#----------------------------------------------------------------------------------------------------------------------
#Function: Converts sheet of cells to list of rows, where rows contain values of cells
def cellsToValues(sheet):
    data = []
    for row in sheet.rows:
        rowValues = []
        for cell in row:
            value = cell.value
            value = str.replace(str(value), '\n', ' ')
            rowValues.append(value)
        data.append(rowValues)
    del data[0:7]
    #for row in data:
    #    print(row)
    return data
#----------------------------------------------------------------------------------------------------------------------

#----------------------------------------------------------------------------------------------------------------------
#Function: input a list of [registers:startbit-stopbit]. output list of lists [register, bit, length]
def splitRegisterContent(listOfReg):
    registerAddress = []
    startBit = []
    numOfBits = []
    registerAddress.append("Register Address")
    startBit.append("Start Bit")
    numOfBits.append("Bitwise Length")
    #need to skip first row of list
    for i in range(1, len(listOfReg)):
        val = str(listOfReg[i])
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
#----------------------------------------------------------------------------------------------------------------------

#======================================================================================================================
#entry point if calling the script directly
def main(argv):
    #------------------------------------------------------------------------------------------------------------------
    #Input Args
    parser = argparse.ArgumentParser()
    parser.add_argument("-f", "--fileType", required=True, help="Specifies the file type of the output file")
    parser.add_argument("-i", "--inputName", required=True, help="Specifies the name of the workbook you wish to import for parsing")
    parser.add_argument("-s", "--sheetName", required=True, help="Specifies the name of the sheet in the workbook you wish to parse")
    parser.add_argument("-c", "--columns", help="Specifies the name of the text file with a list of columns the user wants in the output")
    parser.add_argument("-o", "--outputName", required=True, help="Specifies the name of the output file you wish to write the parsed data to")
    args = parser.parse_args()
    outputFileType = args.fileType
    inputPDSName = args.inputName
    sheetName = args.sheetName
    inputHeadersName = args.columns
    outputFileName = args.outputName
    #------------------------------------------------------------------------------------------------------------------

    #------------------------------------------------------------------------------------------------------------------
    #get worksheet and store it in memory
    sheet = getSheet(inputPDSName, sheetName)
    #------------------------------------------------------------------------------------------------------------------

    #------------------------------------------------------------------------------------------------------------------
    #convert cell objects to values at those cells
    data = cellsToValues(sheet) # we now have a list of lists (rows)
    #------------------------------------------------------------------------------------------------------------------

    #------------------------------------------------------------------------------------------------------------------
    #Row dependent operations go here
    #First Row is headers, store that row for later use
    originalHeaders = data[0]
    #find the index of the header that contains Parameter Name and store it for future use
    paramNameIndex = -1
    for i in range(0, len(originalHeaders)):
        if originalHeaders[i] == "Parameter Description":
            paramDescriptionIndex = i
    #Remove Parameters that are undefined or reserved
    for row in data:
        if ("Reserved" in row[paramDescriptionIndex]) | ("Undefined" in row[paramDescriptionIndex]):
            data.remove(row)
            #print(row[paramNameIndex] + " removed!")
    #------------------------------------------------------------------------------------------------------------------

    #------------------------------------------------------------------------------------------------------------------
    # Column dependent operations go here
    # convert from rows to columns
    data = zip(*data)
    # default to keeping all columns if no --columns arg specified
    desiredHeaders = []
    if inputHeadersName != None:
        try:
            desiredHeaders = readTxt(inputHeadersName)
        except:
            desiredHeaders = originalHeaders
            desiredHeaders = [x for x in desiredHeaders if x != None]
            print(desiredHeaders)
            print("No default header values found. Using Defaults")
    else:
        desiredHeaders = originalHeaders
        desiredHeaders = [x for x in desiredHeaders if x != "None"]
        print(desiredHeaders)
        print("No default header values found. Using Defaults")

    #checks for headers that actually exist
    matchingHeaders = []
    for header in desiredHeaders:
        if header in originalHeaders:
            matchingHeaders.append(header)

    #if the user specified a register address column, extra parsing is required otherwise, return everything
    finalDataTable = []
    if 'Register Address' in matchingHeaders:
        for column in data:
            header = column[0]
            if header == 'Register Address':
                splitRegister = splitRegisterContent(column)
                finalDataTable.append(splitRegister[0])
                finalDataTable.append(splitRegister[1])
                finalDataTable.append(splitRegister[2])
            else:
                for name in matchingHeaders:
                    if (name == header) & (header != 'Register Address'):
                        finalDataTable.append(column)
    else:
        #no registerColumn specified. don't need to unpack a register column then, just build a list without it
        for column in data:
            header = column[0]
            for name in matchingHeaders:
                if name == header:
                    finalDataTable.append(column)#will only add columns to the final data table if the headers match
    #------------------------------------------------------------------------------------------------------------------

    #------------------------------------------------------------------------------------------------------------------
    #make a csv or txt based on user input
    try:
        if outputFileType == "csv":
            finalDataTable = zip(*finalDataTable)
            writeCSV(inputPDSName, sheetName, outputFileName, finalDataTable)
        elif outputFileType == "txt":
            finalDataTable = zip(*finalDataTable)  # to rows
            finalDataTableUnzipped = [list(row) for row in finalDataTable]  # operations to create dictionary don't play well with zips
            pdsDictionary = listOfRowsToDictionary(finalDataTableUnzipped)
            writeTxt(inputPDSName, sheetName, outputFileName, pdsDictionary)
        else:
            #maybe a default would be nice?
            exit("No valid output file type specified.")
    except FileNotFoundError as err:
        exit(err)
    except PermissionError as err:
        exit(err)
    except TypeError as err:
        exit(err)
    #------------------------------------------------------------------------------------------------------------------

    #------------------------------------------------------------------------------------------------------------------
    #Script Terminus
    print("OK")
    exit(0)
    #------------------------------------------------------------------------------------------------------------------
#======================================================================================================================

if __name__ == "__main__":
    main(sys.argv[1:])


