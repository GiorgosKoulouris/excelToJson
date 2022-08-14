# Needed libs for venv
#     pandas
#     openpyxl

from genericpath import isdir, isfile
import pandas as pd
import sys
import os
import datetime

# Store xls full path
xlsFilePath = sys.argv[1]
xlsFilePath = os.path.abspath(xlsFilePath)

# Store log file full path
logFilePath = sys.argv[2]
logFilePath = os.path.abspath(logFilePath)

# Get the file's name in order to name the json files accordingly
xlsFileNameWithExtension = os.path.normpath(os.path.basename(xlsFilePath))
xlsFileName = os.path.splitext(xlsFileNameWithExtension)[0]

# Initialize XLS and JSON parent directories
xlsParentDirPath = os.path.abspath(os.path.join(xlsFilePath, os.pardir))
jsonFileParentDir = xlsParentDirPath

# Create the XLS object and get its needed properties
xls = pd.ExcelFile(xlsFilePath)
multiSheet = False
sheets = xls.sheet_names
sheetCount = len(sheets)

# Get current time
t = datetime.datetime.now()

# Write a title for each excel in log file
with open(logFilePath, 'a') as f:
    f.write("Excel file: " + xlsFileNameWithExtension + "\n")

# If XLS has more than 1 sheets then JSON files are created in a folder
if sheetCount > 1:
    multiSheet = True
    jsonFileParentDir = jsonFileParentDir + "/" + xlsFileName + "_jsonFiles"
    # If the direcotry already exists create a new one with timestamp
    if os.path.isdir(jsonFileParentDir):
        tFormatted = t.strftime("%Y%m%d_%H%M%S")
        jsonFileParentDir = xlsParentDirPath + "/" + xlsFileName + "_jsonFiles_" + tFormatted

    os.mkdir(jsonFileParentDir)

    with open(logFilePath, 'a') as f:
        f.write("   Folder created: " + jsonFileParentDir + "\n")

print("Converting " + xlsFileNameWithExtension + "...")

# Iterate through sheets and create their respective JSON files
for sheet in sheets:
    df = pd.read_excel(xlsFilePath, sheet_name=sheet)
    jsonString = df.to_json()

    # If it's multisheet then sheet name is included in the JSON file name scheme
    if multiSheet:
        jsonFile = jsonFileParentDir + "/" + xlsFileName + "_" + sheet + ".json"
    else:
        jsonFile = jsonFileParentDir + "/" + xlsFileName + ".json"
        # Check if the file exists. If yes, rename the JSON file by including a timestamp
        if os.path.isfile(jsonFile):
            tFormatted = t.strftime("%Y%m%d_%H%M%S")
            jsonFile = jsonFileParentDir + "/" + xlsFileName + "_" + tFormatted + ".json"

    # Write JSON content
    with open(jsonFile, 'w') as f:
        f.write(jsonString)

    # Update log file
    with open(logFilePath, 'a') as f:
        if multiSheet:
            f.write("       " + jsonFile + "\n")
        else:
            f.write("   " + jsonFile + "\n")

print ("Updated log file")

# Append a new line to separate xls entries in log file
with open(logFilePath, 'a') as f:
    f.write("\n")
