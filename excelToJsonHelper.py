# Needed libs for venv
#     pandas
#     openpyxl

from genericpath import isdir, isfile
import pandas as pd
import sys
import os
import datetime
import subprocess


# Store xls full path
xlsFilePath = sys.argv[1]
xlsFilePath = os.path.abspath(xlsFilePath)

# Get the file's name in order to name the json files accordingly
xlsFileName = os.path.splitext(os.path.normpath(os.path.basename(xlsFilePath)))[0]

# Initialize XLS and JSON parent directories
xlsParentDirPath = os.path.abspath(os.path.join(xlsFilePath, os.pardir))
jsonFileParentDir = xlsParentDirPath

# Create the XLS object and get its needed properties
xls = pd.ExcelFile(xlsFilePath)
multiSheet = False
sheets = xls.sheet_names
sheetCount = len(sheets)

# If XLS has more than 1 sheets then JSON files are created in a folder
if sheetCount > 1:
    multiSheet = True
    jsonFileParentDir = jsonFileParentDir + "/" + xlsFileName + "_jsonFiles"
    # If the direcotry already exists create a new one with timestamp
    if os.path.isdir(jsonFileParentDir):
        t = e = datetime.datetime.now()
        tFormatted = t.strftime("%Y%m%d_%H%M%S")
        jsonFileParentDir = xlsParentDirPath + "/" + xlsFileName + "_jsonFiles_" + tFormatted

    os.mkdir(jsonFileParentDir)

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
            t = e = datetime.datetime.now()
            tFormatted = t.strftime("%Y%m%d_%H%M%S")
            jsonFile = jsonFileParentDir + "/" + xlsFileName + "_" + tFormatted + ".json"
       
    # Write JSON content
    with open(jsonFile, 'w') as f:
        f.write(jsonString)

    # Popup a finder window to locate the json file
    if multiSheet:
        subprocess.call(["open", jsonFileParentDir])  
    else:
        subprocess.call(["open", "-R", jsonFile])


