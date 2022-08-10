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

xlsParentDirPath = os.path.abspath(os.path.join(xlsFilePath, os.pardir))
jsonFileParentDir = xlsParentDirPath

xls = pd.ExcelFile(xlsFilePath)
multiSheet = False

sheets = xls.sheet_names
sheetCount = len(sheets)

if sheetCount > 1:
    multiSheet = True
    jsonFileParentDir = jsonFileParentDir + "/" + xlsFileName + "_jsonFiles"
    if os.path.isdir(jsonFileParentDir):
        t = e = datetime.datetime.now()
        tFormatted = t.strftime("%Y%m%d_%H%M%S")
        jsonFileParentDir = xlsParentDirPath + "/" + xlsFileName + "_jsonFiles_" + tFormatted

    os.mkdir(jsonFileParentDir)

for sheet in sheets:
    df = pd.read_excel(xlsFilePath, sheet_name=sheet)

    jsonString = df.to_json()

    if multiSheet:
        jsonFile = jsonFileParentDir + "/" + xlsFileName + "_" + sheet + ".json"
    else:
        jsonFile = jsonFileParentDir + "/" + xlsFileName + ".json"
        if os.path.isfile(jsonFile):
            t = e = datetime.datetime.now()
            tFormatted = t.strftime("%Y%m%d_%H%M%S")
            jsonFile = jsonFileParentDir + "/" + xlsFileName + "_" + tFormatted + ".json"
       
    
    with open(jsonFile, 'w') as f:
        f.write(jsonString)

    # Popup a finder window to locate the json file
    if multiSheet:
        subprocess.call(["open", jsonFileParentDir])  
    else:
        subprocess.call(["open", "-R", jsonFile])


