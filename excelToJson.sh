#!/bin/bash

# Get script parent Directory so you can locate its helper file
scriptPath=$( cd -- "$( dirname -- "${BASH_SOURCE[0]}" )" &> /dev/null && pwd )

xlsPath=$1
xlsParentDir="$(dirname "$xlsPath")"

# Delete Later
source /Users/louris/Dev_Projects/Coding/Scripts/excelToJson/venv/bin/activate

# venvPath="${scriptPath}/tempVenv$RANDOM"
# python3 -m venv $venvPath
# source "${venvPath}/bin/activate"
# python3 -m pip install --upgrade pip
# python3 -m pip install pandas
# python3 -m pip install openpyxl

command="${scriptPath}/excelToJsonHelper.py ${xlsPath}"
python3 $command

deactivate

# rm -rf $venvPath
