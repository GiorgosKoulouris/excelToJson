#!/bin/bash

canExecute=true

# Check that there is at least one argument
if [ "$#" -eq 0 ];
then
    echo "You need at least one XLS or XLSX file as an argument. Usage example:"
    echo ""
    echo "  /path/to/excelToJason.sh /path1/to1/excel1.xls /path2/to2/excel2.xls"
    canExecute=false
fi

# Check that files exist and are in a valid format
for file in "$@"
do
    if [ ! -f "$file" ]; then
        echo "File doesn't exist:"
        echo ""
        echo "  ${file}"
        canExecute=false
    fi

    filename=$(basename -- "$file")
    extension="${filename##*.}"
    if [ $extension != "xls" ] && [ $extension != "xlsx" ]
    then
        echo "File doesn't seem to be an excel file (.xls or .xlsx)"
        echo ""
        echo "  ${file}"
        canExecute=false
    fi
done

if $canExecute = true ; then

    # Get script parent Directory so you can locate its helper file
    scriptPath=$( cd -- "$( dirname -- "${BASH_SOURCE[0]}" )" &> /dev/null && pwd )

    # Scenario 2
    source /Users/louris/Dev_Projects/Coding/Scripts/excelToJson/venv/bin/activate

    # # Main Scenario
    # venvPath="${scriptPath}/tempVenv$RANDOM"
    # python3 -m venv $venvPath
    # source "${venvPath}/bin/activate"
    # python3 -m pip install --upgrade pip
    # python3 -m pip install pandas
    # python3 -m pip install openpyxl

    for xlsPath in "$@"
    do
        pyScript="${scriptPath}/excelToJsonHelper.py"
        python3 "$pyScript" "$xlsPath"
    done

    deactivate

    # # Main Scenario
    # rm -rf $venvPath

else
    echo ""
    echo "Quiting...."
fi

