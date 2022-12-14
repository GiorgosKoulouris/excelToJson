#!/bin/bash

canExecute=true

echo "Validating files..."

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
    # echo "Creating python environment..."
    # venvPath="${scriptPath}/tempVenv$RANDOM"
    # # Check if a directory with the same name exists and rename venv folder if needed
    # while [ -d "${venvPath}" ]
    # do
    #     venvPath="${scriptPath}/tempVenv$RANDOM"
    # done
    # python3 -m venv $venvPath
    # source "${venvPath}/bin/activate"
    # echo "Installing dependencies..."
    # python3 -m pip install --upgrade pip
    # requirementsFile="${scriptPath}/requirements.txt"
    # python3 -m pip install -r "$requirementsFile"

    # Create the log file to pass it to the python script
    echo "Creating log file..."
    currentTimeFull=`date +"%Y%m%d_%H%M%S"`
    logFile="${scriptPath}/xlsToJason_${currentTimeFull}.txt"
    touch "$logFile"
    currentDate=`date +"%d/%m/%Y"`
    currentTime=`date +"%H:%M:%S"`
    echo -e "Script executed on ${currentDate} at ${currentTime}\n\n" >> "$logFile"

    for xlsPath in "$@"
    do
        pyScript="${scriptPath}/excelToJsonHelper.py"
        python3 "$pyScript" "$xlsPath" "$logFile"
    done

    deactivate

    # # Main Scenario
    # echo "Cleaning up..."
    # rm -rf "$venvPath"

    open -R ${logFile}

else
    echo ""
    echo "Quiting...."
fi
