@echo off


REM Check if Notepad++ exists
if exist "C:\Program Files\Notepad++\notepad++.exe" (
    echo Notepad++ found. Launching Notepad++...
    start "" "C:\Program Files\Notepad++\notepad++.exe" "C:\Program Files\Yoink Software\Yoink4CM\monthly.ps1"
) else (
    echo Notepad++ not found. Launching Notepad...
    start "" "notepad.exe" "C:\Program Files\Yoink Software\Yoink4CM\monthly.ps1"
)

REM Exit the batch file
exit /b