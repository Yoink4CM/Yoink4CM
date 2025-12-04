@echo off
REM This batch file copies the necessary files and folders for Yoink Software.

REM Define source directory (assuming the batch file is in the parent directory of "Yoink4CM" and "Yoink Software.xml")
REM You might need to adjust this if the batch file is in a different location.
set "SOURCE_DIR=%~dp0"

REM Define destination paths
set "DEST_DIR_YOINK_SOFTWARE=C:\Program Files\Yoink Software"
set "DEST_DIR_YOINK_CM=C:\Program Files\Yoink Software\Yoink4CM"
set "DEST_DIR_XML_1=C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\XmlStorage\Extensions\Actions\3ad39fd0-efd6-11d0-bdcf-00a0c909fdd7"
set "DEST_DIR_XML_2=C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\XmlStorage\Extensions\Actions\d2e2cba7-98f5-4d3b-bc2f-b670f0621207"

echo.
echo --- Starting Yoink Software File Copy Operations ---
echo.

REM 1. Copy the subfolder "Yoink4CM" to "C:\Program Files\Yoink Software\Yoink4CM"
echo Creating destination directory if it doesn't exist: "%DEST_DIR_YOINK_CM%"
md "%DEST_DIR_YOINK_CM%" 2>nul
echo Copying "Yoink4CM" folder...
xcopy "%SOURCE_DIR%Yoink4CM" "%DEST_DIR_YOINK_CM%\" /E /I /H /K /Y
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Failed to copy Yoink4CM folder.
) else (
    echo Yoink4CM folder copied successfully.
)
echo.

REM Copy config.ps1 to the correct location
echo Creating destination directory if it doesn't exist: "%DEST_DIR_YOINK_SOFTWARE%"
md "%DEST_DIR_YOINK_SOFTWARE%" 2>nul
echo Copying "config.ps1" to "%DEST_DIR_YOINK_SOFTWARE%"...
copy "%SOURCE_DIR%config.ps1" "%DEST_DIR_YOINK_SOFTWARE%\" /Y
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Failed to copy config.ps1. Please ensure config.ps1 exists in the source directory.
) else (
    echo config.ps1 copied successfully to "%DEST_DIR_YOINK_SOFTWARE%".
)
echo.

REM Define the path to config.ps1
set "CONFIG_PS1_PATH=%DEST_DIR_YOINK_SOFTWARE%\config.ps1"

cls

REM Ask for Site Code and append to config.ps1
:GET_SITE_CODE
set "SITE_CODE="
set /p "SITE_CODE=Enter the 3-character Site Code (e.g., ABC): "

REM Validate that the input is exactly 3 characters
if not defined SITE_CODE (
    echo ERROR: Site Code cannot be empty.
    goto GET_SITE_CODE
)
if "%SITE_CODE:~3,1%"=="" (
    REM Check if the length is exactly 3
    if not "%SITE_CODE:~2,1%"=="" (
        echo Site Code entered: %SITE_CODE%
    ) else (
        echo ERROR: Site Code must be exactly 3 characters.
        goto GET_SITE_CODE
    )
) else (
    echo ERROR: Site Code must be exactly 3 characters.
    goto GET_SITE_CODE
)

REM Append a colon to the site code
set "SITE_CODE=%SITE_CODE%:"
echo Final Site Code with colon: %SITE_CODE%
echo.

REM Append the $SiteCode variable to config.ps1
echo Appending Site Code to "%CONFIG_PS1_PATH%"...
echo $SiteCode = "%SITE_CODE%" >> "%CONFIG_PS1_PATH%"
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Failed to append Site Code to config.ps1.
) else (
    echo Site Code appended successfully to config.ps1.
)
echo.

cls

REM --- Ask for Network Share Name and append to config.ps1 ---
:GET_NETWORK_SHARE
set "NETWORK_SHARE="

set /p "NETWORK_SHARE=Enter your network share name where applications and packages will be stored (e.g., \\Server\Share): "

if not defined NETWORK_SHARE (
    echo ERROR: Network Share Name cannot be empty.
    goto GET_NETWORK_SHARE
)

REM Check if the network share name ends with a backslash and remove it
if "%NETWORK_SHARE:~-1%"=="\" (
    set "NETWORK_SHARE=%NETWORK_SHARE:~0,-1%"
    echo Removed trailing backslash from Network Share Name.
)

REM Add "\Auto Update" subfolder to the network share name
set "NETWORK_SHARE=%NETWORK_SHARE%\Auto Update\"

echo Network Share Name (with Auto Update) entered: %NETWORK_SHARE%
echo.

REM Append the $NetworkShare variable to config.ps1
echo Appending Network Share Name to "%CONFIG_PS1_PATH%"...
echo $NetworkShare = "%NETWORK_SHARE%" >> "%CONFIG_PS1_PATH%"
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Failed to append Network Share Name to config.ps1.
) else (
    echo Network Share Name appended successfully to config.ps1.
)
echo.

cls

REM --- Ask for Device Collection ID and append to config.ps1 ---
:GET_DEVICE_COLLECTION_ID
set "DEVICE_COLLECTION_ID="
echo Enter the Device Collection ID for test deployments (e.g., ABC00001).
echo Press 1 and then enter if you do not wish to automatically deploy to a test collection (can be changed later).
set /p "DEVICE_COLLECTION_ID=Enter the Device Collection ID: "

if not defined DEVICE_COLLECTION_ID (
    echo ERROR: Device Collection ID cannot be empty.
    goto GET_DEVICE_COLLECTION_ID
)
echo Device Collection ID entered: %DEVICE_COLLECTION_ID%
echo.

REM Append the $DeviceCollectionID variable to config.ps1
echo Appending Device Collection ID to "%CONFIG_PS1_PATH%"...
echo $DeviceCollectionID = "%DEVICE_COLLECTION_ID%" >> "%CONFIG_PS1_PATH%"
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Failed to append Device Collection ID to config.ps1.
) else (
    echo Device Collection ID appended successfully to config.ps1.
)
echo.

cls

REM --- Ask for Distribution Point Group Name and append to config.ps1 ---
:GET_DP_GROUP_NAME
set "DP_GROUP_NAME="
set /p "DP_GROUP_NAME=Enter the Distribution Point Group Name new packages and applications should be distributed to (e.g., Paris Office): "

if not defined DP_GROUP_NAME (
    echo ERROR: Distribution Point Group Name cannot be empty.
    goto GET_DP_GROUP_NAME
)
echo Distribution Point Group Name entered: %DP_GROUP_NAME%
echo.

REM Append the $DPGroupName variable to config.ps1
echo Appending Distribution Point Group Name to "%CONFIG_PS1_PATH%"...
echo $DistributionPointGroup = "%DP_GROUP_NAME%" >> "%CONFIG_PS1_PATH%"
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Failed to append Distribution Point Group Name to config.ps1.
) else (
    echo Distribution Point Group Name appended successfully to config.ps1.
)
echo.

cls


REM --- Ask for Device Collection Creation and append to config.ps1 ---

:GET_CREATE_DEVICE_COLLECTIONS
set "CreateDeviceCollections="
set /p "CreateDeviceCollections=Create a new Device Collection for each application/package downloaded? (y/n): "

REM Validate that the input is 'y' or 'n' (case-insensitive)
if /i not "%CreateDeviceCollections%"=="y" if /i not "%CreateDeviceCollections%"=="n" (
    echo ERROR: Invalid input. Please enter 'y' for Yes or 'n' for No.
    goto GET_CREATE_DEVICE_COLLECTIONS
)

REM Standardize to lowercase 'y' or 'n' for the variable
if /i "%CreateDeviceCollections%"=="Y" set "CreateDeviceCollections=y"
if /i "%CreateDeviceCollections%"=="N" set "CreateDeviceCollections=n"

echo Create Device Collections setting entered: %CreateDeviceCollections%
echo.

REM Append the $CreateDeviceCollections variable to config.ps1
echo Appending Create Device Collections setting to "%CONFIG_PS1_PATH%"...
echo $CreateDeviceCollections = "%CreateDeviceCollections%" >> "%CONFIG_PS1_PATH%"
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Failed to append Create Device Collections setting to config.ps1.
) else (
    echo Create Device Collections setting appended successfully to config.ps1.
)
echo.
REM --------------------------------------------------------------------------

REM 2. Copy "Yoink Software.xml" to "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\XmlStorage\Extensions\Actions\3ad39fd0-efd6-11d0-bdcf-00a0c909fdd7\Yoink Software.xml"
echo Creating destination directory if it doesn't exist: "%DEST_DIR_XML_1%"
md "%DEST_DIR_XML_1%" 2>nul
echo Copying "Yoink Software.xml" to first destination...
copy "%SOURCE_DIR%Yoink Software.xml" "%DEST_DIR_XML_1%\" /Y
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Failed to copy Yoink Software.xml to "%DEST_DIR_XML_1%\".
) else (
    echo Yoink Software.xml copied successfully to first location.
)
echo.

REM 3. Copy "Yoink Software.xml" again to "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\XmlStorage\Extensions\Actions\d2e2cba7-98f5-4d3b-bc2f-b670f0621207\Yoink Software.xml"
echo Creating destination directory if it doesn't exist: "%DEST_DIR_XML_2%"
md "%DEST_DIR_XML_2%" 2>nul
echo Copying "Yoink Software.xml" to second destination...
copy "%SOURCE_DIR%Yoink Software.xml" "%DEST_DIR_XML_2%\" /Y
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Failed to copy Yoink Software.xml to the second location.
) else (
    echo Yoink Software.xml copied successfully to "%DEST_DIR_XML_2%\".
)
echo.
echo -------------------------------------
echo --- All copy operations complete. ---
echo -------------------------------------
echo.
echo.
echo Add your applications to: C:\Program Files\Yoink Software\Yoink4CM\monthly.ps1
echo C:\Program Files\Yoink Software\Yoink4CM\quarterly.ps1
echo.
echo Your server settings have been saved to C:\Program Files\Yoink Software\config.ps1
echo.
set /p userAnswer="Do you wish to open the documentation in your browser? (y/n): "
if /i "%userAnswer%"=="y" goto OPENDOC
if /i "%userAnswer%"=="Y" goto OPENDOC

goto CONTINUE

:OPENDOC
echo Opening documentation in your browser...
rundll32 url.dll,FileProtocolHandler https://www.yoink4cm.com/yoink4cm-documentation/

:CONTINUE

echo.
echo Please relaunch Microsoft Configuration Manager Console to finish integration.
echo Yoink Software integration can be found in the context menus under Application Management --> Applications, and Application Management --> Packages.
echo.
echo Questions?  Email support@yoink4cm.com

pause