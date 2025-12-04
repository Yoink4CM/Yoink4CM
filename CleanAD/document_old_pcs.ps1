
#########################################################################################################################################

# Variables

. .\config.ps1



# Likely don't need to touch these variables

$ErrorActionPreference = "Continue"  # Tell script to continue even if errors occur (eg. permission denied to move a PC)
$time = (Get-Date).Adddays(-($DaysInactive))  #Calculate todays date - number of days inactive

#########################################################################################################################################






#########################################################################################################################################
  
# STEP 1 - log all PC names to Excel files


Get-ADComputer -Filter {LastLogonTimeStamp -lt $time} -SearchBase $SearchOU -Properties LastLogonTimeStamp | select-object Name,@{Name="Stamp"; Expression={[DateTime]::FromFileTime($_.lastLogonTimestamp)}} | export-csv ".\Documented Systems\Disabled_Systems.csv" -notypeinformation


#########################################################################################################################################










#########################################################################################################################################

# STEP 2 - move all identified PC's to Disabled OU

# Get-ADComputer -Filter {LastLogonTimeStamp -lt $time} -SearchBase $SearchOU -Properties LastLogonTimeStamp | Move-ADObject -TargetPath  $FilteredOU  -ErrorAction SilentlyContinue

#########################################################################################################################################


