## Define variables here

#Do Not Change 
$Pattern = "Silent:"											#used extract silent install command from msi or exe yaml
$CMPSSuppressFastNotUsedCheck = $true							#Likely doesn't need to be changed.. suppress messages from Config Manager cmdlets.. cleaner execution screen, doesn't affect outcome
$url = "https://github.com/Microsoft/Microsoft-Win32-Content-Prep-Tool"


## Determine month so we can download monthly, quarterly updates

$Month = (Get-Date -UFormat "%m")





# Get the script's directory for reliable pathing
$scriptDirectory = $PSScriptRoot

# Check and create "Downloads" subfolder
$TempFolder = Join-Path -Path $scriptDirectory -ChildPath "Downloads"
if (-not (Test-Path -Path $TempFolder)) {
    Write-Host "Creating 'Downloads' folder at: $TempFolder"
    New-Item -Path $TempFolder -ItemType Directory | Out-Null
} else {
	Remove-Item -Path $TempFolder\* -Recurse -Force
}

# Check and create "Packaged" subfolder
$packagedFolder = Join-Path -Path $scriptDirectory -ChildPath "Packaged"
if (-not (Test-Path -Path $packagedFolder)) {
    Write-Host "Creating 'Packaged' folder at: $packagedFolder"
    New-Item -Path $packagedFolder -ItemType Directory | Out-Null
} else {
    
}

# Check and create "Prep Tool" subfolder
$ToolFolder = Join-Path -Path $scriptDirectory -ChildPath "Prep Tool"
if (-not (Test-Path -Path $ToolFolder)) {
	Write-Host ""
    Write-Host "Creating 'Prep Tool' folder at: $ToolFolder"
    New-Item -Path $ToolFolder -ItemType Directory | Out-Null
	Write-Host "Please download the Microsoft Win32 Content Prep Tool from: https://github.com/Microsoft/Microsoft-Win32-Content-Prep-Tool"
	Write-Host "Please copy the extracted contents to .\Prep Tool" 
	$response = Read-Host "Open the web page for the Microsoft Win32 Content Prep Tool? (y/n)"

	# Check the user's response (case-insensitive)
	if ($response -match "^y(es)?$") {
		# If the user says "yes" or "y", open the URL
		Write-Host "Opening $url"
		Start-Process $url
	}
} else {
    
	if ((Test-Path -Path $ToolFolder) -and (Get-ChildItem -Path $ToolFolder -ErrorAction SilentlyContinue | Measure-Object).Count -eq 0) {
		# The folder is empty, proceed to ask the user
		Write-Host "You must download the Microsoft Win32 Content Prep Tool from: https://github.com/Microsoft/Microsoft-Win32-Content-Prep-Tool"
		Write-Host "Please copy the extracted contents to .\Prep Tool" 
		$response = Read-Host "Open the web page for the Microsoft Win32 Content Prep Tool? (y/n)"
    
		# Check the user's response (case-insensitive)
		if ($response -match "^y(es)?$") {
			Write-Host "Opening $url"
			Start-Process $url
		}
		exit 
	}
}

## Download Monthly Updates


. .\monthly.ps1


## Download Quarterly Updates

if ($Month -eq '01' -or $Month -eq '04' -or $Month -eq '07' -or $Month -eq '10') {

    . .\quarterly.ps1

}



cd $TempFolder


$files = Get-ChildItem *.yaml -Path '.'

foreach ($file in $files) {
    # Get the base name without the extension
    $Filename2 = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)

    Write-Host "Processing $($file.Name)..."
    $Newfolder = Join-Path -Path $TempFolder -ChildPath $Filename2

    # Create new folder
    New-Item -ItemType Directory -Force -Path $Newfolder

    # Copy the specific .yaml file
    Copy-Item -Path $file.FullName -Destination $Newfolder
	
	
	Select-String -Path $file.FullName -Pattern $Pattern | ForEach {
    		$VarIndex = $_.Line.IndexOf($Pattern)
    		$SilentArgument = ($_.Line.Substring($VarIndex,($_.Line.Length - $VarIndex)) -replace "^$Pattern").Trim()
	}
	$silentArgumentFilePath = Join-Path -Path $packagedFolder -ChildPath "$Filename2.txt"
    Write-Host "  Creating silent argument file: $silentArgumentFilePath"
	$FinalArgument = "$Filename2 $SilentArgument"
    Set-Content -Path $silentArgumentFilePath -Value $FinalArgument

    # Define paths for both .exe and .msi files
    $exePath = Join-Path -Path $file.DirectoryName -ChildPath "$Filename2.exe"
    $msiPath = Join-Path -Path $file.DirectoryName -ChildPath "$Filename2.msi"

    # Check for the .exe file first
    if (Test-Path -Path $exePath) {
        Write-Host "Found corresponding .exe file: $exePath"
        Copy-Item -Path $exePath -Destination $Newfolder
		& "..\Prep Tool\IntuneWinAppUtil.exe" -c "$Newfolder" -s "$exepath" -o "$packagedFolder" -q
    } 
    # If the .exe file is not found, check for the .msi file
    elseif (Test-Path -Path $msiPath) {
        Write-Host "Found corresponding .msi file: $msiPath"
        Copy-Item -Path $msiPath -Destination $Newfolder
		& "..\Prep Tool\IntuneWinAppUtil.exe" -c "$Newfolder" -s "$msipath" -o "$packagedFolder" -q
    } 
    # If neither file is found
    else {
        Write-Host "  No corresponding .exe or .msi file found for $Filename2.yaml, skipping."
    }
}
