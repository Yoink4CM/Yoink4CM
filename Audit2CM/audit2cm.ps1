param(
    [Parameter(Mandatory=$true)]
    [string]$TextFilePath,

    [Parameter(Mandatory=$true)]
    [string]$CollectionName,
	[string]$SiteCode
)




#Connect to Configuration Manager

Write-Host "Attempting to connect to Configuration Manager site: $SiteCode..."

try {
    # Determine the path to the Configuration Manager PowerShell module
    $RestoreLocation = get-location
    Set-Location "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin"
    Import-Module .\ConfigurationManager.psd1

    # Set the current location to the ConfigMgr site drive
    Set-Location "$($SiteCode):" -ErrorAction Stop

    Write-Host "Successfully connected to Configuration Manager site: $SiteCode."
}
catch {
    Write-Error "Failed to connect to Configuration Manager. Please ensure the console is installed and you have sufficient permissions. Error: $($_.Exception.Message)"
    exit 1
}



#Validate Collection

Write-Host "Validating collection: $CollectionName..."

try {
    # Attempt to retrieve the existing collection
    $collection = Get-CMDeviceCollection -Name $CollectionName -ErrorAction Stop
    Write-Host "Collection '$CollectionName' found. Collection ID: $($collection.CollectionID)"
}
catch {
    # If the collection is not found, create it
    Write-Warning "Collection '$CollectionName' not found. Creating a new device collection..."
    
    # Set the limiting collection name
    $LimitingCollectionName = "All Systems" 

    try {
        # Get the Limiting Collection object (optional step, but good for error checking)
        $limitingCollection = Get-CMDeviceCollection -Name $LimitingCollectionName -ErrorAction Stop
        
        # Create the new device collection
        $newCollection = New-CMDeviceCollection -Name $CollectionName -LimitingCollectionName $LimitingCollectionName

        Write-Host "Successfully created new device collection '$CollectionName'. Collection ID: $($newCollection.CollectionID)"
        
    }
    catch {
        # Handle errors during the creation process
        Write-Error "Failed to create device collection '$CollectionName' using '$LimitingCollectionName'. Error: $($_.Exception.Message)"
        exit 1
    }
}



#Read Hostnames and Add to Collection

Write-Host "Reading hostnames from: $TextFilePath..."

Set-Location "$RestoreLocation"

if (-not (Test-Path $TextFilePath)) {
    Write-Error "The specified text file '$TextFilePath' does not exist."
    exit 1
}

$hostnames = Get-Content $TextFilePath

if ($hostnames.Count -eq 0) {
    Write-Warning "The text file '$TextFilePath' is empty. No devices to add."
    exit 0
}

Write-Host "Found $($hostnames.Count) hostnames in the file. Starting import..."

$addedCount = 0
$skippedCount = 0
$errorLog = @()

Set-Location "$($SiteCode):"

foreach ($hostname in $hostnames) {
    $hostname = $hostname.Trim() # Remove leading/trailing whitespace
    if ([string]::IsNullOrWhiteSpace($hostname)) {
        continue # Skip empty lines
    }

    Write-Host "Processing hostname: $hostname" -ForegroundColor DarkCyan

    try {
        $device = Get-CMDevice -Name $hostname -ErrorAction SilentlyContinue
        
        if ($null -ne $device) {
            # Check if the device is already a member to avoid errors and redundant operations
            if (-not (Get-CMDeviceCollectionDirectMembershipRule -CollectionName $CollectionName -ResourceName $hostname -ErrorAction SilentlyContinue)) {
                Add-CMDeviceCollectionDirectMembershipRule -CollectionName $CollectionName -ResourceId $device.ResourceID -ErrorAction Stop
                Write-Host "Successfully added '$hostname' to collection '$CollectionName'." -ForegroundColor Green
                $addedCount++
            } else {
                Write-Host "Device '$hostname' is already a member of collection '$CollectionName'. Skipping." -ForegroundColor Yellow
                $skippedCount++
            }
        } else {
            Write-Warning "Device '$hostname' not found in Configuration Manager. Skipping."
            $skippedCount++
            $errorLog += "Skipped: Device '$hostname' not found in ConfigMgr."
        }
    }
    catch {
        Write-Error "Failed to add '$hostname' to collection '$CollectionName'. Error: $($_.Exception.Message)"
        $skippedCount++
        $errorLog += "Failed: '$hostname' - $($_.Exception.Message)"
    }
}



#Summary

Write-Host "`n--- Import Summary ---"
Write-Host "Total hostnames processed: $($hostnames.Count)"
Write-Host "Devices successfully added: $addedCount" -ForegroundColor Green
Write-Host "Devices skipped/failed: $skippedCount" -ForegroundColor Yellow

if ($errorLog.Count -gt 0) {
    Write-Host "`nDetails of skipped/failed devices:" -ForegroundColor Red
    $errorLog | ForEach-Object { Write-Host $_ }
}

Write-Host "Script execution complete."
Set-Location "$RestoreLocation"
