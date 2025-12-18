param (
    [Parameter(Mandatory=$false)]
    [string]$FilePath
)

if ($FilePath -and (Test-Path -Path $FilePath)) {
    
    $fileName = Split-Path -Leaf $FilePath

    
    try {
        $windowsInstaller = New-Object -ComObject WindowsInstaller.Installer
        # Open database in Read-Only mode (0)
        $database = $windowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $windowsInstaller, @($FilePath, 0))
        
        # --- MSI SUBJECT EXTRACTION ---
        # Property ID 3 corresponds to 'Subject' in the Summary Information Stream
        #$summaryInfo = $database.GetType().InvokeMember("SummaryInformation", "GetProperty", $null, $database, @(0))
        #$subject = $summaryInfo.GetType().InvokeMember("Property", "GetProperty", $null, $summaryInfo, @(3))
        
        #Write-Host "MSI Subject: "$subject""
        

        # --- MSI INSTALLATION PARAMETERS EXTRACTION ---
		
		Write-Host "--- Product Information ---" -ForegroundColor Green
		$PropertiesToGet = @("ProductName", "ProductVersion", "Manufacturer", "ProductCode")

# 1. Create an empty Hash Table to store the results
		$ProductInfo = @{}

		foreach ($Prop in $PropertiesToGet) {
			$view = $database.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $database, @("SELECT Value FROM Property WHERE Property = '$Prop'"))
			$view.GetType().InvokeMember("Execute", "InvokeMethod", $null, $view, $null)
			$record = $view.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $view, $null)
    
			if ($record) {
				$val = $record.GetType().InvokeMember("StringData", "GetProperty", $null, $record, @(1))
        
        # 2. Save the value into the Hash Table using the Property name as the Key
				$ProductInfo[$Prop] = $val
        
				Write-Host "$($Prop): $val" -ForegroundColor White
			}
		}
		
        Write-Host "`n--- MSI Property Table ---" -ForegroundColor Green
        
        # Query for ALL properties, then filter for "Public" ones (All Caps)
        $view = $database.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $database, @("SELECT Property, Value FROM Property"))
        $view.GetType().InvokeMember("Execute", "InvokeMethod", $null, $view, $null)

        while ($record = $view.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $view, $null)) {
            $propName = $record.GetType().InvokeMember("StringData", "GetProperty", $null, $record, @(1))
            $propVal  = $record.GetType().InvokeMember("StringData", "GetProperty", $null, $record, @(2))
            
            # Regex: Only show properties that are entirely Uppercase (Public Properties)
            if ($propName -cmatch '^[A-Z0-9_]+$') {
                Write-Host "$($propName): " -NoNewline -ForegroundColor White
                Write-Host $propVal -ForegroundColor DarkGray
            }
        }
        
        
    } catch {
        Write-Host "Warning: Could not extract MSI metadata. Ensure the file is a valid MSI." -ForegroundColor DarkYellow
    }

    $commandString = "msiexec /i ""$fileName"" /qn"

	
    Write-Host "`n--- Command Preview ---" -ForegroundColor Green
    Write-Host $commandString
    Write-Host "`nModify the command below if needed (or just press Enter to accept)"
	write-host 'If modifying, ensure to include: msiexec /i "filename"'
    $userInput = Read-Host -Prompt "> "
    
    $finalCommand = if (-not [string]::IsNullOrWhiteSpace($userInput)) { $userInput } else { $commandString }

    Write-Host "`nFinal Command: $finalCommand" -ForegroundColor Green
    Write-Host "-----------------------"

    # Yoink4CM integration check
    $configPath = "C:\Program Files\Yoink Software\config.ps1"
    if (-not (Test-Path -Path $configPath)) {
        Write-Host "ERROR: Yoink4CM must be integrated with Configuration Manager console first." -ForegroundColor Red
        Read-Host "`nPress Enter to exit"
        exit
    }

    # Write-Host "Validation Successful: Yoink4CM integration found." -ForegroundColor Green
	. "C:\Program Files\Yoink Software\config.ps1"
	
	$ApplicationParent = 'Application\Automatic Applications'		#used to create a subfolder in Config Mgr for Applications
	$CMPSSuppressFastNotUsedCheck = $true							#Likely doesn't need to be changed.. suppress messages from Config Manager cmdlets.. cleaner execution screen, doesn't affect outcome

	$NewFolder = $networkshare + $ProductInfo.ProductName + "\" + $ProductInfo.ProductVersion + "\"
	
	if (Test-Path -Path $NewFolder) {
    Write-Host "`nVersion"$ProductInfo.ProductVersion"of"$ProductInfo.ProductName"already exists on your share." -ForegroundColor Yellow
	$NewFolder
	Read-Host -Prompt "`nPress Enter to exit script"
    exit
}
	
	New-Item -Path $NewFolder -ItemType Directory -Force
	
	$ExcludeList = @("__auto_app.ps1", "__Drop Apps Here.lnk")

	# Copy everything from the current directory to the destination
	# -Recurse ensures subfolders are included
	# -Exclude filters out the specific files mentioned
	Get-ChildItem -Path ".\*" -Exclude $ExcludeList -Recurse | Copy-Item -Destination $NewFolder -Container
	
	$RestoreLocation = get-location
	Set-Location "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin"
	Import-Module .\ConfigurationManager.psd1
	set-location $Sitecode
	
	if ($CreateDeviceCollections -eq "y") {
    
		$Schedule = New-CMSchedule -DurationInterval Days -DurationCount 0 -RecurInterval Days -RecurCount 7 	#For Device Collection Creation.. Update the membership weekly
		$TestingCollectionsParentPath = 'DeviceCollection'
		$TestingCollectionsFolderName = 'Testing Collections'
		$TestingCollectionsPath = "$TestingCollectionsParentPath\$TestingCollectionsFolderName"

		# Check and create the primary "Testing Collections" folder
		if (-not (Test-Path -Path $TestingCollectionsPath)) {
			echo "Creating primary Device Collection folder: '$TestingCollectionsFolderName'"
			# New-CMFolder requires the parent path and the name of the new folder
			New-CMFolder -ParentFolderPath $TestingCollectionsParentPath -Name $TestingCollectionsFolderName | Out-Null
		}
    
	}
	$FullName = $ProductInfo.ProductName + " " + $ProductInfo.ProductVersion

	$testapp = Get-CMApplication -Name $FullName -ErrorAction SilentlyContinue
	if ($testapp) {
    	Write-Host "Application '$FullName' already exists in Configuration Manager, skipping.  You may have to browse to a different folder in Applications, then return to Automatic Applications folder to see it."
	} else {
		$ContentPath = $NewFolder + "\" + $fileName
		New-CMApplication -Name $FullName -SoftwareVersion $ProductInfo.ProductVersion -AutoInstall $True -Description "Automated by Yoink4CM"
		Add-CMMSiDeploymentType -ApplicationName $FullName -DeploymentTypeName $FullName -ContentLocation $ContentPath -Comment "Automated by Yoink4CM" -InstallCommand $finalCommand -EstimatedRuntimeMins 60 -InstallationBehaviorType "InstallForSystem" -LogonRequirementType "WhetherOrNotUserLoggedOn" -MaximumRuntimeMins 120 -force
		$app = Get-CMApplication -Name $FullName
		Move-CMObject -FolderPath $ApplicationParent -InputObject $app
		Start-CMContentDistribution -ApplicationName $FullName -DistributionPointGroupName $DistributionPointGroup
		if ($DeviceCollectionID.Length -eq 8) {

			New-CMApplicationDeployment -ApplicationName $FullName -CollectionID $DeviceCollectionID -DeployAction Install -DeadlineDateTime (get-date) -DeployPurpose Available -UserNotification DisplaySoftwareCenterOnly
			
		}

		if ($CreateDeviceCollections -eq "y") {
			
			
			$DCName = $ProductInfo.ProductName + " < " + $ProductInfo.ProductVersion
			
			New-CMDeviceCollection -Name $DCName -LimitingCollectionName "All Systems" -RefreshType Periodic -RefreshSchedule $Schedule
			$Query = @"
select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System 
inner join SMS_G_System_ADD_REMOVE_PROGRAMS on SMS_G_System_ADD_REMOVE_PROGRAMS.ResourceID = SMS_R_System.ResourceId 
inner join SMS_G_System_ADD_REMOVE_PROGRAMS_64 on SMS_G_System_ADD_REMOVE_PROGRAMS_64.ResourceID = SMS_R_System.ResourceId 
where 
    (SMS_G_System_ADD_REMOVE_PROGRAMS.DisplayName like '{0}%' and SMS_G_System_ADD_REMOVE_PROGRAMS.Version < '{1}') 
    or 
    (SMS_G_System_ADD_REMOVE_PROGRAMS_64.DisplayName like '{0}%' and SMS_G_System_ADD_REMOVE_PROGRAMS_64.Version < '{1}')
"@ -f $ProductInfo.ProductName, $ProductInfo.ProductVersion
			$RuleName = "Devices with outdated $($ProductInfo.ProductName) (below $($ProductInfo.ProductVersion))"
			Add-CMDeviceCollectionQueryMembershipRule -CollectionName $DCName -RuleName $RuleName -QueryExpression $Query
			Invoke-CMCollectionUpdate -Name $DCname
			$CollectionToMove = Get-CMDeviceCollection -Name $DCName
			Move-CMObject -InputObject $CollectionToMove -FolderPath $TestingCollectionsPath
		}
	}

	set-location $RestoreLocation

} else {
    Write-Host "Error: No valid file detected." -ForegroundColor Red
}

Read-Host -Prompt "`nPress Enter to exit"