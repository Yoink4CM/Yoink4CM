## Define variables here

. ..\config.ps1

#Do Not Change 
$Pattern = "Silent:"											#used extract silent install command from msi or exe yaml
$MSIXPattern = "PackageFamilyName:"								#used extract silent install command from msix yaml
$PackageParent = 'Package\Automatic Packages'					#used to create a subfolder in Config Mgr for Packages
$ApplicationParent = 'Application\Automatic Applications'			#used to create a subfolder in Config Mgr for Applications
$CMPSSuppressFastNotUsedCheck = $true							#Likely doesn't need to be changed.. suppress messages from Config Manager cmdlets.. cleaner execution screen, doesn't affect outcome


#Tweakable as needed




# Get the current user's temporary directory
# This ensures the temporary files are stored in the user's specific temp folder,
# which is typically C:\Users\[current user]\AppData\Local\Temp
$Tempfolder1 = [System.IO.Path]::GetTempPath()

# Navigate to the temporary directory
Set-Location $Tempfolder1

# Create a subfolder named 'Yoink4CM' within the user's temporary directory
# and then update $Tempfolder to point to this new subfolder.
$Tempfolder = Join-Path -Path $Tempfolder1 -ChildPath "Yoink4CM"
New-Item -Path $Tempfolder -ItemType Directory -Force | Out-Null # Out-Null to suppress output

cd Yoink4CM



## Determine month so we can download monthly, quarterly updates

$Month = (Get-Date -UFormat "%m") 


## Download Monthly Updates


. "C:\Program Files\Yoink Software\Yoink4CM\monthly.ps1"


## Download Quarterly Updates

if ($Month -eq '01' -or $Month -eq '04' -or $Month -eq '07' -or $Month -eq '10') {

	. "C:\Program Files\Yoink Software\Yoink4CM\quarterly.ps1"

}



## Create monthly patch folder on network share
$folderdate = (Get-date -f "yyyy_MM") 

$Monthlyfolder = $networkshare + $folderdate + "\"


New-Item -Path $Monthlyfolder -ItemType Directory -Force


#Set location to config manager and create new package folder for this month
$RestoreLocation = get-location
Set-Location "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin"
Import-Module .\ConfigurationManager.psd1

set-location $Sitecode
$PackagePath = $PackageParent + "\" + $folderdate
$ApplicationPath = $ApplicationParent + "\" + $folderdate






if ($CreateDeviceCollections -eq "y") {
    
	$Schedule = New-CMSchedule -DurationInterval Days -DurationCount 0 -RecurInterval Days -RecurCount 7 	#For Device Collection Creation.. Update the membership weekly
    $TestingCollectionsParentPath = 'DeviceCollection'
    $TestingCollectionsFolderName = 'Testing Collections'
    $TestingCollectionsPath = "$TestingCollectionsParentPath\$TestingCollectionsFolderName"

    # 1. Check and create the primary "Testing Collections" folder
    if (-not (Test-Path -Path $TestingCollectionsPath)) {
        echo "Creating primary Device Collection folder: '$TestingCollectionsFolderName'"
        # New-CMFolder requires the parent path and the name of the new folder
        New-CMFolder -ParentFolderPath $TestingCollectionsParentPath -Name $TestingCollectionsFolderName | Out-Null
    }

    # Check and create the monthly sub folder using $folderdate
    $MonthlyCollectionPath = "$TestingCollectionsPath\$folderdate"
    if (-not (Test-Path -Path $MonthlyCollectionPath)) {
        echo "Creating monthly Device Collection sub folder: '$folderdate'"
        # The parent path for the subfolder is the "Testing Collections" folder
        New-CMFolder -ParentFolderPath $TestingCollectionsPath -Name $folderdate | Out-Null
    }
}



if (Test-Path -Path $PackageParent) {

} else {
	echo 'Need to make Automatic Updates package folder!'
	 New-CMFolder -ParentFolderPath 'Package' -Name 'Automatic Packages'
}

if (Test-Path -Path $PackagePath) {

} else {
	echo 'Need to make monthly package folder!'
	 New-CMFolder -ParentFolderPath 'Package\Automatic Packages' -Name $folderdate
}

if (Test-Path -Path $ApplicationParent) {

} else {
	echo 'Need to make Automatic Updates application folder!'
	 New-CMFolder -ParentFolderPath 'Application' -Name 'Automatic Applications'
}


if (Test-Path -Path $ApplicationPath) {

} else {
	echo 'Need to make monthly application folder!'
	 New-CMFolder -ParentFolderPath 'Application\Automatic Applications' -Name $folderdate
}

Set-Location $RestoreLocation




## Copy files to share

$files = Get-ChildItem *.yaml -Path '.' 				# There's 1 yaml per package.. get its filename

foreach ($file in $files) {
	$Filename = $file.Name
	$Filename2 = $Filename.Substring(0, $Filename.Length - 5)  	# Remove .yaml extension
	$Newfolder = $Monthlyfolder + $Filename2

	New-Item -ItemType Directory -Force -Path $Newfolder		# Create package folder on share
	copy-item .\$Filename2* -Destination $Newfolder		# Copy files to share
	$Yaml = $Newfolder + '\' + $Filename				# Build new string containing folder location + yaml filename
	Select-String -Path $Yaml -Pattern $Pattern | ForEach {
    		$VarIndex = $_.Line.IndexOf($Pattern)
    		$SilentArgument = ($_.Line.Substring($VarIndex,($_.Line.Length - $VarIndex)) -replace "^$Pattern").Trim()
	}
	Select-String -Path $Yaml -Pattern $MSIXPattern | ForEach {
    		$VarIndex = $_.Line.IndexOf($MSIXPattern)
    		$MSIXPFN = ($_.Line.Substring($VarIndex,($_.Line.Length - $VarIndex)) -replace "^$MSIXPattern").Trim()
	}
	$EXE = $Newfolder + '\' + $Filename2 + '.exe' 
	$MSI = $Newfolder + '\' + $Filename2 + '.msi'
	$MSIX = $Newfolder + '\' + $Filename2 + '.msix' 
	$Versionpattern = '(?<=_).+?(?=_)'

	## Test if file is .exe

	if (Test-Path -Path $EXE) {
   		$SilentInstaller = '"' + $Filename2 + '.exe" ' + $SilentArgument
		$Exefiles = Get-ChildItem *.exe -Path $Newfolder
		foreach ($Exefile in $Exefiles) {
			$Packagename = $Exefile.Name 
			$Filesizecalc = $Exefile.Length
		}
	$filesize = ($Filesizecalc/1MB)
	$Rounded = [Math]::Round($filesize)
	$Length = $Packagename.Length

	if($Length -lt '49'){
   		$Programname = $Packagename
	}else {
   		$Programname = $Packagename.SubString(0,48)
	}
	
	
	$version = [regex]::Matches($Packagename, $Versionpattern).Value | Select -First 1
	$sharepath = $networkshare + $folderdate + '\'+ $Filename2 + '\'
	set-location $Sitecode
	Start-Sleep -Seconds 2
	

	$testpackage = Get-CMPackage -Name $PackageName -ErrorAction SilentlyContinue
		if ($testpackage) {
    			Write-Host "Package '$PackageName' already exists, skipping."
		} else {
   	 		New-CMPackage -Name $Packagename -Language English -version $version -Path $sharepath -Description "Automated by Yoink4CM"
			New-CMProgram -PackageName $Packagename -StandardProgramName $Programname -DiskSpaceRequirement $Rounded -DiskSpaceUnit MB -DriveMode RunWithUnc -Duration 120 -ProgramRunType WhetherOrNotUserIsLoggedOn -RunMode RunWithAdministrativeRights -RunType Hidden -CommandLine $SilentInstaller
			$app = Get-CMPackage -Name $Packagename
			Set-CMProgram -InputObject $app -EnableTaskSequence $True -standardprogram
			Move-CMObject -FolderPath $PackagePath -InputObject $app
			Start-CMContentDistribution -PackageName $Packagename -DistributionPointGroupName $DistributionPointGroup


			if ($DeviceCollectionID.Length -eq 8) {

				New-CMPackageDeployment -CollectionID $DeviceCollectionID -StandardProgram -ProgramName $Programname -PackageName $Packagename -DeployPurpose Available -ScheduleEvent AsSoonAsPossible -FastNetworkOption DownloadContentFromDistributionPointAndRunLocally -SlowNetworkOption DownloadContentFromDistributionPointAndLocally
			
			}
			
			if ($CreateDeviceCollections -eq "y") {
				
				$truncPackagename = $Packagename -replace '\s*[\(_].*'
				
				$DCName = "$truncPackagename < $version"
				New-CMDeviceCollection -Name $DCName -LimitingCollectionName "All Systems" -RefreshType Periodic -RefreshSchedule $Schedule
				$Query = @"
select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System 
inner join SMS_G_System_ADD_REMOVE_PROGRAMS on SMS_G_System_ADD_REMOVE_PROGRAMS.ResourceID = SMS_R_System.ResourceId 
inner join SMS_G_System_ADD_REMOVE_PROGRAMS_64 on SMS_G_System_ADD_REMOVE_PROGRAMS_64.ResourceID = SMS_R_System.ResourceId 
where 
    (SMS_G_System_ADD_REMOVE_PROGRAMS.DisplayName like '{0}%' and SMS_G_System_ADD_REMOVE_PROGRAMS.Version < '{1}') 
    or 
    (SMS_G_System_ADD_REMOVE_PROGRAMS_64.DisplayName like '{0}%' and SMS_G_System_ADD_REMOVE_PROGRAMS_64.Version < '{1}')
"@ -f $truncPackagename, $version
				$RuleName = "Devices with outdated $($truncPackagename) (below $($version))"
				Add-CMDeviceCollectionQueryMembershipRule -CollectionName $DCName -RuleName $RuleName -QueryExpression $Query
				Invoke-CMCollectionUpdate -Name $DCname
				$CollectionToMove = Get-CMDeviceCollection -Name $DCName
				Move-CMObject -InputObject $CollectionToMove -FolderPath $MonthlyCollectionPath
				
			}
		}


	
	
	
	}

	## Test if file is .msi

	if (Test-Path -Path $MSI) {
   		$SilentInstaller = 'msiexec /i "' + $Filename2 + '.msi" ' + $SilentArgument
		$Msifiles = Get-ChildItem *.msi -Path $Newfolder
		foreach ($Msifile in $Msifiles) {
			$Packagename = $Msifile.Name
			$Filesizecalc = $Msifile.Length
		}
	$filesize = ($Filesizecalc/1MB)
	$Rounded = [Math]::Round($filesize)
	$version = [regex]::Matches($Packagename, $Versionpattern).Value | Select -First 1
	$sharepath = $networkshare + $folderdate + '\'+ $Filename2 + '\' + $Packagename

	set-location $Sitecode
	Start-Sleep -Seconds 3

	$testapp = Get-CMApplication -Name $PackageName -ErrorAction SilentlyContinue
		if ($testapp) {
    			Write-Host "Application '$PackageName' already exists, skipping."
		} else {

			New-CMApplication -Name $Packagename -SoftwareVersion $version -AutoInstall $True -Description "Automated by Yoink4CM"
			Add-CMMSiDeploymentType -ApplicationName $Packagename -DeploymentTypeName $Packagename -ContentLocation $sharepath -Comment "Automated by Yoink4CM" -InstallCommand $SilentInstaller  -EstimatedRuntimeMins 60 -InstallationBehaviorType "InstallForSystem" -LogonRequirementType "WhetherOrNotUserLoggedOn" -MaximumRuntimeMins 120 -force
			$app = Get-CMApplication -Name $Packagename
			Move-CMObject -FolderPath $ApplicationPath -InputObject $app
			Start-CMContentDistribution -ApplicationName $Packagename -DistributionPointGroupName $DistributionPointGroup
			if ($DeviceCollectionID.Length -eq 8) {

			New-CMApplicationDeployment -ApplicationName $Packagename -CollectionID $DeviceCollectionID -DeployAction Install -DeadlineDateTime (get-date) -DeployPurpose Available -UserNotification DisplaySoftwareCenterOnly
			
			}
			
			if ($CreateDeviceCollections -eq "y") {
				
				$truncPackagename = $Packagename -replace '\s*[\(_].*'
				
				$DCName = "$truncPackagename < $version"
				New-CMDeviceCollection -Name $DCName -LimitingCollectionName "All Systems" -RefreshType Periodic -RefreshSchedule $Schedule
				$Query = @"
select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System 
inner join SMS_G_System_ADD_REMOVE_PROGRAMS on SMS_G_System_ADD_REMOVE_PROGRAMS.ResourceID = SMS_R_System.ResourceId 
inner join SMS_G_System_ADD_REMOVE_PROGRAMS_64 on SMS_G_System_ADD_REMOVE_PROGRAMS_64.ResourceID = SMS_R_System.ResourceId 
where 
    (SMS_G_System_ADD_REMOVE_PROGRAMS.DisplayName like '{0}%' and SMS_G_System_ADD_REMOVE_PROGRAMS.Version < '{1}') 
    or 
    (SMS_G_System_ADD_REMOVE_PROGRAMS_64.DisplayName like '{0}%' and SMS_G_System_ADD_REMOVE_PROGRAMS_64.Version < '{1}')
"@ -f $truncPackagename, $version
				$RuleName = "Devices with outdated $($truncPackagename) (below $($version))"
				Add-CMDeviceCollectionQueryMembershipRule -CollectionName $DCName -RuleName $RuleName -QueryExpression $Query
				Invoke-CMCollectionUpdate -Name $DCname
				$CollectionToMove = Get-CMDeviceCollection -Name $DCName
				Move-CMObject -InputObject $CollectionToMove -FolderPath $MonthlyCollectionPath
				
			}
		}
	}

	## Test if file is .msix

	if (Test-Path -Path $MSIX) {
   		$SilentInstaller = 'msiexec /i "' + $Filename2 + '.msi" ' + $SilentArgument
		$Msixfiles = Get-ChildItem *.msix -Path $Newfolder
		foreach ($Msixfile in $Msixfiles) {
			$Packagename = $Msixfile.Name
			$Filesizecalc = $Msixfile.Length
		}
	$filesize = ($Filesizecalc/1MB)
	$Rounded = [Math]::Round($filesize)
	$version = [regex]::Matches($Packagename, $Versionpattern).Value | Select -First 1
	$sharepath = $networkshare + $folderdate + '\'+ $Filename2 + '\' + $Packagename

	set-location $Sitecode
	Start-Sleep -Seconds 1

	$testapp = Get-CMApplication -Name $PackageName -ErrorAction SilentlyContinue
		if ($testapp) {
    			Write-Host "Application '$PackageName' already exists, skipping."
		} else {

			New-CMApplication -Name $Packagename -SoftwareVersion $version -AutoInstall $True -Description "Automated by Yoink4CM"
			Add-CMWindowsAppxDeploymentType -ApplicationName $Packagename -DeploymentTypeName $Packagename -ContentLocation $sharepath -Comment "Automated by Yoink4CM" -force
			$app = Get-CMApplication -Name $Packagename
			Move-CMObject -FolderPath $ApplicationPath -InputObject $app
			Start-CMContentDistribution -ApplicationName $Packagename -DistributionPointGroupName $DistributionPointGroup
			
			if ($DeviceCollectionID.Length -eq 8) {

				New-CMApplicationDeployment -ApplicationName $Packagename -CollectionID $DeviceCollectionID -DeployAction Install -DeadlineDateTime (get-date) -DeployPurpose Available -UserNotification DisplaySoftwareCenterOnly
			
			}
			
			if ($CreateDeviceCollections -eq "y") {
				
				$truncPackagename = $Packagename -replace '\s*[\(_].*'
				
				$DCName = "$TruckPackagename < $version"
				New-CMDeviceCollection -Name $DCName -LimitingCollectionName "All Systems" -RefreshType Periodic -RefreshSchedule $Schedule
				$Query = @"
select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System 
inner join SMS_G_System_ADD_REMOVE_PROGRAMS on SMS_G_System_ADD_REMOVE_PROGRAMS.ResourceID = SMS_R_System.ResourceId 
inner join SMS_G_System_ADD_REMOVE_PROGRAMS_64 on SMS_G_System_ADD_REMOVE_PROGRAMS_64.ResourceID = SMS_R_System.ResourceId 
where 
    (SMS_G_System_ADD_REMOVE_PROGRAMS.DisplayName like '{0}%' and SMS_G_System_ADD_REMOVE_PROGRAMS.Version < '{1}') 
    or 
    (SMS_G_System_ADD_REMOVE_PROGRAMS_64.DisplayName like '{0}%' and SMS_G_System_ADD_REMOVE_PROGRAMS_64.Version < '{1}')
"@ -f $truncPackagename, $version
				$RuleName = "Devices with outdated $($truncPackagename) (below $($version))"
				Add-CMDeviceCollectionQueryMembershipRule -CollectionName $DCName -RuleName $RuleName -QueryExpression $Query
				Invoke-CMCollectionUpdate -Name $DCname
				$CollectionToMove = Get-CMDeviceCollection -Name $DCName
				Move-CMObject -InputObject $CollectionToMove -FolderPath $MonthlyCollectionPath
				
			}
		}
	}

	""
Set-Location $RestoreLocation
}




Remove-Item ".\*.*"


Stop-Process -Name "Microsoft.ConfigurationManagement"
Start-Process -FilePath "Microsoft.ConfigurationManagement.exe" -WorkingDirectory "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin"