param (
    [Parameter(Mandatory=$false)]
    [string]$FilePath
)

if ($FilePath -and (Test-Path -Path $FilePath)) {
    
    $fileName = Split-Path -Leaf $FilePath

    
    try {
        
		$info = (Get-Item $FilePath).VersionInfo
		$ProductName = $info.ProductName
		$ProductVersion = $info.ProductVersion
		if ([string]::IsNullOrWhiteSpace($ProductName)) {

		$ProductName=Read-Host -Prompt "`nProduct name is missing.  Please enter one"

		}
		
		if ([string]::IsNullOrWhiteSpace($ProductVersion)) {
			$ProductVersion = Read-Host -Prompt "`nProduct version is missing.  Please enter one"
			exit
		}
		Write-Host "Product Name: " $ProductName
		Write-Host "Version: " $ProductVersion

        $Round = [Math]::Round((Get-ChildItem . -Recurse -File | Measure-Object Length -Sum).Sum / 1MB)
        Write-Host "Estimated file size: " $Round 
		
    } catch {
        Write-Host "Warning: Could not extract EXE metadata." -ForegroundColor DarkYellow
    }

    $commandString = $fileName + " "

	Write-Host "`n--- Command Preview ---" -ForegroundColor Green
    Write-Host $commandString
    Write-Host "`nModify the command below if needed (or just press Enter to accept)"
	write-host ""$commandString" is already included.  Just add necessary switches."
    $userInput = Read-Host -Prompt "> "
    
    $finalCommand = $commandString + $userInput

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
	
	$PackageParent = 'Package\Automatic Packages'
	$CMPSSuppressFastNotUsedCheck = $true							

	$NewFolder = $networkshare + $info.ProductName + "\"
	
	if (Test-Path -Path $NewFolder) {
		Write-Host "`n"$ProductName"already exists on your share." -ForegroundColor Yellow
		$NewFolder
		Read-Host -Prompt "`nPress Enter to exit script"
		exit
	}
	
	New-Item -Path $NewFolder -ItemType Directory -Force
	
	$ExcludeList = @("__auto_package.ps1", "__Drop Apps Here.lnk")

	# Copy everything from the current directory to the destination
	# -Recurse ensures subfolders are included
	# -Exclude filters out the specific files mentioned
	Get-ChildItem -Path ".\*" -Exclude $ExcludeList -Recurse | Copy-Item -Destination $NewFolder -Container
	
	$RestoreLocation = get-location
	Set-Location "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin"
	Import-Module .\ConfigurationManager.psd1
	set-location $Sitecode
	

	$PackageName = $ProductName

	$testpackage = Get-CMPackage -Name $PackageName -ErrorAction SilentlyContinue
	if ($testpackage) {
    		Write-Host "Package "$PackageName" already exists, skipping."
	} else {
   		New-CMPackage -Name $Packagename -Language English -version $ProductVersion -Path $NewFolder -Description "Automated by Yoink4CM"
		New-CMProgram -PackageName $Packagename -StandardProgramName $PackageName -DiskSpaceRequirement $Round -DiskSpaceUnit MB -DriveMode RunWithUnc -Duration 120 -ProgramRunType whetherOrNotUserIsLoggedOn -RunMode RunWithAdministrativeRights -RunType Hidden -CommandLine $finalCommand
		$app = Get-CMPackage -Name $Packagename
		Set-CMProgram -InputObject $app -EnableTaskSequence $True -standardprogram
		Move-CMObject -FolderPath $PackageParent -InputObject $app
		Start-CMContentDistribution -PackageName $Packagename -DistributionPointGroupName $DistributionPointGroup

		if ($DeviceCollectionID.Length -eq 8) {

			New-CMPackageDeployment -CollectionID $DeviceCollectionID -StandardProgram -ProgramName $PackageName -PackageName $Packagename -DeployPurpose Available -ScheduleEvent AsSoonAsPossible -FastNetworkOption DownloadContentFromDistributionPointAndRunLocally -SlowNetworkOption DownloadContentFromDistributionPointAndLocally
			
		}
	}

	set-location $RestoreLocation

} else {
    Write-Host "Error: No valid file detected." -ForegroundColor Red
}

Read-Host -Prompt "`nPress Enter to exit"