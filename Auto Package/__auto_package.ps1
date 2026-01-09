param (
    [Parameter(Mandatory=$false)]
    [string]$FilePath
)

if ($FilePath -and (Test-Path -Path $FilePath)) {
    
    $fileName = Split-Path -Leaf $FilePath
    $directory = Split-Path -Parent $FilePath
    
    try {
        $info = (Get-Item $FilePath).VersionInfo
        $ProductName = $info.ProductName
        $ProductVersion = $info.ProductVersion
        
        if ([string]::IsNullOrWhiteSpace($ProductName)) {
            $ProductName = Read-Host -Prompt "`nProduct name is missing. Please enter one"
        }
        
        if ([string]::IsNullOrWhiteSpace($ProductVersion)) {
            $ProductVersion = Read-Host -Prompt "`nProduct version is missing. Please enter one"
            exit
        }
        
        Write-Host "Product Name: " $ProductName
        Write-Host "Version: " $ProductVersion

        $Round = [Math]::Round((Get-ChildItem . -Recurse -File | Measure-Object Length -Sum).Sum / 1MB)
        Write-Host "Estimated file size in MB: " $Round 
        
    } catch {
        Write-Host "Warning: Could not extract EXE metadata." -ForegroundColor DarkYellow
    }


    $Yaml = $fileName -replace "\.exe$", ".yaml"
    $yamlPath = Join-Path -Path $directory -ChildPath $Yaml
    
    $foundSilentParams = $false
    $finalCommand = ""

    if (Test-Path -Path $yamlPath) {
        Write-Host "Found config file: $Yaml" -ForegroundColor Cyan

		$content = Get-Content -Path $yamlPath -Raw
		$lines = $content -split '\r?\n'
    

		$silentLine = $lines | Where-Object { $_ -match "^\s*Silent:" } | Select-Object -First 1
        
        if ($silentLine) {
  
			$installParams = ($silentLine -replace "^\s*Silent:\s*", "").Trim()
    
			$finalCommand = "$fileName $installParams"
			$foundSilentParams = $true
			Write-Host "Extracted silent parameters: $installParams" -ForegroundColor Green
		}
    }

    # If YAML doesn't exist or doesn't contain the "Silent:" line, fall back to manual prompt
    if (-not $foundSilentParams) {
        $commandString = $fileName + " "
        Write-Host "`n--- Command Preview ---" -ForegroundColor Green
        Write-Host $commandString
        Write-Host "`nModify the command below if needed (or just press Enter to accept)"
        Write-Host "'$commandString' is already included. Just add necessary switches."
        $userInput = Read-Host -Prompt "> "
        $finalCommand = $commandString + $userInput
    }


    Write-Host "`nFinal Command: $finalCommand" -ForegroundColor Green
    Write-Host "-----------------------"

    # Yoink4CM integration check
    $configPath = "C:\Program Files\Yoink Software\config.ps1"
    if (-not (Test-Path -Path $configPath)) {
        Write-Host "ERROR: Yoink4CM must be integrated with Configuration Manager console first." -ForegroundColor Red
        Read-Host "`nPress Enter to exit"
        exit
    }

    . "C:\Program Files\Yoink Software\config.ps1"
    
    $PackageParent = 'Package\Automatic Packages'
    $CMPSSuppressFastNotUsedCheck = $true                            

    $NewFolder = $networkshare + $ProductName + "\"
    
    if (Test-Path -Path $NewFolder) {
        Write-Host "`n$ProductName already exists on your share." -ForegroundColor Yellow
        Read-Host -Prompt "`nPress Enter to exit script"
        exit
    }
    
    New-Item -Path $NewFolder -ItemType Directory -Force
    
    $ExcludeList = @("__auto_package.ps1", "__Drop Apps Here.lnk")
    Get-ChildItem -Path ".\*" -Exclude $ExcludeList -Recurse | Copy-Item -Destination $NewFolder -Container
    
    $RestoreLocation = Get-Location
    Set-Location "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin"
    Import-Module .\ConfigurationManager.psd1
    Set-Location $Sitecode
    
    $PackageName = $ProductName

    $testpackage = Get-CMPackage -Name $PackageName -ErrorAction SilentlyContinue
    if ($testpackage) {
            Write-Host "Package $PackageName already exists, skipping."
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

    Set-Location $RestoreLocation

} else {
    Write-Host "Error: No valid file detected." -ForegroundColor Red
}

Read-Host -Prompt "`nPress Enter to exit"