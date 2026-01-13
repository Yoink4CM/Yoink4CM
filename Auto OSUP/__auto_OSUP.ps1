# No params needed as we are scanning the current folder for the ISO directory
$currentDirectory = Get-Location


$isoFolder = Get-ChildItem -Path $currentDirectory -Directory | Where-Object { 
    Test-Path (Join-Path $_.FullName "setup.exe") 
} | Select-Object -First 1

if (-not $isoFolder) {
    Write-Host "Error: Could not find a folder containing 'setup.exe' in the current directory." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit
}

$FilePath = Join-Path $isoFolder.FullName "sources\install.wim"

if (Test-Path -Path $FilePath) {
    
    $tempDataPath = "$env:TEMP\wim_metadata.json"
    if (Test-Path $tempDataPath) { Remove-Item $tempDataPath }


    Write-Host "Requesting Admin permission to inspect WIM metadata..." -ForegroundColor Cyan
    
    $adminScript = {
        param($path, $outputPath)
        try {
            Import-Module Dism
 
            $image = Get-WindowsImage -ImagePath $path -Index 1 -ErrorAction SilentlyContinue
            
            $data = @{
                Name    = $image.ImageName
                Version = $image.Version
                Success = $true
            }
        } catch {
            $data = @{ Success = $false; Error = $_.Exception.Message }
        }
        $data | ConvertTo-Json | Set-Content -Path $outputPath
    }

    Start-Process powershell -Verb RunAs -Wait -ArgumentList "-Command & {$adminScript}", "'$FilePath'", "'$tempDataPath'"


    if (Test-Path $tempDataPath) {
        $wimMetadata = Get-Content $tempDataPath | ConvertFrom-Json
        Remove-Item $tempDataPath 
        
        if ($wimMetadata.Success) {
            $wimName = $wimMetadata.Name
            $wimVersion = $wimMetadata.Version
            Write-Host "Success! Metadata retrieved: $wimName ($wimVersion)" -ForegroundColor Green
        } else {
            Write-Host "Admin script failed to read WIM: $($wimMetadata.Error)" -ForegroundColor Red
            exit
        }
    } else {
        Write-Host "Admin elevation was cancelled or failed." -ForegroundColor Red
        exit
    }

    # Yoink4CM integration check
    $configPath = "C:\Program Files\Yoink Software\config.ps1"
    if (-not (Test-Path -Path $configPath)) {
        Write-Host "ERROR: Yoink4CM must be integrated first." -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit
    }

    . "C:\Program Files\Yoink Software\config.ps1"
    
    # Define destination on share 
    $cleanName = $wimName.Replace(" ","_")
    $NewFolder = Join-Path $networkshare ("OSUP\" + $cleanName + "\" + $wimVersion)
	
	$FullName = "$wimName ($wimVersion) Upgrade"
    
    if (-not (Test-Path -Path $NewFolder)) {
        Write-Host "Creating directory: $NewFolder" -ForegroundColor Gray
        New-Item -Path $NewFolder -ItemType Directory -Force | Out-Null
    }
    
    Write-Host "Copying $FullName Upgrade Package to network share..." -ForegroundColor Gray
    # Copy the ISO folder contents
    Copy-Item -Path "$($isoFolder.FullName)\*" -Destination $NewFolder -Recurse -Force


    $RestoreLocation = Get-Location
    Set-Location "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin"
    Import-Module .\ConfigurationManager.psd1
    Set-Location "$Sitecode"
    


    
    Write-Host "Creating OS Upgrade Package in Configuration Manager..." -ForegroundColor Green

    $objPackage = New-CMOperatingSystemUpgradePackage -Name $FullName -Path $NewFolder -Version $wimVersion -Description "Automated OSUP upload via Yoink4CM"
    
    # Move to Folder
    $OSUPParent = 'OperatingSystemInstaller\Automatic Uploads'
    if (-not (Test-Path -Path "$OSUPParent")) {
        New-CMFolder -ParentFolderPath "OperatingSystemInstaller" -Name "Automatic Uploads" | Out-Null
    }
    Move-CMObject -FolderPath $OSUPParent -InputObject $objPackage
	
    if ($DistributionPointGroup) {
        Write-Host "Distributing to $DistributionPointGroup..." -ForegroundColor Gray
        Start-CMContentDistribution  -OperatingSystemInstallerName $FullName -DistributionPointGroupName $DistributionPointGroup
    }

    Set-Location $RestoreLocation
    Write-Host "`nDone!" -ForegroundColor Green

} else {
    Write-Host "Error: install.wim not found in $($isoFolder.FullName)\sources" -ForegroundColor Red
}

Read-Host -Prompt "`nPress Enter to exit"