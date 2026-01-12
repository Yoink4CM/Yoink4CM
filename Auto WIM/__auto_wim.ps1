param (
    [Parameter(Mandatory=$false)]
    [string]$FilePath
)

if ($FilePath -and (Test-Path -Path $FilePath)) {
    
    $fileName = Split-Path -Leaf $FilePath
    $tempDataPath = "$env:TEMP\wim_metadata.json"
    if (Test-Path $tempDataPath) { Remove-Item $tempDataPath }

    # --- PHASE 1: ELEVATED DATA EXTRACTION ---
    Write-Host "Requesting Admin permission to inspect WIM metadata..." -ForegroundColor Cyan
    
    $adminScript = {
        param($path, $outputPath)
        try {
            Import-Module Dism
            $image = Get-WindowsImage -ImagePath $path -Index 1
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

    # Launch secondary process as Admin
    Start-Process powershell -Verb RunAs -Wait -ArgumentList "-Command & {$adminScript}", "'$FilePath'", "'$tempDataPath'"

    # --- PHASE 2: RETURN TO USER CONTEXT ---
    if (Test-Path $tempDataPath) {
        $wimMetadata = Get-Content $tempDataPath | ConvertFrom-Json
        Remove-Item $tempDataPath # Cleanup
        
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
        Write-Host "ERROR: Yoink4CM must be integrated with Configuration Manager console first." -ForegroundColor Red
        Read-Host "`nPress Enter to exit"
        exit
    }

    . "C:\Program Files\Yoink Software\config.ps1"
    
    # Define destination on share
    $NewFolder = Join-Path $networkshare ("OS_Images\" + $wimName.Replace(" ","_") + "\" + $wimVersion)
    if (-not (Test-Path -Path $NewFolder)) {
        New-Item -Path $NewFolder -ItemType Directory -Force
    }
    
    Write-Host "Copying WIM to network share..." -ForegroundColor Gray
    Copy-Item -Path $FilePath -Destination $NewFolder -Force
    $ContentPath = Join-Path $NewFolder $fileName

    # --- PHASE 3: CONFIG MGR UPLOAD (User Context) ---
    $RestoreLocation = Get-Location
    Set-Location "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin"
    Import-Module .\ConfigurationManager.psd1
    Set-Location "$Sitecode"
    
    $FullName = "$wimName ($wimVersion)"
    
    Write-Host "Creating OS Image in Configuration Manager..." -ForegroundColor Green
    $objImage = New-CMOperatingSystemImage -Name $FullName -Path $ContentPath -Version $wimVersion -Description "Automated upload via Yoink4CM"
    
    $OSImageParent = 'OperatingSystemImage\Automatic Uploads'
    if (-not (Test-Path -Path "OSImage:\$OSImageParent")) {
        New-CMFolder -ParentFolderPath "OperatingSystemImage" -Name "Automatic Uploads" | Out-Null
    }
    Move-CMObject -FolderPath $OSImageParent -InputObject $objImage

    if ($DistributionPointGroup) {
        Write-Host "Distributing to $DistributionPointGroup..." -ForegroundColor Gray
        Start-CMContentDistribution -OperatingSystemImageName $FullName -DistributionPointGroupName $DistributionPointGroup
    }

    Set-Location $RestoreLocation
    Write-Host "`nDone!" -ForegroundColor Green

} else {
    Write-Host "Error: No valid file detected." -ForegroundColor Red
}

Read-Host -Prompt "`nPress Enter to exit"