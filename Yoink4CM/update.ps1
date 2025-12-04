$LocalPath = "C:\Program Files\Yoink Software\Yoink4CM\yoink4cm.ps1"
$RemoteUrl = "https://www.yoink4cm.com/updates/yoink4cm/yoink4cm.ps1"


 if (-not (Test-Path $LocalPath)) {
        Write-Warning "Local script not found at '$LocalPath'. Cannot compare or update."
        return $false
    }

    try {
        $localHash = (Get-FileHash -Path $LocalPath -Algorithm SHA256).Hash
        Write-Host "Local script SHA256 hash: $localHash"
    }
    catch {
        Write-Error "Failed to get hash for local script: $($_.Exception.Message)"
        return $false
    }

    # --- Step 2: Download the remote script and get its hash ---
    $tempFilePath = Join-Path ([System.IO.Path]::GetTempPath()) "yoink4cm.ps1"

    try {
        Write-Host "Downloading remote script from '$RemoteUrl' to '$tempFilePath'..."
        Invoke-WebRequest -Uri $RemoteUrl -OutFile $tempFilePath -UseBasicParsing -ErrorAction Stop
        Write-Host "Remote script downloaded successfully."

        $remoteHash = (Get-FileHash -Path $tempFilePath -Algorithm SHA256).Hash
        Write-Host "Remote script SHA256 hash: $remoteHash"
    }
    catch {
        Write-Error "Failed to download or get hash for remote script: $($_.Exception.Message)"
        return $false
    }
    
    # --- Step 3: Compare the hashes ---
    if ($localHash -ne $remoteHash) {
        Write-Host ""
        Write-Host "UPDATE AVAILABLE!" -ForegroundColor Yellow

        # --- Step 4: Overwrite the local file if requested ---
       
            try {
                # Ensure the directory exists before copying
                $localDirPath = Split-Path -Path $LocalPath -Parent
                if (-not (Test-Path $localDirPath)) {
                    New-Item -ItemType Directory -Path $localDirPath -Force | Out-Null
                }

                Copy-Item -Path $tempFilePath -Destination $LocalPath -Force -ErrorAction Stop
                Write-Host "Local script successfully updated to the latest version!" -ForegroundColor Green
            }
            catch {
                Write-Error "Failed to overwrite local script: $($_.Exception.Message)"
				Write-Host "Please run Configuration Manager console as Administrator, or manually run C:\Program Files\Yoink Software\Yoink4CM\update.ps1 as administrator." -ForegroundColor Red
            }
        
        

    }
    else {
        Write-Host ""
      
        Write-Host "Script is up to date!" -ForegroundColor Green
    

    }
    
        # --- Step 5: Clean up the temporary file as the very last step ---
    if (Test-Path $tempFilePath) {
        Remove-Item $tempFilePath -Force -ErrorAction SilentlyContinue
        Write-Host "Cleaned up temporary file: $tempFilePath"    
    }
	

