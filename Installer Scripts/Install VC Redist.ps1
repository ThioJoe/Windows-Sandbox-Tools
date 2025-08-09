# URLs for the latest Visual C++ Redistributables
$urls = @(
    "https://aka.ms/vs/17/release/vc_redist.x86.exe",
    "https://aka.ms/vs/17/release/vc_redist.x64.exe"
)
if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64') {
    $urls += "https://aka.ms/vs/17/release/vc_redist.arm64.exe"
}

# Directory to save the downloads
$downloadPath = "$env:TEMP"

# To improve download performance, the progress bar is suppressed. [2, 6]
$ProgressPreference = 'SilentlyContinue'

foreach ($url in $urls) {
    $fileName = $url.Split('/')[-1]
    $filePath = Join-Path $downloadPath $fileName

    Write-Host "Downloading $fileName..."
    # Download the file without a progress bar [1, 4]
    Invoke-WebRequest -Uri $url -OutFile $filePath

    if (Test-Path $filePath) {
        Write-Host "Installing $fileName..."
        # Silently install the redistributable and wait for it to complete [3, 5, 9]
        Start-Process -FilePath $filePath -ArgumentList "/install /quiet /norestart" -Wait
        Write-Host "$fileName has been installed."
        # Optional: Remove the installer after installation
        # Remove-Item -Path $filePath
    } else {
        Write-Host "Error: Failed to download $fileName."
    }
}

# Restore the default progress preference
$ProgressPreference = 'Continue'

Write-Host "Script execution finished."
