# Downloads the latest VC Redistributables from Microsoft

param(
    # Optional path to a local directory containing the installation files. If provided, the download steps will be skipped.
    #   - Just copy the entire "VCRedist Install" folder with the files the script normally downloads, and put it in your mounted folder to avoid having to re-download it.
    #   - Make sure to use the mounted path from the perspective of within the sandbox.
    [string]$ExistingInstallerFilesPath,

    # If set, forces using the installer files from ExistingInstallerFilesPath even if a new version exists. Still warns about there being a new version.
    [switch]$ForceCachedFilesOnly,

    # If set, implies ForceCachedFilesOnly and uses cached installer files only, but does not even bother checking for a latest version.
    [switch]$NoCheckLatestVersion
)

# --- Parameter Usage Examples ---
# Standard run (Download & Install):
#    .\Install-VC-Redist.ps1
#
# Install from existing files instead of downloading:
#    .\Install-VC-Redist.ps1 -ExistingInstallerFilesPath "C:\Users\WDAGUtilityAccount\Desktop\HostShared\VCRedist Install"
#
# Use cached files only, even if outdated (still warns about new versions):
#    .\Install-VC-Redist.ps1 -ExistingInstallerFilesPath "C:\Users\WDAGUtilityAccount\Desktop\HostShared\VCRedist Install" -ForceCachedFilesOnly
#
# Use cached files only, skip all version checks:
#    .\Install-VC-Redist.ps1 -ExistingInstallerFilesPath "C:\Users\WDAGUtilityAccount\Desktop\HostShared\VCRedist Install" -NoCheckLatestVersion

# =======================================================

# Validate parameter combinations
if ($NoCheckLatestVersion) {
    $ForceCachedFilesOnly = [switch]::new($true)
}
if ($ForceCachedFilesOnly -and [string]::IsNullOrWhiteSpace($ExistingInstallerFilesPath)) {
    Write-Host "Error: -ForceCachedFilesOnly requires -ExistingInstallerFilesPath to be specified." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# URLs for the latest Visual C++ Redistributables
$urls = @(
    "https://aka.ms/vs/17/release/vc_redist.x86.exe",
    "https://aka.ms/vs/17/release/vc_redist.x64.exe"
)
if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64') {
    $urls += "https://aka.ms/vs/17/release/vc_redist.arm64.exe"
}

# Directory to save the downloads. This will save it into the user "Downloads" folder.
$folderName = "VCRedist Install"
$userDownloadsFolder = (New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path
$downloadPath = Join-Path -Path $userDownloadsFolder -ChildPath $folderName

# To improve download performance, the progress bar is suppressed.
$ProgressPreference = 'SilentlyContinue'

# Load stored version hashes from the pre-existing installer cache (read-only mount)
$storedHashes = @{}
if (-not [string]::IsNullOrWhiteSpace($ExistingInstallerFilesPath)) {
    $hashFilePath = Join-Path $ExistingInstallerFilesPath "versions.txt"
    if (Test-Path $hashFilePath) {
        Get-Content $hashFilePath | ForEach-Object {
            $parts = $_ -split '=', 2
            if ($parts.Count -eq 2) { $storedHashes[$parts[0]] = $parts[1] }
        }
    }
}

# Track current remote hashes to save at the end
$currentHashes = @{}
$pauseAtEnd = $false

foreach ($url in $urls) {
    $fileName = $url.Split('/')[-1]
    $downloadFilePath = Join-Path $downloadPath $fileName
    $installerToRun = $null
    $updateDetected = $false
    $remoteHash = $null

    Write-Host "`nChecking $fileName..."

    if (-not $NoCheckLatestVersion) {
        try {
            # Get the final redirected URL to extract the version hash
            $headResponse = Invoke-WebRequest -Uri $url -Method Head -UseBasicParsing -ErrorAction Stop
            if ($headResponse.BaseResponse.ResponseUri) {
                # PowerShell 5.1
                $finalUrl = $headResponse.BaseResponse.ResponseUri.AbsoluteUri
            } else {
                # PowerShell 7+
                $finalUrl = $headResponse.BaseResponse.RequestMessage.RequestUri.AbsoluteUri
            }
            $remoteHash = $finalUrl.Split('/')[-2]
            $currentHashes[$fileName] = $remoteHash
            Write-Host "Remote hash: $remoteHash"
        }
        catch {
            Write-Host "Warning: Could not retrieve remote info for $fileName." -ForegroundColor Yellow
            Write-Host "Error Info: $_" -ForegroundColor Yellow
        }
    }

    # Check if a pre-existing cached installer can be used
    if (-not [string]::IsNullOrWhiteSpace($ExistingInstallerFilesPath)) {
        $cachedFilePath = Join-Path $ExistingInstallerFilesPath $fileName
        if (Test-Path $cachedFilePath) {
            if ($null -ne $remoteHash -and $storedHashes.ContainsKey($fileName) -and $storedHashes[$fileName] -eq $remoteHash) {
                Write-Host "Cached version is up to date. Using $cachedFilePath"
                $installerToRun = $cachedFilePath
            } elseif ($ForceCachedFilesOnly) {
                if ($null -ne $remoteHash) {
                    Write-Host "WARNING: Cached installer is out of date, but using it anyway (-ForceCachedFilesOnly)." -ForegroundColor Yellow
                    Write-Host "Please update your cache at: $ExistingInstallerFilesPath" -ForegroundColor Yellow
                    $pauseAtEnd = $true
                } else {
                    Write-Host "Using cached file (version check skipped). Using $cachedFilePath"
                }
                $installerToRun = $cachedFilePath
            } else {
                Write-Host "Newer version detected. Cached installer is out of date." -ForegroundColor Yellow
                $updateDetected = $true
                $pauseAtEnd = $true
            }
        } elseif ($ForceCachedFilesOnly) {
            Write-Host "Error: Cached file not found at $cachedFilePath and -ForceCachedFilesOnly is set." -ForegroundColor Red
            continue
        }
    }

    # Download if no valid cache was found or an update is needed
    if ($null -eq $installerToRun) {
        if ($ForceCachedFilesOnly) {
            Write-Host "Error: No cached installer available for $fileName and -ForceCachedFilesOnly is set. Skipping." -ForegroundColor Red
            continue
        }

        Write-Host "Downloading $fileName..."

        # Create the directory if it doesn't exist
        if (-not (Test-Path -Path $downloadPath)) {
            New-Item -Path $downloadPath -ItemType Directory -Force | Out-Null
        }

        Invoke-WebRequest -Uri $url -OutFile $downloadFilePath -UseBasicParsing
        $installerToRun = $downloadFilePath
        
        if ($updateDetected) {
            Write-Host "ACTION REQUIRED: A new version of $fileName was downloaded." -ForegroundColor Yellow
            Write-Host "Please update your cache at: $ExistingInstallerFilesPath" -ForegroundColor Yellow
            Write-Host "Then update versions.txt with: $fileName=$remoteHash" -ForegroundColor Yellow
            $pauseAtEnd = $true
        }
    }

    if (Test-Path $installerToRun) {
        Write-Host "Installing $fileName from $installerToRun..."
        # Silently install the redistributable and wait for it to complete
        Start-Process -FilePath $installerToRun -ArgumentList "/install /quiet /norestart" -Wait
        Write-Host "$fileName has been installed."
        
        # Optional: Remove the downloaded installer if it was downloaded to TEMP
        if ($installerToRun -eq $downloadFilePath) {
            # Remove-Item -Path $downloadFilePath
        }
    } else {
        Write-Host "Error: Failed to locate installer for $fileName."
    }
}

# Save versions.txt to the download folder so it can be copied alongside the installers
if ($currentHashes.Count -gt 0) {
    if (-not (Test-Path -Path $downloadPath)) {
        New-Item -Path $downloadPath -ItemType Directory -Force | Out-Null
    }
    $versionsFilePath = Join-Path $downloadPath "versions.txt"
    $currentHashes.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" } | Set-Content $versionsFilePath
}

# Restore the default progress preference
$ProgressPreference = 'Continue'

Write-Host "`nScript execution finished."
if ($pauseAtEnd) {
    Read-Host "`nPress Enter to exit"
}
