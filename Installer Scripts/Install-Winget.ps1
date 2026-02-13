# This script will install Winget from within the Windows Sandbox
# It fetches the necessary files and dependencies from Microsoft's winget-cli, and installs them

# Author: ThioJoe
# Repo Url: https://github.com/ThioJoe/Windows-Sandbox-Tools
# Last Updated: February 12, 2026

param(
    # If switch is included, it will remove the 'msstore' source after installing winget, which doesn't work with Sandbox, unless the Microsoft Store is also installed
    [switch]$removeMsStoreAsSource = $false,

    # Optional path to a local directory containing the installation files. If provided, the download steps will be skipped.
    #   - Just copy the entire "Winget Install" folder with the files the script normally downloads, and put it in your mounted folder.
    #   - Make sure to use the mounted path from the perspective of within the sandbox.
    [string]$ExistingInstallerFilesPath
)

# --- Parameter Usage Examples ---
# Standard run (Download & Install):
#    .\Install-Winget.ps1
#
# Install from existing files instead of downloading:
#    .\Install-Winget.ps1 -ExistingInstallerFilesPath "C:\Users\WDAGUtilityAccount\Desktop\HostShared\Winget Install"

function Get-LatestRelease {
    param(
        [string]$repoOwner = 'microsoft',
        [string]$repoName = 'winget-cli'
    )
    try {
        $releasesUrl = "https://api.github.com/repos/$repoOwner/$repoName/releases"
        $releases = Invoke-RestMethod -Uri $releasesUrl -UseBasicParsing
    } catch {
        Write-Error "Failed to fetch releases from GitHub API: $($_.Exception.Message)"
        return $null
    }

    if (-not $releases) { Write-Error "No releases found for $repoOwner/$repoName."; return $null; }

    # Pick the top entry once sorted by published_at descending
    $latestRelease = $releases | Sort-Object -Property published_at -Descending | Select-Object -First 1
    return $latestRelease
}

function Get-AssetUrl {
    param(
        [Parameter(Mandatory=$true)]
        $release,
        [Parameter(Mandatory=$true)]
        [string]$assetName
    )

    if ($release.assets -and $release.assets.Count -gt 0) {
        $asset = $release.assets | Where-Object { $_.name -eq $assetName }
        if ($asset) {
            return $asset.browser_download_url
        }
    }
    return $null
}

function Install-WingetDependencies {
    param([string]$depsFolder)

    # Look for DesktopAppInstaller_Dependencies.json to determine explicit install order
    $jsonFile = Join-Path $depsFolder "DesktopAppInstaller_Dependencies.json"
    if (Test-Path $jsonFile) {
        Write-Host "Installing dependencies based on DesktopAppInstaller_Dependencies.json"
        $jsonContent = Get-Content $jsonFile -Raw | ConvertFrom-Json
        $dependencies = $jsonContent.Dependencies

        foreach ($dep in $dependencies) {
            # For example: "Microsoft.VCLibs.140.00.UWPDesktop" + "14.0.33728.0"
            $matchingFiles = Get-ChildItem -Path $depsFolder -Filter *.appx -Recurse |
                Where-Object { $_.Name -like "*$($dep.Name)*" -and $_.Name -like "*$($dep.Version)*" }

            foreach ($file in $matchingFiles) {
                Write-Host "Installing dependency: $($file.Name)"
                Add-AppxPackage -Path $file.FullName
            }
        }
    }
    else {
        # If the JSON doesn't exist, install all .appx in the folder
        Write-Warning "No DesktopAppInstaller_Dependencies.json found, installing all .appx in $depsFolder"
        foreach ($appxFile in Get-ChildItem $depsFolder -Filter *.appx -Recurse) {
            Write-Host "Installing: $($appxFile.Name)"
            Add-AppxPackage -Path $appxFile.FullName
        }
    }
}

# --- Define Working Directory ---
$userDownloadsFolder = (New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path
$subfolderName = "Winget Install"
$msixName = "Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"

if ($ExistingInstallerFilesPath) {
    if (Test-Path -Path $ExistingInstallerFilesPath) {
        $workingDir = $ExistingInstallerFilesPath
        Write-Host "Using local source path: $workingDir" -ForegroundColor Yellow
    } else {
        Write-Error "The specified local source path does not exist: $ExistingInstallerFilesPath"
        return
    }
} else {
    # Combine them to create the full working directory path
    $workingDir = Join-Path -Path $userDownloadsFolder -ChildPath $subfolderName
    
    # Create the directory if it doesn't exist
    if (-not (Test-Path -Path $workingDir)) {
        New-Item -Path $workingDir -ItemType Directory -Force | Out-Null
    }
}

# Prevents progress bar from showing (often speeds downloads)
$ProgressPreference = 'SilentlyContinue'

# --- Determine Architecture ---
# Figure out the OS architecture using environment variable
$procArch = $env:PROCESSOR_ARCHITECTURE
switch -Wildcard ($procArch) {
    "AMD64"   { $arch = "x64" }
    "x86"     { $arch = "x86" }
    "*ARM64*" { $arch = "arm64" }
    "*ARM*"   { $arch = "arm" }
    default {
        $arch = "x64"
        Write-Warning "Unrecognized architecture: $procArch. Defaulting to x64."
    }
}

# --- Download Steps (Skipped if using existing files) ---
if (-not $ExistingInstallerFilesPath) {
    $latestRelease = Get-LatestRelease
    if (-not $latestRelease) { Write-Error "Could not retrieve the latest release. Exiting."; return; }

    $latestTag = $latestRelease.tag_name
    Write-Host "Latest winget version tag is: $latestTag"

    # Download the MSIX bundle
    $msixUrl = Get-AssetUrl -release $latestRelease -assetName $msixName
    if (-not $msixUrl) { Write-Error "Could not find $msixName in the latest release assets."; return; }

    Write-Host "Downloading $msixName..."
    $msixPath = Join-Path $workingDir $msixName
    Invoke-WebRequest -Uri $msixUrl -OutFile $msixPath

    # Download the dependencies zip
    $depsZipName = "DesktopAppInstaller_Dependencies.zip"
    $depsZipUrl  = Get-AssetUrl -release $latestRelease -assetName $depsZipName

    # We'll expand to a base 'Dependencies' folder
    $topDepsFolder = Join-Path $workingDir "Dependencies"

    if ($depsZipUrl) {
        Write-Host "Downloading $depsZipName..."
        $depsZipPath = Join-Path $workingDir $depsZipName
        Invoke-WebRequest -Uri $depsZipUrl -OutFile $depsZipPath

        # Remove existing Dependencies folder and expand the zip
        if (Test-Path $topDepsFolder) { Remove-Item -Path $topDepsFolder -Recurse -Force }
        
        # Use Expand-Archive cmdlet by default because it's safe for constrained language mode. Fall back to .NET assembly if it fails.
        try {
            Expand-Archive -LiteralPath $depsZipPath -DestinationPath $topDepsFolder -Force -ErrorAction Stop
        }
        catch {
            Write-Warning "Standard extraction failed, attempting .NET fallback. The error was: $($_.Exception.Message)"
            # Fallback using .NET System.IO.Compression (Fixes issues in non-EN Windows Sandbox)
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            [System.IO.Compression.ZipFile]::ExtractToDirectory($depsZipPath, $topDepsFolder)
        }

        # Cleanup the zip file
        if (Test-Path $depsZipPath) {
            Remove-Item -Path $depsZipPath -Force
        }
    } 
    else { Write-Warning "No $depsZipName found in $latestTag, skipping dependency download."; }
}

# Restore progress preference
$ProgressPreference = 'Continue'

# --- Installation Steps ---

# Define paths based on working directory
$msixPath = Join-Path $workingDir $msixName
$topDepsFolder = Join-Path $workingDir "Dependencies"
$depsFolder = Join-Path $topDepsFolder $arch

# Install Dependencies
if (Test-Path $depsFolder) {
    Install-WingetDependencies -depsFolder $depsFolder
} else {
    if ($ExistingInstallerFilesPath) {
        Write-Error "Dependencies folder not found at: $depsFolder`nEnsure the 'Dependencies' folder is present in your source directory."
    } else {
        Write-Warning "No architecture-specific dependencies found at $depsFolder"
    }
}

# Install Winget MSIX bundle
if (Test-Path $msixPath) {
    Write-Host "Installing $msixName..."
    Add-AppxPackage -Path $msixPath
} else {
    Write-Error "Winget package not found at: $msixPath"
    return
}

# Remove msstore source if set to do so
if ($removeMsStoreAsSource.IsPresent) {
    Write-Host "Attempting to remove 'msstore' source from winget..."
    try {
        winget source remove -n msstore --ignore-warnings
    } catch {
        Write-Warning "An error occurred while trying to execute 'winget source remove msstore': $($_.Exception.Message)"
    }
} else {
    # Automatically accept source agreements to avoid prompts. Mostly applies to msstore.
    winget list --accept-source-agreements | Out-Null
}