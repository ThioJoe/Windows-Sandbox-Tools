# This script will install Winget from within the Windows Sandbox
# It fetches the necessary files and dependencies from Microsoft's winget-cli, and installs them

# Author: ThioJoe
# Repo Url: https://github.com/ThioJoe/Windows-Sandbox-Tools
# Last Updated: August 4, 2025

param(
        [switch]$removeMsStoreAsSource = $false # If switch is included, it will remove the 'msstore' source after installing winget, which doesn't work with Sandbox, unless the Microsoft Store is also installed
    )

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

# Prevents progress bar from showing (often speeds downloads)
$ProgressPreference = 'SilentlyContinue'

$downloadPath = Join-Path $env:USERPROFILE "Downloads"
$latestRelease = Get-LatestRelease
if (-not $latestRelease) { Write-Error "Could not retrieve the latest release. Exiting."; return; }

$latestTag = $latestRelease.tag_name
Write-Host "Latest winget version tag is: $latestTag"

# Download the MSIX bundle
$msixName = "Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"
$msixUrl = Get-AssetUrl -release $latestRelease -assetName $msixName
if (-not $msixUrl) { Write-Error "Could not find $msixName in the latest release assets."; return; }

Write-Host "Downloading $msixName..."
$msixPath = Join-Path $downloadPath $msixName
Invoke-WebRequest -Uri $msixUrl -OutFile $msixPath

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

# Download the dependencies zip
$depsZipName = "DesktopAppInstaller_Dependencies.zip"
$depsZipUrl  = Get-AssetUrl -release $latestRelease -assetName $depsZipName

# We'll expand to a base 'Dependencies' folder
$topDepsFolder = Join-Path $downloadPath "Dependencies"
# Then pick the sub-folder for the architecture
$depsFolder    = Join-Path $topDepsFolder $arch

if ($depsZipUrl) {
    Write-Host "Downloading $depsZipName..."
    $depsZipPath = Join-Path $downloadPath $depsZipName
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
} 
else { Write-Warning "No $depsZipName found in $latestTag, skipping dependency download."; }

# Restore progress preference
$ProgressPreference = 'Continue'

# If dependencies exist for this architecture, install them
if (Test-Path $depsFolder) {
    Install-WingetDependencies -depsFolder $depsFolder
} else {
    Write-Warning "No architecture-specific dependencies found at $depsFolder"
}

# Finally, install the winget MSIX bundle
Write-Host "Installing $msixName..."
Add-AppxPackage -Path $msixPath

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
