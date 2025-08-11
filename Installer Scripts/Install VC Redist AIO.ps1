# URLs for the latest Visual C++ Redistributables
$urls = @(
    "https://github.com/abbodi1406/vcredist/releases/latest/download/VisualCppRedist_AIO_x86_x64.exe"  # https://github.com/abbodi1406/vcredist
)

# VisualCppRedist AIO does not include any ARM64 installers,
# and it's not planned to have ARM64 support.
# https://github.com/abbodi1406/vcredist/issues/110
if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64') {
    $urls += "https://aka.ms/vs/17/release/vc_redist.arm64.exe"
}

# Directory to save the downloads
$downloadPath = "$env:TEMP"

# https://stackoverflow.com/a/25127597
Function Get-RedirectedUrl {
    Param (
        [Parameter(Mandatory=$true)]
        [String]$URL
    )
    $request = [System.Net.WebRequest]::Create($URL)
    $request.Method = "HEAD"
    $request.AllowAutoRedirect=$false
    $response=$request.GetResponse()

    If ($response.StatusCode -eq "Found")
    {
        $response.GetResponseHeader("Location")
    }
}

# To improve download performance, the progress bar is suppressed. [2, 6]
$ProgressPreference = 'SilentlyContinue'

# There is a bug that makes MSI installs take *FOREVER* to finish.
# https://github.com/microsoft/Windows-Sandbox/issues/68
# There is a solution: temporarily turn off Smart App Control.
# Thanks, Traxof63!
Function Set-SmartAppControl {
    Param (
        [Parameter(Mandatory=$true)]
        [String]$Num
    )
    Write-Host "Setting CI Policy 'VerifiedAndReputablePolicyState' to $Num..."
    Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\CI\Policy" -Name "VerifiedAndReputablePolicyState" -Value $Num
    # Only problem is, CiTool does not exit without user interaction. This means
    # that we create a process that uses around 664k of Memory... Ugh, Microsoft!
    Start-Process -FilePath "CiTool.exe" -ArgumentList "-r" -WindowStyle Hidden
}

Set-SmartAppControl "0"

foreach ($url in $urls) {
    $redirectUrl = [System.IO.Path]::GetFileName((Get-RedirectedUrl $url))
    $fileName = $redirectUrl.Split('/')[-1]
    $filePath = Join-Path $downloadPath $fileName

    Write-Host "Downloading $fileName..."
    Write-Host $filePath
    # Download the file without a progress bar [1, 4]
    Invoke-WebRequest -Uri $url -OutFile $filePath

    if (Test-Path $filePath) {
        Write-Host "Installing $fileName..."
        # Silently install the redistributable and wait for it to complete [3, 5, 9]
        if ($url -match "https://aka.ms/vs/17/release/vc_redist.arm64.exe") {
            Start-Process -FilePath $filePath -ArgumentList "/install /quiet /norestart" -Wait
        } else {
            Start-Process -FilePath $filePath -ArgumentList "/ai" -Wait
        }
        Write-Host "$fileName has been installed."
        # Optional: Remove the installer after installation
        # Remove-Item -Path $filePath
    } else {
        Write-Host "Error: Failed to download $fileName."
    }
}

# Turn Smart App Control back on
Set-SmartAppControl "1"

# Restore the default progress preference
$ProgressPreference = 'Continue'

Write-Host "Script execution finished."
