# This script will install the Microsoft Store from Within the Windows Sandbox
#     It uses the Windows Update API to fetch the necessary installation files directly from Microsoft
#     Unlike many similar scripts, it uses NO dependencies or third party APIs

# Author: ThioJoe
# Repo Url: https://github.com/ThioJoe/Windows-Sandbox-Tools
# Last Updated: February 12, 2026

param(
    # Optional switch to output the generated XML files to the working directory
    [switch]$debugSaveFiles,
    
    # Optional switch to skip the installation of Microsoft Store, but still download the files
    [switch]$noInstall,

    # Optional switch to skip the download and install, but still show the packages found 
    [switch]$noDownload,

    # Optional path to a local directory containing the installation files. If provided, the download steps will be skipped.
    #   - Just copy the entire "MSStore Install" folder with the files the script normally downloads, and put it in your mounted folder to avoid having to re-download it.
    #        - You can find the "MSStore Install" folder in the "Downloads" folder.
    #   - Make sure to use the mounted path from the perspective of within the sandbox.
    [string]$ExistingInstallerFilesPath
)

# --- Parameter Usage Examples ---
# Standard run (Download & Install):
#    .\Install-Microsoft-Store.ps1
#
# Install from existing files instead of downloading:
#    .\Install-Microsoft-Store.ps1 -ExistingInstallerFilesPath "C:\Users\WDAGUtilityAccount\Desktop\HostShared\MSStore Install"
#
# Download only (Don't Install):
#    .\Install-Microsoft-Store.ps1 -noInstall
#
# Debug mode (Save SOAP XML logs to disk):
#    .\Install-Microsoft-Store.ps1 -debugSaveFiles

# =======================================================

# --- Configuration ---
# Note: These defaults should work for the regular current build of Microsoft Store, but I haven't tested using any of the other values. So fetching insider builds of MS Store (if any) might not work.
$flightRing = "Retail"             # Apparently accepts 'Retail', 'Internal', and 'External'
$flightingBranchName = ""          # Empty ( "" ) for normal release. Otherwise apparent possible values: Dev, Beta, ReleasePreview, MSIT, CanaryChannel, external
$currentBranch = "ge_release"      # "rs_prerelease" for insider, "ni_release" for normal release on Windows build below 26100, "ge_release" for normal release equal or above 26100

# Random Notes:
#   flightRing should be "Internal" if flightingBranchName is "MSIT"
#   MAYBE need flightRing as "External" if setting flightingBranchName anything besides empty or MSIT?

# ------ Check that we're running in the Windows Sandbox ------
# This script is intended to be run from within the Windows Sandbox. Warn the user if not. We'll do a rudamentary check for if the current user is named "WDAGUtilityAccount"
if ($env:USERNAME -ne "WDAGUtilityAccount") {
    Write-Warning "`nThis script is intended to be run from WITHIN the Windows Sandbox.`nRunning it outside the Sandbox will just install the MS Store to the current system."
    Write-host "`nPress Enter to continue installing to the current system anyway, or just close this window to exit." -ForegroundColor Yellow
    Read-Host
}

# --- Define Working Directory ---
# Get the path to the user's personal Downloads folder in a reliable way
$userDownloadsFolder = (New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path

# Define the subfolder name for all our files
$subfolderName = "MSStore Install"

# Category ID for the Microsoft Store app package
$storeCategoryId = "64293252-5926-453c-9494-2d4021f1c78d" 

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

If ($debugSaveFiles) {
    Write-Host "All files (logs, downloads) will be saved to: '$workingDir'" -ForegroundColor Yellow
}

# --- XML Templates ---

# Step 1: GetCookie request body.
# See: https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-wusp/36a5d99a-a3ca-439d-bcc5-7325ff6b91e2
$cookieXmlTemplate = @"
<s:Envelope xmlns:s="http://www.w3.org/2003/05/soap-envelope" xmlns:a="http://www.w3.org/2005/08/addressing">
    <s:Header>
        <a:Action s:mustUnderstand="1">http://www.microsoft.com/SoftwareDistribution/Server/ClientWebService/GetCookie</a:Action>
        <a:MessageID>urn:uuid:$(New-Guid)</a:MessageID>
        <a:To s:mustUnderstand="1">https://fe3.delivery.mp.microsoft.com/ClientWebService/client.asmx</a:To>
        <o:Security s:mustUnderstand="1" xmlns:o="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd">
            <wuws:WindowsUpdateTicketsToken wsu:id="ClientMSA" xmlns:wsu="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd" xmlns:wuws="http://schemas.microsoft.com/msus/2014/10/WindowsUpdateAuthorization">
                <TicketType Name="MSA" Version="1.0" Policy="MBI_SSL"><user></user></TicketType>
            </wuws:WindowsUpdateTicketsToken>
        </o:Security>
    </s:Header>
    <s:Body><GetCookie xmlns="http://www.microsoft.com/SoftwareDistribution/Server/ClientWebService" /></s:Body>
</s:Envelope>
"@

# Step 2: SyncUpdates request body. Based on intercepted XML request using Microsoft Store.
# Info about attributes found here: https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-wusp/6b654980-ae63-4b0d-9fae-2abb516af894
$fileListXmlTemplate = @"
<s:Envelope xmlns:a="http://www.w3.org/2005/08/addressing" xmlns:s="http://www.w3.org/2003/05/soap-envelope">
    <s:Header>
        <a:Action s:mustUnderstand="1">http://www.microsoft.com/SoftwareDistribution/Server/ClientWebService/SyncUpdates</a:Action>
        <a:MessageID>urn:uuid:$(New-Guid)</a:MessageID>
        <a:To s:mustUnderstand="1">https://fe3cr.delivery.mp.microsoft.com/ClientWebService/client.asmx</a:To>
        <o:Security s:mustUnderstand="1" xmlns:o="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd">
            <Timestamp xmlns="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">
                <Created>$((Get-Date).ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'"))</Created>
                <Expires>$((Get-Date).AddMinutes(5).ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'"))</Expires>
            </Timestamp>
            <wuws:WindowsUpdateTicketsToken wsu:id="ClientMSA" xmlns:wsu="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd" xmlns:wuws="http://schemas.microsoft.com/msus/2014/10/WindowsUpdateAuthorization">
                <TicketType Name="MSA" Version="1.0" Policy="MBI_SSL">
                    <user/>
                </TicketType>
            </wuws:WindowsUpdateTicketsToken>
        </o:Security>
    </s:Header>
    <s:Body>
        <SyncUpdates xmlns="http://www.microsoft.com/SoftwareDistribution/Server/ClientWebService">
            <cookie>
                <Expiration>$((Get-Date).AddYears(10).ToUniversalTime().ToString('u').Replace(' ','T'))</Expiration>
                <EncryptedData>{0}</EncryptedData>
            </cookie>
            <parameters>
                <ExpressQuery>false</ExpressQuery>
                <InstalledNonLeafUpdateIDs>
                    <int>1</int><int>2</int><int>3</int><int>11</int><int>19</int><int>2359974</int><int>5169044</int>
                    <int>8788830</int><int>23110993</int><int>23110994</int><int>54341900</int><int>59830006</int><int>59830007</int>
                    <int>59830008</int><int>60484010</int><int>62450018</int><int>62450019</int><int>62450020</int><int>98959022</int>
                    <int>98959023</int><int>98959024</int><int>98959025</int><int>98959026</int><int>104433538</int><int>129905029</int>
                    <int>130040031</int><int>132387090</int><int>132393049</int><int>133399034</int><int>138537048</int><int>140377312</int>
                    <int>143747671</int><int>158941041</int><int>158941042</int><int>158941043</int><int>158941044</int><int>159123858</int>
                    <int>159130928</int><int>164836897</int><int>164847386</int><int>164848327</int><int>164852241</int><int>164852246</int>
                    <int>164852253</int>
                </InstalledNonLeafUpdateIDs>
                <SkipSoftwareSync>false</SkipSoftwareSync>
                <NeedTwoGroupOutOfScopeUpdates>false</NeedTwoGroupOutOfScopeUpdates>
                <FilterAppCategoryIds>
                    <CategoryIdentifier>
                        <Id>{1}</Id>
                    </CategoryIdentifier>
                </FilterAppCategoryIds>
                <TreatAppCategoryIdsAsInstalled>true</TreatAppCategoryIdsAsInstalled>
                <AlsoPerformRegularSync>false</AlsoPerformRegularSync>
                <ComputerSpec/>
                <ExtendedUpdateInfoParameters>
                    <XmlUpdateFragmentTypes>
                        <XmlUpdateFragmentType>Extended</XmlUpdateFragmentType>
                    </XmlUpdateFragmentTypes>
                    <Locales>
                        <string>en-US</string>
                        <string>en</string>
                    </Locales>
                </ExtendedUpdateInfoParameters>
                <ClientPreferredLanguages>
                    <string>en-US</string>
                </ClientPreferredLanguages>
                <ProductsParameters>
                    <SyncCurrentVersionOnly>false</SyncCurrentVersionOnly>
                    <DeviceAttributes>E:BranchReadinessLevel=CB&amp;CurrentBranch={2}&amp;OEMModel=Virtual%20Machine&amp;FlightRing={3}&amp;AttrDataVer=321&amp;InstallLanguage=en-US&amp;OSUILocale=en-US&amp;InstallationType=Client&amp;FlightingBranchName={4}&amp;OSSkuId=48&amp;App=WU_STORE&amp;ProcessorManufacturer=GenuineIntel&amp;OEMName_Uncleaned=Microsoft%20Corporation&amp;AppVer=1407.2503.28012.0&amp;OSArchitecture=AMD64&amp;IsFlightingEnabled=1&amp;TelemetryLevel=1&amp;DefaultUserRegion=39070&amp;WuClientVer=1310.2503.26012.0&amp;OSVersion=10.0.26100.3915&amp;DeviceFamily=Windows.Desktop</DeviceAttributes>
                    <CallerAttributes>Interactive=1;IsSeeker=1;</CallerAttributes>
                    <Products/>
                </ProductsParameters>
            </parameters>
        </SyncUpdates>
    </s:Body>
</s:Envelope>
"@

# Step 3: GetExtendedUpdateInfo2 - After getting the list of matched files (app version and dependencies), this lets us get the actual download URLs
# See: https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-wusp/2f66a682-164f-47ec-968e-e43c0a85dc21
$fileUrlXmlTemplate = @"
<s:Envelope xmlns:s="http://www.w3.org/2003/05/soap-envelope" xmlns:a="http://www.w3.org/2005/08/addressing">
    <s:Header>
        <a:Action s:mustUnderstand="1">http://www.microsoft.com/SoftwareDistribution/Server/ClientWebService/GetExtendedUpdateInfo2</a:Action>
        <a:MessageID>urn:uuid:$(New-Guid)</a:MessageID>
        <a:To s:mustUnderstand="1">https://fe3cr.delivery.mp.microsoft.com/ClientWebService/client.asmx/secured</a:To>
        <o:Security s:mustUnderstand="1" xmlns:o="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd">
            <u:Timestamp u:Id="_0" xmlns:u="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">
                <u:Created>$((Get-Date).ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'"))</u:Created>
                <u:Expires>$((Get-Date).AddMinutes(5).ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'"))</u:Expires>
            </u:Timestamp>
            <wuws:WindowsUpdateTicketsToken wsu:id="ClientMSA" xmlns:wsu="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd" xmlns:wuws="http://schemas.microsoft.com/msus/2014/10/WindowsUpdateAuthorization">
                <TicketType Name="MSA" Version="1.0" Policy="MBI_SSL"><user>{0}</user></TicketType>
            </wuws:WindowsUpdateTicketsToken>
        </o:Security>
    </s:Header>
    <s:Body>
        <GetExtendedUpdateInfo2 xmlns="http://www.microsoft.com/SoftwareDistribution/Server/ClientWebService">
            <updateIDs><UpdateIdentity><UpdateID>{1}</UpdateID><RevisionNumber>{2}</RevisionNumber></UpdateIdentity></updateIDs>
            <infoTypes><XmlUpdateFragmentType>FileUrl</XmlUpdateFragmentType></infoTypes>
            <DeviceAttributes>E:BranchReadinessLevel=CB&amp;CurrentBranch={3}&amp;OEMModel=Virtual%20Machine&amp;FlightRing={4}&amp;AttrDataVer=321&amp;InstallLanguage=en-US&amp;OSUILocale=en-US&amp;InstallationType=Client&amp;FlightingBranchName={5}&amp;OSSkuId=48&amp;App=WU_STORE&amp;ProcessorManufacturer=GenuineIntel&amp;OEMName_Uncleaned=Microsoft%20Corporation&amp;AppVer=1407.2503.28012.0&amp;OSArchitecture=AMD64&amp;IsFlightingEnabled=1&amp;TelemetryLevel=1&amp;DefaultUserRegion=39070&amp;WuClientVer=1310.2503.26012.0&amp;OSVersion=10.0.26100.3915&amp;DeviceFamily=Windows.Desktop</DeviceAttributes>
        </GetExtendedUpdateInfo2>
    </s:Body>
</s:Envelope>
"@

# --- Script Execution ---
$headers = @{ "Content-Type" = "application/soap+xml; charset=utf-8" }
$baseUri = "https://fe3.delivery.mp.microsoft.com/ClientWebService/client.asmx"

try {
    if (-not $ExistingInstallerFilesPath) {
        # Step 1: Get Cookie
        Write-Host "Step 1: Getting authentication cookie..."
        $cookieRequestPayload = $cookieXmlTemplate
        If ($debugSaveFiles) { $cookieRequestPayload | Set-Content -Path (Join-Path $LogDirectory "01_Step1_Request.xml") }
        
        $cookieResponse = Invoke-WebRequest -Uri $baseUri -Method Post -Body $cookieRequestPayload -Headers $headers -UseBasicParsing
        If ($debugSaveFiles) { $cookieResponse.Content | Set-Content -Path (Join-Path $LogDirectory "01_Step1_Response.xml"); Write-Host "  -> Saved request and response logs for Step 1." }

        $cookieResponseXml = [xml]$cookieResponse.Content
        $encryptedCookieData = $cookieResponseXml.Envelope.Body.GetCookieResponse.GetCookieResult.EncryptedData
        Write-Host "Success. Cookie received." -ForegroundColor Green

        # Step 2: Get File List
        Write-Host "Step 2: Getting file list..."
        $fileListRequestPayload = $fileListXmlTemplate -f $encryptedCookieData, $storeCategoryId, $currentBranch, $flightRing, $flightingBranchName
        If ($debugSaveFiles) { [System.IO.File]::WriteAllText((Join-Path $LogDirectory "02_Step2_Request_AUTOMATED.xml"), $fileListRequestPayload, [System.Text.UTF8Encoding]::new($false)) }

        $fileListResponse = Invoke-WebRequest -Uri $baseUri -Method Post -Body $fileListRequestPayload -Headers $headers -UseBasicParsing
        If ($debugSaveFiles) { $fileListResponse.Content | Set-Content -Path (Join-Path $LogDirectory "02_Step2_Response_SUCCESS.xml") }

        # The response contains XML fragments that are HTML-encoded. We must decode this before treating it as XML.
        Add-Type -AssemblyName System.Web
        $decodedContent = [System.Web.HttpUtility]::HtmlDecode($fileListResponse.Content)
        $fileListResponseXml = [xml]$decodedContent
        Write-Host "Successfully received and DECODED Step 2 response." -ForegroundColor Green
        
        $fileIdentityMap = @{}
        
        # Get the two main lists of updates from the now correctly-decoded response
        $newUpdates = $fileListResponseXml.Envelope.Body.SyncUpdatesResponse.SyncUpdatesResult.NewUpdates.UpdateInfo
        $allExtendedUpdates = $fileListResponseXml.Envelope.Body.SyncUpdatesResponse.SyncUpdatesResult.ExtendedUpdateInfo.Updates.Update

        Write-Host "--- Correlating Update Information ---" -ForegroundColor Magenta

        # Filter the 'NewUpdates' list to only include items that are actual downloadable files.
        # These are identified by the presence of the <SecuredFragment> tag inside their inner XML.
        $downloadableUpdates = $newUpdates | Where-Object { $_.Xml.Properties.SecuredFragment }

        Write-Host "Found $($downloadableUpdates.Count) potentially downloadable packages." -ForegroundColor Cyan

        # Now, process each downloadable update
        foreach ($update in $downloadableUpdates) {
            $lookupId = $update.ID
            
            # Find the matching entry in the 'ExtendedUpdateInfo' list using the same numeric ID.
            $extendedInfo = $allExtendedUpdates | Where-Object { $_.ID -eq $lookupId } | Select-Object -First 1
            
            if (-not $extendedInfo) {
                Write-Warning "Could not find matching ExtendedInfo for downloadable update ID $lookupId. Skipping."
                continue
            }
            
            # From the extended info, get the actual package file and ignore the metadata .cab files.
            $fileNode = $extendedInfo.Xml.Files.File | Where-Object { $_.FileName -and $_.FileName -notlike "Abm_*" } | Select-Object -First 1

            if (-not $fileNode) {
                Write-Warning "Found matching ExtendedInfo for ID $lookupId, but it contains no valid file node. Skipping."
                continue
            }

            # Additional parsing
            $fileName = $fileNode.FileName
            $updateGuid = $update.Xml.UpdateIdentity.UpdateID
            $revNum = $update.Xml.UpdateIdentity.RevisionNumber
            $fullIdentifier = $fileNode.GetAttribute("InstallerSpecificIdentifier")

            # Define the regex based on the official package identity structure.
            # <Name>_<Version>_<Architecture>_<ResourceId>_<PublisherId>
            $regex = "^(?<Name>.+?)_(?<Version>\d+\.\d+\.\d+\.\d+)_(?<Architecture>[a-zA-Z0-9]+)_(?<ResourceId>.*?)_(?<PublisherId>[a-hjkmnp-tv-z0-9]{13})$"
            
            $packageInfo = [PSCustomObject]@{
                FullName       = $fullIdentifier
                FileName       = $fileName
                UpdateID       = $updateGuid
                RevisionNumber = $revNum
            }

            if ($fullIdentifier -match $regex) {
                # If the regex matches, populate the object with the named capture groups
                $packageInfo | Add-Member -MemberType NoteProperty -Name "PackageName" -Value $matches.Name
                $packageInfo | Add-Member -MemberType NoteProperty -Name "Version" -Value $matches.Version
                $packageInfo | Add-Member -MemberType NoteProperty -Name "Architecture" -Value $matches.Architecture
                $packageInfo | Add-Member -MemberType NoteProperty -Name "ResourceId" -Value $matches.ResourceId
                $packageInfo | Add-Member -MemberType NoteProperty -Name "PublisherId" -Value $matches.PublisherId
            } else {
                # Fallback for any identifiers that don't match the pattern
                $packageInfo | Add-Member -MemberType NoteProperty -Name "PackageName" -Value "Unknown (Parsing Failed)"
                $packageInfo | Add-Member -MemberType NoteProperty -Name "Architecture" -Value "unknown"
            }

            # Use the full, unique identifier as the key in the map
            $fileIdentityMap[$fullIdentifier] = $packageInfo
            
            Write-Host "  -> CORRELATED: '$($packageInfo.PackageName)' ($($packageInfo.Architecture))" -ForegroundColor Green
        }

        Write-Host "--- Correlation Complete ---" -ForegroundColor Magenta
        Write-Host "Found and prepared $($fileIdentityMap.Count) downloadable files." -ForegroundColor Green


        # --- Step 3: Filter, Get URLs, and Download ---
        try {
            # Get the current system's processor architecture and map it to the script's naming convention
            $systemArch = switch ($env:PROCESSOR_ARCHITECTURE) {
                "AMD64" { "x64" }
                "ARM64" { "arm64" }
                "x86"   { "x86" }
                default { "unknown" }
            }
            
            if ($systemArch -eq "unknown") {
                throw "Could not determine system architecture from '$($env:PROCESSOR_ARCHITECTURE)'."
            }
            Write-Host "Step 3: Filtering packages for your system architecture ('$systemArch')..." -ForegroundColor Magenta

            # --- Filter the packages ---

            # 1. Isolate the Microsoft.WindowsStore packages and find the latest version
            $latestStorePackage = $fileIdentityMap.Values |
                Where-Object { $_.PackageName -eq 'Microsoft.WindowsStore' } |
                Sort-Object { [version]$_.Version } -Descending |
                Select-Object -First 1

            # 2. Get all other dependencies that match the system architecture (or are neutral)
            $filteredDependencies = $fileIdentityMap.Values |
                Where-Object {
                    ($_.PackageName -ne 'Microsoft.WindowsStore') -and
                    ( ($_.Architecture -eq $systemArch) -or ($_.Architecture -eq 'neutral') )
                }

            # 3. Combine the lists for the final download queue
            $packagesToDownload = @()
            if ($latestStorePackage) {
                $packagesToDownload += $latestStorePackage
                Write-Host "  -> Found latest Store package: $($latestStorePackage.FullName)" -ForegroundColor Green
            } else {
                Write-Warning "Could not find any Microsoft.WindowsStore package."
            }

            $packagesToDownload += $filteredDependencies
            Write-Host "  -> Found $($filteredDependencies.Count) dependencies for '$systemArch' architecture." -ForegroundColor Green
            Write-Host "Total files to download: $($packagesToDownload.Count)" -ForegroundColor Cyan
            Write-Host "------------------------------------------------------------"


            # --- Loop through the filtered list, get URLs, and download ---
            Write-Host "Step 4: Fetching URLs and downloading files..." -ForegroundColor Magenta

            $originalPref = $ProgressPreference
            $ProgressPreference = 'SilentlyContinue'
            
            foreach ($package in $packagesToDownload) {
                Write-Host "Processing: $($package.FullName)"

                # Get the download URL for this specific package
                $fileUrlRequestPayload = $fileUrlXmlTemplate -f $encryptedCookieData, $package.UpdateID, $package.RevisionNumber, $currentBranch, $flightRing, $flightingBranchName
                $fileUrlResponse = Invoke-WebRequest -Uri "$baseUri/secured" -Method Post -Body $fileUrlRequestPayload -Headers $headers -UseBasicParsing
                $fileUrlResponseXml = [xml]$fileUrlResponse.Content

                $fileLocations = $fileUrlResponseXml.Envelope.Body.GetExtendedUpdateInfo2Response.GetExtendedUpdateInfo2Result.FileLocations.FileLocation
                $baseFileName = [System.IO.Path]::GetFileNameWithoutExtension($package.FileName)
                $downloadUrl = ($fileLocations | Where-Object { $_.Url -like "*$baseFileName*" }).Url

                if (-not $downloadUrl) {
                    Write-Warning "  -> Could not retrieve download URL for $($package.FileName). Skipping."
                    continue
                }
                if ($noDownload) {
                    Write-Host "  -> Skipping download for $($package.FullName) because of -noDownload switch." -ForegroundColor Yellow
                    continue
                }

                # Download the file
                # Construct a more descriptive filename using the package's full name and its original extension
                $fileExtension = [System.IO.Path]::GetExtension($package.FileName)
                $newFileName = "$($package.FullName)$($fileExtension)"
                $filePath = Join-Path $workingDir $newFileName
                
                Write-Host "  -> Downloading from: $downloadUrl" -ForegroundColor Gray
                Write-Host "  -> Saving to: $filePath"

                try {
                    Invoke-WebRequest -Uri $downloadUrl -OutFile $filePath -UseBasicParsing
                    Write-Host "  -> SUCCESS: Download complete." -ForegroundColor Green
                } catch {
                    Write-Error "  -> FAILED to download $($newFileName). Error: $($_.Exception.Message)"
                }
                Write-Host ""
            }
            
            $ProgressPreference = $originalPref

            Write-Host "------------------------------------------------------------"
            Write-Host "Finished downloading packages to: $workingDir" -ForegroundColor Green

        } catch {
            Write-Host "An error occurred during the filtering or downloading phase:" -ForegroundColor Red
            Write-Host $_.Exception.ToString()
        }
    }

    If ($noDownload) {
        Write-Host "Skipping download step because of -noDownload switch." -ForegroundColor Yellow
        return
    }
    If ($noInstall) {
        Write-Host "Skipping installation step because of -noInstall switch." -ForegroundColor Yellow
        return
    }
    
    # --- Step 5: Install Downloaded Packages ---
    Write-Host "------------------------------------------------------------"
    Write-Host "Step 5: Installing packages..." -ForegroundColor Magenta
    Write-Host "This step requires Administrator privileges."

    # 1. Check for Administrator rights
    if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Error "Installation failed. Please re-run the script with 'Run as Administrator'."
        # Add a pause so the user can see the message before the window closes.
        Read-Host "Press Enter to exit"
        exit
    }

    # 2. Define the installation order for dependencies based on their base names
    #    The order here is critical for dependencies.
    $dependencyInstallOrder = @(
        'Microsoft.VCLibs',
        'Microsoft.NET.Native.Framework',
        'Microsoft.NET.Native.Runtime',
        'Microsoft.UI.Xaml'
    )

    # 3. Get all downloaded package files and separate the main app from dependencies
    try {
        $allDownloadedFiles = Get-ChildItem -Path $workingDir -File | Where-Object { $_.Extension -in '.appx', '.msix', '.appxbundle', '.msixbundle' }
        
        $storePackageFile = $allDownloadedFiles | Where-Object { $_.Name -like 'Microsoft.WindowsStore*' } | Select-Object -First 1
        $dependencyFiles = $allDownloadedFiles | Where-Object { $_.Name -notlike 'Microsoft.WindowsStore*' }

        if (-not $dependencyFiles -and -not $storePackageFile) {
            Write-Warning "No package files found in '$workingDir' to install."
            return # Exits this part of the script gracefully
        }

        # 4. Install dependencies in the correct, predefined order
        Write-Host "Installing dependencies..."
        foreach ($baseName in $dependencyInstallOrder) {
            # Find all packages that start with the current base name (e.g., 'Microsoft.VCLibs*')
            # Sorting by name ensures a consistent order if multiple versions exist
            $packagesInGroup = $dependencyFiles | Where-Object { $_.Name -like "$baseName*" } | Sort-Object Name
            
            foreach ($package in $packagesInGroup) {
                Write-Host "  -> Installing $($package.Name)"
                try {
                    Add-AppxPackage -Path $package.FullName
                    Write-Host "     SUCCESS." -ForegroundColor Green
                } catch {
                    Write-Error "     FAILED to install $($package.Name). Error: $($_.Exception.Message)"
                }
            }
        }

        # 5. Install the main Microsoft Store package last
        if ($storePackageFile) {
            Write-Host "Installing the main application..."
            Write-Host "  -> Installing $($storePackageFile.Name)"
            try {
                Add-AppxPackage -Path $storePackageFile.FullName
                Write-Host "     SUCCESS: Microsoft Store has been installed/updated." -ForegroundColor Green
            } catch {
                Write-Error "     FAILED to install $($storePackageFile.Name). Error: $($_.Exception.Message)"
            }
        } else {
            Write-Warning "Microsoft Store package was not found in the download folder."
        }

        Write-Host "------------------------------------------------------------"
        Write-Host "Installation process finished." -ForegroundColor Cyan

    } catch {
        Write-Error "A critical error occurred during the installation phase: $($_.Exception.Message)"
    }
    
    # --- Set Region to US so the store will work. Default 'World' region does not work. ---
    try {
        # Define the path to the registry key
        $geoKeyPath = "HKCU:\Control Panel\International\Geo"

        # Check if the 'Geo' key exists. If not, create it.
        # The -Force switch ensures that parent keys ('International') are also created if they are missing.
        if (-not (Test-Path $geoKeyPath)) {
            Write-Host "  -> Registry key not found. Creating: $geoKeyPath"
            New-Item -Path $geoKeyPath -Force | Out-Null
        }

        # Set the 'Nation' value. This is equivalent to: reg add ... /v Nation ...
        Set-ItemProperty -Path $geoKeyPath -Name "Nation" -Value "244"
        Write-Host "  -> Set 'Nation' value to '244'."

        # Set the 'Name' value. This is equivalent to: reg add ... /v Name ...
        Set-ItemProperty -Path $geoKeyPath -Name "Name" -Value "US"
        Write-Host "  -> Set 'Name' value to 'US'."

        Write-Host "  -> Registry configuration complete." -ForegroundColor Green
    }
    catch {
        Write-Error "FAILED to configure registry settings. Error: $($_.Exception.Message)"
    }
    

} catch {
    Write-Host "An error occurred:" -ForegroundColor Red
    if ($_.Exception.Response) {
        $statusCode = $_.Exception.Response.StatusCode.value__
        $statusDescription = $_.Exception.Response.StatusDescription
        $errorLogPath = Join-Path $LogDirectory "ERROR_Response.txt"
        try {
            $stream = $_.Exception.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($stream)
            If ($debugSaveFiles) { $reader.ReadToEnd() | Set-Content -Path $errorLogPath }
        } catch { "Could not read error response body." | Set-Content -Path $errorLogPath }
        Write-Host "Status Code: $statusCode"
        Write-Host "Status Description: $statusDescription"
        Write-Host "Server Response saved to '$errorLogPath'"
    } else {
        Write-Host $_.Exception.ToString()
    }
}
