# This script downloads and installs any app or framework package that the Microsoft Store distributes, along with its dependencies.
#     - It does NOT require the Microsoft Store be installed first.
#     - It uses the same Windows Update API that the Microsoft Store uses, to fetch the files directly from Microsoft.
#     - It can also search the Microsoft Update Catalog, which lists the updates published through the regular Windows Update and Microsoft Update services.
#     - Unlike many similar scripts, it uses NO dependencies or third party APIs.

# Repo Url: https://github.com/ThioJoe/Windows-Sandbox-Tools
# Last Updated: October 19, 2026

param(
    # The app to fetch, given as its Store product ID. This is the 12-character ID in a Store link, such as "9WZDNCRFJBMP" for the Microsoft Store itself.
    #   - A full Store link that contains the ID also works, such as "https://apps.microsoft.com/detail/9WZDNCRFJBMP"
    #   - The app's Windows Update category ID and package name are looked up from the Store's display catalog.
    [string]$ProductId,

    # The app to fetch, given as its Windows Update category ID (a GUID) instead of a product ID. Use one or the other.
    [string]$CategoryId,

    # Optional package name of the main app, such as "Microsoft.WindowsStore". It's normally found automatically, so this is only needed if that fails.
    [string]$MainPackageName,

    # Optional wildcard filter for package names, such as "Microsoft.VCLibs*". Only matching packages are downloaded and installed.
    [string]$NameFilter,

    # Optional folder to save the downloaded files to. Defaults to a "<package name> Install" folder in the Downloads folder.
    [string]$DownloadPath,

    # Optional: Instead of fetching an app, search the Microsoft Update Catalog for this text. The results are numbered and shown 25 at a time,
    #   and you can choose which ones to get the download links for, such as 1,3 or 2-5. Enter N to load the next 25 results.
    #   - The catalog lists updates published through regular Windows Update and Microsoft Update, such as Windows updates, drivers, Office, .NET and Visual Studio.
    #   - Store apps are not listed there. Those are fetched with -ProductId or -CategoryId instead.
    [string]$SearchCatalog,

    # Optional switch to output the generated XML files to a debug folder in the Downloads folder
    [switch]$debugSaveFiles,

    # Optional switch to skip the installation, but still download the files
    [switch]$noInstall,

    # Optional switch to skip the download and install, but still show the packages found
    [switch]$noDownload,

    # Optional path to a local directory containing the installation files. If provided, the download steps will be skipped if the cached files are up to date.
    #   - Just copy the entire folder with the files the script normally downloads, and put it in your mounted folder to avoid having to re-download it.
    #        - By default the files are saved to a "<package name> Install" folder in the "Downloads" folder.
    #   - Make sure to use the mounted path from the perspective of within the sandbox.
    [string]$ExistingInstallerFilesPath,

    # If set, forces using the installer files from ExistingInstallerFilesPath even if a new version exists. Still warns about there being a new version.
    [switch]$ForceCachedFilesOnly,

    # If set, implies ForceCachedFilesOnly and uses cached installer files only, but does not even bother checking for a latest version.
    [switch]$NoCheckLatestVersion
)

# --- Parameter Usage Examples ---
# Interactive mode (type in a product ID or Store link, or search the Store or the Microsoft Update Catalog):
#    .\Install-Store-Apps.ps1
#
# Download & install an app by its Store product ID (this one is the Microsoft Store):
#    .\Install-Store-Apps.ps1 -ProductId "9WZDNCRFJBMP"
#
# Same, using a Store link:
#    .\Install-Store-Apps.ps1 -ProductId "https://apps.microsoft.com/detail/9N0DX20HK701"
#
# Download & install an app by its Windows Update category ID:
#    .\Install-Store-Apps.ps1 -CategoryId "64293252-5926-453c-9494-2d4021f1c78d"
#
# Only show which packages would be fetched (no download or install):
#    .\Install-Store-Apps.ps1 -ProductId "9WZDNCRFJBMP" -noDownload
#
# Download only the VCLibs framework packages an app depends on (Don't Install):
#    .\Install-Store-Apps.ps1 -ProductId "9WZDNCRFJBMP" -NameFilter "Microsoft.VCLibs*" -noInstall
#
# Install from existing files instead of downloading:
#    .\Install-Store-Apps.ps1 -ProductId "9WZDNCRFJBMP" -ExistingInstallerFilesPath "C:\Users\WDAGUtilityAccount\Desktop\HostShared\Microsoft.WindowsStore Install"
#
# Search the Microsoft Update Catalog:
#    .\Install-Store-Apps.ps1 -SearchCatalog "Visual C++ Redistributable"
#
# Debug mode (Save SOAP XML logs to disk):
#    .\Install-Store-Apps.ps1 -ProductId "9WZDNCRFJBMP" -debugSaveFiles

# =======================================================

# Hide PowerShell's progress bars for the whole script. Drawing them slows down downloads a lot, and makes app installs take about a third longer.
$ProgressPreference = 'SilentlyContinue'

# Validate parameter combinations
if ($NoCheckLatestVersion) {
    $ForceCachedFilesOnly = [switch]::new($true)
}
if ($ForceCachedFilesOnly -and [string]::IsNullOrWhiteSpace($ExistingInstallerFilesPath)) {
    Write-Host "Error: -ForceCachedFilesOnly requires -ExistingInstallerFilesPath to be specified." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}
if ($ProductId -and $CategoryId) {
    Write-Host "Error: Specify either -ProductId or -CategoryId, not both." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}
if ($CategoryId -and -not ($CategoryId -as [guid])) {
    Write-Host "Error: -CategoryId must be a GUID, such as 64293252-5926-453c-9494-2d4021f1c78d" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# --- Store Product Lookup ---
# Looks up a Store product's Windows Update category ID and package name from the Store's display catalog
#   Accepts a product ID, or a Store link that contains one. Throws an error if the product can't be found or has no category ID.
function Resolve-StoreProduct([string]$productIdOrLink) {
    # The product IDs of packaged Store apps are 12 characters long and start with 9. Desktop installer apps have 14-character IDs that start with XP.
    $id = $productIdOrLink.Trim()
    if ($id -notmatch '^[A-Za-z0-9]{12,14}$' -and $id -match '(?i)(?<![a-z0-9])(9[a-z0-9]{11}|XP[a-z0-9]{12})(?![a-z0-9])') {
        $id = $matches[1]
    }
    $id = $id.ToUpper()

    if ($id -like 'XP*') {
        throw "'$id' is a desktop installer app. The Store installs those with their own installers rather than through Windows Update, so this script can't fetch them."
    }

    try {
        $displayCatalogUri = "https://displaycatalog.mp.microsoft.com/v7.0/products/$($id)?market=US&languages=en-us,neutral&MS-CV=DGU1mcuYo0WMMp+F.1"
        $product = (Invoke-RestMethod -Uri $displayCatalogUri -UseBasicParsing -ErrorAction Stop).Product
    } catch {
        if ($_.Exception.Response -and [int]$_.Exception.Response.StatusCode -eq 404) {
            throw "No Store product has the ID '$id'."
        }
        throw
    }

    # Each SKU has its own fulfillment data. It's returned as an object, but parse it in case it ever comes back as a JSON string
    $fulfillmentData = $product.DisplaySkuAvailabilities |
        ForEach-Object { $_.Sku.Properties.FulfillmentData } |
        Where-Object { $_ } |
        ForEach-Object { if ($_ -is [string]) { $_ | ConvertFrom-Json } else { $_ } } |
        Where-Object { $_.WuCategoryId } |
        Select-Object -First 1

    if (-not $fulfillmentData) {
        throw "The display catalog has no Windows Update category ID for '$($product.LocalizedProperties[0].ProductTitle)' ($id). Only packaged (MSIX / APPX) Store apps have one."
    }

    [PSCustomObject]@{
        ProductId   = $id
        Title       = $product.LocalizedProperties[0].ProductTitle
        CategoryId  = $fulfillmentData.WuCategoryId
        PackageName = ($fulfillmentData.PackageFamilyName -split '_')[0]
    }
}

# --- Store Search ---
# Searches the Microsoft Store for apps by name. Returns the matching packaged apps (the kind this script can fetch), with their titles and product IDs.
#   Uses the Store search that winget uses for its 'msstore' source. If that can't be reached, it falls back to the display catalog's title suggestions, which only match the start of app names.
function Search-StoreApps([string]$searchText) {
    $results = @()
    try {
        # The Market filter is the same one winget sends with its Store searches
        $searchBody = @{
            Query   = @{ KeyWord = $searchText; MatchType = "Substring" }
            Filters = @( @{ PackageMatchField = "Market"; RequestMatch = @{ KeyWord = "US"; MatchType = "CaseInsensitive" } } )
        } | ConvertTo-Json -Compress -Depth 5
        $response = Invoke-RestMethod -Uri "https://storeedgefd.dsx.mp.microsoft.com/v9.0/manifestSearch" -Method Post -Body $searchBody -ContentType "application/json" -UseBasicParsing -ErrorAction Stop
        $results = @($response.Data | ForEach-Object { [PSCustomObject]@{ Title = $_.PackageName; Publisher = $_.Publisher; ProductId = $_.PackageIdentifier } })
    } catch {
        Write-Host "  (Could not reach the Store search, so using the display catalog's suggestions instead, which only match the start of app names.)" -ForegroundColor Yellow
        try {
            $suggestUri = "https://displaycatalog.mp.microsoft.com/v7.0/productFamilies/autosuggest?market=US&languages=en-US&query=$([uri]::EscapeDataString($searchText))&productFamilyNames=apps"
            $response = Invoke-RestMethod -Uri $suggestUri -UseBasicParsing -ErrorAction Stop
            $results = @($response.Results | ForEach-Object { $_.Products } | Where-Object { $_.ProductId } | ForEach-Object { [PSCustomObject]@{ Title = $_.Title; Publisher = $null; ProductId = $_.ProductId } })
        } catch {
            Write-Host "Error: Could not search the Store. $($_.Exception.Message)" -ForegroundColor Red
            return
        }
    }

    # Only packaged apps (product IDs starting with 9) are distributed through Windows Update. The others, such as desktop installer apps (IDs starting with XP), can't be fetched.
    $packagedResults = @($results | Where-Object { $_.ProductId -match '^9[A-Za-z0-9]{11}$' })
    $skippedCount = $results.Count - $packagedResults.Count
    if ($skippedCount -gt 0) {
        Write-Host "  ($skippedCount result(s) were desktop installer apps, which aren't distributed through Windows Update, so they're left out.)" -ForegroundColor DarkGray
    }
    return $packagedResults
}

# --- Reading Answers ---
# Reads a line of input from the user and returns it without surrounding spaces.
#   - Entering Q goes back to the main menu. This throws an exception that the main menu's loop catches, so that every prompt doesn't have to check for it.
#     When there's no main menu to go back to, such as with -SearchCatalog, the exception is caught where the search is started and the script exits instead.
#   - With -AllowQ, Q is returned like any other answer instead. The main menu uses this so it can quit the script.
#   - Exits the script if there is no more input, such as when the input was redirected from an empty file.
function Read-Answer([string]$prompt, [switch]$AllowQ) {
    $answer = Read-Host $prompt
    if ($null -eq $answer) { exit 1 }
    $answer = $answer.Trim()
    if (-not $AllowQ -and $answer -eq 'q') {
        throw [System.OperationCanceledException]::new("Q was entered to go back to the main menu.")
    }
    return $answer
}

# --- Microsoft Update Catalog Search ---
# The Microsoft Update Catalog is the title search for everything published through regular Windows Update and Microsoft Update.
# This is separate from the Store's channel used by the rest of the script, which can only be queried by category ID, not by name.

# Gets one page of search results from the Microsoft Update Catalog. Each page has up to 25 results, and the page index starts at 0.
#   Returns $null if the catalog couldn't be reached.
function Get-CatalogSearchPage([string]$searchText, [int]$pageIndex) {
    # The 'p' parameter is the page index, the same one the Next link on the catalog website uses
    $catalogSearchUri = "https://www.catalog.update.microsoft.com/Search.aspx?q=$([uri]::EscapeDataString($searchText))&p=$pageIndex"
    try {
        $catalogHtml = (Invoke-WebRequest -Uri $catalogSearchUri -UseBasicParsing -ErrorAction Stop).Content
    } catch {
        Write-Host "Error: Could not reach the Microsoft Update Catalog. $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }

    # The count text looks like "26 - 50 of 120 (page 2 of 5)"
    $countText = [regex]::Match($catalogHtml, '(?s)id="ctl00_catalogBody_searchDuration"[^>]*>(.*?)</span>').Groups[1].Value.Trim()
    $totalPages = 0
    if ($countText -match '\(page \d+ of (\d+)\)') {
        $totalPages = [int]$matches[1]
    }

    # The results are the rows of the table with the ID ctl00_catalogBody_updateMatches
    #   Cells: icon | Title | Products | Classification | Last Updated | Version | Size | Download button (its ID is the update's GUID)
    $resultsTable = [regex]::Match($catalogHtml, '(?s)<table[^>]*id="ctl00_catalogBody_updateMatches"[^>]*>(.*?)</table>').Groups[1].Value
    $results = @(foreach ($row in [regex]::Matches($resultsTable, '(?s)<tr[^>]*>(.*?)</tr>')) {
        if ($row.Value -match 'id="headerRow"') { continue }
        $cells = @([regex]::Matches($row.Groups[1].Value, '(?s)<td[^>]*>(.*?)</td>') | ForEach-Object {
            ([System.Net.WebUtility]::HtmlDecode(($_.Groups[1].Value -replace '<[^>]+>', ' ')) -replace '\s+', ' ').Trim()
        })
        if ($cells.Count -lt 8) { continue }
        [PSCustomObject]@{
            Title          = $cells[1]
            Products       = $cells[2]
            Classification = $cells[3]
            LastUpdated    = $cells[4]
            Size           = (($cells[6] -split ' ') | Select-Object -First 2) -join ' '
            UpdateId       = [regex]::Match($row.Groups[1].Value, '<input[^>]*id="([0-9a-fA-F-]{36})"').Groups[1].Value
        }
    })

    [PSCustomObject]@{
        Results    = $results
        CountText  = $countText
        TotalPages = $totalPages
    }
}

# Gets the download links for one catalog update, the same way the Download button on the catalog website does
function Get-CatalogDownloadLinks([string]$updateId) {
    $dialogBody = @{ updateIDs = "[{`"size`":0,`"updateID`":`"$updateId`",`"uidInfo`":`"$updateId`"}]" }
    $dialogHtml = (Invoke-WebRequest -Uri "https://www.catalog.update.microsoft.com/DownloadDialog.aspx" -Method Post -Body $dialogBody -UseBasicParsing -ErrorAction Stop).Content
    [regex]::Matches($dialogHtml, "https?://(?:[a-z0-9.-]+\.)?(?:download\.windowsupdate\.com|delivery\.mp\.microsoft\.com)/[^'`"\s]+") | ForEach-Object { $_.Value } | Select-Object -Unique
}

# Turns a selection of result numbers into a sorted list of the numbers, such as "1,3" or "2-5" or "1, 4 - 6"
#   Throws an error that describes the problem if the selection isn't valid, or has a number that isn't between 1 and $maxNumber.
function ConvertFrom-Selection([string]$selectionText, [int]$maxNumber) {
    $numbers = [System.Collections.Generic.List[int]]::new()
    foreach ($part in ($selectionText -split ',')) {
        $part = $part.Trim()
        if (-not $part) { continue }
        if ($part -match '^(\d{1,6})$') {
            $start = [int]$matches[1]
            $end = $start
        } elseif ($part -match '^(\d{1,6})\s*-\s*(\d{1,6})$') {
            $start = [int]$matches[1]
            $end = [int]$matches[2]
            if ($start -gt $end) { $start, $end = $end, $start }
        } else {
            throw "'$part' isn't a number or a range of numbers. Enter numbers like 1,3 or a range like 2-5."
        }
        if ($start -lt 1 -or $end -gt $maxNumber) {
            throw "'$part' is outside the results shown so far, which are 1 to $maxNumber."
        }
        foreach ($number in $start..$end) {
            if (-not $numbers.Contains($number)) { $numbers.Add($number) }
        }
    }
    if ($numbers.Count -eq 0) {
        throw "Enter numbers like 1,3 or a range like 2-5."
    }
    $numbers.Sort()
    return ,$numbers.ToArray()
}

# Shows catalog results with their numbers. The numbers continue from page to page, so the second page starts at 26.
function Show-CatalogResults($results, [int]$firstNumber) {
    for ($i = 0; $i -lt $results.Count; $i++) {
        $result = $results[$i]
        $products = $result.Products
        if ($products.Length -gt 70) { $products = $products.Substring(0, 67) + "..." }
        Write-Host ("{0,5}) {1}" -f ($firstNumber + $i), $result.Title)
        Write-Host ("       {0} | {1} | {2} | {3}" -f $products, $result.Classification, $result.LastUpdated, $result.Size) -ForegroundColor DarkGray
    }
}

# Searches the Microsoft Update Catalog and shows the results 25 at a time.
#   The user can load the next page, and get the download links for the results they choose by number, from any of the pages shown so far.
#   Returns $false if the catalog couldn't be reached.
function Invoke-CatalogSearch([string]$searchText) {
    Write-Host "Searching the Microsoft Update Catalog for '$searchText'..."
    $page = Get-CatalogSearchPage $searchText 0
    if (-not $page) { return $false }
    if ($page.Results.Count -eq 0) {
        Write-Host "The catalog has no updates matching '$searchText'." -ForegroundColor Yellow
        return $true
    }

    $allResults = [System.Collections.Generic.List[object]]::new()
    $pageIndex = 0
    $totalPages = $page.TotalPages

    while ($true) {
        # Show the page that was just loaded
        if ($page) {
            Write-Host ""
            Show-CatalogResults $page.Results ($allResults.Count + 1)
            $allResults.AddRange([object[]]$page.Results)
            Write-Host ""
            Write-Host "Results $($page.CountText)" -ForegroundColor Gray
            $page = $null
        }

        $morePages = ($pageIndex + 1) -lt $totalPages
        Write-Host ""
        Write-Host "Enter result numbers to get their download links, such as 1,3 or 2-5." -ForegroundColor Gray
        if ($morePages) {
            Write-Host "Enter N to load the next page of results." -ForegroundColor Gray
        }
        Write-Host "Press Enter without typing anything when you're done." -ForegroundColor Gray
        Write-Host ""
        $selection = Read-Answer "Enter selection"

        if (-not $selection) {
            return $true
        }

        if ($selection -eq 'n') {
            if (-not $morePages) {
                Write-Host "That was the last page of results." -ForegroundColor Yellow
                continue
            }
            $nextPage = Get-CatalogSearchPage $searchText ($pageIndex + 1)
            if ($nextPage -and $nextPage.Results.Count -gt 0) {
                $pageIndex++
                $page = $nextPage
            } else {
                Write-Host "Could not load the next page." -ForegroundColor Yellow
            }
            continue
        }

        try {
            $numbers = ConvertFrom-Selection $selection $allResults.Count
        } catch {
            Write-Host $_.Exception.Message -ForegroundColor Yellow
            continue
        }

        Write-Host ""
        foreach ($number in $numbers) {
            $result = $allResults[$number - 1]
            Write-Host ("{0,5}) {1}" -f $number, $result.Title) -ForegroundColor Green
            try {
                $downloadUrls = @(Get-CatalogDownloadLinks $result.UpdateId)
                if ($downloadUrls.Count -eq 0) {
                    Write-Host "       No download links were found for this update." -ForegroundColor Yellow
                }
                foreach ($downloadUrl in $downloadUrls) { Write-Host "       $downloadUrl" -ForegroundColor Gray }
            } catch {
                Write-Host "       Could not get the download links. Error: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
    }
}

# --- Catalog Search Mode ---
if ($SearchCatalog) {
    try {
        $catalogReached = Invoke-CatalogSearch $SearchCatalog
    } catch [System.OperationCanceledException] {
        # Q was entered. There's no main menu in this mode, so just exit
        $catalogReached = $true
    }
    if ($catalogReached) {
        exit 0
    }
    Read-Host "Press Enter to exit"
    exit 1
}

# --- Interactive Mode ---
# If the script was run without an app to fetch or a search to do, ask what to do

# Shows the results of a Store search with their numbers
function Show-StoreResults($shownResults, [int]$totalCount) {
    for ($i = 0; $i -lt $shownResults.Count; $i++) {
        $publisherText = if ($shownResults[$i].Publisher) { " - $($shownResults[$i].Publisher)" } else { "" }
        Write-Host ("  {0,2}) {1}{2}  [{3}]" -f ($i + 1), $shownResults[$i].Title, $publisherText, $shownResults[$i].ProductId)
    }
    if ($totalCount -gt $shownResults.Count) {
        Write-Host "  (Showing the first $($shownResults.Count) of $totalCount results. Search for something more specific to narrow them down.)" -ForegroundColor DarkGray
    }
    Write-Host ""
}

# Looks up the app that was entered. A category ID is used as it is, and anything else is looked up as a Store product ID or link.
#   Returns an object describing the app, or nothing if it couldn't be found.
function Find-EnteredApp([string]$appEntry) {
    if ($appEntry -as [guid]) {
        return [PSCustomObject]@{ Product = $null; CategoryId = $appEntry; Description = "the app with category ID $appEntry" }
    }
    Write-Host "Looking up Store product '$appEntry'..."
    try {
        $product = Resolve-StoreProduct $appEntry
    } catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        return
    }
    return [PSCustomObject]@{ Product = $product; CategoryId = $product.CategoryId; Description = "'$($product.Title)'" }
}

# Asks what to do with the chosen app. Returns 'install', 'download' or 'list', or nothing if Enter was pressed to go back.
function Read-AppAction([string]$appDescription) {
    Write-Host ""
    Write-Host "What should be done with $($appDescription)?" -ForegroundColor Cyan
    Write-Host "  1) Download and install it"
    Write-Host "  2) Only download it"
    Write-Host "  3) Only list the packages that would be downloaded"
    Write-Host ""
    while ($true) {
        $actionChoice = Read-Answer "Enter a choice"
        if (-not $actionChoice) { return }
        if ($actionChoice -eq '1') { return 'install' }
        if ($actionChoice -eq '2') { return 'download' }
        if ($actionChoice -eq '3') { return 'list' }
        Write-Host "'$actionChoice' isn't one of the choices. Enter 1, 2 or 3, or press Enter to go back." -ForegroundColor Yellow
    }
}

# Shows examples of what to enter for each option in interactive mode, and what each option does
function Show-Examples {
    $appExampleFormat = "      {0,-48} {1}"
    $catalogExampleFormat = "      {0,-40} {1}"

    Write-Host ""
    Write-Host "1) Fetch an app by its Store product ID or Store link" -ForegroundColor Cyan
    Write-Host "   Downloads an app from the Microsoft Store along with the framework packages it needs (.appx, .msix and"
    Write-Host "   .msixbundle files), and installs them if you choose to."
    Write-Host "   The product ID is the 12-character code after /detail/ in the app's Store link. To get the link, find the app"
    Write-Host "   on apps.microsoft.com, or use the Share button on the app's page in the Store app."
    Write-Host "   Examples:"
    Write-Host ($appExampleFormat -f "9WZDNCRFJBMP", "Microsoft Store") -ForegroundColor Green
    Write-Host ($appExampleFormat -f "9N0DX20HK701", "Windows Terminal") -ForegroundColor Green
    Write-Host ($appExampleFormat -f "9WZDNCRFHVN5", "Windows Calculator") -ForegroundColor Green
    Write-Host ($appExampleFormat -f "9MZ95KL8MR0L", "Snipping Tool") -ForegroundColor Green
    Write-Host ($appExampleFormat -f "9NBLGGH3FRZM", "VCLibs, a C++ runtime framework that many Store apps need") -ForegroundColor Green
    Write-Host ($appExampleFormat -f "https://apps.microsoft.com/detail/9N0DX20HK701", "A Store link works too") -ForegroundColor Green
    Write-Host "   Only apps that the Store installs as packages can be fetched. Their IDs start with 9. Apps with IDs that start" -ForegroundColor Gray
    Write-Host "   with XP are installed by their own installers instead, so they can't be fetched." -ForegroundColor Gray

    Write-Host ""
    Write-Host "2) Search the Microsoft Store for an app" -ForegroundColor Cyan
    Write-Host "   Finds Store apps by name, so you can pick one from a list instead of looking up its product ID."
    Write-Host "   Examples:"
    Write-Host "      Windows Terminal" -ForegroundColor Green
    Write-Host "      Calculator" -ForegroundColor Green
    Write-Host "      Snipping Tool" -ForegroundColor Green
    Write-Host "      WhatsApp" -ForegroundColor Green

    Write-Host ""
    Write-Host "3) Search the Microsoft Update Catalog" -ForegroundColor Cyan
    Write-Host "   The Microsoft Update Catalog (catalog.update.microsoft.com) is Microsoft's list of the updates delivered through"
    Write-Host "   Windows Update, Microsoft Update and WSUS. It has Windows updates, drivers, and updates for other Microsoft products"
    Write-Host "   such as .NET, Office, Visual Studio, Microsoft Edge and Microsoft Defender. Store apps aren't in it."
    Write-Host "   The search matches words in the update titles, so KB numbers, product names and driver names work best."
    Write-Host "   The results are numbered and shown 25 at a time. Enter N to load the next 25, or enter result numbers (such as"
    Write-Host "   1,3 or 2-5) to see the download links for those updates. The files are mostly .msu and .cab files for Windows"
    Write-Host "   updates, drivers and Office, and .exe files for things like .NET, Visual Studio, Edge and Defender. This script"
    Write-Host "   only shows the links. It doesn't download or install the files."
    Write-Host "   Examples:"
    Write-Host ($catalogExampleFormat -f "KB5129195", "One update by its KB number (a September 2026 Windows 11 update)") -ForegroundColor Green
    Write-Host ($catalogExampleFormat -f "Cumulative Update Windows 11 24H2 x64", "Monthly updates for Windows 11 version 24H2 on 64-bit PCs") -ForegroundColor Green
    Write-Host ($catalogExampleFormat -f ".NET 8.0", ".NET 8.0 runtime updates") -ForegroundColor Green
    Write-Host ($catalogExampleFormat -f "Visual Studio 2022", "Visual Studio 2022 updates") -ForegroundColor Green
    Write-Host ($catalogExampleFormat -f "Microsoft Defender Antivirus", "Defender antimalware platform updates") -ForegroundColor Green
    Write-Host ($catalogExampleFormat -f "Intel Bluetooth", "Intel Bluetooth driver updates") -ForegroundColor Green
    Write-Host ($catalogExampleFormat -f "Visual C++ Redistributable", "Visual C++ 2005-2010 runtime updates (newer versions aren't listed)") -ForegroundColor Green

    Write-Host ""
    Read-Answer "Press Enter to go back to the menu" | Out-Null
}

$interactiveMode = $false
$resolvedProduct = $null
if (-not $ProductId -and -not $CategoryId -and -not $NoCheckLatestVersion) {
    $interactiveMode = $true
    Write-Host "================ Windows App & Package Fetcher ================" -ForegroundColor Gray
    Write-Host "=============        Thio's Sandbox Tools         =============" -ForegroundColor Gray

    # Each menu and prompt below is a loop. Pressing Enter without typing anything leaves that loop, which goes back one step.
    # Q goes back to the main menu from any prompt (see Read-Answer). Once an app and what to do with it are chosen, 'break mainMenu' leaves all the loops.
    $chosenApp = $null
    $chosenAction = $null
    $tipShown = $false
    :mainMenu while ($true) {
        try {
            Write-Host ""
            Write-Host "What would you like to do?" -ForegroundColor Cyan
            Write-Host "  1) Fetch an app by its Store product ID or Store link"
            Write-Host "  2) Search the Microsoft Store for an app"
            Write-Host "  3) Search the Microsoft Update Catalog"
            Write-Host ""
            Write-Host "  E) See examples of what to enter for each option"
            Write-Host "  Q) Quit (at any other prompt, Q goes back to this menu)"
            Write-Host ""
            # Only shown the first time the menu comes up
            if (-not $tipShown) {
                Write-Host "  Tip: Pressing Enter without typing anything goes back one step." -ForegroundColor Gray
                Write-Host ""
                $tipShown = $true
            }
            $choice = Read-Answer "Enter a choice (or paste a product ID or Store link)" -AllowQ
            Write-Host ""

            if ($choice -eq '1') {
                # Ask for an app until one is found and an action is chosen
                $entryPrompt = "Enter a Store product ID, a Store link, or a Windows Update category ID"
                while ($true) {
                    $appEntry = Read-Answer $entryPrompt
                    if (-not $appEntry) { break }
                    $app = Find-EnteredApp $appEntry
                    if (-not $app) {
                        $entryPrompt = "Enter a different product ID, Store link or category ID"
                        continue
                    }
                    $action = Read-AppAction $app.Description
                    if ($action) {
                        $chosenApp = $app
                        $chosenAction = $action
                        break mainMenu
                    }
                    # Enter was pressed at the action menu, so ask for an app again
                    Write-Host ""
                    $entryPrompt = "Enter a Store product ID, a Store link, or a Windows Update category ID"
                }
            } elseif ($choice -eq '2') {
                :storeSearch while ($true) {
                    $searchText = Read-Answer "Search the Store for"
                    if (-not $searchText) { break }
                    $storeResults = @(Search-StoreApps $searchText)
                    if ($storeResults.Count -eq 0) {
                        Write-Host "No apps that can be fetched were found for '$searchText'." -ForegroundColor Yellow
                        continue
                    }
                    $shownResults = @($storeResults | Select-Object -First 25)

                    # Show the results again whenever Enter is pressed at the action menu
                    while ($true) {
                        Show-StoreResults $shownResults $storeResults.Count
                        while ($true) {
                            $pick = Read-Answer "Enter a number to fetch that app, or just press Enter to go back"
                            if (-not $pick) {
                                Write-Host ""
                                continue storeSearch
                            }
                            if ($pick -match '^\d{1,4}$' -and [int]$pick -ge 1 -and [int]$pick -le $shownResults.Count) { break }
                            Write-Host "Enter a number from 1 to $($shownResults.Count), or Q to go back to the main menu." -ForegroundColor Yellow
                        }
                        $app = Find-EnteredApp $shownResults[[int]$pick - 1].ProductId
                        if ($app) {
                            $action = Read-AppAction $app.Description
                            if ($action) {
                                $chosenApp = $app
                                $chosenAction = $action
                                break mainMenu
                            }
                        }
                        Write-Host ""
                    }
                }
            } elseif ($choice -eq '3') {
                while ($true) {
                    $searchText = Read-Answer "Search the Microsoft Update Catalog for"
                    if (-not $searchText) { break }
                    Invoke-CatalogSearch $searchText | Out-Null
                    Write-Host ""
                }
            } elseif ($choice -in 'e', 'examples') {
                Show-Examples
            } elseif ($choice -in 'q', 'quit', 'exit') {
                exit 0
            } elseif (($choice -as [guid]) -or ($choice -match '(?i)(?<![a-z0-9])(9[a-z0-9]{11}|XP[a-z0-9]{12})(?![a-z0-9])')) {
                # A product ID, Store link or category ID was pasted in directly. Pressing Enter at the next prompts goes back to the main menu.
                $appEntry = $choice
                while ($appEntry) {
                    $app = Find-EnteredApp $appEntry
                    if ($app) {
                        $action = Read-AppAction $app.Description
                        if ($action) {
                            $chosenApp = $app
                            $chosenAction = $action
                            break mainMenu
                        }
                        break
                    }
                    $appEntry = Read-Answer "Enter a different product ID, Store link or category ID"
                }
            } elseif ($choice) {
                Write-Host "'$choice' isn't one of the choices." -ForegroundColor Yellow
            }
        } catch [System.OperationCanceledException] {
            # Q was entered at a prompt, so go back to the main menu
        }
    }

    # Use the app and the action that were chosen
    if ($chosenApp.Product) {
        $resolvedProduct = $chosenApp.Product
        $ProductId = $resolvedProduct.ProductId
    } else {
        $CategoryId = $chosenApp.CategoryId
    }
    if ($chosenAction -eq 'download') {
        $noInstall = [switch]::new($true)
    } elseif ($chosenAction -eq 'list') {
        $noDownload = [switch]::new($true)
    }
    Write-Host ""
}

# --- Configuration ---
# Note: These defaults are the values for regular Store releases. Fetching insider builds of an app (if any) might need other values.
$flightRing = "Retail"             # Apparently accepts 'Retail', 'Internal', and 'External'
$flightingBranchName = ""          # Empty ( "" ) for normal release. Otherwise apparent possible values: Dev, Beta, ReleasePreview, MSIT, CanaryChannel, external

$osBuildNumber = [Environment]::OSVersion.Version.Build
if ($osBuildNumber -ge 26100) {
    $currentBranch = "ge_release"
} elseif ($osBuildNumber -ge 22000) {
    $currentBranch = "ni_release"
} else {
    $currentBranch = "vb_release"
}

# The Windows version, used to skip package versions that need a newer version of Windows than this one
$osVersion = [version]"10.0.$osBuildNumber.0"

# Random Notes:
#   flightRing should be "Internal" if flightingBranchName is "MSIT"
#   MAYBE need flightRing as "External" if setting flightingBranchName anything besides empty or MSIT?

# ------ Check that we're running in the Windows Sandbox ------
# This script is intended to be run from within the Windows Sandbox. Warn the user if not, unless nothing will be installed. We'll do a rudamentary check for if the current user is named "WDAGUtilityAccount"
if (-not $noInstall -and -not $noDownload -and $env:USERNAME -ne "WDAGUtilityAccount") {
    Write-Warning "`nThis script is intended to be run from WITHIN the Windows Sandbox.`nRunning it outside the Sandbox will install the packages to the current system."
    Write-host "`nPress Enter to continue installing to the current system anyway, or just close this window to exit. (Use -noInstall to only download the files.)" -ForegroundColor Yellow
    Read-Host
}

# --- Define Working Directory ---
# Get the path to the user's personal Downloads folder in a reliable way
try {
    $userDownloadsFolder = (New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path
} catch {
    $userDownloadsFolder = Join-Path -Path $HOME -ChildPath "Downloads"
}

# Folder for the debug files from -debugSaveFiles
$debugDir = Join-Path -Path $userDownloadsFolder -ChildPath "WU Package Fetcher Debug"
$pauseAtEnd = $false

# In interactive mode, always pause at the end so the window doesn't close before the results can be read
if ($interactiveMode) {
    $pauseAtEnd = $true
}

if ($ExistingInstallerFilesPath) {
    if (-not (Test-Path -Path $ExistingInstallerFilesPath)) {
        Write-Error "The specified local source path does not exist: $ExistingInstallerFilesPath"
        Read-Host "Press Enter to exit"
        return
    }
}

If ($debugSaveFiles) {
    if (-not (Test-Path -Path $debugDir)) {
        New-Item -Path $debugDir -ItemType Directory -Force | Out-Null
    }
    Write-Host "Debug files will be saved to: '$debugDir'" -ForegroundColor Yellow
}

# --- Find the App to Fetch ---
$targetCategoryId = $CategoryId
$mainPackageName = $MainPackageName   # If this is still empty, it's found from the server's response later

if ($ProductId -and -not $NoCheckLatestVersion) {
    # Look up the app's Windows Update category ID and package name from the Store's display catalog, unless interactive mode already did
    if (-not $resolvedProduct) {
        Write-Host "Looking up Store product '$ProductId'..."
        try {
            $resolvedProduct = Resolve-StoreProduct $ProductId
        } catch {
            if ($ExistingInstallerFilesPath) {
                # Can't check for a newer version, but the cached files can still be installed
                Write-Host "WARNING: Could not look up the product, so the cached files will be used without checking for a newer version." -ForegroundColor Yellow
                Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Yellow
                $NoCheckLatestVersion = [switch]::new($true)
                $pauseAtEnd = $true
            } else {
                Write-Host "Error: Could not look up Store product '$ProductId'. $($_.Exception.Message)" -ForegroundColor Red
                Read-Host "Press Enter to exit"
                exit 1
            }
        }
    }
    if ($resolvedProduct) {
        $targetCategoryId = $resolvedProduct.CategoryId
        if (-not $mainPackageName) {
            $mainPackageName = $resolvedProduct.PackageName
        }
        Write-Host "  -> Found '$($resolvedProduct.Title)': package '$($resolvedProduct.PackageName)', category ID $targetCategoryId" -ForegroundColor Green
    }
}

# Set up the download directory. If the main package name isn't known yet, this is set after it's found in the server's response
if ($DownloadPath) {
    $workingDir = $DownloadPath
} elseif ($mainPackageName) {
    $workingDir = Join-Path -Path $userDownloadsFolder -ChildPath "$mainPackageName Install"
} else {
    $workingDir = $null
}

# --- Package Compatibility Check ---
# Checks the target platforms in a package's applicability data to see if the package can be installed on this version of Windows
#   - Each target platform has a platform ID and a minimum Windows version. Desktop Windows can install packages that target 0 (Windows.Universal) or 3 (Windows.Desktop).
#   - Other IDs are for other devices, such as 5 for Windows Mobile, which is what the package versions ending in .70 are for.
#   - The minimum version is a 64-bit number with each of the 4 version parts packed into 16 bits.
function Test-PackageSupportsThisWindows($applicabilityBlob) {
    if (-not $applicabilityBlob) { return $true }
    try { $blob = $applicabilityBlob | ConvertFrom-Json } catch { return $true }

    $targetPlatforms = $blob.'content.targetPlatforms'
    if (-not $targetPlatforms) { return $true }

    foreach ($targetPlatform in $targetPlatforms) {
        if ($targetPlatform.'platform.target' -notin 0, 3) { continue }
        $packedVersion = [uint64]$targetPlatform.'platform.minVersion'
        $minVersion = [version]::new([int]($packedVersion -shr 48), [int](($packedVersion -shr 32) -band 0xFFFF), [int](($packedVersion -shr 16) -band 0xFFFF), [int]($packedVersion -band 0xFFFF))
        if ($minVersion -le $osVersion) { return $true }
    }
    return $false
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
#   The category ID in FilterAppCategoryIds selects which app is returned, along with its dependencies.
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
                    <DeviceAttributes>E:BranchReadinessLevel=CB&amp;CurrentBranch={2}&amp;OEMModel=Virtual%20Machine&amp;FlightRing={3}&amp;AttrDataVer=321&amp;InstallLanguage=en-US&amp;OSUILocale=en-US&amp;InstallationType=Client&amp;FlightingBranchName={4}&amp;OSSkuId=48&amp;App=WU_STORE&amp;ProcessorManufacturer=GenuineIntel&amp;OEMName_Uncleaned=Microsoft%20Corporation&amp;AppVer=1407.2503.28012.0&amp;OSArchitecture=AMD64&amp;IsFlightingEnabled=1&amp;TelemetryLevel=1&amp;DefaultUserRegion=39070&amp;WuClientVer=1310.2503.26012.0&amp;OSVersion=10.0.$osBuildNumber.0&amp;DeviceFamily=Windows.Desktop</DeviceAttributes>
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
                <u:Created>{6}</u:Created>
                <u:Expires>{7}</u:Expires>
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
            <deviceAttributes>E:BranchReadinessLevel=CB&amp;CurrentBranch={3}&amp;OEMModel=Virtual%20Machine&amp;FlightRing={4}&amp;AttrDataVer=321&amp;InstallLanguage=en-US&amp;OSUILocale=en-US&amp;InstallationType=Client&amp;FlightingBranchName={5}&amp;OSSkuId=48&amp;App=WU_STORE&amp;ProcessorManufacturer=GenuineIntel&amp;OEMName_Uncleaned=Microsoft%20Corporation&amp;AppVer=1407.2503.28012.0&amp;OSArchitecture=AMD64&amp;IsFlightingEnabled=1&amp;TelemetryLevel=1&amp;DefaultUserRegion=39070&amp;WuClientVer=1310.2503.26012.0&amp;OSVersion=10.0.$osBuildNumber.0&amp;DeviceFamily=Windows.Desktop</deviceAttributes>
        </GetExtendedUpdateInfo2>
    </s:Body>
</s:Envelope>
"@

# --- Script Execution ---
$headers = @{ "Content-Type" = "application/soap+xml; charset=utf-8" }
$baseUri = "https://fe3.delivery.mp.microsoft.com/ClientWebService/client.asmx"

# If cached files were provided, first check that the update server can be reached. If it can't (for example, when networking is disabled for the sandbox), use the cached files without checking for a newer version.
if ($ExistingInstallerFilesPath -and -not $NoCheckLatestVersion) {
    try {
        Invoke-WebRequest -Uri $baseUri -Method Post -Body $cookieXmlTemplate -Headers $headers -UseBasicParsing -ErrorAction Stop | Out-Null
    } catch {
        Write-Host "WARNING: Could not reach the update server, so the cached files will be used without checking for a newer version." -ForegroundColor Yellow
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Yellow
        $NoCheckLatestVersion = [switch]::new($true)
        $pauseAtEnd = $true
    }
}

try {
    if (-not $NoCheckLatestVersion) {
        # Step 1: Get Cookie
        Write-Host "Step 1: Getting authentication cookie..."
        $cookieRequestPayload = $cookieXmlTemplate
        If ($debugSaveFiles) { $cookieRequestPayload | Set-Content -Path (Join-Path $debugDir "01_Step1_Request.xml") }

        $cookieResponse = Invoke-WebRequest -Uri $baseUri -Method Post -Body $cookieRequestPayload -Headers $headers -UseBasicParsing
        If ($debugSaveFiles) { $cookieResponse.Content | Set-Content -Path (Join-Path $debugDir "01_Step1_Response.xml"); Write-Host "  -> Saved request and response logs for Step 1." }

        $cookieResponseXml = [xml]$cookieResponse.Content
        $encryptedCookieData = $cookieResponseXml.Envelope.Body.GetCookieResponse.GetCookieResult.EncryptedData
        Write-Host "Success. Cookie received." -ForegroundColor Green

        # Step 2: Get File List
        Write-Host "Step 2: Getting file list..."
        $fileListRequestPayload = $fileListXmlTemplate -f $encryptedCookieData, $targetCategoryId, $currentBranch, $flightRing, $flightingBranchName
        If ($debugSaveFiles) { [System.IO.File]::WriteAllText((Join-Path $debugDir "02_Step2_Request_AUTOMATED.xml"), $fileListRequestPayload, [System.Text.UTF8Encoding]::new($false)) }

        $fileListResponse = Invoke-WebRequest -Uri $baseUri -Method Post -Body $fileListRequestPayload -Headers $headers -UseBasicParsing
        If ($debugSaveFiles) { $fileListResponse.Content | Set-Content -Path (Join-Path $debugDir "02_Step2_Response_SUCCESS.xml") }

        # The response contains XML fragments that are HTML-encoded. We must decode this before treating it as XML.
        $decodedContent = [System.Net.WebUtility]::HtmlDecode($fileListResponse.Content)
        $fileListResponseXml = [xml]$decodedContent
        Write-Host "Successfully received and DECODED Step 2 response." -ForegroundColor Green

        $fileIdentityMap = @{}

        # Get the two main lists of updates from the now correctly-decoded response
        $newUpdates = $fileListResponseXml.Envelope.Body.SyncUpdatesResponse.SyncUpdatesResult.NewUpdates.UpdateInfo
        $allExtendedUpdates = $fileListResponseXml.Envelope.Body.SyncUpdatesResponse.SyncUpdatesResult.ExtendedUpdateInfo.Updates.Update

        # If the main package name isn't known yet, get it from the app's own category entry in the response, which has the app's package family name
        if (-not $mainPackageName) {
            $categoryUpdate = $newUpdates | Where-Object { $_.Xml.UpdateIdentity.UpdateID -eq $targetCategoryId } | Select-Object -First 1
            $mainPackageName = $categoryUpdate.Xml.ApplicabilityRules.Metadata.AppxPackageMetadata.AppxFamilyMetadata.Name
            if ($mainPackageName) {
                Write-Host "The main app's package name is '$mainPackageName'." -ForegroundColor Green
            } else {
                Write-Warning "Could not tell which package is the main app, so all packages will be treated as dependencies. Use -MainPackageName to specify it."
            }
        }
        if (-not $workingDir) {
            $folderName = if ($mainPackageName) { $mainPackageName } else { $targetCategoryId }
            $workingDir = Join-Path -Path $userDownloadsFolder -ChildPath "$folderName Install"
        }

        Write-Host "--- Correlating Update Information ---" -ForegroundColor Cyan

        # Filter the 'NewUpdates' list to only include items that are actual downloadable files.
        # These are identified by the presence of the <SecuredFragment> tag inside their inner XML.
        $downloadableUpdates = $newUpdates | Where-Object { $_.Xml.Properties.SecuredFragment }

        Write-Host "Found $(@($downloadableUpdates).Count) potentially downloadable packages." -ForegroundColor Gray

        # Now, process each downloadable update
        foreach ($update in $downloadableUpdates) {
            $lookupId = $update.ID

            # Find the matching entry in the 'ExtendedUpdateInfo' list using the same numeric ID. Only the entry with the file list is needed.
            $extendedInfo = $allExtendedUpdates | Where-Object { ($_.ID -eq $lookupId) -and $_.Xml.Files } | Select-Object -First 1

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
            $applicabilityBlob = $update.Xml.ApplicabilityRules.Metadata.AppxPackageMetadata.AppxMetadata.ApplicabilityBlob

            # Define the regex based on the official package identity structure.
            # <Name>_<Version>_<Architecture>_<ResourceId>_<PublisherId>
            $regex = "^(?<Name>.+?)_(?<Version>\d+\.\d+\.\d+\.\d+)_(?<Architecture>[a-zA-Z0-9]+)_(?<ResourceId>.*?)_(?<PublisherId>[a-hjkmnp-tv-z0-9]{13})$"

            $packageInfo = [PSCustomObject]@{
                FullName         = $fullIdentifier
                FileName         = $fileName
                UpdateID         = $updateGuid
                RevisionNumber   = $revNum
                PackageRank      = $update.Xml.Properties.PackageRank -as [int]
                SupportsThisOS   = (Test-PackageSupportsThisWindows $applicabilityBlob)
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

            Write-Host "  -> CORRELATED: '$($packageInfo.PackageName)' $($packageInfo.Version) ($($packageInfo.Architecture))" -ForegroundColor Green
        }

        Write-Host "--- Correlation Complete ---" -ForegroundColor Cyan
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
            Write-Host "Step 3: Filtering packages for your system architecture ('$systemArch') and Windows version ($osVersion)..." -ForegroundColor Cyan

            # --- Filter the packages ---

            # 1. Keep the packages for this architecture (or neutral ones) that support this version of Windows
            $architecturePackages = @($fileIdentityMap.Values | Where-Object { ($_.Architecture -eq $systemArch) -or ($_.Architecture -eq 'neutral') })
            $usablePackages = @($architecturePackages | Where-Object { $_.SupportsThisOS })
            $skippedCount = $architecturePackages.Count - $usablePackages.Count
            if ($skippedCount -gt 0) {
                Write-Host "  -> Skipped $skippedCount package version(s) that are for other devices or need a newer version of Windows." -ForegroundColor Yellow
            }

            # 2. Find the main app's package to install: the one with the highest PackageRank. The rank is the Store's own order of preference for the app's packages.
            #    The highest version number isn't always the newest build, because some apps restarted their version numbers. For example, Phone Link's old builds are 2021.x and its current ones are 1.26.x, and Media Player's old Groove Music builds are 2019.x while its current ones are 11.x.
            #    The version number is only used to break a tie between equal ranks.
            $latestMainPackage = $null
            if ($mainPackageName) {
                $latestMainPackage = $usablePackages |
                    Where-Object { $_.PackageName -eq $mainPackageName } |
                    Sort-Object { $_.PackageRank }, { $_.Version -as [version] } -Descending |
                    Select-Object -First 1
            }

            # 3. Get the latest version of each dependency
            $filteredDependencies = @($usablePackages |
                Where-Object { $_.PackageName -ne $mainPackageName } |
                Group-Object PackageName, Architecture |
                ForEach-Object { $_.Group | Sort-Object { $_.Version -as [version] } -Descending | Select-Object -First 1 })

            # 4. Combine the lists for the final download queue
            $packagesToDownload = @()
            if ($latestMainPackage) {
                $packagesToDownload += $latestMainPackage
                Write-Host "  -> Found latest $mainPackageName package: $($latestMainPackage.FullName)" -ForegroundColor Green
            } elseif ($mainPackageName) {
                if ($architecturePackages | Where-Object { $_.PackageName -eq $mainPackageName }) {
                    Write-Warning "There are $mainPackageName packages, but none of them support this version of Windows ($osVersion)."
                } else {
                    Write-Warning "Could not find any $mainPackageName package for '$systemArch'."
                }
                $pauseAtEnd = $true
            }

            $packagesToDownload += $filteredDependencies
            Write-Host "  -> Found $($filteredDependencies.Count) dependencies for '$systemArch' architecture." -ForegroundColor Green

            # 5. Apply -NameFilter, if given
            if ($NameFilter) {
                $packagesToDownload = @($packagesToDownload | Where-Object { $_.PackageName -like $NameFilter })
                Write-Host "  -> $($packagesToDownload.Count) of those match -NameFilter '$NameFilter'." -ForegroundColor Green
            }

            if ($packagesToDownload.Count -eq 0) {
                Write-Host "No packages to download." -ForegroundColor Yellow
                $pauseAtEnd = $true
                return # Exits the script, but the 'finally' block at the end still runs
            }

            Write-Host "Total files to download: $($packagesToDownload.Count)" -ForegroundColor Gray
            foreach ($package in $packagesToDownload) { Write-Host "  -> $($package.FullName)" }
            Write-Host "------------------------------------------------------------"

            # --- Check Cache ---
            $useExistingFiles = $false
            $updateDetected = $false

            if ($ExistingInstallerFilesPath) {
                $allCached = $true
                foreach ($package in $packagesToDownload) {
                    $fileExtension = [System.IO.Path]::GetExtension($package.FileName)
                    $expectedFile = Join-Path $ExistingInstallerFilesPath "$($package.FullName)$($fileExtension)"
                    if (-not (Test-Path $expectedFile)) {
                        $allCached = $false
                        break
                    }
                }

                if ($allCached) {
                    Write-Host "Cached files are up to date. Using local source path: $ExistingInstallerFilesPath" -ForegroundColor Green
                    $useExistingFiles = $true
                    $workingDir = $ExistingInstallerFilesPath
                } elseif ($ForceCachedFilesOnly) {
                    Write-Host "WARNING: Cached files are out of date or missing, but using cache anyway (-ForceCachedFilesOnly)." -ForegroundColor Yellow
                    Write-Host "Please update your cache at: $ExistingInstallerFilesPath" -ForegroundColor Yellow
                    $useExistingFiles = $true
                    $workingDir = $ExistingInstallerFilesPath
                    $updateDetected = $true
                    $pauseAtEnd = $true
                } else {
                    Write-Host "Newer versions detected or cached files missing. Cache is out of date." -ForegroundColor Yellow
                    $updateDetected = $true
                    $pauseAtEnd = $true
                }
            }


            # --- Loop through the filtered list, get URLs, and download ---
            if (-not $useExistingFiles) {
                if (-not $noDownload -and -not (Test-Path -Path $workingDir)) {
                    New-Item -Path $workingDir -ItemType Directory -Force | Out-Null
                }

                Write-Host "Step 4: Fetching URLs and downloading files..." -ForegroundColor Cyan

                $originalPref = $ProgressPreference
                $ProgressPreference = 'SilentlyContinue'

                foreach ($package in $packagesToDownload) {
                    Write-Host "Processing: $($package.FullName)"

                    # Get the download URL for this specific package
                    # The timestamps are generated for each request, because values written directly in the template are fixed when the script starts and expire 5 minutes later
                    $requestTime = (Get-Date).ToUniversalTime()
                    $fileUrlRequestPayload = $fileUrlXmlTemplate -f $encryptedCookieData, $package.UpdateID, $package.RevisionNumber, $currentBranch, $flightRing, $flightingBranchName, $requestTime.ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'"), $requestTime.AddMinutes(5).ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'")
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
                        Write-Host "  -> FAILED to download $($newFileName). Error: $($_.Exception.Message)" -ForegroundColor Red
                        $pauseAtEnd = $true
                    }
                    Write-Host ""
                }

                $ProgressPreference = $originalPref

                Write-Host "------------------------------------------------------------"
                if (-not $noDownload) { Write-Host "Finished downloading packages to: $workingDir" -ForegroundColor Green }
            } else {
                Write-Host "Step 4: Skipping download step (Using existing cached files)." -ForegroundColor Gray
                Write-Host "------------------------------------------------------------"
            }

        } catch {
            Write-Host "An error occurred during the filtering or downloading phase:" -ForegroundColor Red
            Write-Host $_.Exception.ToString()
            $pauseAtEnd = $true
        }
    } else {
        $workingDir = $ExistingInstallerFilesPath
        Write-Host "Skipping update checks (-NoCheckLatestVersion). Using existing files from: $workingDir" -ForegroundColor Gray
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
    Write-Host "Step 5: Installing packages..." -ForegroundColor Cyan

    # Installs one package file. Returns nothing if it worked, otherwise the error message.
    function Install-PackageFile($packageFile) {
        try {
            Add-AppxPackage -Path $packageFile.FullName -ErrorAction Stop
            Write-Host "     SUCCESS." -ForegroundColor Green
            return $null
        } catch {
            # 0x80073D06 means a higher version of the package is already installed, so this one isn't needed
            if ($_.Exception.Message -match '0x80073D06') {
                Write-Host "     A newer version is already installed, so this one isn't needed." -ForegroundColor Green
                return $null
            }
            return $_.Exception.Message
        }
    }

    # 1. Get all downloaded package files and separate the main app from dependencies
    try {
        $allDownloadedFiles = Get-ChildItem -Path $workingDir -File | Where-Object { $_.Extension -in '.appx', '.msix', '.appxbundle', '.msixbundle' }

        # If this run found the list of packages to install and all of those files are in the folder, ignore any other files in it
        if ($packagesToDownload) {
            $expectedFileNames = $packagesToDownload | ForEach-Object { "$($_.FullName)$([System.IO.Path]::GetExtension($_.FileName))" }
            $expectedFiles = $allDownloadedFiles | Where-Object { $_.Name -in $expectedFileNames }
            if (@($expectedFiles).Count -eq @($expectedFileNames).Count) {
                $allDownloadedFiles = $expectedFiles
            }
        }

        # Apply -NameFilter to the files too, for when the package list came from the cache folder instead of the server
        if ($NameFilter) {
            $allDownloadedFiles = $allDownloadedFiles | Where-Object { ($_.BaseName -split '_')[0] -like $NameFilter }
        }

        # Keep only the highest version of each package (by name and architecture), in case the folder also has older files, such as a cache folder with old and new versions
        # File names follow the package identity format: <Name>_<Version>_<Architecture>_<ResourceId>_<PublisherId>
        $allDownloadedFiles = $allDownloadedFiles |
            Group-Object { ($_.BaseName -split '_')[0,2] -join '_' } |
            ForEach-Object { $_.Group | Sort-Object { ($_.BaseName -split '_')[1] -as [version] } -Descending | Select-Object -First 1 }

        $mainPackageFile = $null
        if ($mainPackageName) {
            $mainPackageFile = $allDownloadedFiles | Where-Object { $_.Name -like "$($mainPackageName)_*" } | Select-Object -First 1
        }
        $dependencyFiles = @($allDownloadedFiles | Where-Object { $_.FullName -ne $mainPackageFile.FullName })

        if ($dependencyFiles.Count -eq 0 -and -not $mainPackageFile) {
            Write-Warning "No package files found in '$workingDir' to install."
            $pauseAtEnd = $true
            return # Exits the script, but the 'finally' block at the end still runs
        }

        # 2. Install the dependencies first.
        #    A dependency can itself need another one installed first, so any that fail are tried again after the others, until a round of attempts makes no progress.
        if ($mainPackageFile -and $dependencyFiles.Count -gt 0) {
            Write-Host "Installing dependencies..."
        }
        $remainingFiles = @($dependencyFiles | Sort-Object Name)
        $lastErrors = @{}
        while ($remainingFiles.Count -gt 0) {
            $failedFiles = @()
            foreach ($package in $remainingFiles) {
                Write-Host "  -> Installing $($package.Name)"
                $errorMessage = Install-PackageFile $package
                if ($errorMessage) {
                    Write-Host "     Failed. Error: $errorMessage" -ForegroundColor Yellow
                    $failedFiles += $package
                    $lastErrors[$package.FullName] = $errorMessage
                }
            }
            if ($failedFiles.Count -eq 0 -or $failedFiles.Count -eq $remainingFiles.Count) {
                $remainingFiles = $failedFiles
                break
            }
            Write-Host "  -> Trying the $($failedFiles.Count) failed package(s) again, since some packages need others installed first..." -ForegroundColor Gray
            $remainingFiles = $failedFiles
        }
        foreach ($package in $remainingFiles) {
            Write-Host "  -> FAILED to install $($package.Name). Error: $($lastErrors[$package.FullName])" -ForegroundColor Red
            $pauseAtEnd = $true
        }

        # 3. Install the main app's package last
        if ($mainPackageFile) {
            Write-Host "Installing the main application..."
            Write-Host "  -> Installing $($mainPackageFile.Name)"
            $errorMessage = Install-PackageFile $mainPackageFile
            if ($errorMessage) {
                Write-Host "     FAILED to install $($mainPackageFile.Name). Error: $errorMessage" -ForegroundColor Red
                $pauseAtEnd = $true
            } else {
                Write-Host "     $mainPackageName has been installed/updated." -ForegroundColor Green
            }
        } elseif ($mainPackageName -and (-not $NameFilter -or $mainPackageName -like $NameFilter)) {
            Write-Warning "$mainPackageName package was not found in the download folder."
            $pauseAtEnd = $true
        }

        Write-Host "------------------------------------------------------------"
        Write-Host "Installation process finished." -ForegroundColor Gray

    } catch {
        Write-Host "A critical error occurred during the installation phase: $($_.Exception.Message)" -ForegroundColor Red
        $pauseAtEnd = $true
    }

} catch {
    Write-Host "An error occurred:" -ForegroundColor Red
    $pauseAtEnd = $true
    if ($_.Exception.Response) {
        $statusCode = $_.Exception.Response.StatusCode.value__
        $statusDescription = $_.Exception.Response.StatusDescription
        $errorLogPath = Join-Path $debugDir "ERROR_Response.txt"
        try {
            $stream = $_.Exception.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($stream)
            If ($debugSaveFiles) { $reader.ReadToEnd() | Set-Content -Path $errorLogPath }
        } catch { If ($debugSaveFiles) { "Could not read error response body." | Set-Content -Path $errorLogPath } }
        Write-Host "Status Code: $statusCode"
        Write-Host "Status Description: $statusDescription"
        If ($debugSaveFiles) { Write-Host "Server Response saved to '$errorLogPath'" }
    } else {
        Write-Host $_.Exception.ToString()
    }
} finally {
    # This also runs when the script exits early with 'return', so these messages and the pause aren't skipped
    if ($updateDetected) {
        Write-Host "`nACTION REQUIRED: Newer file versions are available." -ForegroundColor Yellow
        if (-not $useExistingFiles -and -not $noDownload) { Write-Host "The newer files were downloaded to: $workingDir" -ForegroundColor Yellow }
        Write-Host "Please update your cache at: $ExistingInstallerFilesPath" -ForegroundColor Yellow
    }

    # After an app was chosen in interactive mode, show the commands that do the same thing without the menus
    if ($interactiveMode -and ($ProductId -or $CategoryId)) {
        # The parameters for the commands. A switch has a Value of $null.
        $commandParameters = @()
        if ($ProductId) {
            $commandParameters += [PSCustomObject]@{ Name = 'ProductId'; Value = $ProductId }
        } else {
            $commandParameters += [PSCustomObject]@{ Name = 'CategoryId'; Value = $CategoryId }
        }

        # Include any other parameters the script was started with. -noInstall and -noDownload are added below from what was actually done.
        foreach ($parameter in $PSBoundParameters.GetEnumerator()) {
            if ($parameter.Key -in 'ProductId', 'CategoryId', 'noInstall', 'noDownload') { continue }
            if ($parameter.Value -is [switch]) {
                if ($parameter.Value.IsPresent) { $commandParameters += [PSCustomObject]@{ Name = $parameter.Key; Value = $null } }
            } else {
                $commandParameters += [PSCustomObject]@{ Name = $parameter.Key; Value = "$($parameter.Value)" }
            }
        }

        if ($noDownload) {
            $commandParameters += [PSCustomObject]@{ Name = 'noDownload'; Value = $null }
            $actionText = "list these packages"
        } elseif ($noInstall) {
            $commandParameters += [PSCustomObject]@{ Name = 'noInstall'; Value = $null }
            $actionText = "do this download"
        } else {
            $actionText = "do this install"
        }

        # Use the script's actual file name and path, in case it was renamed or moved
        $scriptPath = if ($PSCommandPath) { $PSCommandPath } else { "Install-Store-Apps.ps1" }
        $scriptName = Split-Path -Path $scriptPath -Leaf

        # 1. A command for a PowerShell window in the script's folder. Text values go in double quotes, with the characters that are special there escaped.
        if ($scriptName -match '^[\w.-]+$') {
            $consoleCommand = ".\$scriptName"
        } else {
            $consoleCommand = "& '.\$($scriptName -replace "'", "''")'"
        }
        foreach ($commandParameter in $commandParameters) {
            if ($null -eq $commandParameter.Value) {
                $consoleCommand += " -$($commandParameter.Name)"
            } elseif ($commandParameter.Value -match '^[\w.-]+$') {
                $consoleCommand += " -$($commandParameter.Name) $($commandParameter.Value)"
            } else {
                $consoleCommand += " -$($commandParameter.Name) `"$($commandParameter.Value -replace '([`$"])', '`$1')`""
            }
        }

        # 2. A line for another .ps1 script. It starts a new Windows PowerShell that bypasses the execution policy, and uses the script's full path.
        #    Text values go in single quotes, so nothing in them is expanded when the other script runs.
        $scriptCommand = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File '$($scriptPath -replace "'", "''")'"
        foreach ($commandParameter in $commandParameters) {
            if ($null -eq $commandParameter.Value) {
                $scriptCommand += " -$($commandParameter.Name)"
            } elseif ($commandParameter.Value -match '^[\w.-]+$') {
                $scriptCommand += " -$($commandParameter.Name) $($commandParameter.Value)"
            } else {
                # Windows PowerShell passes a value that has spaces to powershell.exe in double quotes, where a backslash at the end would escape the closing quote, so remove it
                $commandValue = if ($commandParameter.Value -match ' ') { $commandParameter.Value.TrimEnd('\') } else { $commandParameter.Value }
                $scriptCommand += " -$($commandParameter.Name) '$($commandValue -replace "'", "''")'"
            }
        }

        Write-Host ""
        Write-Host "To $actionText automatically in the future, use this command:" -ForegroundColor Cyan
        Write-Host "    $consoleCommand"
        Write-Host ""
        Write-Host "Or add this line to run it from another PowerShell .ps1 script (Update the path as needed):" -ForegroundColor Cyan
        Write-Host "    $scriptCommand"

        # If files were downloaded, mention that the download folder can be used as a cache next time. Not needed if a cache folder was already used, since it's in the commands above.
        if (-not $noDownload -and -not $ExistingInstallerFilesPath -and $workingDir) {
            $downloadFolderName = Split-Path -Path $workingDir -Leaf
            # Say where the folder is: the Downloads folder normally, or wherever -DownloadPath put it
            $downloadFolderParent = Split-Path -Path $workingDir -Parent
            $downloadFolderLocation = if ($downloadFolderParent.TrimEnd('\', '/') -eq $userDownloadsFolder.TrimEnd('\', '/')) { "Downloads" } else { $downloadFolderParent }
            Write-Host ""
            Write-Host "Optional: Copy the '$downloadFolderName' folder (in $downloadFolderLocation) somewhere, and add this parameter:" -ForegroundColor Cyan
            Write-Host "    -ExistingInstallerFilesPath `"C:\Whatever\Path\$downloadFolderName`""
        }
    }

    if ($pauseAtEnd) {
        Read-Host "`nPress Enter to exit"
    }
}
