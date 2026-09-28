@echo off
setlocal DisableDelayedExpansion

:: By: ThioJoe
:: Source Repo: https://github.com/ThioJoe/Windows-Sandbox-Tools

:: PURPOSE OF THIS SCRIPT: 
::    Lets you drag and drop files onto this batch file, and it will generate a .wsb config for
::    a sandbox which launch with those files pre-copied into it, WITHOUT needing to map a folder.

:: How it works (basics):
::    It compresses then base64 encodes the files, and embeds the data directly into the .wsb file itself.
::    Then it decompresses it from within the sandbox.

:: LIMITATIONS:
::    Each startup command must fit within the Windows command-line limit of around 32 KiB.
::    The script works around this by chunking the data into multiple <command> blocks,
::    but this takes extra time at the sandbox's launch. Progress is shown.

:: ======================== USER SETTINGS ========================
:: Sandbox switches: 1 = Enable, 0 = Disable.
set "WSB_NETWORKING=1"
set "WSB_MEMORY_MB=8192"
set "WSB_CLIPBOARD_REDIRECTION=1"
set "WSB_PROTECTED_CLIENT=1"
set "WSB_VIDEO_INPUT=0"
set "WSB_AUDIO_INPUT=0"
set "WSB_PRINTER_REDIRECTION=0"
set "WSB_VGPU=1"

	:: Absolute folder INSIDE the sandbox. This is not a mapped host folder.
set "WSB_SANDBOX_FOLDER=C:\Users\WDAGUtilityAccount\Desktop\Files"

:: ---WSB Output Options---
	:: "auto" uses the single dropped file's name stem. With multiple files, auto uses GeneratedSandbox.wsb.
set "WSB_OUTPUT_NAME=auto"
	:: 1 = overwrite any existing WSB. 0 = Don't overwrite, add numbered increment to file name
set "OVERWRITE_ANY_EXISTING_WSB=0"

:: ---Post Launch Scripting---
	:: Configure the wsb to run a specific .ps1 script file after copying the files.
	:: 1 = enabled,  0 = disabled
set "WSB_RUN_SCRIPT=1"
	:: Which script to run, if WSB_RUN_SCRIPT=1. 
	:: "auto" will run the first dropped PS1 onto the batch. Or set a specific filename.
set "WSB_SCRIPT_NAME=auto"

:: ---Other Options---
	:: Optional PowerShell -File arguments, for example: -launchingSandbox
set "WSB_SCRIPT_ARGUMENTS="
	:: Configure Windows in the sandbox to allow running .ps1 scripts after launch
	:: 1 = yes, 0 = leave unchanged.
set "WSB_ALLOW_SCRIPTS=1"

:: ---Large File Options---
	:: 1 = split oversized payloads into multiple startup commands. 0 = warn only.
set "WSB_SPLIT_LARGE_FILES=1"
:: =====================================================================

:: %~dp0 means next to the batch file. Change it if you want the wsb to go somewhere specific.
set "WSB_OUTPUT_FOLDER=%~dp0"

set "WSB_BUILDER_SELF=%~f0"
set "WSB_INPUT_COUNT=0"

:collect_files
if "%~1"=="" goto build_wsb
set /a WSB_INPUT_COUNT+=1 >nul
set "WSB_INPUT_%WSB_INPUT_COUNT%=%~f1"
shift
goto collect_files

:build_wsb
if not "%WSB_INPUT_COUNT%"=="0" goto run_builder
echo Drag and drop one or more files onto this batch file.
echo Edit the settings at the top of this file before using it.
echo(
pause
exit /b 1

:run_builder
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop';try{$s=[IO.File]::ReadAllText($env:WSB_BUILDER_SELF);$m='# == POWERSHELL ==';$i=$s.LastIndexOf($m);if($i -lt 0){throw 'Embedded PowerShell section not found.'};& ([scriptblock]::Create($s.Substring($i+$m.Length)))}catch{Write-Host ('ERROR: '+$_.Exception.Message) -ForegroundColor Red;exit 1}"
set "WSB_EXIT_CODE=%ERRORLEVEL%"
echo(
pause
exit /b %WSB_EXIT_CODE%

# == POWERSHELL ==

# --- Setup and helpers ---

if ($PSVersionTable.PSVersion -lt [version]'5.1') {
    throw 'Windows PowerShell 5.1 or later is required.'
}

function Get-SandboxSwitch {
    param([string]$Name)
    $value = [Environment]::GetEnvironmentVariable($Name)
    if ($value -ceq '1') { return 'Enable' }
    if ($value -ceq '0') { return 'Disable' }
    throw "$Name must be 1 or 0."
}

function ConvertTo-PowerShellLiteral {
    param([AllowEmptyString()][string]$Text)
    return "'" + $Text.Replace("'", "''") + "'"
}

function ConvertTo-WindowsArgument {
    param([AllowEmptyString()][string]$Text)
    $quoted = [regex]::Replace($Text, '(\\*)"', '$1$1\"')
    $quoted = [regex]::Replace($quoted, '(\\+)$', '$1$1')
    return '"' + $quoted + '"'
}

function ConvertTo-SandboxCommand {
    param([string]$Code)
    return 'powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -Command ' + (ConvertTo-WindowsArgument $Code)
}

function ConvertTo-VisibleSandboxCommand {
    param([string]$Code)
    $arguments = '-NoLogo -NoProfile -NoExit -ExecutionPolicy Bypass -Command ' + (ConvertTo-WindowsArgument $Code)
    $launcher = 'Start-Process -FilePath powershell.exe -WindowStyle Normal -ArgumentList ' + (ConvertTo-PowerShellLiteral $arguments)
    return ConvertTo-SandboxCommand $launcher
}

# --- Validate settings ---

$memoryMB = 0
if (-not [int]::TryParse($env:WSB_MEMORY_MB, [ref]$memoryMB) -or $memoryMB -le 0) {
    throw 'WSB_MEMORY_MB must be a positive whole number of megabytes.'
}

$options = [ordered]@{
    Networking = (Get-SandboxSwitch 'WSB_NETWORKING')
    MemoryInMB = $memoryMB.ToString([Globalization.CultureInfo]::InvariantCulture)
    ClipboardRedirection = (Get-SandboxSwitch 'WSB_CLIPBOARD_REDIRECTION')
    ProtectedClient = (Get-SandboxSwitch 'WSB_PROTECTED_CLIENT')
    VideoInput = (Get-SandboxSwitch 'WSB_VIDEO_INPUT')
    AudioInput = (Get-SandboxSwitch 'WSB_AUDIO_INPUT')
    PrinterRedirection = (Get-SandboxSwitch 'WSB_PRINTER_REDIRECTION')
    vGPU = (Get-SandboxSwitch 'WSB_VGPU')
}

$runScript = (Get-SandboxSwitch 'WSB_RUN_SCRIPT') -eq 'Enable'
$allowScripts = (Get-SandboxSwitch 'WSB_ALLOW_SCRIPTS') -eq 'Enable'
$overwriteExisting = (Get-SandboxSwitch 'OVERWRITE_ANY_EXISTING_WSB') -eq 'Enable'
$splitLargeFiles = (Get-SandboxSwitch 'WSB_SPLIT_LARGE_FILES') -eq 'Enable'

$destination = $env:WSB_SANDBOX_FOLDER
if ([string]::IsNullOrWhiteSpace($destination) -or $destination -notmatch '^[A-Za-z]:[\\/]') {
    throw 'WSB_SANDBOX_FOLDER must be an absolute sandbox path, such as C:\Users\WDAGUtilityAccount\Desktop\Files.'
}

# --- Input files and startup script ---

$inputFiles = [Collections.Generic.List[string]]::new()
$names = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$totalBytes = [long]0

for ($n = 1; $n -le [int]$env:WSB_INPUT_COUNT; $n++) {
    $path = [Environment]::GetEnvironmentVariable('WSB_INPUT_' + $n)
    if (-not [IO.File]::Exists($path)) {
        throw "File not found, inaccessible, or a folder was dropped: $path"
    }

    $name = [IO.Path]::GetFileName($path)
    if (-not $names.Add($name)) {
        throw "Two dropped files have the same filename: $name. Rename one before trying again."
    }

    $inputFiles.Add($path)
    $totalBytes += ([IO.FileInfo]::new($path)).Length
}

if ($inputFiles.Count -eq 0) { throw 'No input files were provided.' }

$scriptName = $null
if ($runScript -and $env:WSB_SCRIPT_NAME -ieq 'auto') {
    foreach ($path in $inputFiles) {
        if ([IO.Path]::GetExtension($path) -ieq '.ps1') {
            $scriptName = [IO.Path]::GetFileName($path)
            break
        }
    }
    $runScript = $null -ne $scriptName
}
elseif ($runScript) {
    $scriptName = $env:WSB_SCRIPT_NAME
    if ([string]::IsNullOrWhiteSpace($scriptName) -or [IO.Path]::GetExtension($scriptName) -ine '.ps1' -or -not $names.Contains($scriptName)) {
        throw 'WSB_SCRIPT_NAME must be the filename of a dropped .ps1 file when WSB_RUN_SCRIPT=1.'
    }
}

# --- Choose the output filename ---

$outputFolder = $env:WSB_OUTPUT_FOLDER
$outputName = $env:WSB_OUTPUT_NAME

if ($outputName -ieq 'auto') {
    if ($inputFiles.Count -eq 1) {
        $outputName = [IO.Path]::GetFileNameWithoutExtension($inputFiles[0]) + '.wsb'
    }
    else {
        $outputName = 'GeneratedSandbox.wsb'
    }
}

if ([string]::IsNullOrWhiteSpace($outputFolder)) { throw 'WSB_OUTPUT_FOLDER cannot be empty.' }
if ([string]::IsNullOrWhiteSpace($outputName) -or [IO.Path]::GetFileName($outputName) -cne $outputName -or [IO.Path]::GetExtension($outputName) -ine '.wsb') {
    throw 'WSB_OUTPUT_NAME must be auto or a filename ending in .wsb, without a folder path.'
}

$outputFolder = [IO.Path]::GetFullPath($outputFolder)
[IO.Directory]::CreateDirectory($outputFolder) | Out-Null
$outputPath = [IO.Path]::Combine($outputFolder, $outputName)

$suffix = 2
while (-not $overwriteExisting -and ([IO.File]::Exists($outputPath) -or [IO.Directory]::Exists($outputPath))) {
    $outputPath = [IO.Path]::Combine($outputFolder, [IO.Path]::GetFileNameWithoutExtension($outputName) + ' ' + $suffix + '.wsb')
    $suffix++
}

# --- ZIP and Base64 ---

# Read raw bytes to preserve each file exactly.
Write-Host ('Packing {0} file(s)...' -f $inputFiles.Count)
Add-Type -AssemblyName System.IO.Compression

$buffer = [IO.MemoryStream]::new()
try {
    $archive = [IO.Compression.ZipArchive]::new($buffer, [IO.Compression.ZipArchiveMode]::Create, $true)
    try {
        foreach ($path in $inputFiles) {
            $entry = $archive.CreateEntry([IO.Path]::GetFileName($path), [IO.Compression.CompressionLevel]::Optimal)
            $source = [IO.File]::OpenRead($path)
            try {
                $target = $entry.Open()
                try { $source.CopyTo($target) }
                finally { $target.Dispose() }
            }
            finally { $source.Dispose() }
        }
    }
    finally { $archive.Dispose() }

    $zipBytes = $buffer.Length
    $zipData = $buffer.ToArray()
    $base64 = [Convert]::ToBase64String($zipData)
}
finally { $buffer.Dispose() }

# --- Build the sandbox startup command ---

# Extract the files inside the sandbox, then remove the temporary ZIP.
$extractCode = 'Write-Host ''Decoding the embedded ZIP...'';Add-Type -AssemblyName System.IO.Compression.FileSystem;[IO.Directory]::CreateDirectory($d)|Out-Null;$z=[IO.Path]::Combine([IO.Path]::GetTempPath(),[IO.Path]::GetRandomFileName()+''.zip'');try{[IO.File]::WriteAllBytes($z,[Convert]::FromBase64String($b));Write-Host ''Extracting files...'';[IO.Compression.ZipFile]::ExtractToDirectory($z,$d)}finally{if([IO.File]::Exists($z)){[IO.File]::Delete($z)}};'

$launchCode = ''
if ($runScript) {
    # Only the short -File command is passed to this second PowerShell window.
    $scriptPath = [IO.Path]::Combine($destination, $scriptName)
    $scriptArguments = '-NoLogo -NoProfile -NoExit -ExecutionPolicy Bypass -File ' + (ConvertTo-WindowsArgument $scriptPath)
    if (-not [string]::IsNullOrWhiteSpace($env:WSB_SCRIPT_ARGUMENTS)) {
        $scriptArguments += ' ' + $env:WSB_SCRIPT_ARGUMENTS
    }
    $launchCode = 'Write-Host ''Launching startup script...'';Start-Process -FilePath powershell.exe -WorkingDirectory $d -ArgumentList ' + (ConvertTo-PowerShellLiteral $scriptArguments) + ';'
}

$errorCode = '$message=$_.Exception.ToString();Write-Progress -Id 1 -Activity ''Receiving and merging files'' -Completed;Write-Host (''ERROR: ''+$message) -ForegroundColor Red;$log=[IO.Path]::Combine([Environment]::GetFolderPath(''Desktop''),''Sandbox-Setup-Error.txt'');$message|Set-Content -LiteralPath $log -Encoding UTF8;Write-Host (''Error details saved to: ''+$log);Write-Host ''Close this window when finished.'';'

$policyCode = ''
if ($allowScripts) {
    $policyCode = 'Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;'
}

$progressStartCode = '$ProgressPreference=''Continue'';$Host.UI.RawUI.WindowTitle=''Preparing Windows Sandbox files'';Write-Host ''Preparing Windows Sandbox files...'' -ForegroundColor Cyan;Write-Host ''Close this window or press Ctrl+C to cancel.'';'
$progressDoneCode = 'Write-Host (''Files are ready in: ''+$d) -ForegroundColor Green;exit 0;'
$setupCode = '$ErrorActionPreference=''Stop'';try{' + $progressStartCode + $policyCode + '$d=' + (ConvertTo-PowerShellLiteral $destination) + ';$b=' + (ConvertTo-PowerShellLiteral $base64) + ';' + $extractCode + $launchCode + $progressDoneCode + '}catch{' + $errorCode + '}'
$command = ConvertTo-VisibleSandboxCommand $setupCode

# --- Split oversized payloads into independent commands ---

$commands = [Collections.Generic.List[string]]::new()
$chunkCount = 0
if ($splitLargeFiles -and $command.Length -gt 32766) {
    # A multiple of 4 lets each Base64 chunk decode independently.
    $chunkCharacters = 30000
    $chunkCount = [int][Math]::Ceiling($base64.Length / [double]$chunkCharacters)
    $partDirectory = 'WSB-Payload-' + [Guid]::NewGuid().ToString('N')
    $partDirectoryCode = '$r=[IO.Path]::Combine([IO.Path]::GetTempPath(),' + (ConvertTo-PowerShellLiteral $partDirectory) + ');'

    $hashProvider = [Security.Cryptography.SHA256]::Create()
    try { $expectedHash = [BitConverter]::ToString($hashProvider.ComputeHash($zipData)).Replace('-', '') }
    finally { $hashProvider.Dispose() }

    for ($i = 0; $i -lt $chunkCount; $i++) {
        $offset = $i * $chunkCharacters
        $chunk = $base64.Substring($offset, [Math]::Min($chunkCharacters, $base64.Length - $offset))
        $partName = $i.ToString('D8')

        # Publish the .part file only after its bytes have been written and closed.
        $writerCode = '$ErrorActionPreference=''Stop'';' + $partDirectoryCode +
            '$f=[IO.Path]::Combine($r,' + (ConvertTo-PowerShellLiteral $partName) + ');try{' +
            '[IO.Directory]::CreateDirectory($r)|Out-Null;' +
            '[IO.File]::WriteAllBytes($f+''.tmp'',[Convert]::FromBase64String(' + (ConvertTo-PowerShellLiteral $chunk) + '));' +
            '[IO.File]::Move($f+''.tmp'',$f+''.part'')' +
            '}catch{try{[IO.File]::WriteAllText($f+''.error'',$_.Exception.ToString())}catch{};exit 1}'
        $chunkCommand = ConvertTo-SandboxCommand $writerCode
        if ($chunkCommand.Length -gt 32766) { throw 'A chunk command exceeds the Windows command-line limit.' }
        $commands.Add($chunkCommand)
    }

    # This coordinator tolerates writers completing in any order.
    $joinCode = $partDirectoryCode + '$n=' + $chunkCount +
        ';$expectedSize=' + $zipBytes.ToString([Globalization.CultureInfo]::InvariantCulture) +
        ';$expectedHash=' + (ConvertTo-PowerShellLiteral $expectedHash) + ';' + (@(
            '[IO.Directory]::CreateDirectory($r)|Out-Null;'
            '$watch=[Diagnostics.Stopwatch]::StartNew();$progressState=@{LastUpdate=-500};'
            'function Show-ChunkProgress([int]$merged){'
            '$now=$watch.ElapsedMilliseconds;if($now-$progressState.LastUpdate -lt 500 -and $merged -lt $n){return};$progressState.LastUpdate=$now;'
            '$received=[IO.Directory]::GetFiles($r,''*.part'').Length;'
            '$status=(''{0}/{1} chunks received; {2}/{1} merged; {3}s elapsed'' -f $received,$n,$merged,[int]$watch.Elapsed.TotalSeconds);'
            'Write-Progress -Id 1 -Activity ''Receiving and merging files'' -Status $status -PercentComplete ([int][Math]::Floor(100.0*$merged/$n))};'
            'Write-Host (''Waiting for ''+$n+'' chunks. There is no timeout.'');Show-ChunkProgress 0;'
            '$z=[IO.Path]::Combine($r,''payload.zip'');'
            '$outputStream=[IO.File]::Open($z,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write);'
            'try{for($i=0;$i -lt $n;$i++){'
            '$f=[IO.Path]::Combine($r,$i.ToString(''D8''));'
            'while(-not [IO.File]::Exists($f+''.part'')){'
            'if([IO.File]::Exists($f+''.error'')){throw (''Chunk ''+($i+1)+'' failed: ''+[IO.File]::ReadAllText($f+''.error''))};'
            'Show-ChunkProgress $i;'
            'Start-Sleep -Milliseconds 200};'
            '$partStream=[IO.File]::OpenRead($f+''.part'');'
            'try{$partStream.CopyTo($outputStream)}finally{$partStream.Dispose()}'
            ';Show-ChunkProgress ($i+1)'
            '}}finally{$outputStream.Dispose()};'
            'Write-Progress -Id 1 -Activity ''Receiving and merging files'' -Completed;Write-Host ''All chunks received and merged.'';'
            'Write-Host ''Verifying ZIP size and SHA-256...'';'
            'if(([IO.FileInfo]::new($z)).Length -ne $expectedSize){throw ''Reassembled ZIP size does not match.''};'
            '$hashProvider=[Security.Cryptography.SHA256]::Create();$zipStream=[IO.File]::OpenRead($z);'
            'try{$actualHash=[BitConverter]::ToString($hashProvider.ComputeHash($zipStream)).Replace(''-'','''')}finally{$zipStream.Dispose();$hashProvider.Dispose()};'
            'if($actualHash -cne $expectedHash){throw ''Reassembled ZIP SHA-256 does not match.''};'
            'Add-Type -AssemblyName System.IO.Compression.FileSystem;[IO.Directory]::CreateDirectory($d)|Out-Null;'
            'Write-Host ''Extracting files...'';[IO.Compression.ZipFile]::ExtractToDirectory($z,$d);'
            'Write-Host ''Removing temporary chunks...'';[IO.Directory]::Delete($r,$true);'
        ) -join '')
    $coordinatorCode = '$ErrorActionPreference=''Stop'';try{' + $progressStartCode + $policyCode + '$d=' +
        (ConvertTo-PowerShellLiteral $destination) + ';' + $joinCode + $launchCode + $progressDoneCode + '}catch{' + $errorCode + '}'
    $commands.Insert(0, (ConvertTo-VisibleSandboxCommand $coordinatorCode))
}
else {
    $commands.Add($command)
}

$maxCommandLength = 0
foreach ($startupCommand in $commands) {
    if ($startupCommand.Length -gt $maxCommandLength) { $maxCommandLength = $startupCommand.Length }
}

# --- Write the WSB file ---

# XML escapes special characters in paths and commands.
$document = [Xml.XmlDocument]::new()
$configuration = $document.CreateElement('Configuration')
$document.AppendChild($configuration) | Out-Null

foreach ($option in $options.GetEnumerator()) {
    $element = $document.CreateElement($option.Key)
    $element.InnerText = [string]$option.Value
    $configuration.AppendChild($element) | Out-Null
}

$logon = $document.CreateElement('LogonCommand')
foreach ($startupCommand in $commands) {
    $commandElement = $document.CreateElement('Command')
    $commandElement.InnerText = $startupCommand
    $logon.AppendChild($commandElement) | Out-Null
}
$configuration.AppendChild($logon) | Out-Null

$xmlSettings = [Xml.XmlWriterSettings]::new()
$xmlSettings.Indent = $true
$xmlSettings.NewLineChars = "`r`n"
$xmlSettings.Encoding = [Text.UTF8Encoding]::new($false)
$xmlSettings.OmitXmlDeclaration = $true

$fileMode = [IO.FileMode]::CreateNew
if ($overwriteExisting) { $fileMode = [IO.FileMode]::Create }
$stream = [IO.File]::Open($outputPath, $fileMode, [IO.FileAccess]::Write)
try {
    $writer = [Xml.XmlWriter]::Create($stream, $xmlSettings)
    try { $document.Save($writer) }
    finally { $writer.Dispose() }
}
finally { $stream.Dispose() }

# --- Show the result ---

Write-Host ''
Write-Host ('Created: ' + $outputPath) -ForegroundColor Green
Write-Host ''
Write-Host ('Sandbox folder: ' + $destination)
Write-Host ('Files: {0}; original bytes: {1:N0}; ZIP bytes: {2:N0}' -f $inputFiles.Count, $totalBytes, $zipBytes)
Write-Host ('Base64 characters: {0:N0}; startup commands: {1:N0}; longest command: {2:N0} characters' -f $base64.Length, $commands.Count, $maxCommandLength)

if ($chunkCount -gt 0) {
    Write-Host ('Chunked transfer: {0:N0} chunks; visible progress window with no timeout.' -f $chunkCount)
	Write-Host ''
    Write-Warning 'File(s) had to be chunked because of their size and will be recombined at launch. This will take extra time.'
}
if ($maxCommandLength -gt 32766) {
    Write-Warning ('Longest startup command is {0:N0} characters; the expected maximum is 32,766 (32,767 including the terminator). Windows Sandbox may not run it. The .wsb file was still generated.' -f $maxCommandLength)
}

if ($runScript) { Write-Host ('Run after extraction: ' + $scriptName) }
Write-Host ''
Write-Host 'Double-click the generated .wsb file to start the sandbox.'
