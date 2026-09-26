#requires -Version 5.1
<#
.SYNOPSIS
Creates instant disposable remote SSH access inside Windows Sandbox.
.DESCRIPTION
Run this file in an elevated Windows PowerShell window INSIDE a fresh x64 Windows Sandbox.

--- How it Works ---

Setup:
    - Downloads Microsoft Win32-OpenSSH and the latest stable cloudflared release
    - Creates one temporary administrator account,
    - Sets up a loopback-only SSH service
    - Establishes an outbound Cloudflare Quick Tunnel

Creates Usage Tools:
    - Creates instruction files that can be provided directly to the AI, depending on its operating system:
        SANDBOX-CONNECT-INSTRUCTIONS-LINUX.zip
        SANDBOX-CONNECT-INSTRUCTIONS-WINDOWS.zip
        User-Instructions.txt (More detailed instructions for the user)
    - Each ZIP contains the private key, pinned host key, SSH configuration, INSTRUCTIONS.MD, and matching cloudflared binaries.
    - The Linux ZIP also includes CONNECT.sh to prepare the client,  install OpenSSH if missing on supported Linux distributions, and connect.

---------

Important Notes:
- Cloudflare Quick Tunnels are a free development service without an uptime guarantee. 
      - Currently this script doesn't support persistent tunnels or custom domains.
- This script pins Microsoft's OpenSSH release 10.0.0.0p2-Preview (current latest preview).
      - It's not updated very often, but the script will warn you if a newer version is available.

More Information:
- Use -TestConnection to require local AND public-route SSH tests before export.
- The client needs an OpenSSH client; cloudflared is included in the ZIPs.
- No Cloudflare account, domain, router changes, winget, Store, or Python needed.
- Internet access is required; "self-contained" does not mean offline.
- No SSH agent, firewall rule, scheduled task, or host-side service is installed.
- All changes occur within the already-running Windows Sandbox.


Keep the window open. Ctrl+C stops access; closing Windows Sandbox destroys
the guest, including any processes or files created by the remote user.
Access expires based on the value of -SessionMinutes, or its default value.

The account has administrator access INSIDE THE GUEST. Existing mapped host
folders, shared clipboard, and reachable LAN resources remain accessible.

Use a fresh sandbox without private files or writable host-folder mappings.

The WDAGUtilityAccount check prevents accidental host execution; it is not
proof of isolation. Do not rename a host account to bypass it.

.PARAMETER SessionMinutes
How long to keep remote access available after successful setup. Default 120.
.PARAMETER Port
The guest loopback SSH port. Default 2222. This is not an Internet-facing port.
.PARAMETER TestConnection
Run local and public-route SSH tests before exporting credentials. Off by default.
.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Start-AiSandbox.ps1
.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Start-AiSandbox.ps1 -SessionMinutes 240
.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Start-AiSandbox.ps1 -TestConnection
.LINK
https://github.com/ThioJoe/Windows-Sandbox-Tools
.LINK
https://github.com/PowerShell/Win32-OpenSSH
.LINK
https://github.com/cloudflare/cloudflared
#>
[CmdletBinding()]
param(
    [ValidateRange(1, 4320)]    # Default max is 3 days (4320 minutes), but you can change this if needed.
    [int]$SessionMinutes = 180, # 3 Hours
    [ValidateRange(1024, 65535)]
    [int]$Port = 2222,
    [switch]$TestConnection
)

# Strict mode, stop on any error, and hide progress bars (they slow down downloads)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# ------------------------------------------------------------------------------------------------------------------
# The OpenSSH pin is from the official release asset as of 2026. Use OpenSSH-Win64.zip.
# This is not updated very often so worth pinning. The script will check and warn if a newer version is available.
$openSshPinnedTag = '10.0.0.0p2-Preview'
$openSshUrl = "https://github.com/PowerShell/Win32-OpenSSH/releases/download/$openSshPinnedTag/OpenSSH-Win64.zip"
$openSshSha256 = '23f50f3458c4c5d0b12217c6a5ddfde0137210a30fa870e98b29827f7b43aba5'
# ------------------------------------------------------------------------------------------------------------------

# File & Folder Names
$userInstructionsFileName = 'User-Instructions.txt'
$connectionFolderName = 'Sandbox-Connect-Files'

# Directory to store connection files. Default is the Desktop under 'Sandbox-Connect-Files'
$connectionBaseDirectory = [Environment]::GetFolderPath('Desktop')
$connectionDirectory = Join-Path ($connectionBaseDirectory) $connectionFolderName



# -------------------------------------- Check that we're running in the Windows Sandbox --------------------------------------
# This script is intended to be run from within the Windows Sandbox.
# Remove this section if you intend to run it outside the Windows Sandbox for some reason.

$notInSandboxErrorString = "ERROR: This script is intended to be run from WITHIN the Windows Sandbox.`nIt appears you are running this from outside the sandbox.`nIf you need to run it outside the sandbox for some reason, you can comment out this section within the script near the top."
if ($env:USERNAME -ne "WDAGUtilityAccount") {
    Write-host "`n`n$notInSandboxErrorString" -ForegroundColor Red
    Write-host "`n`nPress Enter to exit." -ForegroundColor Yellow
    Read-Host
    exit
}
# Alternative check for running inside the Windows Sandbox
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
if (($identity.Name -split '\\')[-1] -ine 'WDAGUtilityAccount') {
    throw "$notInSandboxErrorString"
}

# -------------------------------------------------------------------------------------------------------------------------------


# Other validation checks
if ($env:OS -ne 'Windows_NT' -or $PSVersionTable.PSEdition -ne 'Desktop') {
    throw 'Use Windows PowerShell 5.1 (powershell.exe) inside Windows Sandbox.'
}
if (-not ([Security.Principal.WindowsPrincipal]$identity).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw 'Open Windows PowerShell as administrator INSIDE Windows Sandbox, then run this script again.'
}
if (-not [Environment]::Is64BitProcess -or $env:PROCESSOR_ARCHITECTURE -ne 'AMD64') {
    throw 'This script targets x64 Windows Sandbox and requires 64-bit Windows PowerShell.'
}
if (Get-Service -Name sshd -ErrorAction SilentlyContinue) {
    throw 'An sshd service already exists. Use a fresh Windows Sandbox; this script will not replace an existing SSH installation.'
}
foreach ($configName in @('config.yml', 'config.yaml')) {
    if (Test-Path -LiteralPath (Join-Path $env:USERPROFILE ".cloudflared\$configName")) {
        throw 'An existing cloudflared configuration can interfere with Quick Tunnels. Use a fresh Windows Sandbox.'
    }
}

# -------------------------------------- Session values --------------------------------------
# Unique names and paths for this run
$sessionId = [Guid]::NewGuid().ToString('N')
$userName = 'ai' + $sessionId.Substring(0, 10)
$hostAlias = 'ai-sandbox-' + $sessionId.Substring(0, 12)
$root = Join-Path $env:ProgramData "AiSandbox-$sessionId"
$sshData = Join-Path $env:ProgramData 'ssh'
$shellPath = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
$registryPath = 'HKLM:\SOFTWARE\OpenSSH'

# Tracking variables so the finally block at the end knows what to clean up
$createdSshData = $false
$createdUser = $false
$createdService = $false
$changedShell = $false
$hadShellValue = $false
$oldShellValue = $null
$oldShellKind = $null
$tunnel = $null
$tunnelJob = $null
$connectionFiles = @()
$clientBundleDirectory = $null
$ready = $false
$oldTls = [Net.ServicePointManager]::SecurityProtocol

# -------------------------------------- Helper functions --------------------------------------

# Restrict a file or folder to Administrators and SYSTEM only (no inherited permissions)
function Set-SystemAdminAcl {
    param([string]$Path, [switch]$Directory)
    $admins = [Security.Principal.SecurityIdentifier]'S-1-5-32-544'
    $system = [Security.Principal.SecurityIdentifier]'S-1-5-18'
    if ($Directory) {
        $acl = New-Object Security.AccessControl.DirectorySecurity
        $inheritance = [Security.AccessControl.InheritanceFlags]'ContainerInherit, ObjectInherit'
    } else {
        $acl = New-Object Security.AccessControl.FileSecurity
        $inheritance = [Security.AccessControl.InheritanceFlags]::None
    }
    $acl.SetAccessRuleProtection($true, $false)
    $acl.SetOwner($admins)
    foreach ($sid in @($admins, $system)) {
        $rule = New-Object Security.AccessControl.FileSystemAccessRule($sid, 'FullControl', $inheritance, 'None', 'Allow')
        $acl.AddAccessRule($rule)
    }
    Set-Acl -LiteralPath $Path -AclObject $acl
}

# Restrict a file to the current user and SYSTEM only (used for the private key and the ZIPs)
function Set-ClientKeyAcl {
    param([string]$Path)
    $acl = New-Object Security.AccessControl.FileSecurity
    $acl.SetAccessRuleProtection($true, $false)
    $acl.SetOwner($identity.User)
    foreach ($sid in @($identity.User, [Security.Principal.SecurityIdentifier]'S-1-5-18')) {
        $acl.AddAccessRule((New-Object Security.AccessControl.FileSystemAccessRule($sid, 'FullControl', 'Allow')))
    }
    Set-Acl -LiteralPath $Path -AclObject $acl
}

# Build quoted command-line argument strings for native executables
function ConvertTo-NativeArgument {
    param([AllowEmptyString()][string]$Value)
    # Quote according to the Windows native argv rules, including empty values.
    $escaped = [regex]::Replace($Value, '(\\*)"', '$1$1\"')
    $escaped = [regex]::Replace($escaped, '(\\+)$', '$1$1')
    return '"' + $escaped + '"'
}

function Join-NativeArguments {
    param([string[]]$Values)
    return (($Values | ForEach-Object { ConvertTo-NativeArgument $_ }) -join ' ')
}

# Run a native executable with a timeout, returning its exit code, stdout, and stderr
function Invoke-Native {
    param([string]$FilePath, [string[]]$Arguments, [int]$TimeoutSeconds = 30)
    $info = New-Object Diagnostics.ProcessStartInfo
    $info.FileName = $FilePath
    $info.Arguments = Join-NativeArguments $Arguments
    $info.UseShellExecute = $false
    $info.CreateNoWindow = $true
    $info.RedirectStandardInput = $true
    $info.RedirectStandardOutput = $true
    $info.RedirectStandardError = $true
    $process = New-Object Diagnostics.Process
    $process.StartInfo = $info
    try {
        [void]$process.Start()
        $process.StandardInput.Close()
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
            # Terminate ProxyCommand children too, before the parent's PID vanishes.
            $killer = Start-Process -FilePath (Join-Path $env:WINDIR 'System32\taskkill.exe') -ArgumentList "/PID $($process.Id) /T /F" -WindowStyle Hidden -PassThru
            [void]$killer.WaitForExit(5000)
            $killer.Dispose()
            if (-not $process.HasExited) { $process.Kill() }
            throw "Timed out after $TimeoutSeconds seconds: $([IO.Path]::GetFileName($FilePath))"
        }
        if (-not $stdout.Wait(5000) -or -not $stderr.Wait(5000)) {
            throw "A child process kept output pipes open: $([IO.Path]::GetFileName($FilePath))"
        }
        return [pscustomobject]@{ ExitCode = $process.ExitCode; Output = $stdout.Result; Error = $stderr.Result }
    } finally {
        $process.Dispose()
    }
}

# Same as Invoke-Native, but throws on a non-zero exit code and returns only stdout
function Invoke-Checked {
    param([string]$FilePath, [string[]]$Arguments, [int]$TimeoutSeconds = 30)
    $result = Invoke-Native -FilePath $FilePath -Arguments $Arguments -TimeoutSeconds $TimeoutSeconds
    if ($result.ExitCode -ne 0) {
        throw "$([IO.Path]::GetFileName($FilePath)) exited with $($result.ExitCode): $($result.Error) $($result.Output)"
    }
    return $result.Output
}

# Download a file (up to 3 attempts) and verify its SHA256 hash
function Get-VerifiedDownload {
    param([string]$Url, [string]$Path, [string]$Sha256)
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        try {
            Invoke-WebRequest -Uri $Url -OutFile $Path -UseBasicParsing -TimeoutSec 180
            break
        } catch {
            if ($attempt -eq 3) { throw }
            Start-Sleep -Seconds 2
        }
    }
    $actual = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
    if ($actual -ine $Sha256) {
        Remove-Item -LiteralPath $Path -Force
        throw "SHA256 verification failed for $([IO.Path]::GetFileName($Path)). The download was not executed."
    }
    Unblock-File -LiteralPath $Path
}

# Get the download URLs and SHA256 hashes for the latest stable cloudflared release from GitHub
function Get-CloudflaredRelease {
    $headers = @{ Accept = 'application/vnd.github+json'; 'User-Agent' = 'AiSandbox' }
    $release = Invoke-RestMethod -Uri 'https://api.github.com/repos/cloudflare/cloudflared/releases/latest' -Headers $headers -TimeoutSec 60
    if ($release.draft -or $release.prerelease -or $release.tag_name -notmatch '^\d{4}\.\d+\.\d+$') {
        throw 'GitHub did not return a supported stable cloudflared release.'
    }
    $assets = @{}
    foreach ($name in @('cloudflared-windows-amd64.exe', 'cloudflared-linux-amd64', 'cloudflared-linux-arm64')) {
        $matchingAssets = @($release.assets | Where-Object { $_.name -ceq $name })
        if ($matchingAssets.Count -ne 1) { throw "The latest cloudflared release is missing a unique asset: $name" }
        $asset = $matchingAssets[0]
        $digest = $asset.PSObject.Properties['digest']
        if (-not $digest -or [string]$digest.Value -cnotmatch '^sha256:[a-f0-9]{64}$') {
            throw "The latest cloudflared asset has no valid SHA256 digest: $name"
        }
        $url = 'https://github.com/cloudflare/cloudflared/releases/download/' + $release.tag_name + '/' + $name
        if ($asset.browser_download_url -cne $url) { throw "Unexpected cloudflared asset URL: $name" }
        $assets[$name] = [pscustomobject]@{ Url = $url; Sha256 = ([string]$digest.Value).Substring(7) }
    }
    return [pscustomobject]@{ Version = $release.tag_name; Assets = $assets }
}

# Warn if Win32-OpenSSH has a newer release than the pinned version
function Test-OpenSshPinCurrent {
    param([string]$PinnedTag)
    try {
        $headers = @{ Accept = 'application/vnd.github+json'; 'User-Agent' = 'AiSandbox' }
        $latest = Invoke-RestMethod -Uri 'https://api.github.com/repos/PowerShell/Win32-OpenSSH/releases/latest' -Headers $headers -TimeoutSec 5  # Short timeout
        if ($latest.tag_name -and $latest.tag_name -cne $PinnedTag) {
            Write-Warning "A newer Win32-OpenSSH release is available: $($latest.tag_name) (this script is pinned to use $PinnedTag). Consider updating the version pinned in this script (set near the top)."
        }
    } catch {
        Write-Warning "Could not check for a newer Win32-OpenSSH release: $($_.Exception.Message)"
    }
}

# Write text as UTF-8 without a BOM
function Write-Utf8 {
    param([string]$Path, [string]$Text)
    [IO.File]::WriteAllText($Path, $Text, (New-Object Text.UTF8Encoding($false)))
}

# Read a text file even while another process (cloudflared) has it open for writing
function Read-SharedText {
    param([string]$Path)
    $stream = $null
    $reader = $null
    try {
        $sharing = [IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete
        $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, $sharing)
        $reader = New-Object IO.StreamReader($stream, [Text.Encoding]::UTF8, $true)
        return $reader.ReadToEnd()
    } finally {
        if ($reader) { $reader.Dispose() }
        elseif ($stream) { $stream.Dispose() }
    }
}

# -------------------------------------- Main setup --------------------------------------
# The finally block at the bottom undoes the changes (tunnel, SSH service, account, registry, credentials) when the session ends or fails.

try {
    # Enable TLS 1.2 for downloads and create the locked-down session folder
    [Net.ServicePointManager]::SecurityProtocol = $oldTls -bor [Net.SecurityProtocolType]::Tls12
    [void](New-Item -Path $root -ItemType Directory)
    Set-SystemAdminAcl -Path $root -Directory
    Write-Host "`nPreparing disposable SSH access inside Windows Sandbox..." -ForegroundColor Cyan
    Write-Verbose "Session files: $root"

    # --- Job object that kills cloudflared when this script exits ---
    # This handle is private to the controller. Windows kills cloudflared if
    # the controller crashes or its window closes, as well as during cleanup.
    if (-not ('AiSandbox.KillOnCloseJob' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
namespace AiSandbox {
    public sealed class KillOnCloseJob : IDisposable {
        [StructLayout(LayoutKind.Sequential)]
        struct BasicLimits {
            public long ProcessTime, JobTime;
            public uint Flags;
            public UIntPtr MinimumWorkingSet, MaximumWorkingSet;
            public uint ActiveProcessLimit;
            public UIntPtr Affinity;
            public uint PriorityClass, SchedulingClass;
        }
        [StructLayout(LayoutKind.Sequential)]
        struct IoCounters {
            public ulong ReadOperations, WriteOperations, OtherOperations;
            public ulong ReadBytes, WriteBytes, OtherBytes;
        }
        [StructLayout(LayoutKind.Sequential)]
        struct ExtendedLimits {
            public BasicLimits Basic;
            public IoCounters Io;
            public UIntPtr ProcessMemory, JobMemory, PeakProcessMemory, PeakJobMemory;
        }
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        static extern IntPtr CreateJobObject(IntPtr attributes, string name);
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool SetInformationJobObject(IntPtr job, int infoClass, IntPtr info, uint size);
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool AssignProcessToJobObject(IntPtr job, IntPtr process);
        [DllImport("kernel32.dll")]
        static extern bool CloseHandle(IntPtr handle);
        IntPtr handle;
        public KillOnCloseJob() {
            handle = CreateJobObject(IntPtr.Zero, null);
            if (handle == IntPtr.Zero) throw new Win32Exception();
            ExtendedLimits limits = new ExtendedLimits();
            limits.Basic.Flags = 0x2000; // JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
            int size = Marshal.SizeOf(typeof(ExtendedLimits));
            IntPtr buffer = Marshal.AllocHGlobal(size);
            try {
                Marshal.StructureToPtr(limits, buffer, false);
                if (!SetInformationJobObject(handle, 9, buffer, (uint)size)) {
                    int error = Marshal.GetLastWin32Error();
                    Dispose();
                    throw new Win32Exception(error);
                }
            } finally { Marshal.FreeHGlobal(buffer); }
        }
        public void Add(IntPtr processHandle) {
            if (!AssignProcessToJobObject(handle, processHandle)) throw new Win32Exception();
        }
        public void Dispose() {
            if (handle != IntPtr.Zero) { CloseHandle(handle); handle = IntPtr.Zero; }
            GC.SuppressFinalize(this);
        }
        ~KillOnCloseJob() { Dispose(); }
    }
}
'@
    }

    # --- Download and verify OpenSSH and cloudflared ---
    Test-OpenSshPinCurrent -PinnedTag $openSshPinnedTag

    Write-Host "Downloading and verifying Microsoft OpenSSH and cloudflared..."
    $archive = Join-Path $root 'OpenSSH.zip'
    $cloudflared = Join-Path $root 'cloudflared.exe'
    $cloudflaredRelease = Get-CloudflaredRelease
    $cloudflaredAssets = $cloudflaredRelease.Assets
    Write-Verbose "Bundling cloudflared $($cloudflaredRelease.Version) for Windows x64 and Linux x64/ARM64..."
    Get-VerifiedDownload $openSshUrl $archive $openSshSha256
    Get-VerifiedDownload $cloudflaredAssets['cloudflared-windows-amd64.exe'].Url $cloudflared $cloudflaredAssets['cloudflared-windows-amd64.exe'].Sha256
    foreach ($name in @('cloudflared-linux-amd64', 'cloudflared-linux-arm64')) {
        Get-VerifiedDownload $cloudflaredAssets[$name].Url (Join-Path $root $name) $cloudflaredAssets[$name].Sha256
    }

    # Extract OpenSSH
    Expand-Archive -LiteralPath $archive -DestinationPath $root
    Remove-Item -LiteralPath $archive -Force
    $bin = Join-Path $root 'OpenSSH-Win64'

    $icacls = Join-Path $env:WINDIR 'System32\icacls.exe'
    # OpenSSH's unprivileged helpers need to traverse the parent and execute
    # the binaries. Only the binary directory receives inheritable access.
    [void](Invoke-Checked $icacls @($root, '/grant', '*S-1-5-11:(RX)', '/Q'))
    [void](Invoke-Checked $icacls @($bin, '/grant', '*S-1-5-11:(OI)(CI)(RX)', '/T', '/Q'))

    # Confirm the required OpenSSH executables exist
    $sshd = Join-Path $bin 'sshd.exe'
    $ssh = Join-Path $bin 'ssh.exe'
    $keygen = Join-Path $bin 'ssh-keygen.exe'
    foreach ($file in @($sshd, $ssh, $keygen, (Join-Path $bin 'sshd-session.exe'), (Join-Path $bin 'sshd-auth.exe'), (Join-Path $bin 'sftp-server.exe'))) {
        if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { throw "Missing required OpenSSH file: $file" }
    }

    # --- Generate SSH keys ---
    # Client key (given to the AI) and host key (pinned by the client to verify this server)
    $clientKey = Join-Path $root 'sandbox-key'
    $hostKey = Join-Path $root 'ssh_host_ed25519_key'
    $authorizedKeys = Join-Path $root 'authorized_keys'
    $knownHosts = Join-Path $root 'sandbox-known-hosts'
    $serverConfig = Join-Path $root 'sshd_config'
    $sshLog = Join-Path $sshData 'logs\sshd.log'

    [void](Invoke-Checked $keygen @('-q', '-t', 'ed25519', '-N', '', '-C', $hostAlias, '-f', $clientKey))
    [void](Invoke-Checked $keygen @('-q', '-t', 'ed25519', '-N', '', '-C', $hostAlias, '-f', $hostKey))

    # Lock down the keys and authorize the client key for login
    Set-ClientKeyAcl $clientKey
    Set-SystemAdminAcl $hostKey
    Write-Utf8 $authorizedKeys ([IO.File]::ReadAllText("$clientKey.pub"))
    Set-SystemAdminAcl $authorizedKeys

    # Create the known-hosts file for the client, and get the host fingerprint for display
    $hostPublicKey = ([IO.File]::ReadAllText("$hostKey.pub").Trim() -split '\s+')[0..1] -join ' '
    Write-Utf8 $knownHosts "$hostAlias $hostPublicKey`n"
    $fingerprint = (Invoke-Checked $keygen @('-l', '-E', 'sha256', '-f', "$hostKey.pub")).Trim()

    # --- Create the temporary administrator account ---
    # The random password is only needed to create the account. It is never shown or used, since login is key-only.
    $passwordBytes = New-Object byte[] 32
    $rng = [Security.Cryptography.RandomNumberGenerator]::Create()
    try { $rng.GetBytes($passwordBytes) } finally { $rng.Dispose() }
    $password = ConvertTo-SecureString ('aA1!' + [Convert]::ToBase64String($passwordBytes)) -AsPlainText -Force
    try {
        $account = New-LocalUser -Name $userName -Password $password -PasswordNeverExpires -UserMayNotChangePassword -Description 'Disposable AI Windows Sandbox SSH account'
        $createdUser = $true
    } finally { $password.Dispose(); [Array]::Clear($passwordBytes, 0, $passwordBytes.Length) }
    Add-LocalGroupMember -SID ([Security.Principal.SecurityIdentifier]'S-1-5-32-544') -Member $account

    # --- Configure and start the SSH server ---
    # SSH can use a separate config path, but Win32-OpenSSH also expects its
    # standard ProgramData directory for its Windows-specific runtime files.
    if (-not (Test-Path -LiteralPath $sshData)) {
        [void](New-Item -Path $sshData -ItemType Directory)
        $createdSshData = $true
        Set-SystemAdminAcl -Path $sshData -Directory
        [void](Invoke-Checked $icacls @($sshData, '/grant', '*S-1-5-11:(RX)', '/Q'))
    }

    $sshLogDirectory = Join-Path $sshData 'logs'
    if (-not (Test-Path -LiteralPath $sshLogDirectory)) {
        [void](New-Item -Path $sshLogDirectory -ItemType Directory)
        Set-SystemAdminAcl -Path $sshLogDirectory -Directory
    }

    # Set Windows PowerShell as the SSH default shell, saving any existing value so cleanup can restore it
    if (-not (Test-Path -LiteralPath $registryPath)) { [void](New-Item -Path $registryPath -Force) }
    $registryKey = Get-Item -LiteralPath $registryPath
    $hadShellValue = $registryKey.GetValueNames() -contains 'DefaultShell'
    if ($hadShellValue) {
        $oldShellValue = $registryKey.GetValue('DefaultShell', $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
        $oldShellKind = $registryKey.GetValueKind('DefaultShell').ToString()
    }
    New-ItemProperty -LiteralPath $registryPath -Name DefaultShell -Value $shellPath -PropertyType String -Force | Out-Null
    $changedShell = $true

    # Write sshd_config: loopback-only listener, public key login only, no forwarding
    $hostKeyForConfig = $hostKey.Replace('\', '/')
    $authForConfig = $authorizedKeys.Replace('\', '/')
    $sftpForConfig = (Join-Path $bin 'sftp-server.exe').Replace('\', '/')
    
    Write-Utf8 $serverConfig @"
Port $Port
AddressFamily inet
ListenAddress 127.0.0.1
HostKey "$hostKeyForConfig"
AuthorizedKeysFile "$authForConfig"
AllowUsers $userName
AuthenticationMethods publickey
PubkeyAuthentication yes
PasswordAuthentication no
KbdInteractiveAuthentication no
PermitEmptyPasswords no
StrictModes yes
LoginGraceTime 30
MaxAuthTries 3
MaxStartups 3:30:10
DisableForwarding yes
PermitTTY yes
SyslogFacility LOCAL0
LogLevel VERBOSE
Subsystem sftp "$sftpForConfig"
"@

    # Lock down and validate the config
    Set-SystemAdminAcl $serverConfig
    [void](Invoke-Checked $icacls @($serverConfig, '/grant', '*S-1-5-11:(R)', '/Q'))
    [void](Invoke-Checked $sshd @('-t', '-f', $serverConfig))

    # Register sshd as a service and start it
    $serviceCommand = (ConvertTo-NativeArgument $sshd) + ' -f ' + (ConvertTo-NativeArgument $serverConfig)
    New-Service -Name sshd -BinaryPathName $serviceCommand -DisplayName 'Disposable AI Sandbox SSH' -StartupType Manual | Out-Null
    $createdService = $true
    # Same required privileges used by Microsoft's install-sshd.ps1.
    [void](Invoke-Checked (Join-Path $env:WINDIR 'System32\sc.exe') @('privs', 'sshd', 'SeAssignPrimaryTokenPrivilege/SeTcbPrivilege/SeBackupPrivilege/SeRestorePrivilege/SeImpersonatePrivilege'))
    Start-Service -Name sshd

    # --- Optional: test SSH login locally (only with -TestConnection) ---
    if ($TestConnection) {
        $commonSshArgs = @('-F', 'NUL', '-i', $clientKey, '-o', 'BatchMode=yes', '-o', 'IdentitiesOnly=yes', '-o', 'StrictHostKeyChecking=yes', '-o', "UserKnownHostsFile=$($knownHosts.Replace('\', '/'))", '-o', "HostKeyAlias=$hostAlias", '-o', 'ConnectTimeout=8', '-o', 'ConnectionAttempts=1', '-T')
        Write-Host 'Testing the local SSH login and PowerShell shell...'
        $localTestPassed = $false
        $lastTestError = ''
        $deadline = [DateTime]::UtcNow.AddSeconds(45)
        while ([DateTime]::UtcNow -lt $deadline) {
            try {
                $localResult = Invoke-Native $ssh ($commonSshArgs + @('-vvv', '-p', "$Port", "$userName@127.0.0.1", "Write-Output 'AI_SANDBOX_READY'")) 20
                Write-Utf8 (Join-Path $root 'ssh-local-client.log') ($localResult.Output + $localResult.Error)
                if ($localResult.ExitCode -eq 0 -and $localResult.Output -match 'AI_SANDBOX_READY') { $localTestPassed = $true; break }
                $lastTestError = ($localResult.Error.TrimEnd() -split '\r?\n')[-1]
            } catch { $lastTestError = $_.Exception.Message }
            Start-Sleep -Seconds 1
        }
        if (-not $localTestPassed) { throw "Local SSH login test failed. $lastTestError" }
    }

    # --- Start the Cloudflare Quick Tunnel ---
    Write-Host "Opening an outbound Cloudflare Quick Tunnel for $SessionMinutes minutes.`n    (Note: Duration set via the SessionMinutes parameter, or change its default value in the script.)" 
    $tunnelLog = Join-Path $root 'cloudflared.log'
    $tunnelJob = New-Object AiSandbox.KillOnCloseJob
    $tunnelArguments = @('tunnel', '--no-autoupdate', '--protocol', 'http2', '--url', "ssh://127.0.0.1:$Port", '--metrics', '127.0.0.1:0', '--loglevel', 'info', '--logfile', $tunnelLog)
    $tunnel = Start-Process -FilePath $cloudflared -ArgumentList (Join-NativeArguments $tunnelArguments) -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $root 'tunnel-stdout.log') -RedirectStandardError (Join-Path $root 'tunnel-stderr.log')
    $tunnelJob.Add($tunnel.Handle)

    # Wait up to 90 seconds for cloudflared to log the assigned trycloudflare.com hostname
    $tunnelHost = $null
    $lastTunnelLogError = $null
    $deadline = [DateTime]::UtcNow.AddSeconds(90)
    while ([DateTime]::UtcNow -lt $deadline) {
        if ($tunnel.HasExited) { throw 'cloudflared exited before a tunnel was established. See tunnel-stderr.log.' }
        if (Test-Path -LiteralPath $tunnelLog) {
            try {
                $logText = Read-SharedText $tunnelLog
                $lastTunnelLogError = $null
                $match = [regex]::Match($logText, 'https://([a-z0-9]+(?:-[a-z0-9]+)*\.trycloudflare\.com)')
                if ($match.Success) { $tunnelHost = $match.Groups[1].Value; break }
            } catch [IO.IOException] {
                $lastTunnelLogError = $_.Exception.GetBaseException().Message
            }
        }
        Start-Sleep -Milliseconds 500
    }
    if (-not $tunnelHost) {
        if ($lastTunnelLogError) { throw "Could not read the Cloudflare log within 90 seconds: $lastTunnelLogError" }
        throw 'Cloudflare did not return a Quick Tunnel hostname within 90 seconds.'
    }

    # --- Optional: test SSH through the public tunnel (only with -TestConnection) ---
    if ($TestConnection) {
        Write-Host 'Testing SSH through the public tunnel, including the pinned host key (can take about two minutes)...'
        $proxyCommand = 'ProxyCommand="' + $cloudflared.Replace('\', '/') + '" access ssh --hostname %h'
        $lastTestError = ''
        $deadline = [DateTime]::UtcNow.AddSeconds(120)
        $publicTestPassed = $false
        while ([DateTime]::UtcNow -lt $deadline) {
            if ($tunnel.HasExited) { throw 'cloudflared exited while testing the public route.' }
            try {
                $result = Invoke-Native $ssh ($commonSshArgs + @('-o', $proxyCommand, "$userName@$tunnelHost", "Write-Output 'AI_SANDBOX_READY'")) 20
                if ($result.ExitCode -eq 0 -and $result.Output -match 'AI_SANDBOX_READY') { $publicTestPassed = $true; break }
                $lastTestError = $result.Error
            } catch { $lastTestError = $_.Exception.Message }
            Start-Sleep -Seconds 2
        }
        if (-not $publicTestPassed) { throw "Public-route SSH test failed. $lastTestError" }
    } else {
        Write-Verbose "`nSSH connection tests skipped. Use -TestConnection to enable them."
    }

    # --- Build the content of the client connection files ---
    $expires = [DateTime]::UtcNow.AddMinutes($SessionMinutes)

    # SSH client config (sandbox-ssh.conf)
    $clientConfig = @"
Host ai-sandbox
    HostName $tunnelHost
    User $userName
    Port 22
    IdentityFile ./sandbox-key
    IdentitiesOnly yes
    BatchMode yes
    UserKnownHostsFile ./sandbox-known-hosts
    HostKeyAlias $hostAlias
    StrictHostKeyChecking yes
    ServerAliveInterval 15
    ServerAliveCountMax 3
    ProxyCommand cloudflared access ssh --hostname %h
"@

    # INSTRUCTIONS.MD for the AI: this common part, plus the Linux or Windows part below
    $sessionInstructions = @"
# My Windows Sandbox connection

I am the AI agent connecting to the user's Windows Sandbox. I extract this archive into a private directory in my execution environment and use the supplied files directly. I do not recreate the keys or print their contents. My remote shell is Windows PowerShell 5.1, regardless of my local operating system.

- Remote hostname: $tunnelHost
- SSH user: $userName
- Expires UTC: $($expires.ToString('yyyy-MM-dd HH:mm:ss'))
- SSH host fingerprint: $fingerprint
- Bundled cloudflared version: $($cloudflaredRelease.Version)
- Remote access: administrator inside Windows Sandbox.

I use sandbox-key for authentication, sandbox-known-hosts to verify the server identity, and sandbox-ssh.conf for the complete connection settings. I keep these files together and use only files from this same archive.

I connect through the configuration's Cloudflare WebSocket/HTTPS proxy; the hostname is not a public TCP port 22. I do not need a Cloudflare account or browser login. My environment must permit the outbound connection. If it is denied, I report the exact hostname and error so the user can adjust their allowlist. I keep host-key verification enabled. If the newly created route is not ready, I can retry briefly. I remove my local credential copies when finished.

"@

    # Linux INSTRUCTIONS.MD
    $linuxInstructions = $sessionInstructions + @'
## My Linux setup

I use the included CONNECT.sh with Bash. It selects and verifies the bundled x86_64 or ARM64 cloudflared binary, protects my key files, configures its own working directory and PATH, and connects. I do not need to download cloudflared.

I run this from the extracted directory to prepare my client and execute a harmless remote identity check:

```bash
bash ./CONNECT.sh
```

If either ssh or sftp is missing, CONNECT.sh automatically attempts to install the OpenSSH client with the detected package manager. It uses my existing root privileges or passwordless sudo, without prompting. Package installation needs access to the distribution's package repositories. Bash, sha256sum, and standard Linux utilities are prerequisites; the helper does not install those.

| My distribution/package manager | Package the helper installs |
| --- | --- |
| Debian/Ubuntu, apt-get | openssh-client |
| Fedora/RHEL family, dnf or yum | openssh-clients |

If automatic installation fails, I inspect its error and use the matching manual command below when my environment permits it. I run these as root, or prefix each privileged command with sudo -n when passwordless sudo is available. I do not disable repository verification or bypass my environment's access restrictions. If my distribution uses a different package manager, I install its OpenSSH client package through that distribution's supported tools.

Debian/Ubuntu:

```bash
apt-get update && DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends openssh-client
```

Fedora/RHEL with dnf:

```bash
dnf install -y openssh-clients
```

Older RHEL-family systems with yum:

```bash
yum install -y openssh-clients
```

I then rerun CONNECT.sh. It checks that both ssh and sftp are available before connecting. Its status output explains permissions, package installation when needed, and binary verification. It sets PATH only for itself and child processes; repeated runs reuse the binaries and check their checksum again.

To run a headless PowerShell command, I use:

```bash
bash ./CONNECT.sh 'Get-ComputerInfo | Select-Object WindowsProductName, OsVersion'
```

When I need an interactive terminal and my tools support one, I use:

```bash
bash ./CONNECT.sh --shell
```

For file transfer, I use:

```bash
bash ./CONNECT.sh --sftp
```

I run connection commands through CONNECT.sh so its local PATH and working directory are correct. I write remote commands for Windows PowerShell 5.1. In Bash, I single-quote remote PowerShell commands to keep dollar signs from expanding locally. In SFTP, I use its normal put/get commands.
'@

    # Windows INSTRUCTIONS.MD
    $windowsInstructions = $sessionInstructions + @'
## My Windows setup

I use the bundled portable x64 cloudflared.exe from Windows PowerShell 5.1 or newer on an x64 Windows client. I need ssh and sftp on PATH. I run the commands below from the extracted directory, in the same PowerShell session. I do not need to download or install cloudflared.

I check that OpenSSH is available:

```powershell
Get-Command ssh, sftp -ErrorAction Stop
```

If OpenSSH is missing, I install its client component or obtain the official portable Microsoft Win32-OpenSSH client before continuing. This Windows archive does not automatically install OpenSSH: https://github.com/PowerShell/Win32-OpenSSH/releases

I verify the bundled cloudflared executable before running it:

```powershell
if ((Get-FileHash -LiteralPath .\cloudflared.exe -Algorithm SHA256).Hash -ine '__CLOUDFLARED_WINDOWS_SHA256__') { throw 'Bundled cloudflared SHA256 mismatch. Do not run this executable.' }
```

I restrict the private key to my current client account:

```powershell
$p = (Resolve-Path .\sandbox-key).Path; $s = [Security.Principal.WindowsIdentity]::GetCurrent().User; $a = New-Object Security.AccessControl.FileSecurity; $a.SetOwner($s); $a.SetAccessRuleProtection($true, $false); $a.AddAccessRule((New-Object Security.AccessControl.FileSystemAccessRule($s, 'FullControl', 'Allow'))); Set-Acl -LiteralPath $p -AclObject $a
```

I add the extracted directory to PATH for this PowerShell process and its children:

```powershell
$env:PATH = (Get-Location).Path + ';' + $env:PATH
```

I verify the connection with a harmless remote identity check:

```powershell
ssh -F ./sandbox-ssh.conf -T ai-sandbox "Write-Output 'AI_SANDBOX_CONNECTED'; whoami; hostname"
```

To run a headless PowerShell command, I use:

```powershell
ssh -F ./sandbox-ssh.conf -T ai-sandbox "Get-ComputerInfo | Select-Object WindowsProductName, OsVersion"
```

For an interactive PowerShell session, I use:

```powershell
ssh -F ./sandbox-ssh.conf ai-sandbox
```

For file transfer, I use:

```powershell
sftp -F ./sandbox-ssh.conf ai-sandbox
```

I keep the extracted directory as my working directory when connecting. I write remote commands for Windows PowerShell 5.1 and use normal SFTP put/get commands for file transfer.
'@
    $windowsInstructions = $windowsInstructions.Replace('__CLOUDFLARED_WINDOWS_SHA256__', $cloudflaredAssets['cloudflared-windows-amd64.exe'].Sha256)

    # CONNECT.sh for Linux clients: checks files, verifies cloudflared, installs OpenSSH if missing, then connects
    $linuxConnector = @'
#!/usr/bin/env bash
set -euo pipefail
umask 077
cd -- "$(dirname -- "${BASH_SOURCE[0]}")"

if [[ "${1:-}" == --help ]]; then
    printf '%s\n' 'Prepare this Linux client and check the remote identity:' '  bash ./CONNECT.sh' 'Run a PowerShell command:' "  bash ./CONNECT.sh 'Get-ComputerInfo | Select-Object WindowsProductName, OsVersion'" 'Open a terminal:' '  bash ./CONNECT.sh --shell' 'Transfer files:' '  bash ./CONNECT.sh --sftp'
    exit 0
fi
for required in sha256sum uname chmod ln; do
    command -v "$required" >/dev/null 2>&1 || { printf 'Required client tool is missing: %s\n' "$required" >&2; exit 1; }
done
for file in sandbox-key sandbox-known-hosts sandbox-ssh.conf; do
    [[ -f "$file" ]] || { printf 'Required client file is missing: %s\n' "$file" >&2; exit 1; }
done
[[ "$(uname -s)" == Linux ]] || { printf 'Use this archive on Linux.\n' >&2; exit 1; }
case "$(uname -m)" in
    x86_64) cf_binary=cloudflared-linux-amd64; cf_sha=__CLOUDFLARED_AMD64_SHA256__ ;;
    aarch64|arm64) cf_binary=cloudflared-linux-arm64; cf_sha=__CLOUDFLARED_ARM64_SHA256__ ;;
    *) printf 'This archive supports x86_64 and ARM64 Linux.\n' >&2; exit 1 ;;
esac
chmod 700 .
chmod 600 sandbox-key sandbox-known-hosts sandbox-ssh.conf
printf 'Protected the client directory and SSH credential files.\n'
[[ -f "$cf_binary" ]] || { printf 'Bundled binary is missing: %s\n' "$cf_binary" >&2; exit 1; }
if ! printf '%s  %s\n' "$cf_sha" "$cf_binary" | sha256sum --check --status; then
    printf 'Bundled cloudflared SHA256 mismatch. Nothing was executed.\n' >&2
    exit 1
fi
ensure_ssh_client() {
    if command -v ssh >/dev/null 2>&1 && command -v sftp >/dev/null 2>&1; then
        return 0
    fi
    local -a elevate=()
    if (( EUID != 0 )); then
        if ! command -v sudo >/dev/null 2>&1 || ! sudo -n true; then
            printf 'OpenSSH is missing; installation needs root or passwordless sudo. See INSTRUCTIONS.MD for packages and manual commands.\n' >&2
            return 1
        fi
        elevate=(sudo -n)
    fi
    if command -v apt-get >/dev/null 2>&1; then
        printf 'Installing openssh-client with apt-get because ssh or sftp is missing.\n'
        if ! "${elevate[@]}" apt-get update || ! "${elevate[@]}" env DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends openssh-client; then
            printf 'OpenSSH installation failed. See INSTRUCTIONS.MD for the manual apt-get command.\n' >&2
            return 1
        fi
    elif command -v dnf >/dev/null 2>&1; then
        printf 'Installing openssh-clients with dnf because ssh or sftp is missing.\n'
        if ! "${elevate[@]}" dnf install -y openssh-clients; then
            printf 'OpenSSH installation failed. See INSTRUCTIONS.MD for the manual dnf command.\n' >&2
            return 1
        fi
    elif command -v yum >/dev/null 2>&1; then
        printf 'Installing openssh-clients with yum because ssh or sftp is missing.\n'
        if ! "${elevate[@]}" yum install -y openssh-clients; then
            printf 'OpenSSH installation failed. See INSTRUCTIONS.MD for the manual yum command.\n' >&2
            return 1
        fi
    else
        printf 'No supported package manager found. Install the OpenSSH client for this distribution; see INSTRUCTIONS.MD.\n' >&2
        return 1
    fi
    hash -r
    if ! command -v ssh >/dev/null 2>&1 || ! command -v sftp >/dev/null 2>&1; then
        printf 'Installation finished but ssh or sftp is still unavailable on PATH. See INSTRUCTIONS.MD.\n' >&2
        return 1
    fi
}
ensure_ssh_client
chmod 700 "$cf_binary"
ln -sfn -- "$cf_binary" ./cloudflared
export PATH="$PWD:$PATH"
printf 'Using bundled cloudflared __CLOUDFLARED_VERSION__ (%s); SHA256 verified.\n' "$cf_binary"
printf 'Connecting with the supplied SSH key and pinned server identity.\n'
case "${1:-}" in
    --shell) exec ssh -F ./sandbox-ssh.conf ai-sandbox ;;
    --sftp) command -v sftp >/dev/null 2>&1 || { printf 'Required client tool is missing: sftp\n' >&2; exit 1; }; exec sftp -F ./sandbox-ssh.conf ai-sandbox ;;
    '') exec ssh -F ./sandbox-ssh.conf -T ai-sandbox "Write-Output 'AI_SANDBOX_CONNECTED'; whoami; hostname" ;;
    *) exec ssh -F ./sandbox-ssh.conf -T ai-sandbox "$@" ;;
esac
'@
    $linuxConnector = $linuxConnector.Replace('__CLOUDFLARED_AMD64_SHA256__', $cloudflaredAssets['cloudflared-linux-amd64'].Sha256).Replace('__CLOUDFLARED_ARM64_SHA256__', $cloudflaredAssets['cloudflared-linux-arm64'].Sha256).Replace('__CLOUDFLARED_VERSION__', $cloudflaredRelease.Version)

    # --- Write the output files to the connection folder ---
    if (-not (Test-Path -LiteralPath $connectionDirectory -PathType Container)) {
        [void](New-Item -Path $connectionDirectory -ItemType Directory)
        Set-SystemAdminAcl -Path $connectionDirectory -Directory
    }

    # User-Instructions.txt (for the human)
    $userInstructionsPath = Join-Path $connectionDirectory $userInstructionsFileName
    $userInstructions = @"
Instructions and tips

This session expires at:  $($expires.ToString('yyyy-MM-dd HH:mm:ss')) UTC.
----------------------------------------------------------

USER INSTRUCTIONS (For you the human):

    1. Upload the ZIP for the AI's execution environment:
        - SANDBOX-CONNECT-INSTRUCTIONS-LINUX.zip for a Linux AI client.
        - SANDBOX-CONNECT-INSTRUCTIONS-WINDOWS.zip for a Windows x64 AI client.
    
        (The above means the OS environment your agent is running on. It has nothing to do with this Sandbox, which obviously runs Windows.)

    2. Ask the AI to extract its ZIP and read INSTRUCTIONS.MD (usually it figures that out on its own). The Linux client can run bash ./CONNECT.sh; that prepares the connection and attempts to install OpenSSH if missing.

    3. If the AI has a network allowlist, allow this session's hostname:
        $tunnelHost

    4. Keep this server PowerShell window open while the AI is connected. Ctrl+C stops access. Closing Windows Sandbox destroys the guest.

----------------------------------------------------------

TIPS: You can customize the Sandbox configuration using a Windows Sandbox configuration file (.wsb):

    - Enable the "protected client" mode using:   <ProtectedClient>Enable</ProtectedClient>
        - I haven't found any drawbacks to using this, so may as well enable it.

    - Enable GPU passthrough in the .wsb using:   <vGPU>Enable</vGPU>
        - This way your AI can utilize hardware acceleration via this sandbox.
        - Note: There is a bug where the sandbox won't work if you enable vGPU and have multiple GPUs (including a dedicated card + an integrated GPU). Disable one of the GPUs in the BIOS or device manager if you want to enable GPU passthrough.

    - Provide the Sandbox with more memory using:  <MemoryInMB>16384</MemoryInMB>   (adjust the value as needed for your system)

    - More Info: https://learn.microsoft.com/en-us/windows/security/application-security/application-isolation/windows-sandbox/windows-sandbox-configure-using-wsb-file

----------------------------------------------------------

Other Notes:
    - The AI receives administrator access inside Windows Sandbox.
    - Both ZIPs contain the private login key. Share them only with the intended AI.
        - Their originals are removed when this session stops. If you need to keep a copy, do so before the session ends.
        - This User-Instructions.txt file remains for reference.
    - Each new server run replaces these output files and generates new credentials and a new hostname.

The server script downloads the latest stable cloudflared release and packages verified Windows x64 and Linux x64/ARM64 binaries. The AI does not need to download cloudflared. Linux OpenSSH installation needs root or passwordless sudo and access to package repositories; the agent's INSTRUCTIONS.MD lists the exact packages and manual commands if automatic setup fails.
"@
    Write-Utf8 $userInstructionsPath ($userInstructions.TrimEnd() + "`n")

    # Build the Linux and Windows ZIPs in a staging folder, then copy them to the connection folder
    $clientBundleDirectory = Join-Path $root 'client-packages'
    [void](New-Item -Path $clientBundleDirectory -ItemType Directory)
    Set-SystemAdminAcl -Path $clientBundleDirectory -Directory
    foreach ($platform in @('LINUX', 'WINDOWS')) {
        $clientDirectory = Join-Path $clientBundleDirectory $platform
        [void](New-Item -Path $clientDirectory -ItemType Directory)
        # Files included for both platforms
        Copy-Item -LiteralPath $clientKey -Destination (Join-Path $clientDirectory 'sandbox-key')
        Copy-Item -LiteralPath $knownHosts -Destination (Join-Path $clientDirectory 'sandbox-known-hosts')
        Write-Utf8 (Join-Path $clientDirectory 'sandbox-ssh.conf') ($clientConfig.TrimEnd() + "`n")

        # Platform-specific files
        if ($platform -eq 'LINUX') {
            foreach ($name in @('cloudflared-linux-amd64', 'cloudflared-linux-arm64')) {
                Copy-Item -LiteralPath (Join-Path $root $name) -Destination (Join-Path $clientDirectory $name)
            }
            Write-Utf8 (Join-Path $clientDirectory 'CONNECT.sh') ($linuxConnector.Replace("`r`n", "`n").TrimEnd() + "`n")
            Write-Utf8 (Join-Path $clientDirectory 'INSTRUCTIONS.MD') ($linuxInstructions.TrimEnd() + "`n")
        } else {
            Copy-Item -LiteralPath $cloudflared -Destination (Join-Path $clientDirectory 'cloudflared.exe')
            Write-Utf8 (Join-Path $clientDirectory 'INSTRUCTIONS.MD') ($windowsInstructions.TrimEnd() + "`n")
        }

        # Zip the files
        $zipName = "SANDBOX-CONNECT-INSTRUCTIONS-$platform.zip"
        $stagedZip = Join-Path $clientBundleDirectory $zipName
        Compress-Archive -LiteralPath @((Get-ChildItem -LiteralPath $clientDirectory -File).FullName) -DestinationPath $stagedZip -CompressionLevel Optimal

        # Create the destination file empty and restrict its permissions first, then write the ZIP contents into it
        $destination = Join-Path $connectionDirectory $zipName
        try { [IO.File]::WriteAllBytes($destination, [byte[]]@()) } catch { throw "Cannot save the connection ZIP in Sandbox-Connect-Files beside this script. Run it from a writable folder inside Windows Sandbox. $($_.Exception.Message)" }
        $connectionFiles += $destination
        Set-ClientKeyAcl $destination
        [IO.File]::WriteAllBytes($destination, [IO.File]::ReadAllBytes($stagedZip))
    }
    Remove-Item -LiteralPath $clientBundleDirectory -Recurse -Force

    # Print important connection information
    Write-Host "`n`n---------------------------------------------------------------------------------------" -ForegroundColor Green
    Write-Host "   HOST TO WHITELIST:      " -NoNewline
    Write-Host "$tunnelHost" -ForegroundColor Green
    Write-Host "   CONNECTION VALID FOR:   " -NoNewline
    Write-Host "$SessionMinutes minutes" -ForegroundColor Green -NoNewline
    Write-Host "`n---------------------------------------------------------------------------------------" -ForegroundColor Green
    Write-Host ''

    # Prints the step telling the user where the connection folder is
    function Write-FolderLocation {
        Write-Host ""

        if ($connectionBaseDirectory -eq [Environment]::GetFolderPath('Desktop')) {
            Write-Host "`t2. On the " -NoNewline
            Write-Host 'Desktop' -ForegroundColor Yellow -NoNewline
            Write-Host " open the " -NoNewline
            Write-Host $connectionFolderName -ForegroundColor Yellow -NoNewline
            Write-Host " folder."
        } else {
            Write-Host "`t2. Open the connection folder at: " -NoNewline
            Write-Host $connectionDirectory -ForegroundColor Yellow -NoNewline
        }
    }

    Write-Host " How To Let Your Agent Connect:`n"
    Write-Host "`t1. Whitelist the cloudflared host (URL) above for your agent, if necessary. (Remove it when done)"
    Write-FolderLocation
    Write-Host "`n`t3. Provide the correct ZIP file to your agent, depending on your agent's OS (Linux or Windows)."
    Write-Host "`n`n`t-- For more details and tips, see: " $userInstructionsFileName
    
    # Full file paths (only shown with -Verbose)
    Write-Verbose "`nConnection File Paths (choose the AI client operating system):"
    $connectionFiles | ForEach-Object { Write-Verbose "`t$_" }
    Write-Verbose "`n`t$userInstructionsPath"

    Write-Host "`n`n---------------------------------------------------------------------------------------"
    Write-Host "`n Ready. Keep this window open. Ctrl+C stops access. Auto-stop in $SessionMinutes minutes.`n" -ForegroundColor Cyan

    Write-Host "`n Advanced Connection Details:"
    Write-Host "    SSH user: $userName | Expires UTC: $($expires.ToString('yyyy-MM-dd HH:mm:ss'))"
    Write-Host "    SSH host fingerprint: $fingerprint"
    Write-Host "`n"
    
    Write-Host "---------------------------------------------------------------------------------------"
    Write-Host ' Close Windows Sandbox itself when finished to destroy all guest state.'

    # --- Setup succeeded. Keep access open until expiry, ending early if the tunnel or SSH service stops ---
    $ready = $true
    while ([DateTime]::UtcNow -lt $expires) {
        if ($tunnel.HasExited) { throw 'The Cloudflare tunnel stopped. Access is being torn down.' }
        if ((Get-Service -Name sshd).Status -ne 'Running') { throw 'The SSH service stopped. Access is being torn down.' }
        Start-Sleep -Milliseconds 500
    }
    Write-Host 'Session expired.' -ForegroundColor Yellow
} catch {
    # --- On failure: stop SSH and print recent log lines for diagnosis ---
    $sessionError = $_
    Write-Host "Setup/session failed: $($sessionError.Exception.Message)" -ForegroundColor Red
    Write-Host "Diagnostics remain inside this sandbox at: $root"
    if ($createdService) {
        try { Stop-Service -Name sshd -Force -ErrorAction Stop } catch { Write-Warning "Could not stop SSH before reading logs: $($_.Exception.Message)" }
    }

    # Show the last 30 lines of each log, and copy the SSH logs into the session folder
    $diagnosticPaths = @((Join-Path $sshData 'logs\sshd.log'), (Join-Path $sshData 'logs\sshd-session.log'), (Join-Path $sshData 'logs\sshd-auth.log'), (Join-Path $root 'ssh-local-client.log'), (Join-Path $root 'tunnel-stderr.log'))
    foreach ($logPath in $diagnosticPaths) {
        if (Test-Path -LiteralPath $logPath) {
            try {
                $logName = [IO.Path]::GetFileName($logPath)
                $lines = @(Get-Content -LiteralPath $logPath -Tail 30 -ErrorAction Stop)
                if ($lines.Count) {
                    Write-Host "Recent $logName entries:"
                    $lines | ForEach-Object { Write-Host $_ }
                    if ([IO.Path]::GetDirectoryName($logPath) -ne $root) {
                        Write-Utf8 (Join-Path $root $logName) ($lines -join "`r`n")
                    }
                }
            } catch { Write-Warning "Could not read diagnostic log ${logPath}: $($_.Exception.Message)" }
        }
    }
    throw $sessionError
} finally {
    # --- Cleanup: runs on success, failure, or Ctrl+C ---
    # Revoke the Internet route first, then the account and server.
    if ($tunnelJob) { $tunnelJob.Dispose() }
    if ($tunnel) {
        try { if (-not $tunnel.HasExited) { $tunnel.Kill(); [void]$tunnel.WaitForExit(5000) } } catch { Write-Warning $_.Exception.Message }
        $tunnel.Dispose()
    }

    # Disable the account right away. It is removed after the SSH service is deleted.
    if ($createdUser) {
        try { Disable-LocalUser -Name $userName } catch { Write-Warning $_.Exception.Message }
    }

    if ($createdService) {
        try {
            Stop-Service -Name sshd -Force -ErrorAction SilentlyContinue
            [void](Invoke-Checked (Join-Path $env:WINDIR 'System32\sc.exe') @('delete', 'sshd'))
        } catch { Write-Warning "Could not remove the temporary SSH service: $($_.Exception.Message)" }
    }

    if ($createdUser) {
        try { Remove-LocalUser -Name $userName } catch { Write-Warning $_.Exception.Message }
    }

    # Restore the previous DefaultShell registry value
    if ($changedShell) {
        try {
            if ($hadShellValue) {
                New-ItemProperty -LiteralPath $registryPath -Name DefaultShell -Value $oldShellValue -PropertyType $oldShellKind -Force | Out-Null
            } else { Remove-ItemProperty -LiteralPath $registryPath -Name DefaultShell -ErrorAction SilentlyContinue }
        } catch { Write-Warning $_.Exception.Message }
    }

    # Erase credentials even if setup failed; leave diagnostic logs on failure.
    foreach ($secretName in @('sandbox-key', 'sandbox-key.pub', 'authorized_keys', 'ssh_host_ed25519_key')) {
        Remove-Item -LiteralPath (Join-Path $root $secretName) -Force -ErrorAction SilentlyContinue
    }
    foreach ($connectionFile in $connectionFiles) { Remove-Item -LiteralPath $connectionFile -Force -ErrorAction SilentlyContinue }
    if ($clientBundleDirectory) { Remove-Item -LiteralPath $clientBundleDirectory -Recurse -Force -ErrorAction SilentlyContinue }
    if ($ready) { Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue }
    if ($createdSshData) {
        # The guest may have acquired other software while connected. Only
        # remove this shared standard directory if it is still empty.
        if ((Test-Path -LiteralPath $sshData) -and -not (Get-ChildItem -LiteralPath $sshData -Force)) {
            Remove-Item -LiteralPath $sshData -Force -ErrorAction SilentlyContinue
        }
    }

    [Net.ServicePointManager]::SecurityProtocol = $oldTls
    Write-Host 'Remote access stopped. Close Windows Sandbox to discard the guest.' -ForegroundColor Cyan
}
