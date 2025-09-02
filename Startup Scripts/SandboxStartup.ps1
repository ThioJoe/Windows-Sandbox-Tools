
param(
    # Can include this switch when running from the .wsb file to indicate it's the first launch of the sandbox
    # Useful if re-running this script within the sandbox as a test, but don't want certain parts to run again
    [switch]$launchingSandbox
)

#Set-PSDebug -Trace 1

# ------ Check that we're running in the Windows Sandbox ------
# This script is intended to be run from within the Windows Sandbox. We'll do a rudamentary check for if the current user is named "WDAGUtilityAccount"
if ($env:USERNAME -ne "WDAGUtilityAccount") {
    Write-host "`n`nERROR: This script is intended to be run from WITHIN the Windows Sandbox.`nIt appears you are running this from outside the sandbox.`n" -ForegroundColor Red
    Write-host "`nPress Enter to exit." -ForegroundColor Yellow
    Read-Host
    exit
}

# Change context menu to old style
reg.exe add "HKCU\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32" /f /ve

# Show file extensions
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "HideFileExt" /t REG_DWORD /d 0 /f

# Show hidden files
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "Hidden" /t REG_DWORD /d 1 /f

# Fix for slow MSI package install. See: https://github.com/microsoft/Windows-Sandbox/issues/68#issuecomment-2754867968
reg add "HKLM\SYSTEM\CurrentControlSet\Control\CI\Policy" /v "VerifiedAndReputablePolicyState" /t REG_DWORD /d 0 /f
CiTool.exe --refresh --json | Out-Null # Refreshes policy. Use json output param or else it will prompt for confirmation, even with Out-Null

# Change execution policy for powershell to allow running scripts. Normally it shows an error about a more specific policy (Process level Bypass policy), but it doesn't matter so we hide it via try/catch
try { Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope LocalMachine -ErrorAction Stop | Out-Null } catch {}

# -----------------------------------------------------------------------------------------

# ---- Add 'Open PowerShell Here' and 'Open CMD Here' to context menu -------
$powershellPath = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
$cmdPath = "C:\Windows\System32\cmd.exe"
Write-Host "`nAdding 'Open PowerShell/CMD Here' context menu options"
reg add "HKEY_CLASSES_ROOT\Directory\Background\shell\MyPowerShell" /ve /d "Open PowerShell Here" /f
reg add "HKEY_CLASSES_ROOT\Directory\Background\shell\MyPowerShell" /v "Icon" /t REG_SZ /d "$powershellPath,0" /f
reg add "HKEY_CLASSES_ROOT\Directory\Background\shell\MyPowerShell\command" /ve /d "powershell.exe -noexit -command Set-Location -literalPath '%V'" /f
reg add "HKEY_CLASSES_ROOT\Directory\Background\shell\Mycmd" /ve /d "Open CMD Here" /f
reg add "HKEY_CLASSES_ROOT\Directory\Background\shell\Mycmd" /v "Icon" /t REG_SZ /d "$cmdPath,0" /f
reg add "HKEY_CLASSES_ROOT\Directory\Background\shell\Mycmd\command" /ve /d "cmd.exe /s /k cd /d `"\`"%V`"\`"" /f

# ---- Add File Types to Context Menu > New ----
# ShellNew Text Document - .txt
Write-host "`nAdding txt document new file option"
reg add "HKEY_CLASSES_ROOT\txtfile" /ve /d "Text Document" /f
reg add "HKEY_CLASSES_ROOT\.txt\ShellNew" /f
# Use --% to not have powershell parse the arguments, otherwise it won't pass the empty string for the /d parameter
reg --% add "HKEY_CLASSES_ROOT\.txt\ShellNew" /v "NullFile" /t REG_SZ /d "" /f
reg add "HKEY_CLASSES_ROOT\.txt\ShellNew" /v "ItemName" /t REG_SZ /d "New Text Document" /f

# ShellNew PowerShell Script - .ps1 -- Also happens to make .ps1 scripts clickable to run because of the association with "ps1file"
Write-host "`nAdding PowerShell new file option"
reg add "HKEY_CLASSES_ROOT\.ps1" /ve /d "ps1file" /f
reg add "HKEY_CLASSES_ROOT\ps1file" /ve /d "PowerShell Script" /f
reg add "HKEY_CLASSES_ROOT\ps1file\DefaultIcon" /ve /d "%SystemRoot%\System32\imageres.dll,-5372" /f
reg add "HKEY_CLASSES_ROOT\.ps1\ShellNew" /ve /d "ps1file" /f
reg add "HKEY_CLASSES_ROOT\.ps1\ShellNew" /f
reg --% add "HKEY_CLASSES_ROOT\.ps1\ShellNew" /v "NullFile" /t REG_SZ /d "" /f
reg add "HKEY_CLASSES_ROOT\.ps1\ShellNew" /v "ItemName" /t REG_SZ /d "script" /f

# =============================== OPTIONAL - Notepad and Notepad++ ===============================

# If you've included Notepad and/or Notepad++ in the shared folder, set the path to them here. Since notepad isn't included in the sandbox for some reason
# If they aren't found, each step will be skipped, so you can run this script without them.

# NotePad Tip: Go to C:\Windows on your main computer and copy Notepad.exe, then copy notepad.exe.mui from your main language folder, such as C:\Windows\en-US
#    Important: Notepad.exe.mui can't simply go next to notepad.exe. You need to actually create the language folder (like en-US) again next to notepad.exe and put it in that. Otherwise notepad won't run.
$notepadPath = "C:\Users\WDAGUtilityAccount\Desktop\HostShared\notepad.exe"
# For Notepad++, use the portable version
$notepadPlusPlusPath = "C:\Users\WDAGUtilityAccount\Desktop\HostShared\Notepad++\Notepad++.exe"

# Check if the Notepad and Notepad++ paths exist, if not, set them to null
If (!(Test-Path $notepadPath)) { $notepadPath = $null; Write-Host "Notepad not found, context menu options will not be added." }
If (!(Test-Path $notepadPlusPlusPath)) { $notepadPlusPlusPath = $null; Write-Host "Notepad++ not found, context menu options will not be added." }

# ---- Add 'Edit With Notepad' and 'Open Notepad' to context menu (If Available) -------
If ($null -ne $notepadPath) {
	Write-Host "`nAdding Edit with notepad context menu options"
	reg add "HKEY_CLASSES_ROOT\*\shell\Edit with Notepad" /f
	reg add "HKEY_CLASSES_ROOT\*\shell\Edit with Notepad" /v "Icon" /t REG_SZ /d "$notepadPath,0" /f
	reg add "HKEY_CLASSES_ROOT\*\shell\Edit with Notepad\command" /ve /d "`"$notepadPath`" `"%1`"" /f
	reg add "HKEY_CLASSES_ROOT\Directory\Background\shell\Notepad" /ve /d "Open Notepad" /f
	reg add "HKEY_CLASSES_ROOT\Directory\Background\shell\Notepad" /v "Icon" /t REG_SZ /d "$notepadPath,0" /f
	reg add "HKEY_CLASSES_ROOT\Directory\Background\shell\Notepad\command" /ve /d "`"$notepadPath`"" /f
}

# ---- Add 'Edit With Notepad++' and 'Open Notepad++' to context menu (If Available) -------
If ($null -ne $notepadPlusPlusPath) {
	Write-Host "`nAdding Edit/Open with Notepad++ context menu options"
	reg add "HKEY_CLASSES_ROOT\*\shell\Edit with Notepad++" /f
	reg add "HKEY_CLASSES_ROOT\*\shell\Edit with Notepad++" /v "Icon" /t REG_SZ /d "$notepadPlusPlusPath,0" /f
	# Notepad++ needs the file path to be in quotes. To get the quotes to show up in registry, need to use two quotes, and need to escape both with powershell `"
	reg add "HKEY_CLASSES_ROOT\*\shell\Edit with Notepad++\command" /ve /t REG_EXPAND_SZ /d "`"$notepadPlusPlusPath`" -settingsDir=`"%appdata%`" `"`"%1`"`"" /f
	reg add "HKEY_CLASSES_ROOT\Directory\Background\shell\Notepad++" /ve /d "Open Notepad++" /f
	reg add "HKEY_CLASSES_ROOT\Directory\Background\shell\Notepad++" /v "Icon" /t REG_SZ /d "$notepadPlusPlusPath,0" /f
	reg add "HKEY_CLASSES_ROOT\Directory\Background\shell\Notepad++\command" /ve /t REG_EXPAND_SZ /d "`"$notepadPlusPlusPath`" -settingsDir=`"%appdata%`"" /f
}

# Set .txt files to open with Notepad, or Notepad++ if available
cmd /c assoc .txt=txtfile
If (($null -ne $notepadPath) -or ($null -ne $notepadPlusPlusPath)) {
	If (!(Test-Path 'HKLM:\SOFTWARE\Classes\txtfile\shell\open\command')) { New-Item -Path 'HKLM:\SOFTWARE\Classes\txtfile\shell\open\command' -Force }
	# Prefer Notepad++ if available, otherwise use Notepad
	If ($null -ne $notepadPlusPlusPath) { 
		Set-ItemProperty -Path 'HKLM:\SOFTWARE\Classes\txtfile\shell\open\command' -Name '(Default)' -Value ('"{0}" -settingsDir=%appdata% "%1"' -f $editorPath) -Type ExpandString -Force
	} Else {  # If Npp isn't available, condition above means we know notepadPath is still available
		cmd /c ftype txtfile=`"$notepadPath`" "%1"
	} 
}

# ================================ FINALIZATION ================================

# Restart Explorer so changes take effect
Stop-Process -Name explorer -Force

# Open an explorer window to the host-shared folder on first launch
if ($launchingSandbox) { Start-Process explorer.exe C:\Users\WDAGUtilityAccount\Desktop\HostShared }

# Uncomment to pause after running
#Read-Host "Pause"

