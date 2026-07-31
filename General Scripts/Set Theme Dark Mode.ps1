<#
.SYNOPSIS
Applies the Windows Dark Mode theme and wallpaper.

.PARAMETER AutoRange
Optional. A 24-hour time range string indicating when dark mode should be applied.
Format: "HH:MM-HH:MM" (e.g., "18:00-06:00" for 6 PM to 6 AM).
If the current time is outside this range, the script will exit without making changes.
If this parameter is omitted, the theme is applied immediately.

You can put this in your .wsb config file or run it from another .ps1 script.

.EXAMPLE
.\Set Theme Dark Mode.ps1 -AutoRange "18:00-06:00"
#>
param(
    [string]$AutoRange
)

if (-not [string]::IsNullOrWhiteSpace($AutoRange)) {
    $times = $AutoRange -split '-'
    if ($times.Count -eq 2) {
        $start = [timespan]$times[0]
        $end = [timespan]$times[1]
        $now = (Get-Date).TimeOfDay
        
        $applyTheme = $false
        if ($start -lt $end) {
            # Range doesn't cross midnight (e.g., "08:00-17:00")
            if ($now -ge $start -and $now -le $end) { $applyTheme = $true }
        } else {
            # Range crosses midnight (e.g., "18:00-06:00")
            if ($now -ge $start -or $now -le $end) { $applyTheme = $true }
        }
        
        if (-not $applyTheme) {
            Write-Host "Current time is outside AutoRange ($AutoRange). Dark mode will not be applied."
            exit
        }
    }
}

# Enable Dark Mode for Apps
Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "AppsUseLightTheme" -Value 0

# Enable Dark Mode for System
Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "SystemUsesLightTheme" -Value 0

# Set the Wallpaper
$wallpaperPath = "C:\Windows\Web\Wallpaper\Windows\img19.jpg"
$code = @'
using System.Runtime.InteropServices;
public class Wallpaper {
    [DllImport("user32.dll", CharSet=CharSet.Auto)]
    public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
}
'@

Add-Type $code
$SPI_SETDESKWALLPAPER = 0x0014
$UPDATE_INI_FILE = 0x01
$SEND_CHANGE = 0x02

[Wallpaper]::SystemParametersInfo($SPI_SETDESKWALLPAPER, 0, $wallpaperPath, ($UPDATE_INI_FILE -bor $SEND_CHANGE))

# Restart Explorer to apply changes
Write-Host "Restarting Explorer..."
Stop-Process -Name explorer -Force
Start-Process explorer
Write-Host "Dark mode enabled and wallpaper updated successfully! Explorer has been restarted."
