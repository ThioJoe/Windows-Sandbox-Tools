# Scripts & Tools for Windows Sandbox
Various useful scripts for use within Windows Sandbox

## 📥 Installer Scripts

PowerShell scripts that can install apps or components not normally included in the Sandbox

- `Install-Winget.ps1`: Fetches necessary files and dependencies from Microsoft's [winget-cli](https://github.com/microsoft/winget-cli) repo, and installs it
- `Install-Microsoft-Store.ps1`: Installs the Microsoft Store via the Windows Update APIs
  - Unlike other similar scripts, mine does NOT use any third party APIs like UUP Dump or RG-Adguard at all
  - (All web requests go directly to Microsoft servers. Even those fetching the download URLs in the first place)
- `Install VC Redist.ps1`: Installs each architecture for [the latest Microsoft Visual C++ Redistributable](https://learn.microsoft.com/en-us/cpp/windows/latest-supported-vc-redist?view=msvc-170#latest-microsoft-visual-c-redistributable-version) (2015-2022), which is commonly required by other software.
  - `Install VC Redist AIO.ps1`: An alternative script that installs every Visual C++ Redistributable via [VisualCppRedist AIO Pack](https://github.com/abbodi1406/vcredist) by abbodi1406. Recommended if you're frequently running older programs.

## 🛠️ General Scripts

- `Set Theme Dark Mode.ps1` - Sets the Sandbox to Dark theme and also changes to a dark wallpaper
- `Set Theme Light Mode.ps1` - Restores the Light theme. Doesn't currently change the wallpaper back though.

---------

## 🕑 SandboxStartup.ps1 (Startup Script)

This script runs within the Sandbox at launch and does a bunch of random registry tweaks to set up the Sandbox based on my own preference

Specifically:

- Enables the old context menu
- Enables Explorer settings: Show hidden files, Show file extensions
- Adds "Open PowerShell Here" and "Open CMD Here" options to the context menu
- Adds `.txt` and `.ps1` files to the Context Menu > "New" list
- Sets PowerShell execution policy to allow running scripts, and also makes them runnable by double clicking
- If you put Notepad or Notepad++ into the shared host folder (described in next section), it will add context menu options to Edit files with them, and set them as default editor for `.txt` files
- Finally restarts Explorer to apply any changes and opens the shared host folder


### How to Use `SandboxStartup.ps1`

It is written assuming it will be run from _within_ the sandbox, so to automatically run it you'll need to put it into a mapped shared folder.

1. Create some new folder location (not in the sandbox) which you'll map into the sandbox. It doesn't matter what it's called or where it goes, but maybe something like `C:\Users\WhateverUsername\MySharedSandboxFolder`
2. In this repo I have the  [`MyDefaultSandbox.wsb`](Sandbox%20Configurations/MyDefaultSandbox.wsb) configuration file which is already set up to map the folder to the location the script expects. So in there you just need to update the `<HostFolder>` setting to use the path you selected in the previous step. 

    For Example:
    ```
    <HostFolder>C:\Users\WhateverUsername\MySharedSandboxFolder</HostFolder>
    ```
    
3. Update any other options to your liking in the `.wsb` file, such as amount of RAM.
4. Launch Sandbox using the configuration by double clicking `MyDefaultSandbox.wsb`. It will map the folder to the Desktop as a folder called `HostShared`, and run the script automatically. You can also add other scripts and things to the shared folder you may want to run manually.


