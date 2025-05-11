# Windows-Update-Manager-Shell
A PowerShell script that automates Windows updates - checks, downloads, and installs missing updates with progress tracking. Runs via Microsoft.Update.Session COM object.


A PowerShell and Batch solution for managing Windows updates with administrative control.

## Features
- One-click menu interface (`WindowsUpdateHelper.bat`)
- Automated update installation with progress tracking
- Force-clear update cache with permission override
- Reboot requirement detection
- Admin privilege verification

## Files
- `WindowsUpdateHelper.bat` - Main menu interface
- `Install-Updates.ps1` - Installs available Windows updates
- `Clear-UpdateCache.ps1` - Clears Windows Update cache forcefully

## Usage
1. Right-click `WindowsUpdateHelper.bat` and select "Run as administrator"
2. Choose an option:
   - `1`: Install all available updates
   - `2`: Clear Windows Update cache
   - `3`: Exit

Alternatively run the PowerShell scripts directly as admin:
```powershell
# Install updates
.\Install-Updates.ps1

# Clear cache
.\Clear-UpdateCache.ps1
