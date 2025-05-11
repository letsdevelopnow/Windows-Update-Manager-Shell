# Windows-Update-Manager-Shell
A PowerShell script that automates Windows updates - checks, downloads, and installs missing updates with progress tracking. Runs via Microsoft.Update.Session COM object.


A PowerShell and Batch solution for managing Windows updates with administrative control.


## Important Deployment Note

For proper functionality:
1. Maintain all scripts in the same directory
2. Preserve the original filenames
3. Keep the file structure intact

This ensures:
- Correct inter-script dependencies
- Proper path resolution
- Seamless batch file operation

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
## Requirements
- Windows 7/8/10/11
- PowerShell 5.1+
- Administrator rights

## Technical Details
- Uses `Microsoft.Update.Session` COM object for update management
- Implements `takeown` + `icacls` for cache clearance
- Services are automatically stopped/restarted during cache cleaning
- Shows per-update download/install progress

## Notes
- System restart may be required after updates
- Cache clearing will force re-download of updates
- Always create a restore point before major updates

Alternatively run the PowerShell scripts directly as admin:
```powershell
# Install updates
.\Install-Updates.ps1

# Clear cache
.\Clear-UpdateCache.ps1



