<#
.SYNOPSIS
    Automatically checks, downloads, and installs missing Windows updates.
.DESCRIPTION
    Uses Microsoft.Update.Session to find, download, and install missing updates. Must be run as Administrator.
.NOTES
    Automatically installs updates if download succeeds. Shows per-update title and percent progress.
#>

# Check for admin rights
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "This script must be run as Administrator!" -ForegroundColor Red
    exit
}

# Initialize update session
try {
    $Session = New-Object -ComObject Microsoft.Update.Session
    $Searcher = $Session.CreateUpdateSearcher()
} catch {
    Write-Host "Failed to initialize update session." -ForegroundColor Red
    exit
}

Write-Host "Searching for missing updates..." -ForegroundColor Cyan
try {
    $SearchResult = $Searcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")
} catch {
    Write-Host "Failed to search for updates. Please check your network connection." -ForegroundColor Red
    exit
}

# If no updates found
if ($SearchResult.Updates.Count -eq 0) {
    Write-Host "`nNo update found to be installed." -ForegroundColor Yellow
    Write-Host "`nPress 0 to exit..." -ForegroundColor Cyan
    do {
        $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    } while ($key.Character -ne '0')
    exit
}

# Display available updates with sizes
Write-Host "`nFound $($SearchResult.Updates.Count) updates available:" -ForegroundColor Yellow
$i = 1
$totalSize = 0
$SearchResult.Updates | ForEach-Object {
    $sizeMB = [math]::Round($_.MaxDownloadSize / 1MB, 2)
    $totalSize += $sizeMB
    Write-Host "$i. $($_.Title) (KB$($_.KBArticleIDs)) - Size: $sizeMB MB" -ForegroundColor White
    $i++
}
Write-Host "`nTotal download size: $totalSize MB" -ForegroundColor Yellow

# Ask for download confirmation
$choice = Read-Host "`nDo you want to download and install ALL these updates? (Y/N)"
if ($choice -ne "Y" -and $choice -ne "y") {
    Write-Host "Update process cancelled by user." -ForegroundColor Yellow
    exit
}

# Create downloader
$Downloader = $Session.CreateUpdateDownloader()
$Downloader.Updates = $SearchResult.Updates

Write-Host "`nStarting download process..." -ForegroundColor Cyan

# Per-update download simulation
$total = $SearchResult.Updates.Count
for ($i = 0; $i -lt $total; $i++) {
    $update = $SearchResult.Updates.Item($i)
    Write-Host "`nDownloading: $($update.Title)" -ForegroundColor Cyan

    $individualUpdateList = New-Object -ComObject Microsoft.Update.UpdateColl
    $null = $individualUpdateList.Add($update)
    $Downloader.Updates = $individualUpdateList

    try {
        $DownloadResult = $Downloader.Download()
        if ($DownloadResult.ResultCode -eq 2) {
            Write-Host "Download progress: 100% completed." -ForegroundColor Green
        } else {
            Write-Host "Failed to download update: $($update.Title)" -ForegroundColor Red
        }
    } catch {
        Write-Host "Error downloading: $($update.Title)" -ForegroundColor Red
    }
}

# Verify all updates were downloaded
if ($SearchResult.Updates | Where-Object { -not $_.IsDownloaded }) {
    Write-Host "`nSome updates failed to download. Installation aborted." -ForegroundColor Red
    exit
}

Write-Host "`nAll updates downloaded successfully. Proceeding with installation..." -ForegroundColor Green

# Create installer
$Installer = $Session.CreateUpdateInstaller()
$Installer.Updates = $SearchResult.Updates

Write-Host "`nInstalling updates... (This may take a while, system may restart)" -ForegroundColor Cyan

# Per-update installation
for ($i = 0; $i -lt $SearchResult.Updates.Count; $i++) {
    $update = $SearchResult.Updates.Item($i)
    Write-Host "`nInstalling: $($update.Title)" -ForegroundColor Cyan

    $individualUpdateList = New-Object -ComObject Microsoft.Update.UpdateColl
    $null = $individualUpdateList.Add($update)

    $Installer.Updates = $individualUpdateList

    try {
        $InstallResult = $Installer.Install()
        if ($InstallResult.ResultCode -eq 2) {
            Write-Host "Install progress: 100% completed." -ForegroundColor Green
        } else {
            Write-Host "Failed to install update: $($update.Title)" -ForegroundColor Red
        }
    } catch {
        Write-Host "Error installing: $($update.Title)" -ForegroundColor Red
    }
}

# Check if reboot is required
$rebootRequired = $false
$SearchResult.Updates | ForEach-Object {
    if ($_.InstallationBehavior.RebootBehavior -gt 0) {
        $rebootRequired = $true
    }
}

if ($rebootRequired) {
    Write-Host "`nSome updates require a system restart to complete installation." -ForegroundColor Yellow
    Write-Host "Please restart your computer as soon as possible." -ForegroundColor Yellow
}

# Exit prompt
Write-Host "`nPress 0 to exit..." -ForegroundColor Cyan
do {
    $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
} while ($key.Character -ne '0')

exit
