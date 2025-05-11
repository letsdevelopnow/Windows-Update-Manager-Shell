<#
.SYNOPSIS
    Force-clears Windows Update cache with advanced permissions
.DESCRIPTION
    Uses Takeown + ICACLS to override file permissions before deletion
#>

# Run as Admin check
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Run as Administrator!" -ForegroundColor Red
    exit 1
}

$downloadPath = "$env:windir\SoftwareDistribution\Download"

try {
    # Stop dependent services
    Stop-Service -Name wuauserv -Force -ErrorAction Stop
    Stop-Service -Name TrustedInstaller -Force -ErrorAction Stop

    # Take ownership recursively
    Write-Host "Taking ownership of files..." -ForegroundColor Cyan
    takeown /F $downloadPath /A /R /D Y | Out-Null

    # Grant full permissions
    Write-Host "Resetting permissions..." -ForegroundColor Cyan
    icacls $downloadPath /reset /T /C /L /Q | Out-Null
    icacls $downloadPath /grant Administrators:F /T /C /L /Q | Out-Null

    # Delete contents
    Write-Host "Removing files..." -ForegroundColor Yellow
    Remove-Item -Path "$downloadPath\*" -Recurse -Force -ErrorAction Stop

    # Restart services
    Start-Service -Name TrustedInstaller -ErrorAction Stop
    Start-Service -Name wuauserv -ErrorAction Stop

    Write-Host "Update cache cleared successfully." -ForegroundColor Green
}
catch {
    Write-Host "ERROR: $_" -ForegroundColor Red
    # Attempt service restart if failed
    Start-Service -Name TrustedInstaller -ErrorAction SilentlyContinue
    Start-Service -Name wuauserv -ErrorAction SilentlyContinue
    exit 1
}