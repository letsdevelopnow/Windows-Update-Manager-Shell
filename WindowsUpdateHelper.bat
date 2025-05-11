@echo off
:: Windows Update Manager
:: Optimized version of your clean implementation

SETLOCAL
set "VERSION=1.0"

:menu
cls
echo ================================
echo    WINDOWS UPDATE MANAGER v%VERSION%
echo ================================
echo.
echo  [1] Install Windows Updates
echo  [2] Clear Windows Update Cache
echo  [3] Exit
echo.
set /p choice="Select option (1-3): "

if "%choice%"=="1" (
    set "script=Install-Updates.ps1"
    set "action=Starting update installation"
    goto run
) else if "%choice%"=="2" (
    set "script=Clear-UpdateCache.ps1"
    set "action=Starting cache cleanup"
    goto run
) else if "%choice%"=="3" (
    goto exit
)
goto menu

:run
cls
echo %action%...
echo (Administrator privileges required)
PowerShell.exe -NoProfile -ExecutionPolicy Bypass -Command ^
    "Start-Process PowerShell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File \"%~dp0%script%\"' -Verb RunAs -Wait"

echo Operation completed successfully

if "%choice%"=="1" echo.& echo NOTE: Some updates may require a system restart
pause
goto menu

:exit
echo.
echo Thank you for using Windows Update Manager, please restart your PC to complete the process.
timeout /t 6 >nul
exit
