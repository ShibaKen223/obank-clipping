@echo off
rem ---------------------------------------------------------------
rem  Daily Clipping - one-click installer (double-click this file)
rem
rem  This file is intentionally ASCII-only: cmd.exe reads .bat in the
rem  OEM codepage (cp950 on Traditional Chinese Windows), so Chinese
rem  text here would come out as garbage. All messages the user reads
rem  are printed by tools\install.ps1, which handles UTF-8 properly.
rem
rem  -ExecutionPolicy Bypass applies to this one run only. It does not
rem  change any system setting.
rem ---------------------------------------------------------------
setlocal

set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

if not exist "%~dp0tools\install.ps1" (
    echo.
    echo   Cannot find tools\install.ps1
    echo   The folder looks incomplete - please unzip it again,
    echo   and make sure you run this from the unzipped folder.
    echo.
    pause
    exit /b 1
)

"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0tools\install.ps1"

if errorlevel 1 (
    echo.
    echo   Install did not finish. See the messages above.
    pause
)

endlocal
