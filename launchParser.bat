:: Launches the log parsing powershell script and sends all the nonsene data that was
:: output to the console into $null. This greatly decreases parsing time.
:: Philip Otter 2024
@echo off
echo Loading...
powershell -command ".\problemDevices.ps1 > $null"
pause