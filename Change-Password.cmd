@ECHO OFF

if exist .\debug.txt (
    rem file exists
    PowerShell.exe -STA -NoProfile -ExecutionPolicy Bypass -Command "& '%~dpn0.ps1'"
    Pause
) else (
    rem file doesn't exist
    PowerShell.exe -STA -WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -Command "& '%~dpn0.ps1'"
    Exit
)
