$PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition

If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {    
    Write-Host "Not running with administrative rights. Attempting to elevate..."
    $command = "/c '.\test.cmd'"
    Start-Process cmd.exe -verb runas -argumentlist $command
    Exit
}


if(!(Test-Admin)) {
}

$Host.UI.RawUI.WindowTitle="Boxstarter Shell"
cd $env:SystemDrive\
Write-Output @"
Welcome to the Boxstarter shell!
The Boxstarter commands have been imported from $here and are available for you to run in this shell.
You may also import them into the shell of your choice. 
Here are some commands to get you started:
Install a Package:   Install-BoxstarterPackage
Create a Package:    New-BoxstarterPackage
Build a Package:     Invoke-BoxstarterBuild
Enable a VM:         Enable-BoxstarterVM
For Command help:    Get-Help <Command Name> -Full
For Boxstarter documentation, source code, to report bugs or participate in discussions, please visit http://boxstarter.org
"@
