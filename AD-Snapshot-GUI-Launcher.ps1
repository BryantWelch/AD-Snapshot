#requires -version 5.0

<#
.SYNOPSIS
    Active Directory Snapshot GUI Tool
.DESCRIPTION
    This script launches a Windows Forms GUI for the AD-Snapshot PowerShell script.
    It provides a user-friendly interface to configure and run AD snapshot reports.
.NOTES
    File Name      : AD-Snapshot-GUI-Launcher.ps1
    Author         : Bryant Welch
    Prerequisite   : PowerShell 5.0, Active Directory module (recommended)
    Version        : 1.0.0
#>

# Check if running as administrator
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Warning "This script should be run as Administrator to ensure full access to Active Directory. Some features may not work correctly."
    $continue = Read-Host "Do you want to continue anyway? (Y/N)"
    if ($continue -ne "Y") {
        exit
    }
}

# Check if Active Directory module is available
try {
    Import-Module ActiveDirectory -ErrorAction Stop
    Write-Host "Active Directory module loaded successfully." -ForegroundColor Green
}
catch {
    Write-Warning "The Active Directory module could not be loaded. Some features may not work correctly."
    Write-Host "You can install RSAT tools by running the following command in PowerShell as Administrator:" -ForegroundColor Yellow
    Write-Host "Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0" -ForegroundColor Cyan
    
    $continue = Read-Host "Do you want to continue without the Active Directory module? (Y/N)"
    if ($continue -ne "Y") {
        exit
    }
}

# Launch the main GUI script
& "$PSScriptRoot\AD-Snapshot-GUI-Main.ps1"
