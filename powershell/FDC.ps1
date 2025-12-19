<#
.SYNOPSIS
    Configures domain controller settings to force clients to use specific domain controllers
.DESCRIPTION
    This script modifies registry settings to enforce domain controller preferences for:
    - Site-specific domain controller usage (user selectable)
    - AvoidSelf setting (prevent using itself as DC)
    - TryNextClosestSite setting (disable failover to other sites)

    This version ONLY runs on Windows 11. Other OS versions will be blocked.
.NOTES
    Created: 2025-11-19
    Requirements: 
    - Windows 11 only
    - PowerShell 5.1 or later
    - Administrator privileges
    - Domain-joined Windows machine
#>

# ---------------------------
# ✔ Windows 11 Check
# ---------------------------
$osBuild = (Get-CimInstance Win32_OperatingSystem).BuildNumber
$osVersion = [System.Environment]::OSVersion.Version

# Windows 11 has build numbers: 22000+ (up to 26xxx)
if ($osBuild -lt 22000) {
    Write-Host "ERROR: This script is allowed only on Windows 11." -ForegroundColor Red
    Write-Host "Detected OS Build: $osBuild" -ForegroundColor Yellow
    exit 1
}

# ---------------------------
# Site Selection
# ---------------------------
$sites = @(
    "AWS-SG-SITE",
    "PH-GL2-SITE",
    "PH-JKA-SITE",
    "PH-MAT-SITE",
    "PH-RSC-SITE"
)

Write-Host "Select the SiteName to configure:`n" -ForegroundColor Cyan
for ($i = 0; $i -lt $sites.Count; $i++) {
    Write-Host "$($i+1). $($sites[$i])"
}
$selection = Read-Host "`nEnter the number of your choice"

if ($selection -match '^[1-5]$') {
    $selectedSite = $sites[$selection - 1]
    Write-Host "You selected: $selectedSite" -ForegroundColor Green
} else {
    Write-Host "Invalid selection. Exiting..." -ForegroundColor Red
    exit 1
}

# ---------------------------
# Execution Policy Check
# ---------------------------
$currentPolicy = Get-ExecutionPolicy -Scope CurrentUser
if ($currentPolicy -eq "Restricted") {
    Write-Host "Enabling script execution for current user..." -ForegroundColor Yellow
    Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
}

# ---------------------------
# Admin Check (Auto-Elevate)
# ---------------------------
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole] "Administrator")) {

    $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    Start-Process powershell.exe -ArgumentList $arguments -Verb RunAs -Wait
    exit
}

# ---------------------------
# Apply Registry Settings
# ---------------------------
try {
    Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon\Parameters" `
        -Name "SiteName" -Value $selectedSite -Type String -ErrorAction Stop

    Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon\Parameters" `
        -Name "AvoidSelf" -Value 1 -Type DWord -ErrorAction Stop

    Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon\Parameters" `
        -Name "TryNextClosestSite" -Value 0 -Type DWord -ErrorAction Stop  

    Write-Host "`nRegistry settings updated successfully:" -ForegroundColor Green
    Write-Host "Site: $selectedSite"
    Write-Host "AvoidSelf: Enabled"
    Write-Host "TryNextClosestSite: Disabled"
}
catch {
    Write-Host "Failed to configure registry settings: $_" -ForegroundColor Red
    exit 1
}

# ---------------------------
# GPUpdate
# ---------------------------
try {
    gpupdate /force | Out-Null
    Write-Host "`nGroup Policy updated successfully" -ForegroundColor Green
}
catch {
    Write-Host "Warning: Group Policy update failed" -ForegroundColor Yellow
}

# ---------------------------
# Restart Prompt
# ---------------------------
if ([Environment]::UserInteractive) {
    $restartChoice = Read-Host "`nDo you want to restart the computer now to apply changes? (Y/N)"
    if ($restartChoice -match '^[Yy]$') {
        Write-Host "Restarting computer..." -ForegroundColor Cyan
        Restart-Computer -Force
    }
    else {
        Write-Host "Changes will take effect after the next manual restart." -ForegroundColor Yellow
    }
}
