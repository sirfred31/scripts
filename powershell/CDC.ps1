<#
.SYNOPSIS
    Retrieves and displays domain controller and group policy information for a specific site
.DESCRIPTION
    This diagnostic script checks and reports:
    - Current domain controller information for the selected site using nltest
    - Group Policy application source using gpresult
    - Any Group Policy processing errors
    Results are displayed in a graphical window with categorized sections for easy analysis.

    This version works ONLY on Windows 11 and will block other OS versions.
.NOTES
    Created: 2025-11-19
    Requirements:
    - Windows 11 (Build 22000+)
    - PowerShell 5.1 or later
    - Administrator privileges
    - Active Directory environment
    - nltest.exe available in system path
#>

# ---------------------------
# ✔ Windows 11 Check
# ---------------------------
$osBuild = [int](Get-CimInstance Win32_OperatingSystem).BuildNumber

if ($osBuild -lt 22000) {
    Write-Host "❌ ERROR: This script is allowed ONLY on Windows 11." -ForegroundColor Red
    Write-Host "Detected OS Build: $osBuild" -ForegroundColor Yellow
    exit 1
}

# Ensure execution policy allows script
$currentPolicy = Get-ExecutionPolicy -Scope CurrentUser
if ($currentPolicy -eq "Restricted") {
    Write-Host "Enabling script execution for current user..." -ForegroundColor Yellow
    Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
}

# Relaunch as admin if not elevated
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole] "Administrator")) {
    $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    Start-Process powershell.exe -ArgumentList $arguments -Verb RunAs -Wait
    exit
}

# --- Site selection ---
$sites = @(
    "AWS-SG-SITE",
    "PH-GL2-SITE",
    "PH-JKA-SITE",
    "PH-MAT-SITE",
    "PH-RSC-SITE"
)

Write-Host "Select the SiteName to query:`n" -ForegroundColor Cyan
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

# --- Command execution helper ---
function Invoke-CommandCapture {
    param (
        [string]$Command
    )
    try {
        $process = Start-Process -FilePath "cmd.exe" `
            -ArgumentList "/c $Command" `
            -NoNewWindow -Wait -PassThru `
            -RedirectStandardOutput "$env:TEMP\cmdout.txt" `
            -RedirectStandardError "$env:TEMP\cmderr.txt"

        $output = Get-Content "$env:TEMP\cmdout.txt" -Raw
        $errorOutput = Get-Content "$env:TEMP\cmderr.txt" -Raw
        
        if ($process.ExitCode -ne 0) {
            return "Command failed with exit code $($process.ExitCode):`r`n$errorOutput"
        }
        return $output
    }
    catch {
        return "Error executing command: $_"
    }
    finally {
        if (Test-Path "$env:TEMP\cmdout.txt") { Remove-Item "$env:TEMP\cmdout.txt" }
        if (Test-Path "$env:TEMP\cmderr.txt") { Remove-Item "$env:TEMP\cmderr.txt" }
    }
}

# --- Run diagnostics ---
$dcInfo   = Invoke-CommandCapture "nltest /dsgetdc:openaccess.bpo /site:$selectedSite"
$gpInfo   = Invoke-CommandCapture "gpresult /r"
$gpSource = $gpInfo | Select-String "Group Policy was applied from" | Out-String
$gpErrors = $gpInfo | Select-String "error" | Out-String

$result = @"
=== DOMAIN CONTROLLER INFORMATION (Site: $selectedSite) ===
$dcInfo

=== GROUP POLICY SOURCE ===
$gpSource

=== GROUP POLICY ERRORS ===
$($gpErrors.Trim())
"@

# --- Show results in GUI ---
Add-Type -AssemblyName System.Windows.Forms
$form = New-Object System.Windows.Forms.Form
$form.Text = "Domain Controller and Policy Results - $selectedSite"
$form.Width = 800
$form.Height = 600
$form.StartPosition = "CenterScreen"

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Multiline = $true
$textBox.Width = 780
$textBox.Height = 550
$textBox.ScrollBars = "Both"
$textBox.Font = New-Object System.Drawing.Font("Consolas", 10)
$textBox.Text = $result
$textBox.Dock = "Fill"
$textBox.ReadOnly = $true

$form.Controls.Add($textBox)
[void]$form.ShowDialog()
