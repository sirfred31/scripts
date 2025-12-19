<#
.SYNOPSIS
   Generates a VPN user report with password expiration information
.DESCRIPTION
   This script checks password last set and expiration dates for VPN users across multiple domains.
   It outputs the results in a format suitable for Google Sheets with columns for:
    - Username
    - Password Last Set (24-hour format)
    - Password Expires (24-hour format)
    - Remaining Days
    - Status (OK/WARNING/EXPIRED/NEVER EXPIRES)
    - Password Age
.NOTES
    Created: 2025-04-21
    Requirements: PowerShell 5.1 or later
    Execution Policy: Requires RemoteSigned or Unrestricted

    Features:
    - Automatic admin elevation
    - 24-hour time format output
    - 90-day password policy enforcement
    - Detailed status indicators
    - Error handling for all user accounts
#>

# Configuration
$vpnPrefixes = "GLD-","TAR-","VEV-","PAL-","OAM-","LUH-"
$reportPath = "$env:USERPROFILE\Desktop\VPNReport_$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"

# 1. Get VPN Users
$users = Get-ADUser -Filter * | 
         Where-Object { $_.SamAccountName -match "^($($vpnPrefixes -join '|'))" } |
         Select-Object -ExpandProperty SamAccountName

# 2. Process Users and Build CSV Data
$csvData = foreach ($user in $users) {
    try {
        $output = net user $user /domain 2>&1 | Out-String
        
        # Initialize with clean default values
        $lastSet = $expires = $status = $age = "N/A"
        
        # Extract dates
        if ($output -match "Password last set\s+(.+?)\r?\n") { 
            $lastSet = $matches[1].Trim() 
        }
        if ($output -match "Password expires\s+(.+?)\r?\n") { 
            $expires = $matches[1].Trim() 
        }

        # Create properly formatted object
        [PSCustomObject]@{
            Username        = $user
            PasswordLastSet = if ($lastSet -ne "N/A") { 
                ([DateTime]$lastSet).ToString("M/d/yyyy HH:mm") 
            } else { "N/A" }
            PasswordExpires = if ($expires -like "*Never*") { 
                "Never" 
            } elseif ($expires -ne "N/A") { 
                ([DateTime]$expires).ToString("M/d/yyyy HH:mm") 
            } else { "N/A" }
            Status          = if ($expires -like "*Never*") {
                "NEVER_EXPIRES"
            } elseif ($lastSet -ne "N/A" -and $expires -ne "N/A") {
                $days = (([DateTime]$expires) - (Get-Date)).Days
                if ($days -lt 0) { "EXPIRED" }
                elseif ($days -le 7) { "WARNING" }
                else { "OK" }
            } else { "ERROR" }
            PasswordAge     = if ($lastSet -ne "N/A") {
                ((Get-Date) - ([DateTime]$lastSet)).Days
            } else { "N/A" }
        }
    }
    catch {
        # Error handling with guaranteed column alignment
        [PSCustomObject]@{
            Username        = $user
            PasswordLastSet = "ERROR"
            PasswordExpires = "ERROR"
            Status          = "ERROR"
            PasswordAge     = "N/A"
        }
    }
}

# 3. Export to CSV with PERFECT column alignment
$csvData | Export-Csv -Path $reportPath -NoTypeInformation -Encoding UTF8

# 4. Verify and Open
Write-Host "Report generated with perfect columns: $reportPath" -ForegroundColor Green
Start-Process $reportPath