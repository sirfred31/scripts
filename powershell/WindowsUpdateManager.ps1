# ==================================
# Windows Update Manager (PowerShell)
# ==================================
# Safely disables/enables Windows Update settings
# Must be run as Administrator

function Check-Admin {
    if (-not ([Security.Principal.WindowsPrincipal] `
        [Security.Principal.WindowsIdentity]::GetCurrent()
    ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Host "`n[!] This script requires administrative privileges."
        Write-Host "    Right-click and choose 'Run as Administrator'."
        Pause
        exit
    }
}

function Show-Menu {
    Clear-Host
    Write-Host "==============================="
    Write-Host "   Windows Update Manager"
    Write-Host "==============================="
    Write-Host ""
    Write-Host "1. Disable Windows Update"
    Write-Host "2. Enable Windows Update"
    Write-Host ""
    $choice = Read-Host "Enter your choice (1 or 2)"
    switch ($choice) {
        "1" { Disable-WindowsUpdate }
        "2" { Enable-WindowsUpdate }
        default {
            Write-Host "`n[!] Invalid choice."
            Pause
            exit
        }
    }
}

function Disable-WindowsUpdate {
    Write-Host "`n=== Disabling Windows Update ==="

    # Stop and disable services
    Stop-Service -Name "wuauserv","UsoSvc","WaaSMedicSvc" -Force -ErrorAction SilentlyContinue
    Set-Service -Name "wuauserv" -StartupType Disabled
    Set-Service -Name "UsoSvc" -StartupType Disabled
    Set-Service -Name "WaaSMedicSvc" -StartupType Disabled

    # Registry tweaks
    New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Force | Out-Null
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" `
        -Name "NoAutoUpdate" -Value 1 -Type DWord
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" `
        -Name "AUOptions" -Value 1 -Type DWord

    # Hosts file block
    Write-Host "[SC] Blocking Microsoft Update Servers..."
    Start-Sleep -Seconds 2
    Write-Host "[SC] SUCCESS"

    $hostsPath = "$env:SystemRoot\System32\drivers\etc\hosts"
    $blockList = @(
        "127.0.0.1 windowsupdate.microsoft.com",
        "127.0.0.1 update.microsoft.com",
        "127.0.0.1 download.windowsupdate.com",
        "127.0.0.1 wustat.windows.com",
        "127.0.0.1 ntservicepack.microsoft.com"
    )
    Add-Content -Path $hostsPath -Value $blockList

    Write-Host "`n[SC] Windows Update has been DISABLED"
    Pause
}

function Enable-WindowsUpdate {
    Write-Host "`n=== Enabling Windows Update ==="

    # Enable and start services
    Set-Service -Name "wuauserv" -StartupType Manual
    Set-Service -Name "UsoSvc" -StartupType Manual
    Set-Service -Name "WaaSMedicSvc" -StartupType Manual -ErrorAction SilentlyContinue
    Start-Service -Name "wuauserv","UsoSvc","WaaSMedicSvc" -ErrorAction SilentlyContinue

    # Remove registry tweaks
    Remove-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" `
        -Name "NoAutoUpdate" -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" `
        -Name "AUOptions" -ErrorAction SilentlyContinue

    # Unblock hosts
    Write-Host "[SC] Allowing Microsoft Update Servers..."
    Start-Sleep -Seconds 2
    Write-Host "[SC] SUCCESS"

    $hostsPath = "$env:SystemRoot\System32\drivers\etc\hosts"
    $filter = "windowsupdate.microsoft.com","update.microsoft.com","download.windowsupdate.com","wustat.windows.com","ntservicepack.microsoft.com"
    $updatedHosts = Get-Content -Path $hostsPath | Where-Object {
        $include = $true
        foreach ($entry in $filter) {
            if ($_ -match [regex]::Escape($entry)) { $include = $false }
        }
        $include
    }
    Set-Content -Path $hostsPath -Value $updatedHosts

    Write-Host "`n[SC] Windows Update has been ENABLED."
    Pause
}

# Main
Check-Admin
Show-Menu
