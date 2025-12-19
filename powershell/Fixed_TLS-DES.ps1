<#
.SYNOPSIS
    Configures Windows registry settings to disable older TLS protocols and enforce TLS 1.2+.

.DESCRIPTION
    This script modifies Windows registry settings to:
    - Disable TLS 1.0 and TLS 1.1
    - Enable and enforce TLS 1.2
    - Enable support for TLS 1.3 (HTTP/3 if supported)
    - Disable Triple DES (3DES) cipher
    - Disable legacy features like Remote Desktop, Samba, and SmartScreen

    It ensures the required registry keys exist before attempting to set their values.
    At the end of execution, it displays a confirmation message.

.NOTES
    Requirements: 
    - PowerShell 5.1 or later
    - Must be run with Administrator privileges
    - Applies to Windows 10 and Windows Server 2016+

    Registry Changes:
    - Disables: TLS 1.0, TLS 1.1, Triple DES, Remote Desktop, SmartScreen, Samba signing
    - Enables: TLS 1.2 (client & server), TLS 1.3 (HTTP/3)

    Post-Configuration:
    - Shows a message confirming success
#>
# Requires admin privileges
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Start-Process powershell -Verb runAs -ArgumentList $MyInvocation.MyCommand.Definition
    exit
}

function Ensure-RegistryKey {
    param (
        [string]$Path
    )
    if (-not (Test-Path -Path $Path)) {
        New-Item -Path $Path -Force | Out-Null
    }
}

function Set-RegistryValueSafe {
    param (
        [string]$Path,
        [string]$Name,
        [Object]$Value,
        [Microsoft.Win32.RegistryValueKind]$Type
    )
    Ensure-RegistryKey -Path $Path

    $currentValue = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue | Select-Object -ExpandProperty $Name -ErrorAction SilentlyContinue
    if ($null -ne $currentValue -and $currentValue -eq $Value) {
        Write-Host "[SKIP] $Path\$Name already set to $Value"
    } else {
        Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type $Type
        Write-Host "[APPLY] $Path\$Name set to $Value"
    }
}

# Base path
$baseTLS = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols"

# Disable TLS 1.0
Set-RegistryValueSafe "$baseTLS\TLS 1.0\Client" "DisabledByDefault" 1 DWord
Set-RegistryValueSafe "$baseTLS\TLS 1.0\Client" "Enabled" 0 DWord
Set-RegistryValueSafe "$baseTLS\TLS 1.0\Server" "DisabledByDefault" 1 DWord
Set-RegistryValueSafe "$baseTLS\TLS 1.0\Server" "Enabled" 0 DWord

# Disable TLS 1.1
Set-RegistryValueSafe "$baseTLS\TLS 1.1\Client" "DisabledByDefault" 1 DWord
Set-RegistryValueSafe "$baseTLS\TLS 1.1\Client" "Enabled" 0 DWord
Set-RegistryValueSafe "$baseTLS\TLS 1.1\Server" "DisabledByDefault" 1 DWord
Set-RegistryValueSafe "$baseTLS\TLS 1.1\Server" "Enabled" 0 DWord

# Enable TLS 1.2
Set-RegistryValueSafe "$baseTLS\TLS 1.2\Client" "DisabledByDefault" 0 DWord
Set-RegistryValueSafe "$baseTLS\TLS 1.2\Client" "Enabled" 1 DWord
Set-RegistryValueSafe "$baseTLS\TLS 1.2\Server" "DisabledByDefault" 0 DWord
Set-RegistryValueSafe "$baseTLS\TLS 1.2\Server" "Enabled" 1 DWord

# Enable TLS 1.3 (HTTP/3)
Set-RegistryValueSafe "HKLM:\SYSTEM\CurrentControlSet\Services\HTTP\Parameters" "EnabledHttp3" 1 DWord

# Disable 3DES (Triple DES)
Ensure-RegistryKey "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\Triple DES 168"
Set-RegistryValueSafe "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\Triple DES 168" "Enabled" 0 DWord

# Disable Samba signing requirement
Set-RegistryValueSafe "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters" "requiresecuritysignature" 0 DWord

# Disable Remote Desktop
Set-RegistryValueSafe "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server" "fDenyTSConnections" 1 DWord

# Disable SmartScreen
Set-RegistryValueSafe "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System" "EnableSmartScreen" 0 DWord
Set-RegistryValueSafe "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System" "ShellSmartScreenLevel" "Warn" String

Write-Host "`n✅ TLS settings have been successfully applied to the registry."
