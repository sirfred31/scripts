@echo off
setlocal EnableDelayedExpansion
:: ==========================================================
:: Advanced Windows Network Toolkit
:: ==========================================================
title Advanced Windows Network Toolkit v8.1
if /i not "%~1"=="RUN" ( cmd /k ""%~f0" RUN" & exit /b )
mode con: cols=138 lines=70 >nul 2>&1
color 0A
net session >nul 2>&1
if errorlevel 1 (
    color 0C
    echo.
    echo  ERROR: Please Run as Administrator
    pause >nul & exit /b
)
set "BLINK_ON=0"

:: ----------------------------------------------------------
:: ENSURE POWERSHELL NUGET & PYTHON PIP ARE READY
:: ----------------------------------------------------------
cls
color 0E
echo  ==================================================
echo    CHECKING SYSTEM DEPENDENCIES...
echo  ==================================================
color 0A

:: Check NuGet
powershell -NoProfile -ExecutionPolicy Bypass -Command "if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) { exit 1 } else { exit 0 }"
if %errorlevel% neq 0 (
    color 0E
    echo  [!] NuGet package is not installed... Attempting to install...
    color 0B
    powershell -NoProfile -ExecutionPolicy Bypass -Command "Install-PackageProvider -Name NuGet -Force -Scope CurrentUser | Out-Null"
    timeout /t 2 >nul
)

:: Check Python
where python >nul 2>&1
if errorlevel 1 (
    color 0E
    echo  [!] Python not found. Some advanced tools will require manual Python installation.
    timeout /t 3 >nul
) else (
    color 0A
    echo  [OK] NuGet and Python dependencies are ready.
    timeout /t 1 >nul
)

:MENU
cls
call :HEADER
echo.
color 0E
echo   +----------------------------------------+  +----------------------------------------+  +----------------------------------------+
echo   ^|            DIAGNOSTICS                 ^|  ^|             WIFI TOOLS                 ^|  ^|           REPAIR / RESET               ^|
echo   +----------------------------------------+  +----------------------------------------+  +----------------------------------------+
color 0B
echo   ^|  [1]  Full IP Configuration            ^|  ^|  [25] Open WiFi Settings               ^|  ^|  [45] Flush DNS Cache                  ^|
echo   ^|  [2]  Ping Google                      ^|  ^|  [26] WiFi Profile / Password          ^|  ^|  [46] Release IP Address               ^| 
echo   ^|  [3]  Continuous Ping (8.8.8.8)        ^|  ^|  [27] Show All WiFi Profiles           ^|  ^|  [47] Renew IP Address                 ^|
echo   ^|  [4]  Routing Table                    ^|  ^|  [28] Show WiFi Interfaces             ^|  ^|  [48] Reset Winsock                    ^|
echo   ^|  [5]  ARP Cache                        ^|  ^|  [29] Scan WiFi Networks               ^|  ^|  [49] Reset TCP/IP Stack               ^|
echo   ^|  [6]  Active Sessions                  ^|  ^|  [30] Export Profiles to Desktop       ^|  ^|  [50] Reset DNS Settings               ^|
echo   ^|  [7]  DNS Cache                        ^|  ^|  [31] WiFi Drivers Info                ^|  ^|  [51] Reset DHCP Configuration         ^|
echo   ^|  [8]  System Network Info              ^|  ^|  [32] WiFi Radio / Signal Status       ^|  ^|  [52] Reset IPv6 Stack                 ^|
echo   ^|  [9]  Test-Connection (PowerShell)     ^|  ^|  [33] WiFi Hosted Network Status       ^|  ^|  [53] Reset Firewall to Default        ^|
echo   ^|  [10] Test-NetConnection (Port Check)  ^|  ^|  [34] WiFi Blocked Networks            ^|  ^|  [54] Reset Proxy Settings             ^|
echo   ^|  [11] Trace Route (tracert)            ^|  ^|  [35] WiFi WLAN Report                 ^|  ^|  [55] Reset NetBIOS over TCP/IP        ^|
echo   ^|  [12] PathPing (Latency per Hop)       ^|  ^|  [36] Connect to WiFi Profile          ^|  ^|  [56] Restart WLAN AutoConfig          ^|
echo   ^|  [13] DNS Lookup (nslookup)            ^|  ^|  [37] Disconnect WiFi                  ^|  ^|  [57] Restart Network Adapters         ^|
echo   ^|  [14] NetBIOS Cache (nbtstat)          ^|  ^|  [38] WiFi AutoConfig On/Off           ^|  ^|  [58] Restart Windows Explorer         ^|
echo   ^|  [15] Check Listening Ports            ^|  ^|  [39] WiFi Profile Priority            ^|  ^|  [59] FULL NETWORK REPAIR  ^<^< REC      ^|
echo   ^|  [16] DNS Server Check (PowerShell)    ^|  ^|  [40] WiFi Random MAC Status           ^|  ^|                                        ^|
echo   ^|  [17] Get-NetIPConfiguration           ^|  ^|  [41] WiFi Wireless Capabilities       ^|  ^|  [99] EXIT                             ^|
echo   ^|  [18] Get-NetAdapter Status            ^|  ^|  [42] WiFi Settings Overview           ^|  ^|                                        ^|
echo   ^|  [19] Get-NetTCPConnection             ^|  ^|  [43] WiFi Tracing On/Off              ^|  ^|                                        ^|
echo   ^|  [20] Get-DnsClientServerAddress       ^|  ^|  [44] Export WiFi Profiles to Desktop  ^|  ^|                                        ^|
echo   ^|  [21] Get-NetAdapterBinding            ^|  ^|                                        ^|  ^|                                        ^|
echo   ^|  [22] Test-DnsServer                   ^|  ^|                                        ^|  ^|                                        ^|
echo   ^|  [23] Get-NetRoute                     ^|  ^|                                        ^|  ^|                                        ^|
echo   ^|  [24] Get-NetIPAddress                 ^|  ^|                                        ^|  ^|                                        ^|
color 0E
echo   +----------------------------------------+  +----------------------------------------+  +----------------------------------------+
echo   ^|          ADVANCED NETWORKING           ^|  ^|               MISC / INFO              ^|  ^|             QUICK DIAGNOSTIC           ^|
echo   +----------------------------------------+  +----------------------------------------+  +----------------------------------------+
color 0B
echo   ^|  [60] Packet Capture (pktmon)          ^|  ^|  [68] Network Connections              ^|  ^|  [85] Quick Connectivity Test          ^|
echo   ^|  [61] Speed Test (Ookla)               ^|  ^|  [69] Device Manager                   ^|  ^|  [86] Quick DNS Test                   ^|
echo   ^|  [62] Ping Sweep (Host Discovery)      ^|  ^|  [70] Network Troubleshooter           ^|  ^|  [87] Quick WiFi Signal Test           ^|
echo   ^|  [63] Find Public IP ^& GeoIP           ^|  ^|  [71] Windows Firewall                 ^|  ^|  [88] Quick Port Scan (1-1024)         ^|
echo   ^|  [64] Show Network Errors (netstat)    ^|  ^|  [72] Proxy Settings                   ^|  ^|  [89] Wake-on-LAN (WOL)                ^|
echo   ^|  [65] Test MTU Size (Ping DF)          ^|  ^|  [73] Network Reset (Windows)          ^|  ^|  [90] Test ALL Google DNS              ^|
echo   ^|  [66] Check DoH (DNS over HTTPS)       ^|  ^|  [74] Windows Update (Network)         ^|  ^|  [91] Show DHCP Lease Info             ^|
echo   ^|  [67] Open Advanced Firewall           ^|  ^|                                        ^|  ^|  [92] Show WiFi QR Code                ^|
color 0E
echo   +----------------------------------------+  +----------------------------------------+  +----------------------------------------+
echo   ^|            SECURITY CHECKS             ^|  ^|          NETWORK MONITORING            ^|  ^|             LOCAL TOOLS                ^|
echo   +----------------------------------------+  +----------------------------------------+  +----------------------------------------+
color 0B
echo   ^|  [75] Show SMB Shares ^& Sessions       ^|  ^|  [93] Bandwidth Monitor (Live)         ^|  ^|  [96] View DNS Cache                   ^|
echo   ^|  [76] Windows Firewall Active Rules    ^|  ^|  [94] DNS Benchmark (Fastest)          ^|  ^|  [97] View ARP Cache                   ^|
echo   ^|  [77] SMBv1 Protocol Check             ^|  ^|  [95] Traceroute via TCP (tcping)      ^|  ^|  [98] View Routing Table               ^|
echo   ^|  [78] NTLM Auditing Check              ^|  ^|                                        ^|  ^|  [95] View Active Sessions             ^|
echo   ^|  [79] Show TCP Parameters (Registry)   ^|  ^|                                        ^|  ^|                                        ^|
echo   ^|  [80] Show Network Bindings            ^|  ^|                                        ^|  ^|                                        ^|
color 0E
echo   +----------------------------------------+  +----------------------------------------+  +----------------------------------------+
echo   ^|            NEW TOOLS ADDED             ^|  ^|               NEW TOOLS                ^|  ^|               NEW TOOLS                ^|
echo   +----------------------------------------+  +----------------------------------------+  +----------------------------------------+
color 0B
echo   ^|  [100] MAC Address Lookup (Vendor)     ^|  ^|  [104] View Open UDP Ports             ^|  ^|  [108] Get Active HTTP Connections     ^|
echo   ^|  [101] Check Public DNS (Google/CF)    ^|  ^|  [105] View Open TCP Ports             ^|  ^|  [109] Check HTTP Proxy Auto-Detect    ^|
echo   ^|  [102] Find IP of a Domain             ^|  ^|  [106] Show Active Multicast Groups    ^|  ^|  [110] Test Network Port Speed (iPerf) ^|
echo   ^|  [103] Resolve Reverse DNS (PTR)       ^|  ^|  [107] Show IPv6 Global Addresses      ^|  ^|  [111] Reset WINS Cache                ^|
color 0E
echo   +----------------------------------------+  +----------------------------------------+  +----------------------------------------+
color 0A
echo.
set /p choice=  Select Option: 

:: ==========================================================
:: ROUTING LOGIC
:: ==========================================================
if "%choice%"=="1"  goto IPCONFIG
if "%choice%"=="2"  goto PINGGOOGLE
if "%choice%"=="3"  goto CONTPING
if "%choice%"=="4"  goto ROUTE
if "%choice%"=="5"  goto ARP
if "%choice%"=="6"  goto NETSTAT
if "%choice%"=="7"  goto DNSCACHE
if "%choice%"=="8"  goto SYSINFO
if "%choice%"=="9"  goto PS_TESTCONN
if "%choice%"=="10" goto PS_TESTNETCONN
if "%choice%"=="11" goto TRACERT
if "%choice%"=="12" goto PATHPING
if "%choice%"=="13" goto NSLOOKUP
if "%choice%"=="14" goto NBTSTAT
if "%choice%"=="15" goto LISTENPORTS
if "%choice%"=="16" goto PS_DNSCHECK
if "%choice%"=="17" goto PS_GETNETIPCONFIG
if "%choice%"=="18" goto PS_GETNETADAPTER
if "%choice%"=="19" goto PS_GETTCPCONN
if "%choice%"=="20" goto PS_GETDNSSERVER
if "%choice%"=="21" goto PS_BINDINGS
if "%choice%"=="22" goto PS_TESTDNSSERVER
if "%choice%"=="23" goto PS_GETROUTE
if "%choice%"=="24" goto PS_GETIPADDRESS

if "%choice%"=="25" goto WIFI
if "%choice%"=="26" goto WIFIPROFILE
if "%choice%"=="27" goto WIFIALLLIST
if "%choice%"=="28" goto WIFIINT
if "%choice%"=="29" goto WIFISCAN
if "%choice%"=="30" goto WIFIEXPORT
if "%choice%"=="31" goto WIFIDRIVERS
if "%choice%"=="32" goto WIFIRADIO
if "%choice%"=="33" goto WIFIHOSTED
if "%choice%"=="34" goto WIFIBLOCKED
if "%choice%"=="35" goto WIFIREPORT
if "%choice%"=="36" goto WIFICONNECT
if "%choice%"=="37" goto WIFIDISCONNECT
if "%choice%"=="38" goto WIFIAUTOCONFIG
if "%choice%"=="39" goto WIFIPRIORITY
if "%choice%"=="40" goto WIFIRANDOM
if "%choice%"=="41" goto WIFICAPS
if "%choice%"=="42" goto WIFISETTINGS
if "%choice%"=="43" goto WIFITRACE
if "%choice%"=="44" goto WIFIEXPORT

if "%choice%"=="45" goto FLUSHDNS
if "%choice%"=="46" goto RELEASE
if "%choice%"=="47" goto RENEW
if "%choice%"=="48" goto WINSOCK
if "%choice%"=="49" goto TCPRESET
if "%choice%"=="50" goto DNSRESET
if "%choice%"=="51" goto DHCPRESET
if "%choice%"=="52" goto IPV6RESET
if "%choice%"=="53" goto FWRESET
if "%choice%"=="54" goto PROXYRESET
if "%choice%"=="55" goto NETBIOSRESET
if "%choice%"=="56" goto WLANRESTART
if "%choice%"=="57" goto ADAPTERRESTART
if "%choice%"=="58" goto EXPLORER
if "%choice%"=="59" goto FULLREPAIR

if "%choice%"=="60" goto PKTCAP
if "%choice%"=="61" goto SPEEDTEST
if "%choice%"=="62" goto PINGSWEEP
if "%choice%"=="63" goto PUBLICIP
if "%choice%"=="64" goto NETERRORS
if "%choice%"=="65" goto MTUTEST
if "%choice%"=="66" goto DOHCHECK
if "%choice%"=="67" goto FIREWALLADV

if "%choice%"=="68" goto NCPA
if "%choice%"=="69" goto DEVICE
if "%choice%"=="70" goto TROUBLE
if "%choice%"=="71" goto FIREWALLPANEL
if "%choice%"=="72" goto PROXYPANEL
if "%choice%"=="73" goto NETRESET
if "%choice%"=="74" goto WU_NETWORK

if "%choice%"=="75" goto SMBCHECK
if "%choice%"=="76" goto FWRULES
if "%choice%"=="77" goto SMBV1
if "%choice%"=="78" goto NTLM
if "%choice%"=="79" goto TCPPARAMS
if "%choice%"=="80" goto BINDINGS

if "%choice%"=="85" goto QUICKCONN
if "%choice%"=="86" goto QUICKDNS
if "%choice%"=="87" goto QUICKWIFI
if "%choice%"=="88" goto QUICKPORTSCAN
if "%choice%"=="89" goto WOL
if "%choice%"=="90" goto QUICKALLDNS
if "%choice%"=="91" goto DHCPLEASE
if "%choice%"=="92" goto WIFIQR

if "%choice%"=="93" goto BANDWIDTH
if "%choice%"=="94" goto DNSBENCH
if "%choice%"=="95" goto TCPING

if "%choice%"=="96" goto DNSCACHE
if "%choice%"=="97" goto ARP
if "%choice%"=="98" goto ROUTE

if "%choice%"=="100" goto MACLOOKUP
if "%choice%"=="101" goto PUBDNS
if "%choice%"=="102" goto DOMAINIP
if "%choice%"=="103" goto REVDNS
if "%choice%"=="104" goto UDPPORTS
if "%choice%"=="105" goto TCPPORTS
if "%choice%"=="106" goto MULTICAST
if "%choice%"=="107" goto IPV6ADDR
if "%choice%"=="108" goto HTTPCONN
if "%choice%"=="109" goto PROXYWPAD
if "%choice%"=="110" goto IPERF
if "%choice%"=="111" goto WINSRESET

if "%choice%"=="99" goto QUIT

color 0C
echo. & echo  Invalid selection. Returning... & timeout /t 2 >nul & goto MENU

:: ==========================================================
:: HELPERS
:: ==========================================================
:HEADER
set /a BLINK_ON=(!BLINK_ON!+1) %% 2
if !BLINK_ON!==1 (color 0E) else (color 0B)
echo  ========================================================================================================================================
color 0A
echo                         A D V A N C E D   N E T W O R K   T O O L K I T
color 0C
echo                                        ~ Diagnose. Repair. Dominate. ~
color 0A
echo  ========================================================================================================================================
exit /b

:SECTION
color 0E & echo. & echo  ================================================== & echo    %~1 & echo  ================================================== & color 0A & echo. & exit /b
:WARN
color 0C & echo. & echo   ^!^! %~1 ^!^! & color 0A & exit /b
:OK
color 0A & echo. & echo   [OK] %~1 & color 0A & exit /b
:PAUSE_RETURN
echo. & color 0D & echo  ------------------------------------------------ & echo    Press any key to return... & color 0A & pause >nul & goto MENU

:: ==========================================================
:: DEPENDENCY INSTALLERS
:: ==========================================================
:INSTALL_SPEEDTEST
color 0E & echo  [!] speedtest-cli not found. Installing via pip...
color 0B & python -m pip install speedtest-cli
color 0A & echo  [OK] Installation complete. & timeout /t 2 >nul & exit /b

:INSTALL_MACLOOKUP
color 0E & echo  [!] mac-vendor-lookup not found. Installing via pip...
color 0B & python -m pip install mac-vendor-lookup
color 0A & echo  [OK] Installation complete. & timeout /t 2 >nul & exit /b

:: ==========================================================
:: DIAGNOSTICS
:: ==========================================================
:IPCONFIG
cls & call :SECTION "FULL IP CONFIGURATION" & color 0B & ipconfig /all & goto PAUSE_RETURN

:PINGGOOGLE
cls & call :SECTION "PING GOOGLE (10)" & color 0B & ping google.com -n 10 & goto PAUSE_RETURN

:CONTPING
cls & call :SECTION "CONTINUOUS PING 8.8.8.8" & call :WARN "CTRL+C to stop" & color 0B & ping 8.8.8.8 -t & goto PAUSE_RETURN

:ROUTE
cls & call :SECTION "ROUTING TABLE" & color 0B & route print & goto PAUSE_RETURN

:ARP
cls & call :SECTION "ARP CACHE" & color 0B & arp -a & goto PAUSE_RETURN

:NETSTAT
cls & call :SECTION "ACTIVE SESSIONS" & color 0B & netstat -ano & goto PAUSE_RETURN

:DNSCACHE
cls & call :SECTION "DNS CACHE" & color 0B & ipconfig /displaydns & goto PAUSE_RETURN

:SYSINFO
cls & call :SECTION "SYSTEM NETWORK INFO" & color 0E & echo  Host: %COMPUTERNAME%  User: %USERNAME%  %DATE% %TIME% & color 0A & echo. & color 0B
for /f "tokens=2 delims=:" %%A in ('ipconfig ^| findstr /C:"IPv4"') do ( for /f "tokens=* delims= " %%B in ("%%A") do echo  IPv4 : %%B )
for /f "tokens=2 delims=:" %%A in ('ipconfig ^| findstr /C:"Default Gateway"') do ( for /f "tokens=* delims= " %%B in ("%%A") do echo  GW   : %%B )
echo. & color 0E & echo  === DNS === & color 0B & ipconfig /all | findstr /C:"DNS Servers" & goto PAUSE_RETURN

:PS_TESTCONN
cls & call :SECTION "Test-Connection (Gateway + Internet)" & color 0B
set /p "tg= Enter target IP/host (default 8.8.8.8): " & if "!tg!"=="" set "tg=8.8.8.8"
powershell -NoProfile -Command "Test-Connection -TargetName '%tg%' -Count 4 -ErrorAction SilentlyContinue" & goto PAUSE_RETURN

:PS_TESTNETCONN
cls & call :SECTION "Test-NetConnection (TCP Port Check)" & color 0B
set /p "th= Host (default google.com): " & if "!th!"=="" set "th=google.com"
set /p "tp= Port (default 443): " & if "!tp!"=="" set "tp=443"
echo. & color 0E & echo  Testing %th%:%tp% ... & color 0B
powershell -NoProfile -Command "Test-NetConnection -ComputerName '%th%' -Port %tp% -InformationLevel Detailed" & goto PAUSE_RETURN

:TRACERT
cls & call :SECTION "TRACE ROUTE" & color 0B
set /p "tt= Target (default 8.8.8.8): " & if "!tt!"=="" set "tt=8.8.8.8"
tracert -d "!tt!" & goto PAUSE_RETURN

:PATHPING
cls & call :SECTION "PATHPING (Latency + Packet Loss per Hop)" & color 0B
set /p "tp2= Target (default 8.8.8.8): " & if "!tp2!"=="" set "tp2=8.8.8.8"
pathping -n "!tp2!" & goto PAUSE_RETURN

:NSLOOKUP
cls & call :SECTION "DNS LOOKUP (nslookup)" & color 0B
set /p "nd= Domain (default google.com): " & if "!nd!"=="" set "nd=google.com"
nslookup "!nd!" & goto PAUSE_RETURN

:NBTSTAT
cls & call :SECTION "NetBIOS CACHE" & color 0B & nbtstat -c & nbtstat -n & goto PAUSE_RETURN

:LISTENPORTS
cls & call :SECTION "LISTENING PORTS" & color 0B & netstat -an | findstr "LISTENING" & goto PAUSE_RETURN

:PS_DNSCHECK
cls & call :SECTION "DNS Server Check (Get-DnsClientServerAddress)" & color 0B
powershell -NoProfile -Command "if (-not (Get-Command Get-DnsClientServerAddress -ErrorAction SilentlyContinue)) { Enable-WindowsOptionalFeature -Online -FeatureName Microsoft.Windows.PowerShell.NetTCPIP -All -NoRestart | Out-Null } Get-DnsClientServerAddress | Format-Table -AutoSize" & goto PAUSE_RETURN

:PS_GETNETIPCONFIG
cls & call :SECTION "Get-NetIPConfiguration" & color 0B
powershell -NoProfile -Command "if (-not (Get-Command Get-NetIPConfiguration -ErrorAction SilentlyContinue)) { Enable-WindowsOptionalFeature -Online -FeatureName Microsoft.Windows.PowerShell.NetTCPIP -All -NoRestart | Out-Null } Get-NetIPConfiguration | Format-List" & goto PAUSE_RETURN

:PS_GETNETADAPTER
cls & call :SECTION "Get-NetAdapter Status" & color 0B
powershell -NoProfile -Command "if (-not (Get-Command Get-NetAdapter -ErrorAction SilentlyContinue)) { Enable-WindowsOptionalFeature -Online -FeatureName Microsoft.Windows.PowerShell.NetTCPIP -All -NoRestart | Out-Null } Get-NetAdapter | Format-Table Name, Status, LinkSpeed, MacAddress -AutoSize" & goto PAUSE_RETURN

:PS_GETTCPCONN
cls & call :SECTION "Get-NetTCPConnection" & color 0B
powershell -NoProfile -Command "if (-not (Get-Command Get-NetTCPConnection -ErrorAction SilentlyContinue)) { Enable-WindowsOptionalFeature -Online -FeatureName Microsoft.Windows.PowerShell.NetTCPIP -All -NoRestart | Out-Null } Get-NetTCPConnection | Sort-Object State | Format-Table LocalAddress,LocalPort,RemoteAddress,RemotePort,State -AutoSize" & goto PAUSE_RETURN

:PS_GETDNSSERVER
cls & call :SECTION "Get-DnsClientServerAddress" & color 0B
powershell -NoProfile -Command "Get-DnsClientServerAddress | Format-Table -AutoSize" & goto PAUSE_RETURN

:PS_BINDINGS
cls & call :SECTION "Get-NetAdapterBinding" & color 0B
powershell -NoProfile -Command "Get-NetAdapterBinding | Format-Table Name, DisplayName, Enabled -AutoSize" & goto PAUSE_RETURN

:PS_TESTDNSSERVER
cls & call :SECTION "Test-DnsServer (Resolve)" & color 0B
set /p "td= Domain to test (default google.com): " & if "!td!"=="" set "td=google.com"
powershell -NoProfile -Command "Resolve-DnsName -Name '!td!' | Format-Table -AutoSize" & goto PAUSE_RETURN

:PS_GETROUTE
cls & call :SECTION "Get-NetRoute" & color 0B
powershell -NoProfile -Command "Get-NetRoute | Format-Table -AutoSize" & goto PAUSE_RETURN

:PS_GETIPADDRESS
cls & call :SECTION "Get-NetIPAddress" & color 0B
powershell -NoProfile -Command "Get-NetIPAddress | Format-Table -AutoSize" & goto PAUSE_RETURN

:: ==========================================================
:: WIFI TOOLS
:: ==========================================================
:WIFI
cls & call :SECTION "OPENING WIFI SETTINGS" & start ms-settings:network-wifi & goto PAUSE_RETURN

:WIFIPROFILE
cls & call :SECTION "SELECT WIFI PROFILE (Show Password)" & set "i=0"
for /f "tokens=2 delims=:" %%A in ('netsh wlan show profiles ^| findstr /C:"All User Profile"') do (
    set /a i+=1 & set "profile[!i!]=%%A"
    for /f "tokens=* delims= " %%B in ("!profile[!i!]!") do set "profile[!i!]=%%B"
    color 0B & echo    [!i!] !profile[!i!]!
)
color 0A & if !i!==0 ( call :WARN "No profiles" & goto PAUSE_RETURN )
echo. & set /p "sel=  Profile # (0 cancel): " & if "%sel%"=="0" goto MENU
if not defined profile[%sel%] ( call :WARN "Invalid" & goto PAUSE_RETURN )
set "profile=!profile[%sel%]!" & call :SECTION "PASSWORD: !profile!" & color 0E
netsh wlan show profiles name="!profile!" key=clear & goto PAUSE_RETURN

:WIFIALLLIST
cls & call :SECTION "ALL WIFI PROFILES" & color 0B & netsh wlan show profiles & goto PAUSE_RETURN

:WIFIINT
cls & call :SECTION "WIFI INTERFACES" & color 0B & netsh wlan show interfaces & goto PAUSE_RETURN

:WIFISCAN
cls & call :SECTION "SCANNING WIFI NETWORKS" & color 0E & echo  Scanning... & color 0B & netsh wlan show networks mode=bssid & goto PAUSE_RETURN

:WIFIEXPORT
cls & call :SECTION "EXPORT WIFI PROFILES" & set "OUTDIR=%USERPROFILE%\Desktop\WiFiProfiles_Export"
if not exist "%OUTDIR%" mkdir "%OUTDIR%" >nul 2>&1 & color 0E & echo  Export to: %OUTDIR% & color 0B
netsh wlan export profile folder="%OUTDIR%" key=clear >nul 2>&1
if errorlevel 1 ( call :WARN "Export failed" ) else ( call :OK "Saved to Desktop" & start "" "%OUTDIR%" )
goto PAUSE_RETURN

:WIFIDRIVERS
cls & call :SECTION "WIFI DRIVERS INFO" & color 0B & netsh wlan show drivers & goto PAUSE_RETURN

:WIFIRADIO
cls & call :SECTION "WIFI RADIO / SIGNAL" & color 0B & netsh wlan show interfaces & goto PAUSE_RETURN

:WIFIHOSTED
cls & call :SECTION "HOSTED NETWORK STATUS" & color 0B & netsh wlan show hostednetwork & goto PAUSE_RETURN

:WIFIBLOCKED
cls & call :SECTION "BLOCKED NETWORKS" & color 0B & netsh wlan show blockednetworks & goto PAUSE_RETURN

:WIFIREPORT
cls & call :SECTION "GENERATE WLAN REPORT" & color 0E & echo  Generating (may take 10-30 sec)... & color 0B
netsh wlan show wlanreport & echo. & echo  Report saved to: C:\ProgramData\Microsoft\Windows\WlanReport\ & goto PAUSE_RETURN

:WIFICONNECT
cls & call :SECTION "CONNECT TO WIFI PROFILE" & color 0B & netsh wlan show profiles & echo.
set /p "cn= Profile name: " & if "!cn!"=="" goto PAUSE_RETURN
netsh wlan connect name="!cn!" & goto PAUSE_RETURN

:WIFIDISCONNECT
cls & call :SECTION "DISCONNECT WIFI" & color 0B & netsh wlan disconnect & call :OK "Disconnected" & goto PAUSE_RETURN

:WIFIAUTOCONFIG
cls & call :SECTION "WIFI AUTOCONFIG" & color 0B & netsh wlan show autoconfig
echo. & set /p "ac= Enable auto-config? (yes/no, blank=skip): "
if /i "!ac!"=="yes" ( netsh wlan set autoconfig enabled=yes interface="Wi-Fi" 2>nul & call :OK "Enabled" )
if /i "!ac!"=="no" ( netsh wlan set autoconfig enabled=no interface="Wi-Fi" 2>nul & call :OK "Disabled" )
goto PAUSE_RETURN

:WIFIPRIORITY
cls & call :SECTION "SET PROFILE PRIORITY" & color 0B & netsh wlan show profiles & echo.
set /p "pn= Profile name: " & set /p "pp= Priority (1=highest): "
netsh wlan set profileorder name="!pn!" interface="Wi-Fi" priority=!pp! & goto PAUSE_RETURN

:WIFIRANDOM
cls & call :SECTION "MAC RANDOMIZATION" & color 0B & netsh wlan show randomization & goto PAUSE_RETURN

:WIFICAPS
cls & call :SECTION "WIRELESS CAPABILITIES" & color 0B & netsh wlan show wirelesscapabilities & goto PAUSE_RETURN

:WIFISETTINGS
cls & call :SECTION "WIFI SETTINGS" & color 0B & netsh wlan show settings & goto PAUSE_RETURN

:WIFITRACE
cls & call :SECTION "WIFI TRACING" & color 0B & netsh wlan show tracing & goto PAUSE_RETURN

:: ==========================================================
:: REPAIR / RESET
:: ==========================================================
:FLUSHDNS
cls & call :SECTION "FLUSH DNS" & color 0B & ipconfig /flushdns & call :OK "Done" & goto PAUSE_RETURN
:RELEASE
cls & call :SECTION "RELEASE IP" & color 0B & ipconfig /release & call :OK "Done" & goto PAUSE_RETURN
:RENEW
cls & call :SECTION "RENEW IP" & color 0B & ipconfig /renew & call :OK "Done" & goto PAUSE_RETURN
:WINSOCK
cls & call :SECTION "RESET WINSOCK" & color 0B & netsh winsock reset & call :OK "Reboot required" & goto PAUSE_RETURN
:TCPRESET
cls & call :SECTION "RESET TCP/IP" & color 0B & netsh int ip reset & call :OK "Reboot required" & goto PAUSE_RETURN
:DNSRESET
cls & call :SECTION "RESET DNS" & color 0B & netsh int ip set dns "Local Area Connection" dhcp >nul 2>&1 & netsh int ip set dns "Wi-Fi" dhcp >nul 2>&1 & netsh int ip reset & call :OK "Done" & goto PAUSE_RETURN
:DHCPRESET
cls & call :SECTION "RESET DHCP" & color 0B & ipconfig /release & ipconfig /renew & netsh int ip reset & netsh winsock reset & call :OK "Reboot required" & goto PAUSE_RETURN
:IPV6RESET
cls & call :SECTION "RESET IPv6" & color 0B & netsh int ipv6 reset & call :OK "Reboot required" & goto PAUSE_RETURN
:FWRESET
cls & call :SECTION "RESET FIREWALL" & call :WARN "Removes custom rules" & set /p "cf= Type YES: "
if /i not "!cf!"=="YES" goto MENU & color 0B & netsh advfirewall reset & call :OK "Done" & goto PAUSE_RETURN
:PROXYRESET
cls & call :SECTION "RESET PROXY" & color 0B & netsh winhttp reset proxy & reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings" /v ProxyEnable /t REG_DWORD /d 0 /f >nul 2>&1 & call :OK "Done" & goto PAUSE_RETURN
:NETBIOSRESET
cls & call :SECTION "RESET NetBIOS" & color 0B & wmic nicconfig where TcpipNetbiosOptions=0 call SetTcpipNetbios 1 >nul 2>&1 & wmic nicconfig where TcpipNetbiosOptions=1 call SetTcpipNetbios 0 >nul 2>&1 & call :OK "Reboot recommended" & goto PAUSE_RETURN
:WINSRESET
cls & call :SECTION "RESET WINS CACHE" & color 0B & nbtstat -R & nbtstat -RR & call :OK "WINS cache cleared" & goto PAUSE_RETURN
:WLANRESTART
cls & call :SECTION "RESTART WLAN AutoConfig" & color 0B & net stop WlanSvc /y >nul 2>&1 & net start WlanSvc >nul 2>&1 & call :OK "Done" & goto PAUSE_RETURN
:ADAPTERRESTART
cls & call :SECTION "RESTART ADAPTERS" & call :WARN "Brief disconnect" & color 0B & powershell -NoProfile -Command "Get-NetAdapter | Restart-NetAdapter -Confirm:$false" & call :OK "Done" & goto PAUSE_RETURN
:EXPLORER
cls & call :SECTION "RESTART EXPLORER" & color 0B & taskkill /f /im explorer.exe >nul 2>&1 & timeout /t 2 >nul & start explorer.exe & call :OK "Done" & goto PAUSE_RETURN

:FULLREPAIR
cls & call :HEADER & call :WARN "FULL NETWORK REPAIR" & set /p "cf= Type YES: "
if /i not "!cf!"=="YES" goto MENU & echo.
color 0E & echo  [1/9] Flush DNS... & color 0B & ipconfig /flushdns
color 0E & echo  [2/9] Release... & color 0B & ipconfig /release >nul
color 0E & echo  [3/9] Renew... & color 0B & ipconfig /renew >nul
color 0E & echo  [4/9] Winsock... & color 0B & netsh winsock reset
color 0E & echo  [5/9] TCP/IP... & color 0B & netsh int ip reset
color 0E & echo  [6/9] IPv6... & color 0B & netsh int ipv6 reset
color 0E & echo  [7/9] Proxy... & color 0B & netsh winhttp reset proxy
color 0E & echo  [8/9] WLAN... & color 0B & net stop WlanSvc /y >nul 2>&1 & net start WlanSvc >nul 2>&1
color 0E & echo  [9/9] Flush DNS again... & color 0B & ipconfig /flushdns
color 0A & call :OK "FULL REPAIR COMPLETE -- REBOOT RECOMMENDED" & goto PAUSE_RETURN

:: ==========================================================
:: ADVANCED NETWORKING
:: ==========================================================
:PKTCAP
cls & call :SECTION "PACKET CAPTURE (pktmon)" & color 0E
echo  This captures packets for 15 seconds and exports to Desktop.
echo  Press CTRL+C to cancel.
echo.
set /p "cap= Type YES to start: "
if /i not "!cap!"=="YES" goto MENU
color 0B & pktmon start --capture --pkt-size 128 --file-name "%USERPROFILE%\Desktop\pktmon_capture.etl"
echo. & color 0E & echo  Capturing... (15 seconds) & timeout /t 15 >nul
color 0B & pktmon stop
echo. & call :OK "Capture saved to Desktop as pktmon_capture.etl" & goto PAUSE_RETURN

:SPEEDTEST
cls & call :SECTION "SPEED TEST (Ookla / Speedtest.net)" & color 0B
echo  Checking for speedtest-cli...
where speedtest >nul 2>&1
if errorlevel 1 call :INSTALL_SPEEDTEST
color 0B & echo  Running Speed Test...
speedtest & call :OK "Speed Test Complete" & goto PAUSE_RETURN

:PINGSWEEP
cls & call :SECTION "PING SWEEP (Host Discovery)" & color 0E
echo  Enter the first three octets of your subnet. Example: 192.168.1
set /p "subnet= Subnet: " & if "!subnet!"=="" goto MENU
color 0B & echo  Pinging !subnet!.1 to !subnet!.254 ...
for /l %%i in (1,1,254) do (
    ping -n 1 -w 100 !subnet!.%%i | find "TTL=" >nul && echo  Host UP: !subnet!.%%i
)
call :OK "Sweep complete" & goto PAUSE_RETURN

:PUBLICIP
cls & call :SECTION "PUBLIC IP & GEOIP" & color 0B
echo  Fetching your public IP...
powershell -NoProfile -Command "(Invoke-RestMethod -Uri 'https://ipinfo.io/json') | Format-List"
goto PAUSE_RETURN

:NETERRORS
cls & call :SECTION "NETWORK ERRORS (netstat -e)" & color 0B & netstat -e & goto PAUSE_RETURN

:MTUTEST
cls & call :SECTION "TEST MTU SIZE (Ping DF)" & color 0E
echo  Enter target (default: 8.8.8.8)
set /p "mtut= Target: " & if "!mtut!"=="" set "mtut=8.8.8.8"
echo. & echo  Testing max payload before fragmentation (Standard is 1500)...
color 0B & ping -f -l 1472 !mtut! & goto PAUSE_RETURN

:DOHCHECK
cls & call :SECTION "DNS OVER HTTPS (DoH) CHECK" & color 0B & netsh dns show global & goto PAUSE_RETURN

:FIREWALLADV
cls & call :SECTION "ADVANCED FIREWALL" & start wf.msc & goto PAUSE_RETURN

:: ==========================================================
:: PANELS
:: ==========================================================
:NCPA
cls & call :SECTION "NETWORK CONNECTIONS" & start ncpa.cpl & goto PAUSE_RETURN
:DEVICE
cls & call :SECTION "DEVICE MANAGER" & start devmgmt.msc & goto PAUSE_RETURN
:TROUBLE
cls & call :SECTION "NETWORK TROUBLESHOOTER" & color 0E & msdt.exe /id NetworkDiagnosticsNetworkAdapter & goto PAUSE_RETURN
:FIREWALLPANEL
cls & call :SECTION "FIREWALL" & start control firewall.cpl & goto PAUSE_RETURN
:PROXYPANEL
cls & call :SECTION "PROXY SETTINGS" & start ms-settings:network-proxy & goto PAUSE_RETURN
:NETRESET
cls & call :SECTION "WINDOWS NETWORK RESET" & start ms-settings:network-status & goto PAUSE_RETURN
:WU_NETWORK
cls & call :SECTION "WINDOWS UPDATE" & start ms-settings:windowsupdate & goto PAUSE_RETURN

:: ==========================================================
:: SECURITY CHECKS & LOCAL TOOLS
:: ==========================================================
:SMBCHECK
cls & call :SECTION "SMB SHARES & SESSIONS" & color 0B & net share & echo. & net session & goto PAUSE_RETURN

:FWRULES
cls & call :SECTION "WINDOWS FIREWALL ACTIVE RULES" & color 0B
netsh advfirewall firewall show rule name=all | findstr /C:"Rule Name" /C:"Enabled" /C:"Direction" /C:"Action" & goto PAUSE_RETURN

:SMBV1
cls & call :SECTION "SMBv1 PROTOCOL CHECK" & color 0B
powershell -NoProfile -Command "Get-WindowsOptionalFeature -Online -FeatureName SMB1Protocol | Select-Object FeatureName, State" 
call :WARN "If State is 'Enabled', you are vulnerable to WannaCry." & goto PAUSE_RETURN

:NTLM
cls & call :SECTION "NTLM AUDITING CHECK" & color 0B
reg query "HKLM\SYSTEM\CurrentControlSet\Control\Lsa" /v LmCompatibilityLevel & call :WARN "Level 3 or higher is recommended" & goto PAUSE_RETURN

:TCPPARAMS
cls & call :SECTION "TCP PARAMETERS (Registry)" & color 0B
reg query "HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters" & goto PAUSE_RETURN

:BINDINGS
cls & call :SECTION "NETWORK BINDINGS" & color 0B
powershell -NoProfile -Command "Get-NetAdapterBinding | Format-Table Name, DisplayName, Enabled -AutoSize" & goto PAUSE_RETURN

:: ==========================================================
:: QUICK DIAGNOSTIC & MONITORING
:: ==========================================================
:QUICKCONN
cls & call :SECTION "QUICK CONNECTIVITY TEST" & color 0B
echo  [1/3] Gateway... & ping -n 2 (ipconfig | findstr "Default Gateway" | for /f "tokens=2 delims=:" %%a in ('more') do @echo %%a) 2>nul
echo  [2/3] Google DNS (8.8.8.8)... & ping -n 2 8.8.8.8
echo  [3/3] google.com (DNS)... & ping -n 2 google.com
call :OK "Quick test done" & goto PAUSE_RETURN

:QUICKDNS
cls & call :SECTION "QUICK DNS TEST" & color 0B
echo  Default DNS: & ipconfig /all | findstr /C:"DNS Servers"
echo. & echo  Resolving google.com... & nslookup google.com
call :OK "DNS test done" & goto PAUSE_RETURN

:QUICKWIFI
cls & call :SECTION "QUICK WIFI SIGNAL TEST" & color 0B & netsh wlan show interfaces & goto PAUSE_RETURN

:QUICKPORTSCAN
cls & call :SECTION "QUICK PORT SCAN (localhost 1-1024)" & call :WARN "This may take a minute" & color 0B
powershell -NoProfile -Command "1..1024 | ForEach-Object { $r = Test-NetConnection -ComputerName 127.0.0.1 -Port $_ -WarningAction SilentlyContinue; if ($r.TcpTestSucceeded) { Write-Host \"Port $_ : OPEN\" } }"
call :OK "Scan complete" & goto PAUSE_RETURN

:WOL
cls & call :SECTION "WAKE-ON-LAN (WOL)" & color 0E
echo  Enter the MAC Address to wake (format: XX-XX-XX-XX-XX-XX)
set /p "mac= MAC: " & if "!mac!"=="" goto MENU
color 0B & powershell -NoProfile -Command "$mac = '!mac!'.Replace('-','').Replace(':',''); $bytes = [byte[]](0..5 | ForEach-Object { 0xFF }) + [byte[]](0..15 | ForEach-Object { [Convert]::ToByte($mac.Substring($_*2, 2), 16) }); $udp = New-Object System.Net.Sockets.UdpClient; $udp.Connect([System.Net.IPAddress]::Broadcast, 9); $udp.Send($bytes, $bytes.Length) | Out-Null; $udp.Close(); Write-Host 'Magic Packet Sent!'"
call :OK "WOL Packet Broadcasted" & goto PAUSE_RETURN

:QUICKALLDNS
cls & call :SECTION "TEST ALL MAJOR DNS" & color 0B
echo  8.8.8.8... & ping -n 2 8.8.8.8
echo  8.8.4.4... & ping -n 2 8.8.4.4
echo  1.1.1.1... & ping -n 2 1.1.1.1
echo  9.9.9.9... & ping -n 2 9.9.9.9
call :OK "Done" & goto PAUSE_RETURN

:DHCPLEASE
cls & call :SECTION "DHCP LEASE INFO" & color 0B
ipconfig /all | findstr /C:"Lease Obtained" /C:"Lease Expires" /C:"DHCP Server" /C:"Physical Address" & goto PAUSE_RETURN

:WIFIQR
cls & call :SECTION "GENERATE WIFI QR CODE" & color 0E
set /p "ssid= WiFi Name (SSID): " & set /p "pass= Password: "
if "!ssid!"=="" goto MENU
color 0B
powershell -NoProfile -Command "if (-not (Get-Command New-QRCodeText -ErrorAction SilentlyContinue)) { Install-Module -Name QRCodeGenerator -Force -Scope CurrentUser -AllowClobber -ErrorAction SilentlyContinue } $url = \"https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=WIFI:T:WPA;S:$ssid;P:$pass;;\"; Start-Process $url"
call :OK "Opened QR Code in Browser" & goto PAUSE_RETURN

:BANDWIDTH
cls & call :SECTION "BANDWIDTH MONITOR (Live)" & color 0E
echo  Monitoring network usage. Press CTRL+C to stop.
echo.
powershell -NoProfile -Command "while($true) { $a = Get-NetAdapterStatistics; $rx = ($a | Measure-Object -Property ReceivedBytes -Sum).Sum; $tx = ($a | Measure-Object -Property SentBytes -Sum).Sum; Start-Sleep -Seconds 1; $a2 = Get-NetAdapterStatistics; $rx2 = ($a2 | Measure-Object -Property ReceivedBytes -Sum).Sum; $tx2 = ($a2 | Measure-Object -Property SentBytes -Sum).Sum; $drx = [math]::Round(($rx2 - $rx)/1MB, 2); $dtx = [math]::Round(($tx2 - $tx)/1MB, 2); Write-Host -NoNewline \"``rDownload: $drx MB/s ^| Upload: $dtx MB/s   \" }"
goto PAUSE_RETURN

:DNSBENCH
cls & call :SECTION "DNS BENCHMARK (Fastest)" & color 0E
echo  Testing response times for major DNS providers...
echo.
color 0B
powershell -NoProfile -Command "$providers = @{'Google'='8.8.8.8'; 'Cloudflare'='1.1.1.1'; 'Quad9'='9.9.9.9'; 'OpenDNS'='208.67.222.222'}; foreach ($p in $providers.GetEnumerator()) { $t = Measure-Command { Resolve-DnsName -Name google.com -Server $p.Value -ErrorAction SilentlyContinue }; Write-Host \"$($p.Key) [$($p.Value)] : $([math]::Round($t.TotalMilliseconds, 2)) ms\" }"
call :OK "Benchmark Complete" & goto PAUSE_RETURN

:TCPING
cls & call :SECTION "TCPING (Traceroute via TCP)" & color 0E
echo  Tests TCP port connectivity (bypasses ICMP blocks).
set /p "tcph= Host (default google.com): " & if "!tcph!"=="" set "tcph=google.com"
set /p "tcpp= Port (default 443): " & if "!tcpp!"=="" set "tcpp=443"
color 0B
powershell -NoProfile -Command "if (-not (Get-Command tnc -ErrorAction SilentlyContinue)) { Write-Host 'Test-NetConnection is built-in.' } 1..5 | ForEach-Object { $r = Test-NetConnection -ComputerName '!tcph!' -Port !tcpp! -WarningAction SilentlyContinue; if ($r.TcpTestSucceeded) { Write-Host 'Port Open: TCP !tcpp! on !tcph!' } else { Write-Host 'Port Closed/Filtered' }; Start-Sleep -Seconds 1 }"
call :OK "TCPing Complete" & goto PAUSE_RETURN

:: ==========================================================
:: NEW TOOLS ADDED (100-111)
:: ==========================================================
:MACLOOKUP
cls & call :SECTION "MAC ADDRESS LOOKUP (Vendor)" & color 0E
echo  Enter MAC Address (e.g., 00-1A-2B-3C-4D-5E)
set /p "mac= MAC: " & if "!mac!"=="" goto MENU
where python >nul 2>&1
if errorlevel 1 (
    call :WARN "Python is required for this tool."
    goto PAUSE_RETURN
)
python -c "import mac_vendor_lookup" >nul 2>&1
if errorlevel 1 call :INSTALL_MACLOOKUP
color 0B & echo  Querying vendor...
python -c "from mac_vendor_lookup import MacLookup; print(MacLookup().lookup('!mac!'))" 2>nul
call :OK "Lookup Complete" & goto PAUSE_RETURN

:PUBDNS
cls & call :SECTION "CHECK PUBLIC DNS" & color 0B
echo  Testing connectivity to Google (8.8.8.8) and Cloudflare (1.1.1.1)...
echo.
echo  Google DNS...
ping -n 2 8.8.8.8
echo.
echo  Cloudflare DNS...
ping -n 2 1.1.1.1
call :OK "Done" & goto PAUSE_RETURN

:DOMAINIP
cls & call :SECTION "FIND IP OF A DOMAIN" & color 0E
set /p "dom= Enter Domain (e.g., google.com): " & if "!dom!"=="" goto MENU
color 0B & nslookup "!dom!" & goto PAUSE_RETURN

:REVDNS
cls & call :SECTION "REVERSE DNS (PTR)" & color 0E
set /p "ip= Enter IP Address: " & if "!ip!"=="" goto MENU
color 0B & nslookup "!ip!" & goto PAUSE_RETURN

:UDPPORTS
cls & call :SECTION "VIEW OPEN UDP PORTS" & color 0B
netstat -an | findstr "UDP" & goto PAUSE_RETURN

:TCPPORTS
cls & call :SECTION "VIEW OPEN TCP PORTS" & color 0B
netstat -an | findstr "TCP" & goto PAUSE_RETURN

:MULTICAST
cls & call :SECTION "ACTIVE MULTICAST GROUPS" & color 0B
netsh interface ipv4 show joins & goto PAUSE_RETURN

:IPV6ADDR
cls & call :SECTION "IPv6 GLOBAL ADDRESSES" & color 0B
netsh interface ipv6 show address & goto PAUSE_RETURN

:HTTPCONN
cls & call :SECTION "ACTIVE HTTP CONNECTIONS" & color 0B
netstat -ano | findstr ":80" & netstat -ano | findstr ":443" & goto PAUSE_RETURN

:PROXYWPAD
cls & call :SECTION "CHECK HTTP PROXY AUTO-DETECT" & color 0B
netsh winhttp show proxy & goto PAUSE_RETURN

:IPERF
cls & call :SECTION "TEST NETWORK PORT SPEED (iPerf3)" & color 0E
echo  Requires iPerf3 installed and a server running.
echo  (e.g., iperf3 -s on another PC).
set /p "iperf_ip= Server IP: " & if "!iperf_ip!"=="" goto MENU
color 0B
where iperf3 >nul 2>&1
if errorlevel 1 (
    call :WARN "iPerf3 not found. Please download from iperf.fr and add to PATH."
    goto PAUSE_RETURN
)
iperf3 -c !iperf_ip!
call :OK "Speed Test Complete" & goto PAUSE_RETURN

:: ==========================================================
:: QUIT
:: ==========================================================
:QUIT
cls & color 0A & echo. & echo  ================================================== & echo    Thank you for using Network Toolkit v8.1 & echo  ==================================================
echo. & echo  Closing in 3 sec... & timeout /t 3 >nul & exit /b
