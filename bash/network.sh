#!/bin/bash
# ==========================================================
# Advanced Linux Network Toolkit v1.0 - Ultimate Edition
# ==========================================================

# Check for root
if [ "$EUID" -ne 0 ]; then
  echo -e "\e[31m[!] ERROR: Please run this script as root (sudo ./network.sh)\e[0m"
  exit 1
fi

BLINK_ON=0
OS_TYPE=""

# Detect Distro and Package Manager
if command -v apt-get >/dev/null 2>&1; then OS_TYPE="debian"
elif command -v dnf >/dev/null 2>&1; then OS_TYPE="fedora"
elif command -v pacman >/dev/null 2>&1; then OS_TYPE="arch"
else OS_TYPE="unknown"
fi

# Dependency Installer Function
install_pkg() {
    local pkg=$1
    echo -e "\e[33m[!] $pkg not found. Attempting to install...\e[0m"
    case $OS_TYPE in
        debian) apt-get update -qq && apt-get install -y "$pkg" >/dev/null 2>&1 ;;
        fedora) dnf install -y "$pkg" >/dev/null 2>&1 ;;
        arch) pacman -S --noconfirm "$pkg" >/dev/null 2>&1 ;;
        *) echo -e "\e[31m[!] Unknown package manager. Please install $pkg manually.\e[0m"; return 1 ;;
    esac
    echo -e "\e[32m[OK] $pkg installed successfully.\e[0m"
}

# Dependency Checker
check_dep() {
    if ! command -v "$1" >/dev/null 2>&1; then
        install_pkg "$2"
    fi
}

pause_return() {
    echo -e "\n\e[35m------------------------------------------------\e[0m"
    echo -e "\e[35m   Press any key to return...\e[0m"
    read -n 1 -s -r
    menu
}

header() {
    BLINK_ON=$(( (BLINK_ON + 1) % 2 ))
    echo -e "\e[1;37m========================================================================================================================================\e[0m"
    if [ $BLINK_ON -eq 1 ]; then
        echo -e "\e[1;33m                     A D V A N C E D   L I N U X   N E T W O R K   T O O L K I T\e[0m"
    else
        echo -e "\e[1;36m                     A D V A N C E D   L I N U X   N E T W O R K   T O O L K I T\e[0m"
    fi
    echo -e "\e[1;31m                                  ~ Diagnose. Repair. Dominate. ~\e[0m"
    echo -e "\e[1;37m========================================================================================================================================\e[0m"
}

section() {
    echo -e "\n\e[1;33m==================================================\e[0m"
    echo -e "\e[1;33m   $1\e[0m"
    echo -e "\e[1;33m==================================================\e[0m\n"
}

menu() {
    clear
    header
    echo ""
    echo -e "\e[1;33m  +----------------------------------------+  +----------------------------------------+  +----------------------------------------+\e[0m"
    echo -e "\e[1;33m  |            DIAGNOSTICS                 |  |             NETWORK TOOLS              |  |           REPAIR / RESET               |\e[0m"
    echo -e "\e[1;33m  +----------------------------------------+  +----------------------------------------+  +----------------------------------------+\e[0m"
    echo -e "\e[1;36m  |  [1]  Full IP Configuration            |  |  [25] Open WiFi Settings (nmtui)       |  |  [45] Flush DNS Cache                  |\e[0m"
    echo -e "\e[1;36m  |  [2]  Ping Google                      |  |  [26] WiFi Profile / Password          |  |  [46] Restart Network Manager          |\e[0m"
    echo -e "\e[1;36m  |  [3]  Continuous Ping (8.8.8.8)        |  |  [27] Show All WiFi Profiles           |  |  [47] Renew DHCP Lease                 |\e[0m"
    echo -e "\e[1;36m  |  [4]  Routing Table                    |  |  [28] Show WiFi Interfaces             |  |  [48] Restart Network Interfaces       |\e[0m"
    echo -e "\e[1;36m  |  [5]  ARP Cache                        |  |  [29] Scan WiFi Networks               |  |  [49] Reset Firewall (iptables)        |\e[0m"
    echo -e "\e[1;36m  |  [6]  Active Sessions                  |  |  [30] Export Profiles to Desktop       |  |  [50] Reset IPv6 Stack                 |\e[0m"
    echo -e "\e[1;36m  |  [7]  DNS Cache                        |  |  [31] WiFi Drivers Info                |  |  [51] Clear ARP Cache                  |\e[0m"
    echo -e "\e[1;36m  |  [8]  System Network Info              |  |  [32] WiFi Radio / Signal Status       |  |  [52] FULL NETWORK REPAIR  << REC      |\e[0m"
    echo -e "\e[1;36m  |  [9]  Test-Connection (Ping)           |  |  [33] WiFi Hosted Network Status       |  |                                        |\e[0m"
    echo -e "\e[1;36m  |  [10] Test-NetConnection (Port Check)  |  |  [34] WiFi Blocked Networks            |  |  [99] EXIT                             |\e[0m"
    echo -e "\e[1;36m  |  [11] Trace Route (traceroute)        |  |  [35] WiFi WLAN Report                 |  |                                        |\e[0m"
    echo -e "\e[1;36m  |  [12] PathPing (MTR)                  |  |  [36] Connect to WiFi Profile          |  |                                        |\e[0m"
    echo -e "\e[1;36m  |  [13] DNS Lookup (dig)                |  |  [37] Disconnect WiFi                  |  |                                        |\e[0m"
    echo -e "\e[1;36m  |  [14] NetBIOS Cache (nbtstat)          |  |  [38] WiFi AutoConfig On/Off           |  |                                        |\e[0m"
    echo -e "\e[1;36m  |  [15] Check Listening Ports            |  |  [39] WiFi Profile Priority            |  |                                        |\e[0m"
    echo -e "\e[1;36m  |  [16] DNS Server Check                 |  |  [40] WiFi Random MAC Status           |  |                                        |\e[0m"
    echo -e "\e[1;36m  |  [17] Get IP Configuration             |  |  [41] WiFi Wireless Capabilities       |  |                                        |\e[0m"
    echo -e "\e[1;36m  |  [18] Get Adapter Status               |  |  [42] WiFi Settings Overview           |  |                                        |\e[0m"
    echo -e "\e[1;36m  |  [19] Get TCP Connections              |  |  [43] WiFi Tracing On/Off              |  |                                        |\e[0m"
    echo -e "\e[1;36m  |  [20] Get DNS Client Server            |  |  [44] Export WiFi Profiles             |  |                                        |\e[0m"
    echo -e "\e[1;36m  |  [21] Get Adapter Binding              |  |                                        |  |                                        |\e[0m"
    echo -e "\e[1;36m  |  [22] Test DNS Server                  |  |                                        |  |                                        |\e[0m"
    echo -e "\e[1;36m  |  [23] Get Route Table                  |  |                                        |  |                                        |\e[0m"
    echo -e "\e[1;36m  |  [24] Get IP Address                   |  |                                        |  |                                        |\e[0m"
    echo -e "\e[1;33m  +----------------------------------------+  +----------------------------------------+  +----------------------------------------+\e[0m"
    echo -e "\e[1;33m  |          ADVANCED NETWORKING           |  |               MISC / INFO              |  |             QUICK DIAGNOSTIC           |\e[0m"
    echo -e "\e[1;33m  +----------------------------------------+  +----------------------------------------+  +----------------------------------------+\e[0m"
    echo -e "\e[1;36m  |  [60] Packet Capture (tcpdump)         |  |  [68] Network Connections              |  |  [85] Quick Connectivity Test          |\e[0m"
    echo -e "\e[1;36m  |  [61] Speed Test (Ookla)               |  |  [69] Hardware Info                    |  |  [86] Quick DNS Test                   |\e[0m"
    echo -e "\e[1;36m  |  [62] Ping Sweep (Host Discovery)      |  |  [70] System Logs                     |  |  [87] Quick WiFi Signal Test           |\e[0m"
    echo -e "\e[1;36m  |  [63] Find Public IP & GeoIP           |  |  [71] Firewall Status                  |  |  [88] Quick Port Scan (1-1024)         |\e[0m"
    echo -e "\e[1;36m  |  [64] Show Network Errors              |  |  [72] Proxy Settings                   |  |  [89] Wake-on-LAN (WOL)                |\e[0m"
    echo -e "\e[1;36m  |  [65] Test MTU Size                    |  |  [73] Network Reset (Linux)            |  |  [90] Test ALL Google DNS              |\e[0m"
    echo -e "\e[1;36m  |  [66] Check DNS over HTTPS             |  |  [74] Update Packages                  |  |  [91] Show DHCP Lease Info             |\e[0m"
    echo -e "\e[1;36m  |  [67] Advanced Firewall (ufw/iptables) |  |                                        |  |  [92] Show WiFi QR Code                |\e[0m"
    echo -e "\e[1;33m  +----------------------------------------+  +----------------------------------------+  +----------------------------------------+\e[0m"
    echo -e "\e[1;33m  |            SECURITY CHECKS             |  |          NETWORK MONITORING            |  |             LOCAL TOOLS                |\e[0m"
    echo -e "\e[1;33m  +----------------------------------------+  +----------------------------------------+  +----------------------------------------+\e[0m"
    echo -e "\e[1;36m  |  [75] Show Samba Shares                |  |  [93] Bandwidth Monitor (Live)         |  |  [96] View DNS Cache                   |\e[0m"
    echo -e "\e[1;36m  |  [76] Firewall Active Rules            |  |  [94] DNS Benchmark (Fastest)          |  |  [97] View ARP Cache                   |\e[0m"
    echo -e "\e[1;36m  |  [77] SMBv1 Protocol Check             |  |  [95] Traceroute via TCP (tcping)      |  |  [98] View Routing Table               |\e[0m"
    echo -e "\e[1;36m  |  [78] NTLM Auditing Check              |  |                                        |  |                                        |\e[0m"
    echo -e "\e[1;36m  |  [79] Show Kernel Parameters           |  |                                        |  |                                        |\e[0m"
    echo -e "\e[1;36m  |  [80] Show Network Interfaces          |  |                                        |  |                                        |\e[0m"
    echo -e "\e[1;33m  +----------------------------------------+  +----------------------------------------+  +----------------------------------------+\e[0m"
    echo -e "\e[1;33m  |            NEW TOOLS ADDED             |  |               NEW TOOLS                |  |               NEW TOOLS                |\e[0m"
    echo -e "\e[1;33m  +----------------------------------------+  +----------------------------------------+  +----------------------------------------+\e[0m"
    echo -e "\e[1;36m  |  [100] MAC Address Lookup (Vendor)     |  |  [104] View Open UDP Ports             |  |  [108] Get Active HTTP Connections     |\e[0m"
    echo -e "\e[1;36m  |  [101] Check Public DNS                |  |  [105] View Open TCP Ports             |  |  [109] Check HTTP Proxy Auto-Detect    |\e[0m"
    echo -e "\e[1;36m  |  [102] Find IP of a Domain             |  |  [106] Show Active Multicast Groups    |  |  [110] Test Network Port Speed (iPerf) |\e[0m"
    echo -e "\e[1;36m  |  [103] Resolve Reverse DNS (PTR)       |  |  [107] Show IPv6 Global Addresses      |  |  [111] Reset WINS Cache                |\e[0m"
    echo -e "\e[1;33m  +----------------------------------------+  +----------------------------------------+  +----------------------------------------+\e[0m"
    echo ""
    read -p "  Select Option: " choice
    echo ""

    case $choice in
        1) section "FULL IP CONFIGURATION"; ip a; pause_return ;;
        2) section "PING GOOGLE (10)"; ping -c 10 google.com; pause_return ;;
        3) section "CONTINUOUS PING 8.8.8.8"; echo -e "\e[31m[!] Press CTRL+C to stop\e[0m"; ping 8.8.8.8; pause_return ;;
        4) section "ROUTING TABLE"; ip route; pause_return ;;
        5) section "ARP CACHE"; ip neigh; pause_return ;;
        6) section "ACTIVE SESSIONS"; ss -tulpn; pause_return ;;
        7) section "DNS CACHE"; if command -v systemd-resolve >/dev/null; then systemd-resolve --statistics; else echo "systemd-resolved not found."; fi; pause_return ;;
        8) section "SYSTEM NETWORK INFO"; hostnamectl; echo ""; ip -br a; pause_return ;;
        9) section "Test-Connection (Ping)"; read -p "Enter target (default 8.8.8.8): " tg; tg=${tg:-8.8.8.8}; ping -c 4 $tg; pause_return ;;
        10) section "Test-NetConnection (Port Check)"; read -p "Host (default google.com): " th; th=${th:-google.com}; read -p "Port (default 443): " tp; tp=${tp:-443}; nc -zv $th $tp; pause_return ;;
        11) section "TRACE ROUTE"; check_dep traceroute traceroute; read -p "Target (default 8.8.8.8): " tt; tt=${tt:-8.8.8.8}; traceroute $tt; pause_return ;;
        12) section "MTR (Latency + Packet Loss per Hop)"; check_dep mtr mtr; read -p "Target (default 8.8.8.8): " tm; tm=${tm:-8.8.8.8}; mtr -r -c 10 $tm; pause_return ;;
        13) section "DNS LOOKUP (dig)"; check_dep dig dnsutils; read -p "Domain (default google.com): " nd; nd=${nd:-google.com}; dig $nd; pause_return ;;
        14) section "NetBIOS CACHE"; check_dep nbtscan nbtscan; read -p "IP to scan: " nbtip; nbtscan $nbtip; pause_return ;;
        15) section "LISTENING PORTS"; ss -tulpn | grep LISTEN; pause_return ;;
        16) section "DNS Server Check"; cat /etc/resolv.conf; pause_return ;;
        17) section "Get IP Configuration"; ip a; pause_return ;;
        18) section "Get Adapter Status"; ip link show; pause_return ;;
        19) section "Get TCP Connections"; ss -t; pause_return ;;
        20) section "Get DNS Client Server"; cat /etc/resolv.conf; pause_return ;;
        21) section "Get Adapter Binding"; ethtool -i $(ip route | grep default | awk '{print $5}'); pause_return ;;
        22) section "Test DNS Server"; read -p "Domain to test (default google.com): " td; td=${td:-google.com}; dig $td; pause_return ;;
        23) section "Get Route Table"; ip route; pause_return ;;
        24) section "Get IP Address"; ip addr show; pause_return ;;

        25) section "OPEN WIFI SETTINGS"; nmtui; pause_return ;;
        26) section "WIFI PROFILE / PASSWORD"; read -p "Enter WiFi SSID: " ssid; if [ -f "/etc/NetworkManager/system-connections/$ssid.nmconnection" ]; then grep psk= "/etc/NetworkManager/system-connections/$ssid.nmconnection" | cut -d'=' -f2; else echo "Profile not found."; fi; pause_return ;;
        27) section "ALL WIFI PROFILES"; ls /etc/NetworkManager/system-connections/; pause_return ;;
        28) section "WIFI INTERFACES"; nmcli device status; pause_return ;;
        29) section "SCANNING WIFI NETWORKS"; nmcli device wifi list; pause_return ;;
        30) section "EXPORT WIFI PROFILES"; cp -r /etc/NetworkManager/system-connections/ ~/Desktop/WiFi_Export/; echo "Copied to Desktop."; pause_return ;;
        31) section "WIFI DRIVERS INFO"; lspci -k | grep -A 3 -i "network"; pause_return ;;
        32) section "WIFI RADIO / SIGNAL"; nmcli -f SSID,SIGNAL,BARS,SECURITY device wifi list; pause_return ;;
        33) section "HOSTED NETWORK STATUS"; nmcli connection show; pause_return ;;
        34) section "BLOCKED NETWORKS"; rfkill list; pause_return ;;
        35) section "WLAN REPORT"; nmcli general; pause_return ;;
        36) section "CONNECT TO WIFI"; read -p "SSID: " ssid; read -p "Password: " pass; nmcli device wifi connect "$ssid" password "$pass"; pause_return ;;
        37) section "DISCONNECT WIFI"; nmcli networking off; echo "Disconnected"; pause_return ;;
        38) section "WIFI AUTOCONFIG"; nmcli radio wifi on; pause_return ;;
        39) section "PROFILE PRIORITY"; nmcli connection modify "$(nmcli -t -f NAME connection show --active)" connection.autoconnect-priority 10; echo "Priority set."; pause_return ;;
        40) section "MAC RANDOMIZATION"; macchanger -s $(ip route | grep default | awk '{print $5}'); pause_return ;;
        41) section "WIRELESS CAPABILITIES"; iw list; pause_return ;;
        42) section "WIFI SETTINGS"; nmcli connection show; pause_return ;;
        43) section "WIFI TRACING"; echo "Use dmesg | grep -i wifi to trace."; pause_return ;;
        44) section "EXPORT WIFI PROFILES"; cp -r /etc/NetworkManager/system-connections/ ~/Desktop/WiFi_Export/; pause_return ;;

        45) section "FLUSH DNS"; resolvectl flush-caches; systemctl restart systemd-resolved; echo "DNS Flushed"; pause_return ;;
        46) section "RESTART NETWORK MANAGER"; systemctl restart NetworkManager; echo "Restarted"; pause_return ;;
        47) section "RENEW DHCP"; dhclient -r && dhclient; echo "Renewed"; pause_return ;;
        48) section "RESTART ADAPTERS"; systemctl restart networking; pause_return ;;
        49) section "RESET FIREWALL"; iptables -F; iptables -X; echo "iptables reset."; pause_return ;;
        50) section "RESET IPv6"; sysctl -w net.ipv6.conf.all.disable_ipv6=0; pause_return ;;
        51) section "CLEAR ARP CACHE"; ip -s -s neigh flush all; echo "ARP cache cleared"; pause_return ;;
        52) section "FULL REPAIR"; resolvectl flush-caches; systemctl restart NetworkManager; echo "Repair complete."; pause_return ;;

        60) section "PACKET CAPTURE (tcpdump)"; check_dep tcpdump tcpdump; read -p "Interface (default eth0): " iface; iface=${iface:-eth0}; timeout 15 tcpdump -i $iface -w ~/capture.pcap; echo "Saved to ~/capture.pcap"; pause_return ;;
        61) section "SPEED TEST"; check_dep speedtest speedtest-cli; speedtest; pause_return ;;
        62) section "PING SWEEP"; read -p "Enter subnet (e.g., 192.168.1): " sub; for i in {1..254}; do ping -c 1 -W 1 $sub.$i | grep "64 bytes" & done; wait; pause_return ;;
        63) section "PUBLIC IP & GEOIP"; check_dep curl curl; curl ipinfo.io; pause_return ;;
        64) section "NETWORK ERRORS"; ip -s link; pause_return ;;
        65) section "TEST MTU SIZE"; read -p "Target (default 8.8.8.8): " mtut; mtut=${mtut:-8.8.8.8}; ping -M do -s 1472 -c 2 $mtut; pause_return ;;
        66) section "DNS OVER HTTPS CHECK"; resolvectl status | grep -i "dns over https"; pause_return ;;
        67) section "ADVANCED FIREWALL"; if command -v ufw >/dev/null; then ufw status verbose; else iptables -L -n -v; fi; pause_return ;;

        68) section "NETWORK CONNECTIONS"; nmcli connection show; pause_return ;;
        69) section "HARDWARE INFO"; lspci | grep -i network; pause_return ;;
        70) section "SYSTEM LOGS"; journalctl -u NetworkManager --no-pager | tail -n 20; pause_return ;;
        71) section "FIREWALL STATUS"; ufw status; pause_return ;;
        72) section "PROXY SETTINGS"; env | grep -i proxy; pause_return ;;
        73) section "NETWORK RESET"; systemctl restart NetworkManager; pause_return ;;
        74) section "UPDATE PACKAGES"; apt-get update && apt-get upgrade -y; pause_return ;;

        75) section "SAMBA SHARES"; check_dep smbclient smbclient; smbclient -L localhost -N; pause_return ;;
        76) section "FIREWALL ACTIVE RULES"; iptables -L -n -v; pause_return ;;
        77) section "SMBv1 CHECK"; testparm -s | grep -i "server min protocol"; pause_return ;;
        78) section "NTLM AUDITING CHECK"; grep -i "ntlm" /etc/samba/smb.conf; pause_return ;;
        79) section "KERNEL PARAMETERS"; sysctl -a | grep net.ipv4; pause_return ;;
        80) section "NETWORK INTERFACES"; ip link show; pause_return ;;

        85) section "QUICK CONNECTIVITY TEST"; ping -c 2 $(ip route | grep default | awk '{print $3}'); ping -c 2 8.8.8.8; ping -c 2 google.com; pause_return ;;
        86) section "QUICK DNS TEST"; cat /etc/resolv.conf; dig google.com; pause_return ;;
        87) section "QUICK WIFI SIGNAL TEST"; nmcli -f SSID,SIGNAL device wifi; pause_return ;;
        88) section "QUICK PORT SCAN"; check_dep nmap nmap; nmap -p 1-1024 localhost; pause_return ;;
        89) section "WAKE-ON-LAN (WOL)"; check_dep wakeonlan wakeonlan; read -p "MAC Address: " mac; wakeonlan $mac; pause_return ;;
        90) section "TEST ALL MAJOR DNS"; ping -c 2 8.8.8.8; ping -c 2 1.1.1.1; ping -c 2 9.9.9.9; pause_return ;;
        91) section "DHCP LEASE INFO"; cat /var/lib/dhcp/dhclient.leases; pause_return ;;
        92) section "WIFI QR CODE"; check_dep qrencode qrencode; read -p "SSID: " ssid; read -p "Password: " pass; qrencode -t ANSIUTF8 "WIFI:T:WPA;S:$ssid;P:$pass;;"; pause_return ;;

        93) section "BANDWIDTH MONITOR"; check_dep iftop iftop; iftop; pause_return ;;
        94) section "DNS BENCHMARK"; check_dep namebench namebench; namebench; pause_return ;;
        95) section "TRACEROUTE VIA TCP (tcping)"; check_dep tcptraceroute tcptraceroute; read -p "Host: " th; tcptraceroute $th 443; pause_return ;;

        96) section "VIEW DNS CACHE"; resolvectl statistics; pause_return ;;
        97) section "VIEW ARP CACHE"; ip neigh; pause_return ;;
        98) section "VIEW ROUTING TABLE"; ip route; pause_return ;;

        100) section "MAC LOOKUP"; check_dep curl curl; read -p "MAC Address: " mac; curl -s "https://api.macvendors.com/$mac"; echo ""; pause_return ;;
        101) section "CHECK PUBLIC DNS"; ping -c 2 8.8.8.8; ping -c 2 1.1.1.1; pause_return ;;
        102) section "FIND IP OF DOMAIN"; check_dep dig dnsutils; read -p "Domain: " dom; dig +short $dom; pause_return ;;
        103) section "REVERSE DNS"; check_dep dig dnsutils; read -p "IP: " ip; dig -x $ip +short; pause_return ;;
        104) section "VIEW OPEN UDP PORTS"; ss -ulpn; pause_return ;;
        105) section "VIEW OPEN TCP PORTS"; ss -tlpn; pause_return ;;
        106) section "ACTIVE MULTICAST GROUPS"; ip maddr show; pause_return ;;
        107) section "IPv6 ADDRESSES"; ip -6 addr show; pause_return ;;
        108) section "ACTIVE HTTP CONNECTIONS"; ss -tupn | grep -E ':80|:443'; pause_return ;;
        109) section "CHECK HTTP PROXY"; env | grep -i proxy; pause_return ;;
        110) section "IPERF3 TEST"; check_dep iperf3 iperf3; read -p "Server IP: " sip; iperf3 -c $sip; pause_return ;;
        111) section "RESET WINS CACHE"; nbtstat -R 2>/dev/null || echo "nbtstat not found (Linux uses different NetBIOS tools)"; pause_return ;;

        99) clear; exit 0 ;;
        *) echo -e "\e[31mInvalid selection.\e[0m"; sleep 2; menu ;;
    esac
}

# Start Script
menu
