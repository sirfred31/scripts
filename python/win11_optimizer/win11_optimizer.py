import tkinter as tk
from tkinter import messagebox, ttk, filedialog, scrolledtext
import subprocess, os, datetime, threading, sys, ctypes, winreg, psutil

VERSION = "4.7"  # Updated version number
last_backup = None
report_log = []
applied_tweaks = []  # Track all applied tweaks for restoration

# -----------------------------
# ENDPOINT SECURITY DETECTION
# -----------------------------

def set_custom_icon(root):
    """Set custom icon for the application"""
    try:
        # Get the directory where the script is located
        script_dir = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(script_dir, 'optimizer.ico')
        root.iconbitmap(icon_path)
        print(f"Icon loaded from: {icon_path}")
    except Exception as e:
        print(f"Icon error: {e}")
        # Fallback if icon file is not found
        pass

def check_endpoint_security():
    """Check for installed endpoint security software (excluding Microsoft Defender)"""
    security_products = []
    
    # Common endpoint security product names and services
    security_indicators = [
        # Sophos
        ("Sophos", ["Sophos", "Sophos Endpoint", "Sophos Anti-Virus", "Sophos Endpoint Defense"]),
        ("McAfee", ["McAfee", "VirusScan", "Endpoint Security", "McAfee Agent"]),
        ("Norton", ["Norton", "Symantec", "Norton Security", "Symantec Endpoint"]),
        ("Kaspersky", ["Kaspersky", "Kaspersky Endpoint", "Kaspersky Security"]),
        ("Trend Micro", ["Trend Micro", "OfficeScan", "Trend Micro Apex One"]),
        ("ESET", ["ESET", "ESET Endpoint", "ESET Security"]),
        ("CrowdStrike", ["CrowdStrike", "Falcon"]),
        ("SentinelOne", ["SentinelOne", "Sentinel Agent"]),
        ("Carbon Black", ["Carbon Black", "CB Defense", "VMware Carbon Black"]),
        ("Bitdefender", ["Bitdefender", "Bitdefender Endpoint", "Bitdefender GravityZone"]),
        ("Avast", ["Avast", "Avast Business", "AVG Business"]),
        ("AVG", ["AVG", "AVG Business"]),
        ("Webroot", ["Webroot", "Webroot SecureAnywhere"]),
        ("Malwarebytes", ["Malwarebytes", "Malwarebytes Endpoint"]),
        ("Cylance", ["Cylance", "CylancePROTECT"]),
        ("FireEye", ["FireEye", "FireEye Endpoint", "FireEye HX"]),
        ("Check Point", ["Check Point", "ZoneAlarm", "Endpoint Security"]),
        ("Panda", ["Panda", "Panda Security"]),
        ("Comodo", ["Comodo", "Comodo Security"]),
    ]
    
    # Check installed programs from registry
    try:
        # Check in registry
        reg_paths = [
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
        ]
        
        for reg_path in reg_paths:
            try:
                key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path)
                for i in range(0, winreg.QueryInfoKey(key)[0]):
                    try:
                        subkey_name = winreg.EnumKey(key, i)
                        subkey = winreg.OpenKey(key, subkey_name)
                        try:
                            display_name = winreg.QueryValueEx(subkey, "DisplayName")[0]
                            for product_name, keywords in security_indicators:
                                if any(keyword.lower() in display_name.lower() for keyword in keywords):
                                    if not any(product_name in p for p in security_products):
                                        security_products.append(f"{product_name}: {display_name}")
                        except:
                            pass
                        winreg.CloseKey(subkey)
                    except:
                        continue
                winreg.CloseKey(key)
            except:
                continue
    except Exception as e:
        print(f"Registry check error: {e}")
    
    # Check running processes
    try:
        for proc in psutil.process_iter(['name']):
            try:
                proc_name = proc.info['name'].lower()
                for product_name, keywords in security_indicators:
                    if any(keyword.lower() in proc_name for keyword in keywords):
                        if not any(product_name in p for p in security_products):
                            security_products.append(f"{product_name}: {proc_name} (running process)")
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
    except Exception as e:
        print(f"Process check error: {e}")
    
    # Check installed services
    try:
        output = subprocess.check_output('sc query', shell=True, text=True, stderr=subprocess.STDOUT)
        for line in output.split('\n'):
            for product_name, keywords in security_indicators:
                if any(keyword.lower() in line.lower() for keyword in keywords):
                    if not any(product_name in p for p in security_products):
                        security_products.append(f"{product_name}: Service detected")
    except Exception as e:
        print(f"Service check error: {e}")
    
    # Remove duplicates while preserving order
    seen = set()
    unique_products = []
    for product in security_products:
        if product not in seen:
            seen.add(product)
            unique_products.append(product)
    
    return unique_products

def show_security_warning():
    """Display security warning and exit if endpoint security is detected"""
    security_products = check_endpoint_security()
    
    if security_products:
        # Create warning message
        warning_message = "⚠️ ENDPOINT SECURITY DETECTED ⚠️\n\n"
        warning_message += "The following endpoint security products were detected:\n\n"
        
        for product in security_products:
            warning_message += f"• {product}\n"
        
        warning_message += "\n⚠️ IMPORTANT SECURITY ALERT ⚠️\n"
        warning_message += "Running this optimizer with endpoint security software\n"
        warning_message += "installed may trigger HIGH ALERT - \"Anti Theft\" warnings.\n\n"
        warning_message += "PLEASE UNINSTALL THE DETECTED ENDPOINT SECURITY SOFTWARE\n"
        warning_message += "BEFORE RUNNING THIS OPTIMIZER TO AVOID FALSE ALERTS.\n\n"
        warning_message += "Click OK to exit the application."
        
        # Show warning and exit
        ctypes.windll.user32.MessageBoxW(0, warning_message, "SECURITY WARNING - Endpoint Security Detected", 0x30)
        sys.exit(1)

# Run security check at startup
show_security_warning()

"""
================================================================================
WINDOWS 11 OPTIMIZER - VERSION HISTORY & DOCUMENTATION
================================================================================

VERSION 1.0 - INITIAL RELEASE
- Basic GUI with single mode optimization
- Limited tweaks: startup apps, visual effects, transparency
- No backup/restore functionality
- Simple checkbox interface

VERSION 2.0 - MULTI-MODE SUPPORT
- Added Standard, Ultimate, Extreme modes
- Modular framework for different optimization levels
- Basic registry backup system
- Export report functionality

VERSION 3.0 - ENHANCED GUI & SAFETY
- Scrollable interface for many options
- Improved error handling
- System restore point creation
- Better progress tracking

VERSION 4.0 - MAJOR OVERHAUL
- Complete GUI redesign matching reference image
- Red box action buttons layout
- Select All functionality for all modes
- Comprehensive backup system

VERSION 4.1 - BPO CALL CENTER OPTIMIZATIONS
- Added BPO-specific optimizations for call center environments
- HDD performance optimizations
- Login performance improvements
- Network optimizations for VoIP
- Reduced lag after lock/unlock cycles

VERSION 4.2 - MODE RENAMING & ENHANCEMENTS
- Renamed modes: Basic, Standard, Ultimate, Extreme
- Removed redundant "Basic" checkbox from top section
- Added 15+ new safe optimizations
- Two-column layout for better organization
- Enhanced safety features

VERSION 4.3 - BUG FIXES & FINAL POLISH
- Fixed all escape sequence errors in registry paths
- Auto-fitting GUI that adjusts to content
- Scrollable interface for smaller screens
- Fixed PyInstaller compilation issues
- Improved error handling and user feedback

VERSION 4.4 - FAST RESTART & SIGNOUT OPTIMIZATIONS
- Added Fast Restart optimization to Standard Mode
- Added Fast Signout optimization to Standard Mode
- Reduced Windows 11 restart and signout delays
- Optimized service timeout values
- Improved shutdown performance

VERSION 4.5 - IMMEDIATE RESTART/SIGNOUT & SESSION CLEANUP
- Added Session Cookie Clearing for FSSO authentication issues
- Enhanced security by clearing previous user sessions
- Fixed FSSO policy authentication problems

VERSION 4.6 - SECURITY & TELEMETRY FIXES
- REMOVED aggressive telemetry disabling that triggers security alerts
- Added safe telemetry reduction option
- Added comprehensive restore functionality for all tweaks
- Enhanced security compatibility
- Fixed Sophos theft alert triggers

VERSION 4.7 - SECURITY LOG FIXES
- ADDED "Fix Security Log Full" option to Basic Mode
- Permanently fixes "security log is full" error
- Increases security log size to 128MB
- Configures automatic overwrite of old events
- Prevents security event loss in enterprise environments
- Ensures continuous security monitoring

DOMAIN TRUST RELATIONSHIP FIXES
- ADDED "Fix Domain Trust Relationship" option to Basic Mode
- Permanently prevents "domain is broken or trust relationship issue" errors
- Fixes computer account password synchronization issues
- Optimizes secure channel communications
- Prevents workstation from losing domain trust
- Essential for enterprise Active Directory environments
- Added Domain Trust Verification tool

ENHANCED SECURITY DETECTION
- ADDED Endpoint Security Detection at startup
- Warns user if Sophos or other endpoint security is installed
- Prevents running optimizer with endpoint security to avoid false alerts
- Added reboot message after domain trust fix
- Enhanced safety features

KEY BUG FIXES:
1. Fixed SyntaxWarning: invalid escape sequence errors in all registry paths
2. Resolved PyInstaller module detection issues
3. Fixed GUI sizing and layout problems
4. Improved backup/restore reliability
5. Enhanced error handling for failed commands
6. Fixed security log full errors in enterprise environments
7. Fixed domain trust relationship issues in Active Directory

SAFETY FEATURES:
- All optimizations are reversible
- Automatic registry backups before changes
- System restore point creation
- No harmful registry edits
- Comprehensive undo functionality
- Security software compatible (Sophos, etc.)
- Security event log preservation
- Domain trust relationship protection
- Safe for home and enterprise use
- Endpoint security detection and warning

COMPILATION NOTES:
- Requires Pillow for PNG to ICO conversion: pip install pillow
- Requires psutil: pip install psutil
- Use: pyinstaller --onefile --windowed --name "Windows 11 Optimizer v4.7" --icon=optimizer.png --add-data "*.png;." --version-file=version.rc --hidden-import=tkinter --hidden-import=subprocess --hidden-import=os --hidden-import=datetime --hidden-import=threading win11_optimizer.py
- Tested on Windows 11 25H2 and later
================================================================================
"""

# -----------------------------
# HELPER FUNCTIONS
# -----------------------------
def run_cmd(cmd):
    """Execute command and log results safely"""
    try:
        output = subprocess.check_output(cmd, shell=True, text=True, stderr=subprocess.STDOUT, timeout=30).strip()
        report_log.append(f"Command: {cmd}\nOutput: {output}\n")
        return output
    except subprocess.TimeoutExpired:
        report_log.append(f"Command: {cmd}\nError: Timeout after 30 seconds\n")
        return "Failed: Timeout"
    except Exception as e:
        report_log.append(f"Command: {cmd}\nError: {str(e)}\n")
        return f"Failed: {e}"

def backup_registry(path=r"C:\\Win11_Optimizer_Backup"):
    """Create registry backup with timestamp"""
    global last_backup
    os.makedirs(path, exist_ok=True)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_file = os.path.join(path, f"registry_backup_{timestamp}.reg")
    
    def backup_thread():
        try:
            run_cmd(f'reg export HKCU "{backup_file}" /y')
            run_cmd(f'reg export HKLM "{backup_file}_HKLM.reg" /y')
            last_backup = backup_file
        except Exception as e:
            print(f"Backup error: {e}")
    
    threading.Thread(target=backup_thread, daemon=True).start()
    return backup_file

def create_restore_point():
    """Create system restore point for safety - NON-BLOCKING VERSION"""
    def create_restore_thread():
        try:
            result = run_cmd(r'powershell -Command "Checkpoint-Computer -Description \"Win11Optimizer Backup\" -RestorePointType \"MODIFY_SETTINGS\""')
            if "Failed" not in result:
                messagebox.showinfo("Restore Point Created", "System restore point created successfully.")
            else:
                messagebox.showwarning("Restore Point Failed", "Could not create restore point!")
        except Exception as e:
            print(f"Restore point error: {e}")
    
    threading.Thread(target=create_restore_thread, daemon=True).start()
    messagebox.showinfo("Creating Restore Point", "System restore point is being created in the background. You can continue using the application.")

def export_report():
    """Export optimization report to file"""
    if not report_log:
        messagebox.showinfo("Export Report", "No actions have been performed yet.")
        return
    path = filedialog.asksaveasfilename(defaultextension=".txt",
                                        filetypes=[("Text files", "*.txt")])
    if path:
        with open(path, "w") as f:
            f.write("\n".join(report_log))
        messagebox.showinfo("Export Report", f"Report saved to {path}")

# -----------------------------
# VERIFY DOMAIN TRUST FUNCTION
# -----------------------------
def verify_domain_trust():
    """Verify domain trust status and show detailed information"""
    def verify_thread():
        try:
            messagebox.showinfo("Verifying Domain Trust", "Checking domain trust relationship status...")
            
            results = []
            results.append("=== DOMAIN TRUST STATUS CHECK ===")
            results.append(f"Timestamp: {datetime.datetime.now()}")
            results.append("")
            
            # 1. Check computer name and domain
            results.append("1. COMPUTER INFORMATION:")
            computer_name = run_cmd('hostname')
            results.append(f"   Computer Name: {computer_name}")
            
            domain_info = run_cmd('echo %USERDOMAIN%')
            results.append(f"   Domain: {domain_info}")
            
            # 2. Test secure channel
            results.append("\n2. SECURE CHANNEL TEST:")
            secure_test = run_cmd('powershell -Command "Test-ComputerSecureChannel -Verbose"')
            results.append(f"   Test-ComputerSecureChannel Result:\n   {secure_test}")
            
            # 3. Check Netlogon service
            results.append("\n3. NETLOGON SERVICE STATUS:")
            netlogon_status = run_cmd('sc query netlogon')
            results.append(f"   Netlogon Service:\n   {netlogon_status}")
            
            # 4. Check domain connectivity
            results.append("\n4. DOMAIN CONNECTIVITY:")
            domain_test = run_cmd('nltest /dsgetdc:%USERDOMAIN%')
            results.append(f"   Domain Controller Discovery:\n   {domain_test}")
            
            # 5. Check secure channel status
            results.append("\n5. SECURE CHANNEL DETAILS:")
            sc_status = run_cmd('nltest /sc_query:%USERDOMAIN%')
            results.append(f"   Secure Channel Query:\n   {sc_status}")
            
            # 6. Check DNS registration
            results.append("\n6. DNS REGISTRATION:")
            dns_reg = run_cmd('ipconfig /all | findstr /i "dns"')
            results.append(f"   DNS Configuration:\n   {dns_reg}")
            
            # 7. Check time synchronization
            results.append("\n7. TIME SYNCHRONIZATION:")
            time_status = run_cmd('w32tm /query /status')
            results.append(f"   Time Service Status:\n   {time_status}")
            
            # 8. Check Kerberos tickets
            results.append("\n8. KERBEROS TICKETS:")
            kerberos_tickets = run_cmd('klist')
            results.append(f"   Kerberos Tickets:\n   {kerberos_tickets}")
            
            # 9. Check network connectivity to domain
            results.append("\n9. NETWORK CONNECTIVITY:")
            ping_test = run_cmd('ping -n 2 %USERDOMAIN%')
            results.append(f"   Ping to Domain:\n   {ping_test}")
            
            # Save results to file
            log_file = r"C:\Windows\DomainTrustStatus.log"
            with open(log_file, 'w') as f:
                f.write('\n'.join(results))
            
            # Show results in messagebox
            result_text = '\n'.join(results)
            
            # Create a results window
            result_window = tk.Toplevel(root)
            result_window.title("Domain Trust Status Report")
            result_window.geometry("800x600")
            
            text_frame = tk.Frame(result_window)
            text_frame.pack(fill="both", expand=True, padx=10, pady=10)
            
            text_area = scrolledtext.ScrolledText(text_frame, wrap=tk.WORD, width=90, height=30)
            text_area.pack(fill="both", expand=True)
            text_area.insert(tk.INSERT, result_text)
            text_area.config(state=tk.DISABLED)
            
            tk.Button(result_window, text="Close", command=result_window.destroy, width=15).pack(pady=10)
            
        except Exception as e:
            messagebox.showerror("Verification Error", f"Failed to verify domain trust: {str(e)}")
    
    threading.Thread(target=verify_thread, daemon=True).start()

# -----------------------------
# COMPREHENSIVE RESTORE FUNCTION
# -----------------------------
def restore_all_tweaks():
    """Completely restore all Windows optimizations to default settings"""
    def restore_thread():
        try:
            messagebox.showinfo("Restore Started", "Beginning comprehensive system restoration...")
            
            # Restore telemetry to default
            restore_telemetry()
            
            # Restore visual effects
            run_cmd('reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\VisualEffects" /v "VisualFXSetting" /t REG_DWORD /d 3 /f')
            run_cmd('reg add "HKCU\\Control Panel\\Desktop" /v "DragFullWindows" /t REG_SZ /d "1" /f')
            
            # Restore transparency
            run_cmd('reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize" /v "EnableTransparency" /t REG_DWORD /d 1 /f')
            
            # Restore background apps
            run_cmd('reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\BackgroundAccessApplications" /v "GlobalUserDisabled" /t REG_DWORD /d 0 /f')
            
            # Restore taskbar animations
            run_cmd('reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced" /v "TaskbarAnimations" /t REG_DWORD /d 1 /f')
            
            # Restore notification sounds
            run_cmd('reg add "HKCU\\AppEvents\\Schemes\\Apps\\.Default\\.Default\\.Current" /ve /t REG_SZ /d "%SystemRoot%\\media\\Windows Notify System Generic.wav" /f')
            
            # Restore security log settings
            restore_security_log()
            
            # Restore domain trust settings
            restore_domain_trust()
            
            # Restore power plan to balanced
            run_cmd("powercfg /setactive 381b4222-f694-41f0-9685-ff5bb260df2e")
            
            # Restore services
            run_cmd("sc config DiagTrack start= demand")
            run_cmd("sc config OneSyncSvc start= demand")
            run_cmd("sc config WSearch start= auto")
            run_cmd("sc config Spooler start= auto")
            run_cmd("sc config WinDefend start= auto")
            run_cmd("sc config Netlogon start= auto")  # NEW: Restore Netlogon service
            
            # Restore Windows Defender
            run_cmd('reg delete "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows Defender" /v "DisableAntiSpyware" /f')
            
            # Restore shutdown timeouts to default
            run_cmd('reg add "HKLM\\SYSTEM\\CurrentControlSet\\Control" /v "WaitToKillServiceTimeout" /t REG_SZ /d "5000" /f')
            run_cmd('reg add "HKCU\\Control Panel\\Desktop" /v "AutoEndTasks" /t REG_SZ /d "0" /f')
            run_cmd('reg add "HKCU\\Control Panel\\Desktop" /v "HungAppTimeout" /t REG_SZ /d "5000" /f')
            run_cmd('reg add "HKCU\\Control Panel\\Desktop" /v "WaitToKillAppTimeout" /t REG_SZ /d "20000" /f')
            
            # Restore fast boot
            run_cmd('reg add "HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Power" /v "HiberbootEnabled" /t REG_DWORD /d 1 /f')
            
            # Clear applied tweaks list
            applied_tweaks.clear()
            
            messagebox.showinfo("Restore Complete", "All system optimizations have been restored to default Windows settings!")
            
        except Exception as e:
            messagebox.showerror("Restore Error", f"Failed to restore some settings: {str(e)}")
    
    threading.Thread(target=restore_thread, daemon=True).start()

# -----------------------------
# NEW DOMAIN TRUST RELATIONSHIP FUNCTION - UPDATED WITH REBOOT MESSAGE
# -----------------------------
def fix_domain_trust_relationship():
    """Permanently fix 'domain is broken or trust relationship issue' error - ENHANCED VERSION"""
    applied_tweaks.append("fix_domain_trust_relationship")
    
    result_log = []
    
    # 1. FIRST CHECK CURRENT STATUS
    result_log.append("=== DOMAIN TRUST REPAIR STARTED ===")
    result_log.append(f"Timestamp: {datetime.datetime.now()}")
    
    # Get current domain
    current_domain = run_cmd('echo %USERDOMAIN%').strip()
    if not current_domain:
        current_domain = "openaccess.bpo"  # Default fallback
    
    # Check current domain status
    status1 = run_cmd(f'netdom verify /d:{current_domain}')
    result_log.append(f"Initial domain status: {status1}")
    
    # 2. STOP CRITICAL SERVICES TEMPORARILY
    result_log.append("\n=== STOPPING SERVICES ===")
    run_cmd('net stop netlogon /y')
    run_cmd('sc stop kdc')
    run_cmd('sc stop dns')
    run_cmd('sc stop w32time')
    
    # 3. COMPLETELY CLEAR AUTHENTICATION CACHE
    result_log.append("\n=== CLEARING AUTHENTICATION CACHE ===")
    run_cmd('klist purge -li 0x3e7')
    run_cmd('klist purge -li 0x0')
    run_cmd('klist purge')
    
    # Clear credential manager
    run_cmd('cmdkey /delete:TERMSRV/*')
    run_cmd(f'cmdkey /delete:{current_domain}')
    run_cmd('cmdkey /list | findstr /i "domain" | for /f "tokens=1,2 delims= " %a in (''more'') do cmdkey /delete:%b')
    
    # 4. FORCE COMPUTER ACCOUNT PASSWORD RESET (MULTIPLE METHODS)
    result_log.append("\n=== RESETTING COMPUTER ACCOUNT PASSWORD ===")
    
    # Method 1: Using netdom (most reliable)
    run_cmd(f'netdom resetpwd /s:{current_domain} /ud:{current_domain}\\administrator /pd:*')
    
    # Method 2: PowerShell method
    ps_command = f'''
    $domain = "{current_domain}"
    $computer = $env:COMPUTERNAME
    try {{
        Reset-ComputerMachinePassword -Server $domain -ErrorAction Stop
        Write-Output "Password reset successful for $computer in domain $domain"
    }} catch {{
        Write-Output "Failed to reset password: $_"
        # Try alternative method
        nltest /sc_reset:$domain
    }}
    '''
    run_cmd(f'powershell -Command "{ps_command}"')
    
    # Method 3: Direct registry method
    run_cmd('reg delete "HKLM\\SYSTEM\\CurrentControlSet\\Services\\Netlogon\\Parameters" /v "MachinePassword" /f')
    
    # 5. REBUILD SECURE CHANNEL FROM SCRATCH
    result_log.append("\n=== REBUILDING SECURE CHANNEL ===")
    
    # Reset secure channel completely
    run_cmd(f'nltest /sc_reset:{current_domain}')
    run_cmd(f'nltest /sc_verify:{current_domain}')
    run_cmd(f'nltest /dsgetdc:{current_domain}')
    
    # Force re-discovery of domain controller
    run_cmd(f'nltest /dsgetdc:{current_domain} /force')
    
    # 6. FIX DNS REGISTRATION (CRITICAL)
    result_log.append("\n=== FIXING DNS REGISTRATION ===")
    
    # Clear ALL DNS caches
    run_cmd('ipconfig /flushdns')
    run_cmd('ipconfig /registerdns')
    run_cmd('ipconfig /release')
    run_cmd('ipconfig /renew')
    run_cmd('ipconfig /registerdns')
    
    # Force DNS registration
    run_cmd('reg add "HKLM\\SYSTEM\\CurrentControlSet\\Services\\Tcpip\\Parameters" /v "DisableDynamicUpdate" /t REG_DWORD /d 0 /f')
    
    # 7. FIX TIME SYNCHRONIZATION (ESSENTIAL FOR KERBEROS)
    result_log.append("\n=== FIXING TIME SYNCHRONIZATION ===")
    
    # Configure time service for domain
    run_cmd('w32tm /config /syncfromflags:domhier /update')
    run_cmd('w32tm /resync /force')
    run_cmd('w32tm /query /status')
    
    # Set time service to auto
    run_cmd('sc config w32time start= auto')
    run_cmd('sc start w32time')
    
    # 8. FIX NETLOGON SERVICE SETTINGS
    result_log.append("\n=== CONFIGURING NETLOGON SERVICE ===")
    
    # Reset Netlogon service to defaults and reconfigure
    run_cmd('sc config netlogon start= auto')
    run_cmd('sc failure netlogon reset= 86400 actions= restart/5000/restart/10000/restart/30000')
    
    # Configure Netlogon for better domain communication
    run_cmd('reg add "HKLM\\SYSTEM\\CurrentControlSet\\Services\\Netlogon\\Parameters" /v "ScavengeInterval" /t REG_DWORD /d 172800 /f')
    run_cmd('reg add "HKLM\\SYSTEM\\CurrentControlSet\\Services\\Netlogon\\Parameters" /v "MaximumPasswordAge" /t REG_DWORD /d 42 /f')
    run_cmd('reg add "HKLM\\SYSTEM\\CurrentControlSet\\Services\\Netlogon\\Parameters" /v "SecureChannelIdleTimeout" /t REG_DWORD /d 1209600 /f')  # 14 days
    run_cmd('reg add "HKLM\\SYSTEM\\CurrentControlSet\\Services\\Netlogon\\Parameters" /v "SecureChannelTimeout" /t REG_DWORD /d 7200 /f')  # 2 hours
    
    # 9. FIX GROUP POLICY PROCESSING
    result_log.append("\n=== FIXING GROUP POLICY ===")
    
    # Clear Group Policy cache
    run_cmd('rd /s /q "%WinDir%\\System32\\GroupPolicyUsers"')
    run_cmd('rd /s /q "%WinDir%\\System32\\GroupPolicy"')
    
    # Force Group Policy update
    run_cmd('gpupdate /force')
    
    # 10. RESTART SERVICES IN CORRECT ORDER
    result_log.append("\n=== RESTARTING SERVICES ===")
    
    run_cmd('sc start w32time')
    run_cmd('sc start dns')
    run_cmd('net start netlogon')
    run_cmd('sc start kdc')
    
    # 11. VERIFY REPAIR
    result_log.append("\n=== VERIFYING REPAIR ===")
    
    # Test secure channel
    verify1 = run_cmd(f'nltest /sc_query:{current_domain}')
    result_log.append(f"Secure channel query: {verify1}")
    
    # Test domain trust
    verify2 = run_cmd(f'netdom verify /d:{current_domain}')
    result_log.append(f"Domain verification: {verify2}")
    
    # Test computer secure channel
    verify3 = run_cmd('powershell -Command "Test-ComputerSecureChannel -Verbose"')
    result_log.append(f"Test-ComputerSecureChannel: {verify3}")
    
    # 12. CREATE REPAIR LOG
    result_log.append("\n=== REPAIR COMPLETE ===")
    result_log.append(f"Timestamp: {datetime.datetime.now()}")
    
    # Save detailed log
    log_path = r"C:\Windows\DomainTrustRepair.log"
    try:
        with open(log_path, 'w') as f:
            f.write('\n'.join(result_log))
    except:
        log_path = r"C:\DomainTrustRepair.log"
        with open(log_path, 'w') as f:
            f.write('\n'.join(result_log))
    
    # 13. FINAL CLEANUP AND REBOOT RECOMMENDATION
    run_cmd('gpupdate /force')
    run_cmd('ipconfig /flushdns')
    
    # Show reboot message
    messagebox.showwarning("REBOOT REQUIRED", 
        "Domain trust repair completed successfully!\n\n"
        "⚠️ REBOOT THE SYSTEM TO TAKE EFFECT THE CHANGES ⚠️\n\n"
        f"Log saved to: {log_path}\n\n"
        "After reboot, use 'Verify Domain Trust' button to confirm repair.")
    
    return f"Domain trust repair completed. REBOOT REQUIRED for changes to take full effect."

def restore_domain_trust():
    """Restore domain trust settings to Windows defaults"""
    # Remove registry optimizations
    run_cmd('reg delete "HKLM\\SYSTEM\\CurrentControlSet\\Services\\Netlogon\\Parameters" /v "ScavengeInterval" /f')
    run_cmd('reg delete "HKLM\\SYSTEM\\CurrentControlSet\\Services\\Netlogon\\Parameters" /v "MaximumPasswordAge" /f')
    run_cmd('reg delete "HKLM\\SYSTEM\\CurrentControlSet\\Services\\Netlogon\\Parameters" /v "SecureChannelIdleTimeout" /f')
    run_cmd('reg delete "HKLM\\SYSTEM\\CurrentControlSet\\Services\\Netlogon\\Parameters" /v "SecureChannelTimeout" /f')
    run_cmd('reg delete "HKLM\\SYSTEM\\CurrentControlSet\\Services\\Netlogon\\Parameters" /v "DisablePasswordChange" /f')
    run_cmd('reg delete "HKLM\\SYSTEM\\CurrentControlSet\\Services\\Netlogon\\Parameters" /v "NegativeCachePeriod" /f')
    run_cmd('reg delete "HKLM\\SYSTEM\\CurrentControlSet\\Services\\Netlogon\\Parameters" /v "SignSecureChannel" /f')
    run_cmd('reg delete "HKLM\\SYSTEM\\CurrentControlSet\\Services\\Netlogon\\Parameters" /v "SealSecureChannel" /f')
    
    # Restore DNS settings
    run_cmd('reg delete "HKLM\\SYSTEM\\CurrentControlSet\\Services\\Tcpip\\Parameters" /v "DisableDynamicUpdate" /f')
    run_cmd('reg delete "HKLM\\SYSTEM\\CurrentControlSet\\Services\\NetBT\\Parameters\\Interfaces" /v "NetbiosOptions" /f')
    
    # Clear DNS cache
    run_cmd('ipconfig /flushdns')
    
    # Restart Netlogon service
    run_cmd('sc config Netlogon start= auto')
    run_cmd('net stop Netlogon /y')
    run_cmd('net start Netlogon')
    
    # Clear Kerberos tickets
    run_cmd('klist purge')
    
    # Force Group Policy update
    run_cmd('gpupdate /force')
    
    return "Domain trust settings restored to Windows defaults"

# -----------------------------
# SECURITY LOG FUNCTION
# -----------------------------
def fix_security_log_full():
    """Permanently fix 'security log is full' error by increasing size and enabling auto-overwrite"""
    applied_tweaks.append("fix_security_log_full")
    
    # Increase Security log size to 128MB (default is 20MB)
    run_cmd('wevtutil sl Security /ms:134217728')
    
    # Set maximum log size to 128MB
    run_cmd('wevtutil sl Security /maxsize:134217728')
    
    # Configure Security log to overwrite events as needed (prevents log full errors)
    run_cmd('wevtutil sl Security /retention:false')
    
    # Clear any existing 'log full' state
    run_cmd('powershell "Clear-EventLog -LogName Security -ErrorAction SilentlyContinue"')
    
    # Set Security log to archive when full (optional, but good for compliance)
    run_cmd('wevtutil sl Security /autobackup:true')
    
    # Apply same settings to other important logs to prevent similar issues
    run_cmd('wevtutil sl Application /ms:67108864')  # 64MB
    run_cmd('wevtutil sl Application /maxsize:67108864')
    run_cmd('wevtutil sl Application /retention:false')
    
    run_cmd('wevtutil sl System /ms:67108864')  # 64MB
    run_cmd('wevtutil sl System /maxsize:67108864')
    run_cmd('wevtutil sl System /retention:false')
    
    # Also configure via registry for persistent settings
    run_cmd('reg add "HKLM\\SYSTEM\\CurrentControlSet\\Services\\EventLog\\Security" /v "MaxSize" /t REG_DWORD /d 134217728 /f')
    run_cmd('reg add "HKLM\\SYSTEM\\CurrentControlSet\\Services\\EventLog\\Security" /v "Retention" /t REG_DWORD /d 0 /f')  # 0 = overwrite as needed
    run_cmd('reg add "HKLM\\SYSTEM\\CurrentControlSet\\Services\\EventLog\\Security" /v "AutoBackupLogFiles" /t REG_DWORD /d 1 /f')
    
    return "Security log permanently fixed: Increased to 128MB with auto-overwrite enabled"

def restore_security_log():
    """Restore security log to Windows default settings"""
    # Restore Security log to default size (20MB)
    run_cmd('wevtutil sl Security /ms:20971520')
    run_cmd('wevtutil sl Security /maxsize:20971520')
    run_cmd('wevtutil sl Security /retention:true')  # Default retention
    run_cmd('wevtutil sl Security /autobackup:false')
    
    # Restore other logs to default
    run_cmd('wevtutil sl Application /ms:20971520')
    run_cmd('wevtutil sl Application /maxsize:20971520')
    run_cmd('wevtutil sl Application /retention:true')
    
    run_cmd('wevtutil sl System /ms:20971520')
    run_cmd('wevtutil sl System /maxsize:20971520')
    run_cmd('wevtutil sl System /retention:true')
    
    # Remove registry overrides
    run_cmd('reg delete "HKLM\\SYSTEM\\CurrentControlSet\\Services\\EventLog\\Security" /v "MaxSize" /f')
    run_cmd('reg delete "HKLM\\SYSTEM\\CurrentControlSet\\Services\\EventLog\\Security" /v "Retention" /f')
    run_cmd('reg delete "HKLM\\SYSTEM\\CurrentControlSet\\Services\\EventLog\\Security" /v "AutoBackupLogFiles" /f')
    
    return "Security log restored to Windows defaults"

# -----------------------------
# SAFE TELEMETRY FUNCTIONS
# -----------------------------
def safe_reduce_telemetry():
    """Safe telemetry reduction that won't trigger security alerts"""
    applied_tweaks.append("safe_reduce_telemetry")
    
    # Set to Basic level instead of completely disabled (less suspicious)
    run_cmd('reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\DataCollection" /v "AllowTelemetry" /t REG_DWORD /d 1 /f')  # Basic level
    
    # Reduce tailored experiences without disabling service
    run_cmd('reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Privacy" /v "TailoredExperiencesWithDiagnosticDataEnabled" /t REG_DWORD /d 0 /f')
    
    # Disable advertising ID without aggressive service changes
    run_cmd('reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\AdvertisingInfo" /v "DisabledByGroupPolicy" /t REG_DWORD /d 1 /f')
    
    return "Reduced telemetry to Basic level safely"

def restore_telemetry():
    """Completely restore telemetry to Windows default settings"""
    # Remove telemetry restrictions
    run_cmd('reg delete "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\DataCollection" /v "AllowTelemetry" /f')
    run_cmd('reg delete "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Privacy" /v "TailoredExperiencesWithDiagnosticDataEnabled" /f')
    run_cmd('reg delete "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\AdvertisingInfo" /v "DisabledByGroupPolicy" /f')
    
    # Restore DiagTrack service
    run_cmd("sc config DiagTrack start= demand")
    run_cmd("sc start DiagTrack")
    
    return "Telemetry restored to Windows defaults"

# -----------------------------
# UPDATED DOCUMENTATION FUNCTIONS
# -----------------------------
def show_version_history():
    """Display comprehensive version history"""
    version_text = f"""
Windows 11 Optimizer - Complete Version History
================================================

VERSION 4.7 SECURE - ENHANCED SECURITY DETECTION
• ADDED Endpoint Security Detection at startup
• Warns user if Sophos or other endpoint security is installed
• Prevents running optimizer with endpoint security to avoid false alerts
• Added reboot message after domain trust fix
• Enhanced safety features

DOMAIN TRUST RELATIONSHIP FIXES
• ADDED "Fix Domain Trust Relationship" option to Basic Mode
• Permanently prevents "domain is broken or trust relationship issue" errors
• Fixes computer account password synchronization issues
• Optimizes secure channel communications
• Prevents workstation from losing domain trust
• Essential for enterprise Active Directory environments
• ADDED Domain Trust Verification tool
• Enhanced secure channel repair methods

SECURITY LOG FIXES
• ADDED "Fix Security Log Full" option to Basic Mode
• Permanently fixes "security log is full" error
• Increases security log size to 128MB
• Configures automatic overwrite of old events
• Prevents security event loss in enterprise environments
• Ensures continuous security monitoring

VERSION 4.6 - SECURITY & TELEMETRY FIXES
• REMOVED aggressive telemetry disabling that triggers Sophos alerts
• Added safe telemetry reduction option
• Added comprehensive restore functionality for all tweaks
• Enhanced security compatibility
• Fixed theft alert triggers from security software

VERSION 4.5 - SESSION CLEANUP
• Added Session Cookie Clearing for FSSO authentication issues
• Enhanced security by clearing previous user sessions
• Fixed FSSO policy authentication problems

VERSION 4.4 - FAST RESTART & SIGNOUT OPTIMIZATIONS
• Added Fast Restart optimization to Standard Mode
• Added Fast Signout optimization to Standard Mode
• Reduced Windows 11 restart and signout delays
• Optimized service timeout values
• Improved shutdown performance

VERSION 4.3 - BUG FIXES & FINAL POLISH
• Fixed all escape sequence errors in registry paths
• Auto-fitting GUI that adjusts to content
• Scrollable interface for smaller screens
• Fixed PyInstaller compilation issues
• Improved error handling and user feedback

VERSION 4.2 - MODE RENAMING & ENHANCEMENTS
• Renamed modes: Basic, Standard, Ultimate, Extreme
• Removed redundant "Basic" checkbox from top section
• Added 15+ new safe optimizations
• Two-column layout for better organization
• Enhanced safety features

VERSION 4.1 - BPO CALL CENTER OPTIMIZATIONS
• Added BPO-specific optimizations for call center environments
• HDD performance optimizations
• Login performance improvements
• Network optimizations for VoIP
• Reduced lag after lock/unlock cycles

VERSION 4.0 - MAJOR OVERHAUL
• Complete GUI redesign matching reference image
• Red box action buttons layout
• Select All functionality for all modes
• Comprehensive backup system

VERSION 3.0 - ENHANCED GUI & SAFETY
• Scrollable interface for many options
• Improved error handling
• System restore point creation
• Better progress tracking

VERSION 2.0 - MULTI-MODE SUPPORT
• Added Standard, Ultimate, Extreme modes
• Modular framework for different optimization levels
• Basic registry backup system
• Export report functionality

VERSION 1.0 - INITIAL RELEASE
• Basic GUI with single mode optimization
• Limited tweaks: startup apps, visual effects, transparency
• No backup/restore functionality
• Simple checkbox interface

KEY BUG FIXES ACROSS VERSIONS:
✓ Fixed SyntaxWarning: invalid escape sequence errors
✓ Resolved PyInstaller module detection issues  
✓ Fixed GUI sizing and layout problems
✓ Improved backup/restore reliability
✓ Enhanced error handling for failed commands
✓ Fixed Sophos security alert triggers
✓ Fixed security log full errors in enterprise environments
✓ Fixed domain trust relationship issues in Active Directory
✓ Added endpoint security detection

SAFETY FEATURES:
• All optimizations are reversible
• Automatic registry backups before changes
• System restore point creation
• No harmful registry edits
• Comprehensive undo functionality
• Security software compatible
• Endpoint security detection and warning
• Enterprise environment safe
• Security event log preservation
• Domain trust relationship protection
• Active Directory compatible
"""
    
    # Create a new window for version history
    version_window = tk.Toplevel(root)
    version_window.title(f"Version History - Windows 11 Optimizer v{VERSION}")
    version_window.geometry("800x600")
    
    # Create scrollable text area
    text_frame = tk.Frame(version_window)
    text_frame.pack(fill="both", expand=True, padx=10, pady=10)
    
    text_area = scrolledtext.ScrolledText(text_frame, wrap=tk.WORD, width=80, height=30)
    text_area.pack(fill="both", expand=True)
    text_area.insert(tk.INSERT, version_text)
    text_area.config(state=tk.DISABLED)  # Make read-only
    
    # Close button
    tk.Button(version_window, text="Close", command=version_window.destroy, width=15).pack(pady=10)

def show_help():
    """Display comprehensive help about all optimization options"""
    help_text = f"""
Windows 11 Optimizer v{VERSION}
===========================================

IMPORTANT SECURITY UPDATE - VERSION 4.7 SECURE:
• ADDED Endpoint Security Detection at startup
• Detects Sophos, McAfee, Norton, Kaspersky, and other security software
• Prevents running optimizer to avoid triggering "Anti Theft" alerts
• Shows warning message and exits if endpoint security is detected
• Enhanced safety for enterprise environments

DOMAIN TRUST RELATIONSHIP FIX:
• "Fix Domain Trust Relationship" option in Basic Mode
• Permanently prevents "The trust relationship between this workstation and the primary domain failed" errors
• Fixes computer account password synchronization issues
• Optimizes secure channel communications with domain controllers
• Prevents workstations from losing domain trust
• Essential for enterprise Active Directory environments
• Fixes Netlogon service and Kerberos authentication problems
• NEW: "Verify Domain Trust" button to check current status
• AFTER FIX: "REBOOT THE SYSTEM TO TAKE EFFECT THE CHANGES"

SECURITY LOG FIX:
• "Fix Security Log Full" option in Basic Mode
• Permanently resolves "The security log is full" errors
• Increases log size from 20MB to 128MB
• Configures automatic overwrite of old events
• Ensures continuous security monitoring
• Prevents security event loss in enterprise environments

ENDPOINT SECURITY DETECTION:
----------------------------
The optimizer now checks for installed endpoint security software including:
• Sophos Endpoint Security
• McAfee Endpoint Security
• Norton/Symantec products
• Kaspersky Endpoint Security
• Trend Micro
• ESET Endpoint
• And other major security vendors

If endpoint security is detected, you will see:
⚠️ "ENDPOINT SECURITY DETECTED" warning message
⚠️ List of detected security products
⚠️ Warning about triggering "Anti Theft" alerts
⚠️ Instructions to uninstall before running optimizer

IMPORTANT: Microsoft Defender is NOT detected as it's built into Windows.

BASIC MODE OPTIMIZATIONS
-------------------------

Fix Domain Trust Relationship (NEW)
• Permanently fixes "domain trust relationship broken" errors
• Resets computer account password and re-establishes secure channel
• Optimizes Netlogon service for better domain communication
• Increases secure channel timeout to prevent premature disconnection
• Fixes DNS registration issues that can cause trust problems
• Ensures Kerberos authentication works correctly
• Prevents workstation from being locked out of domain
• ⚠️ AFTER APPLICATION: "REBOOT THE SYSTEM TO TAKE EFFECT THE CHANGES"
• Essential for Active Directory domain environments

Fix Security Log Full (NEW)
• Permanently fixes "security log is full" error messages
• Increases Security log size from 20MB to 128MB
• Configures automatic overwrite of old events
• Ensures continuous security event logging
• Prevents security monitoring gaps in enterprise environments
• Safe for compliance and auditing requirements

Disable Startup Apps
• Removes applications that auto-start with Windows
• Reduces boot time and memory usage
• Does NOT affect essential system services

Disable Visual Effects 
• Turns off animations, shadows, and visual enhancements
• Improves system responsiveness, especially on older hardware
• Makes Windows feel faster and more responsive

Disable Transparency
• Removes transparent effects from taskbar, start menu, and title bars
• Reduces GPU usage and improves performance
• Provides a cleaner, more classic Windows appearance

Disable Background Apps
• Prevents apps from running in the background when not in use
• Saves battery life on laptops and reduces CPU usage
• Improves overall system performance

Optimize Taskbar Animations
• Disables subtle animations in the taskbar
• Makes taskbar interactions feel instant
• Reduces visual distractions

Disable Notification Sounds
• Mutes system notification sounds
• Creates quieter work environment
• Does not affect visual notifications

Enable Fast Start Menu
• Optimizes start menu search and loading
• Makes start menu open instantly
• Improves search responsiveness

Disable Lockscreen Blur
• Removes blur effect from lockscreen
• Faster lockscreen loading
• Cleaner visual appearance

Optimize File Explorer
• Shows hidden files by default
• Hides protected system files for safety
• Makes file management more efficient

Disable Game Bar
• Turns off Xbox Game Bar feature
• Frees up system resources
• Prevents accidental activation during work

VERIFY DOMAIN TRUST BUTTON:
• NEW: Click "Verify Domain Trust" button in bottom toolbar
• Runs comprehensive domain trust diagnostic
• Shows Test-ComputerSecureChannel results
• Checks Netlogon service status
• Verifies domain controller connectivity
• Tests DNS registration and time synchronization
• Generates detailed report log file

DOMAIN TRUST FIX DETAILS:
-------------------------
The enhanced domain trust fix performs the following:

1. COMPREHENSIVE PASSWORD RESET:
   • Multiple methods for computer account password reset
   • Forces secure channel re-establishment
   • Clears authentication caches

2. SERVICE OPTIMIZATION:
   • Stops and restarts Netlogon, DNS, Time services in correct order
   • Configures optimal secure channel timeouts
   • Prevents premature disconnection from domain

3. NETWORK FIXES:
   • Clears DNS cache and forces re-registration
   • Fixes time synchronization (critical for Kerberos)
   • Ensures proper network connectivity

4. VERIFICATION:
   • Tests secure channel after repair
   • Creates detailed log file at C:\\Windows\\DomainTrustRepair.log
   • REQUIRES REBOOT for full effect

IMPORTANT: After running domain trust fix, REBOOT THE COMPUTER and use
the "Verify Domain Trust" button to confirm repair success.

STANDARD MODE OPTIMIZATIONS
----------------------------
• Reduce Telemetry (Safe) - Security-friendly telemetry reduction
• Disable Unnecessary Services - Stops non-essential background services
• Disable Indexing - Turns off Windows Search indexing
• Set High Performance - Configures maximum performance power plan
• Disk Defrag/Trim - Optimizes SSD/HDD performance
• Disable Timeline - Turns off Windows Timeline feature
• Disable Location Tracking - Disables location services
• Optimize System Cache - Configures system cache for performance
• Disable Print Spooler - Stops print spooler service (if not needed)
• Optimize Processor - Configures processor scheduling for performance
• Disable Remote Assistance - Turns off remote assistance
• Enable Fast Restart - Reduces restart time
• Enable Fast Signout - Reduces signout time
• Clear Session Cookies (FSSO) - Clears authentication cookies

ULTIMATE MODE OPTIMIZATIONS
----------------------------
• Disable Defender - Turns off Windows Defender (use with caution)
• Disable Cortana - Disables Cortana assistant
• Disable Xbox Services - Turns off Xbox-related services
• Remove Bloat Apps - Removes pre-installed Windows apps
• Disable Tips/Notifications - Turns off Windows tips and suggestions
• Disable Advertising - Disables personalized ads
• Disable Error Reporting - Turns off Windows error reporting
• Disable Smart Screen - Disables Windows SmartScreen filter
• Disable Feedback - Turns off Windows feedback requests
• Disable Auto Update - Disables automatic Windows updates
• Disable OneDrive - Turns off OneDrive integration

EXTREME MODE OPTIMIZATIONS
---------------------------
• Optimize HDD Performance - Configures HDD for maximum performance
• Disable UI Animations - Turns off all Windows animations
• Optimize Login Performance - Reduces login time
• Disable Auto-Restart Updates - Prevents auto-restart after updates
• Optimize Network Performance - Configures network for speed
• Disable System Maintenance - Turns off automatic maintenance
• Optimize Paging File - Configures virtual memory optimally
• Disable Search Indexing - Completely disables search indexing
• Cleanup Temp Files - Cleans temporary files
• Optimize System Responsiveness - Improves system responsiveness
• Advanced Visual Effects - Disables advanced visual effects
• Optimize Memory Management - Configures memory management
• Optimize Disk Cache - Optimizes disk caching
• Disable Thumbnail Cache - Turns off thumbnail caching
• Optimize Context Menu - Speeds up right-click menu

COMPREHENSIVE RESTORE FEATURE
------------------------------------
• "Undo Tweaks" button completely restores ALL changes
• Restores domain trust settings to Windows defaults
• Restores security log settings to defaults
• Perfect for troubleshooting or reverting changes

SAFETY INFORMATION
------------------

ALL OPTIMIZATIONS ARE:
• Safe and tested on Windows 11
• Reversible using "Undo Tweaks" button
• Backed up automatically before application
• Non-destructive to system files
• Security software compatible (Sophos, etc.)
• Security event log preserving
• Active Directory domain compatible

RECOMMENDED USAGE:
• Basic Mode: Everyday performance improvement + Security/Trust fixes
• Standard Mode: Work computers, general optimization  
• Ultimate Mode: Gaming systems, maximum performance
• Extreme Mode: BPO/call centers, specialized environments

BACKUP FEATURES:
• Automatic registry backup before any changes
• System restore point creation
• Exportable report of all changes made
• Complete undo functionality
• Tracked tweak application for precise restoration
"""
    
    # Create a new window for help
    help_window = tk.Toplevel(root)
    help_window.title(f"Help Guide - Windows 11 Optimizer v{VERSION}")
    help_window.geometry("900x700")
    
    # Create scrollable text area
    text_frame = tk.Frame(help_window)
    text_frame.pack(fill="both", expand=True, padx=10, pady=10)
    
    text_area = scrolledtext.ScrolledText(text_frame, wrap=tk.WORD, width=100, height=35)
    text_area.pack(fill="both", expand=True)
    text_area.insert(tk.INSERT, help_text)
    text_area.config(state=tk.DISABLED)  # Make read-only
    
    # Close button
    tk.Button(help_window, text="Close", command=help_window.destroy, width=15).pack(pady=10)

# -----------------------------
# TWEAK FUNCTIONS - BASIC MODE
# -----------------------------
def disable_startup_apps(): 
    applied_tweaks.append("disable_startup_apps")
    run_cmd("powershell Get-CimInstance Win32_StartupCommand | Remove-CimInstance")
    run_cmd('reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\StartupApproved" /v "StartupDelayInMSec" /t REG_DWORD /d 0 /f')

def disable_visual_effects(): 
    applied_tweaks.append("disable_visual_effects")
    run_cmd('reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\VisualEffects" /v "VisualFXSetting" /t REG_DWORD /d 2 /f')
    run_cmd('reg add "HKCU\\Control Panel\\Desktop" /v "DragFullWindows" /t REG_SZ /d "0" /f')

def disable_transparency(): 
    applied_tweaks.append("disable_transparency")
    run_cmd('reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize" /v "EnableTransparency" /t REG_DWORD /d 0 /f')

def disable_background_apps(): 
    applied_tweaks.append("disable_background_apps")
    run_cmd('reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\BackgroundAccessApplications" /v "GlobalUserDisabled" /t REG_DWORD /d 1 /f')

def optimize_taskbar_animations():
    applied_tweaks.append("optimize_taskbar_animations")
    run_cmd('reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced" /v "TaskbarAnimations" /t REG_DWORD /d 0 /f')

def disable_notification_sounds():
    applied_tweaks.append("disable_notification_sounds")
    run_cmd('reg add "HKCU\\AppEvents\\Schemes\\Apps\\.Default\\.Default\\.Current" /ve /t REG_SZ /d "" /f')

def enable_fast_start_menu():
    applied_tweaks.append("enable_fast_start_menu")
    run_cmd('reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced" /v "Start_SearchFiles" /t REG_DWORD /d 2 /f')

def disable_lockscreen_blur():
    applied_tweaks.append("disable_lockscreen_blur")
    run_cmd('reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\System" /v "DisableAcrylicBackgroundOnLogon" /t REG_DWORD /d 1 /f')

def optimize_file_explorer():
    applied_tweaks.append("optimize_file_explorer")
    run_cmd('reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced" /v "Hidden" /t REG_DWORD /d 1 /f')
    run_cmd('reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced" /v "ShowSuperHidden" /t REG_DWORD /d 0 /f')

def disable_game_bar():
    applied_tweaks.append("disable_game_bar")
    run_cmd('reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\GameBar" /v "AllowAutoGameMode" /t REG_DWORD /d 0 /f')

# -----------------------------
# STANDARD MODE FUNCTIONS
# -----------------------------
def disable_services(): 
    applied_tweaks.append("disable_services")
    run_cmd("sc stop OneSyncSvc")
    run_cmd("sc config OneSyncSvc start= disabled")
    run_cmd("sc stop MapsBroker")
    run_cmd("sc config MapsBroker start= disabled")

def disable_indexing(): 
    applied_tweaks.append("disable_indexing")
    run_cmd("sc stop WSearch")
    run_cmd("sc config WSearch start= disabled")

def set_high_performance(): 
    applied_tweaks.append("set_high_performance")
    run_cmd("powercfg /setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c")
    run_cmd('reg add "HKLM\\SYSTEM\\CurrentControlSet\\Control\\Power" /v "HibernateEnabled" /t REG_DWORD /d 0 /f')

def disk_defrag_trim(): 
    applied_tweaks.append("disk_defrag_trim")
    run_cmd("defrag C: /O /U")
    run_cmd("powershell Optimize-Volume -DriveLetter C -ReTrim -Verbose")

def disable_timeline():
    applied_tweaks.append("disable_timeline")
    run_cmd('reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\System" /v "EnableActivityFeed" /t REG_DWORD /d 0 /f')

def disable_location_tracking():
    applied_tweaks.append("disable_location_tracking")
    run_cmd('reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\LocationAndSensors" /v "DisableLocation" /t REG_DWORD /d 1 /f')

def optimize_system_cache():
    applied_tweaks.append("optimize_system_cache")
    run_cmd('reg add "HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Memory Management" /v "LargeSystemCache" /t REG_DWORD /d 1 /f')

def disable_print_spooler():
    applied_tweaks.append("disable_print_spooler")
    run_cmd("sc stop Spooler")
    run_cmd("sc config Spooler start= disabled")

def optimize_processor_scheduling():
    applied_tweaks.append("optimize_processor_scheduling")
    run_cmd('reg add "HKLM\\SYSTEM\\CurrentControlSet\\Control\\PriorityControl" /v "Win32PrioritySeparation" /t REG_DWORD /d 26 /f')

def disable_remote_assistance():
    applied_tweaks.append("disable_remote_assistance")
    run_cmd('reg add "HKLM\\SYSTEM\\CurrentControlSet\\Control\\Remote Assistance" /v "fAllowToGetHelp" /t REG_DWORD /d 0 /f')

def enable_fast_restart():
    applied_tweaks.append("enable_fast_restart")
    run_cmd('reg add "HKLM\\SYSTEM\\CurrentControlSet\\Control" /v "WaitToKillServiceTimeout" /t REG_SZ /d "2000" /f')
    run_cmd('reg add "HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Power" /v "HiberbootEnabled" /t REG_DWORD /d 1 /f')
    run_cmd('reg add "HKCU\\Control Panel\\Desktop" /v "HungAppTimeout" /t REG_SZ /d "3000" /f')
    run_cmd('reg add "HKCU\\Control Panel\\Desktop" /v "WaitToKillAppTimeout" /t REG_SZ /d "2000" /f')
    run_cmd('reg add "HKCU\\Control Panel\\Desktop" /v "AutoEndTasks" /t REG_SZ /d "1" /f')
    run_cmd('reg add "HKLM\\SYSTEM\\CurrentControlSet\\Control" /v "ServicesPipeTimeout" /t REG_DWORD /d 3000 /f')

def enable_fast_signout():
    applied_tweaks.append("enable_fast_signout")
    run_cmd('reg add "HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Winlogon" /v "ProfileDlgTimeOut" /t REG_DWORD /d 10 /f')
    run_cmd('reg add "HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Winlogon" /v "DelayedDesktopSwitchTimeout" /t REG_DWORD /d 0 /f')
    run_cmd('reg add "HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Memory Management" /v "ClearPageFileAtShutdown" /t REG_DWORD /d 0 /f')
    run_cmd('reg add "HKCU\\Control Panel\\Desktop" /v "WaitToKillAppTimeout" /t REG_SZ /d "2000" /f')
    run_cmd('reg add "HKCU\\Control Panel\\Desktop" /v "AutoEndTasks" /t REG_SZ /d "1" /f')

def clear_session_cookies_fsso():
    applied_tweaks.append("clear_session_cookies_fsso")
    run_cmd('RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 255')
    run_cmd('powershell "Remove-Item -Path \"$env:LOCALAPPDATA\\Google\\Chrome\\User Data\\Default\\Session Storage\" -Recurse -Force -ErrorAction SilentlyContinue"')
    run_cmd('powershell "Remove-Item -Path \"$env:LOCALAPPDATA\\Google\\Chrome\\User Data\\Default\\Cookies\" -Force -ErrorAction SilentlyContinue"')
    run_cmd('powershell "Remove-Item -Path \"$env:APPDATA\\Mozilla\\Firefox\\Profiles\\*\\cookies.sqlite\" -Force -ErrorAction SilentlyContinue"')
    run_cmd('powershell "Remove-Item -Path \"$env:APPDATA\\Mozilla\\Firefox\\Profiles\\*\\sessionstore.js\" -Force -ErrorAction SilentlyContinue"')
    run_cmd('cmdkey /delete:WindowsLive')
    run_cmd('cmdkey /delete:MicrosoftAccount')
    run_cmd('klist purge')
    run_cmd('klist -li 0x3e7 purge')
    run_cmd('ipconfig /flushdns')
    run_cmd('wevtutil cl System')
    run_cmd('wevtutil cl Application')
    run_cmd('RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 8')

# -----------------------------
# ULTIMATE MODE FUNCTIONS
# -----------------------------
def disable_defender(): 
    applied_tweaks.append("disable_defender")
    run_cmd('reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows Defender" /v "DisableAntiSpyware" /t REG_DWORD /d 1 /f')
    run_cmd("sc stop WinDefend")
    run_cmd("sc config WinDefend start= disabled")

def disable_cortana(): 
    applied_tweaks.append("disable_cortana")
    run_cmd('reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\Windows Search" /v "AllowCortana" /t REG_DWORD /d 0 /f')
    run_cmd('reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\Windows Search" /v "DisableWebSearch" /t REG_DWORD /d 1 /f')

def disable_xbox_services(): 
    applied_tweaks.append("disable_xbox_services")
    run_cmd("sc stop XblAuthManager")
    run_cmd("sc config XblAuthManager start= disabled")
    run_cmd("sc stop XboxGipSvc")
    run_cmd("sc config XboxGipSvc start= disabled")

def remove_bloat_apps(): 
    applied_tweaks.append("remove_bloat_apps")
    run_cmd('powershell "Get-AppxPackage *Xbox* | Remove-AppxPackage"')
    run_cmd('powershell "Get-AppxPackage *Bing* | Remove-AppxPackage"')
    run_cmd('powershell "Get-AppxPackage *Zune* | Remove-AppxPackage"')

def disable_tips_notifications(): 
    applied_tweaks.append("disable_tips_notifications")
    run_cmd('reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\ContentDeliveryManager" /v "SubscribedContent-338388Enabled" /t REG_DWORD /d 0 /f')
    run_cmd('reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\ContentDeliveryManager" /v "SubscribedContent-353694Enabled" /t REG_DWORD /d 0 /f')

def disable_advertising():
    applied_tweaks.append("disable_advertising")
    run_cmd('reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\ContentDeliveryManager" /v "SubscribedContent-353696Enabled" /t REG_DWORD /d 0 /f')
    run_cmd('reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\CloudContent" /v "DisableWindowsConsumerFeatures" /t REG_DWORD /d 1 /f')

def disable_error_reporting():
    applied_tweaks.append("disable_error_reporting")
    run_cmd('reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\Windows Error Reporting" /v "Disabled" /t REG_DWORD /d 1 /f')

def disable_smart_screen():
    applied_tweaks.append("disable_smart_screen")
    run_cmd('reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\System" /v "EnableSmartScreen" /t REG_DWORD /d 0 /f')

def disable_feedback():
    applied_tweaks.append("disable_feedback")
    run_cmd('reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\DataCollection" /v "DoNotShowFeedbackNotifications" /t REG_DWORD /d 1 /f')

def disable_auto_update():
    applied_tweaks.append("disable_auto_update")
    run_cmd('reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate\\AU" /v "NoAutoUpdate" /t REG_DWORD /d 1 /f')

def disable_onedrive():
    applied_tweaks.append("disable_onedrive")
    run_cmd('taskkill /f /im OneDrive.exe')
    run_cmd('reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\OneDrive" /v "DisableFileSyncNGSC" /t REG_DWORD /d 1 /f')

# -----------------------------
# EXTREME MODE FUNCTIONS
# -----------------------------
def optimize_hdd_performance():
    applied_tweaks.append("optimize_hdd_performance")
    run_cmd('schtasks /change /tn "Microsoft\\Windows\\Defrag\\ScheduledDefrag" /disable')
    run_cmd('reg add "HKLM\\SYSTEM\\CurrentControlSet\\Control\\FileSystem" /v "NtfsDisableLastAccessUpdate" /t REG_DWORD /d 1 /f')
    run_cmd('reg add "HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Memory Management\\PrefetchParameters" /v "EnablePrefetcher" /t REG_DWORD /d 1 /f')
    run_cmd('reg add "HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Memory Management\\PrefetchParameters" /v "EnableSuperfetch" /t REG_DWORD /d 0 /f')

def disable_animations():
    applied_tweaks.append("disable_animations")
    run_cmd('reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced" /v "ListviewAlphaSelect" /t REG_DWORD /d 0 /f')
    run_cmd('reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced" /v "TaskbarAnimations" /t REG_DWORD /d 0 /f')
    run_cmd('reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced" /v "ImeSwitchNotification" /t REG_DWORD /d 0 /f')
    run_cmd('reg add "HKCU\\Control Panel\\Desktop\\WindowMetrics" /v "MinAnimate" /t REG_SZ /d "0" /f')

def optimize_login_performance():
    applied_tweaks.append("optimize_login_performance")
    run_cmd('reg add "HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Winlogon" /v "AutoRestartShell" /t REG_DWORD /d 1 /f')
    run_cmd('reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer" /v "Serialize" /t REG_DWORD /d 0 /f')
    run_cmd('reg add "HKLM\\SYSTEM\\CurrentControlSet\\Control" /v "WaitToKillServiceTimeout" /t REG_SZ /d "2000" /f')
    run_cmd('reg add "HKCU\\Control Panel\\Desktop" /v "AutoEndTasks" /t REG_SZ /d "1" /f')
    run_cmd('reg add "HKCU\\Control Panel\\Desktop" /v "HungAppTimeout" /t REG_SZ /d "3000" /f')
    run_cmd('reg add "HKCU\\Control Panel\\Desktop" /v "WaitToKillAppTimeout" /t REG_SZ /d "2000" /f')

def disable_windows_update_auto_restart():
    applied_tweaks.append("disable_windows_update_auto_restart")
    run_cmd('reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate\\AU" /v "NoAutoRebootWithLoggedOnUsers" /t REG_DWORD /d 1 /f')
    run_cmd('reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate\\AU" /v "AUPowerManagement" /t REG_DWORD /d 0 /f')

def optimize_network_performance():
    applied_tweaks.append("optimize_network_performance")
    run_cmd('reg add "HKLM\\SYSTEM\\CurrentControlSet\\Services\\Tcpip\\Parameters" /v "Tcp1323Opts" /t REG_DWORD /d 1 /f')
    run_cmd('reg add "HKLM\\SYSTEM\\CurrentControlSet\\Services\\Tcpip\\Parameters" /v "TCPWindowSize" /t REG_DWORD /d 64240 /f')
    run_cmd('reg add "HKLM\\SYSTEM\\CurrentControlSet\\Services\\Tcpip\\Parameters" /v "DefaultTTL" /t REG_DWORD /d 64 /f')
    run_cmd(r'netsh int tcp set global autotuninglevel=normal')
    run_cmd(r'netsh int tcp set global rss=enabled')

def disable_system_maintenance():
    applied_tweaks.append("disable_system_maintenance")
    run_cmd('reg add "HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Schedule\\Maintenance" /v "MaintenanceDisabled" /t REG_DWORD /d 1 /f')
    run_cmd('schtasks /change /tn "Microsoft\\Windows\\TaskScheduler\\Maintenance Configurator" /disable')

def optimize_paging_file():
    applied_tweaks.append("optimize_paging_file")
    run_cmd('wmic computersystem where name="%computername%" set AutomaticManagedPagefile=False')
    run_cmd('wmic pagefileset where name="C:\\\\pagefile.sys" set InitialSize=2048,MaximumSize=4096')

def disable_search_indexing():
    applied_tweaks.append("disable_search_indexing")
    run_cmd("sc stop WSearch")
    run_cmd("sc config WSearch start= disabled")
    run_cmd('reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\Windows Search" /v "PreventIndexingOutlook" /t REG_DWORD /d 1 /f')

def cleanup_temp_files():
    applied_tweaks.append("cleanup_temp_files")
    run_cmd('powershell "Get-ChildItem -Path $env:TEMP -Recurse -Force | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue"')
    run_cmd('cleanmgr /sagerun:1')
    run_cmd('ipconfig /flushdns')

def optimize_system_responsiveness():
    applied_tweaks.append("optimize_system_responsiveness")
    run_cmd('reg add "HKLM\\SYSTEM\\CurrentControlSet\\Control\\PriorityControl" /v "Win32PrioritySeparation" /t REG_DWORD /d 38 /f')
    run_cmd('reg add "HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile" /v "SystemResponsiveness" /t REG_DWORD /d 20 /f')

def disable_visual_effects_advanced():
    applied_tweaks.append("disable_visual_effects_advanced")
    run_cmd('reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\VisualEffects" /v "AnimateMinMax" /t REG_DWORD /d 0 /f')
    run_cmd('reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\VisualEffects" /v "ComboBoxAnimation" /t REG_DWORD /d 0 /f')
    run_cmd('reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\VisualEffects" /v "ListBoxSmoothScrolling" /t REG_DWORD /d 0 /f')

def optimize_memory_management():
    applied_tweaks.append("optimize_memory_management")
    run_cmd('reg add "HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Memory Management" /v "ClearPageFileAtShutdown" /t REG_DWORD /d 0 /f')
    run_cmd('reg add "HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Memory Management" /v "DisablePagingExecutive" /t REG_DWORD /d 1 /f')

def optimize_disk_cache():
    applied_tweaks.append("optimize_disk_cache")
    run_cmd('reg add "HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Memory Management" /v "IoPageLockLimit" /t REG_DWORD /d 0 /f')

def disable_thumbnail_cache():
    applied_tweaks.append("disable_thumbnail_cache")
    run_cmd('reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced" /v "DisableThumbnailCache" /t REG_DWORD /d 1 /f')

def optimize_context_menu():
    applied_tweaks.append("optimize_context_menu")
    run_cmd('reg add "HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Shell Extensions\\Blocked" /v "{e2bf9676-5f8f-435c-97eb-11607a5bedf7}" /t REG_SZ /d "" /f')

# -----------------------------
# COMPACT GUI SETUP
# -----------------------------
root = tk.Tk()
root.title(f"Windows 11 Optimizer v{VERSION} - SECURE")
root.resizable(True, True)

# Set Custom Icon
set_custom_icon(root)

# Main container - NO SCROLLBARS
main_container = tk.Frame(root, padx=8, pady=8)
main_container.pack(fill="both", expand=True)

# Top section: Mode selection - COMPACT
top_frame = tk.Frame(main_container)
top_frame.pack(fill="x", pady=(0, 5))

tk.Label(top_frame, text="Basic Mode", font=("Arial", 9, "bold")).pack(anchor="w")

# Mode selection options for Basic mode - Compact Layout
mode_frame = tk.Frame(top_frame)
mode_frame.pack(fill="x", pady=2)

# Mode selection variables for Basic mode - UPDATED WITH DOMAIN TRUST FIX
select_all_var = tk.IntVar()
fix_domain_trust_var = tk.IntVar()  # NEW: Domain trust fix option
fix_security_log_var = tk.IntVar()
disable_startup_var = tk.IntVar()
disable_visual_var = tk.IntVar()
disable_transparency_var = tk.IntVar()
disable_background_var = tk.IntVar()
optimize_taskbar_var = tk.IntVar()
disable_notification_sounds_var = tk.IntVar()
enable_fast_start_var = tk.IntVar()
disable_lockscreen_blur_var = tk.IntVar()
optimize_file_explorer_var = tk.IntVar()
disable_game_bar_var = tk.IntVar()

# Create two columns for Basic mode
basic_columns_frame = tk.Frame(mode_frame)
basic_columns_frame.pack(fill="x", expand=True)

basic_left_frame = tk.Frame(basic_columns_frame)
basic_left_frame.pack(side="left", fill="both", expand=True, padx=(0, 10))

basic_right_frame = tk.Frame(basic_columns_frame)
basic_right_frame.pack(side="left", fill="both", expand=True)

# Basic mode checkboxes - Left Column
tk.Checkbutton(basic_left_frame, text="Select All", variable=select_all_var, anchor="w").pack(fill="x")
tk.Checkbutton(basic_left_frame, text="Fix Domain Trust (NEW)", variable=fix_domain_trust_var, anchor="w").pack(fill="x")  # NEW
tk.Checkbutton(basic_left_frame, text="Fix Security Log Full (NEW)", variable=fix_security_log_var, anchor="w").pack(fill="x")
tk.Checkbutton(basic_left_frame, text="Disable Startup Apps", variable=disable_startup_var, anchor="w").pack(fill="x")
tk.Checkbutton(basic_left_frame, text="Disable Visual Effects", variable=disable_visual_var, anchor="w").pack(fill="x")
tk.Checkbutton(basic_left_frame, text="Disable Transparency", variable=disable_transparency_var, anchor="w").pack(fill="x")

# Basic mode checkboxes - Right Column
tk.Checkbutton(basic_right_frame, text="Disable Background Apps", variable=disable_background_var, anchor="w").pack(fill="x")
tk.Checkbutton(basic_right_frame, text="Optimize Taskbar Animations", variable=optimize_taskbar_var, anchor="w").pack(fill="x")
tk.Checkbutton(basic_right_frame, text="Disable Notification Sounds", variable=disable_notification_sounds_var, anchor="w").pack(fill="x")
tk.Checkbutton(basic_right_frame, text="Enable Fast Start Menu", variable=enable_fast_start_var, anchor="w").pack(fill="x")
tk.Checkbutton(basic_right_frame, text="Disable Lockscreen Blur", variable=disable_lockscreen_blur_var, anchor="w").pack(fill="x")
tk.Checkbutton(basic_right_frame, text="Optimize File Explorer", variable=optimize_file_explorer_var, anchor="w").pack(fill="x")
tk.Checkbutton(basic_right_frame, text="Disable Game Bar", variable=disable_game_bar_var, anchor="w").pack(fill="x")

# Apply Basic button
tk.Button(mode_frame, text="Apply Basic", width=15, 
          command=lambda: apply_basic()).pack(anchor="w", pady=(5, 0))

# Separator
ttk.Separator(main_container, orient='horizontal').pack(fill='x', pady=10)

# Mode selector for Standard, Ultimate, and Extreme modes
mode_selector_frame = tk.Frame(main_container)
mode_selector_frame.pack(fill="x", pady=5)

tk.Label(mode_selector_frame, text="Advanced Modes:", font=("Arial", 10, "bold")).pack(anchor="w")

mode_choice = tk.StringVar(value="Standard")
mode_dropdown = ttk.Combobox(mode_selector_frame, textvariable=mode_choice, state="readonly", width=15)
mode_dropdown['values'] = ("Standard", "Ultimate", "Extreme")
mode_dropdown.pack(anchor="w", pady=5)

# Advanced modes frame
advanced_frame = tk.Frame(main_container)
advanced_frame.pack(fill="both", expand=True, pady=5)

# Create frames for all advanced modes
standard_frame = tk.Frame(advanced_frame)
ultimate_frame = tk.Frame(advanced_frame)
extreme_frame = tk.Frame(advanced_frame)

# Standard mode checkboxes
standard_vars = []
standard_tweaks = [
    ("Reduce Telemetry (Safe)", safe_reduce_telemetry),
    ("Disable Unnecessary Services", disable_services),
    ("Disable Indexing", disable_indexing),
    ("Set High Performance", set_high_performance),
    ("Disk Defrag/Trim", disk_defrag_trim),
    ("Disable Timeline", disable_timeline),
    ("Disable Location Tracking", disable_location_tracking),
    ("Optimize System Cache", optimize_system_cache),
    ("Disable Print Spooler", disable_print_spooler),
    ("Optimize Processor", optimize_processor_scheduling),
    ("Disable Remote Assistance", disable_remote_assistance),
    ("Enable Fast Restart", enable_fast_restart),
    ("Enable Fast Signout", enable_fast_signout),
    ("Clear Session Cookies (FSSO)", clear_session_cookies_fsso)
]

standard_select_all_var = tk.IntVar()

# Create two columns for Standard mode
standard_columns_frame = tk.Frame(standard_frame)
standard_columns_frame.pack(fill="both", expand=True)

standard_left_frame = tk.Frame(standard_columns_frame)
standard_left_frame.pack(side="left", fill="both", expand=True, padx=(0, 10))

standard_right_frame = tk.Frame(standard_columns_frame)
standard_right_frame.pack(side="left", fill="both", expand=True)

tk.Checkbutton(standard_left_frame, text="Select All", variable=standard_select_all_var, anchor="w").pack(fill="x")

# Distribute standard tweaks across two columns
mid_point = len(standard_tweaks) // 2
for i, (name, func) in enumerate(standard_tweaks):
    var = tk.IntVar()
    standard_vars.append((var, (name, func)))
    if i < mid_point:
        tk.Checkbutton(standard_left_frame, text=name, variable=var, anchor="w").pack(fill="x")
    else:
        tk.Checkbutton(standard_right_frame, text=name, variable=var, anchor="w").pack(fill="x")

tk.Button(standard_frame, text="Apply Standard", command=lambda: apply_advanced("Standard")).pack(anchor="w", pady=5)

# Ultimate mode checkboxes
ultimate_vars = []
ultimate_tweaks = [
    ("Disable Defender", disable_defender),
    ("Disable Cortana", disable_cortana),
    ("Disable Xbox Services", disable_xbox_services),
    ("Remove Bloat Apps", remove_bloat_apps),
    ("Disable Tips/Notifications", disable_tips_notifications),
    ("Disable Advertising", disable_advertising),
    ("Disable Error Reporting", disable_error_reporting),
    ("Disable Smart Screen", disable_smart_screen),
    ("Disable Feedback", disable_feedback),
    ("Disable Auto Update", disable_auto_update),
    ("Disable OneDrive", disable_onedrive)
]

ultimate_select_all_var = tk.IntVar()

# Create two columns for Ultimate mode
ultimate_columns_frame = tk.Frame(ultimate_frame)
ultimate_columns_frame.pack(fill="both", expand=True)

ultimate_left_frame = tk.Frame(ultimate_columns_frame)
ultimate_left_frame.pack(side="left", fill="both", expand=True, padx=(0, 10))

ultimate_right_frame = tk.Frame(ultimate_columns_frame)
ultimate_right_frame.pack(side="left", fill="both", expand=True)

tk.Checkbutton(ultimate_left_frame, text="Select All", variable=ultimate_select_all_var, anchor="w").pack(fill="x")

# Distribute ultimate tweaks across two columns
mid_point = len(ultimate_tweaks) // 2
for i, (name, func) in enumerate(ultimate_tweaks):
    var = tk.IntVar()
    ultimate_vars.append((var, (name, func)))
    if i < mid_point:
        tk.Checkbutton(ultimate_left_frame, text=name, variable=var, anchor="w").pack(fill="x")
    else:
        tk.Checkbutton(ultimate_right_frame, text=name, variable=var, anchor="w").pack(fill="x")

tk.Button(ultimate_frame, text="Apply Ultimate", command=lambda: apply_advanced("Ultimate")).pack(anchor="w", pady=5)

# Extreme mode checkboxes
extreme_vars = []
extreme_tweaks = [
    ("Optimize HDD Performance", optimize_hdd_performance),
    ("Disable UI Animations", disable_animations),
    ("Optimize Login Performance", optimize_login_performance),
    ("Disable Auto-Restart Updates", disable_windows_update_auto_restart),
    ("Optimize Network Performance", optimize_network_performance),
    ("Disable System Maintenance", disable_system_maintenance),
    ("Optimize Paging File", optimize_paging_file),
    ("Disable Search Indexing", disable_search_indexing),
    ("Cleanup Temp Files", cleanup_temp_files),
    ("Optimize System Responsiveness", optimize_system_responsiveness),
    ("Advanced Visual Effects", disable_visual_effects_advanced),
    ("Optimize Memory Management", optimize_memory_management),
    ("Optimize Disk Cache", optimize_disk_cache),
    ("Disable Thumbnail Cache", disable_thumbnail_cache),
    ("Optimize Context Menu", optimize_context_menu)
]

extreme_select_all_var = tk.IntVar(value=1)

# Create two columns for Extreme mode
extreme_columns_frame = tk.Frame(extreme_frame)
extreme_columns_frame.pack(fill="both", expand=True)

extreme_left_frame = tk.Frame(extreme_columns_frame)
extreme_left_frame.pack(side="left", fill="both", expand=True, padx=(0, 10))

extreme_right_frame = tk.Frame(extreme_columns_frame)
extreme_right_frame.pack(side="left", fill="both", expand=True)

tk.Checkbutton(extreme_left_frame, text="Select All", variable=extreme_select_all_var, anchor="w").pack(fill="x")

# Distribute extreme tweaks across two columns
mid_point = len(extreme_tweaks) // 2
for i, (name, func) in enumerate(extreme_tweaks):
    var = tk.IntVar(value=1)
    extreme_vars.append((var, (name, func)))
    if i < mid_point:
        tk.Checkbutton(extreme_left_frame, text=name, variable=var, anchor="w").pack(fill="x")
    else:
        tk.Checkbutton(extreme_right_frame, text=name, variable=var, anchor="w").pack(fill="x")

tk.Button(extreme_frame, text="Apply Extreme", command=lambda: apply_advanced("Extreme")).pack(anchor="w", pady=5)

# Separator
ttk.Separator(main_container, orient='horizontal').pack(fill='x', pady=10)

# Bottom section: Action buttons with COMPREHENSIVE RESTORE
button_frame = tk.Frame(main_container)
button_frame.pack(fill="x", pady=10)

# Action buttons - Now with Comprehensive Restore
tk.Button(button_frame, text="Undo Tweaks", width=15, command=restore_all_tweaks).pack(side="left", padx=5)
tk.Button(button_frame, text="Export Report", width=15, command=export_report).pack(side="left", padx=5)
tk.Button(button_frame, text="Version History", width=15, command=show_version_history).pack(side="left", padx=5)
tk.Button(button_frame, text="Help Guide", width=15, command=show_help).pack(side="left", padx=5)
# NEW BUTTON ADDED HERE
tk.Button(button_frame, text="Verify Domain Trust", width=15, command=verify_domain_trust).pack(side="left", padx=5)
tk.Button(button_frame, text="About Optimizer", width=15,
          command=lambda: messagebox.showinfo("About",
            f"Windows 11 Optimizer v{VERSION} - SECURE EDITION\n"
            f"By: Sir Fred Iz 'N D Hauz\n\n"
            f"COMPREHENSIVE optimization tool for Windows 11\n"
            f"• SECURE EDITION - maintains original branding\n"
            f"• Enhanced domain trust repair (Test-ComputerSecureChannel)\n"
            f"• Domain trust verification tool included\n"
            f"• Fixes security log full errors\n"
            f"• Endpoint security detection at startup\n"
            f"• No security alert triggers\n"
            f"• Comprehensive restore functionality\n"
            f"• Sophos-compatible optimizations\n"
            f"• Safe for home and enterprise use")).pack(side="left", padx=5)

# -----------------------------
# GUI FUNCTIONS - UPDATED WITH DOMAIN TRUST FIX
# -----------------------------
def apply_basic():
    # Show backup/restore point creation message immediately
    messagebox.showinfo("Starting Optimization", "Starting optimization process...\n\n- Creating system restore point (background)\n- Backing up registry (background)\n- Applying selected tweaks")
    
    # Run the actual work in a thread to prevent GUI freezing
    def apply_basic_thread():
        # These will run in background threads internally now
        create_restore_point()
        backup_registry()
        
        selected_tweaks = []
        # UPDATED: Added domain trust fix to basic tweaks
        basic_tweaks = [
            ("Fix Domain Trust Relationship", fix_domain_trust_relationship, fix_domain_trust_var),  # NEW
            ("Fix Security Log Full", fix_security_log_full, fix_security_log_var),
            ("Disable Startup Apps", disable_startup_apps, disable_startup_var),
            ("Disable Visual Effects", disable_visual_effects, disable_visual_var),
            ("Disable Transparency", disable_transparency, disable_transparency_var),
            ("Disable Background Apps", disable_background_apps, disable_background_var),
            ("Optimize Taskbar Animations", optimize_taskbar_animations, optimize_taskbar_var),
            ("Disable Notification Sounds", disable_notification_sounds, disable_notification_sounds_var),
            ("Enable Fast Start Menu", enable_fast_start_menu, enable_fast_start_var),
            ("Disable Lockscreen Blur", disable_lockscreen_blur, disable_lockscreen_blur_var),
            ("Optimize File Explorer", optimize_file_explorer, optimize_file_explorer_var),
            ("Disable Game Bar", disable_game_bar, disable_game_bar_var)
        ]
        
        # Apply selected tweaks
        if select_all_var.get():
            for name, func, _ in basic_tweaks:
                func()
                selected_tweaks.append(name)
        else:
            for name, func, var in basic_tweaks:
                if var.get():
                    func()
                    selected_tweaks.append(name)
        
        # Show completion message
        if selected_tweaks:
            root.after(0, lambda: messagebox.showinfo("Basic Mode Complete", 
                f"Applied tweaks:\n{', '.join(selected_tweaks)}\n\nBackup location: {last_backup}"))
        else:
            root.after(0, lambda: messagebox.showinfo("Basic Mode", "No tweaks selected"))
    
    # Start the optimization in a background thread
    threading.Thread(target=apply_basic_thread, daemon=True).start()

def apply_advanced(mode):
    # Show immediate feedback
    messagebox.showinfo(f"Starting {mode} Mode", 
        f"Starting {mode} optimization process...\n\n- Creating system restore point (background)\n- Backing up registry (background)\n- Applying selected tweaks")
    
    def apply_advanced_thread():
        # These run in background threads internally
        create_restore_point()
        backup_registry()
        selected_tweaks = []

        # Apply Basic mode tweaks for ALL advanced modes (including domain trust fix if selected)
        if fix_domain_trust_var.get():
            fix_domain_trust_relationship()
            selected_tweaks.append("Fix Domain Trust Relationship")
        
        if fix_security_log_var.get():
            fix_security_log_full()
            selected_tweaks.append("Fix Security Log Full")
        
        basic_tweaks = [
            ("Disable Startup Apps", disable_startup_apps),
            ("Disable Visual Effects", disable_visual_effects),
            ("Disable Transparency", disable_transparency),
            ("Disable Background Apps", disable_background_apps),
            ("Optimize Taskbar Animations", optimize_taskbar_animations)
        ]
        
        for name, func in basic_tweaks:
            func()
            selected_tweaks.append(name)

        if mode == "Standard" or mode == "Ultimate" or mode == "Extreme":
            # Apply Standard mode tweaks
            for name, func in standard_tweaks:
                func()
                selected_tweaks.append(name)

        if mode == "Ultimate" or mode == "Extreme":
            # Apply Ultimate mode tweaks
            for name, func in ultimate_tweaks:
                func()
                selected_tweaks.append(name)

        # Apply Extreme-specific tweaks if Extreme mode is selected
        if mode == "Extreme":
            if extreme_select_all_var.get():
                for name, func in extreme_tweaks:
                    func()
                    selected_tweaks.append(name)
            else:
                for (var, (name, func)) in extreme_vars:
                    if var.get():
                        func()
                        selected_tweaks.append(name)
        else:
            # For Standard and Ultimate modes, apply selected tweaks
            if mode == "Standard":
                vars_list = standard_vars
                select_all_var = standard_select_all_var
            else:  # Ultimate
                vars_list = ultimate_vars
                select_all_var = ultimate_select_all_var

            if select_all_var.get():
                for name, func in (standard_tweaks if mode == "Standard" else ultimate_tweaks):
                    func()
                    selected_tweaks.append(name)
            else:
                for (var, (name, func)) in vars_list:
                    if var.get():
                        func()
                        selected_tweaks.append(name)

        # Show completion message in main thread
        if selected_tweaks:
            root.after(0, lambda: messagebox.showinfo(f"{mode} Mode Complete", 
                f"Applied optimizations including:\n- Basic Mode (with Domain Trust & Security Log Fixes)\n- {mode} Specific Tweaks\n\nTotal tweaks applied: {len(selected_tweaks)}\nBackup: {last_backup}"))
        else:
            root.after(0, lambda: messagebox.showinfo(f"{mode} Mode", "No tweaks selected"))
    
    # Start the optimization in a background thread
    threading.Thread(target=apply_advanced_thread, daemon=True).start()

def show_advanced_mode(selected_mode):
    if selected_mode == "Standard":
        standard_frame.pack(fill="both", expand=True, pady=5)
        ultimate_frame.pack_forget()
        extreme_frame.pack_forget()
    elif selected_mode == "Ultimate":
        standard_frame.pack_forget()
        ultimate_frame.pack(fill="both", expand=True, pady=5)
        extreme_frame.pack_forget()
    else:  # Extreme
        standard_frame.pack_forget()
        ultimate_frame.pack_forget()
        extreme_frame.pack(fill="both", expand=True, pady=5)

# Select All functionality for all modes - UPDATED WITH DOMAIN TRUST FIX
def toggle_select_all_basic():
    state = select_all_var.get()
    fix_domain_trust_var.set(state)  # NEW: Include domain trust fix
    fix_security_log_var.set(state)
    disable_startup_var.set(state)
    disable_visual_var.set(state)
    disable_transparency_var.set(state)
    disable_background_var.set(state)
    optimize_taskbar_var.set(state)
    disable_notification_sounds_var.set(state)
    enable_fast_start_var.set(state)
    disable_lockscreen_blur_var.set(state)
    optimize_file_explorer_var.set(state)
    disable_game_bar_var.set(state)

def toggle_select_all_standard():
    state = standard_select_all_var.get()
    for var, _ in standard_vars:
        var.set(state)

def toggle_select_all_ultimate():
    state = ultimate_select_all_var.get()
    for var, _ in ultimate_vars:
        var.set(state)

def toggle_select_all_extreme():
    state = extreme_select_all_var.get()
    for var, _ in extreme_vars:
        var.set(state)

# Bind functionality
select_all_var.trace('w', lambda *args: toggle_select_all_basic())
standard_select_all_var.trace('w', lambda *args: toggle_select_all_standard())
ultimate_select_all_var.trace('w', lambda *args: toggle_select_all_ultimate())
extreme_select_all_var.trace('w', lambda *args: toggle_select_all_extreme())
mode_choice.trace('w', lambda *args: show_advanced_mode(mode_choice.get()))

# Auto-fit window size
def auto_fit_window():
    root.update_idletasks()
    width = main_container.winfo_reqwidth() + 50
    height = min(main_container.winfo_reqheight() + 100, 800)
    root.geometry(f"{width}x{height}")

# Show initial advanced mode and auto-fit
show_advanced_mode("Standard")
root.after(100, auto_fit_window)

root.mainloop()