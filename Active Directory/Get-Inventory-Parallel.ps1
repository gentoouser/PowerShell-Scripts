<#
.SYNOPSIS
  Name: Get-Inventory-Parallel.ps1
  The purpose of this script is to create a simple inventory.
  
.DESCRIPTION
  This is a simple script to retrieve all computer objects in Active Directory and then connect
  to each one and gather basic hardware information using Cim. The information includes Manufacturer,
  Model,Serial Number, CPU, RAM, Disks, Operating System, Sound Listening, Graphics Card Controller, etc ...

.RELATED LINKS
  https://www.sconstantinou.com
  https://gallery.technet.microsoft.com/scriptcenter/PowerShell-Hardware-f99336f6


.PARAMETER OutputFolder
    Folder where the inventory output file is saved.
.PARAMETER MaxJobs
    Maximum number of computers to inventory at once.
.PARAMETER SleepTime
    Time to wait before testing for more finished jobs.
.PARAMETER TimeOut
    Maximum time in seconds to wait for before canceling command/operation. 
.PARAMETER LogHistory
    Negative number for number of days to log at past IIS log files.
.PARAMETER MaxMemoryUseage
    Maximum percentage of ram to use parsing log files. Without this large logs would crash the server while parsing. 
.PARAMETER WSO
    Share where WSUSOffline files are stored. Used with -InstallWMF
.PARAMETER InstallWMF
    Try to install Windows Management Framework update. Needs valid -WSO parameter to work.
.PARAMETER SetupWinRM
    Try to setup Windows Remote Management on remote computer if cannot connect to WinRM.
.PARAMETER RemoveSMB1
    Try to Remove or Disable SMB 1 protocol from computer. 
.PARAMETER Excel
    [Experimental] Export output to excel file. Exported file shows up corrupted. Also Need to see how to format cells to have multiple lines instead of commas.
.PARAMETER Verbose
    Debugging output
.PARAMETER VIServers
    Array of vCenters servers to connect and get VM inventory from. 
.NOTES
  Updated:      01-06-2018        - Replaced Get-CimInstance cmdlet with Get-CimInstance
                                  - Added Serial Number Information
                                  - Added Sound Device Information
                                  - Added Video Controller Information
                                  - Added option to send CSV file through email
                                  - Added parameters to enable email function option
                03/21/2019        - Added IP address
                                  - Added Software
                                  - Added Installed Roles
                                  - Added Installed Features
                09/24/2019        - Added DNS Server
                                  - Added DNS Search Suffix
                01/27/2020        - Added Certificates Expiration info for Listening connections.
                                  - Added Jobs to process computers in parallel
                11/09/2020        - Added Services and IIS info   
                11/10/2020        - Filtered disabled computer
                                  - Cleanup errors 
                                  - Reduce AD Calls
                                  - Fix video bug
                11/11/2020        - Cleaned up names
                                  - Get DNS Record IP
                                  - Ping by FQDN
                                  - Fixed Services Name
                                  - Fixed Drive info
                11/16/2020        - Fixed DNS Records null issue
                                  - Formatting fixes
                12/03/2020        - Added ability to remove SMB1 ability
                                  - Added Check for Reboot needed
                                  - Fixed Formatting
                12/05/2020        - Added PSVersion Fixes for SMB1
                12/08/2020        - Moved DN to end of list and more fixed for SMB1
                12/10/2020        - SMB1 detection fixed.  
                01/06/2021        - Added Hit Count to IIS sites.  
                01/07/2021        - Removed Hit Count as it caused computers to lock up
                03/24/2021        - IIS log memory work around
                04/07/2021        - Added SQL databases, FortiClient Version,LanDesk Agent detection.
                06/21/2021        - Remove AD module dependency in threads. Fix Errors.
                08/24/2021        - Added LAN Manager authentication level
                08/31/2021        - Added Last Hotfixes and Export to Excel directly
                09/29/2021        - Fixed output formatting. Added TLS info.
                09/30/2021        - Create one WMI and PSSession per computer instead of for each command. Also cleaned up reported IPs. Added information from vCenter about the VMs. Cleaned up progressbar reporting.
                11/08/2021        - Added switch SetupWinRM to configure WinRM on remote computer using PSExec. Tweaked progressbar reporting. Fixed issue with LANDesk Agent reporting wrong status
                02/28/2022        - Switched to use class instead of PSCustomObject.
                05/10/2022        - Added more variable checks and TRY{}/Catch{}'s to reduce errors in transcript. 

  Release Date: 10-02-2018
   
  Author: Stephanos Constantinou

.EXAMPLES
  Get-Inventory-Parallel.ps1
  Find the output under Logs in the script directory

#>

Param(
    [String]$OutputFolder = ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\Logs"),
    [Int]$MaxJobs = 120,
    [Int]$SleepTime = 5,
    [Int]$TimeOut = 90,
    [Int]$LogHistory = -90,
    [Int]$MaxMemoryUseage = 80,
    [string]$WSO = "\\ucn\share\wsusoffline-12_CE\client",
    [switch]$InstallWMF,
    [switch]$SetupWinRM,
    [switch]$RemoveSMB1,
    [switch]$Excel,
    [switch]$Verbose,
    [array]$VIServers = @()
)
$FileDate = (Get-Date -format yyyyMMdd-hhmm)
$LogFile = ($OutputFolder + "\" + `
           ($MyInvocation.MyCommand.Name -replace ".ps1","") + "_" + `
		   $env:computername + "_" + `
           $FileDate + ".log")
$ScriptVersion = "2.2.01"
$sw = [Diagnostics.Stopwatch]::StartNew()
$Inventory = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
$VMs = [System.Collections.ArrayList]@()
$ScanCount = 1
$UNCPSExecPath = "\\plsfinancial.com\share\IT\Utilities\PSTools\PsExec.exe"
$RemoveServices=@(
    "ActiveX Installer (AxInstSV) - Disabled",
    "App Readiness - Manual",
    "AllJoyn Router Service - Disabled",
    "Application Experience - Manual",
    "Application Identity - Manual",
    "Application Information - Manual",
    "Application Layer Gateway Service - Manual",
    "Application Layer Gateway Service - Disabled",
    "Application Host Helper Service - Auto",
    "Application Management - Manual",
    "AppX Deployment Service (AppXSVC) - Manual",
    "Auto Time Zone Updater - Disabled",
    "AVCTP service - Disabled",
    "Background Intelligent Transfer Service - Manual",
    "Background Tasks Infrastructure Service - Auto",
    "Base Filtering Engine - Auto",
    "Bluetooth Audio Gateway Service - Disabled",
    "Bluetooth Support Service - Disabled",
    "Certificate Propagation - Manual",
    "CNG Key Isolation - Manual",
    "COM+ Event System - Auto",
    "COM+ System Application - Manual",
    "Connected Devices Platform Service - Disabled",
    "Connected User Experiences and Telemetry - Disabled",
    "CoreMessaging - Auto",
    "Credential Manager - Manual",
    "Cryptographic Services - Auto",
    "Data Sharing Service - Manual",
    "DCOM Server Process Launcher - Auto",
    "Delivery Optimization - Manual",
    "Device Association Service - Manual",
    "Device Install Service - Manual",
    "Device Management Enrollment Service - Manual",
    "Device Management Wireless Application Protocol (WAP) Push message Routing Service - Manual",
    "Device Setup Manager - Manual",
    "DHCP Client - Auto",
    "Diagnostic Policy Service - Auto",
    "Diagnostic Policy Service - Disabled",
    "Diagnostic Service Host - Manual",
    "Diagnostic Service Host - Disabled",
    "Diagnostic System Host - Manual",
    "Diagnostic System Host - Disabled",
    "Diagnostics Tracking Service - Auto",
    "Distributed Link Tracking Client - Disabled",
    "Distributed Link Tracking Client - Manual",
    "Distributed Link Tracking Client - Auto",
    "Distributed Transaction Coordinator - Auto",
    "Downloaded Maps Manager - Disabled",
    "DNS Client - Auto",
    "Embedded Mode - Manual",
    "Encrypting File System (EFS) - Manual",
    "Encrypting File System (EFS) - Disabled",
    "Enterprise App Management Service - Manual",
    "Extensible Authentication Protocol - Manual",
    "Function Discovery Provider Host - Manual",
    "Function Discovery Provider Host - Disabled",
    "Function Discovery Resource Publication - Manual",
    "Function Discovery Resource Publication - Disabled",
    "Geolocation Service - Disabled"
    "Group Policy Client - Auto",
    "GraphicsPerfSvc - Disabled",
    "Human Interface Device Service - Manual",
    "HV Host Service - Disabled",
    "Hyper-V Data Exchange Service - Manual",
    "Hyper-V Data Exchange Service - Disabled",
    "Hyper-V Guest Service Interface - Manual",
    "Hyper-V Guest Service Interface - Disabled",
    "Hyper-V Guest Shutdown Service - Manual",
    "Hyper-V Guest Shutdown Service - Disabled",
    "Hyper-V Heartbeat Service - Manual",
    "Hyper-V Heartbeat Service - Disabled",
    "Hyper-V Remote Desktop Virtualization Service - Manual",
    "Hyper-V Remote Desktop Virtualization Service - Disabled",
    "Hyper-V Time Synchronization Service - Manual",
    "Hyper-V Time Synchronization Service - Disabled",
    "Hyper-V Volume Shadow Copy Requestor - Manual",
    "Hyper-V Volume Shadow Copy Requestor - Disabled",
    "IKE and AuthIP IPsec Keying Modules - Auto",
    "Intel Local Scheduler Service - Auto",
    "Intel PDS - Auto",
    "Interactive Services Detection - Manual",
    "Internet Connection Sharing (ICS) - Disabled",
    "Internet Explorer ETW Collector Service - Manual",
    "IP Helper - Auto",
    "IP Helper - Manual",
    "IP Helper - Disabled",
    "IPsec Policy Agent - Manual",
    "KDC Proxy Server service (KPS) - Manual",
    "KtmRm for Distributed Transaction Coordinator - Manual",
    "Link-Layer Topology Discovery Mapper - Manual",
    "Link-Layer Topology Discovery Mapper - Disabled",
    "Local Session Manager - Auto",
    "Microsoft (R) Diagnostics Hub Standard Collector Service - Disabled",
    "Microsoft App-V Client - Disabled",
    "Microsoft iSCSI Initiator Service - Manual",
    "Microsoft iSCSI Initiator Service - Disabled",
    "Microsoft Passport - Manual",
    "Microsoft Passport Container - Manual",
    "Microsoft Software Shadow Copy Provider - Manual",
    "Microsoft Storage Spaces SMP - Auto",
    "Microsoft Storage Spaces SMP - Manual",
    "Microsoft Store Install Service - Disabled",
    "Multimedia Class Scheduler - Manual",
    "Net Driver HPZ12 - Auto",
    "Net.Tcp Port Sharing Service - Disabled",
    "Netlogon - Auto",
    "Network Access Protection Agent - Manual",
    "Network Connections - Manual",
    "Network Connection Broker - Auto",
    "Network Connectivity Assistant - Manual",
    "Network Connectivity Assistant - Disabled",
    "Network List Service - Manual",
    "Network Location Awareness - Auto",
    "Network Setup Service - Manual",
    "Network Store Interface Service - Auto",
    "Offline Files - Disabled",
    "Optimize drives - Manual",
    "Payments and NFC/SE Manager - Disabled",
    "Performance Counter DLL Host - Manual",
    "Performance Logs & Alerts - Manual",
    "Plug and Play - Manual",
    "Phone Service - Disabled",
    "Pml Driver HPZ12 - Auto",
    "Portable Device Enumerator Service - Manual",
    "Portable Device Enumerator Service - Disabled",
    "Power - Disabled",
    "Print Spooler - Manual",
    "Print Spooler - Disabled",
    "Printer Extensions and Notifications - Manual",
    "Problem Reports and Solutions Control Panel Support - Disabled",
    "Program Compatibility Assistant Service - Disabled",
    "Problem Reports and Solutions Control Panel Support - Manual",
    "Quality Windows Audio Video Experience - Disabled",
    "Radio Management Service - Disabled",
    "Remote Access Auto Connection Manager - Manual",
    "Remote Access Connection Manager - Manual",
    "Remote Desktop Configuration - Manual",
    "Remote Desktop Services - Manual",
    "Remote Desktop Services UserMode Port Redirector - Manual",
    "Remote Procedure Call (RPC) - Auto",
    "Remote Procedure Call (RPC) Locator - Manual",
    "Remote Procedure Call (RPC) Locator - Disabled",
    "Resultant Set of Policy Provider - Disabled",
    "Resultant Set of Policy Provider - Manual",
    "Routing and Remote Access - Disabled",
    "RPC Endpoint Mapper - Auto",
    "Secondary Logon - Manual",
    "Secure Socket Tunneling Protocol Service - Manual",
    "Security Accounts Manager - Auto",
    "Sensor Data Service - Disabled",
    "Sensor Monitoring Service - Disabled",
    "Sensor Service - Disabled",
    "Shared PC Account Manager - Disabled",
    "Server - Auto",
    "Shell Hardware Detection - Auto",
    "Smart Card - Auto",
    "Smart Card - Disabled",
    "Smart Card Device Enumeration Service - Manual",
    "Smart Card Device Enumeration Service - Disabled",
    "Smart Card Removal Policy - Manual",
    "Software Protection - Auto",
    "SNMP Trap - Disabled",
    "Special Administration Console Helper - Manual",
    "Spot Verifier - Manual",
    "Spot Verifier - Disabled",
    "SSDP Discovery - Disabled",
    "State Repository Service - Manual",
    "Still Image Acquisition Events - Disabled",
    "Storage Service - Auto",
    "Storage Tiers Management - Manual",
    "Superfetch - Manual",
    "System Event Notification Service - Auto",
    "System Events Broker - Auto",
    "System Guard Runtime Monitor Broker - Manual",
    "Task Scheduler - Auto",
    "TCP/IP NetBIOS Helper - Auto",
    "TCP/IP NetBIOS Helper - Manual",
    "Telephony - Manual",
    "Telephony - Disabled",
    "Themes - Disabled",
    "Themes - Auto",
    "Thread Ordering Server - Manual",
    "Time Broker - Manual",
    "Touch Keyboard and Handwriting Panel Service - Disabled",
    "Update Orchestrator Service - Auto",
    "UPnP Device Host - Disabled",
    "User Access Logging Service - Auto",
    "User Experience Virtualization Service - Disabled",
    "User Profile Service - Auto",
    "Virtual Disk - Manual",
    "VMware Alias Manager and Ticket Service - Auto",
    "VMware Snapshot Provider - Manual",
    "VMware SVGA Helper Service - Auto",
    "VMware Tools - Auto",
    "Volume Shadow Copy - Manual",
    "WalletService - Disabled",
    "Windows Audio - Manual",
    "Windows Audio - Disabled",
    "Windows Audio Endpoint Builder - Manual",
    "Windows Audio Endpoint Builder - Disabled",
    "Windows Color System - Manual",
    "Windows Connection Manager - Auto",
    "Windows Camera Frame Server - Disabled",
    "Windows Defender Advanced Threat Protection Service - Manual",
    "Windows Defender Firewall - Auto",
    "Windows Driver Foundation - User-mode Driver Framework - Manual",
    "Windows Encryption Provider Host Service - Manual",
    "Windows Encryption Provider Host Service - Disabled",
    "Windows Error Reporting Service - Manual",
    "Windows Error Reporting Service - Disabled",
    "Windows Event Collector - Manual",
    "Windows Event Log - Auto",
    "Windows Firewall - Auto",
    "Windows Font Cache Service - Auto",
    "Windows Insider Service - Disabled",
    "Windows Installer - Manual",
    "Windows Management Instrumentation - Auto",
    "Windows Modules Installer - Manual",
    "Windows Mobile Hotspot Service - Disabled",
    "Windows Push Notifications System Service - Disabled",
    "Windows PushToInstall Service - Disabled",
    "Windows Remote Management (WS-Management) - Auto",
    "Windows Search - Disabled",
    "Windows Security Service - Manual",
    "Windows Store Service (WSService) - Manual",
    "Windows Time - Manual",
    "Windows Time - Auto",
    "Windows Update - Manual",
    "Windows Update Medic Service - Manual",
    "WinHTTP Web Proxy Auto-Discovery Service - Manual",
    "Wired AutoConfig - Manual",
    "WMI Performance Adapter - Manual",
    "Workstation - Auto",
    "Xbox Live Auth Manager - Disabled",
    "Xbox Live Game Save - Disabled"
)
[string[]]$KillProcesses=@(
    "TrustedInstaller",

)
$winrmcommands=@(
    'netsh advfirewall firewall add rule dir=in name="DCOM" program=%systemroot%\system32\svchost.exe service=rpcss action=allow protocol=TCP localport=135 remoteip=10.0.0.0/8,172.16.0.0/12,192.168.0.0/16,LocalSubnet profile=any',
    'netsh advfirewall firewall add rule dir=in name="WMI-In" program=%systemroot%\system32\svchost.exe service=winmgmt action = allow protocol=TCP localport=any remoteip=10.0.0.0/8,172.16.0.0/12,192.168.0.0/16,LocalSubnet profile=any',
    'netsh advfirewall firewall add rule dir=in name="UnsecApp" program=%systemroot%\system32\wbem\unsecapp.exe action=allow remoteip=10.0.0.0/8,172.16.0.0/12,192.168.0.0/16,LocalSubnet profile=any',
    'netsh advfirewall firewall add rule dir=in name="WINRM-HTTP" protocol=tcp localport=5985 action=allow remoteip=10.0.0.0/8,172.16.0.0/12,192.168.0.0/16,LocalSubnet profile=any',
    'netsh advfirewall firewall add rule dir=in name="WINRM-HTTPS" protocol=tcp localport=5986 action=allow remoteip=10.0.0.0/8,172.16.0.0/12,192.168.0.0/16,LocalSubnet profile=any',
    'netsh advfirewall firewall add rule name="ICMP Allow incoming V4 echo request" protocol=icmpv4:8,any dir=in action=allow remoteip=10.0.0.0/8,172.16.0.0/12,192.168.0.0/16,LocalSubnet',
    'netsh advfirewall firewall add rule dir=out name="WMI-OUT" program=%systemroot%\system32\svchost.exe service=winmgmt action=allow protocol=TCP localport=any remoteip=10.0.0.0/8,172.16.0.0/12,192.168.0.0/16,LocalSubnet profile=any',
    'PowerShell -Command {New-Item -Path WSMan:\LocalHost\Listener -Transport HTTPS -Address * -CertificateThumbPrint $(Get-ChildItem -Path cert:\LocalMachine\My | Sort-Object -Descending -property NotAfter | Where {$_.Subject -match [System.Net.Dns]::GetHostByName(($env:computerName)).hostname} | Select-Object -first 1).Thumbprint -Force"}',
    'net stop winrm && net start winrm'
)

function FormatElapsedTime($ts) {
  #https://stackoverflow.com/questions/3513650/timing-a-commands-execution-in-powershell
  $elapsedTime = ""
  if ( $ts.Hours -gt 0 ){
      $elapsedTime = [string]::Format( "{0:00} hours {1:00} min. {2:00}.{3:00} sec.", $ts.Hours, $ts.Minutes, $ts.Seconds, $ts.Milliseconds / 10 );
  }else{
      if ( $ts.Minutes -gt 0 ){
          $elapsedTime = [string]::Format( "{0:00} min. {1:00}.{2:00} sec.", $ts.Minutes, $ts.Seconds, $ts.Milliseconds / 10 );
      }else{
          $elapsedTime = [string]::Format( "{0:00}.{1:00} sec.", $ts.Seconds, $ts.Milliseconds / 10 );
      }

      if ($ts.Hours -eq 0 -and $ts.Minutes -eq 0 -and $ts.Seconds -eq 0){
          $elapsedTime = [string]::Format("{0:00} ms.", $ts.Milliseconds);
      }
      if ($ts.Milliseconds -eq 0){
          $elapsedTime = [string]::Format("{0} ms", $ts.TotalMilliseconds);
      }
  }
  return $elapsedTime
}

#Log File
If (-Not [string]::IsNullOrEmpty($LogFile)) {
    try {
        Stop-transcript -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    } catch {

    }
  If (-Not( Test-Path (Split-Path -Path $LogFile -Parent))) {
    New-Item -ItemType directory -Path (Split-Path -Path $LogFile -Parent) | Out-Null
    $Acl = Get-Acl (Split-Path -Path $LogFile -Parent)
    $Ar = New-Object system.Security.AccessControl.FileSystemAccessRule('Users', "FullControl", "ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.Setaccessrule($Ar)
    Set-Acl (Split-Path -Path $LogFile -Parent) $Acl
  }
  try { 
    Start-Transcript -Path $LogFile -Append
  } catch { 
    Stop-transcript
    Start-Transcript -Path $LogFile -Append
  } 
  Write-Host ("Script: " + $MyInvocation.MyCommand.Name)
  Write-Host ("Version: " + $ScriptVersion)
  Write-Host (" ")
}	      

if($verbose) {
    $oldverbose = $VerbosePreference
    $VerbosePreference = "continue" 
}Else {
    $VerbosePreference ="SilentlyContinue"
}
#region Load AD PSSnapin
If (Get-Module -ListAvailable -Name "ActiveDirectory") {
    If (-Not (Get-Module "ActiveDirectory" -ErrorAction SilentlyContinue)) {
        Import-Module "ActiveDirectory"
    } Else {
        write-host "ActiveDirectory PowerShell Module Already Loaded" -foregroundcolor "Green"
    } 
} Else {
	If ((Get-ComputerInfo).OsProductType -eq "Server") {
		Install-WindowsFeature RSAT-AD-PowerShell
		Import-Module "ActiveDirectory"
		If (-Not (Get-Module "ActiveDirectory" -ErrorAction SilentlyContinue)) {
			write-error ("Please install ActiveDirectory Powershell Modules" , "Error")
			exit
		}
	}Else{
		Get-WindowsCapability -Name Rsat.ActiveDirectory* -Online | Add-WindowsCapability -Online
		Import-Module "ActiveDirectory"
		If (-Not (Get-Module "ActiveDirectory" -ErrorAction SilentlyContinue)) {
			write-error ("Please install ActiveDirectory Powershell Modules" , "Error")
			exit
		}
	}
}
#endregion Load AD PSSnapin
#region Load VMWare PSSnapin
$swv = [Diagnostics.Stopwatch]::StartNew()
If (Get-Module -ListAvailable -Name "VMware.PowerCLI") {
    If (-Not (Get-Module "VMware.PowerCLI" -ErrorAction SilentlyContinue)) {
        Import-Module "VMware.PowerCLI"
    } Else {
        write-host "VMware.PowerCLI PowerShell Module Already Loaded" -foregroundcolor "Green"
    } 
} Else {
    Install-Module -Name "VMware.PowerCLI" -Force -Confirm:$false -Scope:AllUsers -SkipPublisherCheck -AllowClobber
    Import-Module "VMware.PowerCLI"
    Set-PowerCLIConfiguration -Scope User -ParticipateInCEIP $false -confirm:$false
}

write-host "Connecting to vCenters"
$ConnectedVI = Connect-VIServer -Server $VIServers
write-host "Getting VM Information"
ForEach ($VI in $ConnectedVI) {
    $shutdownCA = Get-CustomAttribute -Name "Shutdown Date" -Server $VI
    $VM = (Get-View -ViewType VirtualMachine -Server $VI | Select-Object Name, 
        @{N ='Power State'; E = {$_.Runtime.PowerState}},
        @{N ='Notes'; E = {$_.Summary.Config.Annotation.Replace("`n",'')}},
        @{N ="Guest OS"; E= {$_.Guest.Guestfullname}},
        @{N="Tools Status"; E={$_.Guest.Toolsstatus}},
        @{N="Shutdown Date"; E={($_.Summary.CustomValue | Where-Object {$_.Key -eq $shutdownCA.Key}).value}}
    ) 
    $null = $VMs.Add($VM) 
    # [void]$VMs.Add($VM)
}
$swv.Stop()
write-verbose ("Section took: " + (FormatElapsedTime($swv.Elapsed)) + " to run.")
#endregion Load VMWare PSSnapin
#region PSExecPath
If (Test-Path ($env:ProgramFiles + "\SysinternalsSuite\PsExec.exe")){
    $PSExecPath = ($env:ProgramFiles + "\SysinternalsSuite\PsExec.exe")
}Elseif(Test-Path (${env:ProgramFiles(x86)} + "\SysinternalsSuite\PsExec.exe")) {
    $PSExecPath = (${env:ProgramFiles(x86)} + "\SysinternalsSuite\PsExec.exe")
}Else{
    $PSExecPath = $UNCPSExecPath
}
#endregion PSExecPath
if (Get-Job -name *)
{
    Write-Verbose "Removing old jobs."
    Get-Job -name * -ErrorAction SilentlyContinue| Remove-Job -Force -ErrorAction SilentlyContinue | Out-Null
}
$swv = [Diagnostics.Stopwatch]::StartNew()
write-host "Getting AD Computers"
$AllComputers = Get-ADComputer -Filter * -Properties Name,Enabled,OperatingSystem,OperatingSystemServicePack,OperatingSystemVersion,LastLogonDate,description,dNSHostName
$AllComputersNames = $AllComputers.Name
$AllComputersNamesCount = [int]$AllComputersNames.Count
$swv.Stop()
write-verbose ("Section took: " + (FormatElapsedTime($swv.Elapsed)) + " to run.")
write-host "Starting Inventory"
Write-Progress  -Id 0 -Activity "Active Directory Computers Count: $AllComputersNamesCount" -Status "Starting Scan" -Percentcomplete (0)
Foreach ($ComputerName in $AllComputersNames) {
    #Write-Progress  -Id 0 -Activity ("Processing Computer") -Status ("( " + $count + "\" + $AllComputersNames.count + "): " + $CurrentComputer) -percentComplete ($count / $AllComputersNames.count*100)
    #Write-Host ("( " + $count + "\" + $Computers.count + ") Computer: " + $ComputerName)
    If ($null -ne $ComputerName) {
        $ComputerOS = $AllComputers | Where-object {$_.Name -eq $ComputerName}
        Write-Verbose "Starting job ($((Get-Job -name * | Measure-Object).Count+1)/$MaxJobs) for $ComputerName."
        Start-Job -Name  $ComputerName -ArgumentList $ComputerName,$ComputerOS,$RemoveServices,$RemoveSMB1,$InstallWMF,$LogHistory,$KillProcesses,$VMs,$winrmcommands,$SetupWinRM,$PSExecPath,$TimeOut -ScriptBlock{
            param
            (
                $ComputerName=$ComputerName,
                $ComputerOS=$ComputerOS,
                $RemoveServices=$RemoveServices,
                $RemoveSMB1=$RemoveSMB1,
                $InstallWMF=$InstallWMF,
                $LogHistory=$LogHistory,
                $KillProcesses=$KillProcesses,
                $VMs=$VMs,
                $winrmcommands=$winrmcommands,
                $SetupWinRM=$SetupWinRM,
                $PSExecPath=$PSExecPath,
                $TimeOut=$TimeOut
            )
            Class InventoryObject {
                [string]${AD Name}
                [string]${DNS IP}
                [string]$Enabled
                [string]$Description
                [string]${Operating System}
                [string]${Operating System Version}
                [string]${Operating System Service Pack}
                [string]${Last Logon Date}
                [string]${Computer IPs}
                [string]$DNS
                [string]${DNS Suffixs}
                [string]${VM Power State}
                [string]${VM Notes}
                [string]${VM Guest OS}
                [string]${VM Tools Status}
                [string]${VM Shutdown Date}
                [string]$Manufacturer
                [string]$Model
                [string]$Serial
                [string]${Number Of Processors}
                [string]${Processor Manufacturer}
                [string]${Processor Name}
                [string]${Number Of Cores}
                [string]${Number Of Logical Processors}
                [string]$RAM
                [string]${Disk Drive}
                [string]$Graphics  
                [string]${Graphics RAM}
                [string]${Sound Devices}
                [string]${Windows Features} 
                [string]${Windows Roles}
                [string]$Software 
                [string]$Services
                [string]$IIS 
                [string]${SSL Certificates Expiration}
                [string]${SQL Version} 
                [string]${SQL Databases} 
                [string]$FortiClient
                [string]${LANDesk Agent Installed}
                [string]${SMB Status}
                [string]${TLS Status} 
                [string]${Reboot Needed} 
                [string]${PowerShell Version}
                [string]${LAN Manager Authentication Level}
                [string]${Last Hotfixes Install Date}
                [string]${Last Hotfixes Packages}
                [string]${Distinguished Name}
            }
            function Get-CheckUrl{
                [CmdletBinding()]	
                param(
                    [parameter(Mandatory=$true)][string]$url,
                    [int]$timeoutMilliseconds = ($TimeOut * 1000),
                    [int]$MinimumCertAgeDays = 90
                )
                #source: https://stackoverflow.com/questions/39253055/powershell-script-to-get-certificate-expiry-for-a-website-remotely-for-multiple
                [string]$details = $null
                #Ignore Certs
                #https://blog.ukotic.net/2017/08/15/could-not-establish-trust-relationship-for-the-ssltls-invoke-webrequest/
                if (-not ([System.Management.Automation.PSTypeName]'ServerCertificateValidationCallback').Type)
                {
                $certCallback = @"
    using System;
    using System.Net;
    using System.Net.Security;
    using System.Security.Cryptography.X509Certificates;
    public class ServerCertificateValidationCallback
    {
        public static void Ignore()
        {
            if(ServicePointManager.ServerCertificateValidationCallback ==null)
            {
                ServicePointManager.ServerCertificateValidationCallback += 
                    delegate
                    (
                        Object obj, 
                        X509Certificate certificate, 
                        X509Chain chain, 
                        SslPolicyErrors errors
                    )
                    {
                        return true;
                    };
            }
        }
    }
"@
                Add-Type $certCallback
                }
                [ServerCertificateValidationCallback]::Ignore()
                #disabling the cert validation check. This is what makes this whole thing work with invalid certs...
                #[Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

                #Start Request
                $req = [Net.HttpWebRequest]::Create($url)
                # $req.ServerCertificateValidationCallback = [ServerCertificateValidationCallback]::Ignore()
                $req.Timeout = $timeoutMilliseconds
                $req.AllowAutoRedirect = $false
                try 
                {
                    # GET WEB RESPONSE
                    $req.GetResponse().Dispose() | Out-Null
                    if ($null -eq $req.ServicePoint.Certificate) {$details = "No certificate in use for connection"}
                } 
                catch 
                {
                    $details = "Exception while checking URL $url`: $_ "
                }
                if ( $null -eq $details -or $details -eq "")
                {
                    [datetime]$expiration = [System.DateTime]::Parse($req.ServicePoint.Certificate.GetExpirationDateString())
                    [int]$certExpiresIn = ($expiration - $(get-date)).Days
                    $certName = $req.ServicePoint.Certificate.GetName()
                    $certPublicKeyString = $req.ServicePoint.Certificate.GetPublicKeyString()
                    $certSerialNumber = $req.ServicePoint.Certificate.GetSerialNumberString()
                    $certThumbprint = $req.ServicePoint.Certificate.GetCertHashString()
                    $certEffectiveDate = $req.ServicePoint.Certificate.GetEffectiveDateString()
                    $certIssuer = $req.ServicePoint.Certificate.GetIssuerName()
                    if ($certExpiresIn -gt $minimumCertAgeDays) {
                        $returnData += new-object psobject -property  @{Url = $url; CheckResult = "OK"; CertExpiresInDays = [int]$certExpiresIn; ExpirationOn = [datetime]$expiration; CertName = $certname; Details = $details}
                    }else{
                        $details = ""
                        $details += "Cert for site $url expires in $certExpiresIn days [on $expiration]`n"
                        $details += "Threshold is $minimumCertAgeDays days. Check details:`n"
                        $details += "Cert name: $certName`n"
                        $details += "Cert public key: $certPublicKeyString`n"
                        $details += "Cert serial number: $certSerialNumber`n"
                        $details += "Cert thumbprint: $certThumbprint`n"
                        $details += "Cert effective date: $certEffectiveDate`n"
                        $details += "Cert issuer: $certIssuer"
                        $returnData += new-object psobject -property  @{Url = $url; CheckResult = "WARNING"; CertExpiresInDays = [int]$certExpiresIn; ExpirationOn = [datetime]$expiration; CertName = $certname; Details = $details}
                    }
                    Remove-Variable expiration
                    Remove-Variable certExpiresIn
                }else{
                    $returnData += new-object psobject -property  @{Url = $url; CheckResult = "ERROR"; CertExpiresInDays = $null; ExpirationOn = $null; CertName = $certname; Details = $details}
                }

                Remove-Variable req
                #Remove-Variable res
                return $returnData
            }
            Function Get-DomainComputer {
                <#
                .SYNOPSIS
                    The Get-DomainComputer function allows you to get information from an Active Directory Computer object using ADSI.
                
                .DESCRIPTION
                    The Get-DomainComputer function allows you to get information from an Active Directory Computer object using ADSI.
                    You can specify: how many result you want to see, which credentials to use and/or which domain to query.
                
                .PARAMETER ComputerName
                    Specifies the name(s) of the Computer(s) to query
                
                .PARAMETER SizeLimit
                    Specifies the number of objects to output. Default is 100.
                
                .PARAMETER DomainDN
                    Specifies the path of the Domain to query.
                    Examples:     "FX.LAB"
                                "DC=FX,DC=LAB"
                                "Ldap://FX.LAB"
                                "Ldap://DC=FX,DC=LAB"
                .PARAMETER Properties
                    Specifies extra fields to return. 
                .PARAMETER Credential
                    Specifies the alternate credentials to use.
                
                .EXAMPLE
                    Get-DomainComputer
                
                    This will show all the computers in the current domain
                
                .EXAMPLE
                    Get-DomainComputer -ComputerName "Workstation001"
                
                    This will query information for the computer Workstation001.
                
                .EXAMPLE
                    Get-DomainComputer -ComputerName "Workstation001","Workstation002"
                
                    This will query information for the computers Workstation001 and Workstation002.
                
                .EXAMPLE
                    Get-Content -Path c:\WorkstationsList.txt | Get-DomainComputer
                
                    This will query information for all the workstations listed inside the WorkstationsList.txt file.
                
                .EXAMPLE
                    Get-DomainComputer -ComputerName "Workstation0*" -SizeLimit 10 -Verbose
                
                    This will query information for computers starting with 'Workstation0', but only show 10 results max.
                    The Verbose parameter allow you to track the progression of the script.
                
                .EXAMPLE
                    Get-DomainComputer -ComputerName "Workstation0*" -SizeLimit 10 -Verbose -DomainDN "DC=FX,DC=LAB" -Credential (Get-Credential -Credential FX\Administrator)
                
                    This will query information for computers starting with 'Workstation0' from the domain FX.LAB with the account FX\Administrator.
                    Only show 10 results max and the Verbose parameter allows you to track the progression of the script.
                
                .NOTES
                    NAME:    FUNCT-AD-COMPUTER-Get-DomainComputer.ps1
                    AUTHOR:    Francois-Xavier CAT
                    DATE:    2013/10/26
                    WWW:    www.lazywinadmin.com
                    TWITTER: @lazywinadmin
                
                    VERSION HISTORY:
                    1.0 2013.10.26
                        Initial Version
                    1.1 2021.06.21
                        Added ability to have extra properties. 
                #>
                
                    [CmdletBinding()]
                    PARAM(
                        [Parameter(
                            ValueFromPipelineByPropertyName=$true,
                            ValueFromPipeline=$true)]
                        [Alias("Computer")]
                        [String[]]$ComputerName,
                
                        [Alias("ResultLimit","Limit")]
                        [int]$SizeLimit='100',
                
                        [Parameter(ValueFromPipelineByPropertyName=$true)]
                        [Alias("Domain")]
                        [String]$DomainDN=$(([adsisearcher]"").Searchroot.path),

                        [Parameter(ValueFromPipelineByPropertyName=$true)]
                        [Alias("Property")]
                        [String[]]$Properties,

                        [Alias("RunAs")]
                        [System.Management.Automation.Credential()]
                        $Credential = [System.Management.Automation.PSCredential]::Empty
                
                    )#PARAM
                
                    PROCESS{
                        IF ($ComputerName){
                            Write-Verbose -Message "One or more ComputerName specified"
                            FOREACH ($item in $ComputerName){
                                TRY{
                                    # Building the basic search object with some parameters
                                    Write-Verbose -Message "COMPUTERNAME: $item"
                                    $Searcher = New-Object -TypeName System.DirectoryServices.DirectorySearcher -ErrorAction 'Stop' -ErrorVariable ErrProcessNewObjectSearcher
                                    $Searcher.Filter = "(&(objectCategory=Computer)(name=$item))"
                                    $Searcher.SizeLimit = $SizeLimit
                                    $Searcher.SearchRoot = $DomainDN
                
                                    #Specify Other Properties to load
                                    # If ($Properties) {
                                    #     foreach ($Property in $Properties) {
                                    #         [void]$Searcher.PropertiesToLoad.Add($Property)
                                    #     }
                                    # }
                                    # Specify a different domain to query
                                    IF ($PSBoundParameters['DomainDN']){
                                        IF ($DomainDN -notlike "LDAP://*") {$DomainDN = "LDAP://$DomainDN"}#IF
                                        Write-Verbose -Message "Different Domain specified: $DomainDN"
                                        $Searcher.SearchRoot = $DomainDN}#IF ($PSBoundParameters['DomainDN'])
                
                                    # Alternate Credentials
                                    IF ($PSBoundParameters['Credential']) {
                                        Write-Verbose -Message "Different Credential specified: $($Credential.UserName)"
                                        $Domain = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList $DomainDN,$($Credential.UserName),$($Credential.GetNetworkCredential().password) -ErrorAction 'Stop' -ErrorVariable ErrProcessNewObjectCred
                                        $Searcher.SearchRoot = $Domain}#IF ($PSBoundParameters['Credential'])
                
                                    # Querying the Active Directory
                                    Write-Verbose -Message "Starting the ADSI Search..."
                                    FOREACH ($Computer in $($Searcher.FindAll())){
                                        Write-Verbose -Message "$($Computer.properties.name)"
                                        If ($Properties) {
                                            $Output = New-Object -TypeName PSObject -ErrorAction 'Continue' -ErrorVariable ErrProcessNewObjectOutputALL -Property @{}
                                            foreach ($Property in $Properties) {
                                                Try {
                                                    Write-Verbose -Message $([string]$Computer.properties.$Property + " ")
                                                    switch ($Property.ToLower()) {
                                                        "lastlogondate" { 
                                                            Add-Member -InputObject $Output -NotePropertyName ($Property) -NotePropertyValue $([datetime]::FromFileTime([int64][string]$computer.Properties.lastlogon).ToString('MM/dd/yyy hh:mm:ss tt')) -Force
                                                         }
                                                         "enabled" {
                                                             If ([int64][string]$computer.Properties.useraccountcontrol -band 2 -eq 0) {
                                                                Add-Member -InputObject $Output -NotePropertyName ($Property) -NotePropertyValue $true -Force
                                                             }Else {
                                                                Add-Member -InputObject $Output -NotePropertyName ($Property) -NotePropertyValue $false -Force
                                                             }
                                                         }
                                                         "operatingsystemservicepack" {
                                                             #Need to find/figure out
                                                            Add-Member -InputObject $Output -NotePropertyName ($Property) -NotePropertyValue $("") -Force
                                                         }
                                                        Default {
                                                            Add-Member -InputObject $Output -NotePropertyName ($Property) -NotePropertyValue $([string]$Computer.properties.($Property.ToLower())) -Force
                                                        }
                                                    }
                                                   
                                                }
                                                Catch {
                                                    Write-Warning -Message ('{0}: {1}' -f $Property, $_.Exception.Message)
                                                }   
                                            }
                                        }Else {
                                            $Output = New-Object -TypeName PSObject -ErrorAction 'Continue' -ErrorVariable ErrProcessNewObjectOutputALL -Property @{
                                                "Name" = $($Computer.properties.name)
                                                "DNShostName"    = $($Computer.properties.dnshostname)
                                                "Description" = $($Computer.properties.description)
                                                "OperatingSystem"=$($Computer.Properties.operatingsystem)
                                                "WhenCreated" = $($Computer.properties.whencreated)
                                                "DistinguishedName" = $($Computer.properties.distinguishedname)}#New-Object
                                        } #If $Properties
                                        return $Output
                                    }#FOREACH $Computer
                
                                    Write-Verbose -Message "ADSI Search completed"
                                }#TRY
                                CATCH{
                                    Write-Warning -Message ('{0}: {1}' -f $item, $_.Exception.Message)
                                    IF ($ErrProcessNewObjectSearcher){Write-Warning -Message "PROCESS BLOCK - Error during the creation of the searcher object"}
                                    IF ($ErrProcessNewObjectCred){Write-Warning -Message "PROCESS BLOCK - Error during the creation of the alternate credential object"}
                                    IF ($ErrProcessNewObjectOutput){Write-Warning -Message "PROCESS BLOCK - Error during the creation of the output object"}
                                }#CATCH
                            }#FOREACH $item
                
                
                        }#IF $ComputerName
                        ELSE {
                            Write-Verbose -Message "No ComputerName specified"
                            TRY{
                                # Building the basic search object with some parameters
                                Write-Verbose -Message "List All object"
                                $Searcher = New-Object -TypeName System.DirectoryServices.DirectorySearcher -ErrorAction 'Stop' -ErrorVariable ErrProcessNewObjectSearcherALL
                                $Searcher.Filter = "(objectCategory=Computer)"
                                $Searcher.SizeLimit = $SizeLimit
                                
                                #Specify Other Properties to load
                                # If ($Properties) {
                                #     foreach ($Property in $Properties) {
                                #         [void]$Searcher.PropertiesToLoad.Add($Property)
                                #     }
                                # }                
                                # Specify a different domain to query
                                IF ($PSBoundParameters['DomainDN']){
                                    $DomainDN = "LDAP://$DomainDN"
                                    Write-Verbose -Message "Different Domain specified: $DomainDN"
                                    $Searcher.SearchRoot = $DomainDN}#IF ($PSBoundParameters['DomainDN'])
                
                                # Alternate Credentials
                                IF ($PSBoundParameters['Credential']) {
                                    Write-Verbose -Message "Different Credential specified: $($Credential.UserName)"
                                    $DomainDN = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList $DomainDN, $Credential.UserName,$Credential.GetNetworkCredential().password -ErrorAction 'Stop' -ErrorVariable ErrProcessNewObjectCredALL
                                    $Searcher.SearchRoot = $DomainDN}#IF ($PSBoundParameters['Credential'])
                
                                # Querying the Active Directory
                                Write-Verbose -Message "Starting the ADSI Search..."
                                FOREACH ($Computer in $($Searcher.FindAll())){
                                    TRY{
                                        Write-Verbose -Message "$($Computer.properties.name)"
                                        If ($Properties) {
                                            Try {
                                                Write-Verbose -Message $([string]$Computer.properties.$Property + " ")
                                                switch ($Property.ToLower()) {
                                                    "lastlogondate" { 
                                                        Add-Member -InputObject $Output -NotePropertyName ($Property) -NotePropertyValue $([datetime]::FromFileTime([int64][string]$computer.Properties.lastlogon).ToString('MM/dd/yyy hh:mm:ss tt')) -Force
                                                     }
                                                     "enabled" {
                                                         If ([int64][string]$computer.Properties.useraccountcontrol -band 2 -eq 0) {
                                                            Add-Member -InputObject $Output -NotePropertyName ($Property) -NotePropertyValue $true -Force
                                                         }Else {
                                                            Add-Member -InputObject $Output -NotePropertyName ($Property) -NotePropertyValue $false -Force
                                                         }
                                                     }
                                                     "operatingsystemservicepack" {
                                                         #Need to find/figure out
                                                        Add-Member -InputObject $Output -NotePropertyName ($Property) -NotePropertyValue $("") -Force
                                                     }
                                                    Default {
                                                        Add-Member -InputObject $Output -NotePropertyName ($Property) -NotePropertyValue $([string]$Computer.properties.($Property.ToLower())) -Force
                                                    }
                                                }
                                               
                                            }
                                            Catch {
                                                Write-Warning -Message ('{0}: {1}' -f $Property, $_.Exception.Message)
                                            }   
                                        }Else {
                                            $Output = New-Object -TypeName PSObject -ErrorAction 'Continue' -ErrorVariable ErrProcessNewObjectOutputALL -Property @{
                                                "Name" = $($Computer.properties.name)
                                                "DNShostName"    = $($Computer.properties.dnshostname)
                                                "Description" = $($Computer.properties.description)
                                                "OperatingSystem"=$($Computer.Properties.operatingsystem)
                                                "WhenCreated" = $($Computer.properties.whencreated)
                                                "DistinguishedName" = $($Computer.properties.distinguishedname)}#New-Object
                                        } #If $Properties
                                        return $Output
                                    }#TRY
                                    CATCH{
                                        Write-Warning -Message ('{0}: {1}' -f $Computer, $_.Exception.Message)
                                        IF ($ErrProcessNewObjectOutputALL){Write-Warning -Message "PROCESS BLOCK - Error during the creation of the output object"}
                                    }
                                }#FOREACH $Computer
                
                                Write-Verbose -Message "ADSI Search completed"
                
                            }#TRY
                
                            CATCH{
                                Write-Warning -Message "Something Wrong happened"
                                IF ($ErrProcessNewObjectSearcherALL){Write-Warning -Message "PROCESS BLOCK - Error during the creation of the searcher object"}
                                IF ($ErrProcessNewObjectCredALL){Write-Warning -Message "PROCESS BLOCK - Error during the creation of the alternate credential object"}
                
                            }#CATCH
                        }#ELSE
                    }#PROCESS
                    END{Write-Verbose -Message "Function Completed"}
            }#function
            Function Start-PSExec {
                param(
                    [Parameter(Mandatory=$true)][string]$Computer,
                    [Parameter(Mandatory=$true)][string]$Command,
                    [Parameter(Mandatory=$true)][string]$PSExecPath,
                    [Parameter(Mandatory=$false)]$maximumRuntimeSeconds = $TimeOut,
                    [Parameter(Mandatory=$false)][string]$User,
                    [Parameter(Mandatory=$false)][string]$Pass, 
                    [Parameter(Mandatory=$false)][switch]$Copy
            
                    )
            
                If ($Command) {
                    Write-Verbose ("`t`t[" + $Computer +  "] Running program: " + $Command)
                    if ( $User -and $Pass) {
                        If ($Copy) {
                            $process = Start-Process -FilePath $PSExecPath -ArgumentList $("\\" + $Computer + " -h -c -v -i -accepteula -nobanner -u " + $User + " -p " + $Pass + " " + $Command) -PassThru -NoNewWindow
                        } else {
                            $process = Start-Process -FilePath $PSExecPath -ArgumentList $("\\" + $Computer + " -h -i -accepteula -nobanner -u " + $User + " -p " + $Pass + " " + $Command) -PassThru -NoNewWindow
                        }
                    }else{
                        If ($Copy) {
                            $process = Start-Process -FilePath $PSExecPath -ArgumentList $("\\" + $Computer + " -h -c -v -i -accepteula -nobanner " + $Command) -PassThru -NoNewWindow
                        } else {
                            $process = Start-Process -FilePath $PSExecPath -ArgumentList $("\\" + $Computer + " -h -i -accepteula -nobanner " + $Command) -PassThru -NoNewWindow
                        }
                    }
                    try 
                    {
                        $process | Wait-Process -Timeout $maximumRuntimeSeconds -ErrorAction Stop 
                        If ($process.ExitCode -le 0) {
                            Write-Verbose ("`t`t`t PSExec successfully completed within timeout.")
                            If ($ALLCSV) {
                                If (Test-Path -Path ($LogFile + "_all.csv")) {
                                    #"Date,Computer,Command,Status"
                                    Add-Content ($LogFile + "_all.csv") ((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + "," + $Command + ",success")
                                }
                            }
                        }else{
                            Write-Warning -Message $('PSExec could not run command. Exit Code: ' + $process.ExitCode)
                            If ($ALLCSV) {
                                If (Test-Path -Path ($LogFile + "_all.csv")) {
                                    #"Date,Computer,Command,Status"
                                    Add-Content ($LogFile + "_all.csv") ((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + "," + $Command + ",Failed Error:" + $process.ExitCode)
                                }
                            }
                            continue
                        }
                    }catch{
                        Write-Warning -Message 'PSExec exceeded timeout, will be killed now.' 
                        If ($ALLCSV) {
                            If (Test-Path -Path ($LogFile + "_all.csv")) {
                                #"Date,Computer,Command,Status"
                                Add-Content ($LogFile + "_all.csv") ((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + "," + $Command + ",Timed Out")
                            }
                        }
                        $process | Stop-Process -Force
                        continue
                    } 
                }else{
                    Write-Warning -Message "`t`t NO Commands"
                }
                
            }

            #Create New Object
            $ComputerInfo = [InventoryObject]::new()

            If ($null -eq $ComputerOS) {
                $ComputerOS = Get-DomainComputer -ComputerName "$ComputerName" -Properties Name,Enabled,OperatingSystem,OperatingSystemServicePack,OperatingSystemVersion,LastLogonDate,description,dNSHostName -ErrorAction SilentlyContinue
            }
            
            if ($null -ne $ComputerOS.Name) {
                $ComputerInfo.'AD Name'= $ComputerOS.Name
            }else {
                $ComputerInfo.'AD Name'= $ComputerName
            }
            If ($null -ne $ComputerOS.dNSHostName) {
                $ComputerInfo.'DNS IP' = (Resolve-DnsName -Name $ComputerOS.dNSHostName -ErrorAction SilentlyContinue).IPAddress
            }Else {
                $ComputerInfo.'DNS IP' = (Resolve-DnsName -Name $ComputerName -ErrorAction SilentlyContinue).IPAddress
            }
            If ($null -ne $ComputerOS) {
                $ComputerInfo.Enabled = $ComputerOS.Enabled
                $ComputerInfo.Description = $ComputerOS.description
                $ComputerInfo.'Operating System' = $ComputerOS.OperatingSystem
                $ComputerInfo.'Operating System Version' = $ComputerOS.OperatingSystemVersion
                $ComputerInfo.'Operating System Service Pack' = $ComputerOS.OperatingSystemServicePack
                $ComputerInfo.'Last Logon Date' = $ComputerOS.LastLogonDate
                $ComputerInfo.'Distinguished Name' = $ComputerOS.DistinguishedName
            }
            $CurrentVM = ($VMs.ToArray())[0] | Where-object {$_.Name -eq $ComputerName}
            If ($CurrentVM) {
                $ComputerInfo.'VM Power State' = $CurrentVM.'Power State'
                $ComputerInfo.'VM Notes' = $CurrentVM.'Notes'
                $ComputerInfo.'VM Guest OS' = $CurrentVM.'Guest OS'
                $ComputerInfo.'VM Tools Status' = $CurrentVM.'Tools Status'
                $ComputerInfo.'VM Shutdown Date' = $CurrentVM.'Shutdown Date'
            }
            If ($ComputerOS.Enabled -eq $true -and $null -ne $ComputerInfo.'DNS IP') {
                #Create remote connections
                Try{
                    $cimSession = New-CimSession -ComputerName $ComputerName -ErrorAction SilentlyContinue
                }Catch{
                    if (-Not $cimSession) {
                        Try{   
                            $cimSession = New-CimSession -ComputerName $ComputerName -SkipTestConnection -ErrorAction SilentlyContinue 
                        }Catch{
                            Write-Warning ($ComputerName + ": Error getting WMI Servcie")
                        }
                    }
                }
                                 
                $psSessionoptions = New-PSSessionOption -SkipCAChec -SkipCNCheck -SkipRevocationCheck
                Try{
                $psSession =  New-PSSession -UseSSL -Authentication Kerberos -ComputerName $ComputerName -ErrorAction SilentlyContinue -SessionOption $psSessionoptions
                }Catch{
                    if (-Not $psSession -or $psSession.Availability -ne "Available" ) {
                        $psSession =  New-PSSession -ComputerName $ComputerName -ErrorAction SilentlyContinue 
                    }
                }
                #Trying Setting up WINRM
                if ((-Not $psSession -or $psSession.Availability -ne "Available" ) -and $SetupWinRM) {                   
                    ForEach ($Command in $winrmcommands) {
                        Start-PSExec -Computer $ComputerName -Command $Command -PSExecPath $PSExecPath -maximumRuntimeSeconds 120
                    }
                    Try{
                        $psSession =  New-PSSession -UseSSL -Authentication Kerberos -ComputerName $ComputerName -ErrorAction SilentlyContinue -SessionOption $psSessionoptions
                    }Catch{
                        if (-Not $psSession -or $psSession.Availability -ne "Available" ) {
                            $psSession =  New-PSSession -ComputerName $ComputerName -ErrorAction SilentlyContinue 
                        } 
                    }
                }
                If ($cimSession -and $psSession) {
                    Write-verbose ("`tHost is up")
                    #region NIC
                    Try {
                        $ArrComputerIP= (Get-CimInstance -Class Win32_NetworkAdapterConfiguration -CimSession $cimSession -Filter 'IPEnabled = True' -ErrorAction SilentlyContinue | Select-Object IPAddress,Description,DNSServerSearchOrder,DNSDomainSuffixSearchOrder)
                        If ($ArrComputerIP) {                       
                            $ComputerIP=@()
                            $ComputerDNS=@()
                            $ComputerDNSSuffix=@()
                            
                            foreach ($NIC in $ArrComputerIP) {
                                $cIP = @()
                                foreach ($iIP in $NIC.IPAddress) {
                                    If (([ipaddress]$iIP).AddressFamily -eq "InterNetwork" -and $iIP -ne "127.0.0.1" -and $iIP -notlike "169.254.*") {
                                        $cIP += $iIP
                                    }
                                }
                                If ($null -eq $cIP){
                                    $ComputerIP += (($NIC.Description -replace '[,()]',''))
                                }Else{
                                    $ComputerIP += (($NIC.Description -replace '[,()]','') + " (" + ($cIP -join " ") + ")")
                                }
                                if ($NIC.DNSServerSearchOrder) {
                                    $ComputerDNS += $NIC.DNSServerSearchOrder
                                }
                                if ($NIC.DNSDomainSuffixSearchOrder) {
                                    $ComputerDNSSuffix += $NIC.DNSDomainSuffixSearchOrder
                                }
                            }
                            If($Excel) {
                                $ComputerInfo.'Computer IPs' =  ($ComputerIP | Sort-Object -Unique) -join "\n"
                                $ComputerInfo.DNS =  ($ComputerDNS | Sort-Object -Unique) -join "\n"
                                $ComputerInfo.'DNS Suffixs' =  ($ComputerDNSSuffix | Sort-Object -Unique) -join "\n"
                            }Else{
                                $ComputerInfo.'Computer IPs' =  ($ComputerIP | Sort-Object -Unique) -join ","
                                $ComputerInfo.DNS =  ($ComputerDNS | Sort-Object -Unique) -join ","
                                $ComputerInfo.'DNS Suffixs' =  ($ComputerDNSSuffix | Sort-Object -Unique) -join ","
                            }

                        }
                    }Catch{
                        Write-Warning ($ComputerName + ": Error getting NIC Information")
                    }
                    #endregion NIC
                    Try{
                        $ComputerHW = Get-CimInstance -Class Win32_ComputerSystem -CimSession $cimSession | Select-Object Manufacturer,Model,NumberOfProcessors,@{Expression={($_.TotalPhysicalMemory / 1GB).ToString("#,###.##")};Label="TotalPhysicalMemoryGB"}
                        if($ComputerHW){
                            $ComputerInfo.Manufacturer =  $ComputerHW.Manufacturer
                            $ComputerInfo.Model = $ComputerHW.Model
                            $ComputerInfo.'Number Of Processors' = $ComputerHW.NumberOfProcessors
                            $ComputerInfo.RAM = $ComputerHW.TotalPhysicalMemoryGB
                        }
                    }Catch{
                        Write-Warning ($ComputerName + ": Error getting Hardware Information")
                    }
                    Try{
                        $ComputerCPU = Get-CimInstance win32_processor -CimSession $cimSession | Select-Object DeviceID,Name,Manufacturer,NumberOfCores,NumberOfLogicalProcessors
                        if($ComputerCPU){
                            $ComputerInfo.'Processor Manufacturer' = $ComputerCPU.Manufacturer | Select-Object -Unique -First 1
                            $ComputerInfo.'Processor Name' = $ComputerCPU.Name | Select-Object -Unique -First 1
                            $sum = 0
                            $ComputerCPU.NumberOfCores  | ForEach-Object { $sum += $_}
                            $ComputerInfo.'Number Of Cores' = $sum
                            $sum = 0
                            $ComputerCPU.NumberOfLogicalProcessors | ForEach-Object { $sum += $_}
                            $ComputerInfo.'Number Of Logical Processors' = $sum
                        }
                    }Catch{
                        Write-Warning ($ComputerName + ": Error getting CPU Information")
                    }
                    #region Disk
                    $ArrComputerDisks = Get-CimInstance -Class Win32_LogicalDisk -Filter "DriveType=3" -CimSession $cimSession |
                        Select-Object DeviceID,VolumeName,@{Expression={($_.Size / 1GB).ToString("###.##")};Label="SizeGB"}
                        If (($ArrComputerDisks | Measure-Object).count -gt 1) {
                            If (-Not ([string]::IsNullOrWhiteSpace($_.VolumeName))) {
                                If($Excel) {
                                    $ComputerDisks = ($ArrComputerDisks | ForEach-Object {($_.DeviceID + " (" + $_.VolumeName + ") " + $_.SizeGB  + " GB")})  -join "\n"
                                }Else{
                                    $ComputerDisks = ($ArrComputerDisks | ForEach-Object {($_.DeviceID + " (" + $_.VolumeName + ") " + $_.SizeGB  + " GB")})  -join ","
                                }
                            }else{
                                If($Excel) {
                                    $ComputerDisks = ($ArrComputerDisks | ForEach-Object {($_.DeviceID + " " + $_.SizeGB + " GB")})  -join "\n"
                                }Else{
                                    $ComputerDisks = ($ArrComputerDisks | ForEach-Object {($_.DeviceID + " " + $_.SizeGB  + " GB")})  -join ","
                                }
                            }
                        } else {
                            If (-Not ([string]::IsNullOrWhiteSpace($ArrComputerDisks.VolumeName))) {
                                $ComputerDisks = ( $ArrComputerDisks.DeviceID + "  (" + $ArrComputerDisks.VolumeName + ") " +  $ArrComputerDisks.SizeGB  + " GB" )
                            }else {
                                $ComputerDisks = ( $ArrComputerDisks.DeviceID + " " + $ArrComputerDisks.SizeGB  + " GB" )
                            }
                        }   
                    $ComputerInfo.'Disk Drive' = $ComputerDisks
                    #endregion Disk
                    Try{
                        $ComputerInfo.Serial = (Get-CimInstance Win32_Bios -CimSession $cimSession).SerialNumber
                    }Catch{
                        Write-Warning ($ComputerName + ": Error getting Serial Number")
                    }
                    Try{
                        $ComputerGraphics = Get-CimInstance -Class Win32_VideoController -CimSession $cimSession |Select-Object Name,@{Expression={($_.AdapterRAM / 1MB).ToString("#,###.##")};Label="GraphicsRAM"}
                        if($ComputerGraphics){
                            $ComputerInfo.Graphics = $ComputerGraphics.Name
                            $ComputerInfo.'Graphics RAM' = $ComputerGraphics.GraphicsRAM
                        }
                    }Catch{
                        Write-Warning ($ComputerName + ": Error getting Graphics Information")
                    }
                    Try{
                        $ComputerInfo.'Sound Devices' = ((Get-CimInstance -Class Win32_SoundDevice -CimSession $cimSession).Name  -join ", ") 
                    }Catch{
                        Write-Warning ($ComputerName + ": Error getting Sound Information")
                    }

                    #region Windows Optional Features
                    Try{
                        If($Excel) {
                            $ComputerInfo.'Windows Features' =  ((Get-CimInstance -CimSession $cimSession -query "select Name from win32_optionalfeature where installstate= 1").name  -join "\n") 
                        }else{
                            $ComputerInfo.'Windows Features' =  ((Get-CimInstance -CimSession $cimSession -query "select Name from win32_optionalfeature where installstate= 1").name  -join ", ") 
                        }
                        If (($ComputerOS.OperatingSystem).Contains("Server")) {
                            $ComputerRolesTemp = Invoke-Command -Session $psSession -ScriptBlock { 
                                If (Get-Module -ListAvailable -Name "servermanager") {
                                    If (-Not (Get-Module "servermanager" -ErrorAction SilentlyContinue)) {
                                        Import-Module "servermanager"
                                    }
                                    get-windowsfeature | Where-Object { $_.installed -eq $true -and $_.featuretype -eq 'Role'} | Select-Object name
                                }
                            }
                            If($ComputerRolesTemp){
                                If($Excel) {
                                    $ComputerInfo.'Windows Roles' =  $ComputerRolesTemp.Name -join "\n"
                                }Else{
                                    $ComputerInfo.'Windows Roles' =  (($ComputerRolesTemp.Name) -join ",") 
                                }
                            }
                        }
                    }Catch{
                        Write-Warning ($ComputerName + ": Error getting Windows Optional Features")
                    }
                    #endregion Windows Optional Features
                    #region Software
                    Try{
                        $ArrComputerSoftware = Get-CimInstance -Class Win32_Product -CimSession $cimSession | Select-Object Name,Version,InstallDate | Sort-Object -Unique -Property Name,Version,InstallDate
                        If ($ArrComputerSoftware){
                            If (($ArrComputerSoftware | Measure-Object).count -gt 1) {
                                If($Excel) {
                                    $ComputerSoftware =  ($ArrComputerSoftware | ForEach-Object {($_.Name + " (" + $_.Version + ") - " + $_.InstallDate)}) -join "\n"
                                }Else{
                                    $ComputerSoftware =  (($ArrComputerSoftware | ForEach-Object {($_.Name + " (" + $_.Version + ") - " + $_.InstallDate)}) -join ",")
                                    }
                            } else {
                                $ComputerSoftware =  ($ArrComputerSoftware.Name + " - " + $ArrComputerSoftware.Version) 
                            }
                            $ComputerInfo.Software = $ComputerSoftware
                        }
                    }Catch{
                        Write-Warning ($ComputerName + ": Error getting Install Software")
                    }
                    #endregion Software
                    #region Services
                    Try{
                        $ArrComputerServcies = Get-CimInstance -Class Win32_Service -CimSession $cimSession | Select-Object Name,Caption,StartMode,PathName | Sort-Object -Unique -Property Caption
                        If($ArrComputerServcies){
                            $ComputerServcies=@()
                            ForEach ($ItemServcie in $ArrComputerServcies) {
                                If (($ItemServcie.Caption + " - " + $ItemServcie.StartMode) -notin $RemoveServices ) {
                                    $ComputerServcies += ($ItemServcie.Caption + " - " + $ItemServcie.StartMode) 
                                }
                            }

                            If($Excel) {
                                $ComputerServcies = $ComputerServcies -join "\n"
                            }else{
                                $ComputerServcies = $ComputerServcies -join ","
                            }
                            $ComputerInfo.Services = $ComputerServcies
                        }
                    }Catch{
                        Write-Warning ($ComputerName + ": Error getting Servcies")
                    }
                    #endregion Services                   
                    #region IIS
                    Try{
                        IF ($ArrComputerServcies | Where-Object {$_.Name -eq "W3SVC"}) {
                            $ComputerIIS = Invoke-Command -errorAction SilentlyContinue -HideComputerName -Session $psSession -ScriptBlock {                           
                                $MaxMemoryUseage = 75
                                #Force Process to run with lower priority
                                Get-Process -id $pid | ForEach-Object {$_.PriorityClass='BelowNormal'}

                                function ConvertFrom-IISW3CLog {
                                    [CmdletBinding()]
                                    param (
                                        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
                                        [Alias('PSPath')]
                                        [string[]]
                                        $Path,
                                        [int]$MaxMemoryUseage = 75,
                                        [string[]]$FilterField,
                                        [string[]]$KillProcesses
                                    )
                                    #Source: https://gist.github.com/jstangroome/6189660
                                    process {
                                        If ($KillProcesses) {
                                            Stop-Process -Force -Name $KillProcesses -ErrorAction SilentlyContinue
                                        }
                                        If ($null -eq $MaxMemoryUseage -or $MaxMemoryUseage -le 0) {
                                            $MaxMemoryUseage = 75
                                        }
                                        foreach ($SinglePath in $Path) {
                            
                                            $FieldNames = $null
                                            $Properties = @{}
                                            
                                            #Test Memory Usage
                                            If (Get-Command Get-CimInstance -errorAction SilentlyContinue) {
                                                $OperatingSystem = Get-CimInstance -Class win32_OperatingSystem
                                            }Else{
                                                $OperatingSystem = Get-WmiObject -Class win32_OperatingSystem
                                            }
                                            # Lets grab the free memory
                                            $FreeMemory = $OperatingSystem.FreePhysicalMemory
                                            # Lets grab the total memory
                                            $TotalMemory = $OperatingSystem.TotalVisibleMemorySize
                                            # Calculate used memory
                                            $MemoryUsed = ($TotalMemory-$FreeMemory)
                                            # Lets do some math for percent
                                            $PercentMemoryUsed = "{0:N2}" -f (($MemoryUsed / $TotalMemory) * 100)
                                            
                                            If ( $PercentMemoryUsed -le $MaxMemoryUseage ) {
                                                Get-Content -Path $SinglePath |
                                                    ForEach-Object {
                                                        if ($_ -match '^#') {
                                                            #metadata
                                                            if ($_ -match '^#(?<k>[^:]+):\s*(?<v>.*)$') {
                                                                #key value pair
                                                                if ($Matches.k -eq 'Fields') {
                                                                    $FieldNames  = @(-split $Matches.v)
                                                                }
                                                            }
                                                        } else {
                                                            $FieldValues = @(-split $_)
                                                            $Properties.Clear()
                                                            for ($Index = 0; $Index -lt $FieldValues.Length; $Index++) {
                                                                If( $null -eq $FieldValues[$Index]) {
                                                                    If ($FilterField) {
                                                                        If ($FieldNames -in $FilterField) {
                                                                            $Properties[$FieldNames[$Index]] = ""
                                                                        }
                                                                    }Else{
                                                                        $Properties[$FieldNames[$Index]] = ""
                                                                    }
                                                                }else{
                                                                    If ($null -ne $FieldNames[$Index] ) {
                                                                        If ($FilterField) {
                                                                            If ($FieldNames -in $FilterField) {
                                                                                $Properties[$FieldNames[$Index]] = $FieldValues[$Index]
                                                                            }
                                                                        }Else{
                                                                            $Properties[$FieldNames[$Index]] = $FieldValues[$Index]
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            [pscustomobject]$Properties
                                                        }
                                                    }
                                            }Else{
                                                Write-Error ("Skipped file " + (split-path -path $SinglePath -Leaf) + " on computer " + $env:computername + " Memory usage at " + $PercentMemoryUsed + "%.")
                                                If ($KillProcesses) {
                                                    Stop-Process -Force -Name $KillProcesses -ErrorAction SilentlyContinue
                                                }
                                            }
                                        }
                            
                                    }
                                } 
                                If (Get-Module -ListAvailable -Name "WebAdministration") {
                                    If (-Not (Get-Module "WebAdministration" -ErrorAction SilentlyContinue)) {
                                        Import-Module "WebAdministration" -DisableNameChecking
                                    } 
                                }	
                                If($using:LogHistory -gt 0 -or $null -eq $using:LogHistory ) {
                                    #Do not get Site hits
                                    If (get-psdrive | Where-Object {$_.Name -eq "IIS"}) {
                                        Get-ChildItem "IIS:\Sites" | Where-Object -Property State -eq "Started" | ForEach-Object {

                                            Write-Output ($_.Name + " (" + ([string]$_.Bindings.Collection -join "; ") + "),")
                                            $apps = $_ | Get-ChildItem | Where-Object { $_.nodetype -eq "application" }
                                            Foreach ($app  in $apps) {
                                                Write-Output ("`t" + $app.Name )
                                            }
                                        }
                                    }   
                                }Else{
                                    #Get Site hits
                                    If (get-psdrive | Where-Object {$_.Name -eq "IIS"}) {
                                        Get-ChildItem "IIS:\Sites" | Where-Object -Property State -eq "Started" | ForEach-Object {
                                            #Get number of hits in the last $LogHistory days
                                            If ( Test-Path -Path ([System.Environment]::ExpandEnvironmentVariables($_.logfile.directory) + "\W3SVC" + $_.id)) {
                                                $Logs = Get-ChildItem ([System.Environment]::ExpandEnvironmentVariables($_.logfile.directory) + "\W3SVC" + $_.id) -include *.log -rec | Where-Object {$_.LastWriteTime -gt (Get-Date).AddDays($using:LogHistory)} | Sort-Object -Property LastWriteTime -Descending
                                                $Count = 0
                                                If ($null -ne $logs) {
                                                    foreach ($log in $Logs) {
                                                        $Count += ($log | ConvertFrom-IISW3CLog -FilterField "c-ip" -MaxMemoryUseage $MaxMemoryUseage -KillProcesses $KillProcesses | Where-Object {$null -ne $_."c-ip" }| Measure-Object).Count
                                                    }
                                                }
                                                Write-Output ($_.Name + " (" + ([string]$_.Bindings.Collection -join ";") + ") Hits in " + $using:LogHistory + " : " +  $Count)
                                                $apps = $_ | Get-ChildItem | Where-Object { $_.nodetype -eq "application" }
                                                Foreach ($app  in $apps) {
                                                    If($app.Name) {
                                                        Write-Output ("`t" + $app.Name )
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            #results
                            If($Excel) {
                                $ComputerIIS = $ComputerIIS -join "\n"
                            }else{
                                $ComputerIIS = $ComputerIIS -join ", "
                            }
                            #output
                            $ComputerInfo.IIS = $ComputerIIS 
                        }
                    }Catch{
                        Write-Warning ($ComputerName + ": Error getting IIS Information")
                    }
                    #endregion IIS
                    #region SSL Certs
                    Try{
                        $NetConnections = Get-NetTCPConnection -CimSession $ComputerName -State Listen -ErrorAction SilentlyContinue | Where-Object {$_.RemoteAddress -ne "::1" -and $_.RemoteAddress -ne "127.0.0.1" -and $_.LocalAddress -ne "127.0.0.1" -and $_.LocalAddress -ne "::1" }
                        If($NetConnections){
                            $WebCerts = @()
                            Foreach ($NetConnection in $NetConnections) {
                                $WebCerts += Get-CheckUrl -Url ("Https://" + $ComputerName + ":"  + $NetConnection.LocalPort)
                            }
                            If (( $WebCerts | Measure-Object).count -gt 1) {
                                If($Excel) {
                                    $StrWebCerts = ( $WebCerts | Where-Object {$null -ne $_.ExpirationOn} | ForEach-Object {($_.Url + " - " + $_.ExpirationOn)}) 
                                }Else{
                                    $StrWebCerts = (( $WebCerts | Where-Object {$null -ne $_.ExpirationOn} | ForEach-Object {($_.Url + " - " + $_.ExpirationOn)})  -join ",") 
                                }
                            } else {
                                $StrWebCerts = ( $WebCerts.Url + " - " +  $WebCerts.ExpirationOn) 
                            }
                            $ComputerInfo.'SSL Certificates Expiration' = $StrWebCerts 
                        }
                    }Catch{
                        Write-Warning ($ComputerName + ": Error getting SSL Certificates Information")
                    }
                    #endregion SSL Certs
                    #region Databases
                    Try{
                        $SQLVersion = ""
                        $SQL= ""
                        $SQLInstances = Invoke-Command -errorAction SilentlyContinue -HideComputerName -Session $psSession -ScriptBlock {
                            (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\').InstalledInstances
                        }
                        If ($SQLInstances){
                            $SQLVersion = Invoke-Command -errorAction SilentlyContinue -HideComputerName -Session $psSession -ScriptBlock {
                                If (Get-Command Invoke-SqlCmd -errorAction SilentlyContinue) {
                                    #Get SQL Server Version
                                    $Output = Invoke-SqlCmd -query "select @@version" -ServerInstance $env:computerName
                                    If ($null -eq $Output) {
                                        #Get SQL Server Version for Cluster
                                        [string]$Output = Invoke-SqlCmd -query "select @@version" -ServerInstance ($env:computerName).SubString(0,($env:ComputerName).Length-2)
                                    }
                                    $Output
                                }
                            }
                            If($SQLVersion){
                                $ComputerInfo.'SQL Version' = ($SQLVersion.Column1 -split "\n")[0]
                            }
                            $SQL = Invoke-Command -errorAction SilentlyContinue -HideComputerName -Session $psSession -ScriptBlock {
                                If (Get-Command Invoke-SqlCmd -errorAction SilentlyContinue) {
                                    #Get SQL Server Databases
                                    $Output = Get-ChildItem SQLSERVER:\SQL\$env:computername\Default\Databases | Select-Object Name
                                    If ($null -eq $Output) {
                                        #Get SQL Server Databases for Cluster
                                        $Output = Get-ChildItem SQLSERVER:\SQL\($env:computerName).SubString(0,($env:ComputerName).Length-2)\Default\Databases | Select-Object Name
                                    }
                                    $Output
                                }
                            }
                            If($SQL){
                                If($Excel) {
                                    $ComputerInfo.'SQL Databases' =  (($SQL | Select-Object Name).Name -join "\n") 
                                }Else{
                                    $ComputerInfo.'SQL Databases' =  (($SQL | Select-Object Name).Name -join "," )
                                }                                
                            }
                        }
                    }Catch{
                        Write-Warning ($ComputerName + ": Error getting SQL Information")
                    }
                    #endregion Databases
                    if($ArrComputerSoftware){
                        #region Forti Client
                            $FortiClient = ($ArrComputerSoftware | Where-Object {$_.Name -match "FortiClient"}).version
                            $ComputerInfo.FortiClient = $FortiClient
                        #endregion Forti Client
                        #region LanDesk Agent
                            $LanDeskVersion =  $ArrComputerSoftware | Where-Object {$_.Name -match "LanDesk"  -and $_.Name -match "Common" -and $_.Name -match "Agent"}
                            If ($LanDeskVersion) {
                                $ComputerInfo.'LANDesk Agent Installed' = $LanDeskVersion.Version
                            }
                        #endregion LanDesk Agent
                    }
                    #region RemoveSMB1
                    Try{
                        If ($RemoveSMB1) {
                            $SMB1 = Invoke-Command -errorAction SilentlyContinue -HideComputerName -Session $psSession -ScriptBlock {
                                $SMB1Enabled = $true
                                If (Get-Command Get-WindowsOptionalFeature -errorAction SilentlyContinue) {
                                    $WOF = Get-WindowsOptionalFeature -Online -ErrorAction SilentlyContinue | Where-Object {($_.FeatureName -eq "SMB1Protocol" -or $_.FeatureName -eq "SMB1Protocol-Server") -and $_.state -eq "Enabled" }
                                    If ($WOF.state -eq "Enabled") {
                                        $WOF | Disable-WindowsOptionalFeature -online -NoRestart -WarningAction SilentlyContinue | Out-Null
                                    }
                                    #Test
                                    If (Get-WindowsOptionalFeature -Online -ErrorAction SilentlyContinue | Where-Object {($_.FeatureName -eq "SMB1Protocol" -or $_.FeatureName -eq "SMB1Protocol-Server") -and $_.state -eq "Disabled"}) {
                                        $SMB1Enabled = $false
                                    }
                                }
                                If (Get-Command Get-WindowsFeature -errorAction SilentlyContinue) {
                                    $WF = Get-WindowsFeature -Name "FS-SMB1" -ErrorAction SilentlyContinue
                                    If ( $WF.Installed) {
                                        $WF | Uninstall-WindowsFeature -Restart:$false -WarningAction SilentlyContinue | Out-Null
                                    }
                                    If ((Get-WindowsFeature -Name "FS-SMB1" -ErrorAction SilentlyContinue).Installed -eq $false) {
                                        $SMB1Enabled = $false
                                    }
                                }
                                If (Get-Command Get-SmbServerConfiguration -errorAction SilentlyContinue) {	
                                    If ((Get-SmbServerConfiguration).EnableSMB1Protocol) {
                                        Set-SmbServerConfiguration -EnableSMB1Protocol $false -confirm:$false
                                        If ((Get-SmbServerConfiguration).EnableSMB1Protocol -eq $false) {
                                            $SMB1Enabled = $false
                                        }
                                    }else {
                                        $SMB1Enabled = $false
                                    }
                                }
                                If ([Environment]::OSVersion.Version -lt (new-object 'Version' 6,2)) {
                                    If ((Get-Item HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters | ForEach-Object {Get-ItemProperty $_.pspath}).SMB1 -ne 0) {
                                        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters" SMB1 -Type DWORD = 0 -Force
                                    }
                                    If ((Get-Item HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters | ForEach-Object {Get-ItemProperty $_.pspath}).SMB1 -eq 0) {
                                        $SMB1Enabled = $false
                                    }
                                }
                                Restart-Service -Force -Name "LanmanServer"
                                #Results
                                If ( $SMB1Enabled) {
                                    return ("SMB1 Enabled")
                                }else{
                                    return ("SMB1 Disabled")
                                }
                            }
                        }Else{
                            $SMB1 = Invoke-Command -errorAction SilentlyContinue -HideComputerName -Session $psSession -ScriptBlock {
                                If (Get-Command Get-SmbServerConfiguration -errorAction SilentlyContinue) {	
                                    If ((Get-SmbServerConfiguration).EnableSMB1Protocol ) {
                                        $SMB1Enabled = $true
                                    }else{
                                        $SMB1Enabled = $false
                                    }
                                }Elseif([Environment]::OSVersion.Version -lt (new-object 'Version' 6,2) -and (Get-Item HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters | ForEach-Object {Get-ItemProperty $_.pspath}).SMB1 -eq 0) {
                                    $SMB1Enabled = $false
                                }Else{
                                    $SMB1Enabled = $true
                                }
                                If (Get-Command Get-WindowsOptionalFeature -errorAction SilentlyContinue) {
                                    If (Get-WindowsOptionalFeature -Online -ErrorAction SilentlyContinue | Where-Object {($_.FeatureName -eq "SMB1Protocol" -or $_.FeatureName -eq "SMB1Protocol-Server") -and $_.state -eq "Disabled"}) {
                                        $SMB1Enabled = $false
                                    }
                                }
                                If (Get-Command Get-WindowsFeature -errorAction SilentlyContinue) {
                                    If ((Get-WindowsFeature -Name "FS-SMB1" -ErrorAction SilentlyContinue).Installed -eq $False) {
                                        $SMB1Enabled = $false
                                    }
                                }
                                #Results
                                If ($SMB1Enabled) {
                                    return ("SMB1 Enabled")
                                }else{
                                    return ("SMB1 Disabled")
                                }
                            }
                        }
                        $ComputerInfo.'SMB Status' = $SMB1
                    }Catch{
                        Write-Warning ($ComputerName + ": Error getting SMB1 Status")
                    }
                    #endregion RemoveSMB1
                    #region TLS Status
                    Try{
                        $TLS = Invoke-Command -errorAction SilentlyContinue -HideComputerName -Session $psSession -ScriptBlock {
                            $Protocols =@(
                                "Multi-Protocol Unified Hello"
                                "PCT 1.0"
                                "SSL 2.0"
                                "SSL 3.0"
                                "TLS 1.0"
                                "TLS 1.1"
                                "TLS 1.2"
                                "TLS 1.3"
                            )
                            $Modes =@(
                                "Client"
                                "Server"
                            )
                            
                            Foreach ( $Protocol in $Protocols) {
                                Foreach ($Mode in $Modes) {
                                    If ((Get-ItemProperty -Path ("HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\" + $Protocol + "\" + $Mode + "\") -Name "Enabled").Enabled -ne 0) {
                                        Write-Output ($Protocol + " " + $Mode + " Enabled")
                                    }
                                }
                            }
                        }
                        If($TLS){
                            If($Excel) {
                                $ComputerInfo.'TLS Status' = ($TLS -join "\n") 
                            }Else{
                                $ComputerInfo.'TLS Status' = ($TLS -join "," )
                            }
                        }
                    }Catch{
                        Write-Warning ($ComputerName + ": Error getting TLS Information")
                    }
                    #endregion TLS Status
                    #region Pending Reboot
                    Try{
                        $ComputerInfo.'Reboot Needed' = Invoke-Command -errorAction SilentlyContinue -Session $psSession -ScriptBlock {
                            If ((Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') `
                                -or (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired')`
                                -or (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name 'PendingFileRenameOperations' -ErrorAction Ignore)) {
                                $true
                            }else {
                                $false
                            }
                        }
                    }Catch{
                        Write-Warning ($ComputerName + ": Error getting Reboot Information")
                    }
                    #endregion Pending Reboot
                    #region WMF and PS Version
                    Try{
                        $ComputerInfo.'PowerShell Version' = Invoke-Command -errorAction SilentlyContinue -HideComputerName -Session $psSession -ScriptBlock {
                            If ($using:InstallWMF -and $PSVersionTable.PSVersion.Major -le 5 -and $PSVersionTable.PSVersion.Minor -le 1) {
                                If ((Get-PSDrive -PSProvider FileSystem) | Where-Object -Property Root -eq $useing:WSO) {
                                    $DL = [string]((Get-PSDrive -PSProvider FileSystem) | Where-Object -Property Root -eq $useing:WSO).Name
                                }Else{
                                    $DL = (68..90 | ForEach-Object{$L=[char]$_; if ((Get-PSDrive -PSProvider FileSystem).Name -notContains $L) {$L}}) | Select-Object -last 1
                                    $pinfo = New-Object System.Diagnostics.ProcessStartInfo
                                    $pinfo.FileName = "net.exe"
                                    $pinfo.RedirectStandardError = $true
                                    $pinfo.RedirectStandardOutput = $true
                                    $pinfo.UseShellExecute = $false
                                    $pinfo.Arguments = ("use " + $DL + ": " + $useing:WSO) 
                                    $p = New-Object System.Diagnostics.Process
                                    $p.StartInfo = $pinfo
                                    $p.Start() | Out-Null
                                    $p.WaitForExit()
                                    [System.GC]::Collect()
                                }
                                return ($PSVersionTable.PSVersion + "`tInstalling WMF 5.1")
                                Set-Location -Path ($DL + ":\")
                                Start-Process -FilePath $($DL + ":\Update.cmd") -ArgumentList "/instwmf"
                            }else{
                                return ($PSVersionTable.PSVersion)
                            }
                            return ($PSVersionTable.PSVersion)
                        }
                    }Catch{
                        Write-Warning ($ComputerName + ": Error getting PowerShell Information")
                    }
                    #endregion WMF and PS Version
                    #region NTLM Version
                    Try{
                        $LMCL = Invoke-Command -errorAction SilentlyContinue -Session $psSession -ScriptBlock {
                            If (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa" -Name "LmCompatibilityLevel" -ErrorAction SilentlyContinue) {
                                (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa" -Name "LmCompatibilityLevel").LmCompatibilityLevel
                            }
                        }
                        If($LMCL){
                            $LMCLN =$(switch ($LMCL) {
                                    0 {"Send LM & NTLM responses"}
                                    1 {"Send LM & NTLM - use NTLMv2 session security if negotiated"}
                                    2 {"Send NTLM responses only"}
                                    3 {"Send NTLMv2 responses only"}
                                    4 {"Send NTLMv2 responses only. Refuse LM"}
                                    5 {"Send NTLMv2 responses only. Refuse LM & NTLM"}
                                    Default {"Not Defined"}
                                })
                            $ComputerInfo.'LAN Manager Authentication Level' = $LMCLN
                        }
                    }Catch{
                        Write-Warning ($ComputerName + ": Error getting NTLM Information")
                    }
                    #endregion NTLM Version
                    #region Last Patch Date
                    Try{
                        $MSPatches = Get-CimInstance -Class win32_quickfixengineering  -CimSession $cimSession | Sort-Object {[System.DateTime]($_.InstalledOn) } -Descending -ErrorAction SilentlyContinue
                        If($MSPatches){
                            If ($MSPatches[0].InstalledOn) {
                                $LMSPatches = $MSPatches | Where-Object -Property Installedon -eq ($MSPatches[0].InstalledOn)| Select-Object HotFixID
                                If($Excel) {
                                    $ComputerInfo.'Last Hotfixes Packages' = (($LMSPatches.HotFixID -replace '"',''  -replace "'","") -join "\n")
                                }Else{
                                    $ComputerInfo.'Last Hotfixes Packages' = (($LMSPatches.HotFixID -replace '"','' -replace "'","") -join ",")
                                }
                            }
                            $ComputerInfo.'Last Hotfixes Install Date' =  $MSPatches[0].InstalledOn.ToString("MM/dd/yyyy")
                        }
                    }Catch{
                        Write-Warning ($ComputerName + ": Error getting Patch Information")
                    }
                    #endregion Last Patch Date
                    
                } else {
                    Write-Verbose ("`t`tHost WMI Not Reachable for computer: " + $ComputerName) 
                   
                }
                #Clean up remote connections
                $cimSession | Remove-CimSession
                $psSession | Remove-PSSession 
            }Else{
               
                $ComputerInfo.'Distinguished Name' = $ComputerOS.DistinguishedName 
            }

            If ($ComputerInfo) {
                #$Inventory.Add($ComputerInfo) | Out-Null
                # [void]$Inventory.Add($ComputerInfo)
                return $ComputerInfo
            }
        #Stop Script block
        } | Out-Null
    }
    
    do {
        $CJC=1
        Write-Verbose "Trying get part of data."
        $CJobs = Get-Job -State Completed
        ForEach ($CJ in $CJobs) {
            Write-Verbose "Geting job $($CJ.Name) result."
            $JobResult = Receive-Job -Id ($CJ.Id)

            if($ShowAll) {
                if($ShowInstantly) {
                    if($JobResult.Active -eq $true) {
                        Write-Host "$($JobResult.Name) is active." -ForegroundColor Green
                    } else {
                        Write-Host "$($JobResult.Name) is inactive." -ForegroundColor Red
                    }
                }
            } else {
                if($JobResult.Active -eq $true) {
                    if($ShowInstantly) {
                        Write-Host "$($JobResult.Name) is active." -ForegroundColor Green
                    }
                }
            }
            Write-Verbose "Removing job $($CJ.Name)."
            # $Inventory.Add($JobResult) | Out-Null
            [void]$Inventory.Add($JobResult)
            #$Inventory += $JobResult
            $CJC++
            Remove-Job -Id ($CJ.Id)
            $WPPC = ($((Get-Job -name * | Measure-Object).Count)/$MaxJobs*100)
            If ($WPPC -gt 100) { 
                $WPPC=100
            }
            Write-Progress  -Id 1  -Activity "Retrieving results, please wait..." -Status "$($CJ.Name)" -Percentcomplete $WPPC
        }
       Write-Progress  -Id 0 -Activity "Active Directory Computers Count: $AllComputersNamesCount" -Percentcomplete (($Inventory.count/$AllComputersNamesCount)*100 )
       Write-Verbose ("Active Directory Progress: " + $Inventory.count + "/" + $AllComputersNamesCount + " Percent: " + '{0:N0}' -f (($Inventory.count/$AllComputersNamesCount)*100)) 
        if((Get-Job -name *).Count -ge $MaxJobs) {
            Write-Verbose "Jobs are not completed ($((Get-Job -name * | Measure-Object).Count)/$MaxJobs), please wait..."
            Write-Progress  -Id 1  -Activity ("Jobs are not completed, please wait...") -Status "Sleeping" -Percentcomplete 0
            Start-Sleep $SleepTime
        }
        
    } while((Get-Job -name *).Count -ge $MaxJobs)
    Write-Progress  -Id 0 -Activity "Active Directory Computers Count: $AllComputersNamesCount" -Percentcomplete (($Inventory.count/$AllComputersNamesCount)*100)
    Write-Verbose ("Active Directory Progress: " + $Inventory.count + "/" + $AllComputersNamesCount + " Percent: " + '{0:N0}' -f (($Inventory.count/$AllComputersNamesCount)*100))
    $ScanCount++
}

do {
    $CJC=1    
    Write-Verbose "Trying get last part of data."
    $CJobs = Get-Job -State Completed
    ForEach ($CJ in $CJobs) {
            Write-Verbose "Getting job $($CJ.Name) result."
            $JobResult = Receive-Job -Id ($CJ.Id)

            if($ShowAll) {
                if($ShowInstantly) {
                    if($JobResult.Active -eq $true) {
                        Write-Host "$($JobResult.Name) is active." -ForegroundColor Green
                    } else {
                        Write-Host "$($JobResult.Name) is inactive." -ForegroundColor Red
                    }
                }
                # $Inventory.Add($JobResult) | Out-Null
                [void]$Inventory.Add($JobResult)
                # $Inventory += $JobResult	
            } else {
                if($JobResult.Active -eq $true) {
                    if($ShowInstantly) {
                        Write-Host "$($JobResult.Name) is active." -ForegroundColor Green
                    }
                    # $Inventory.Add($JobResult) | Out-Null
                    [void]$Inventory.Add($JobResult)
                    # $Inventory += $JobResult
                }
            }
            $CJC++
            Write-Verbose "Removing job $($CJ.Name)."
            Remove-Job -Id ($CJ.Id)
            $WPPC = ($((Get-Job -name * | Measure-Object).Count)/$MaxJobs*100)
            If ($WPPC -gt 100) { 
                $WPPC=100
            }
            Write-Progress  -Id 1  -Activity "Retrieving results, please wait..." -Status "$($CJ.Name)" -Percentcomplete $WPPC
        }
        
        if(Get-Job -name *) {
            Write-Verbose "All jobs are not completed ($((Get-Job -name *| Measure-Object).Count)/$MaxJobs), please wait... ($timeOutCounter)"
            Write-Progress  -Id 1  -Activity ("Jobs are not completed, please wait...") -Status "Sleeping" -Percentcomplete 0
            Start-Sleep $SleepTime
            $timeOutCounter += $SleepTime				
            Write-Progress  -Id 0 -Activity "Active Directory Computers Count: $AllComputersNamesCount" -Percentcomplete (($Inventory.count/$AllComputersNamesCount)*100)
            Write-Verbose ("Active Directory Progress: " + $Inventory.count + "/" + $AllComputersNamesCount + " Percent: " + '{0:N0}' -f (($Inventory.count/$AllComputersNamesCount)*100) )
            if($timeOutCounter -ge $TimeOut) {
                Write-Verbose "Time out... $TimeOut. Can't finish some jobs  ($((Get-Job -name * | Measure-Object).Count)/$MaxJobs) try remove it manualy."
                Break
            }
        }
    } while(Get-Job -name *)
    
Write-Verbose "Scan finished."
if ($Excel) {
    $Inventory | Sort-Object ADName | Export-Csv -NoTypeInformation -Path ($OutputFolder + "\Inventory_" + $FileDate + ".csv")
    #region Load ImportExcel
    If(-Not (Get-Module -Name ImportExcel -ListAvailable)){
        Install-Module -Name ImportExcel -Force -Confirm:$false
    }
    If (-Not (Get-Module "ImportExcel" -ErrorAction SilentlyContinue)) {
        Import-Module ImportExcel
    }   
    #endregion Load ImportExcel
    $excel = $Inventory | Export-Excel -ClearSheet -AutoFilter -AutoSize -FreezeTopRowFirstColumn -PassThru -Path ($OutputFolder + "\Inventory_" + $FileDate + ".xlsx") -WorksheetName ("Inventory_" + $FileDate) 
    $ws = $excel.Workbook.Worksheets[("Inventory_" + $FileDate)]
    $LastRow = $ws.Dimension.End.Row
    $LastColumn = $ws.Dimension.End.column
    
    #Replace commas with new lines
    # For ($CROW = 1; $CROW -le $LastRow;$CROW) {
    #     For ($CCOLUMN = 1; $CCOLUMN -le $LastColumn;$CCOLUMN) {
    #         $ws.cells[$CROW,$CCOLUMN].value = $ws.cells[$CROW,$CCOLUMN].value -replace ",", "\n"
    #     }
    # }

    Close-ExcelPackage $excel
}Else{
    $Inventory | Sort-Object ADName | Export-Csv -NoTypeInformation -Path ($OutputFolder + "\Inventory_" + $FileDate + ".csv")
}

$sw.Stop()
Write-Host ("Script took: " + (FormatElapsedTime($sw.Elapsed)) + " to run.")
#============================================================================
#endregion Main Local Machine Cleanup
#============================================================================
if($verbose) {
    $VerbosePreference = $oldverbose
  }
If (-Not [string]::IsNullOrEmpty($LogFile)) {
	Stop-Transcript
}
$VMs = $null
$AllComputers = $null
$Inventory = $null

