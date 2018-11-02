<# 
.SYNOPSIS
    Name: Default_User_Win8.1_+.ps1
    Hardens Fresh installs of Windows

.DESCRIPTION
	* Hardens c:\
	* Caches configureation files
	* Creates Store Users
	* Lock down users by loading registry and appling settings
	* Lock down users by appling GPO
	

    .DEPENDENCIES
		LGPO.exe
	
.PARAMETER LockedDown
	Lock down user accounts more
.PARAMETER LICache
	Location of the local configureation files cache
.PARAMETER Profiles
	Array of users to lockdown. If Store is enabled all store Window users will be added to the array and locked down.
.PARAMETER Store
	Enables more locked down of users and creates store Local Windows accounts 
.PARAMETER RemoteFiles
	Network path of configureation files are to copy down. 
.PARAMETER StartLayoutXML
	XML files that setup the defualt startmenu and task bar items
.PARAMETER CARoot
	Cert file of Domain Root CA
.PARAMETER CAInter
	Cert file of Domain Intermediate  CA
.PARAMETER CSCert
	Cert file of Code Signing cert
.PARAMETER LGPO
	Location of GPO backup to apply to machine and all users
.PARAMETER LGPOSU
	Location of GPO backup to apply to store users
.PARAMETER User
	Username for network configureation share
.PARAMETER Password
	Password that goes with useranme for network configureation share
.PARAMETER UserOnly
	Sets user settings only and no machine settings
.PARAMETER NoCacheUpdate
	Skip updating local cache.
.PARAMETER AllowClientTLS1
	Enables Computer to go to TLS 1.0 and TLS 1.1 sites.
.PARAMETER BackgroundFolder
	Folder name of where customized default windows backgrounds are.
.EXAMPLE
   & Default_User_Win8.1_+.ps1 -Store:$true

.NOTES
 Author: Paul Fuller
 Changes:
	* Version 2.0.01 - Added Notes and locked down Powershell and Administrative tools. Also added Option to skip updating local cache.
	* Version 2.0.02 - Blocking Powershell and MMC from launching.
	* Version 2.0.03 - Added Setting up of Favorites for users. Fixed order of operations. 
	* Version 2.0.04 - Added Chrome default settings. Fixed Hiding Network in Windows Explorer to allow UNC browsing. 
	* Version 2.0.05 - Fixed issue with Favorites
	* Version 2.0.06 - Moved Custiom settings to variables 
	* Version 2.0.07 - Fixed PSRemote does not exist. Setup Auto Arrange Icons. Fixed User Account copying issue. Add OEM Info. Hide last Logged in User.
	* Version 2.0.08 - Getting away from LGPO for Store Users as settings are not kept after sysprep.
	* Version 2.0.09 - Fixing logic issues with BGInfo. Updated Variables. Updated Chrome Settings. Updated Firewall settings. Updated Logon Background issues.
	* Version 2.0.10 - StoreDenyFolderUser files/folders are also hidden. Fixed Issues to Show Control Panel. Testing for Updated NTFS Permissions. Added work around for Logon Screen Cache.
	* Version 2.0.11 - Setting up Chrome base profile.
	* Version 2.0.12 - Exclude Default from Store hardening.
	* Version 2.0.13 - Make launch updated script after updating cache. Disable Windows Defender AntiSpyware
#>
PARAM (
	[switch]$LockedDown	  	= $false,
	[string]$LICache	  	= "C:\IT_Updates",
	[array]$Profiles  	  	= @("Default"),
	[switch]$Store	  	  	= $false,
	[string]$RemoteFiles  	= "\\share\Hardening_Files",
	[string]$StartLayoutXML	= "Win10.xml",
	[string]$CARoot			= "RootCA.cer",
	[string]$CAInter		= "InterCA.cer",
	[string]$CSCert			= "Code Signing.cer",
	[string]$LGPO			= "Windows10Ent",
	[string]$LGPOSU			= "CompletePolicy",
	[String]$User		    = $null,
	[String]$Password	    = $null,
	[switch]$UserOnly		= $false,
	[switch]$NoCacheUpdate	= $false,
	[switch]$AllowClientTLS1= $false,
	[String]$BackgroundFolder = "workstations"

)
#Force Running Script as Admin
If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{   
$arguments = "& '" + $myinvocation.mycommand.definition + "'"
Start-Process powershell -Verb runAs -ArgumentList $arguments
Break
}
#Fix issue for services
cd \
$ScriptVersion = "2.0.13"
#############################################################################
#############################################################################

#############################################################################
#region User Variables
#############################################################################
#region    +++++++ Company Specific Settings +++++++#
$HomePage = "http://github.com"
$IE_Header = ""
$IE_Footer = ""
$IE_Margin_Top = "0.500000"
$IE_Margin_Bottom = "0.500000"
$IE_Margin_Left = "0.166000"
$IE_Margin_Right = "0.166000"
$Custom_Software_Path = (${env:ProgramFiles(x86)} + "\app")
$Custom_Wallpaper_SubFolder = "PLS Wallpapers"
$Custom_User_Account_Pictures_SubFolder = ($Custom_Wallpaper_SubFolder + "\User Account Pictures")
$Custom_OEM_Logo = "LOGO_OEM.bmp"
$NTP_ManualPeerList = "time.nist.gov,0x08 north-america.pool.ntp.org,0x08"
$NTP_ManualPeerList_Store = $NTP_ManualPeerList
$BGInfo_StartupLink = "Bginfo Slient Start x64.lnk"
$BGInfo_StartupLink_Store = "Bginfo Slient Start VDI.lnk"
$SettingsPageVisibility = "showonly:printers;defaultapps;display;mousetouchpad;network-ethernet;notifications;usb"
$ChromeBaseZip = "Google_Profile_Base.zip"
#Versions of Adobe Reader to setup for.
$ARV = ("11.0","2005","DC")
$UserRange = 1..20
#region IE Domain Settings
$ZoneMap = @(
    New-Object PSObject -Property @{Site = "patchmypc.net";  Protocol = "https"; Zone = 2}
    New-Object PSObject -Property @{Site = "microsoft.com"; Protocol = "https"; Zone = 2}
    New-Object PSObject -Property @{Site = "microsoft.com"; Protocol = "http"; Zone = 2}
    New-Object PSObject -Property @{Site = "microsoft.com\download"; Protocol = "http"; Zone = 2}
    New-Object PSObject -Property @{Site = "microsoft.com\download"; Protocol = "https"; Zone = 2}
)
#endregion IE Domain Settings
#region Deny Folder
$StoreDenyFolder = @(
	($env:programdata + "\Microsoft\Windows\Start Menu\Programs\Administrative Tools") #Administrative Tools
	($env:programdata + "\Microsoft\Windows\Start Menu\Programs\Server Manager.lnk") #Server Manager
	($env:programdata + "\Microsoft\Windows\Start Menu\Programs\System Tools") #System Tools	
	($env:programdata + "\Microsoft\Windows\Start Menu\Programs\Accessories\Remote Desktop Connection.lnk") #RDP Client	
	($env:programdata + "\Microsoft\Windows\Start Menu\Programs\Microsoft .NET Compact Framework 1.0 SP3 Installer.lnk") #.Net
	($env:programdata + "\Microsoft\Windows\Start Menu\Programs\HP") #HP
)
$StoreDenyFolderUser = @(
	"\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Windows PowerShell" #PowerShell
	"\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\System Tools" #System Tools
)
#endregion Deny Folder
#region Show Control Panel Items
$RestrictCPL = @(
	"Printer"
	"Mail"
	"Java"
	"Flash"
	"Microsoft.Mouse"
)
#endregion Show Control Panel Items
#region Blacked List Programs
$BlackListPrograms = @(
	"powershell.exe"
	"PowerShell_ISE.exe"
	"mmc.exe"
	"ConfigWizards.exe"
	"nlbmgr.exe"
	"RecoveryDrive.exe"
	"RAMgmtUI.exe"
	"perfmon.exe"
	"ServerManager.exe"
	"ShieldingDataFileWizard.exe"
	"msconfig.exe"
	"msinfo32.exe"
	"TemplateDiskWizard.exe"
	"vmw.exe"
	"MdSched.exe"
	"licmgr.exe"
	"ClusterUpdateUI.exe"
	"dsac.exe"
	"dfrgui.exe"
	"cleanmgr.exe"
	"iscsicpl.exe"
	"odbcad32.exe"
)
#endregion Blacked List Programs
#region Disable Known Folders
$DisableKnownFolders = @(
	"Videos"
	"Music"
	"Pictures"
	"Network"
)
#endregion Disable Known Folders
#endregion +++++++ Company Specific Settings +++++++#

#region Services	
$DisableServices = @(
	"AdobeARMservice"							# Adobe Acrobat Update Service
	"AJRouter"									# AllJoyn Router Service
	"Browser"									# Computer Browser
	"diagnosticshub.standardcollector.service"  # Microsoft (R) Diagnostics Hub Standard Collector Service
	"diagsvc"									# Diagnostic Execution Service
	"DiagTrack"                              	# Diagnostics Tracking Service
	"dmwappushservice"                       	# WAP Push Message Routing Service (see known issues)
	"DPS"										# Diagnostic Policy Service
	#"HomeGroupListener"                      	# HomeGroup Listener
	#"HomeGroupProvider"                      	# HomeGroup Provider
	"HvHost"									# HV Host Service
	"irmon"										# Infrared monitor service
	"lfsvc"                                  	# Geolocation Service
	"icssvc"									# Windows Mobile Hotspot Service
	#"lmhosts"									# TCP/IP NetBIOS Helper	#####Breaks SMB 
	"MapsBroker"                             	# Downloaded Maps Manager
	"MSiSCSI"									# Microsoft iSCSI Initiator Service
	"MyWiFiDHCPDNS"								# Wireless PAN DHCP Server
	"NetTcpPortSharing"                      	# Net.Tcp Port Sharing Service
	"p2pimsvc"									# Peer Networking Identity Manager
	"p2psvc"									# Peer Name Resolution Protocol
	"PNRPAutoReg"								# PNRP Machine Name Publication Service
	"PNRPsvc"									# Peer Name Resolution Protocol
	"RemoteAccess"                           	# Routing and Remote Access
	"RemoteRegistry"                         	# Remote Registry
	"RetailDemo"								# Retail Demo Service
	#"RSoPProv"									# Resultant Set of Policy Provider
	"SEMgrSvc"									# Payments and NFC/SE Manager
	"SharedAccess"                           	# Internet Connection Sharing (ICS)
	"SNMPTRAP"									# SNMP Trap
	#"SSDPSRV"									# SSDP Discovery	#####Breaks SMB
	"Themes"									# Themes
	"TrkWks"                                 	# Distributed Link Tracking Client
	#"upnphost"									# UPnP Device Host    #####Breaks SMB
	"vmicguestinterface"						# Hyper-V Guest Service Interface
	"vmicheartbeat"								# Hyper-V Heartbeat Service
	"vmickvpexchange"							# Hyper-V Data Exchange Service
	"vmicrdv" 									# Hyper-V Remote Desktop Virtualization Service
	"vmicshutdown"								# Hyper-V Guest Shutdown Service
	"vmictimesync"								# Hyper-V Time Synchronization Service
	"vmicvmsession"								# Hyper-V PowerShell Direct Service
	"vmicvss"									# Hyper-V Volume Shadow Copy Requestor
	"WbioSrvc"                               	# Windows Biometric Service
	"WdiServiceHost"							# Diagnostic Service Host
	"WFDSConMgrSvc"								# Wi-Fi Direct Services Connection Manager Service
	#"WlanSvc"                               	# WLAN AutoConfig ##### Breaks Wi-Fi
	"WMPNetworkSvc"                          	# Windows Media Player Network Sharing Service
	#"wscsvc"                                	# Windows Security Center Service
	#"WSearch"                               	# Windows Search
	"XblAuthManager"                        	# Xbox Live Auth Manager
	"XblGameSave"                            	# Xbox Live Game Save Service
	"XboxGipSvc"								# Xbox Accessory Management Service
	"XboxNetApiSvc"                          	# Xbox Live Networking Service
	# Services which cannot be disabled
	#"WdNisSvc"
	#"WinDefend"
	#"WdNisSvc"
	#"SecurityHealthService"
	# "xbgm"
	# "WinHttpAutoProxySvc"
	# "BcastDVRUserService_62ab9"
)
$ManualServices = @(
	"Nameiphlpsvc"								#IP Helper
)
$AutomaticServices = @(
	"W32Time"									#Windows Time
)
#endregion Services	
#region Microsoft Store
	#Windows 10 Rev. 1803 WhiteList
	#APSS to Keep:
	$Keep =  "1527c705-839a-4832-9118-54d4Bd6a0c89",
	"c5e2524a-ea46-4f67-841f-6a9465d9d515",
	"E2A4F912-2574-4A75-9BB0-0D023378592B",
	"F46D4000-FD22-4DB4-AC8E-4E1DDDE828FE",
	"InputApp",
	"Microsoft.AAD.BrokerPlugin",
	"Microsoft.AccountsControl",
	"Microsoft.Appconnector",
	"Microsoft.AsyncTextService",
	"Microsoft.BingWeather", 
	"Microsoft.BioEnrollment",
	"Microsoft.CredDialogHost",
	"Microsoft.ECApp",
	"Microsoft.LockApp",
	"Microsoft.MSPaint",
	"Microsoft.MicrosoftEdge",
	"Microsoft.MicrosoftEdgeDevToolsClient",
	"Microsoft.MicrosoftStickyNotes", 
	"Microsoft.NET.Native.Framework.1.6",
	"Microsoft.NET.Native.Framework.1.7",
	"Microsoft.NET.Native.Framework.2.1",
	"Microsoft.NET.Native.Runtime.1.6",
	"Microsoft.NET.Native.Runtime.1.7",
	"Microsoft.NET.Native.Runtime.2.1",
	"Microsoft.Office.OneNote",
	"Microsoft.PPIProjection",
	"Microsoft.People",
	"Microsoft.Services.Store.Engagement",
	"Microsoft.SkypeApp",
	"Microsoft.StorePurchaseApp",
	"Microsoft.VCLibs.140.00",
	"Microsoft.VCLibs.140.00.UWPDesktop",
	"Microsoft.Wallet",
	"Microsoft.Win32WebViewHost",
	"Microsoft.Windows.Apprep.ChxApp",
	"Microsoft.Windows.AssignedAccessLockApp",
	"Microsoft.Windows.CapturePicker",
	"Microsoft.Windows.CloudExperienceHost",
	"Microsoft.Windows.ContentDeliveryManager",
	"Microsoft.Windows.Cortana",
	"Microsoft.Windows.HolographicFirstRun",
	"Microsoft.Windows.OOBENetworkCaptivePortal",
	"Microsoft.Windows.OOBENetworkConnectionFlow",
	"Microsoft.Windows.ParentalControls",
	"Microsoft.Windows.PeopleExperienceHost",
	"Microsoft.Windows.Photos",
	"Microsoft.Windows.PinningConfirmationDialog",
	"Microsoft.Windows.SecHealthUI",
	"Microsoft.Windows.SecondaryTileExperience",
	"Microsoft.Windows.SecureAssessmentBrowser",
	"Microsoft.Windows.ShellExperienceHost",
	"Microsoft.WindowsAlarms",
	"Microsoft.WindowsCalculator",
	"Microsoft.WindowsCamera",
	"Microsoft.WindowsFeedbackHub",
	"Microsoft.WindowsMaps",
	"Microsoft.WindowsStore",
	"Microsoft.Xbox.TCUI",
	"Microsoft.XboxApp",
	"Microsoft.XboxGameCallableUI",
	"Microsoft.XboxIdentityProvider",
	"Windows.CBSPreview",
	"Windows.MiracastView",
	"Windows.PrintDialog",
	"windows.immersivecontrolpanel",
	"Microsoft.Advertising",
	"Microsoft.Advertising.Xaml"
#endregion Microsoft Store

$LogFile = ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\Logs\" + `
		   $MyInvocation.MyCommand.Name + "_" + `
		   $env:computername + "_" + `
		   (Get-Date -format yyyyMMdd-hhmm) + ".log")
$sw = [Diagnostics.Stopwatch]::StartNew()
$HKEY = "HKU\DEFAULTUSER"
# Some paths that get used more than once
$ContentDeliveryPath = ($HKEY.replace("HKU\","HKU:\") + "\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager")
$HKEYWE = ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Explorer")
$HKEYIE = ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Internet Explorer")
$HKEYIS = ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Internet Settings")
$WindowsSearchPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search"
$UACPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
$HKLWE = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer"
$HKAR = "HKLM:\SOFTWARE\Policies\Adobe\Acrobat Reader"
$HKSCH = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL"
$UsersProfileFolder = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\' -Name "ProfilesDirectory").ProfilesDirectory
$ProfileList =  New-Object System.Collections.ArrayList
#############################################################################
#endregion User Variables
#############################################################################

#############################################################################
#region Setup Sessions
#############################################################################
#Start logging.
If (-Not [string]::IsNullOrEmpty($LogFile)) {
	If (-Not( Test-Path (Split-Path -Path $LogFile -Parent))) {
		New-Item -ItemType directory -Path (Split-Path -Path $LogFile -Parent) | Out-Null
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
#Store Setup
If ($Store) {
	$LockedDown = $True
}
#Add Registry Hives
New-PSDrive -PSProvider Registry -Name HKU -Root HKEY_USERS
New-PSDrive -PSProvider Registry -Name HKCR -Root HKEY_CLASSES_ROOT
#Share Setup
if ( $User -and $Password) {
	$Credential = New-Object System.Management.Automation.PSCredential ($User, (ConvertTo-SecureString $Password -AsPlainText -Force))
}
#Setup ProfileList
ForEach ($Profile in $Profiles) {
	If ($Profile) {
		$ProfileList.Add($Profile)
	}
}

#############################################################################
#endregion Setup Sessions
#############################################################################

#############################################################################
#region Functions
#############################################################################
function FormatElapsedTime($ts) {
    #https://stackoverflow.com/questions/3513650/timing-a-commands-execution-in-powershell
	$elapsedTime = ""

    if ( $ts.Hours -gt 0 )
    {
        $elapsedTime = [string]::Format( "{0:00} hours {1:00} min. {2:00}.{3:00} sec.", $ts.Hours, $ts.Minutes, $ts.Seconds, $ts.Milliseconds / 10 );
    }else {
        if ( $ts.Minutes -gt 0 )
        {
            $elapsedTime = [string]::Format( "{0:00} min. {1:00}.{2:00} sec.", $ts.Minutes, $ts.Seconds, $ts.Milliseconds / 10 );
        }
        else
        {
            $elapsedTime = [string]::Format( "{0:00}.{1:00} sec.", $ts.Seconds, $ts.Milliseconds / 10 );
        }

        if ($ts.Hours -eq 0 -and $ts.Minutes -eq 0 -and $ts.Seconds -eq 0)
        {
            $elapsedTime = [string]::Format("{0:00} ms.", $ts.Milliseconds);
        }

        if ($ts.Milliseconds -eq 0)
        {
            $elapsedTime = [string]::Format("{0} ms", $ts.TotalMilliseconds);
        }
    }
    return $elapsedTime
}
function Take-KeyPermissions {
    # Developed for PowerShell v4.0
    # Required Admin privileges
    # Links:
    #   http://shrekpoint.blogspot.ru/2012/08/taking-ownership-of-dcom-registry.html
    #   http://www.remkoweijnen.nl/blog/2012/01/16/take-ownership-of-a-registry-key-in-powershell/
    #   https://powertoe.wordpress.com/2010/08/28/controlling-registry-acl-permissions-with-powershell/
	# Default SID = S-1-5-32-544 Administrators Group
    param($rootKey, $key, [System.Security.Principal.SecurityIdentifier]$sid = 'S-1-5-32-544', $recurse = $true)

    switch -regex ($rootKey) {
        'HKCU|HKEY_CURRENT_USER'    { $rootKey = 'CurrentUser' }
        'HKLM|HKEY_LOCAL_MACHINE'   { $rootKey = 'LocalMachine' }
        'HKCR|HKEY_CLASSES_ROOT'    { $rootKey = 'ClassesRoot' }
        'HKCC|HKEY_CURRENT_CONFIG'  { $rootKey = 'CurrentConfig' }
        'HKU|HKEY_USERS'            { $rootKey = 'Users' }
    }

    ### Step 1 - escalate current process's privilege
    # get SeTakeOwnership, SeBackup and SeRestore privileges before executes next lines, script needs Admin privilege
    $import = '[DllImport("ntdll.dll")] public static extern int RtlAdjustPrivilege(ulong a, bool b, bool c, ref bool d);'
    $ntdll = Add-Type -Member $import -Name NtDll -PassThru
    $privileges = @{ SeTakeOwnership = 9; SeBackup =  17; SeRestore = 18 }
    foreach ($i in $privileges.Values) {
        $null = $ntdll::RtlAdjustPrivilege($i, 1, 0, [ref]0)
    }

    function Take-KeyPermissions {
        param($rootKey, $key, $sid, $recurse, $recurseLevel = 0)

        ### Step 2 - get ownerships of key - it works only for current key
        $regKey = [Microsoft.Win32.Registry]::$rootKey.OpenSubKey($key, 'ReadWriteSubTree', 'TakeOwnership')
        $acl = New-Object System.Security.AccessControl.RegistrySecurity
        $acl.SetOwner($sid)
        $regKey.SetAccessControl($acl)

        ### Step 3 - enable inheritance of permissions (not ownership) for current key from parent
        $acl.SetAccessRuleProtection($false, $false)
        $regKey.SetAccessControl($acl)

        ### Step 4 - only for top-level key, change permissions for current key and propagate it for subkeys
        # to enable propagations for subkeys, it needs to execute Steps 2-3 for each subkey (Step 5)
        if ($recurseLevel -eq 0) {
            $regKey = $regKey.OpenSubKey('', 'ReadWriteSubTree', 'ChangePermissions')
            $rule = New-Object System.Security.AccessControl.RegistryAccessRule($sid, 'FullControl', 'ContainerInherit', 'None', 'Allow')
            $acl.ResetAccessRule($rule)
            $regKey.SetAccessControl($acl)
        }

        ### Step 5 - recursively repeat steps 2-5 for subkeys
        if ($recurse) {
            foreach($subKey in $regKey.OpenSubKey('').GetSubKeyNames()) {
                Take-KeyPermissions $rootKey ($key+'\'+$subKey) $sid $recurse ($recurseLevel+1)
            }
        }
    }

    Take-KeyPermissions $rootKey $key $sid $recurse
}
function Get-CurrentUserSID {            
	[CmdletBinding()]            
	param(            
	)            
	#Source: https://techibee.com/powershell/find-the-sid-of-current-logged-on-user-using-powershell/2638
	Add-Type -AssemblyName System.DirectoryServices.AccountManagement            
	return ([System.DirectoryServices.AccountManagement.UserPrincipal]::Current).SID.Value            
}
function Set-Reg ($regPath, $name, $value, $type) {
	#Source: https://github.com/nichite/chill-out-windows-10/blob/master/chill-out-windows-10.ps1
	# String: Specifies a null-terminated string. Equivalent to REG_SZ.
	# ExpandString: Specifies a null-terminated string that contains unexpanded references to environment variables that are expanded when the value is retrieved. Equivalent to REG_EXPAND_SZ.
	# Binary: Specifies binary data in any form. Equivalent to REG_BINARY.
	# DWord: Specifies a 32-bit binary number. Equivalent to REG_DWORD.
	# MultiString: Specifies an array of null-terminated strings terminated by two null characters. Equivalent to REG_MULTI_SZ.
	# Qword: Specifies a 64-bit binary number. Equivalent to REG_QWORD.
	# Unknown: Indicates an unsupported registry data type, such as REG_RESOURCE_LIST.
	
	If(!(Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }
    New-ItemProperty -Path $regPath -Name $name -Value $value -PropertyType `
        $type -Force | Out-Null
}
Function Set-Owner {
    <#
        .Source: https://gallery.technet.microsoft.com/scriptcenter/Set-Owner-ff4db177
		.SYNOPSIS
            Changes owner of a file or folder to another user or group.

        .DESCRIPTION
            Changes owner of a file or folder to another user or group.

        .PARAMETER Path
            The folder or file that will have the owner changed.

        .PARAMETER Account
            Optional parameter to change owner of a file or folder to specified account.

            Default value is 'Builtin\Administrators'

        .PARAMETER Recurse
            Recursively set ownership on subfolders and files beneath given folder.

        .NOTES
            Name: Set-Owner
            Author: Boe Prox
            Version History:
                 1.0 - Boe Prox
                    - Initial Version

        .EXAMPLE
            Set-Owner -Path C:\temp\test.txt

            Description
            -----------
            Changes the owner of test.txt to Builtin\Administrators

        .EXAMPLE
            Set-Owner -Path C:\temp\test.txt -Account 'Domain\bprox

            Description
            -----------
            Changes the owner of test.txt to Domain\bprox

        .EXAMPLE
            Set-Owner -Path C:\temp -Recurse 

            Description
            -----------
            Changes the owner of all files and folders under C:\Temp to Builtin\Administrators

        .EXAMPLE
            Get-ChildItem C:\Temp | Set-Owner -Recurse -Account 'Domain\bprox'

            Description
            -----------
            Changes the owner of all files and folders under C:\Temp to Domain\bprox
    #>
    [cmdletbinding(
        SupportsShouldProcess = $True
    )]
    Param (
        [parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [Alias('FullName')]
        [string[]]$Path,
        [parameter()]
        [string]$Account = 'Builtin\Administrators',
        [parameter()]
        [switch]$Recurse
    )
    Begin {
        #Prevent Confirmation on each Write-Debug command when using -Debug
        If ($PSBoundParameters['Debug']) {
            $DebugPreference = 'Continue'
        }
        Try {
            [void][TokenAdjuster]
        } Catch {
            $AdjustTokenPrivileges = @"
            using System;
            using System.Runtime.InteropServices;

             public class TokenAdjuster
             {
              [DllImport("advapi32.dll", ExactSpelling = true, SetLastError = true)]
              internal static extern bool AdjustTokenPrivileges(IntPtr htok, bool disall,
              ref TokPriv1Luid newst, int len, IntPtr prev, IntPtr relen);
              [DllImport("kernel32.dll", ExactSpelling = true)]
              internal static extern IntPtr GetCurrentProcess();
              [DllImport("advapi32.dll", ExactSpelling = true, SetLastError = true)]
              internal static extern bool OpenProcessToken(IntPtr h, int acc, ref IntPtr
              phtok);
              [DllImport("advapi32.dll", SetLastError = true)]
              internal static extern bool LookupPrivilegeValue(string host, string name,
              ref long pluid);
              [StructLayout(LayoutKind.Sequential, Pack = 1)]
              internal struct TokPriv1Luid
              {
               public int Count;
               public long Luid;
               public int Attr;
              }
              internal const int SE_PRIVILEGE_DISABLED = 0x00000000;
              internal const int SE_PRIVILEGE_ENABLED = 0x00000002;
              internal const int TOKEN_QUERY = 0x00000008;
              internal const int TOKEN_ADJUST_PRIVILEGES = 0x00000020;
              public static bool AddPrivilege(string privilege)
              {
               try
               {
                bool retVal;
                TokPriv1Luid tp;
                IntPtr hproc = GetCurrentProcess();
                IntPtr htok = IntPtr.Zero;
                retVal = OpenProcessToken(hproc, TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, ref htok);
                tp.Count = 1;
                tp.Luid = 0;
                tp.Attr = SE_PRIVILEGE_ENABLED;
                retVal = LookupPrivilegeValue(null, privilege, ref tp.Luid);
                retVal = AdjustTokenPrivileges(htok, false, ref tp, 0, IntPtr.Zero, IntPtr.Zero);
                return retVal;
               }
               catch (Exception ex)
               {
                throw ex;
               }
              }
              public static bool RemovePrivilege(string privilege)
              {
               try
               {
                bool retVal;
                TokPriv1Luid tp;
                IntPtr hproc = GetCurrentProcess();
                IntPtr htok = IntPtr.Zero;
                retVal = OpenProcessToken(hproc, TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, ref htok);
                tp.Count = 1;
                tp.Luid = 0;
                tp.Attr = SE_PRIVILEGE_DISABLED;
                retVal = LookupPrivilegeValue(null, privilege, ref tp.Luid);
                retVal = AdjustTokenPrivileges(htok, false, ref tp, 0, IntPtr.Zero, IntPtr.Zero);
                return retVal;
               }
               catch (Exception ex)
               {
                throw ex;
               }
              }
             }
"@
            Add-Type $AdjustTokenPrivileges
        }

        #Activate necessary admin privileges to make changes without NTFS perms
        [void][TokenAdjuster]::AddPrivilege("SeRestorePrivilege") #Necessary to set Owner Permissions
        [void][TokenAdjuster]::AddPrivilege("SeBackupPrivilege") #Necessary to bypass Traverse Checking
        [void][TokenAdjuster]::AddPrivilege("SeTakeOwnershipPrivilege") #Necessary to override FilePermissions
    }
    Process {
        ForEach ($Item in $Path) {
            Write-Verbose "FullName: $Item"
            #The ACL objects do not like being used more than once, so re-create them on the Process block
            $DirOwner = New-Object System.Security.AccessControl.DirectorySecurity
            $DirOwner.SetOwner([System.Security.Principal.NTAccount]$Account)
            $FileOwner = New-Object System.Security.AccessControl.FileSecurity
            $FileOwner.SetOwner([System.Security.Principal.NTAccount]$Account)
            $DirAdminAcl = New-Object System.Security.AccessControl.DirectorySecurity
            $FileAdminAcl = New-Object System.Security.AccessControl.DirectorySecurity
            $AdminACL = New-Object System.Security.AccessControl.FileSystemAccessRule('Builtin\Administrators','FullControl','ContainerInherit,ObjectInherit','InheritOnly','Allow')
            $FileAdminAcl.AddAccessRule($AdminACL)
            $DirAdminAcl.AddAccessRule($AdminACL)
            Try {
                $Item = Get-Item -LiteralPath $Item -Force -ErrorAction Stop
                If (-NOT $Item.PSIsContainer) {
                    If ($PSCmdlet.ShouldProcess($Item, 'Set File Owner')) {
                        Try {
                            $Item.SetAccessControl($FileOwner)
                        } Catch {
                            Write-Warning "Couldn't take ownership of $($Item.FullName)! Taking FullControl of $($Item.Directory.FullName)"
                            $Item.Directory.SetAccessControl($FileAdminAcl)
                            $Item.SetAccessControl($FileOwner)
                        }
                    }
                } Else {
                    If ($PSCmdlet.ShouldProcess($Item, 'Set Directory Owner')) {                        
                        Try {
                            $Item.SetAccessControl($DirOwner)
                        } Catch {
                            Write-Warning "Couldn't take ownership of $($Item.FullName)! Taking FullControl of $($Item.Parent.FullName)"
                            $Item.Parent.SetAccessControl($DirAdminAcl) 
                            $Item.SetAccessControl($DirOwner)
                        }
                    }
                    If ($Recurse) {
                        [void]$PSBoundParameters.Remove('Path')
                        Get-ChildItem $Item -Force | Set-Owner @PSBoundParameters
                    }
                }
            } Catch {
                Write-Warning "$($Item): $($_.Exception.Message)"
            }
        }
    }
    End {  
        #Remove priviledges that had been granted
        [void][TokenAdjuster]::RemovePrivilege("SeRestorePrivilege") 
        [void][TokenAdjuster]::RemovePrivilege("SeBackupPrivilege") 
        [void][TokenAdjuster]::RemovePrivilege("SeTakeOwnershipPrivilege")     
    }
}
function Set-QuickAccess {
	<# 
	 
	.SYNOPSIS 
	Pin or Unpin folders to/from Quick Access in File Explorer. 
	 
	.DESCRIPTION 
	Pin or Unpin folders to/from Quick Access in File Explorer. 
	 
	.EXAMPLE 
	.\Set-QuickAccess.ps1 -Action Pin -Path "\\server\share\redirected_folders\$env:USERNAME\Links" 
	Pin the specified UNC server share to Quick Access in File Explorer. 
	 
	.EXAMPLE 
	.\Set-QuickAccess.ps1 -Action Unpin -Path "\\server\share\redirected_folders\$env:USERNAME\Links" 
	Unpin the specified UNC server share from Quick Access in File Explorer. 
	 
	.NOTES 
	Thanks to the below sources for inspiration :) 
	https://blogs.technet.microsoft.com/heyscriptingguy/2013/04/26/use-powershell-to-work-with-windows-explorer/ 
	https://www.reddit.com/r/sysadmin/comments/6g5hz4/removing_pinned_quick_access_pins_via_powershell/ 
	 
	.LINK 
	https://gallery.technet.microsoft.com/Set-QuickAccess-117e9a89 
	 
	#> 
	 
	[CmdletBinding()] 
	Param 
	( 
		[Parameter(Mandatory=$true,Position=1,HelpMessage="Pin or Unpin folder to/from Quick Access in File Explorer.")] 
		[ValidateSet("Pin", "Unpin")] 
		[string]$Action, 
		[Parameter(Mandatory=$true,Position=2,HelpMessage="Path to the folder to Pin or Unpin to/from Quick Access in File Explorer.")] 
		[string]$Path 
	) 
	 
	Write-Host "$Action to/from Quick Access: $Path.. " -NoNewline 
	 
	#Check if specified path is valid 
	If ((Test-Path -Path $Path) -ne $true) 
		{ 
			Write-Warning "Path does not exist." 
			return 
		} 
	#Check if specified path is a folder 
	If ((Test-Path -Path $Path -PathType Container) -ne $true) 
		{ 
			Write-Warning "Path is not a folder." 
			return 
		} 
	 
	#Pin or Unpin 
	$QuickAccess = New-Object -ComObject shell.application 
	$TargetObject = $QuickAccess.Namespace("shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}").Items() | where {$_.Path -eq "$Path"} 
	If ($Action -eq "Pin") 
		{ 
			If ($TargetObject -ne $null) 
				{ 
					Write-Warning "Path is already pinned to Quick Access." 
					return 
				} 
			Else 
				{ 
					$QuickAccess.Namespace("$Path").Self.InvokeVerb("pintohome")
				} 
		} 
	ElseIf ($Action -eq "Unpin") 
		{ 
			If ($TargetObject -eq $null) 
				{ 
					Write-Warning "Path is not pinned to Quick Access." 
					return 
				} 
			Else 
				{ 
					$TargetObject.InvokeVerb("unpinfromhome") 
				} 
		} 
 }
#############################################################################
#endregion Functions
#############################################################################

#############################################################################
#region Main 
#############################################################################
#============================================================================
#region Main Setup
#============================================================================
#Skip updating local cache
If (-Not $NoCacheUpdate) {
	#Setup Local Install Cache
	If (-Not( Test-Path $LICache)) {
		write-host ("Creating Local Install cache: " + $LICache)
		New-Item -ItemType directory -Path $LICache | Out-Null
		$Acl = Get-Acl $LICache
		$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule('Users', "FullControl", "ContainerInherit, ObjectInherit", "None", "Allow")
		$Acl.Setaccessrule($Ar)
		Set-Acl $LICache $Acl
	}
	#Map UNC path or local path as PSDrive
	If (-Not (Test-Path $RemoteFiles -erroraction 'silentlycontinue')) {
		#Files need explicated credentials
		If (-Not (Test-Path "PSRemote:\")) {
			If ($Credential) {
				#Credentials given as parameter 
				New-PSDrive -Name "PSRemote" -PSProvider "FileSystem" -Root $RemoteFiles -Credential $Credential | out-null
				If ($LASTEXITCODE -gt 0 ) {
					write-error "Cannot Update Local Cache"
					break
				}
			}else{
				#Credentials not given
				$Credential = Get-Credential
				New-PSDrive -Name "PSRemote" -PSProvider "FileSystem" -Root $RemoteFiles -Credential $Credential| out-null
				If ($LASTEXITCODE -gt 0 -and (-Not (Test-Path "PSRemote:"))) {
					write-error "Cannot Update Local Cache"
					break
				}
			}
		}else{
			#PSDrive already Mapped
		}
	}else{
		#Remove files are accessible with out explicated credentials. Mapping
		If (-Not (Test-Path "PSRemote:\")) {
			New-PSDrive -Name "PSRemote" -PSProvider "FileSystem" -Root $RemoteFiles | out-null
			If ($LASTEXITCODE -gt 0 -and (-Not (Test-Path "PSRemote:\"))) {
				write-error "Cannot Update Local Cache"
				break
			}
		}else{
			#PSDrive already Mapped
		}
	}
	#Sync files to local cache
	If (Test-Path "PSRemote:\") {
		write-host ("Copying to Local Install cache: " + $LICache)
		$CurrentScriptUTC = $(Get-Item $MyInvocation.MyCommand.Definition).LastWriteTimeUtc		
		Copy-Item  "PSRemote:\*" -Destination $LICache -Recurse -Force
		If ($(Get-Item ($LICache + "\" + $MyInvocation.MyCommand.Name)).LastWriteTimeUtc -gt $CurrentScriptUTC) {
			write-host ("Starting newer copy of script...")
			Stop-transcript
			&$MyInvocation.MyCommand.Definition  $MyInvocation.MyCommand.Parameters 
			exit
		}
	}
}
#Harden Permission on the c:\
# Remove user the rights to create and modify data on the root of the c:\ drive.
If (-Not $UserOnly) {
	write-host ("Hardening Permissions: " + ($env:systemdrive + "\"))
	$acl = Get-Acl ($env:systemdrive + "\")
	If ($acl.Access | where { $_.IdentityReference -eq "NT AUTHORITY\Authenticated Users" }) {
		$usersid = New-Object System.Security.Principal.Ntaccount ("NT AUTHORITY\Authenticated Users")
		$acl.PurgeAccessRules($usersid)
		$acl | Set-Acl ($env:systemdrive + "\")
	}
	If (Test-Path $Custom_Software_Path) {
		write-host ("Setting Permissions: " + $Custom_Software_Path)
		$acl = Get-Acl $Custom_Software_Path
		If (-Not ($acl.Access | where { $_.IdentityReference -eq "BUILTIN\Users" -and $_.FileSystemRights -eq "FullControl"})) {
			$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule('Users', "FullControl", "ContainerInherit, ObjectInherit", "None", "Allow")
			$Acl.Setaccessrule($Ar)
			Set-Acl $Custom_Software_Path $Acl
		}
	}
}
#============================================================================
#region Main Local Start Menu and Taskbar Settings
#============================================================================
#Import Start Menu Layout
If (-Not $UserOnly) {
	If ([environment]::OSVersion.Version.Major -ge 10) {
		If (Test-Path ($LICache + "\" + $StartLayoutXML)) {
			Import-StartLayout -LayoutPath ($LICache + "\" + $StartLayoutXML) -MountPath ($env:systemdrive + "\")
		}
	}
}
#============================================================================
#endregion Main Local Start Menu and Taskbar Settings
#============================================================================
#Create Local Store users
If ($Store) {
	#Disable Password Requirements for creating new accounts
	#secedit /export /cfg c:\secpol.cfg
	Write-Host 'Changing Password Policy to create "Window" users . . .'
	$process = Start-Process -FilePath ("secedit") -ArgumentList @("/export","/cfg","c:\secpol.cfg") -PassThru -NoNewWindow -wait
	(gc C:\secpol.cfg).replace("PasswordComplexity = 1", "PasswordComplexity = 0") | Out-File C:\secpol.cfg
	(gc C:\secpol.cfg).replace("MinimumPasswordAge = 1", "MinimumPasswordAge = 0") | Out-File C:\secpol.cfg
	(gc C:\secpol.cfg).replace("MinimumPasswordLength = 14", "MinimumPasswordLength = 0") | Out-File C:\secpol.cfg
	#secedit /configure /db c:\windows\security\local.sdb /cfg c:\secpol.cfg /areas SECURITYPOLICY
	$process = Start-Process -FilePath ("secedit") -ArgumentList @("/configure","/db","c:\windows\security\local.sdb","/cfg","c:\secpol.cfg","/areas","SECURITYPOLICY") -PassThru -NoNewWindow -wait
	rm -force c:\secpol.cfg -confirm:$false
	# net accounts /minpwage:0 /minpwlen:0
	ForEach ( $i in $UserRange) {	
		If ($i) {
			If (-Not (Get-LocalUser -Name ("Window" + $i) -erroraction 'silentlycontinue')) {
				write-host ("Creating User: " +("Window" + $i))
				New-LocalUser -Name ("Window" + $i).ToLower() -Description "LiveWire Window User" -FullName ("Window" + $i) -Password (ConvertTo-SecureString ("Window" + $i).ToLower() -AsPlainText -Force) -AccountNeverExpires -UserMayNotChangePassword | Out-Null
				Add-LocalGroupMember -Name 'Administrators' -Member ("Window" + $i) | Out-Null
				Write-Host "`tWorking on Creating user profile: " ("Window" + $i)
				#launch process as user to create user profile
				# https://msdn.microsoft.com/en-us/library/system.diagnostics.processstartinfo(v=vs.110).aspx
				$processStartInfo = New-Object System.Diagnostics.ProcessStartInfo
				$processStartInfo.UserName = ("Window" + $i)
				$processStartInfo.Domain = "."
				$processStartInfo.Password = (ConvertTo-SecureString ("Window" + $i).ToLower() -AsPlainText -Force)
				$processStartInfo.FileName = "cmd"
				$processStartInfo.Arguments = "/C echo . && echo %username% && echo ."
				$processStartInfo.LoadUserProfile = $true
				$processStartInfo.UseShellExecute = $false
				$processStartInfo.RedirectStandardOutput = $false
				$process = [System.Diagnostics.Process]::Start($processStartInfo)
				$Process.WaitForExit()   
				#Add setup user to profiles created to allow registry to be created. 
				If (Test-Path ($UsersProfileFolder + "\Window" + $i) ) {
					$ProfileList.Add(("Window" + $i).ToLower()) | Out-Null
					#Grant Current user rights on new Profiles
					Write-Host ("`tUpdating ACLs and adding to Profile List: " + ($UsersProfileFolder + "\Window" + $i))
					$Folderpath=($UsersProfileFolder + "\Window" + $i)
					$user_account=$env:username
					$Acl = Get-Acl $Folderpath
					$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule($user_account, "FullControl", "ContainerInherit, ObjectInherit", "None", "Allow")
					$Acl.Setaccessrule($Ar)
					Set-Acl $Folderpath $Acl	
				}

			}else{
				If (Test-Path ($UsersProfileFolder + "\Window" + $i) ) {
					Write-Host ("Adding to Profile List:" + ($UsersProfileFolder + "\Window" + $i))
					$ProfileList.Add(("Window" + $i).ToLower()) | Out-Null
				}
			}
		}
	}
}
#============================================================================
#endregion Main Setup
#============================================================================
#============================================================================
#region Main Set User Defaults 
#============================================================================
Write-Host ("-"*[console]::BufferWidth)
Write-Host ("Starting User Profile Setup. . .")
Write-Host ("-"*[console]::BufferWidth)
ForEach ( $CurrentProfile in $ProfileList.ToArray() ) {
	write-host ("Working with user: " + $CurrentProfile) -foregroundcolor "Magenta"
	$HKEY = ("HKU\H_" + $CurrentProfile)
	If (-Not (Test-Path $HKEY)) {
		If ($CurrentProfile.ToUpper() -eq "DEFAULT") {
			If (Test-Path ($UsersProfileFolder + "\Default\ntuser.dat")) {
				#REG LOAD $HKEY ($UsersProfileFolder + "\Default\ntuser.dat")
				[gc]::collect()
				$process = (REG LOAD  $HKEY ($UsersProfileFolder + "\Default\ntuser.dat"))
				If ($LASTEXITCODE -ne 0 ) {
					write-error ( "Cannot load profile for: " + ($UsersProfileFolder + "\" + $CurrentProfile + "\ntuser.dat") )
					continue
				}
			} else {	
				If (Test-Path ($UsersProfileFolder.Replace($UsersProfileFolder.Substring(0,1),($env:systemdrive).Substring(0,1)) + "\Default\ntuser.dat")) {
					# REG LOAD $HKEY ($UsersProfileFolder.Replace($UsersProfileFolder.Substring(0,1),($env:systemdrive).Substring(0,1)) + "\Default\ntuser.dat")
					[gc]::collect()
					$process = (REG LOAD $HKEY ($UsersProfileFolder.Replace($UsersProfileFolder.Substring(0,1),($env:systemdrive).Substring(0,1)) + "\Default\ntuser.dat"))
					If ($LASTEXITCODE -ne 0 ) {
						write-error ( "Cannot load profile for: " + ($UsersProfileFolder.Replace($UsersProfileFolder.Substring(0,1),($env:systemdrive).Substring(0,1)) + "\Default\ntuser.dat") )
						continue
					}
				}
			}
		}else{
			$UserProfile = (gwmi Win32_UserProfile |where { (Split-Path -leaf -Path ($_.LocalPath)) -eq $CurrentProfile} |select Localpath).localpath	
			If (Test-Path ($UserProfile + "\ntuser.dat")) { 
				#Load Default User Hive
				#REG LOAD $HKEY ($UserProfile + "\ntuser.dat")
				[gc]::collect()
				$process = (REG LOAD  $HKEY ($UserProfile + "\ntuser.dat"))
				If ($LASTEXITCODE -ne 0 ) {
					write-error ( "Cannot load profile for: " + ($UsersProfileFolder + "\" + $CurrentProfile + "\ntuser.dat") )
					continue
				}
			}else{
				If (Test-Path $UserProfile.Replace($UserProfile.Substring(0,1),($env:systemdrive).Substring(0,1))) {
					# REG LOAD $HKEY ($UserProfile + "\ntuser.dat")
					[gc]::collect()
					$process = (REG LOAD  $HKEY ($UserProfile + "\ntuser.dat"))
					If ($LASTEXITCODE -ne 0 ) {
						write-error ( "Cannot load profile for: " + ($UsersProfileFolder + "\" + $CurrentProfile + "\ntuser.dat") )
						continue
					}		
				}else{
					write-error ( "Cannot load profile for: " + ($UsersProfileFolder + "\" + $CurrentProfile + "\ntuser.dat") )
					continue
				}
			}
			If ($Store) {
				#Add Deny ACL
				foreach ( $file in $StoreDenyFolder) {
					If (Test-Path $file) {
						$Acl = Get-Acl ($file)
						$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule(($env:computer + "\" + $CurrentProfile), "Read", "Deny")
						$Acl.Setaccessrule($Ar)
						Set-Acl ($file) $Acl	
					}
				}
				#Add Deny ACL User Profile
				foreach ( $file in $StoreDenyFolderUser) {
					If (Test-Path ($UserProfile + "\"+ $file)) {
						$Acl = Get-Acl ($UserProfile + "\"+ $file)
						$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule(($env:computer + "\" + $CurrentProfile), "Read", "Deny")
						$Acl.Setaccessrule($Ar)
						Set-Acl ($UserProfile + "\"+ $file) $Acl	
						Get-ChildItem -path ($UserProfile + "\"+ $file) -Recurse -Force | foreach {$_.attributes = "Hidden"}
					}
				}
			}
		}
	}
	#Set Common variables
	$ContentDeliveryPath = ($HKEY.replace("HKU\","HKU:\") + "\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager")
	$HKEYWE = ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Explorer")
	$HKEYIE = ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Internet Explorer")
	$HKEYIS = ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Internet Settings")
	
	#region Setting up sounds
	write-host ("`t" + $CurrentProfile + ": Setting up sounds")
	##Beep, Sounds  and Hung Apps##
	#disable System Beep
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Control Panel\Sound") "Beep" "NO" "String"
	#Sound ExtendedSounds
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Control Panel\Sound") "ExtendedSounds" "NO" "String"
	#How long (5 seconds by default) the system waits for user processes to end after the user clicks/taps on the End task button in Task Manager
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Control Panel\Desktop") "HungAppTimeout" "5000" "String"
	#Automatically close any apps and continue to restart, shut down, or sign out of Windows 10 without the End Task dialog.
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Control Panel\Desktop") "AutoEndTasks" "1" "String"
	#When you shut down your PC, Windows gives open applications (X) (default 20) seconds to clean up and save their data before offering to close them
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Control Panel\Desktop") "WaitToKillAppTimeout" "4000" "String"
	#Disable Sound when Moving between folders
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\AppEvents\Schemes\Apps\Explorer\Navigating\.Current") "(Default)" "" "String"
	#endregion Setting up sounds
	#region Command Prompt settings
	write-host ("`t" + $CurrentProfile + ": Setting up Command Prompt")		
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Console") "QuickEdit" 1 "DWORD"
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Command Processor") "CompletionChar" 9 "DWORD"
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Command Processor") "PathCompletionChar" 9 "DWORD"
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows NT\CurrentVersion\Network\Persistent Connections") "SaveConnections" "" "ExpandString"
	#endregion Command Prompt settings
	#region Wallpaper and Screen Saver
	write-host ("`t" + $CurrentProfile + ": Setting up Screen Saver")		
	#Set Wallpaper style to stretch
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Control Panel\Desktop") "WallpaperStyle" "2" "STRING"	
	#Setup Screen Saver
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Control Panel\Desktop") "ScreenSaveActive" "1" "STRING"
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Control Panel\Desktop") "ScreenSaverIsSecure" "1" "STRING"
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Control Panel\Desktop") "ScreenSaveTimeOut" "600" "STRING"
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Control Panel\Desktop") "SCRNSAVE.EXE" "C:\Windows\system32\scrnsave.scr" "STRING"	
	#endregion Wallpaper and Screen Saver		
	#region Windows Explorer	
	write-host ("`t" + $CurrentProfile + ": Setting up Policies Windows Explorer")
	Set-Reg ($HKEYWE + "\Serialize") "StartupDelayInMSec" 0 "DWORD"
	Set-Reg ($HKEYWE + "\Advanced") "SeparateProcess" 1 "DWORD"
	Set-Reg ($HKEYWE + "\Advanced") "ServerAdminUI" 0 "DWORD"
	Set-Reg ($HKEYWE + "\Advanced") "Start_AdminToolsRoot" 0 "DWORD"
	Set-Reg ($HKEYWE + "\Advanced") "Start_PowerButtonAction" 1 "DWORD"
	Set-Reg ($HKEYWE + "\Advanced") "Start_ShowMyMusic" 0 "DWORD"
	Set-Reg ($HKEYWE + "\Advanced") "StartMenuFavorites" 0 "DWORD"
	Set-Reg ($HKEYWE + "\AutoComplete") "Append Completion" "YES" "String"
			
	#Windows 8 + navigation settings
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Explorer\StartPage") "OpenAtLogon" "0" "DWORD"
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Explorer\StartPage") "DesktopFirst" "0" "DWORD"
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Explorer\StartPage") "MakeAllAppsDefault" "0" "DWORD"
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Explorer\StartPage") "MonitorOverride" "0" "DWORD"

	#Other Settings
	#Disable AutoPlay
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Explorer\AutoplayHandlers") "DisableAutoplay" "1" "DWORD"
	#Hide File Extensions
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced") "HideFileExt" "1" "DWORD"
	#Hide Files 
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced") "Hidden" "2" "DWORD"
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced") "ShowSuperHidden" "0" "DWORD"
	#Don't create thumb.db (thumbnail) files for local files
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced") "DisableThumbnailCache" "1" "DWORD"
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "DisableThumbsDBOnNetworkFolders" "1" "DWORD"
	#Don't ask to search the internet for the correct program when opening a file with an unknown extension
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoInternetOpenWith" "1" "DWORD"
	#endregion Windows Explorer
	#region Start Menu	
	write-host ("`t" + $CurrentProfile + ": Setting up Start Menu")
	#Show Recycle Bin
	Set-Reg ($HKEYWE + "\HideDesktopIcons\NewStartPanel") "{645FF040-5081-101B-9F08-00AA002F954E}" 0 "DWORD"
	Set-Reg ($HKEYWE + "\HideDesktopIcons\ClassicStartMenu") "{645FF040-5081-101B-9F08-00AA002F954E}" 0 "DWORD"
	#Show Web browser (default)
	Set-Reg ($HKEYWE + "\HideDesktopIcons\NewStartPanel") "{871C5380-42A0-1069-A2EA-08002B30309D}" 0 "DWORD"
	Set-Reg ($HKEYWE + "\HideDesktopIcons\ClassicStartMenu") "{871C5380-42A0-1069-A2EA-08002B30309D}" 0 "DWORD"
	#endregion Start Menu
	If ($LockedDown) {
		write-host ("`t" + $CurrentProfile + ": Setting up LockDown Settings")
		#Enable Auto Arrange Icons
		#https://www.tenforums.com/tutorials/57518-turn-off-auto-arrange-desktop-icons-windows-10-a.html
		#1075839520 = Auto arrange icons = OFF and Align icons to grid = OFF
		#1075839521 = Auto arrange icons = ON and Align icons to grid = OFF
		#1075839524 = Auto arrange icons = OFF and Align icons to grid = ON
		#1075839525 = Auto arrange icons = ON and Align icons to grid = ON
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\SOFTWARE\Microsoft\Windows\Shell\Bags\1\Desktop") "FFlags" 1075839525 "DWORD"
		#region Control Panel
		If (Test-Path($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\RestrictCpl")) {
			Remove-Item ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\RestrictCpl") -Recurse
		}
		$i = 1
		ForEach ( $item in $RestrictCPL) {
				Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\RestrictCpl") $i $item "String"
				$i++
			}	
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "RestrictCpl" 1 "DWORD"
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Control Panel\Desktop") "SCRNSAVE.EXE" "scrnsave.scr" "String"
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Control Panel\Performance Control Panel") "PerfCplEnabled" 0 "DWORD"
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Control Panel\Performance Control Panel") "SolutionsEnabled" 0 "DWORD"
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Personalization") "NoChangingSoundScheme" 1 "DWORD"
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows NT\Printers") "DisableHTTPPrinting" 1 "DWORD"
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows NT\Printers") "DisableWebPnPDownload" 1 "DWORD"
		#endregion Control Panel
		#region LockDown Windows Explorer
		Set-Reg ($HKEYWE + "\Advanced") "Start_ShowDownloads" 0 "DWORD"
		Set-Reg ($HKEYWE + "\Advanced") "StartMenuAdminTools" 0 "DWORD"
		Set-Reg ($HKEYWE + "\Advanced") "Start_AdminToolsRoot" 0 "DWORD"
		Set-Reg ($HKEYWE + "\Advanced") "TaskbarSizeMove" 0 "DWORD"
		Set-Reg ($HKEYWE + "\Advanced") "Start_ShowControlPanel" 1 "DWORD"
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows NT\SharedFolders") "PublishSharedFolders" 0 "DWORD"
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows NT\SharedFolders") "PublishDfsRoots" 0 "DWORD"
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows NT\Terminal Services") "DisablePasswordSaving" 1 "DWORD"
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\WindowsMediaCenter") "MediaCenter" 1 "DWORD"
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\WindowsStore") "RemoveWindowsStore" 1 "DWORD"
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\WindowsStore") "DisableOSUpgrade" 1 "DWORD"
		#endregion LockDown Windows Explorer
		#region Windows Media Player
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\WindowsMediaPlayer") "PreventRadioPresetsRetrieval" 1 "DWORD"
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\WindowsMediaPlayer") "PreventMusicFileMetadataRetrieval" 1 "DWORD"
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\WindowsMediaPlayer") "PreventCDDVDMetadataRetrieval" 1 "DWORD"
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\WindowsMediaPlayer") "EnableScreenSaver" 1 "DWORD"
		#endregion Windows Media Player
		#region LockDown Start Menu
		#Hide This PC
		Set-Reg ($HKEYWE + "\HideDesktopIcons\NewStartPanel") "{20D04FE0-3AEA-1069-A2D8-08002B30309D}" 1 "DWORD"
		Set-Reg ($HKEYWE + "\HideDesktopIcons\ClassicStartMenu") "{20D04FE0-3AEA-1069-A2D8-08002B30309D}" 1 "DWORD"
		#Hide Frequent Access
		Set-Reg ($HKEYWE) "ShowFrequent" 0 "DWORD"
		Set-Reg ($HKEYWE) "ShowRecent" 0 "DWORD"
		# Change Explorer home screen back to "This PC"
		Set-Reg ($HKEYWE + "\Advanced") "LaunchTo" 1 "DWORD"	
		#Hide All Drives Tc
		#endregion LockDown Start Menu
		If (($Store) -and ($CurrentProfile.ToUpper() -ne "DEFAULT" )) {
			#region LockDown Store Windows Explorer
			write-host ("`t" + $CurrentProfile + ": Setting up Store settings Windows Explorer")
			#Prevent Changing Wallpaper
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop") "NoChangingWallPaper" 1 "DWORD"	
			
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "SettingsPageVisibility" $SettingsPageVisibility "String"
			#Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoDrives" 67108863 "DWORD"		
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "ClearRecentDocsOnExit" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "ConfirmFileDelete" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "DisableThumbnails" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "DisableThumbnailsOnNetworkFolders" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "DisallowCPL" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "DisallowRun" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "ForceRunOnStartMenu" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "ForceStartMenuLogOff" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "GreyMSIAds" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "HideSCAHealth" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "LinkResolveIgnoreLinkInfo" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "LockTaskbar" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoAutoTrayNotify" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoCDBurning" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoChangeStartMenu" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoClose" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoComputersNearMe" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoDFSTab" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoDriveTypeAutoRun" 255 "DWORD"		
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoDrives" 33554431 "DWORD"		
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoFileMenu" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoFind" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoFolderOptions" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoHardwareTab" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoInplaceSharing" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoInstrumentation" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoInternetOpenWith" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoLogoff" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoLowDiskSpaceChecks" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoManageMyComputerVerb" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoNetConnectDisconnect" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoNetHood" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoNetworkConnections" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoOnlinePrintsWizard" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoPreviewPane" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoPropertiesMyComputer" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoPropertiesRecycleBin" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoPublishingWizard" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoReadingPane" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoRecentDocsMenu" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoRecentDocsNetHood" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoRun" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoSMBalloonTip" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoSMConfigurePrograms" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoSMHelp" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoSMMyDocs" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoSMMyPictures" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoSearchCommInStartMenu" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoSearchComputerLinkInStartMenu" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoSearchFilesInStartMenu" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoSearchInternetInStartMenu" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoSecurityTab" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoSetTaskbar" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoSharedDocuments" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoShellSearchButton" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoStartBanner" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoStartMenuEjectPC" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoStartMenuMyGames" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoStartMenuMyMusic" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoStartMenuNetworkPlaces" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoStrCmpLogical" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoThemesTab" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoThumbnailCache" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoTrayContextMenu" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoTrayItemsDisplay" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoViewContextMenu" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoViewOnDrive" 67108863 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoWebServices" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoWinKeys" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoWindowsUpdate" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "PreventItemCreationInUsersFilesFolder" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "TurnOffSPIAnimations" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\System") "DisableChangePassword" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\System") "DisableCMD" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\System") "DisableLockWorkstation" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\System") "DisableTaskMgr" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\System") "NoAdminPage" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Uninstall") "NoAddRemovePrograms" 1 "DWORD"	
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "AddSearchInternetLinkInStartMenu" 0 "DWORD"	
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "DisableContextMenusInStart" 1 "DWORD"	
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "ForceStartSize" 1 "DWORD"	
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "HidePeopleBar" 1 "DWORD"	
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "NoSearchEverywhereLinkInStartMenu" 1 "DWORD"	
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "NoStartMenuHomegroup" 1 "DWORD"	
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "NoStartMenuRecordedTV" 1 "DWORD"	
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "NoStartMenuVideos" 1 "DWORD"	
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "NoUninstallFromStart" 1 "DWORD"	
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "DisableNotificationCenter" 1 "DWORD"	
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "NoPinningStoreToTaskbar" 1 "DWORD"	
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "NoSystraySystemPromotion" 1 "DWORD"	
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "ShowWindowsStoreAppsOnTaskbar" 2 "DWORD"	
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "DisableThumbsDBOnNetworkFolders" 1 "DWORD"	
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "DisableSearchBoxSuggestions" 1 "DWORD"	
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "NoSearchInternetTryHarderButton" 1 "DWORD"	
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "DisableIndexedLibraryExperience" 1 "DWORD"	
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "NoAutoplayfornonVolume" 1 "DWORD"	
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "NoUseStoreOpenWith" 1 "DWORD"	
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "PowerButtonAction" 4 "DWORD"	
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "NoBalloonFeatureAdvertisements" 1 "DWORD"	
			#endregion LockDown Store Windows Explorer	
			#region Known Folders Windows Explorer
			write-host ("`t" + $CurrentProfile + ": Setting up Store settings Windows Explorer Known Folders")
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "DisableKnownFolders" 1 "DWORD"	
			#Cleanup old
			If ( Test-Path ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer\DisableKnownFolders")) {
				Remove-Item ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer\DisableKnownFolders") -Recurse
			}
			ForEach ( $item in $DisableKnownFolders) {
				Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer\DisableKnownFolders") $item $item "String"
			}			
			#endregion Known Folders Windows Explorer
			#region Programs
			write-host ("`t" + $CurrentProfile + ": Setting up Store settings Programs")
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Programs") "NoDefaultPrograms" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Programs") "NoGetPrograms" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Programs") "NoInstalledUpdates" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Programs") "NoWindowsMarketplace" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Programs") "NoWindowsFeatures" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Programs") "NoProgramsAndFeatures" 1 "DWORD"
			#endregion Programs
			#region System
			write-host ("`t" + $CurrentProfile + ": Setting up Store settings System")
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\System") "NoDispScrSavPage" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\System") "NoVisualStyleChoice" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\System") "NoColorChoice" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\System") "DisableTaskMgr" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\System") "DisableRegistryTools" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\System") "DisableCMD" 1 "DWORD"
			#endregion System
			#region Updates
			write-host ("`t" + $CurrentProfile + ": Setting up Store settings Updates")
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\WindowsUpdate") "DisableWindowsUpdateAccess" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\WindowsUpdate") "DisableWindowsUpdateAccessMode" 1 "DWORD"
			#endregion Updates
			#region Internet Explorer
			write-host ("`t" + $CurrentProfile + ": Setting up Store settings Internet Explorer")
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Ext") "DisableAddonLoadTimePerformanceNotifications" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Ext") "IgnoreFrameApprovalCheck" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer") "AllowServicePoweredQSA" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer") "DisableImportExportFavorites" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Activities") "NoActivities" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Control Panel") "Autoconfig" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Control Panel") "Certificates" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Control Panel") "Connection Settings" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Control Panel") "Connwiz Admin Lock" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Control Panel") "HomePage" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Control Panel") "Proxy" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Control Panel") "Ratings" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Feed Discovery") "Enabled" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Feeds") "AllowBasicAuthInClear" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Feeds") "BackgroundSyncStatus" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Infodelivery\Restrictions") "NoSearchCustomization" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Infodelivery\Restrictions") "NoBrowserSaveWebComplete" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Infodelivery\Restrictions") "NoScheduledUpdates" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Infodelivery\Restrictions") "NoSubscriptionContent" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\LinksBar") "Enabled" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Main") "DisableFirstRunCustomize" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Main") "Start Page" $HomePage "String"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Main") "NoReportSiteProblems" "yes" "String"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\PhishingFilter") "EnabledV9" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\PhishingFilter") "Enabled" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Privacy") "EnableInPrivateBrowsing" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Restrictions") "NoBrowserOptions" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Restrictions") "NoHelpItemSendFeedback" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Restrictions") "NoHelpItemNetscapeHelp" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Restrictions") "NoHelpItemTipOfTheDay" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Restrictions") "NoHelpItemTutorial" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Restrictions") "NoHelpMenu" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Restrictions") "NoBrowserSaveAs" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Restrictions") "NoViewSource" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Restrictions") "NoSelectDownloadDir" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Safety\PrivacIE") "DisableInPrivateBlocking" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\SearchScopes") "ShowSearchSuggestionsGlobal" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Security") "DisableFixSecuritySettings" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Security") "DisableSecuritySettingsCheck" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Security\ActiveX") "BlockNonAdminActiveXInstall" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\SQM") "DisableCustomerImprovementProgram" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Suggested Sites") "Enabled" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings") "EnableAutoProxyResultCache" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings") "WarnOnBadCertRecving" 1 "DWORD"
			#endregion Internet Explorer
			#region Microsoft Edge
			write-host ("`t" + $CurrentProfile + ": Setting up Store settings Edge")
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\MicrosoftEdge\Main") "FormSuggest Passwords" "no" "String"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\MicrosoftEdge\Main") "PreventFirstRunPage" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\MicrosoftEdge\Main") "SyncFavoritesBetweenIEAndMicrosoftEdge" 1 "DWORD"
			#endregion Microsoft Edge
			#region Messenger
			write-host ("`t" + $CurrentProfile + ": Setting up Store settings Messenger")
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Messenger\Client") "PreventRun" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Messenger\Client") "CEIP" 2 "DWORD"
			#endregion Messenger
			#region Google Chrome
			write-host ("`t" + $CurrentProfile + ": Setting up Store settings Chrome")
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "AdsSettingForIntrusiveAdsSites" 2 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "AllowDinosaurEasterEgg" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "AutofillCreditCardEnabled" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "BookmarkBarEnabled" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "BrowserAddPersonEnabled" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "BrowserGuestModeEnabled" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "BuiltInDnsClientEnabled" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "CloudPrintProxyEnabled" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "CloudPrintSubmitEnabled" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "EnableMediaRouter" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "ForceGoogleSafeSearch" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "HomepageLocation" $HomePage "String"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "ImportAutofillFormData" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "ImportBookmarks" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "ImportHistory" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "ImportHomepage" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "ImportSavedPasswords" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "ImportSearchEngine" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "IncognitoModeAvailability" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "MediaRouterCastAllowAllIPs" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "PasswordManagerEnabled" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "PromotionalTabsEnabled" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "ProxyMode" "direct" "String"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "ReportMachineIDData" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "ReportPolicyData" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "ReportUserIDData" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "ReportVersionData" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "SafeBrowsingEnabled" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "SearchSuggestEnabled" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "ShowHomeButton" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "SyncDisabled" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "TranslateEnabled" 0 "DWORD"
			If (Test-Path ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome\ExtensionInstallBlacklist")) {
				Remove-Item ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome\ExtensionInstallBlacklist") -Recurse | out-null
			}
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome\ExtensionInstallBlacklist") 1 "*" "String"
			If (Test-Path ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome\RestoreOnStartupURLs")) {
				Remove-Item ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome\RestoreOnStartupURLs") -Recurse | out-null
			}
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome\RestoreOnStartupURLs") 1 $HomePage "String"				
			#endregion Google Chrome
			#region Assistance
			write-host ("`t" + $CurrentProfile + ": Setting up Store settings Assistance")
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Assistance\Client\1.0") "NoExplicitFeedback" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Assistance\Client\1.0") "NoImplicitFeedback" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Assistance\Client\1.0") "NoOnlineAssist" 1 "DWORD"
			#endregion Assistance
			#region Conferencing
			write-host ("`t" + $CurrentProfile + ": Setting up Store settings Conferencing")
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Conferencing") "NoChat" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Conferencing") "NoNewWhiteBoard" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Conferencing") "NoOldWhiteBoard" 1 "DWORD"
			#endregion Conferencing
			#region Deny Programs to run
			write-host ("`t" + $CurrentProfile + ": Setting up Store settings Deny Programs")
			#Cleanup old
			If (Test-Path ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun")) {
				Remove-Item ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun") -Recurse | out-null
			}
			$i = 1
			ForEach ( $Exe in $BlackListPrograms) {
				write-host ("`t`tBlackListing: " + $Exe)
				Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun") $i $Exe "String"
				$i++
			}
			#endregion Deny Programs to run
			#region Cloud Content
			write-host ("`t" + $CurrentProfile + ": Setting up Store settings Cloud")
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\CloudContent") "DisableWindowsSpotlightOnSettings" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\CloudContent") "DisableWindowsSpotlightWindowsWelcomeExperience" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\CloudContent") "DisableWindowsSpotlightOnActionCenter" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\CloudContent") "DisableThirdPartySuggestions" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\CloudContent") "DisableTailoredExperiencesWithDiagnosticData" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\CloudContent") "DisableWindowsSpotlightFeatures" 1 "DWORD"
			#endregion Cloud Content
			#region NetCache
			write-host ("`t" + $CurrentProfile + ": Setting up Store settings NetCache")
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\NetCache") "WorkOfflineDisabled" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\NetCache") "NoMakeAvailableOffline" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\NetCache") "NoCacheViewer" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\NetCache") "DisableFRAdminPin" 1 "DWORD"
			#endregion NetCache
			#region Network Connections
			write-host ("`t" + $CurrentProfile + ": Setting up Store settings Network Connection")
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Network Connections") "NC_ChangeBindState" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Network Connections") "NC_AddRemoveComponents" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Network Connections") "NC_LanChangeProperties" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Network Connections") "NC_LanProperties" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Network Connections") "NC_NewConnectionWizard" 0 "DWORD"
			#endregion Network Connections
			#region Powershell
			write-host ("`t" + $CurrentProfile + ": Setting up Store settings PowerShell")
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\PowerShell") "EnableScripts" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\PowerShell") "ExecutionPolicy" "AllSigned" "String"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging") "EnableScriptBlockLogging" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging") "EnableScriptBlockInvocationLogging" 1 "DWORD"
			#endregion Powershell
			#region VMware View
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\VMware, Inc.\VMware VDM\Security") "AcceptTicketSSLAuth" 1 "String"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\VMware, Inc.\VMware VDM\Client\Security") "EnableTicketSSLAuth" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\VMware, Inc.\VMware VDM\Client\Security") "LogInAsCurrentUser" "true" "String"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\VMware, Inc.\VMware VDM\Client\Security") "LogInAsCurrentUser_Display" "true" "String"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\VMware, Inc.\VMware VDM\Client\Security") "SSLCipherList" "All:!NONE:!SSLv3:!TLSv1:!MD5:!RC4:!3DES:!LOW:!EXP:!EXPORT:!DHE+AES-GCM:!DHE+AES:!DES-CBC3-SHA:!ECDHE-RSA-DES-CBC3-SHA:!AES256-SHA256:!AES256-SHA:!AES128-SHA256:!AES128-SHA:!AES256-GCM-SHA384:!AES128-GCM-SHA256:!CAMELLIA256-SHA:!ADH-AES256-GCM-SHA384:!CAMELLIA128-SHA:!ADH-AES128-GCM-SHA256:!DHE-RSA-CAMELLIA128-SHA:!DHE-RSA-CAMELLIA256-SHA" "String"
			#endregion VMware View
			#region Other ??
			write-host ("`t" + $CurrentProfile + ": Setting up Store settings Other Settings")
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\MobilityCenter") "NoMobilityCenter" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Network") "NoEntireNetwork" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\NonEnum") "{20D04FE0-3AEA-1069-A2D8-08002B30309D}" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\NonEnum") "{450D8FBA-AD25-11D0-98A8-0800361B1103}" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\PresentationSettings") "NoPresentationSettings" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\NetworkProjector") "DisableNetworkProjector" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\DataCollection") "AllowTelemetry" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\DriverSearching") "DontSearchFloppies" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\DriverSearching") "DontSearchCD" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\DriverSearching") "DontSearchWindowsUpdate" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\DriverSearching") "DontPromptForWindowsUpdate" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\HandwritingErrorReports") "PreventHandwritingErrorReports" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\RemovableStorageDevices") "Deny_All" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\SideShow") "AutoWakeDisabled" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\SideShow") "Disabled" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\TabletPC") "PreventHandwritingDataSharing" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\WCN\UI") "DisableWcnUi" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Windows Error Reporting") "DontSendAdditionalData" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Windows Error Reporting") "Disabled" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows Mail") "DisableCommunities" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows Mail") "ManualLaunchAllowed" 0 "DWORD"
			#endregion Other ??
			
			#Remove Network
			# If(Test-Path ($HKEYWE + "\Desktop\NameSpace\{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}")) {
				# write-host ("`tNetwork from This PC ") -foregroundcolor "gray"
				# Remove-Item ($HKEYWE + "\Desktop\NameSpace\{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}") -Recurse | Out-Null
			# }	
			# Set-Reg ($HKEYWE + "\CLSID\{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}\ShellFolder") "Attributes" 1048576 "DWORD"	
			# Set-Reg ($HKEYWE.replace("\Software\","\Software\Wow6432Node\") + "\CLSID\{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}\ShellFolder") "Attributes" 1048576 "DWORD"	
			
		}
	}else {
		#region Windows Explorer, Start Menu Continued
		#Show This PC
		Set-Reg ($HKEYWE + "\HideDesktopIcons\NewStartPanel") "{20D04FE0-3AEA-1069-A2D8-08002B30309D}" 0 "DWORD"
		Set-Reg ($HKEYWE + "\HideDesktopIcons\ClassicStartMenu") "{20D04FE0-3AEA-1069-A2D8-08002B30309D}" 0 "DWORD"
		#Show Frequent Access
		Set-Reg ($HKEYWE) "ShowFrequent" 1 "DWORD"
		Set-Reg ($HKEYWE) "ShowRecent" 1 "DWORD"
		# Change Explorer home screen back to ""Quick Access"
		Set-Reg ($HKEYWE + "\Advanced") "LaunchTo" 2 "DWORD"	
		#endregion Windows Explorer, Start Menu Continued
	}
	#region Internet Explorer
	write-host ("`t" + $CurrentProfile + ": Setting up Internet Explorer")
	#MigrateProxy
	Set-Reg $HKEYIS "AutoDetect" "0" "DWORD"
	#ProxyEnable
	Set-Reg $HKEYIS "ProxyEnable" "0" "DWORD"
	#Set DefaultConnectionSettings
	#AutoConfig
	$temp = (Get-ItemProperty -Path ($HKEYIS + "\Connections") -name "DefaultConnectionSettings").DefaultConnectionSettings
	if (!($temp)) {
		$temp = (70,0,0,0,3,0,0,0,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)
	} 
	$temp[8] = 1
	Set-Reg ($HKEYIS + "\Connections") "DefaultConnectionSettings" $temp  "Binary"
	#CacheScripts
	Set-Reg $HKEYIS "EnableAutoProxyResultCache" "0" "DWORD"
	#ChangeAutoConfig
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Control Panel") "Autoconfig" 0 "DWORD"
	#Set SSL Caching WUPOS
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings") "DisableCachingOfSSLPages" 0 "DWORD"
	#region LockDown Store WUPOS and DaVinci IE Settings
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings" "DisableCachingOfSSLPages" 0 "DWORD"
	#endregion LockDown Store WUPOS and DaVinci IE Settings
	$HKEYIE = ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Internet Explorer")
	#Additional Internet Explorer options
	Set-Reg ($HKEYIE + "\TabbedBrowsing") "PopupsUseNewWindow" 0 "DWORD"
	Set-Reg ($HKEYIE + "\PhishingFilter") "Enabled" 1 "DWORD"
	Set-Reg ($HKEYIE + "\Main") "Enable AutoImageResize" "YES" "String"
	Set-Reg ($HKEYIE + "\Main") "Start Page" $HomePage "String"

	#Set Margins for WUPOS
	Set-Reg ($HKEYIE + "\PageSetup") "header" $IE_Header "String"
	Set-Reg ($HKEYIE + "\PageSetup") "footer" $IE_Footer "String"
	Set-Reg ($HKEYIE + "\PageSetup") "margin_bottom" $IE_Margin_Bottom "String"
	Set-Reg ($HKEYIE + "\PageSetup") "margin_top" $IE_Margin_Top "String"
	Set-Reg ($HKEYIE + "\PageSetup") "margin_left" $IE_Margin_Left "String"
	Set-Reg ($HKEYIE + "\PageSetup") "margin_right" $IE_Margin_Right "String"
	Set-Reg ($HKEYIE.replace("\Software\","\Software\Wow6432Node\") + "\PageSetup") "header" $IE_Header "String"
	Set-Reg ($HKEYIE.replace("\Software\","\Software\Wow6432Node\") + "\PageSetup") "footer" $IE_Footer "String"
	Set-Reg ($HKEYIE.replace("\Software\","\Software\Wow6432Node\") + "\PageSetup") "margin_bottom" $IE_Margin_Bottom "String"
	Set-Reg ($HKEYIE.replace("\Software\","\Software\Wow6432Node\") + "\PageSetup") "margin_top" $IE_Margin_Top "String"
	Set-Reg ($HKEYIE.replace("\Software\","\Software\Wow6432Node\") + "\PageSetup") "margin_left" $IE_Margin_Left "String"
	Set-Reg ($HKEYIE.replace("\Software\","\Software\Wow6432Node\") + "\PageSetup") "margin_right" $IE_Margin_Right "String"

	#Clean up old keys
	Remove-Item ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains") -Recurse

	#IE Settings Trusted Sites
	Set-Reg ($HKEYIS + "\ZoneMap\Domains\blank") "about" 2 "DWORD"
	Set-Reg ($HKEYIS + "\ZoneMap\EscDomains\blank") "about" 2 "DWORD"
	Set-Reg ($HKEYIS.replace("\Software\","\Software\Wow6432Node\") + "\ZoneMap\EscDomains\blank") "about" 2 "DWORD"
	#Company Set sites 
	ForEach ( $item in $ZoneMap) {
		write-host ("`t`tAdding Site: " + $item.Site + " to zone: " + $item.Protocol + " for protocol: " + $item.Protocol)
		Set-Reg ($HKEYIS + "\ZoneMap\Domains\" +  $item.Site) $item.Protocol $item.Zone "DWORD"
	}
	#endregion Internet Explorer
	#region Windows Media Player
	write-host ("`t" + $CurrentProfile + ": Setting up Windows Media Player")
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\MediaPlayer\Setup\UserOptions") "DesktopShortcut" "No" "String"
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\MediaPlayer\Setup\UserOptions") "QuickLaunchShortcut" 0 "DWORD"
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\MediaPlayer\Preferences") "AcceptedPrivacyStatement" 1 "DWORD"
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\MediaPlayer\Preferences") "FirstRun" 0 "DWORD"
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\MediaPlayer\Preferences") "DisableMRU" 1 "DWORD"
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\MediaPlayer\Preferences") "AutoCopyCD" 0 "DWORD"
	#endregion Windows Media Player
	#Remove localization - Themes, Feeds, Favorites
	Remove-ItemProperty -Path ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\RunOnce") -Name "mctadmin" -Confirm:$False  -erroraction 'silentlycontinue'
	#Hide VMWare Tools
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\VMware, Inc.\VMware Tools") "ShowTray" 0 "DWORD"
	# Don't let apps use your advertising ID.
	Write-Host ("`t" + $CurrentProfile + ": Disabling use of Advertising Id...")
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\SOFTWARE\Microsoft\Windows\CurrentVersion\AdvertisingInfo") "Enabled" 0 "DWORD"
	# Don't let Microsoft push annoying RSS feeds about its products.
	Write-Host ("`t" + $CurrentProfile + ": Disabling Microsoft RSS Feeds...")
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\SOFTWARE\Microsoft\Feeds")  "SyncStatus" 0 "DWORD"
	# Turn off tips about Windows. If you're to the point of grabbing a script like this
	# off GitHub, chances are you don't need these.
	Write-Host ("`t" + $CurrentProfile + ": Disabling tips about Windows...")
	Set-Reg $ContentDeliveryPath "SoftLandingEnabled" 0 "DWORD"
	# Disable Bing search. No one wants these suggestions.
	Write-Host ("`t" + $CurrentProfile + ": Disabling Bing search...")
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\SOFTWARE\Microsoft\Windows\CurrentVersion\Search") "BingSearchEnabled" 0x0
	#Search 
	write-host ("`t" + $CurrentProfile + ": Search from This PC ...")
	#0 = Hidden
	#1 = Show search or Cortana icon
	#2 = Show search box
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\SOFTWARE\Microsoft\Windows\CurrentVersion\Search") "SearchboxTaskbarMode" 0 "DWORD"
	#region Remove Chrome Settings
	If (Test-Path ($UserProfile + "\AppData\Local\Google")) {
		Remove-Item -Path ($UserProfile + "\AppData\Local\Google") -Recurse -Confirm:$false | out-null
	}
	If (Test-Path ($HKEY.replace("HKU\","HKU:\") + "\SOFTWARE\Google")) {
		Remove-Item -Recurse -Confirm:$false -Path ($HKEY.replace("HKU\","HKU:\") + "\SOFTWARE\Google") | out-null
	}	
	#endregion Remove Chrome Settings
	#region Deploy Chrome Base Profile
	If (Test-Path ($LICache + "\" + $ChromeBaseZip)) {
		Write-Host ("`t" + $CurrentProfile + ": Setting-up Chrome Base Settings...")
		Expand-Archive -Path ($LICache + "\" + $ChromeBaseZip) -DestinationPath ($UserProfile + "\AppData\Local") -Force	
	}
	#endregion Deploy Chrome Base Profile
	# Unload the default profile hive
	Write-Host ("`t" + $CurrentProfile + ": Unloading User Registry")
	[gc]::collect()
	$process = (REG UNLOAD $HKEY)
	If ($LASTEXITCODE -ne 0 ) {
		[gc]::collect()
		sleep 3
		$process = (REG UNLOAD $HKEY)
		If ($LASTEXITCODE -ne 0 ) {
			write-error ("`t" + $CurrentProfile + ": Can not unload user registry!")
		}
	}
	#region Load LGPO User Settings
	# write-host ("`tLoading more setting via LGPO.exe ...") -foregroundcolor "gray"
	# If ($CurrentProfile -ne "Default" -and $Store) {
		# If ([environment]::OSVersion.Version.Major -ge 10) {
			# $RPF = (((((Get-ChildItem -Directory -Path ($LICache + "\Security Templates\" + $LGPOSU) | Select -First 1).GetDirectories()| Where {$_.name -eq "DomainSysvol" }).GetDirectories()| Where {$_.name -eq "GPO" }).GetDirectories()| Where {$_.name -eq "User" }).GetFiles() | Where {$_.name -eq "registry.pol" })
			# If ($RPF.Exists) {
				# $process = Start-Process -FilePath ($LICache + "\LGPO.EXE") -ArgumentList @("/q",('/u:' + $CurrentProfile),('"' + $RPF.FullName + '"')) -PassThru -NoNewWindow -wait
				# If ($process.ExitCode -eq 0 ) {
					# Write-Host ("`t" + $CurrentProfile + ": Applied $LGPOSU GPO")
				# }else {
					# Write-error ("`t" + $CurrentProfile + ": Error Applying $LGPOSU GPO")
				# }
			# }
		# }
	# }
	#endregion Load LGPO User Settings
	#region Replace Favorites
	If (Test-Path ($LICache + "\Favorites")) {
		write-host ("`t" + $CurrentProfile + ": Setting up Favorites")
		If ($CurrentProfile -eq "Default") {
				If (Test-Path ($UsersProfileFolder + "\Default\Favorites")) {
					Remove-Item -path ($UsersProfileFolder + "\Default\Favorites") -recurse -force
					Copy-Item  ($LICache + "\Favorites") -Destination ($UsersProfileFolder + "\Default\Favorites") -recurse -force
				}
				
		}else{
			$UserProfile = (gwmi Win32_UserProfile |where { (Split-Path -leaf -Path ($_.LocalPath)) -eq $CurrentProfile} |select Localpath).localpath	
			If (Test-Path ($UserProfile + "\Favorites")) {
				Remove-Item -path ($UserProfile + "\Favorites") -recurse -force
				Copy-Item  ($LICache + "\Favorites") -Destination ($UserProfile + "\Favorites") -recurse -force
			}
		}
	}

	#endregion Replace Favorites
	
}
Write-Host ("-"*[console]::BufferWidth)
Write-Host ("Ending User Profile Setup. . .")
Write-Host ("-"*[console]::BufferWidth)
#============================================================================
#endregion Main Set User Defaults 
#============================================================================
#============================================================================
#region Main Local Machine
#============================================================================

If (-Not $UserOnly) {

	# Cortana as running in Task View.
	Write-Host "Disabling Cortana..."
	Set-Reg $WindowsSearchPath "AllowCortana" 0 "DWORD"

	# I never liked location-based suggestions in my searches.
	Write-Host "Disabling location-based search suggestions..."
	Set-Reg $WindowsSearchPath "AllowSearchToUseLocation" 0 "DWORD"

	# Web suggestions in my search menu? No thanks.
	Write-Host "Disabling web suggestions in Windows Search..."
	Set-Reg $WindowsSearchPath "ConnectedSearchUseWeb" 0 "DWORD"
	Set-Reg $WindowsSearchPath "DisableWebSearch" 1 "DWORD"

	Write-Host "Disabling collection of OS usage data..."
	Set-Reg "HKLM:\SOFTWARE\Microsoft\SQMClient\Windows" "CEIPEnable" 0 "DWORD"

	Write-Host "Disabling telemetry data collection..."
	Set-Reg "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" "AllowTelemetry" 0 "DWORD"

	Write-Host "Disabling send additional info with error reports..."
	Set-Reg "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Error Reporting" "DontSendAdditionalData" 1 "DWORD"

	Write-Host "Disabling P2P Windows Update download and hosting..."
	Set-Reg "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\DeliveryOptimization\Config" "DownloadMode" 0 "DWORD"
	
	Write-Host "Disabling Windows Defender AntiSpyware ..."
	Set-Reg "HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender" "DisableAntiSpyware" 1 "DWORD"
	
	# WiFi Sense: HotSpot Sharing: Disable
	If (-Not (Test-Path "HKLM:\Software\Microsoft\PolicyManager\default\WiFi\AllowWiFiHotSpotReporting")) {
		Write-Host "WiFi Sense: HotSpot Sharing: Disable"
		New-Item -Path HKLM:\Software\Microsoft\PolicyManager\default\WiFi\AllowWiFiHotSpotReporting | Out-Null
	}
	
	#Remove OneDrive from This PC
	If (Test-Path "HKCR:\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}") {
		Set-Reg "HKCR:\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}" "System.IsPinnedToNameSpaceTree"  0 "DWORD"
		Set-Reg "HKCR:\WOW6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}" "System.IsPinnedToNameSpaceTree"  0 "DWORD"
	}
	#Removes UsersLibraries from This PC
	If (Test-Path "HKCR:\CLSID\{031E4825-7B94-4dc3-B131-E946B44C8DD5}") {
		Take-KeyPermissions "HKCR" "CLSID\{031E4825-7B94-4dc3-B131-E946B44C8DD5}" 
		Set-Reg "HKCR:\CLSID\{031E4825-7B94-4dc3-B131-E946B44C8DD5}" "System.IsPinnedToNameSpaceTree"  0 "DWORD"
		Take-KeyPermissions "HKCR" "WOW6432Node\CLSID\{031E4825-7B94-4dc3-B131-E946B44C8DD5}" 
		Set-Reg "HKCR:\WOW6432Node\CLSID\{031E4825-7B94-4dc3-B131-E946B44C8DD5}" "System.IsPinnedToNameSpaceTree"  0 "DWORD"
	}
	Write-host "Disabling scheduled tasks related to feedback and location."

	# We killed off the CEIP, so we won't need these tasks.
	Write-Host "Disabling CEIP scheduled tasks..."
	Disable-ScheduledTask -TaskName "Microsoft\Windows\Customer Experience Improvement Program\Consolidator" -erroraction 'silentlycontinue'| Out-Null
	Disable-ScheduledTask -TaskName "Microsoft\Windows\Customer Experience Improvement Program\KernelCeipTask" -erroraction 'silentlycontinue'| Out-Null
	Disable-ScheduledTask -TaskName "Microsoft\Windows\Customer Experience Improvement Program\UsbCeip" -erroraction 'silentlycontinue'| Out-Null

	# Remove the DMClient task (also sends feedback)
	Write-Host "Disabling Feedback scheduled tasks..."
	Disable-ScheduledTask -TaskName "Microsoft\Windows\Feedback\Siuf\DmClient" -erroraction 'silentlycontinue'| Out-Null

	# Disable location-based tasks and map tasks
	Write-Host "Disabling location-based scheduled tasks..."
	Disable-ScheduledTask -TaskName "Microsoft\Windows\Location\Notifications" -erroraction 'silentlycontinue'| Out-Null
	Disable-ScheduledTask -TaskName "Microsoft\Windows\Maps\MapsToastTask" -erroraction 'silentlycontinue'| Out-Null
	Disable-ScheduledTask -TaskName "Microsoft\Windows\Maps\MapsUpdateTask" -erroraction 'silentlycontinue'| Out-Null

	#Disable ThumbnailCache
	Set-Reg "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "DisableThumbnailCache" 1 "DWORD"
	#Harden lsass Processing|Print
	# https://windowsforum.com/threads/windows-hardening-guide-securing-the-lsass-process.230793/
	Set-Reg "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa" "RunAsPPL" 1 "DWORD"
	
	write-host ("Setting up Desktop Icons")
	# Start Menu
	Set-Reg ($HKLWE + "\HideDesktopIcons\ClassicStartMenu") "CEIPEnable" 0 "DWORD"
	#Web browser (default)
	Set-Reg ($HKLWE + "\HideDesktopIcons\ClassicStartMenu") "{871C5380-42A0-1069-A2EA-08002B30309D}" 0 "DWORD"
	Set-Reg ($HKLWE + "\HideDesktopIcons\NewStartPanel") "{871C5380-42A0-1069-A2EA-08002B30309D}" 0 "DWORD"
	If ($LockedDown) {
		#This PC
		Set-Reg ($HKLWE + "\HideDesktopIcons\ClassicStartMenu") "{20D04FE0-3AEA-1069-A2D8-08002B30309D}" 1 "DWORD"
		Set-Reg ($HKLWE + "\HideDesktopIcons\NewStartPanel") "{20D04FE0-3AEA-1069-A2D8-08002B30309D}" 1 "DWORD"
		#Recycle Bin
		Set-Reg ($HKLWE + "\HideDesktopIcons\ClassicStartMenu") "{645FF040-5081-101B-9F08-00AA002F954E}" 1 "DWORD"
		Set-Reg ($HKLWE + "\HideDesktopIcons\NewStartPanel") "{645FF040-5081-101B-9F08-00AA002F954E}" 1 "DWORD"
		#Hide Settings
		# Set-Reg "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer" "SettingsPageVisibility" "showonly:printers;defaultapps;display;mousetouchpad;network-ethernet;notifications;usb;windowsupdate" "String"
		Set-Reg "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" "shutdownwithoutlogon" 0 "DWORD"
	}else{
		#This PC
		Set-Reg ($HKLWE + "\HideDesktopIcons\ClassicStartMenu") "{20D04FE0-3AEA-1069-A2D8-08002B30309D}" 0 "DWORD"
		Set-Reg ($HKLWE + "\HideDesktopIcons\NewStartPanel") "{20D04FE0-3AEA-1069-A2D8-08002B30309D}" 0 "DWORD"
		#Recycle Bin
		Set-Reg ($HKLWE + "\HideDesktopIcons\ClassicStartMenu") "{645FF040-5081-101B-9F08-00AA002F954E}" 0 "DWORD"
		Set-Reg ($HKLWE + "\HideDesktopIcons\NewStartPanel") "{645FF040-5081-101B-9F08-00AA002F954E}" 0 "DWORD"
		#Hide Settings
		# Set-Reg "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer" "SettingsPageVisibility" "" "String"
	}

	write-host ("Setting up Windows Explorer Icons")
	# Windows Explorer
	Set-Reg ($HKLWE ) "FolderRedirectionWait" 1000 "DWORD"	
	# Added Recycle Bin to This PC
	If(!(Test-Path ($HKLWE + "\MyComputer\NameSpace\{645FF040-5081-101B-9F08-00AA002F954E}"))) {
		write-host ("`tAdded Recycle Bin to This PC") -foregroundcolor "gray"
		New-Item -Path ($HKLWE + "\MyComputer\NameSpace\{645FF040-5081-101B-9F08-00AA002F954E}") -Force | Out-Null
	}
	#Remove Pictures (folder) from This PC 
	write-host ("`tPictures folder from This PC ")  -foregroundcolor "gray"
	Set-Reg $HKLWE "{24AD3AD4-A569-4530-98E1-AB02F9417AA8}" 1 "DWORD"
	If(Test-Path ($HKLWE + "\MyComputer\NameSpace\{3ADD1653-EB32-4cb0-BBD7-DFA0ABB5ACCA}")) {
		Remove-Item ($HKLWE + "\MyComputer\NameSpace\{3ADD1653-EB32-4cb0-BBD7-DFA0ABB5ACCA}") -Recurse | Out-Null
	}
	If(Test-Path ($HKLWE.replace("\Software\","\Software\Wow6432Node\") + "\MyComputer\NameSpace\{3ADD1653-EB32-4cb0-BBD7-DFA0ABB5ACCA}")) {
		Remove-Item ($HKLWE.replace("\Software\","\Software\Wow6432Node\") + "\MyComputer\NameSpace\{3ADD1653-EB32-4cb0-BBD7-DFA0ABB5ACCA}") -Recurse | Out-Null
	}
	If(Test-Path ($HKLWE + "\MyComputer\NameSpace\{24ad3ad4-a569-4530-98e1-ab02f9417aa8}")) {
		Remove-Item ($HKLWE + "\MyComputer\NameSpace\{24ad3ad4-a569-4530-98e1-ab02f9417aa8}") -Recurse | Out-Null
	}
	If(Test-Path ($HKLWE.replace("\Software\","\Software\Wow6432Node\") + "\MyComputer\NameSpace\{24ad3ad4-a569-4530-98e1-ab02f9417aa8}")) {
		Remove-Item ($HKLWE.replace("\Software\","\Software\Wow6432Node\") + "\MyComputer\NameSpace\{24ad3ad4-a569-4530-98e1-ab02f9417aa8}") -Recurse | Out-Null
	}
	# Removes Music from This PC 
	write-host ("`tMusic folder from This PC ")  -foregroundcolor "gray"
	Set-Reg $HKLWE "{3DFDF296-DBEC-4FB4-81D1-6A3438BCF4DE}" 1 "DWORD"
	If(Test-Path ($HKLWE + "\MyComputer\NameSpace\{1CF1260C-4DD0-4ebb-811F-33C572699FDE}")) {
		Remove-Item ($HKLWE + "\MyComputer\NameSpace\{1CF1260C-4DD0-4ebb-811F-33C572699FDE}") -Recurse | Out-Null
	}
	If(Test-Path ($HKLWE.replace("\Software\","\Software\Wow6432Node\") + "\MyComputer\NameSpace\{1CF1260C-4DD0-4ebb-811F-33C572699FDE}")) {
		Remove-Item ($HKLWE.replace("\Software\","\Software\Wow6432Node\") + "\MyComputer\NameSpace\{1CF1260C-4DD0-4ebb-811F-33C572699FDE}") -Recurse | Out-Null
	}
	If(Test-Path ($HKLWE + "\MyComputer\NameSpace\{3dfdf296-dbec-4fb4-81d1-6a3438bcf4de}")) {
		Remove-Item ($HKLWE + "\MyComputer\NameSpace\{3dfdf296-dbec-4fb4-81d1-6a3438bcf4de}") -Recurse | Out-Null
	}
	If(Test-Path ($HKLWE.replace("\Software\","\Software\Wow6432Node\") + "\MyComputer\NameSpace\{3dfdf296-dbec-4fb4-81d1-6a3438bcf4de}")) {
		Remove-Item ($HKLWE.replace("\Software\","\Software\Wow6432Node\") + "\MyComputer\NameSpace\{3dfdf296-dbec-4fb4-81d1-6a3438bcf4de}") -Recurse | Out-Null
	}
	# Removes Videos from This PC 
	write-host ("`tPictures folder from This PC ") -foregroundcolor "gray"
	Set-Reg $HKLWE "{F86FA3AB-70D2-4FC7-9C99-FCBF05467F3A}" 1 "DWORD"
	If(Test-Path ($HKLWE + "\MyComputer\NameSpace\{f86fa3ab-70d2-4fc7-9c99-fcbf05467f3a}")) {
		Remove-Item ($HKLWE + "\MyComputer\NameSpace\{f86fa3ab-70d2-4fc7-9c99-fcbf05467f3a}") -Recurse | Out-Null
	}
	If(Test-Path ($HKLWE.replace("\Software\","\Software\Wow6432Node\") + "\MyComputer\NameSpace\{f86fa3ab-70d2-4fc7-9c99-fcbf05467f3a}")) {
		Remove-Item ($HKLWE.replace("\Software\","\Software\Wow6432Node\") + "\MyComputer\NameSpace\{f86fa3ab-70d2-4fc7-9c99-fcbf05467f3a}") -Recurse | Out-Null
	}
	If(Test-Path ($HKLWE + "\MyComputer\NameSpace\{A0953C92-50DC-43bf-BE83-3742FED03C9C}")) {
		Remove-Item ($HKLWE + "\MyComputer\NameSpace\{A0953C92-50DC-43bf-BE83-3742FED03C9C}") -Recurse | Out-Null
	}
	If(Test-Path ($HKLWE.replace("\Software\","\Software\Wow6432Node\") + "\MyComputer\NameSpace\{A0953C92-50DC-43bf-BE83-3742FED03C9C}")) {
		Remove-Item ($HKLWE.replace("\Software\","\Software\Wow6432Node\") + "\MyComputer\NameSpace\{A0953C92-50DC-43bf-BE83-3742FED03C9C}") -Recurse | Out-Null
	}
	# Removes 3D Objects from This PC 
	write-host ("`t3D Objects folder from This PC ") -foregroundcolor "gray"
	If(Test-Path ($HKLWE + "\MyComputer\NameSpace\{0DB7E03F-FC29-4DC6-9020-FF41B59E513A}")) {
		Remove-Item ($HKLWE + "\MyComputer\NameSpace\{0DB7E03F-FC29-4DC6-9020-FF41B59E513A}") -Recurse | Out-Null
	}
	If(Test-Path ($HKLWE.replace("\Software\","\Software\Wow6432Node\") + "\MyComputer\NameSpace\{0DB7E03F-FC29-4DC6-9020-FF41B59E513A}")) {
		Remove-Item ($HKLWE.replace("\Software\","\Software\Wow6432Node\") + "\MyComputer\NameSpace\{0DB7E03F-FC29-4DC6-9020-FF41B59E513A}") -Recurse | Out-Null
	}
	If ($LockedDown) {
		# Removes Desktop from This PC 
		write-host ("`tDesktop folder from This PC ") -foregroundcolor "gray"
		If(Test-Path ($HKLWE + "\MyComputer\NameSpace\{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}")) {
			Remove-Item ($HKLWE + "\MyComputer\NameSpace\{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}") -Recurse | Out-Null
		}
		If(Test-Path ($HKLWE.replace("\Software\","\Software\Wow6432Node\") + "\MyComputer\NameSpace\{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}")) {
			Remove-Item ($HKLWE.replace("\Software\","\Software\Wow6432Node\") + "\MyComputer\NameSpace\{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}") -Recurse | Out-Null
		}
		# Removes Documents from This PC 
		write-host ("`tDocuments folder from This PC ") -foregroundcolor "gray"
		If(Test-Path ($HKLWE + "\MyComputer\NameSpace\{A8CDFF1C-4878-43be-B5FD-F8091C1C60D0}")) {
			Remove-Item ($HKLWE + "\MyComputer\NameSpace\{A8CDFF1C-4878-43be-B5FD-F8091C1C60D0}") -Recurse | Out-Null
		}
		If(Test-Path ($HKLWE.replace("\Software\","\Software\Wow6432Node\") + "\MyComputer\NameSpace\{A8CDFF1C-4878-43be-B5FD-F8091C1C60D0}")) {
			Remove-Item ($HKLWE.replace("\Software\","\Software\Wow6432Node\") + "\MyComputer\NameSpace\{A8CDFF1C-4878-43be-B5FD-F8091C1C60D0}") -Recurse | Out-Null
		}
		If(Test-Path ($HKLWE + "\MyComputer\NameSpace\{d3162b92-9365-467a-956b-92703aca08af}")) {
			Remove-Item ($HKLWE + "\MyComputer\NameSpace\{d3162b92-9365-467a-956b-92703aca08af}") -Recurse | Out-Null
		}
		If(Test-Path ($HKLWE.replace("\Software\","\Software\Wow6432Node\") + "\MyComputer\NameSpace\{d3162b92-9365-467a-956b-92703aca08af}")) {
			Remove-Item ($HKLWE.replace("\Software\","\Software\Wow6432Node\") + "\MyComputer\NameSpace\{d3162b92-9365-467a-956b-92703aca08af}") -Recurse | Out-Null
		}
		# Removes Downloads from This PC 
		write-host ("`tDownloads folder from This PC ") -foregroundcolor "gray"
		If(Test-Path ($HKLWE + "\MyComputer\NameSpace\{088e3905-0323-4b02-9826-5d99428e115f}")) {
			Remove-Item ($HKLWE + "\MyComputer\NameSpace\{088e3905-0323-4b02-9826-5d99428e115f}") -Recurse | Out-Null
		}
		If(Test-Path ($HKLWE.replace("\Software\","\Software\Wow6432Node\") + "\MyComputer\NameSpace\{088e3905-0323-4b02-9826-5d99428e115f}")) {
			Remove-Item ($HKLWE.replace("\Software\","\Software\Wow6432Node\") + "\MyComputer\NameSpace\{088e3905-0323-4b02-9826-5d99428e115f}") -Recurse | Out-Null
		}
		If((Test-Path ($HKLWE + "\MyComputer\NameSpace\{374DE290-123F-4565-9164-39C4925E467B}"))) {
			Remove-Item ($HKLWE + "\MyComputer\NameSpace\{374DE290-123F-4565-9164-39C4925E467B}") -Recurse | Out-Null
		}
		If(Test-Path ($HKLWE.replace("\Software\","\Software\Wow6432Node\") + "\MyComputer\NameSpace\{374DE290-123F-4565-9164-39C4925E467B}")) {
			Remove-Item ($HKLWE.replace("\Software\","\Software\Wow6432Node\") + "\MyComputer\NameSpace\{374DE290-123F-4565-9164-39C4925E467B}") -Recurse | Out-Null
		}
		#Removing "Quick access" from Windows 10 File Explorer
		write-host ("`tQuick access from This PC ") -foregroundcolor "gray"
		Set-Reg ($HKLWE ) "HubMode" 1 "DWORD"
		Set-Reg ($HKLWE + "\CLSID\{679f85cb-0220-4080-b29b-5540cc05aab6}\ShellFolder") "Attributes" 2690646016 "DWORD"	
		#Removes OneDrive from This PC
		write-host ("`tOneDrive from This PC ") -foregroundcolor "gray"
		If (Test-Path ("HKCR:\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}")) {
			Set-Reg "HKCR:\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}" "System.IsPinnedToNameSpaceTree" 0 "DWORD"
			Set-Reg "HKCR:\WOW6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}" "System.IsPinnedToNameSpaceTree" 0 "DWORD"
		}
		If(Test-Path ($HKLWE + "\MyComputer\NameSpace\{018D5C66-4533-4307-9B53-224DE2ED1FE6}")) {
			Remove-Item ($HKLWE + "\MyComputer\NameSpace\{018D5C66-4533-4307-9B53-224DE2ED1FE6}") -Recurse | Out-Null
		}
		#Removes UsersLibraries from This PC
		write-host ("`tUsers Libraries from This PC ") -foregroundcolor "gray"
		If (Test-Path ("HKCR:\CLSID\{031E4825-7B94-4dc3-B131-E946B44C8DD5}")) {
			Set-Reg "HKCR:\CLSID\{031E4825-7B94-4dc3-B131-E946B44C8DD5}" "System.IsPinnedToNameSpaceTree" 0 "DWORD"
			Set-Reg "HKCR:\WOW6432Node\CLSID\{031E4825-7B94-4dc3-B131-E946B44C8DD5}" "System.IsPinnedToNameSpaceTree" 0 "DWORD"
		}
		#Remove Homegroup
		If(Test-Path ($HKLWE + "\Desktop\NameSpace\{B4FB3F98-C1EA-428d-A78A-D1F5659CBA93}")) {
			write-host ("`tHomegroup from This PC ") -foregroundcolor "gray"
			Remove-Item ($HKLWE + "\Desktop\NameSpace\{B4FB3F98-C1EA-428d-A78A-D1F5659CBA93}") -Recurse | Out-Null
		}
		#Remove Network
		# If(Test-Path ($HKLWE + "\Desktop\NameSpace\{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}")) {
			# write-host ("`tNetwork from This PC ") -foregroundcolor "gray"
			# Remove-Item ($HKLWE + "\Desktop\NameSpace\{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}") -Recurse | Out-Null
		# }	
		# Set-Reg ($HKLWE + "\CLSID\{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}\ShellFolder") "Attributes" 1048576 "DWORD"	
		# Set-Reg ($HKLWE.replace("\Software\","\Software\Wow6432Node\") + "\CLSID\{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}\ShellFolder") "Attributes" 1048576 "DWORD"	
	

	}else{

	}
	#region Chrome Default Bookmarks
	#Set Default Bookmarks <-- Was not working will revisit later.
	# If ((Test-Path ($LICache + "\Bookmarks.html")) -And (Test-Path (${env:ProgramFiles(x86)} + "\Google\Chrome\Application\master_preferences"))) {
		# write-host ("`tSetting up Chrome Bookmarks...") -foregroundcolor "gray"
		# $MPJson = Get-Content -Raw -Path (${env:ProgramFiles(x86)} + "\Google\Chrome\Application\master_preferences") | ConvertFrom-Json
		# Copy-Item  ($LICache + "\Bookmarks.html") -Destination (${env:ProgramFiles(x86)} + "\Google\Chrome\Application\Bookmarks.html") -Force
		# If (Test-Path (${env:ProgramFiles(x86)} + "\Google\Chrome\Application\Bookmarks.html")) {
			 ##Import skeleton bookmarks
			# If (-Not $MPJson.distribution.import_bookmarks_from_file) {
				# $MPJson.distribution | Add-Member import_bookmarks_from_file (${env:ProgramFiles(x86)} + "\Google\Chrome\Application\Bookmarks.html")
			# }else{
				# $MPJson.distribution.import_bookmarks_from_file = (${env:ProgramFiles(x86)} + "\Google\Chrome\Application\Bookmarks.html")
			# }
			 ##Show Bookmark Bar
			# If (-Not $MPJson.bookmark_bar ) {
				# $MPJson | Add-Member @{"bookmark_bar"=@{"show_on_all_tabs"="true"}}
			# }else{
				# If (-Not $MPJson.bookmark_bar.show_on_all_tabs) {
					# $MPJson.bookmark_bar | Add-Member "show_on_all_tabs" 'true'
				# }else{
					# $MPJson.bookmark_bar.show_on_all_tabs = 'true'
				# }
			# }
			# $MPJson | ConvertTo-Json -Compress | Out-File -filepath (${env:ProgramFiles(x86)} + "\Google\Chrome\Application\master_preferences") -Force
		# }
	# }	
	#endregion Chrome Default Bookmarks
}
#============================================================================
#endregion Main Local Machine
#============================================================================
#============================================================================
#region Main Local Machine Adobe
#============================================================================
If (-Not $UserOnly) {
	Write-Host "Setting up Adobe Policies"
	ForEach ( $CARV in $ARV ) {
		Set-Reg ($HKAR + "\" + $CARV + "\FeatureLockDown") "bUpdater" 1 "DWORD"
		Set-Reg ($HKAR + "\" + $CARV + "\FeatureLockDown") "bUsageMeasurement" 1 "DWORD"
		Set-Reg ($HKAR + "\" + $CARV + "\FeatureLockDown\cIPM") "bAllowUserToChangeMsgPrefs" 0 "DWORD"
		Set-Reg ($HKAR + "\" + $CARV + "\FeatureLockDown\cIPM") "bDontShowMsgWhenViewingDoc" 1 "DWORD"
		Set-Reg ($HKAR + "\" + $CARV + "\FeatureLockDown\cIPM") "bShowMsgAtLaunch" 0 "DWORD"
		Set-Reg ($HKAR + "\" + $CARV + "\FeatureLockDown\cWelcomeScreen") "bShowWelcomeScreen" 0 "DWORD"
		Set-Reg ($HKAR + "\" + $CARV + "\FeatureLockDown\cDefaultExecMenuItems") "tWhiteList" "Close|GeneralInfo|Quit|FirstPage|PrevPage|NextPage|LastPage|ActualSize|FitPage|FitWidth|FitHeight|SinglePage|OneColumn|TwoPages|TwoColumns|ZoomViewIn|ZoomViewOut|ShowHideBookmarks|ShowHideThumbnails|Print|GoToPage|ZoomTo|GeneralPrefs|SaveAs|FullScreenMode|OpenOrganizer|Scan|Web2PDF:OpnURL|AcroSendMail:SendMail|Spelling:Check Spelling|PageSetup|Find|FindSearch|GoBack|GoForward|FitVisible|ShowHideArticles|ShowHideFileAttachment|ShowHideAnnotManager|ShowHideFields|ShowHideOptCont|ShowHideModelTree|ShowHideSignatures|InsertPages|ExtractPages|ReplacePages|DeletePages|CropPages|RotatePages|AddFileAttachment|FindCurrentBookmark|BookmarkShowLocation|GoBackDoc|GoForwardDoc|DocHelpUserGuide|HelpReader|rolReadPage|HandMenuItem|ZoomDragMenuItem|CollectionPreview|CollectionHome|CollectionDetails|CollectionShowRoot|&Pages|Co&ntent|&Forms|Action &Wizard|Recognize &Text|P&rotection|&Sign && Certify|Doc&ument Processing|Print Pro&duction|Ja&vaScript|&Accessibility|Analy&ze|&Annotations|D&rawing Markups|Revie&w" "String"
		Set-Reg ($HKAR + "\" + $CARV + "\FeatureLockDown\cDefaultFindAttachmentPerms") "tSearchAttachmentsWhiteList" "3g2|3gp|3gpp|3gpp2|aac|ac3|aif|aiff|ani|asf|avi|bmp|cdr|cur|divx|djvu|doc|docx|dv|emf|eps|flv|f4v|gif|ico|iff|jbig2|jp2|jpeg|jpg|m2v|m4a|m4b|m4p|m4v|mid|mkv|mov|mpa|mp2|mp3|mp4|mts|nsv|ogg|ogm|ogv|pbm|pgm|png|ppm|ppt|pptx|ps|psd|qt|rtf|riff|svg|tif|ts|txt|ram|rm|rmvb|vob|wav|wma|wmf|wmv|xmb|xls|xlsx" "String"
		Set-Reg ($HKAR + "\" + $CARV + "\FeatureLockDown\cDefaultLaunchAttachmentPerms") "tBuiltInPermList" "version:1|.ade:3|.adp:3|.app:3|.arc:3|.arj:3|.asp:3|.bas:3|.bat:3|.bz:3|.bz2:3|.cab:3|.chm:3|.class:3|.cmd:3|.com:3|.command:3|.cpl:3|.crt:3|.csh:3|.desktop:3|.dll:3|.exe:3|.fxp:3|.gz:3|.hex:3|.hlp:3|.hqx:3|.hta:3|.inf:3|.ini:3|.ins:3|.isp:3|.its:3|.job:3|.js:3|.jse:3|.ksh:3|.lnk:3|.lzh:3|.mad:3|.maf:3|.mag:3|.mam:3|.maq:3|.mar:3|.mas:3|.mat:3|.mau:3|.mav:3|.maw:3|.mda:3|.mdb:3|.mde:3|.mdt:3|.mdw:3|.mdz:3|.msc:3|.msi:3|.msp:3|.mst:3|.ocx:3|.ops:3|.pcd:3|.pi:3|.pif:3|.prf:3|.prg:3|.pst:3|.rar:3|.reg:3|.scf:3|.scr:3|.sct:3|.sea:3|.shb:3|.shs:3|.sit:3|.tar:3|.taz:3|.tgz:3|.tmp:3|.url:3|.vb:3|.vbe:3|.vbs:3|.vsmacros:3|.vss:3|.vst:3|.vsw:3|.webloc:3|.ws:3|.wsc:3|.wsf:3|.wsh:3|.z:3|.zip:3|.zlo:3|.zoo:3|.pdf:2|.fdf:2|.jar:3|.pkg:3|.tool:3|.term:3|.acm:3|.asa:3|.aspx:3|.ax:3|.ad:3|.application:3|.asx:3|.cer:3|.cfg:3|.chi:3|.class:3|.clb:3|.cnt:3|.cnv:3|.cpx:3|.crx:3|.der:3|.drv:3|.fon:3|.gadget:3|.grp:3|.htt:3|.ime:3|.jnlp:3|.local:3|.manifest:3|.mmc:3|.mof:3|.msh:3|.msh1:3|.msh2:3|.mshxml:3|.msh1xml:3|.msh2xml:3|.mui:3|.nls:3|.pl:3|.perl:3|.plg:3|.ps1:3|.ps2:3|.ps1xml:3|.ps2xml:3|.psc1:3|.psc2:3|.py:3|.pyc:3|.pyo:3|.pyd:3|.rb:3|.sys:3|.tlb:3|.tsp:3|.xbap:3|.xnk:3|.xpi:3|.air:3|.appref-ms:3|.desklink:3|.glk:3|.library-ms:3|.mapimail:3|.mydocs:3|.sct:3|.search-ms:3|.searchConnector-ms:3|.vxd:3|.website:3|.zfsendtotarget:3" "String"
		Set-Reg ($HKAR + "\" + $CARV + "\FeatureLockDown\cDefaultLaunchURLPerms") "tSchemePerms" "version:2|shell:3|hcp:3|ms-help:3|ms-its:3|ms-itss:3|its:3|mk:3|mhtml:3|help:3|disk:3|afp:3|disks:3|telnet:3|ssh:3|acrobat:2|mailto:2|file:1|rlogin:3|javascript:4|data:3|jar:3|vbscript:3" "String"	
		Set-Reg ($HKAR + "\" + $CARV + "\FeatureLockDown\cDefaultLaunchURLPerms") "tSponsoredContentSchemeWhiteList" "http|https" "String"
		Set-Reg ($HKAR + "\" + $CARV + "\FeatureLockDown\cDefaultLaunchURLPerms") "tFlashContentSchemeWhiteList" "http|https|ftp|rtmp|rtmpe|rtmpt|rtmpte|rtmps|mailto" "String"
		Set-Reg ($HKAR + "\" + $CARV + "\FeatureLockDown\cServices") "bToggleAdobeDocumentServices" 1 "DWORD"
		#Wow6432Node
		Set-Reg ($HKAR.replace("\Software\","\Software\Wow6432Node\") + "\" + $CARV + "\FeatureLockDown") "bUpdater" 1 "DWORD"
		Set-Reg ($HKAR.replace("\Software\","\Software\Wow6432Node\") + "\" + $CARV + "\FeatureLockDown") "bUsageMeasurement" 1 "DWORD"
		Set-Reg ($HKAR.replace("\Software\","\Software\Wow6432Node\") + "\" + $CARV + "\FeatureLockDown\cIPM") "bAllowUserToChangeMsgPrefs" 0 "DWORD"
		Set-Reg ($HKAR.replace("\Software\","\Software\Wow6432Node\") + "\" + $CARV + "\FeatureLockDown\cIPM") "bDontShowMsgWhenViewingDoc" 1 "DWORD"
		Set-Reg ($HKAR.replace("\Software\","\Software\Wow6432Node\") + "\" + $CARV + "\FeatureLockDown\cIPM") "bShowMsgAtLaunch" 0 "DWORD"
		Set-Reg ($HKAR.replace("\Software\","\Software\Wow6432Node\") + "\" + $CARV + "\FeatureLockDown\cWelcomeScreen") "bShowWelcomeScreen" 0 "DWORD"
		Set-Reg ($HKAR.replace("\Software\","\Software\Wow6432Node\") + "\" + $CARV + "\FeatureLockDown\cDefaultExecMenuItems") "tWhiteList" "Close|GeneralInfo|Quit|FirstPage|PrevPage|NextPage|LastPage|ActualSize|FitPage|FitWidth|FitHeight|SinglePage|OneColumn|TwoPages|TwoColumns|ZoomViewIn|ZoomViewOut|ShowHideBookmarks|ShowHideThumbnails|Print|GoToPage|ZoomTo|GeneralPrefs|SaveAs|FullScreenMode|OpenOrganizer|Scan|Web2PDF:OpnURL|AcroSendMail:SendMail|Spelling:Check Spelling|PageSetup|Find|FindSearch|GoBack|GoForward|FitVisible|ShowHideArticles|ShowHideFileAttachment|ShowHideAnnotManager|ShowHideFields|ShowHideOptCont|ShowHideModelTree|ShowHideSignatures|InsertPages|ExtractPages|ReplacePages|DeletePages|CropPages|RotatePages|AddFileAttachment|FindCurrentBookmark|BookmarkShowLocation|GoBackDoc|GoForwardDoc|DocHelpUserGuide|HelpReader|rolReadPage|HandMenuItem|ZoomDragMenuItem|CollectionPreview|CollectionHome|CollectionDetails|CollectionShowRoot|&Pages|Co&ntent|&Forms|Action &Wizard|Recognize &Text|P&rotection|&Sign && Certify|Doc&ument Processing|Print Pro&duction|Ja&vaScript|&Accessibility|Analy&ze|&Annotations|D&rawing Markups|Revie&w" "String"
		Set-Reg ($HKAR.replace("\Software\","\Software\Wow6432Node\") + "\" + $CARV + "\FeatureLockDown\cDefaultFindAttachmentPerms") "tSearchAttachmentsWhiteList" "3g2|3gp|3gpp|3gpp2|aac|ac3|aif|aiff|ani|asf|avi|bmp|cdr|cur|divx|djvu|doc|docx|dv|emf|eps|flv|f4v|gif|ico|iff|jbig2|jp2|jpeg|jpg|m2v|m4a|m4b|m4p|m4v|mid|mkv|mov|mpa|mp2|mp3|mp4|mts|nsv|ogg|ogm|ogv|pbm|pgm|png|ppm|ppt|pptx|ps|psd|qt|rtf|riff|svg|tif|ts|txt|ram|rm|rmvb|vob|wav|wma|wmf|wmv|xmb|xls|xlsx" "String"
		Set-Reg ($HKAR.replace("\Software\","\Software\Wow6432Node\") + "\" + $CARV + "\FeatureLockDown\cDefaultLaunchAttachmentPerms") "tBuiltInPermList" "version:1|.ade:3|.adp:3|.app:3|.arc:3|.arj:3|.asp:3|.bas:3|.bat:3|.bz:3|.bz2:3|.cab:3|.chm:3|.class:3|.cmd:3|.com:3|.command:3|.cpl:3|.crt:3|.csh:3|.desktop:3|.dll:3|.exe:3|.fxp:3|.gz:3|.hex:3|.hlp:3|.hqx:3|.hta:3|.inf:3|.ini:3|.ins:3|.isp:3|.its:3|.job:3|.js:3|.jse:3|.ksh:3|.lnk:3|.lzh:3|.mad:3|.maf:3|.mag:3|.mam:3|.maq:3|.mar:3|.mas:3|.mat:3|.mau:3|.mav:3|.maw:3|.mda:3|.mdb:3|.mde:3|.mdt:3|.mdw:3|.mdz:3|.msc:3|.msi:3|.msp:3|.mst:3|.ocx:3|.ops:3|.pcd:3|.pi:3|.pif:3|.prf:3|.prg:3|.pst:3|.rar:3|.reg:3|.scf:3|.scr:3|.sct:3|.sea:3|.shb:3|.shs:3|.sit:3|.tar:3|.taz:3|.tgz:3|.tmp:3|.url:3|.vb:3|.vbe:3|.vbs:3|.vsmacros:3|.vss:3|.vst:3|.vsw:3|.webloc:3|.ws:3|.wsc:3|.wsf:3|.wsh:3|.z:3|.zip:3|.zlo:3|.zoo:3|.pdf:2|.fdf:2|.jar:3|.pkg:3|.tool:3|.term:3|.acm:3|.asa:3|.aspx:3|.ax:3|.ad:3|.application:3|.asx:3|.cer:3|.cfg:3|.chi:3|.class:3|.clb:3|.cnt:3|.cnv:3|.cpx:3|.crx:3|.der:3|.drv:3|.fon:3|.gadget:3|.grp:3|.htt:3|.ime:3|.jnlp:3|.local:3|.manifest:3|.mmc:3|.mof:3|.msh:3|.msh1:3|.msh2:3|.mshxml:3|.msh1xml:3|.msh2xml:3|.mui:3|.nls:3|.pl:3|.perl:3|.plg:3|.ps1:3|.ps2:3|.ps1xml:3|.ps2xml:3|.psc1:3|.psc2:3|.py:3|.pyc:3|.pyo:3|.pyd:3|.rb:3|.sys:3|.tlb:3|.tsp:3|.xbap:3|.xnk:3|.xpi:3|.air:3|.appref-ms:3|.desklink:3|.glk:3|.library-ms:3|.mapimail:3|.mydocs:3|.sct:3|.search-ms:3|.searchConnector-ms:3|.vxd:3|.website:3|.zfsendtotarget:3" "String"
		Set-Reg ($HKAR.replace("\Software\","\Software\Wow6432Node\") + "\" + $CARV + "\FeatureLockDown\cDefaultLaunchURLPerms") "tSchemePerms" "version:2|shell:3|hcp:3|ms-help:3|ms-its:3|ms-itss:3|its:3|mk:3|mhtml:3|help:3|disk:3|afp:3|disks:3|telnet:3|ssh:3|acrobat:2|mailto:2|file:1|rlogin:3|javascript:4|data:3|jar:3|vbscript:3" "String"	
		Set-Reg ($HKAR.replace("\Software\","\Software\Wow6432Node\") + "\" + $CARV + "\FeatureLockDown\cDefaultLaunchURLPerms") "tSponsoredContentSchemeWhiteList" "http|https" "String"
		Set-Reg ($HKAR.replace("\Software\","\Software\Wow6432Node\") + "\" + $CARV + "\FeatureLockDown\cDefaultLaunchURLPerms") "tFlashContentSchemeWhiteList" "http|https|ftp|rtmp|rtmpe|rtmpt|rtmpte|rtmps|mailto" "String"
		Set-Reg ($HKAR.replace("\Software\","\Software\Wow6432Node\") + "\" + $CARV + "\FeatureLockDown\cServices") "bToggleAdobeDocumentServices" 1 "DWORD"
	}
}
#============================================================================
#endregion Main Local Machine Adobe
#============================================================================
#Power Settings
# powercfg.exe /setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c
#============================================================================
#region Main Local Machine Services
#============================================================================
If (-Not $UserOnly) {
	Write-Host "Setting up Services. . . "
	# Source: https://github.com/W4RH4WK/Debloat-Windows-10/blob/master/scripts/disable-services.ps1
	#Services to Disable
	foreach ($service in $DisableServices) {
		If ( Get-Service -Name $service -erroraction 'silentlycontinue') {
			write-host ("`tDisabling: " + (Get-Service -Name $service).DisplayName ) -foregroundcolor green 
			Get-Service -Name $service | Stop-Service 
			Get-Service -Name $service | Set-Service -StartupType Disabled
		}
	}
	#Services to set as Manual
	foreach ($service in $ManualServices) {
		If ( Get-Service -Name $service -erroraction 'silentlycontinue') {
			write-host ("`tManual Startup: " + (Get-Service -Name $service).DisplayName ) -foregroundcolor yellow 
			Get-Service -Name $service | Stop-Service 
			Get-Service -Name $service | Set-Service -StartupType Manual
		}
	}
	#Services to set as Automatic
	foreach ($service in $AutomaticServices) {
		If ( Get-Service -Name $service -erroraction 'silentlycontinue') {
			write-host ("`tAutomatic Startup: " + (Get-Service -Name $service).DisplayName ) -foregroundcolor red 
			Get-Service -Name $service | Stop-Service 
			Get-Service -Name $service | Set-Service -StartupType Automatic
			Get-Service -Name $service | Start-Service 
		}
	}
}
#============================================================================
#endregion Main Local Machine Services
#============================================================================
#============================================================================
#region Main Local Machine Tweaks
#============================================================================
If (-Not $UserOnly) {
	Write-Host "Setting up other Tweaks. . . "
	#Disable a Paging Executive
	Set-Reg "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management" "DisablePagingExecutive" 1 "DWORD"
	#Trend-Micro Performance Fix
	#Set-Reg "HKLM:\SYSTEM\CurrentControlSet\Services\TmFilter\Parameters" "DisableCtProcCheck" 1 "DWORD"
	#Force Splwow64.exe process doesn't end after a print job finishes
	#Set-Reg "HKLM:\SYSTEM\CurrentControlSet\Control\Print" "SplWOW64TimeOutSeconds" 10 "DWORD"
	#Disable RDP Drive Redirection
	Set-Reg "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services" "fDisableCpm" 1 "DWORD"
	#Do not set default client printer to be default printer in a session
	Set-Reg "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services" "fForceClientLptDef" 1 "DWORD"
	#Hide Users on login screen
	Set-Reg "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" "dontdisplaylastusername" 1 "DWORD"
}
#============================================================================
#endregion Main Local Machine Tweaks
#============================================================================
#============================================================================
#region Main Local Machine Certs
#============================================================================
If (-Not $UserOnly) {
	Write-Host ("Setting up Certificates...")
	If (Test-Path ($LICache + "\" + $CARoot)) {
		Write-Host ("Importing Domain CA Root: " + $LICache + "\" + $CARoot)
		Import-Certificate -Filepath ($LICache + "\" + $CARoot) -CertStoreLocation cert:\LocalMachine\Root | out-null
	}
	If (Test-Path ($LICache + "\" + $CAInter)) {
		Write-Host ("Importing Domain CA Intermediate : " + $LICache + "\" + $CAInter)
		Import-Certificate -Filepath ($LICache + "\" + $CAInter) -CertStoreLocation cert:\LocalMachine\CA | out-null
	}
	#Error Importing Code Signing Cert
	If (Test-Path ( $LICache + "\" + $CSCert )) {
		Write-Host ("Importing Code Signing Cert : " + $LICache + "\" + $CSCert)
		Import-Certificate -Filepath ($LICache + "\" + $CSCert) -CertStoreLocation cert:\LocalMachine\TrustedPublisher | out-null
	}
}
#============================================================================
#endregion Main Local Machine Certs
#============================================================================
#============================================================================
#region Main Local Schannel for PCI
#============================================================================
If (-Not $UserOnly) {
	Write-Host "Setting up SSL PCI 2018 Standard. . . "
	Set-Reg ($HKSCH + "\Ciphers\AES 128/128") "Enabled" 4294967295 "DWORD"
	Set-Reg ($HKSCH + "\Ciphers\AES 256/256") "Enabled" 4294967295 "DWORD"
	Set-Reg ($HKSCH + "\Ciphers\DES 56/56") "Enabled" 0 "DWORD"
	Set-Reg ($HKSCH + "\Ciphers\NULL") "Enabled" 0 "DWORD"
	Set-Reg ($HKSCH + "\Ciphers\RC2 128/128") "Enabled" 0 "DWORD"
	Set-Reg ($HKSCH + "\Ciphers\RC2 40/128") "Enabled" 0 "DWORD"
	Set-Reg ($HKSCH + "\Ciphers\RC2 56/128") "Enabled" 0 "DWORD"
	Set-Reg ($HKSCH + "\Ciphers\RC4 128/128") "Enabled" 0 "DWORD"
	Set-Reg ($HKSCH + "\Ciphers\RC4 40/128") "Enabled" 0 "DWORD"
	Set-Reg ($HKSCH + "\Ciphers\RC4 56/128") "Enabled" 0 "DWORD"
	Set-Reg ($HKSCH + "\Ciphers\RC4 64/128") "Enabled" 0 "DWORD"
	Set-Reg ($HKSCH + "\Ciphers\Triple DES 168") "Enabled" 0 "DWORD"

	Set-Reg ($HKSCH + "\Hashes\MD5") "Enabled" 0 "DWORD"
	Set-Reg ($HKSCH + "\Hashes\SHA") "Enabled" 0 "DWORD"
	Set-Reg ($HKSCH + "\Hashes\SHA256") "Enabled" 4294967295 "DWORD"
	Set-Reg ($HKSCH + "\Hashes\SHA384") "Enabled" 4294967295 "DWORD"
	Set-Reg ($HKSCH + "\Hashes\SHA512") "Enabled" 4294967295 "DWORD"

	Set-Reg ($HKSCH + "\KeyExchangeAlgorithms\Diffie-Hellman") "Enabled" 4294967295 "DWORD"
	Set-Reg ($HKSCH + "\KeyExchangeAlgorithms\Diffie-Hellman") "ServerMinKeyBitLength" 2048 "DWORD"
	Set-Reg ($HKSCH + "\KeyExchangeAlgorithms\ECDH") "Enabled" 4294967295 "DWORD"
	Set-Reg ($HKSCH + "\KeyExchangeAlgorithms\PKCS") "Enabled" 4294967295 "DWORD"

	Set-Reg ($HKSCH + "\Protocols\Multi-Protocol Unified Hello\Client") "Enabled" 0 "DWORD"
	Set-Reg ($HKSCH + "\Protocols\Multi-Protocol Unified Hello\Client") "DisabledByDefault" 1 "DWORD"
	Set-Reg ($HKSCH + "\Protocols\Multi-Protocol Unified Hello\Server") "Enabled" 0 "DWORD"
	Set-Reg ($HKSCH + "\Protocols\Multi-Protocol Unified Hello\Server") "DisabledByDefault" 1 "DWORD"
	Set-Reg ($HKSCH + "\Protocols\PCT 1.0\Client") "Enabled" 0 "DWORD"
	Set-Reg ($HKSCH + "\Protocols\PCT 1.0\Client") "DisabledByDefault" 1 "DWORD"
	Set-Reg ($HKSCH + "\Protocols\PCT 1.0\Server") "Enabled" 0 "DWORD"
	Set-Reg ($HKSCH + "\Protocols\PCT 1.0\Server") "DisabledByDefault" 1 "DWORD"
	Set-Reg ($HKSCH + "\Protocols\SSL 2.0\Client") "Enabled" 0 "DWORD"
	Set-Reg ($HKSCH + "\Protocols\SSL 2.0\Client") "DisabledByDefault" 1 "DWORD"
	Set-Reg ($HKSCH + "\Protocols\SSL 2.0\Server") "Enabled" 0 "DWORD"
	Set-Reg ($HKSCH + "\Protocols\SSL 2.0\Server") "DisabledByDefault" 1 "DWORD"
	Set-Reg ($HKSCH + "\Protocols\SSL 3.0\Client") "Enabled" 0 "DWORD"
	Set-Reg ($HKSCH + "\Protocols\SSL 3.0\Client") "DisabledByDefault" 1 "DWORD"
	Set-Reg ($HKSCH + "\Protocols\SSL 3.0\Server") "Enabled" 0 "DWORD"
	Set-Reg ($HKSCH + "\Protocols\SSL 3.0\Server") "DisabledByDefault" 1 "DWORD"
	If ($AllowClientTLS1) {
		Set-Reg ($HKSCH + "\Protocols\TLS 1.0\Client") "Enabled" 1"DWORD"
	}else{
		Set-Reg ($HKSCH + "\Protocols\TLS 1.0\Client") "Enabled" 0 "DWORD"
	}
	Set-Reg ($HKSCH + "\Protocols\TLS 1.0\Client") "DisabledByDefault" 1 "DWORD"
	Set-Reg ($HKSCH + "\Protocols\TLS 1.0\Server") "Enabled" 0 "DWORD"
	Set-Reg ($HKSCH + "\Protocols\TLS 1.0\Server") "DisabledByDefault" 1 "DWORD"
	If ($AllowClientTLS1) {
		Set-Reg ($HKSCH + "\Protocols\TLS 1.1\Client") "Enabled" 1 "DWORD"
	}else{
		Set-Reg ($HKSCH + "\Protocols\TLS 1.1\Client") "Enabled" 0 "DWORD"
	}
	Set-Reg ($HKSCH + "\Protocols\TLS 1.1\Client") "DisabledByDefault" 1 "DWORD"
	Set-Reg ($HKSCH + "\Protocols\TLS 1.1\Server") "Enabled" 0 "DWORD"
	Set-Reg ($HKSCH + "\Protocols\TLS 1.1\Server") "DisabledByDefault" 1 "DWORD"
	Set-Reg ($HKSCH + "\Protocols\TLS 1.2\Client") "Enabled" 4294967295 "DWORD"
	Set-Reg ($HKSCH + "\Protocols\TLS 1.2\Client") "DisabledByDefault" 0 "DWORD"
	Set-Reg ($HKSCH + "\Protocols\TLS 1.2\Server") "Enabled" 4294967295 "DWORD"
	Set-Reg ($HKSCH + "\Protocols\TLS 1.2\Server") "DisabledByDefault" 0 "DWORD"
}
#============================================================================
#endregion Main Local Schannel for PCI
#============================================================================
#============================================================================
#region Main Local User Icons
#============================================================================
If (-Not $UserOnly) {
	Write-Host "Setting up User Icons . . . "
	If (Test-Path ($LICache + "\" + $Custom_User_Account_Pictures_SubFolder)) {
		copy-item ($LICache + "\" + $Custom_User_Account_Pictures_SubFolder + "\*.*") -Destination ($env:programdata + "\Microsoft\User Account Pictures") -force
		Remove-Item ($env:programdata + "\Microsoft\User Account Pictures\*.dat") -force
	}
}
#============================================================================
#endregion Main Local User Icons
#============================================================================
#============================================================================
#region Main Local Background
#============================================================================
If (-Not $UserOnly) {
	Write-Host "Setting up Background . . . "
	#Set Default Picture
	Set-Owner -Path ($env:windir + "\Web\Wallpaper\Windows\img0.jpg")
	#Add Administrators with full control
	$Folderpath=($env:windir + "\Web\Wallpaper\Windows\img0.jpg")
	$user_account='Administrators'
	$Acl = Get-Acl $Folderpath
	$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule($user_account,"FullControl", "None", "None", "Allow")
	$Acl.Setaccessrule($Ar)
	Set-Acl $Folderpath $Acl
	If (Test-Path ($LICache + "\" + $Custom_Wallpaper_SubFolder + "\" + $BackgroundFolder + "\img0.jpg")) {	
		copy-item ($LICache + "\" + $Custom_Wallpaper_SubFolder + "\" + $BackgroundFolder + "\img0.jpg") -Destination ($env:windir + "\Web\Wallpaper\Windows\img0.jpg") -force
	}
	If (Test-Path ($LICache + "\" + $Custom_Wallpaper_SubFolder + "\" + $BackgroundFolder + "\Backgrounds")) {	
		If (-Not( Test-Path ($env:windir + "\System32\oobe\info\backgrounds\"))) {
			New-Item -ItemType directory -Path ($env:windir + "\system32\oobe\info\backgrounds")
		}
		copy-item ($LICache + "\" + $Custom_Wallpaper_SubFolder + "\" + $BackgroundFolder + "\Backgrounds\*.*") -Destination ($env:windir + "\System32\oobe\info\backgrounds\") -force
	}
	#region Clear Lock Screen Cache
	#Add Administrators with full control
	$user_account='Administrators'
	$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule($user_account,"FullControl", "None", "None", "Allow")
	$Folderpath=Get-item ($env:programdata + "\Microsoft\Windows\SystemData")
	$Acl = Get-Acl $Folderpath.FullName
	$Acl.Setaccessrule($Ar)
	Set-Acl $Folderpath.FullName $Acl
	ForEach ($F1 in (Get-ChildItem $Folderpath)) {
		#Add Permissions on S-1-5-18
		$Acl = Get-Acl $F1.FullName
		$Acl.Setaccessrule($Ar)
		Set-Acl $F1.FullName $Acl
		ForEach ($F2 in (Get-ChildItem $F1.FullName)) {
			#ReadOnly
			$Acl = Get-Acl $F2.FullName
			$Acl.Setaccessrule($Ar)
			Set-Acl $F2.FullName $Acl
			ForEach ($F3 in (Get-ChildItem $F2.FullName)) {
				#LockScreen
				$Acl = Get-Acl $F3.FullName
				$Acl.Setaccessrule($Ar)
				Set-Acl $F3.FullName $Acl
				ForEach ($File in (Get-ChildItem $F3.FullName)) {
					$Acl = Get-Acl $File.fullname
					$Acl.Setaccessrule($Ar)
					Set-Acl $File.fullname $Acl
					Remove-Item $File.fullname -Force
				}
			}	
		}
	}
	#endregion Clear Lock Screen Cache	
	#High Res BG
	Set-Owner -Path ($env:windir + "\Web\4K\Wallpaper\Windows") -Recurse
	#Add Administrators with full control
	$files = get-childitem -Path ($env:windir + "\Web\4K\Wallpaper\Windows") 
	$user_account='Administrators'
	ForEach ($file in $files) {
		$Acl = Get-Acl $file.VersionInfo.FileName
		$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule($user_account,"FullControl", "None", "None", "Allow")
		$Acl.Setaccessrule($Ar)
		Set-Acl $file.VersionInfo.FileName $Acl
	}
	If (Test-Path ($LICache + "\" + $Custom_Wallpaper_SubFolder + "\" + $BackgroundFolder + "\4K\Wallpaper\Windows")) {	
		copy-item ($LICache + "\" + $Custom_Wallpaper_SubFolder + "\" + $BackgroundFolder + "\4K\Wallpaper\Windows\*.*") -Destination ($env:windir + "\Web\4K\Wallpaper\Windows") -force
    }
}
#============================================================================
#endregion Main Local Background
#============================================================================
#============================================================================
#region Main Local Setup Windows Time
#============================================================================
If (-Not $UserOnly) {
	Write-Host "Setting up Time . . . "
	#Disable Clients being NTP Servers
	Set-Reg "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\TimeProviders\NtpServer" "Enabled" 0 "DWORD"
	If ($Store) {
		net stop w32time | out-null
		W32tm /config /syncfromflags:manual /manualpeerlist:$NTP_ManualPeerList_Store | out-null
		w32tm /config /reliable:yes | out-null
		net start w32time | out-null
		w32tm /resync /rediscover | out-null
	} else {
		net stop w32time | out-null
		W32tm /config /syncfromflags:ALL /manualpeerlist:$NTP_ManualPeerList | out-null
		w32tm /config /reliable:yes | out-null
		net start w32time | out-null
		w32tm /resync /rediscover | out-null
	}
}
#============================================================================
#endregion Main Local Setup Windows Time
#============================================================================
#============================================================================
#region Main Local BGInfo
#============================================================================
If (-Not $UserOnly) {
	Write-Host "Setting up BGInfo . . . "
	If (Test-Path ($LICache + "\BgInfo")) {
		copy-item ($LICache + "\BgInfo") -Destination ($env:programfiles) -Force -Recurse
		Remove-Item ($env:programdata + "\Microsoft\Windows\Start Menu\Programs\StartUp\*BGInfo*.lnk")
		If ($Store) {
			copy-item ($env:programfiles + "\BgInfo\" + $BGInfo_StartupLink_Store) ($env:programdata + "\Microsoft\Windows\Start Menu\Programs\StartUp") -Force
		}else{
			copy-item ($env:programfiles + "\BgInfo\" + $BGInfo_StartupLink) ($env:programdata + "\Microsoft\Windows\Start Menu\Programs\StartUp") -Force
		}
	}
}
#============================================================================
#endregion Main Local BGInfo
#============================================================================
#============================================================================
#region Main Local Firewall Setup
#============================================================================
#Custom Software Firewall
If (-Not $UserOnly) {
	Write-Host "Setting up Firewall . . . "
	Remove-NetFirewallRule -Group (split-path $Custom_Software_Path -Leaf )
	If (Test-Path $Custom_Software_Path) {
		Write-Host ("Adding " + (split-path $Custom_Software_Path -Leaf ) + " to Firewall...") -foregroundcolor darkgray
		Get-ChildItem -Path $Custom_Software_Path -Filter *.exe -Recurse| ForEach {
			Write-Host ("`t Adding rule for: " + $_.Name) -foregroundcolor yellow
			New-NetFirewallRule -DisplayName $_.Name -Direction Inbound -Program $_.VersionInfo.FileName -Group (split-path $Custom_Software_Path -Leaf ) -Action Allow | out-null
		}
	}
	Write-Host "Disabling un-needed Firewall Rules . . . " -foregroundcolor darkgray
	Disable-NetFirewallRule -DisplayGroup "AllJoyn Router"
	Disable-NetFirewallRule -DisplayGroup "Cast to Device functionality"
	Disable-NetFirewallRule -DisplayGroup "cortana"	
	Disable-NetFirewallRule -DisplayGroup "Media Center Extenders"
	Disable-NetFirewallRule -DisplayGroup "Proximity Sharing"
	Disable-NetFirewallRule -DisplayGroup "Routing and Remote Access"
	Disable-NetFirewallRule -DisplayGroup "Wi-Fi Direct Network Discovery"
	Disable-NetFirewallRule -DisplayGroup "Windows Media Player Network Sharing Service"
	Disable-NetFirewallRule -DisplayGroup "Work or school account"
	Disable-NetFirewallRule -DisplayGroup "Your account"
	Disable-NetFirewallRule -DisplayGroup "iSCSI Service"
	Disable-NetFirewallRule -DisplayGroup "Xbox Game UI"
}
#============================================================================
#endregion Main Local Firewall Setup
#============================================================================
#============================================================================
#region Main Local Log and Performance Monitoring
#============================================================================


#============================================================================
#endregion Main Local Log and Performance Monitoring
#============================================================================
#============================================================================
#region Main Local FileShares
#============================================================================

#============================================================================
#endregion Main Local FileShares
#============================================================================
#============================================================================
#region Main Local RDP
#============================================================================
#RDP
If (-Not $UserOnly) {
	Write-Host "Enabling RDP . . . "
	Set-Reg "HKLM:\SYSTEM\CurrentControlSet\control\Terminal Server" "fDenyTSConnections " 0 "DWORD"
	# Set-Reg "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp" "UserAuthentication" 0 "DWORD"
}
#============================================================================
#endregion Main Local RDP
#============================================================================
#============================================================================
#region Main Local FileShares
#============================================================================

#============================================================================
#endregion Main Local FileShares
#============================================================================
#============================================================================
#region Main Local Setup Screen Saver
#============================================================================
If (-Not $UserOnly) {
	Write-Host "Setup Logon Screen Saver . . ."
	Set-Reg "HKU:\.DEFAULT\Control Panel\Desktop" "ScreenSaveActive" "1" "STRING"
	Set-Reg "HKU:\.DEFAULT\Control Panel\Desktop" "ScreenSaverIsSecure" "1" "STRING"
	Set-Reg "HKU:\.DEFAULT\Control Panel\Desktop" "ScreenSaveTimeOut" "600" "STRING"
	Set-Reg "HKU:\.DEFAULT\Control Panel\Desktop" "SCRNSAVE.EXE" "C:\Windows\system32\scrnsave.scr" "STRING"
}

#============================================================================
#endregion Main Local Setup Screen Saver
#============================================================================
#============================================================================
#region Main Local Microsoft Store
#============================================================================
#Disable MS Apps
If (-Not $UserOnly) {
	If ([environment]::OSVersion.Version.Major -ge 6) {
		#region Get list of currently installed and provisioned Appx packages
		$AllInstalled = Get-AppxPackage -AllUsers | Foreach {$_.Name}
		$AllProvisioned = Get-ProvisionedAppxPackage -Online | Foreach {$_.DisplayName}
		#endregion
		 
		#region Remove Appx Packages
		Write-Host "`n"
		Write-Host '#####################################' -ForegroundColor Green
		Write-Host -NoNewline '#' -ForegroundColor Green
		Write-Host -NoNewline '           '
		Write-Host -NoNewline "Appx Packages" -ForegroundColor White
		Write-Host -NoNewline '           '
		Write-Host '#' -ForegroundColor Green
		Write-Host '#####################################' -ForegroundColor Green
		Write-Host "`n"
		Foreach($Appx in $AllInstalled){
			$error.Clear()
			If(-Not $Keep.Contains($Appx)){
				Try{
					#Turn off the progress bar
					$ProgressPreference = 'silentlyContinue'
					Get-AppxPackage -Name $Appx | Remove-AppxPackage
					#Turn on the progress bar
					$ProgressPreference = 'Continue'
				}
				 
				Catch{
					$ErrorMessage = $_.Exception.Message
					$FailedItem = $_.Exception.ItemName
					Write-Host "There was an error removing Appx: $Appx"
					Write-Host $ErrorMessage
					Write-Host $FailedItem
				}
				If(!$error){
					Write-Host "Removed Appx: $Appx"
				}
			}
			Else{
				Write-Host "Appx Package is whitelisted: $Appx" -ForegroundColor yellow
			}
		}
		#endregion
		 
		#region Remove Provisioned Appx Packages
		Write-Host "`n"
		Write-Host '#####################################' -ForegroundColor Green
		Write-Host -NoNewline '#' -ForegroundColor Green
		Write-Host -NoNewline '     '
		Write-Host -NoNewline "Provisioned Appx Packages" -ForegroundColor White
		Write-Host -NoNewline '     '
		Write-Host '#' -ForegroundColor Green
		Write-Host '#####################################' -ForegroundColor Green
		Write-Host "`n"
		Foreach($Appx in $AllProvisioned){
			$error.Clear()
			If(-Not $Keep.Contains($Appx)){
				Try{
					Get-ProvisionedAppxPackage -Online | where {$_.DisplayName -eq $Appx} | Remove-ProvisionedAppxPackage -Online | Out-Null
				}
				 
				Catch{
					$ErrorMessage = $_.Exception.Message
					$FailedItem = $_.Exception.ItemName
					Write-Host "There was an error removing Provisioned Appx: $Appx"
					Write-Host $ErrorMessage
					Write-Host $FailedItem
				}
				If(!$error){
					Write-Host "Removed Provisioned Appx: $Appx"
				}
			}
			Else{
				Write-Host "Appx Package is whitelisted: $Appx" -ForegroundColor yellow
			}
		}
		#endregion

		Write-Host "`n"
	}
}
#============================================================================
#endregion Main Local Microsoft Store
#============================================================================
#============================================================================
#region Main Local Set OEM Info
#============================================================================
$Bios_Info = Get-CimInstance -ClassName Win32_BIOS
Write-Host "Setup System OEM Info . . ."
Set-Reg "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation" "Manufacturer" ((Get-CimInstance -ClassName Win32_ComputerSystem).Manufacturer) "String"
Set-Reg "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation" "Model" ((Get-CimInstance -ClassName Win32_ComputerSystem).model + " (Serial Number: " + (Get-CimInstance -ClassName Win32_BIOS).SerialNumber + ")") "String"
If (-Not (Test-Path ($env:windir + "\system32\oobe\info\"))) {
	New-Item -ItemType directory -Path ($env:windir + "\system32\oobe\info\")
}
Copy-Item  ($LICache + "\" + $Custom_Wallpaper_SubFolder + "\" + $Custom_OEM_Logo) -Destination ($env:windir + "\system32\oobe\info\" + $Custom_OEM_Logo ) -Recurse -Force
Set-Reg "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation" "Logo" ($env:windir + "\system32\oobe\info\" + $Custom_OEM_Logo ) "String"
#============================================================================
#region Main Local Set OEM Info
#============================================================================
#============================================================================
#region Main Local Disable Netbios
#============================================================================

If (-Not $UserOnly) {
	#https://community.spiceworks.com/topic/2010972-disable-netbios-over-tcp-ip-using-gpo-in-ad-environment 
	Write-Host ("Disabling Netbios...") -foregroundcolor darkgray
	$key = "HKLM:SYSTEM\CurrentControlSet\services\NetBT\Parameters\Interfaces"
	Get-ChildItem $key |
	foreach { Set-ItemProperty -Path "$key\$($_.pschildname)" -Name NetbiosOptions -Value 2 -Verbose}
}
#============================================================================
#endregion Main Local Disable Netbios
#============================================================================
#Disable Cast,WiFi
#============================================================================
#region Main Local Load Local GPO
#============================================================================
If (-Not $UserOnly) {
	If ([environment]::OSVersion.Version.Major -ge 10) {
		If ((Test-Path ($LICache + "\LGPO.EXE")) -and (Test-Path ($LICache + '\Security Templates\Windows10Ent'))) {
			$process = Start-Process -FilePath ($LICache + "\LGPO.EXE") -ArgumentList ('/g "' + $LICache + '\Security Templates\Windows10Ent"') -PassThru -NoNewWindow -Wait
			Write-Host
		}
	}
}

#============================================================================
#endregion Main Local Load Local GPO
#============================================================================
#============================================================================
#region Main Local Cleanup
#============================================================================
If (Test-Path "PSRemote:\") {
	Remove-PSDrive -Name "PSRemote"
}

$sw.Stop()
Write-Host ("Script took: " + (FormatElapsedTime($sw.Elapsed)) + " to run.")
#============================================================================
#endregion Main Local Cleanup
#============================================================================
#############################################################################
#endregion Main
#############################################################################
If (-Not [string]::IsNullOrEmpty($LogFile)) {
	Stop-Transcript
}
