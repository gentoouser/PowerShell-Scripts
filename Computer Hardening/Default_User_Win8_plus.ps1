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
	
.PARAMETER Profiles
	Array of users to lockdown. If Store is enabled all store Window users will be added to the array and locked down.	
.PARAMETER LICache
	Location of the local configureation files cache.	
.PARAMETER RemoteFiles
	Network path of configureation files are to copy down. 	
.PARAMETER CARoot
	Cert file of Domain Root CA.
.PARAMETER CAInter
	Cert file of Domain Intermediate  CA.
.PARAMETER CSCert
	Cert file of Code Signing cert.
.PARAMETER StartLayoutXML
	XML files that setup the defualt startmenu and task bar items.
.PARAMETER BackgroundFolder
	Folder name of where customized default windows backgrounds are.	
.PARAMETER User
	Username for network configureation share.
.PARAMETER Password
	Password that goes with useranme for network configureation share.
.PARAMETER Store
	Enables more locked down of users and creates store Local Windows accounts.
.PARAMETER LockedDown
	Lock down user accounts more
.PARAMETER UserOnly
	Sets user settings only and no machine settings
.PARAMETER NoCacheUpdate
	Skip updating local cache.
.PARAMETER AllowClientTLS1
	Enables Computer to go to TLS 1.0 and TLS 1.1 sites.
.PARAMETER NoOEMInfo
	Keeps from reseting the OEM Info.	
.PARAMETER OEMInfoAddSerial
	Added Serial number to the System Preferences.
.PARAMETER NoBgInfo
	Does not setup BGInfo to launch at startup.
.PARAMETER IPv6
	Keeps IPv6 enabled; otherwise IPv6 will be disabled. 
.EXAMPLE
   & Default_User_Win8.1_+.ps1 -AllowClientTLS1
.EXAMPLE
	powershell -executionpolicy unrestricted -file .\Default_User_Win8.1_+.ps1 -Store -AllowClientTLS1 -Profile '"Default","User"' -StartLayoutXML "Win10_VDI.xml" -BackgroundFolder "Store workstations"
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
	* Version 2.0.14 - Fixed Hard Coded Path for LGPO.exe. Update Logic for Store and TLS 1.0 and TLS 1.1. Fixed showing only select Contol Panel items. Added Chrome $ChromeURLBlackList 
	* Version 2.0.15 - Hiding User Accounts from Logon Screen. Test if users need to be created before changing local policy. Added make sure that SHA is enabled when TLS 1.0 is enabled. Hide VMWare Tools
	* Version 2.0.16 - FortiClient copy and run RemoveFCTID.exe for System-Prep. Created new fuction to copy only changed files.
	* Version 2.0.17 - Updated Windows Store Apps White-List. Fixed issue where new files were not copying to Local Cache.
	* Version 2.0.18 - Found how to use RoboCopy for Local Cache update. Adobe Reader Accept EULA.
	* Version 2.1.00 - Added ablity to disable windows Features. Removed need for LGPO. Remove OneDrive. Add Internet Explorer on All Users Desktop
	* Version 2.1.01 - Updated IE Trusted Sites. Updated OneDrive Removeal. Added more services to Automatic to fix issue with SysPrep. 
	* Version 2.1.02 - Fixed issues with Default chrome profile. Fixed issues with automatic services. Added Switch to ignore NoBgInfo. 
	* Version 2.1.03 - .Net Settings for TLS 1.2. Log path update. Update Setting to allow MS to update other products too. Added Windows version checking.
	* Version 2.1.04 - Fix Issue with RemoveFCTID.exe being in wrong place. Added more insanity checks. 
	* Version 2.1.05 - Stopped removing "Microsoft.Windows.Cortana" and "Microsoft.Windows.ShellExperienceHost" due to Start Menu Breaking. Updated WallpaperStyle. Change it so Remote files are where the script is launched from. Disable more visual effects.
	* Version 2.1.06 - More Tweaks and test to see of we are running in VM.
	* Version 2.1.07 - Added ablity to change registry permissions. Fix bug with BGInfo Shortcut. Added $ScriptDateValue to know when the script was ran. Create Shortcut on Desktop for FortiClient ID Remover
	* Version 2.1.08 - Fix Windows Update options. Tweaks to speed up running -store for a second time. If filename has store in it enable store switch. Added switch for OEMInfoAddSerial. Disabled register domain join computers as devices for Azure. Disable Secure Screen saver for store users. Disables "Recently added" Apps List on the Start Menu for Locked down users. Disable all store user accounts after profile is created. Added IPv6 switch; all IPv6 will be disabled by default. Set "Power Button" to shutdown. Disable Windows 10 managed default printer. Disable Network Discovery network rules. Auto Start Custom exe for all "Window users"
	* Version 2.1.09 - Fix Cache Issue. Fixed issue with $Custom_Software_Exec shortcut creation issue. Updated example code to convirt string ot array for powershell.exe launch. 
	* Version 2.1.10 - Add Ping firewall rule. Fix issue with VM only settings. Disabled lock screen for store users.
	* Version 2.1.11 - Enable FontSmoothing. Disable SNMP by default. Configure SNMP Settings. Fixed Custom software auto launching link issue. 
	* Version 2.1.12 - More Disabling of the Lockscreen for stores. Added Custiomized Win+X setting for stores. Fix of Control panel for stores. Copy Icons to all users desktop if there is a Desktop folder in the LICache. Added ScheduledJob on startup to clean temp files.
	* Version 2.1.13 - Copy Custom Icons. Enable PS/2 Mouse. Enable SNMP Legacy mode for Printing. Fix Power button issue.
	* Version 2.1.14 - Fix for screen output. Added check for "bowser" service; disabling this servcie stops SMB. Lockdown VMWare Horizon; Fixed SSLCipherList issue. Commented alot of code for trobleshooting. 
	* Version 2.1.15 - Fixed access admin UNC for local users. Fixed issue where deny files were still showing up for users.
	* Version 2.1.16 - Fixed Issues for Windows Store and Settings Apps not launching.
	* Version 2.1.17 - Chrome GPO updates; Fixed OWA File Upload bug
	* Version 2.1.18 - Random complex 120 chr. password for Store users. Reset Local Administrator password to random password and disables the account. Disable Windows Feature Upgrade
	* Version 2.1.19 - Updated Comments on what each registry entry does. Added reg. to stop download of printer metadata. Added code to deal with reseting profiles. Updated User Theme and Accent color settings. Set account lockout. Fixed Password generation bug. Fixed bug where account was not enabled when trying to recoreate user profile. 
	#>
#Requires -Version 5.1 -PSEdition Desktop
PARAM (
	[array]$Profiles  	  		= @("Default"),	
	[string]$LICache	  		= "C:\IT_Updates",
	[string]$RemoteFiles  		= (Split-Path -Parent -Path $MyInvocation.MyCommand.Definition),
	[string]$CARoot				= "RootCA.cer",
	[string]$CAInter			= "InterCA.cer",
	[string]$CSCert				= "Code Signing.cer",
	[string]$StartLayoutXML		= "Win10_VDI.xml",
	[String]$BackgroundFolder 	= "Store workstations",
	[String]$User		    	= $null,
	[String]$Password	    	= $null,
	[switch]$Store	  	  		= $false,
	[switch]$LockedDown	  		= $false,
	[switch]$UserOnly			= $false,
	[switch]$NoCacheUpdate		= $false,
	[switch]$AllowClientTLS1	= $false,
	[switch]$NoOEMInfo			= $false,
	[switch]$OEMInfoAddSerial	= $false,
	[switch]$NoBgInfo			= $false,
	[switch]$IPv6				= $false
)
#Force Running Script as Admin
If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{   
$arguments = "& '" + $myinvocation.mycommand.definition + "'"
Start-Process powershell -Verb runAs -ArgumentList $arguments
Break
}
#Fix issue for services
Set-Location -Path "\"
$ScriptVersion = "2.1.19"
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
$IE_Cache_Size = 1024
$VMware_Horizon_Server = "https://horizon.github.com"
$VMware_Horizon_NetBIOSDomain = "github"
$VMware_Horizon_SSLCipherList = "TLSv1.2:!aNULL:!SHA:kECDH+AESGCM:ECDH+AESGCM:RSA+AESGCM:kECDH+AES:ECDH+AES:RSA+AES"
$Custom_Software_Path = (${env:ProgramFiles(x86)} + "\Custom_app")
$Custom_Software_Exec = "Custom.exe"
$Custom_Wallpaper_SubFolder = "Wallpapers"
$Custom_User_Account_Pictures_SubFolder = ($Custom_Wallpaper_SubFolder + "\User Account Pictures")
$Custom_OEM_Logo = "LOGO_OEM.bmp"
$Custom_Icon_Path = (${env:ProgramFiles(x86)} + "\github")
$NTP_ManualPeerList = "time.nist.gov,0x08 north-america.pool.ntp.org,0x08"
$NTP_ManualPeerList_Store = $NTP_ManualPeerList
$BGInfo_StartupLink = "Bginfo Slient Start x64.lnk"
$BGInfo_StartupLink_Store = "Bginfo Slient Start VDI.lnk"
$SettingsPageVisibility = "showonly:printers;defaultapps;display;mousetouchpad;network-ethernet;notifications;usb"
$ChromeBaseZip = "Google_Profile_Base.zip"
#Easy way to make Custom Settings: https://www.howtogeek.com/113570/how-to-edit-the-winx-menu-in-windows-8-using-a-free-tool/
$WinXZip = "WinX.zip"
$ChromeDelegateWhiteList = "https://*.git.com"
$ScriptVersionKey = "Git Hub" 
$ScriptVersionValue = "Security Hardening Version"
$ScriptDateValue = "Security Hardening Date"
$SNMPValue = "Public"
#Versions of Adobe Reader to setup for.
$ARV = ("11.0","2005","DC")
$UserRange = 1..20
#region IE Domain Settings
#https://support.microsoft.com/en-us/help/182569/internet-explorer-security-zones-registry-entries-for-advanced-users
# Value    Setting
# ------------------------------
# 0        My Computer
# 1        Local Intranet Zone
# 2        Trusted sites Zone
# 3        Internet Zone
# 4        Restricted Sites Zone
$ZoneMap = @(
    New-Object PSObject -Property @{Site = "patchmypc.net";  Protocol = "https"; Zone = 2}
    New-Object PSObject -Property @{Site = "microsoft.com"; Protocol = "https"; Zone = 2}
    New-Object PSObject -Property @{Site = "microsoft.com"; Protocol = "http"; Zone = 2}
    New-Object PSObject -Property @{Site = "microsoft.com\download"; Protocol = "http"; Zone = 2}
    New-Object PSObject -Property @{Site = "microsoft.com\download"; Protocol = "https"; Zone = 2}
    New-Object PSObject -Property @{Site = "update.microsoft.com"; Protocol = "http"; Zone = 2}
    New-Object PSObject -Property @{Site = "update.microsoft.com"; Protocol = "https"; Zone = 2}
)
#endregion IE Domain Settings
#region Registry Permissions
# Options:
#	FullControl, ReadKey, SetValue, CreateSubKey, Delete
$RegPerms = @(
	New-Object PSObject -Property @{Hive = "HKEY_LOCAL_MACHINE"; Key = "SOFTWARE\WOW6432Node\Custom\app";  User = "Users"; Perm = "FullControl"; Action = "Allow"}
)
#endregion Registry Permissions
#region RoboCopy Options
$LICRoboCopyOptions = @(
	"/E"
	"/R:3"
	"/W:3"
	"/NDL"
	"/NFL"
	"/NJH"
	"/XD Logs"
)
#endregion RoboCopy Options
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
#region Hide Accounts from Logon screen
$HideAccounts = @(
	"Admin"
	"Administrator"
	"ASPNET"
	"Guest"
)
#endregion Hide Accounts from Logon screen
#region Show Control Panel Items
$ShowOnlyCPL = @(
	"Devices and Printers"
	"Microsoft.DevicesAndPrinters"
	"Default Programs"
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
#region Disable Windows Features
$RemoveFeatures = @(
	"ClientForNFS-Infrastructure"
	"Containers"
	"DirectoryServices-ADAM-Client"
	"Hello.Face.Resource"
	"IIS-FTPServer"
	"IIS-WebServer"
	"IIS-WebServerRole"
	"LegacyComponents"
	"MSMQ-Container"
	"MSMQ-Server"
	"Microsoft-Hyper-V"
	"Microsoft-Hyper-V-All"
	"Microsoft-Windows-Subsystem-Linux"
	"MicrosoftWindowsPowerShellV2"
	"MicrosoftWindowsPowerShellV2Root"
	"OpenSSH"
	"Printing-Foundation-InternetPrinting-Client"
	"Printing-Foundation-LPDPrintService"
	"Printing-Foundation-LPRPortMonitor"
	"RasRip"
	"SMB1Protocol"
	"SMB1Protocol-Client"
	"SMB1Protocol-Server"
	"SNMP"
	"ServicesForNFS-ClientOnly"
	"SimpleTCP"
	"TFTP"
	"TelnetClient"
	"WMISnmpProvider"
	"Windows-Defender-Default-Definitions"
)
#endregion Disable Windows Features
#region Disable Known Folders
$DisableKnownFolders = @(
	"Videos"
	"Music"
	"Pictures"
	"Network"
)
#endregion Disable Known Folders
#region Chrome URL BlackList
$ChromeURLBlackList = @(
	"file://*"
	"chrome://settings/clearBrowserData"
	"chrome://settings/clearBrowserData-frame"
	"chrome://extensions"
)
#endregion Chrome URL BlackList
#endregion +++++++ Company Specific Settings +++++++#

#region Services	
$DisableServices = @(
	"AdobeARMservice"							# Adobe Acrobat Update Service
	"AMD External Events Utility"				# AMD External Events Utility
	"AJRouter"									# AllJoyn Router Service
	"ALG"										# Application Layer Gateway Service
	"Browser"									# Computer Browser
	#"DeviceAssociationService"					#Device Association Service	#### Causes Log-on Delays
	"diagnosticshub.standardcollector.service"  # Microsoft (R) Diagnostics Hub Standard Collector Service
	"diagsvc"									# Diagnostic Execution Service
	"DiagTrack"                              	# Diagnostics Tracking Service
	#"dmwappushservice"                       	# WAP Push Message Routing Service (see known issues) ####Breaks SysPrep
	"DPS"										# Diagnostic Policy Service
	"FAX"										# Fax Service
	#"FDResPub"									# Function Discovery Resource Publication Service ### Errors for HomeGroup
	#"HomeGroupListener"                      	# HomeGroup Listener
	#"HomeGroupProvider"                      	# HomeGroup Provider
	"HvHost"									# HV Host Service
	"irmon"										# Infrared monitor service
	"lfsvc"                                  	# Geolocation Service
	"icssvc"									# Windows Mobile Hotspot Service
	"IpxlatCfgSvc"								# IP Translation Configuration Service
	"NaturalAuthentication"						# Natural Authentication ## Face login
	#"lmhosts"									# TCP/IP NetBIOS Helper	#####Breaks SMB 
	"MapsBroker"                             	# Downloaded Maps Manager
	"MSiSCSI"									# Microsoft iSCSI Initiator Service
	"MyWiFiDHCPDNS"								# Wireless PAN DHCP Server
	"NetTcpPortSharing"                      	# Net.Tcp Port Sharing Service
	#"netprofm"									# Network List Service	### Event log errors
	"p2pimsvc"									# Peer Networking Identity Manager
	"p2psvc"									# Peer Name Resolution Protocol
	"PeerDistSvc"								# BranchCache ## Used by Windows Update for download sharing.
	"PhoneSvc"									# Phone Service
	"PNRPAutoReg"								# PNRP Machine Name Publication Service
	"PNRPsvc"									# Peer Name Resolution Protocol
	"QWAVE"										# Quality Windows Audio Video Experience Service
	"RemoteAccess"                           	# Routing and Remote Access
	"RemoteRegistry"                         	# Remote Registry
	"RetailDemo"								# Retail Demo Service
	#"RmSvc"									# Radio Management Service ##### Breaks Wi-Fi
	"RpcLocator"								# Remote Procedure Call (RPC) Locator
	#"RSoPProv"									# Resultant Set of Policy Provider
	"SEMgrSvc"									# Payments and NFC/SE Manager
	"SensorDataService"							# Sensor Data Service
	"SensrSvc"									# Sensor Service
	"SharedAccess"                           	# Internet Connection Sharing (ICS)
	"smphost"									# Microsoft Storage Spaces SMP Service
	"SNMP"										# SNMP Service
	"SNMPTRAP"									# SNMP Trap
	"SSDPSRV"									# SSDP Discovery	#####Breaks SMB
	"svsvc"										# Spot Verifier Service
	"Themes"									# Themes
	"TrkWks"                                 	# Distributed Link Tracking Client
	"upnphost"									# UPnP Device Host    #####Breaks SMB
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
	"WcsPlugInService"							# Windows Color System Service
	"wcncsvc"									# Windows Connect Now - Config Registrar Service
	"WerSvc"									# Windows Error Reporting Service
	"WFDSConMgrSvc"								# Wi-Fi Direct Services Connection Manager Service
	"wisvc"										# Windows insider program
	#"WlanSvc"                               	# WLAN AutoConfig ##### Breaks Wi-Fi
	"wlidsvc"									# Microsoft Account Sign-in Assistant
	"WMPNetworkSvc"                          	# Windows Media Player Network Sharing Service
	#"wscsvc"                                	# Windows Security Center Service
	#"WSearch"                               	# Windows Search
	"XblAuthManager"                        	# Xbox Live Auth Manager
	"XblGameSave"                            	# Xbox Live Game Save Service
	#"xbgm"										# Xbox Game Monitoring Service
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
	"AppXSVC"									# AppX Deployment Service (AppXSVC) ## Windows Store Requirement
	"Nameiphlpsvc"								# IP Helper
	"wuauserv"									# Windows Update
	"LicenseManager"							# Windows License Manager Service ## Windows Store Requirement
)
$AutomaticServices = @(
	"W32Time"									#Windows Time
	"lmhosts"									# TCP/IP NetBIOS Helper	#####Breaks SMB 
	"SSDPSRV"									# SSDP Discovery	#####Breaks SMB
	"upnphost"									# UPnP Device Host    #####Breaks SMB
	"WlanSvc"                               	# WLAN AutoConfig ##### Breaks Wi-Fi
	"dmwappushservice"                       	# WAP Push Message Routing Service (see known issues) #### Breaks SysPrep
	"bits"										# Background Intelligent Transfer Service ### For Windows Update
	"cryptsvc"									# Cryptographic Services ### For Windows Update
	"trustedinstaller"							# Windows Modules Installer ### For Windows Update
	"StorSvc"									# Storage Service ## Windows Store Requirement
)

#endregion Services	
#region Microsoft Store
	#Windows 10 Rev. 1803 WhiteList
	#APSS to Keep:
	$Keep =  @(
	"1527c705-839a-4832-9118-54d4Bd6a0c89"
	"E2A4F912-2574-4A75-9BB0-0D023378592B"
	"F46D4000-FD22-4DB4-AC8E-4E1DDDE828FE"
	"InputApp"
	"Microsoft.AAD.BrokerPlugin"
	"Microsoft.AccountsControl"
	"Microsoft.Advertising"
	"Microsoft.Advertising.Xaml"
	"Microsoft.Appconnector"
	"Microsoft.AsyncTextService"
	"Microsoft.BingWeather" 
	"Microsoft.BioEnrollment"
	"Microsoft.CredDialogHost"
	"Microsoft.DesktopAppInstaller"
	"Microsoft.ECApp"
	"Microsoft.LockApp"
	"Microsoft.MSPaint"
	"Microsoft.Microsoft3DViewer"
	"Microsoft.MicrosoftEdge"
	"Microsoft.MicrosoftEdgeDevToolsClient"
	"Microsoft.MicrosoftStickyNotes" 
	"Microsoft.NET.Native.Framework"
	"Microsoft.NET.Native.Runtime"
	"Microsoft.Office.OneNote"
	"Microsoft.PPIProjection"
	#"Microsoft.People"
	"Microsoft.Services.Store.Engagement"
	#"Microsoft.SkypeApp"
	"Microsoft.StorePurchaseApp"
	"Microsoft.VCLibs"
	"Microsoft.VCLibs.UWPDesktop"
	"Microsoft.UI.Xaml"
	"Microsoft.Wallet"
	"Microsoft.Win32WebViewHost"
	"Microsoft.windowscommunicationsapps" ##Breaks Microsoft Accounts from UWP
	"Microsoft.Windows.Apprep.ChxApp"
	"Microsoft.Windows.AssignedAccessLockApp"
	"Microsoft.Windows.CapturePicker"
	"Microsoft.Windows.CloudExperienceHost"
	"Microsoft.Windows.ContentDeliveryManager"
	"Microsoft.Windows.Cortana" ##Breaks Start Menu if not Installed
	"Microsoft.Windows.HolographicFirstRun"
	"Microsoft.Windows.OOBENetworkCaptivePortal"
	"Microsoft.Windows.OOBENetworkConnectionFlow"
	"Microsoft.Windows.ParentalControls"
	"Microsoft.Windows.PeopleExperienceHost"
	"Microsoft.Windows.Photos"
	"Microsoft.Windows.PinningConfirmationDialog"
	"Microsoft.Windows.SecHealthUI"
	"Microsoft.Windows.SecondaryTileExperience"
	"Microsoft.Windows.SecureAssessmentBrowser"
	"Microsoft.Windows.ShellExperienceHost" ##Breaks Start Menu if not Installed
	"Microsoft.WindowsAlarms"
	"Microsoft.WindowsCalculator"
	"Microsoft.WindowsCamera"
	#"Microsoft.WindowsFeedbackHub"
	#"Microsoft.WindowsMaps"
	"Microsoft.WindowsStore"
	#"Microsoft.Xbox.TCUI"
	#"Microsoft.XboxApp"
	"Microsoft.XboxGameCallableUI"
	"Microsoft.XboxIdentityProvider"
	"Windows.CBSPreview"
	"Windows.MiracastView"
	"Windows.PrintDialog"
	"c5e2524a-ea46-4f67-841f-6a9465d9d515"
	"windows.immersivecontrolpanel"
    "winstore"
	)
#endregion Microsoft Store

$LogFile = ($LICache + "\Logs\" + `
		   $MyInvocation.MyCommand.Name + "_" + `
		   $env:computername + "_" + `
		   (Get-Date -format yyyyMMdd-hhmm) + ".log")
$sw = [Diagnostics.Stopwatch]::StartNew()
$IsVM = $False
$HKEY = "HKU\DEFAULTUSER"
# Some paths that get used more than once
$ContentDeliveryPath = ($HKEY.replace("HKU\","HKU:\") + "\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager")
$HKEYWE = ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Explorer")
$HKEYIE = ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Internet Explorer")
$HKEYIS = ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Internet Settings")
$WindowsSearchPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search"
$VMWare_Horzion_Key = 'HKLM:\SOFTWARE\VMware, Inc.\VMware VDM\Client'
#$UACPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
$HKLWE = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer"
$HKAR = "HKLM:\SOFTWARE\Policies\Adobe\Acrobat Reader"
$HKSCH = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL"
$UsersProfileFolder = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\' -Name "ProfilesDirectory").ProfilesDirectory
$ProfileList =  New-Object System.Collections.ArrayList
$WScriptShell = New-Object -ComObject ("WScript.Shell")
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
#Store Setup
If ($MyInvocation.MyCommand.Name -match "store") {
	$Store = $True
}
If ($Store) {
	$LockedDown = $True
}
#Add Registry Hives
New-PSDrive -PSProvider Registry -Name HKU -Root HKEY_USERS -erroraction 'silentlycontinue' | Out-Null
New-PSDrive -PSProvider Registry -Name HKCR -Root HKEY_CLASSES_ROOT -erroraction 'silentlycontinue' | Out-Null
#Share Setup
if ( $User -and $Password) {
	$Credential = New-Object System.Management.Automation.PSCredential ($User, (ConvertTo-SecureString $Password -AsPlainText -Force))
}
#Powershell.exe launch cleanup
If ($Profiles[0].Contains(",")) {
	#Setup ProfileList
	ForEach ($Profile in ($Profiles[0].split(","))) {
		If ($Profile) {
			$ProfileList.Add($Profile)
			$HideAccounts += $Profile
		}
	}
}else{
	#Setup ProfileList
	ForEach ($Profile in $Profiles) {
		If ($Profile) {
			$ProfileList.Add($Profile)
			$HideAccounts += $Profile
		}
	}
}
#Load Password DLL
Add-Type -AssemblyName System.web
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
function Set-KeyOwnership {
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
    ForEach ($i in $privileges.Values) {
        $null = $ntdll::RtlAdjustPrivilege($i, 1, 0, [ref]0)
    }

    function Set-KeyOwnership {
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
            ForEach($subKey in $regKey.OpenSubKey('').GetSubKeyNames()) {
                Set-KeyOwnership $rootKey ($key+'\'+$subKey) $sid $recurse ($recurseLevel+1)
            }
        }
    }

    Set-KeyOwnership $rootKey $key $sid $recurse
}
function Get-CurrentUserSID {            
	[CmdletBinding()]            
	param(            
	)            
	#Source: https://techibee.com/powershell/find-the-sid-of-current-logged-on-user-using-powershell/2638
	Add-Type -AssemblyName System.DirectoryServices.AccountManagement            
	return ([System.DirectoryServices.AccountManagement.UserPrincipal]::Current).SID.Value            
}
function Set-Reg {
	[CmdletBinding()] 
Param 
( 
	[Parameter(Mandatory=$true,Position=1,HelpMessage="Path to Registry Key")][string]$regPath, 
	[Parameter(Mandatory=$true,Position=2,HelpMessage="Name of Value")][string]$name,
	[Parameter(Mandatory=$true,Position=3,HelpMessage="Data for Value")]$value,
	[Parameter(Mandatory=$true,Position=4,HelpMessage="Type of Value")][ValidateSet("String", "ExpandString","Binary","DWord","MultiString","Qword","Unknown",IgnoreCase =$true)][string]$type 
) 
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

If($type -eq "Binary" -and $value.GetType().Name -eq "String" -and $value -match ",") {
	$value = [byte[]]($value -split ",")
}

New-ItemProperty -Path $regPath -Name $name -Value $value -PropertyType $type -Force | Out-Null							   
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
	$TargetObject = $QuickAccess.Namespace("shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}").Items() | Where-Object {$_.Path -eq "$Path"} 
	If ($Action -eq "Pin") 
		{ 
			If (-Not ([string]::IsNullOrEmpty($TargetObject))) 
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
			If (-Not ([string]::IsNullOrEmpty($TargetObject)))
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
Function Get-MachineType { 
	<# 
	.Synopsis 
	   A quick function to determine if a computer is VM or physical box. 
	.DESCRIPTION 
	   This function is designed to quickly determine if a local or remote 
	   computer is a physical machine or a virtual machine. 
	.NOTES 
	   Created by: Jason Wasser 
	   Modified: 9/11/2015 04:12:51 PM   
	 
	   Changelog:  
		* added credential support 
	 
	   To Do: 
		* Find the Model information for other hypervisor VM's like Xen and KVM. 
	.EXAMPLE 
	   Get-MachineType 
	   Query if the local machine is a physical or virtual machine. 													   
	.LINK 
	   https://gallery.technet.microsoft.com/scriptcenter/Get-MachineType-VM-or-ff43f3a9 
	#> 
    [CmdletBinding()] 
    [OutputType([int])] 
    Param 
    ( 
													 
    ) 
    Begin { 
    } Process { 
		try { 
			#$hostdns = [System.Net.DNS]::GetHostEntry($Computer) 
			If (Get-Command Get-CimInstance -errorAction SilentlyContinue) {
				$ComputerSystemInfo = Get-CimInstance -Class Win32_ComputerSystem  -ErrorAction Stop 
			} Else {
				$ComputerSystemInfo = Get-WmiObject -Class Win32_ComputerSystem  -ErrorAction Stop 
			} 
							
			switch -wildcard ($ComputerSystemInfo.Model) { 	
				# Check for Hyper-V Machine Type 
				"*Virtual Machine*" { 
					$MachineType="VM" 
					} 
				# Check for VMware Machine Type 
				"*VMware*" { 
					$MachineType="VM" 
					} 
				# Check for Oracle VM Machine Type 
				"*VirtualBox*" { 
					$MachineType="VM" 
					} 
				# Check for Xen 
				# I need the values for the Model for which to check. 

				# Check for KVM 
				# I need the values for the Model for which to check. 

				# Otherwise it is a physical Box 
				default { 
					$MachineType="Physical" 
					} 
				} 
				
			# Building MachineTypeInfo Object 
			$MachineTypeInfo = New-Object -TypeName PSObject -Property ([ordered]@{ 
				ComputerName=$ComputerSystemInfo.Name
				Type=$MachineType 
				Manufacturer=$ComputerSystemInfo.Manufacturer 
				Model=$ComputerSystemInfo.Model 
				}) 
			$MachineTypeInfo 
		} catch [Exception] { 								
			Write-Output "Error`: $($_.Exception.Message)" 				  
		} 
	} End { 
 
    } 
}
function Test-RegistryKeyValue {
    <#
    .SYNOPSIS
    Tests if a registry value exists.

    .DESCRIPTION
    The usual ways for checking if a registry value exists don't handle when a value simply has an empty or null value.  This function actually checks if a key has a value with a given name.
	Source: https://stackoverflow.com/questions/5648931/test-if-registry-value-exists
    .EXAMPLE
    Test-RegistryKeyValue -Path 'hklm:\Software\Carbon\Test' -Name 'Title'

    Returns `True` if `hklm:\Software\Carbon\Test` contains a value named 'Title'.  `False` otherwise.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]
        # The path to the registry key where the value should be set.  Will be created if it doesn't exist.
        $Path,

        [Parameter(Mandatory=$true)]
        [string]
        # The name of the value being set.
        $Name
    )

    if( -not (Test-Path -Path $Path -PathType Container) ) {
        return $false
    }

    $properties = Get-ItemProperty -Path $Path 
    if( -not $properties ) {
        return $false
    }
    $member = Get-Member -InputObject $properties -Name $Name
    if( $member ) {
        return $true
    } else {
        return $false
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
#Get Where we are running
If ((Get-MachineType).type -ne "Physical") {
	$IsVM = $True
}
#Skip updating local cache
If (-Not $NoCacheUpdate) {
	#Setup Local Install Cache
	If (-Not( Test-Path $LICache)) {
		write-host ("Creating Local Install cache: " + $LICache)
		New-Item -ItemType directory -Path $LICache | Out-Null
		$Acl = Get-Acl $LICache
		$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule('Users', "FullControl", "ContainerInherit, ObjectInherit", "None", "Allow")
		$Acl.Setaccessrule($Ar) | Out-Null
		Set-Acl $LICache $Acl | Out-Null
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
		write-host ("Copying to Local Install cache: " + $LICache + " Please Wait . . .")
		$CurrentScriptUTC = $(Get-Item $MyInvocation.MyCommand.Definition).LastWriteTimeUtc		
		#Copy-Item  "PSRemote:\*" -Destination $LICache -Recurse -Force
		#Copy-Newer -Source "PSRemote:\" -Destination $LICache -Exclude @("logs") -Overwrite
		$temp = @((Get-Item "PSRemote:\").FullName,$LICache)
		$temp += $LICRoboCopyOptions
		$process = Start-Process -FilePath ("robocopy.exe") -ArgumentList $temp -PassThru -NoNewWindow -wait

		If ($(Get-Item ($LICache + "\" + $MyInvocation.MyCommand.Name)).LastWriteTimeUtc -gt $CurrentScriptUTC) {
			#write-host ("Starting newer copy of script...")
			Stop-transcript
			#Need to fix getting the correct parameters sent to the new script instance
			#$&$MyInvocation.MyCommand.Definition  $MyInvocation.MyCommand.Parameters 
			exit
		}
	}
}
#Harden Permission on the c:\
# Remove user the rights to create and modify data on the root of the c:\ drive.
If (-Not $UserOnly) {
	write-host ("Hardening Permissions: " + ($env:systemdrive + "\"))
	$acl = Get-Acl ($env:systemdrive + "\")
	If ($acl.Access | Where-Object { $_.IdentityReference -eq "NT AUTHORITY\Authenticated Users" }) {
		$usersid = New-Object System.Security.Principal.Ntaccount ("NT AUTHORITY\Authenticated Users")
		$acl.PurgeAccessRules($usersid) | Out-Null
		$acl | Set-Acl ($env:systemdrive + "\") | Out-Null
	}
	If (Test-Path $Custom_Software_Path) {
		write-host ("Setting Permissions: " + $Custom_Software_Path)
		$acl = Get-Acl $Custom_Software_Path
		If (-Not ($acl.Access | Where-Object { $_.IdentityReference -eq "BUILTIN\Users" -and $_.FileSystemRights -eq "FullControl"})) {
			$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule('Users', "FullControl", "ContainerInherit, ObjectInherit", "None", "Allow")
			$Acl.Setaccessrule($Ar) | Out-Null
			Set-Acl $Custom_Software_Path $Acl | Out-Null
		}
	}
	If (Test-Path (Split-path -path $Custom_Software_Path -Parent) {
		write-host ("Setting Permissions: " + (Split-path -path $Custom_Software_Path -Parent))
		$acl = Get-Acl (Split-path -path $Custom_Software_Path -Parent)
		If (-Not ($acl.Access | Where-Object { $_.IdentityReference -eq "BUILTIN\Users" -and $_.FileSystemRights -eq "FullControl"})) {
			$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule('Users', "FullControl", "ContainerInherit, ObjectInherit", "None", "Allow")
			$Acl.Setaccessrule($Ar) | Out-Null
			Set-Acl (Split-path -path $Custom_Software_Path -Parent) $Acl | Out-Null
		}
}
#============================================================================
#region Main Local Start Menu and Taskbar Settings
#============================================================================
#Import Start Menu Layout
If (-Not $UserOnly) {
	If ([environment]::OSVersion.Version.Major -ge 10) {
		If (Test-Path ($LICache + "\" + $StartLayoutXML)) {
			write-host ("Setting Taskbar and Start Menu: " + ($LICache + "\" + $StartLayoutXML))
			Import-StartLayout -LayoutPath ($LICache + "\" + $StartLayoutXML) -MountPath ($env:systemdrive + "\") | Out-Null
		}
	}
}
#============================================================================
#endregion Main Local Start Menu and Taskbar Settings
#============================================================================
#Create Local Store users
If ($Store) {
	#Testing if we need to create any accounts
	Write-Host 'Testing for exsiting "Window" users . . .'
	$CreateUsers = $False
	ForEach ( $i in $UserRange) {	
		If ($i) {
			If (-Not (Get-LocalUser -Name ("Window" + $i) -erroraction 'silentlycontinue')) {
				$CreateUsers = $True
			}else{
				If (Test-Path ($UsersProfileFolder + "\Window" + $i)) {
					Write-Host ("`tAdding to Profile List: " + ($UsersProfileFolder + "\Window" + $i))
					$ProfileList.Add(("Window" + $i).ToLower()) | Out-Null
					$HideAccounts += ("Window" + $i).ToLower()
				}else{
					$CreateUsers = $True
				}
			}
		}
	}
	#Disable Password Requirements for creating new accounts
	If ($CreateUsers) {
		#secedit /export /cfg c:\secpol.cfg
		# Write-Host 'Changing Password Policy to create "Window" users . . .'
		# $process = Start-Process -FilePath ("secedit") -ArgumentList @("/export","/cfg","c:\secpol.cfg") -PassThru -NoNewWindow -wait
		# (Get-Content C:\secpol.cfg).replace("PasswordComplexity = 1", "PasswordComplexity = 0") | Out-File C:\secpol.cfg
		# (Get-Content C:\secpol.cfg).replace("MinimumPasswordAge = 1", "MinimumPasswordAge = 0") | Out-File C:\secpol.cfg
		# (Get-Content C:\secpol.cfg).replace("MinimumPasswordLength = 14", "MinimumPasswordLength = 0") | Out-File C:\secpol.cfg
		#secedit /configure /db c:\windows\security\local.sdb /cfg c:\secpol.cfg /areas SECURITYPOLICY
		# $process = Start-Process -FilePath ("secedit") -ArgumentList @("/configure","/db","c:\windows\security\local.sdb","/cfg","c:\secpol.cfg","/areas","SECURITYPOLICY") -PassThru -NoNewWindow -wait
		# Remove-Item -force c:\secpol.cfg -confirm:$false
		# net accounts /minpwage:0 /minpwlen:0
		ForEach ( $i in $UserRange) {	
			If ($i) {
				#Only create profile if user is a local user
				If (-Not (Get-LocalUser -Name ("Window" + $i) -erroraction 'silentlycontinue')) {
					#Only create profile if no profile exists
					If (-Not (Test-Path ((Get-WmiObject Win32_UserProfile |Where-Object { (Split-Path -leaf -Path ($_.LocalPath)) -eq $CurrentProfile} |Select-Object Localpath).localpath + "\ntuser.dat"))) {
						write-host ("Creating User: " +("Window" + $i))
						#Password same as username
						#New-LocalUser -Name ("Window" + $i).ToLower() -Description "Store Window User" -FullName ("Window" + $i) -Password (ConvertTo-SecureString ("Window" + $i).ToLower() -AsPlainText -Force) -AccountNeverExpires -UserMayNotChangePassword -PasswordNeverExpires | Out-Null
						#Random 120 chr. password
						$TempPass= (ConvertTo-SecureString ([system.web.security.membership]::GeneratePassword(120,32)).tostring() -AsPlainText -Force)
						New-LocalUser -Name ("Window" + $i).ToLower() -Description "Store Window User" -FullName ("Window" + $i) -Password $TempPass -AccountNeverExpires -UserMayNotChangePassword -PasswordNeverExpires | Out-Null
						Add-LocalGroupMember -Name 'Administrators' -Member ("Window" + $i) | Out-Null
						Write-Host "`tWorking on Creating user profile: " ("Window" + $i)
						#launch process as user to create user profile
						# https://msdn.microsoft.com/en-us/library/system.diagnostics.processstartinfo(v=vs.110).aspx
						$processStartInfo = New-Object System.Diagnostics.ProcessStartInfo
						$processStartInfo.UserName = ("Window" + $i)
						$processStartInfo.Domain = "."
						$processStartInfo.Password = $TempPass
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
							$HideAccounts += ("Window" + $i).ToLower()
							#Grant Current user rights on new Profiles
							Write-Host ("`tUpdating ACLs and adding to Profile List: " + ($UsersProfileFolder + "\Window" + $i))
							$Folderpath=($UsersProfileFolder + "\Window" + $i)
							$user_account=$env:username
							$Acl = Get-Acl $Folderpath
							$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule($user_account, "FullControl", "ContainerInherit, ObjectInherit", "None", "Allow")
							$Acl.Setaccessrule($Ar)
							Set-Acl $Folderpath $Acl
							#Disable User.
							Disable-LocalUser -Name ("Window" + $i).ToLower() -Confirm:$false
						}
					}
				} else {
				#Only create profile if no profile exists
				$CurrentUserSID = (Get-LocalUser -Name ("Window" + $i)).SID
				If (Get-Command Get-CimInstance -errorAction SilentlyContinue) {
					$UserProfile = (Get-CimInstance Win32_UserProfile | Where-Object { $_.SID -eq $CurrentUserSID}).localpath
				} Else {
					$UserProfile = (Get-WmiObject Win32_UserProfile | Where-Object { $_.SID -eq $CurrentUserSID}).localpath
				} 
				If (-Not (Test-Path ($UserProfile + "\ntuser.dat"))) {
					If ((Get-LocalUser -Name ("Window" + $i)).Enabled -eq $false) {
						Enable-LocalUser -Name ("Window" + $i).ToLower()
					}
					Write-Host ("`tResetting password for profile: " + "Window" + $i)
					$TempPass= (ConvertTo-SecureString ([system.web.security.membership]::GeneratePassword(120,32)).tostring() -AsPlainText -Force)
					Set-LocalUser -Name ("Window" + $i).ToLower() -Password $TempPass
					Write-Host ("`tWorking on Creating user profile: " + "Window" + $i)
					#launch process as user to create user profile
					# https://msdn.microsoft.com/en-us/library/system.diagnostics.processstartinfo(v=vs.110).aspx
					$processStartInfo = New-Object System.Diagnostics.ProcessStartInfo
					$processStartInfo.UserName = ("Window" + $i)
					$processStartInfo.Domain = "."
					$processStartInfo.Password = $TempPass
					$processStartInfo.FileName = ($env:windir + "\system32\cmd.exe")
					$processStartInfo.WorkingDirectory = $LICache
					$processStartInfo.Arguments = "/C echo . && echo %username% && echo ."
					$processStartInfo.LoadUserProfile = $true
					$processStartInfo.UseShellExecute = $false
					#$processStartInfo.WindowStyle  = "minimized"
					$processStartInfo.RedirectStandardOutput = $false
					$process = [System.Diagnostics.Process]::Start($processStartInfo)
					$Process.WaitForExit()   
					#Add setup user to profiles created to allow registry to be created. 
					If (Test-Path ($UsersProfileFolder + "\Window" + $i)) {
						$ProfileList.Add(("Window" + $i).ToLower()) | Out-Null
						$HideAccounts += ("Window" + $i).ToLower()
						#Grant Current user rights on new Profiles
						Write-Host ("`tUpdating ACLs and adding to Profile List: " + ($UsersProfileFolder + "\Window" + $i))
						$user_account=$env:username
						$Acl = Get-Acl ($UsersProfileFolder + "\Window" + $i)
						$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule($user_account, "FullControl", "ContainerInherit, ObjectInherit", "None", "Allow")
						$Acl.Setaccessrule($Ar)
						Set-Acl ($UsersProfileFolder + "\Window" + $i) $Acl
						#Disable User.
						Disable-LocalUser -Name ("Window" + $i).ToLower() -Confirm:$false
					}
				}
			}
			}
		}
	}
}
#If not logged in as administrator and administrators groups has more than one user set administrator account with random password.
If ($env:username -ne "Administrator") {
	If ((Get-LocalGroupMember -Name 'Administrators').count -gt 1) {
		#Sets Random 265 character password
		set-localuser -Name 'Administrator' -Password (ConvertTo-SecureString ([system.web.security.membership]::GeneratePassword(128,32) + [system.web.security.membership]::GeneratePassword(128,32)).tostring() -AsPlainText -Force )
		Disable-LocalUser -Name 'Administrator' -Confirm:$false
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
			$UserProfile = (Get-WmiObject Win32_UserProfile |Where-Object { (Split-Path -leaf -Path ($_.LocalPath)) -eq $CurrentProfile} |Select-Object Localpath).localpath	
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
				ForEach ( $file in $StoreDenyFolder) {
					If (Test-Path $file) {
						Write-Host ("`t`tDenying: " + $file)
						$Acl = Get-Acl ($file)
						$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule(($env:computer + "\" + $CurrentProfile), "ReadAndExecute", "Deny")
						$Acl.Setaccessrule($Ar)
						Set-Acl ($file) $Acl	
					}
				}
				#Add Deny ACL User Profile
				ForEach ( $file in $StoreDenyFolderUser) {
					If (Test-Path ($UserProfile + "\"+ $file)) {
						Write-Host ("`t`tDenying: " + $file)
						$Acl = Get-Acl ($UserProfile + "\"+ $file)
						$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule(($env:computer + "\" + $CurrentProfile), "ReadAndExecute", "Deny")
						$Acl.Setaccessrule($Ar)
						Set-Acl ($UserProfile + "\"+ $file) $Acl	
						Get-ChildItem -path ($UserProfile + "\"+ $file) -Recurse -Force | ForEach-Object {$_.attributes = "Hidden"}
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
	#Set Wallpaper style
	# 0:  The image is centered if TileWallpaper=0 or tiled if TileWallpaper=1
	# 2:  The image is stretched to fill the screen
	# 6:  The image is resized to fit the screen while maintaining the aspect ratio. (Windows 7 and later)
	#10:  The image is resized and cropped to fill the screen while maintaining the aspect ratio. (Windows 7 and later)
	If ([environment]::OSVersion.Version.Major -ge 10) { 
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Control Panel\Desktop") "WallpaperStyle" "10" "STRING"
	}else{
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Control Panel\Desktop") "WallpaperStyle" "6" "STRING"	
	}
	#Setup Theme settings
	#Use Dark theme
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize") "AppsUseLightTheme" "0" "DWORD"
	#Disable Taskbar Transparency
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize") "EnableTransparency" "0" "DWORD"
	#Set AccentColor
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\DWM") "AccentColor" "4292311040" "DWORD"
	#Setup Screen Saver
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Control Panel\Desktop") "ScreenSaveActive" "1" "STRING"
	If ($Store){
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Control Panel\Desktop") "ScreenSaverIsSecure" "0" "STRING"
	}else{
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Control Panel\Desktop") "ScreenSaverIsSecure" "1" "STRING"
	}
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Control Panel\Desktop") "ScreenSaveTimeOut" "600" "STRING"
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Control Panel\Desktop") "SCRNSAVE.EXE" "C:\Windows\system32\scrnsave.scr" "STRING"	
	#Set FontSmoothing
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Control Panel\Desktop") "FontSmoothing" "2" "STRING"
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
	#Don't ask to search the Internet for the correct program when opening a file with an unknown extension
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
	#UI Tweaks
	If ($IsVM) {
		write-host ("`t" + $CurrentProfile + ": Setting UI Optimizations")
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects") "VisualFXSetting" 3 "DWORD"
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced") "ListviewAlphaSelect" 0 "DWORD"
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced") "ListviewShadow" 0 "DWORD"
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced") "MinAnimate" 0 "DWORD"
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced") "TaskbarAnimations" 0 "DWORD"
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\DWM") "EnableAeroPeek" 0 "DWORD"
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\DWM") "AlwaysHibernateThumbnails" 0 "DWORD"
		#Chrome needs FontSmoothing Enabled
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Control Panel\Desktop") "FontSmoothing" "2" "String"
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Control Panel\Desktop") "MenuShowDelay" "0" "String"
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Control Panel\Desktop") "CursorBlinkRate" "-1" "String"
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Control Panel\Desktop") "UserPreferencesMask" (90,12,01,80) "Binary"
		# DisableTransparency:
		Write-Host "Removing Transparency Effects..." -ForegroundColor Green
		Write-Host ""
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize") "EnableTransparency" 0 "DWORD"
	}
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
		If (Test-RegistryKeyValue -Path ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") -Name "DisallowCPL") {
			Remove-ItemProperty -Path ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") -Name "DisallowCPL" -erroraction 'silentlycontinue' | Out-Null
		}
		If (Test-RegistryKeyValue -Path ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") -Name "DisallowRun") {
			Remove-ItemProperty -Path ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") -Name "DisallowRun" -erroraction 'silentlycontinue' | Out-Null
		}
		#Set ScreenSaver to scrnsave.scr
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Control Panel\Desktop") "SCRNSAVE.EXE" "scrnsave.scr" "String"
		#Disable Windows Experience Index
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Control Panel\Performance Control Panel") "PerfCplEnabled" 0 "DWORD"
		#Turn off access to the solutions to performance problems section
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Control Panel\Performance Control Panel") "SolutionsEnabled" 0 "DWORD"
		#Prevent changing sounds
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Personalization") "NoChangingSoundScheme" 1 "DWORD"
		#Prevent printing over HTTP
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows NT\Printers") "DisableHTTPPrinting" 1 "DWORD"
		#Turn off downloading of print drivers over HTTP
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows NT\Printers") "DisableWebPnPDownload" 1 "DWORD"
		#endregion Control Panel
		#region LockDown Windows Explorer
		Set-Reg ($HKEYWE + "\Advanced") "Start_ShowDownloads" 0 "DWORD"
		Set-Reg ($HKEYWE + "\Advanced") "StartMenuAdminTools" 0 "DWORD"
		Set-Reg ($HKEYWE + "\Advanced") "Start_AdminToolsRoot" 0 "DWORD"
		Set-Reg ($HKEYWE + "\Advanced") "TaskbarSizeMove" 0 "DWORD"
		Set-Reg ($HKEYWE + "\Advanced") "Start_ShowControlPanel" 1 "DWORD"
		#Disables "Recently added" Apps List on the Start Menu
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "HideRecentlyAddedApps" 1 "DWORD"
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
		#endregion LockDown Start Menu
		#region Adobe Reader
		#Accept EULA
		Set-Reg	($HKEY.replace("HKU\","HKU:\") + "\Software\Adobe\Acrobat Reader\DC\AdobeViewer") "EULA" 1 "DWORD"	
		#endregion Adobe Reader
		#Hide All Drives Tc
		If (($Store) -and ($CurrentProfile.ToUpper() -ne "DEFAULT" )) {
			#region Show only select Contol Panel Icons
			If (Test-Path($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\RestrictCpl")) {
				Remove-Item ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\RestrictCpl") -Recurse
			}
			$i = 1
			ForEach ( $item in $ShowOnlyCPL) {
					Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\RestrictCpl") $i $item "String"
					$i++
				}	
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "RestrictCpl" 1 "DWORD"
			#endregion Show only select Contol Panel Icons
			#region LockDown Store Windows Explorer
			write-host ("`t" + $CurrentProfile + ": Setting up Store settings Windows Explorer")
			#Prevent Changing Wallpaper
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop") "NoChangingWallPaper" 1 "DWORD"
			#Hide Items in settings
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "SettingsPageVisibility" $SettingsPageVisibility "String"
			#Removes the contents of the Documents menu when Windows is shut down or the user logs off.	
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "ClearRecentDocsOnExit" 1 "DWORD"
			#Force Delete Confirmation Dialog Box in Windows 
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "ConfirmFileDelete" 1 "DWORD"
			#Disabling the thumbnail previewing feature in Windows Explorer will speed up access to the folder and increase system response time especially when have to browse back and forth between multiple folders.
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "DisableThumbnails" 1 "DWORD"
			#Disable creation of Thumbs.db on Network shares to remove deletion errors
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "DisableThumbnailsOnNetworkFolders" 1 "DWORD"
			#Disable the Thumbnail Cache (thumbs.db)
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoThumbnailCache" 1 "DWORD"
			#Hide Run on Start Menu
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "ForceRunOnStartMenu" 0 "DWORD"
			#Show Log Off on Start Menu
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "ForceStartMenuLogOff" 1 "DWORD"
			#This entry makes it easier for users to distinguish between programs that are fully installed and those that are only partially installed.
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "GreyMSIAds" 1 "DWORD"
			#Greyed out in the Action Center configuration, and you will no longer see the Action Center icon in the system tray.
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "HideSCAHealth" 1 "DWORD"
			#Don't tie new shortcuts to a specific PC
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "LinkResolveIgnoreLinkInfo" 1 "DWORD"
			#Lock all taskbar settings
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "LockTaskbar" 1 "DWORD"
			#Turn off notification area cleanup
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoAutoTrayNotify" 1 "DWORD"
			#Disable CD Burning in Windows 10
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoCDBurning" 1 "DWORD"
			#Prevents users from using the drag-and-drop method to reorder or remove items on the Start menu
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoChangeStartMenu" 1 "DWORD"
			#Allow User to Shutdown 
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoClose" 0 "DWORD"
			#To stop users from listing machines in their local work-groups or domains via Windows Explorer or My Network Places (Network Neighborhood)
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoComputersNearMe" 1 "DWORD"
			#Removes the Dfs tab from Windows Explorer and from other programs that use the Windows Explorer browser, such as My Computer.
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoDFSTab" 1 "DWORD"
			#Disable Autorun for all devices
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoDriveTypeAutoRun" 255 "DWORD"	
			#Hide all drives in Windows Explorer	
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoDrives" 67108863 "DWORD"
			#Prevents users from using My Computer to gain access to the content of selected drives.
			#Allow access to the C:\ and Z:\ Drive. 
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoViewOnDrive" 33554427 "DWORD"
			#Removes the File menu from My Computer and Windows Explorer and disables the File menu in Internet Explorer.
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoFileMenu" 1 "DWORD"
			#Removes the Search item from the Start menu and disables some Windows Explorer search elements.
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoFind" 1 "DWORD"
			#Removes the Folder Options item from all Windows Explorer menus and removes the Folder Options item from Control Panel.
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoFolderOptions" 1 "DWORD"
			#A value of 1 removes the Hardware tab from Mouse, Keyboard, and Sounds and Multimedia in Control Panel. It also removes the Hardware tab from the Properties dialog box for all local drives, including hard drives, floppy disk drives, and CD-ROM drives. As a result, users cannot use the Hardware tab to view or change the device list or device properties, or use the Troubleshoot button to resolve problems with the device.
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoHardwareTab" 1 "DWORD"
			#Dis-Allows users to share files in their profiles to prevent unauthorized access or exposure of sensitive data.
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoInplaceSharing" 1 "DWORD"
			#Disables user tracking and features that require tracking data, such as personalized menus.
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoInstrumentation" 1 "DWORD"
			#This setting prevents unhandled file associations from using the Microsoft Web service to find an application.
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoInternetOpenWith" 1 "DWORD"
			#The user can log off can be seen in start menu.
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoLogoff" 0 "DWORD"
			#Disable the low disk space warning messages 
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoLowDiskSpaceChecks" 1 "DWORD"
			#The Manage item is removed from the Windows Explorer context menu.
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoManageMyComputerVerb" 1 "DWORD"
			#Prevents users from using Windows Explorer or My Network Places to map or disconnect network drives.
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoNetConnectDisconnect" 1 "DWORD"
			#Removes the My Network Places icon from the desktop.
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoNetHood" 1 "DWORD"
			#Prevents users from running Network and Dial-up Connections.
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoNetworkConnections" 0 "DWORD"
			#Turn Off the "Order Prints" Picture Task
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoOnlinePrintsWizard" 1 "DWORD"
			#Preview Pane in File Explorer is hidden and cannot be turned on by the user.
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoReadingPane" 1 "DWORD"
			#The Details Pane is always visible and cannot be hidden by the user. Note: This has a side effect of not being able to toggle to the Preview Pane since the two cannot be displayed at the same time.
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoPreviewPane" 1 "DWORD"
			#Disabling access to System Properties via My Computer is for the current user.
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoPropertiesMyComputer" 1 "DWORD"
			#Removes the Properties option from the Recycle Bin context menu.
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoPropertiesRecycleBin" 1 "DWORD"
			#Turn off the "Publish to Web" task for files and folders
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoPublishingWizard" 1 "DWORD"
			#Turn off Recent Documents in Start Menu
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoRecentDocsMenu" 1 "DWORD"
			#Remote shared folders are not added to My Network Places when you open a document in the shared folder.
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoRecentDocsNetHood" 1 "DWORD"
			#Disables Run program access 
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoRun" 1 "DWORD"
			#Remove Balloon Tips on Start Menu items
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoSMBalloonTip" 1 "DWORD"
			#Remove the "Set Program Access and Defaults" icon from the Start Menu
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoSMConfigurePrograms" 1 "DWORD"
			#Removes the Help command from the Start menu.
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoSMHelp" 1 "DWORD"
			#Removes the My Documents icon from the Start Menu and its submenus.
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoSMMyDocs" 1 "DWORD"
			#Hide My Pictures from the Start menu 
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoSMMyPictures" 1 "DWORD"
			#Disable start menu search of Outlook
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoSearchCommInStartMenu" 1 "DWORD"
			#"See all results" link will not be shown when the user performs a search in the start menu search box.
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoSearchComputerLinkInStartMenu" 1 "DWORD"
			#Start Menu search will not search for Files.
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoSearchFilesInStartMenu" 1 "DWORD"
			# Start Menu search will not search for Internet History
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoSearchInternetInStartMenu" 1 "DWORD"
			# Hide Security Tab in File/Folder Properties. 
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoSecurityTab" 1 "DWORD"
			#Users cannot open or use the Taskbar and Start Menu Properties dialog box
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoSetTaskbar" 1 "DWORD"
			#Remove Shared Documents from My Computer
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoSharedDocuments" 1 "DWORD"
			#The Search button is removed from the Windows Explorer toolbar.
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoShellSearchButton" 1 "DWORD"
			#Removes start banner balloon tip
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoStartBanner" 1 "DWORD"
			#Removes the "Undock PC" button from the Start Menu and prevents undocking of the PC
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoStartMenuEjectPC" 1 "DWORD"
			#Remove Games from Start Menu
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoStartMenuMyGames" 1 "DWORD"
			#Remove Music from Start Menu
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoStartMenuMyMusic" 1 "DWORD"
			#Remove Network from Start Menu
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoStartMenuNetworkPlaces" 1 "DWORD"
			#Disable Numerical Sorting in File Explorer 
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoStrCmpLogical" 1 "DWORD"
			#This setting disables the theme gallery in the Personalization Control Panel.
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoThemesTab" 1 "DWORD"
			#Context-sensitive menus are hidden.
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoTrayContextMenu" 1 "DWORD"
			#Show the Notifications Area
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoTrayItemsDisplay" 0 "DWORD"
			#Disable right-click on Desktop and Windows Explorer
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoViewContextMenu" 1 "DWORD"
			#Turn off Internet download for Web publishing and online ordering wizards	 
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoWebServices" 1 "DWORD"
			#Disabling the Windows-key is a quick operation
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoWinKeys" 1 "DWORD"
			#Disable access to the Windows update feature
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoWindowsUpdate" 1 "DWORD"
			#Prevent users from adding files to the root of their Users Files folder
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "PreventItemCreationInUsersFilesFolder" 1 "DWORD"
			#Turn off common control and window animations
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "TurnOffSPIAnimations" 1 "DWORD"
			#Disable Change Password Option from the CTRL + ALT + DEL Screen
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\System") "DisableChangePassword" 1 "DWORD"
			#Disable Command Prompt
			#To Disable Command Prompt Only: 		Change the data value with 2
			#To Disable Command Prompt and Scripts: Change the data value with 1
			#To Enable Command Prompt: 				Change the data value with 0
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\System") "DisableCMD" 1 "DWORD"
			#Disable Access to the Windows Registry
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\System") "DisableRegistryTools" 1 "DWORD"
			#Enable "Lock Workstation" when I press Ctrl-Alt-Del and (Window+L)
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\System") "DisableLockWorkstation" 0 "DWORD"
			#Disable Task Manager
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\System") "DisableTaskMgr" 1 "DWORD"
			#Hiding the Remote Administration Page
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\System") "NoAdminPage" 1 "DWORD"
			#Removes Add/Remove Programs from Control Panel and removes the Add/Remove Programs item from menus
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Uninstall") "NoAddRemovePrograms" 1 "DWORD"
			#Remove Search the Internet to the Windows Start Menu	
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "AddSearchInternetLinkInStartMenu" 0 "DWORD"
			#Disable Context Menus in the Start Menu 
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "DisableContextMenusInStart" 1 "DWORD"
			#Force Start to be menu size
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "ForceStartSize" 1 "DWORD"
			#Disable People Bar on Taskbar 
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "HidePeopleBar" 1 "DWORD"
			#Remove See More Results / Search Everywhere link
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "NoSearchEverywhereLinkInStartMenu" 1 "DWORD"
			#Remove Homegroup link from Start Menu
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "NoStartMenuHomegroup" 1 "DWORD"	
			#Remove Recorded TV link from Start Menu
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "NoStartMenuRecordedTV" 1 "DWORD"	
			#Remove Videos link from Start Menu
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "NoStartMenuVideos" 1 "DWORD"	
			#Remove the “Uninstall” Option from the Windows 10 Start Menu
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "NoUninstallFromStart" 1 "DWORD"
			#Disable Notification & Action Center 
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "DisableNotificationCenter" 1 "DWORD"	
			#Do not allow pinning Store app to the Taskbar
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "NoPinningStoreToTaskbar" 1 "DWORD"
			#Turn off automatic promotion of notification icons to the taskbar
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "NoSystraySystemPromotion" 1 "DWORD"
			#Disable Show Windows Store apps on the taskbar
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "ShowWindowsStoreAppsOnTaskbar" 2 "DWORD"
			#Disable Creating Thumbs.db on Network Folders
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "DisableThumbsDBOnNetworkFolders" 1 "DWORD"
			#Disable Explorer Search Box Suggestions
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "DisableSearchBoxSuggestions" 1 "DWORD"	
			#Remove the Search the Internet "Search again" link
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "NoSearchInternetTryHarderButton" 1 "DWORD"
			#Turn off Windows Libraries features that rely on indexed file data
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "DisableIndexedLibraryExperience" 1 "DWORD"
			#Turn off autoplay for non-volume devices	
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "NoAutoplayfornonVolume" 1 "DWORD"
			#Disable Look for an app in the Store
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "NoUseStoreOpenWith" 1 "DWORD"
			#Select the Power button action (plugged in)
			# Index	Name	Description
			# 0		Do Nothing		No action is taken when the power button is pressed.
			# 1		Sleep			The system enters sleep when the power button is pressed.
			# 2		Hibernate		The system enters hibernate when the power button is pressed.
			# 3		Shut Down		The system shuts down when the power button is pressed.
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "PowerButtonAction" 3 "DWORD"	
			#Turn Off Annoying Feature Advertisement 
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
			#Hide "Set Program Access and Computer Defaults" page
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Programs") "NoDefaultPrograms" 1 "DWORD"
			#Hide "Get Programs" page
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Programs") "NoGetPrograms" 1 "DWORD"
			#Hide "Installed Updates" page
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Programs") "NoInstalledUpdates" 1 "DWORD"
			#Hide "Windows Marketplace" 
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Programs") "NoWindowsMarketplace" 1 "DWORD"
			#Hide "Windows Features"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Programs") "NoWindowsFeatures" 1 "DWORD"
			#Hide "Programs and Features" page
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Programs") "NoProgramsAndFeatures" 1 "DWORD"
			#endregion Programs
			#region System
			write-host ("`t" + $CurrentProfile + ": Setting up Store settings System")
			#Prevent changing screen saver
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\System") "NoDispScrSavPage" 1 "DWORD"
			#Prevent changing visual style for windows and buttons
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\System") "NoVisualStyleChoice" 1 "DWORD"
			#Prevent changing color scheme 
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\System") "NoColorChoice" 1 "DWORD"
			#endregion System
			#region Updates
			write-host ("`t" + $CurrentProfile + ": Setting up Store settings Updates")
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\WindowsUpdate") "DisableWindowsUpdateAccess" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\WindowsUpdate") "DisableWindowsUpdateAccessMode" 1 "DWORD"
			#endregion Updates
			#region Internet Explorer
			write-host ("`t" + $CurrentProfile + ": Setting up Store settings Internet Explorer")
			#HDisable "Speed up Browsing by Disabling Add-ons" Popup
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Ext") "DisableAddonLoadTimePerformanceNotifications" 1 "DWORD"
			#Disable "The Add-on is Ready for Use" Popup Notification
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Ext") "IgnoreFrameApprovalCheck" 1 "DWORD"
			#Disable Enhanced Search Engine Suggestions in IE 11
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer") "AllowServicePoweredQSA" 0 "DWORD"
			#Turn-off import and export of Internet Explorer settings
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer") "DisableImportExportFavorites" 1 "DWORD"
			#Turn off Accelerators 
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Activities") "NoActivities" 1 "DWORD"
			#Automatic configuration of Internet Explorer is disabled
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Control Panel") "Autoconfig" 1 "DWORD"
			#Prevents users from changing certificate settings in Internet Explorer
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Control Panel") "Certificates" 1 "DWORD"
			#Prevents users from changing dial-up setting
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Control Panel") "Connection Settings" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Control Panel") "Connwiz Admin Lock" 1 "DWORD"
			#Disable or enable homepage setting in Internet Explorer
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Control Panel") "HomePage" 1 "DWORD"
			#user will not be able to configure proxy settings
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Control Panel") "Proxy" 1 "DWORD"
			#Disable changing ratings settings; Content Advisor area on the Content tab in the Internet Options dialog box appear dimmed
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Control Panel") "Ratings" 1 "DWORD"
			#Prevent automatic discovery of feeds and Web Slices
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Feed Discovery") "Enabled" 0 "DWORD"
			#Disable Basic authentication for RSS feeds over HTTP
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Feeds") "AllowBasicAuthInClear" 0 "DWORD"
			#Turn off background synchronization for feeds and Web Slices
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Feeds") "BackgroundSyncStatus" 0 "DWORD"
			#sers cannot change their Search Assistant settings such as setting default search engines for specific tasks
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Infodelivery\Restrictions") "NoSearchCustomization" 1 "DWORD"
			#disables the ability to save complete web pages including images, scripts, linked files and other elements in Internet Explorer.
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Infodelivery\Restrictions") "NoBrowserSaveWebComplete" 1 "DWORD"
			#Disable all scheduled offline pages
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Infodelivery\Restrictions") "NoScheduledUpdates" 1 "DWORD"
			#Disable downloading of site subscription content
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Infodelivery\Restrictions") "NoSubscriptionContent" 1 "DWORD"
			#Turn On Favorites bar
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\LinksBar") "Enabled" 1 "DWORD"
			#Prevent running First Run wizard 
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Main") "DisableFirstRunCustomize" 1 "DWORD"
			#Set Homepage URL
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Main") "Start Page" $HomePage "String"
			#Turn off the ability to launch report site problems using a menu option
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Main") "NoReportSiteProblems" "yes" "String"
			#Turn On Managing SmartScreen Filter
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\PhishingFilter") "EnabledV9" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\PhishingFilter") "EnabledV8" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\PhishingFilter") "Enabled" 1 "DWORD"
			#Disable Private Browsing
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Privacy") "EnableInPrivateBrowsing" 0 "DWORD"
			#Disable Internet Options
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Restrictions") "NoBrowserOptions" 1 "DWORD"
			#Disable Smiley Button in Internet Explorer
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Restrictions") "NoHelpItemSendFeedback" 1 "DWORD"
			# Remove the "For Netscape Users" menu item
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Restrictions") "NoHelpItemNetscapeHelp" 1 "DWORD"
			#Remove the "Tip of the Day" menu item
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Restrictions") "NoHelpItemTipOfTheDay" 1 "DWORD"
			#Remove the "Tour" (Tutorial) menu item
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Restrictions") "NoHelpItemTutorial" 1 "DWORD"
			#Disable the entire help menu
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Restrictions") "NoHelpMenu" 1 "DWORD"
			#Removes the Save As command on the File menu.
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Restrictions") "NoBrowserSaveAs" 1 "DWORD"
			#Turns off the Source command on the View menu. 
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Restrictions") "NoViewSource" 1 "DWORD"
			#Turns off Save on the File Download dialog box.
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Restrictions") "NoSelectDownloadDir" 1 "DWORD"
			#Disable or Turn Off InPrivate Browsing in Internet Explorer
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Safety\PrivacIE") "DisableInPrivateBlocking" 1 "DWORD"
			#Turn off suggestions for all user-installed providers
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\SearchScopes") "ShowSearchSuggestionsGlobal" 0 "DWORD"
			#Turn off the Security Settings Check feature 
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Security") "DisableFixSecuritySettings" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Security") "DisableSecuritySettingsCheck" 1 "DWORD"
			#Prevent per-user installation of ActiveX controls
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Security\ActiveX") "BlockNonAdminActiveXInstall" 1 "DWORD"
			#Turn off the Windows Customer Experience program
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\SQM") "DisableCustomerImprovementProgram" 0 "DWORD"
			#disable suggested sites and URLs in IE11
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Suggested Sites") "Enabled" 0 "DWORD"
			#disable automatic proxy caching in Internet Explorer
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings") "EnableAutoProxyResultCache" 0 "DWORD"
			#Enable Internet Explorer warning about certificate address mismatch
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings") "WarnOnBadCertRecving" 1 "DWORD"
			#Disable Saving of Passwords
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings") "DisablePasswordCaching" 1 "DWORD"
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
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "AbusiveExperienceInterventionEnforce" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "AdsSettingForIntrusiveAdsSites" 2 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "AllowDinosaurEasterEgg" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "AuthNegotiateDelegateWhitelist" $ChromeDelegateWhiteList "String"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "AutofillCreditCardEnabled" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "AlwaysOpenPdfExternally" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "BookmarkBarEnabled" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "BrowserAddPersonEnabled" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "BrowserGuestModeEnabled" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "BrowserSignin" 0 "DWORD"
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
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "RestoreOnStartup" 4 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "SafeBrowsingEnabled" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "SearchSuggestEnabled" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "ShowHomeButton" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "ShowCastIconInToolbar" 0 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "SyncDisabled" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome") "TranslateEnabled" 0 "DWORD"
			#Disables all extensions
			If (Test-Path ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome\ExtensionInstallBlacklist")) {
				Remove-Item ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome\ExtensionInstallBlacklist") -Recurse | out-null
			}
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome\ExtensionInstallBlacklist") 1 "*" "String"
			#Sets Startup page
			If (Test-Path ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome\RestoreOnStartupURLs")) {
				Remove-Item ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome\RestoreOnStartupURLs") -Recurse | out-null
			}
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome\RestoreOnStartupURLs") 1 $HomePage "String"				
			#$ChromeURLBlackList Stops local browsing
			If (Test-Path($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome\URLBlacklist")) {
				Remove-Item ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome\URLBlacklist") -Recurse
			}
			$i = 1
			ForEach ( $item in $ChromeURLBlackList) {
				Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome\URLBlacklist") $i $item "String"
				$i++
			}			
			#endregion Google Chrome
			#region Auto Start Custom EXE
			If ($Custom_Software_Exec -and $Custom_Software_Path) {
				If (-Not (Test-Path ($UsersProfileFolder + "\" + $CurrentProfile + "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\" + $Custom_Software_Exec.Substring(0,$Custom_Software_Exec.IndexOfAny(".")) + ".lnk"))) {
					If (-Not (Test-Path ($UsersProfileFolder + "\" + $CurrentProfile + "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"))) {
						New-Item -Force -ItemType "directory" -Path ($UsersProfileFolder + "\" + $CurrentProfile + "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup")
					}
					If (Test-Path ($Custom_Software_Path + "\" + $Custom_Software_Exec)) {
						$ShortCut = $WScriptShell.CreateShortcut($UsersProfileFolder + "\" + $CurrentProfile + "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\" + $Custom_Software_Exec.Substring(0,$Custom_Software_Exec.IndexOfAny(".")) + ".lnk")
						$ShortCut.TargetPath=($Custom_Software_Path + "\" + $Custom_Software_Exec)
						$ShortCut.WorkingDirectory = ($Custom_Software_Path)
						$ShortCut.IconLocation = ( $Custom_Software_Path + "\" + $Custom_Software_Exec + ",0")
						$ShortCut.Save()
						#Make ShortCut ran as admin https://stackoverflow.com/questions/28997799/how-to-create-a-run-as-administrator-shortcut-using-powershell
						$bytes = [System.IO.File]::ReadAllBytes($UsersProfileFolder + "\" + $CurrentProfile + "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\" + $Custom_Software_Exec.Substring(0,$Custom_Software_Exec.IndexOfAny(".")) + ".lnk")
						$bytes[0x15] = $bytes[0x15] -bor 0x20 #set byte 21 (0x15) bit 6 (0x20) ON
						[System.IO.File]::WriteAllBytes($UsersProfileFolder + "\" + $CurrentProfile + "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\" + $Custom_Software_Exec.Substring(0,$Custom_Software_Exec.IndexOfAny(".")) + ".lnk", $bytes)
					}
				}
			}
			#endregion Auto Start Custom EXE
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
			Remove-item -Recurse -Force -Confirm:$false -ErrorAction SilentlyContinue -Path ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\VMware, Inc.\VMware VDM") 
			#Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\VMware, Inc.\VMware VDM\Security") "AcceptTicketSSLAuth" 1 "String"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\VMware, Inc.\VMware VDM\Client\Security") "LogInAsCurrentUser" 0 "String"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\VMware, Inc.\VMware VDM\Client\Security") "LogInAsCurrentUser_Display" 0 "String"
			#endregion VMware View
			#region Other ??
			write-host ("`t" + $CurrentProfile + ": Setting up Store settings Other Settings")
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\MobilityCenter") "NoMobilityCenter" 1 "DWORD"
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Network") "NoEntireNetwork" 1 "DWORD"
			#Hides "This PC" in Windows Explorer
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
			#Disable Windows 10 managed default printer
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows NT\CurrentVersion\Windows") "LegacyDefaultPrinterMode" 1 "DWORD"
			#Disable PEOPLE Icon From Taskbar
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\SOFTWARE\Microsoft\Windows\Explorer\Advanced") "PEOPLEBAND" 0 "DWORD"		
			#endregion Other ??
			#region Disable Lock Screen
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Personalization") "NoLockScreen " 1 "DWORD"
			#endregion Disable Lock Screen
			#region OnDrive
			Remove-Itemproperty -Path ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Run") -name 'OneDriveSetup' -erroraction 'silentlycontinue'| out-null
			#endregion OnDrive
			#Remove Network
			# If(Test-Path ($HKEYWE + "\Desktop\NameSpace\{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}")) {
				# write-host ("`tNetwork from This PC ") -foregroundcolor "gray"
				# Remove-Item ($HKEYWE + "\Desktop\NameSpace\{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}") -Recurse | Out-Null
			# }	
			# Set-Reg ($HKEYWE + "\CLSID\{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}\ShellFolder") "Attributes" 1048576 "DWORD"	
			# Set-Reg ($HKEYWE.replace("\Software\","\Software\Wow6432Node\") + "\CLSID\{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}\ShellFolder") "Attributes" 1048576 "DWORD"	
			
		}
	} else {
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
	$temp = (Get-ItemProperty -Path ($HKEYIS + "\Connections") -name "DefaultConnectionSettings" -erroraction 'silentlycontinue').DefaultConnectionSettings  | out-null
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
	If ([Environment]::Is64BitOperatingSystem) {
		Set-Reg ($HKEYIE.replace("\Software\","\Software\Wow6432Node\") + "\PageSetup") "header" $IE_Header "String"
		Set-Reg ($HKEYIE.replace("\Software\","\Software\Wow6432Node\") + "\PageSetup") "footer" $IE_Footer "String"
		Set-Reg ($HKEYIE.replace("\Software\","\Software\Wow6432Node\") + "\PageSetup") "margin_bottom" $IE_Margin_Bottom "String"
		Set-Reg ($HKEYIE.replace("\Software\","\Software\Wow6432Node\") + "\PageSetup") "margin_top" $IE_Margin_Top "String"
		Set-Reg ($HKEYIE.replace("\Software\","\Software\Wow6432Node\") + "\PageSetup") "margin_left" $IE_Margin_Left "String"
		Set-Reg ($HKEYIE.replace("\Software\","\Software\Wow6432Node\") + "\PageSetup") "margin_right" $IE_Margin_Right "String"
	}
	#IE Cache Settings Size Stored in KB not MB to convert MB to KB
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\5.0\Cache\Content\CacheLimit") "CacheLimit" ($IE_Cache_Size * 1024) "DWORD"
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Cache\Content\CacheLimit") "CacheLimit" ($IE_Cache_Size * 1024) "DWORD"
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Cache") "Persistent" 0 "DWORD"
	#Clean up old keys
	If (Test-Path ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains")) {
		Remove-Item ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains") -Recurse -Confirm:$false | out-null
	}
	If (Test-Path ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\EscDomains")) {
		Remove-Item ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\EscDomains") -Recurse -Confirm:$false | out-null
	}
	#IE Settings Trusted Sites
	Set-Reg ($HKEYIS + "\ZoneMap\Domains\blank") "about" 2 "DWORD"
	Set-Reg ($HKEYIS + "\ZoneMap\EscDomains\blank") "about" 2 "DWORD"
	If ([Environment]::Is64BitOperatingSystem) {
		Set-Reg ($HKEYIS.replace("\Software\","\Software\Wow6432Node\") + "\ZoneMap\EscDomains\blank") "about" 2 "DWORD"
	}
	#Company Set sites 
	ForEach ( $item in $ZoneMap) {
		write-host ("`t`tAdding Site: " + $item.Site + " to zone: " + $item.Zone + " for protocol: " + $item.Protocol)
		Set-Reg ($HKEYIS + "\ZoneMap\Domains\" +  $item.Site) $item.Protocol $item.Zone "DWORD"
		Set-Reg ($HKEYIS + "\ZoneMap\EscDomains\" +  $item.Site) $item.Protocol $item.Zone "DWORD"
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
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\SOFTWARE\Microsoft\Windows\CurrentVersion\Search") "BingSearchEnabled" 0 "DWORD"
	#Search 
	write-host ("`t" + $CurrentProfile + ": Search from This PC ...")
	#0 = Hidden
	#1 = Show search or Cortana icon
	#2 = Show search box
	Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\SOFTWARE\Microsoft\Windows\CurrentVersion\Search") "SearchboxTaskbarMode" 0 "DWORD"
	#Hide VMWare Tools
	If (Test-Path ("HKLM:\SOFTWARE\VMware, Inc.\VMware Tools") ) {
		Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\VMware, Inc.\VMware Tools") "ShowTray" 0 "DWORD"
	}
	#region Remove Chrome Settings
	If (Test-Path ($UserProfile + "\AppData\Local\Google")) {
		Remove-Item -Path ($UserProfile + "\AppData\Local\Google") -Recurse -Confirm:$false | out-null
	}
	If (Test-Path ($HKEY.replace("HKU\","HKU:\") + "\SOFTWARE\Google")) {
		Remove-Item -Recurse -Confirm:$false -Path ($HKEY.replace("HKU\","HKU:\") + "\SOFTWARE\Google") -erroraction 'silentlycontinue'
	}	
	#endregion Remove Chrome Settings
	#region Deploy Chrome Base Profile
	If (Test-Path ($LICache + "\" + $ChromeBaseZip)) {
		If (Test-Path ($UserProfile + "\AppData\Local")) {
			Write-Host ("`t" + $CurrentProfile + ": Setting-up Chrome Base Settings...")
			Expand-Archive -Path ($LICache + "\" + $ChromeBaseZip) -DestinationPath ($UserProfile + "\AppData\Local") -Force
		}
	}
	#endregion Deploy Chrome Base Profile
	#region Win+X Custom settings
	If (Test-Path ($LICache + "\" + $WinXZip)) {
		If (Test-Path ($UserProfile + "\AppData\Local\Microsoft\Windows\WinX")) {
			Write-Host ("`t" + $CurrentProfile + ": Setting-up Win+X Custom Settings...")
			#Remove standard entries before adding customized ones. 
			Remove-Item -Recurse -Confirm:$false -Path ($UserProfile + "\AppData\Local\Microsoft\Windows\WinX") -erroraction 'silentlycontinue'
			Expand-Archive -Path ($LICache + "\" + $WinXZip) -DestinationPath ($UserProfile + "\AppData\Local\Microsoft\Windows") -Force
		}
	}
	#endregion Win+X Custom settings
	# Unload the default profile hive
	Write-Host ("`t" + $CurrentProfile + ": Unloading User Registry")
	[gc]::collect()
	$process = (REG UNLOAD $HKEY)
	If ($LASTEXITCODE -ne 0 ) {
		[gc]::collect()
		Start-Sleep 3
		$process = (REG UNLOAD $HKEY)
		If ($LASTEXITCODE -ne 0 ) {
			write-error ("`t" + $CurrentProfile + ": Can not unload user registry!")
		}
	}

	#region Replace Favorites
	If (Test-Path ($LICache + "\Favorites")) {
		write-host ("`t" + $CurrentProfile + ": Setting up Favorites")
		If ($CurrentProfile -eq "Default") {
			If (Test-Path ($UsersProfileFolder + "\Default\Favorites")) {
				Remove-Item -path ($UsersProfileFolder + "\Default\Favorites") -recurse -force
				Copy-Item  ($LICache + "\Favorites") -Destination ($UsersProfileFolder + "\Default\Favorites") -recurse -force
			}
		}else{
			$UserProfile = (Get-WmiObject Win32_UserProfile |Where-Object { (Split-Path -leaf -Path ($_.LocalPath)) -eq $CurrentProfile} |Select-Object Localpath).localpath
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
	If ([environment]::OSVersion.Version.Major -ge 10) {
		#region Windows Feature setup
		Write-Host "Disabling Windows Features:"
		ForEach ( $Feature in $RemoveFeatures ) {
			If ((Get-WindowsOptionalFeature -Online -FeatureName $Feature).state -eq "Enabled") {
				Write-Host ("`t" + $Feature) -ForegroundColor gray
				Disable-WindowsOptionalFeature -Online -FeatureName $Feature -NoRestart | out-null
			}
			If ((Get-WindowsCapability -Online | Where-Object {$_.name -like ("*" + $Feature + "*") -and $_.state -eq "Installed"}).state) {
				Write-Host ("`t" + $Feature) -ForegroundColor gray
				Get-WindowsCapability -Online | Where-Object {$_.name -like ("*" + $Feature + "*") -and $_.state -eq "Installed"} | Remove-WindowsCapability -online | out-null
			}
		}
		#endregion Windows Feature setup
		#region Hiding Accounts
		Write-Host "Hiding accounts from login screen ..."
		If (-Not (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon")) {
			New-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" | Out-Null
		}
		If (-Not (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\SpecialAccounts")) {
			New-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\SpecialAccounts" | Out-Null
		}
		If (-Not (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\SpecialAccounts\UserList")) {
			New-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\SpecialAccounts\UserList" | Out-Null
		}	
		ForEach ($Account in $HideAccounts) {
			Write-Host ("`tHiding: " + $Account) -foregroundcolor "gray"
			Set-Reg "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\SpecialAccounts\UserList" $Account 0 "DWORD"
		}
		#endregion Hiding Accounts

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
		# WiFi Sense: HotSpot Sharing: Disable
		If (-Not (Test-Path "HKLM:\Software\Microsoft\PolicyManager\default\WiFi\AllowWiFiHotSpotReporting")) {
			Write-Host "WiFi Sense: HotSpot Sharing: Disable"
			New-Item -Path "HKLM:\Software\Microsoft\PolicyManager\default\WiFi\AllowWiFiHotSpotReporting" | Out-Null
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

		# Disable Scheduled Tasks:
		Write-Host "Disabling Scheduled Tasks..." -ForegroundColor Cyan
		Write-Host ""
		Disable-ScheduledTask -TaskName "\Microsoft\Windows\Autochk\Proxy" -erroraction 'silentlycontinue'| Out-Null
		Disable-ScheduledTask -TaskName "\Microsoft\Windows\Bluetooth\UninstallDeviceTask" -erroraction 'silentlycontinue'| Out-Null
		If ($IsVM) {
			Disable-ScheduledTask -TaskName "\Microsoft\Windows\Defrag\ScheduledDefrag" -erroraction 'silentlycontinue'| Out-Null
			Disable-ScheduledTask -TaskName "\Microsoft\Windows\Diagnosis\Scheduled" -erroraction 'silentlycontinue'| Out-Null
			Disable-ScheduledTask -TaskName "\Microsoft\Windows\DiskDiagnostic\Microsoft-Windows-DiskDiagnosticDataCollector" -erroraction 'silentlycontinue'| Out-Null
			Disable-ScheduledTask -TaskName "\Microsoft\Windows\DiskDiagnostic\Microsoft-Windows-DiskDiagnosticResolver" -erroraction 'silentlycontinue'| Out-Null
			Disable-ScheduledTask -TaskName "\Microsoft\Windows\Maintenance\WinSAT" -erroraction 'silentlycontinue'| Out-Null
			Disable-ScheduledTask -TaskName "\Microsoft\Windows\MemoryDiagnostic\ProcessMemoryDiagnosticEvents" -erroraction 'silentlycontinue'| Out-Null
			Disable-ScheduledTask -TaskName "\Microsoft\Windows\MemoryDiagnostic\RunFullMemoryDiagnostic" -erroraction 'silentlycontinue'| Out-Null
			Disable-ScheduledTask -TaskName "\Microsoft\Windows\Power Efficiency Diagnostics\AnalyzeSystem" -erroraction 'silentlycontinue'| Out-Null
			Disable-ScheduledTask -TaskName "\Microsoft\Windows\RecoveryEnvironment\VerifyWinRE" -erroraction 'silentlycontinue'| Out-Null
		}
		Disable-ScheduledTask -TaskName "\Microsoft\Windows\Mobile Broadband Accounts\MNO Metadata Parser" -erroraction 'silentlycontinue'| Out-Null	
		Disable-ScheduledTask -TaskName "\Microsoft\Windows\Ras\MobilityManager" -erroraction 'silentlycontinue'| Out-Null	
		Disable-ScheduledTask -TaskName "\Microsoft\Windows\Registry\RegIdleBackup" -erroraction 'silentlycontinue'| Out-Null
		Disable-ScheduledTask -TaskName "\Microsoft\Windows\RetailDemo\CleanupOfflineContent" -erroraction 'silentlycontinue'| Out-Null
		Disable-ScheduledTask -TaskName "\Microsoft\Windows\Shell\FamilySafetyMonitor" -erroraction 'silentlycontinue'| Out-Null
		Disable-ScheduledTask -TaskName "\Microsoft\Windows\Shell\FamilySafetyRefresh" -erroraction 'silentlycontinue'| Out-Null
		Disable-ScheduledTask -TaskName "\Microsoft\Windows\SystemRestore\SR" -erroraction 'silentlycontinue'| Out-Null
		Disable-ScheduledTask -TaskName "\Microsoft\Windows\UPnP\UPnPHostConfig" -erroraction 'silentlycontinue'| Out-Null
		Disable-ScheduledTask -TaskName "\Microsoft\Windows\WDI\ResolutionHost" -erroraction 'silentlycontinue'| Out-Null
		Disable-ScheduledTask -TaskName "\Microsoft\Windows\Windows Media Sharing\UpdateLibrary" -erroraction 'silentlycontinue'| Out-Null
		Disable-ScheduledTask -TaskName "\Microsoft\Windows\WOF\WIM-Hash-Management" -erroraction 'silentlycontinue'| Out-Null
		Disable-ScheduledTask -TaskName "\Microsoft\Windows\WOF\WIM-Hash-Validation" -erroraction 'silentlycontinue'| Out-Null
		
		
	}
	Write-Host "Disabling Windows Defender AntiSpyware ..."
	Set-Reg "HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender" "DisableAntiSpyware" 1 "DWORD"
	Set-Reg "HKLM:\SOFTWARE\Policies\Microsoft\Microsoft Antimalware\Real-Time Protection" "DisableScriptScanning" 1 "DWORD"
	Set-Reg "HKLM:\SOFTWARE\Policies\Microsoft\Microsoft Antimalware\Real-Time Protection" "LocalSettingOverrideDisableScriptScanning" 0 "DWORD"
		
	#Remove OneDrive from This PC
	If (Test-Path "HKCR:\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}") {
		Set-Reg "HKCR:\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}" "System.IsPinnedToNameSpaceTree"  0 "DWORD"
		If ([Environment]::Is64BitOperatingSystem) {
			Set-Reg "HKCR:\WOW6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}" "System.IsPinnedToNameSpaceTree"  0 "DWORD"
		}
	}
	#Removes UsersLibraries from This PC
	If (Test-Path "HKCR:\CLSID\{031E4825-7B94-4dc3-B131-E946B44C8DD5}") {
		Set-KeyOwnership "HKCR" "CLSID\{031E4825-7B94-4dc3-B131-E946B44C8DD5}" 
		Set-Reg "HKCR:\CLSID\{031E4825-7B94-4dc3-B131-E946B44C8DD5}" "System.IsPinnedToNameSpaceTree"  0 "DWORD"
		If ([Environment]::Is64BitOperatingSystem) {
			Set-KeyOwnership "HKCR" "WOW6432Node\CLSID\{031E4825-7B94-4dc3-B131-E946B44C8DD5}" 
			Set-Reg "HKCR:\WOW6432Node\CLSID\{031E4825-7B94-4dc3-B131-E946B44C8DD5}" "System.IsPinnedToNameSpaceTree"  0 "DWORD"
		}
	}
	
	#Disable ThumbnailCache
	Set-Reg "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "DisableThumbnailCache" 1 "DWORD"
	#Harden lsass Processing|Print
	# https://windowsforum.com/threads/windows-hardening-guide-securing-the-lsass-process.230793/
	Set-Reg "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa" "RunAsPPL" 1 "DWORD"
	#https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2012-R2-and-2012/dn408187(v=ws.11)
	Set-Reg "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\LSASS.exe" "AuditLevel" 8 "DWORD"
	
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
		#Hide Shutdown on Logon Screen
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

		#Removes UsersLibraries from This PC
		write-host ("`tUsers Libraries from This PC ") -foregroundcolor "gray"
		If (Test-Path ("HKCR:\CLSID\{031E4825-7B94-4dc3-B131-E946B44C8DD5}")) {
			Set-Reg "HKCR:\CLSID\{031E4825-7B94-4dc3-B131-E946B44C8DD5}" "System.IsPinnedToNameSpaceTree" 0 "DWORD"
			If ([Environment]::Is64BitOperatingSystem) {
				Set-Reg "HKCR:\WOW6432Node\CLSID\{031E4825-7B94-4dc3-B131-E946B44C8DD5}" "System.IsPinnedToNameSpaceTree" 0 "DWORD"
			}
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
		
		If ($IsVM) {
			Write-Host "Disabling Hard Disk Timeouts..." -ForegroundColor Yellow
			Write-Host ""
			POWERCFG /SETACVALUEINDEX 381b4222-f694-41f0-9685-ff5bb260df2e 0012ee47-9041-4b5d-9b77-535fba8b1442 6738e2c4-e8a5-4a42-b16a-e040e769756e 0
			POWERCFG /SETDCVALUEINDEX 381b4222-f694-41f0-9685-ff5bb260df2e 0012ee47-9041-4b5d-9b77-535fba8b1442 6738e2c4-e8a5-4a42-b16a-e040e769756e 0

			# Disable Hibernate
			Write-Host "Disabling Hibernate..." -ForegroundColor Green
			Write-Host ""
			POWERCFG -h off
			# Disable System Restore
			Write-Host "Disabling System Restore..." -ForegroundColor Green
			Write-Host ""
			Disable-ComputerRestore -Drive "C:\"
			# Increase Service Startup Timeout:
			Write-Host "Increasing Service Startup Timeout To 180 Seconds..." -ForegroundColor Yellow
			Write-Host ""
			Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control' -Name 'ServicesPipeTimeout' -Value '180000'
			# Increase Disk I/O Timeout to 200 Seconds:
			Write-Host "Increasing Disk I/O Timeout to 200 Seconds..." -ForegroundColor Green
			Write-Host ""
			Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\Disk' -Name 'TimeOutValue' -Value '200'
			# Disable New Network Dialog:
			Write-Host "Disabling New Network Dialog..." -ForegroundColor Green
			Write-Host ""
			New-Item -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Network' -Name 'NewNetworkWindowOff' | Out-Null

		}
		# Setting "Power Button"
		# 0 -- Do nothing.
		# 1 -- Sleep.
		# 2 -- Hibernate.
		# 3 -- Shut down. <-- Current Setting
		# 4 -- Turn off the display.
		Write-Host 'Setting "Power Button"...' -ForegroundColor Green
		powercfg /SETDCVALUEINDEX SCHEME_CURRENT 4f971e89-eebd-4455-a8de-9e59040e7347 7648efa3-dd9c-4e3e-b566-50f929386280 3
		powercfg /SETDCVALUEINDEX SCHEME_CURRENT 4f971e89-eebd-4455-a8de-9e59040e7347 7648efa3-dd9c-4e3e-b566-50f929386280 3
		powercfg -SetActive SCHEME_CURRENT
		
	}
	#Diables lockscreen for stores
	If ($Store) {
		write-host ("`tDisabling Lockscreen") -foregroundcolor "gray"
		Set-Reg "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization" "NoLockScreen" 1 "DWORD"
		#Set-Reg "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI\TestHooks" "Threshold" 0 "DWORD"
	}
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
		Set-Reg ($HKAR + "\" + $CARV + "\FeatureLockDown\cServices") "bToggleWebConnectors" 1 "DWORD"
		Set-Reg ($HKAR + "\" + $CARV + "\FeatureLockDown\cServices") "bDisableSharePointFeatures" 0 "DWORD"
		#Wow6432Node
		If ([Environment]::Is64BitOperatingSystem) {
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
			Set-Reg ($HKAR.replace("\Software\","\Software\Wow6432Node\") + "\" + $CARV + "\FeatureLockDown\cServices") "bToggleWebConnectors" 1 "DWORD"
			Set-Reg ($HKAR.replace("\Software\","\Software\Wow6432Node\") + "\" + $CARV + "\FeatureLockDown\cServices") "bDisableSharePointFeatures" 0 "DWORD"
		}
	}
}
#============================================================================
#endregion Main Local Machine Adobe
#============================================================================
#============================================================================
#region Main Local Machine Services
#============================================================================
If (-Not $UserOnly) {
	Write-Host "Setting up Services. . . "
	# Source: https://github.com/W4RH4WK/Debloat-Windows-10/blob/master/scripts/disable-services.ps1
	#Services to Disable
	ForEach ($service in $DisableServices) {
		If ( Get-Service -Name $service -erroraction 'silentlycontinue') {
			#Windows 10 and 2016 have hidden servcie bowser which will disalbe all SMB traffic if disabled.
			If ((Get-Service -Name $service).Name -ne "bowser") {
				write-host ("`tDisabling: " + (Get-Service -Name $service).DisplayName ) -foregroundcolor green 
				Get-Service -Name $service | Stop-Service 
				Get-Service -Name $service | Set-Service -StartupType Disabled
			}
		}
	}
	#Services to set as Manual
	ForEach ($service in $ManualServices) {
		If ( Get-Service -Name $service -erroraction 'silentlycontinue') {
			write-host ("`tManual Startup: " + (Get-Service -Name $service).DisplayName ) -foregroundcolor yellow 
			Get-Service -Name $service | Stop-Service 
			Get-Service -Name $service | Set-Service -StartupType Manual
		}
	}
	#Services to set as Automatic
	ForEach ($service in $AutomaticServices) {
		If ( Get-Service -Name $service -erroraction 'silentlycontinue') {
			write-host ("`tAutomatic Startup: " + (Get-Service -Name $service).DisplayName ) -foregroundcolor red 
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
	#Allow local user to logon to Admin Shares
	Set-Reg "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" "LocalAccountTokenFilterPolicy" 1 "DWORD"
}
#============================================================================
#endregion Main Local Machine Tweaks
#============================================================================
#============================================================================
#region Main Local Machine Certs
#============================================================================
If (-Not $UserOnly) {
	If ([int]([environment]::OSVersion.Version.Major.tostring() + [environment]::OSVersion.Version.Minor.tostring()) -gt 61) {
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
}
#============================================================================
#endregion Main Local Machine Certs
#============================================================================
#============================================================================
#region Main Local Machine Schannel for PCI
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
	If ($AllowClientTLS1 -or $Store) {
		Set-Reg ($HKSCH + "\Hashes\SHA") "Enabled" 4294967295 "DWORD"
		Write-Host "`tKeeping SHA Enabled . . ." -foregroundcolor Darkred
	}else{
		Set-Reg ($HKSCH + "\Hashes\SHA") "Enabled" 0 "DWORD"
	}
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
	If ($AllowClientTLS1 -or $Store) {
		Set-Reg ($HKSCH + "\Protocols\TLS 1.0\Client") "Enabled" 4294967295 "DWORD"
		Write-Host "`tKeeping TLS 1.0 Enabled . . ." -foregroundcolor Darkred
	}else{
		Set-Reg ($HKSCH + "\Protocols\TLS 1.0\Client") "Enabled" 0 "DWORD"
	}
	Set-Reg ($HKSCH + "\Protocols\TLS 1.0\Client") "DisabledByDefault" 1 "DWORD"
	Set-Reg ($HKSCH + "\Protocols\TLS 1.0\Server") "Enabled" 0 "DWORD"
	Set-Reg ($HKSCH + "\Protocols\TLS 1.0\Server") "DisabledByDefault" 1 "DWORD"
	If ($AllowClientTLS1 -or $Store) {
		Set-Reg ($HKSCH + "\Protocols\TLS 1.1\Client") "Enabled" 4294967295 "DWORD"
		Write-Host "`tKeeping TLS 1.1 Enabled . . ." -foregroundcolor Darkred
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
	Write-Host "Setting up .Net for TLS 1.2. . . "
	#https://jorgequestforknowledge.wordpress.com/2017/03/01/hardening-disabling-weak-ciphers-hashes-and-protocols-on-adfs-wap-aad-connect/
	#https://docs.microsoft.com/en-us/dotnet/framework/network-programming/tls
	Set-Reg ("HKLM:\SOFTWARE\Microsoft\.NETFramework\v2.0.50727") "SchUseStrongCrypto" 1 "DWORD"
	Set-Reg ("HKLM:\SOFTWARE\Microsoft\.NETFramework\v2.0.50727") "SystemDefaultTlsVersions" 1 "DWORD"
	If ([Environment]::Is64BitOperatingSystem) {
		Set-Reg ("HKLM:\SOFTWARE\Wow6432Node\Microsoft\.NETFramework\v2.0.50727") "SchUseStrongCrypto" 1 "DWORD"
		Set-Reg ("HKLM:\SOFTWARE\Wow6432Node\Microsoft\.NETFramework\v2.0.50727") "SystemDefaultTlsVersions" 1 "DWORD"
	}
	Set-Reg ("HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319") "SchUseStrongCrypto" 1 "DWORD"
	Set-Reg ("HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319") "SystemDefaultTlsVersions" 1 "DWORD"
	If ([Environment]::Is64BitOperatingSystem) {
		Set-Reg ("HKLM:\SOFTWARE\Wow6432Node\Microsoft\.NETFramework\v4.0.30319") "SchUseStrongCrypto" 1 "DWORD"
		Set-Reg ("HKLM:\SOFTWARE\Wow6432Node\Microsoft\.NETFramework\v4.0.30319") "SystemDefaultTlsVersions" 1 "DWORD"
	}
	#https://support.microsoft.com/en-us/help/3140245/update-to-enable-tls-1-1-and-tls-1-2-as-a-default-secure-protocols-in
	Set-Reg ("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\WinHttp") "DefaultSecureProtocols" 2688 "DWORD"
	If ([Environment]::Is64BitOperatingSystem) {
		Set-Reg ("HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Internet Settings\WinHttp") "DefaultSecureProtocols" 2688 "DWORD"
	}
	Set-Reg ("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\WinHttp") "SecureProtocols" 2688 "DWORD"
	If ([Environment]::Is64BitOperatingSystem) {
		Set-Reg ("HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Internet Settings\WinHttp") "SecureProtocols" 2688 "DWORD"
	}
	
}
#============================================================================
#endregion Main Local Machine Schannel for PCI
#============================================================================
#============================================================================
#region Main Local Machine User Icons
#============================================================================
If (-Not $UserOnly) {
	If ([int]([environment]::OSVersion.Version.Major.tostring() + [environment]::OSVersion.Version.Minor.tostring()) -gt 61) {
		Write-Host "Setting up User Icons . . . "
		If (Test-Path ($LICache + "\" + $Custom_User_Account_Pictures_SubFolder)) {
			copy-item ($LICache + "\" + $Custom_User_Account_Pictures_SubFolder + "\*.*") -Destination ($env:programdata + "\Microsoft\User Account Pictures") -force
			Remove-Item ($env:programdata + "\Microsoft\User Account Pictures\*.dat") -force
		}
	}
}
#============================================================================
#endregion Main Local Machine User Icons
#============================================================================
#============================================================================
#region Main Local Machine Background
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
		copy-item ($LICache + "\" + $Custom_Wallpaper_SubFolder + "\" + $BackgroundFolder + "\img0.jpg") -Destination ($env:windir + "\Web\Wallpaper\Windows\img0.jpg") -force | out-null
	}
	If (Test-Path ($LICache + "\" + $Custom_Wallpaper_SubFolder + "\" + $BackgroundFolder + "\Backgrounds")) {	
		If (-Not( Test-Path ($env:windir + "\System32\oobe\info\backgrounds\"))) {
			New-Item -ItemType directory -Path ($env:windir + "\system32\oobe\info\backgrounds") | out-null
		}
		copy-item ($LICache + "\" + $Custom_Wallpaper_SubFolder + "\" + $BackgroundFolder + "\Backgrounds\*.*") -Destination ($env:windir + "\System32\oobe\info\backgrounds\") -force | out-null
	}
	Set-Reg ("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI\Background") "OEMBackground" 1 "DWORD"
	#region Clear Lock Screen Cache
	If (Test-Path ($env:programdata + "\Microsoft\Windows\SystemData")) {
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
	}
	#endregion Clear Lock Screen Cache	
	If ([environment]::OSVersion.Version.Major -ge 10) {
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
}
#============================================================================
#endregion Main Local Machine Background
#============================================================================
#============================================================================
#region Main Local Machine Setup Windows Time
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
#endregion Main Local Machine Setup Windows Time
#============================================================================
#============================================================================
#region Main Local Machine BGInfo
#============================================================================
If (-Not $UserOnly) {
	If (-Not $NoBgInfo) {
		Write-Host "Setting up BGInfo . . . "
		If (Test-Path ($LICache + "\BgInfo")) {
			copy-item ($LICache + "\BgInfo") -Destination ($env:programfiles) -Force -Recurse
			Get-ChildItem ($env:programdata + "\Microsoft\Windows\Start Menu\Programs\StartUp") | Where-Object Name -Like "*bginfo*.lnk" | ForEach-Object { Remove-Item $_.fullname}
			If ($Store) {
				copy-item ($env:programfiles + "\BgInfo\" + $BGInfo_StartupLink_Store) ($env:programdata + "\Microsoft\Windows\Start Menu\Programs\StartUp") -Force
			}else{
				copy-item ($env:programfiles + "\BgInfo\" + $BGInfo_StartupLink) ($env:programdata + "\Microsoft\Windows\Start Menu\Programs\StartUp") -Force
			}
		}
	}
}
#============================================================================
#endregion Main Local Machine BGInfo
#============================================================================
#============================================================================
#region Main Local Machine Firewall Setup
#============================================================================
#Custom Software Firewall
If (-Not $UserOnly) {
	If ([int]([environment]::OSVersion.Version.Major.tostring() + [environment]::OSVersion.Version.Minor.tostring()) -gt 61) {
		Write-Host "Setting up Firewall . . . "
		Remove-NetFirewallRule -Group (split-path $Custom_Software_Path -Leaf ) -erroraction 'silentlycontinue'
		If (Test-Path $Custom_Software_Path) {
			Write-Host ("Adding " + (split-path $Custom_Software_Path -Leaf ) + " to Firewall...") -foregroundcolor darkgray
			Get-ChildItem -Path $Custom_Software_Path -Filter *.exe -Recurse| ForEach-Object {
				Write-Host ("`t Adding rule for: " + $_.Name) -foregroundcolor yellow
				New-NetFirewallRule -DisplayName $_.Name -Direction Inbound -Program $_.VersionInfo.FileName -Group (split-path $Custom_Software_Path -Leaf ) -Action Allow | out-null
			}
		}
	}
	If ([environment]::OSVersion.Version.Major -ge 10) {
		Write-Host "Disabling un-needed Firewall Rules . . . " -foregroundcolor darkgray
		Disable-NetFirewallRule -DisplayGroup "AllJoyn Router" -erroraction 'silentlycontinue' | out-null
		Disable-NetFirewallRule -DisplayGroup "Cast to Device functionality" -erroraction 'silentlycontinue' | out-null
		Disable-NetFirewallRule -DisplayGroup "cortana" -erroraction 'silentlycontinue' | out-null
		Disable-NetFirewallRule -DisplayGroup "Media Center Extenders" -erroraction 'silentlycontinue' | out-null
		Disable-NetFirewallRule -DisplayGroup "Proximity Sharing" -erroraction 'silentlycontinue' | out-null
		Disable-NetFirewallRule -DisplayGroup "Routing and Remote Access" -erroraction 'silentlycontinue' | out-null
		Disable-NetFirewallRule -DisplayGroup "Wi-Fi Direct Network Discovery" -erroraction 'silentlycontinue' | out-null
		Disable-NetFirewallRule -DisplayGroup "Windows Media Player Network Sharing Service" -erroraction 'silentlycontinue' | out-null
		Disable-NetFirewallRule -DisplayGroup "Work or school account" -erroraction 'silentlycontinue' | out-null
		Disable-NetFirewallRule -DisplayGroup "Your account" -erroraction 'silentlycontinue' | out-null
		Disable-NetFirewallRule -DisplayGroup "iSCSI Service" -erroraction 'silentlycontinue' | out-null
		Disable-NetFirewallRule -DisplayGroup "Xbox Game UI" -erroraction 'silentlycontinue' | out-null
		Disable-NetFirewallRule -DisplayGroup "Network Discovery" -erroraction 'silentlycontinue' | out-null
	}
	Write-Host ("`t Adding rule for: Ping")
	New-NetFirewallRule -DisplayName "Allow inbound ICMPv4" -Direction Inbound -Protocol ICMPv4 -IcmpType 8  -Action Allow | out-null
}
#============================================================================
#endregion Main Local Machine Firewall Setup
#============================================================================
#============================================================================
#regionMain Local Machine Log and Performance Monitoring
#============================================================================


#============================================================================
#endregion Main Local Machine Log and Performance Monitoring
#============================================================================
#============================================================================
#region Main Local Machine FileShares
#============================================================================

#============================================================================
#endregion Main Local Machine FileShares
#============================================================================
#region Main Local Machine All Users Desktop
#============================================================================
If (-Not $UserOnly) {
	If (-Not (Test-Path ($env:Public + "\Desktop\Internet Explorer.lnk"))) {
		If ( Test-Path ($env:appdata + "\Microsoft\Windows\Start Menu\Programs\Accessories\Internet Explorer.lnk")) {
			Write-Host "Adding Internet Explorer to All Users Desktop"
			copy-item ($env:appdata + "\Microsoft\Windows\Start Menu\Programs\Accessories\Internet Explorer.lnk") ($env:Public + "\Desktop\Internet Explorer.lnk")
		}
	}
	#Add other Icons to all users desktop.
	If (Test-Path($LICache + "\Desktop")) {
		Copy-Item -Force -Recurse -Path ($LICache + "\Desktop") -Destination ($env:Public)
	}
	#Copy Custom Icons.
	If (Test-Path($LICache + "\icons")) {
		If (-Not (Test-path($Custom_Icon_Path))) {
			New-Item -ItemType Directory -Force -Path $Custom_Icon_Path
		}
		Copy-Item -Force -Recurse -Path ($LICache + "\icons\*") -Destination ($Custom_Icon_Path)
	}
}
#============================================================================
#endregion Main Local Machine All Users Desktop
#============================================================================
#============================================================================
#region Main Local Machine RDP
#============================================================================
#RDP
If (-Not $UserOnly) {
	Write-Host "Enabling RDP . . . "
	Set-Reg "HKLM:\SYSTEM\CurrentControlSet\control\Terminal Server" "fDenyTSConnections " 0 "DWORD"
	# Set-Reg "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp" "UserAuthentication" 0 "DWORD"
}
#============================================================================
#endregion Main Local Machine RDP
#============================================================================
#============================================================================
#region Main Local Machine Setup Screen Saver
#============================================================================
If (-Not $UserOnly) {
	Write-Host "Setup Logon Screen Saver . . ."
	Set-Reg "HKU:\.DEFAULT\Control Panel\Desktop" "ScreenSaveActive" "1" "STRING"
	Set-Reg "HKU:\.DEFAULT\Control Panel\Desktop" "ScreenSaverIsSecure" "1" "STRING"
	Set-Reg "HKU:\.DEFAULT\Control Panel\Desktop" "ScreenSaveTimeOut" "600" "STRING"
	Set-Reg "HKU:\.DEFAULT\Control Panel\Desktop" "SCRNSAVE.EXE" "C:\Windows\system32\scrnsave.scr" "STRING"
}

#============================================================================
#endregion Main Local Machine Setup Screen Saver
#============================================================================
#============================================================================
#region Main Local Machine Microsoft Store
#============================================================================
#Disable MS Apps
If (-Not $UserOnly) {
	If ([int]([environment]::OSVersion.Version.Major.tostring() + [environment]::OSVersion.Version.Minor.tostring()) -gt 61) {
		#region Get list of currently installed and provisioned Appx packages
		$AllInstalled = Get-AppxPackage -AllUsers | ForEach-Object {$_.Name}
		$AllProvisioned = Get-ProvisionedAppxPackage -Online | ForEach-Object {$_.DisplayName}
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
		ForEach($Appx in $AllInstalled){
			$error.Clear()
			If(-Not $Keep.Contains([system.String]::Join(".", ($Appx.split(".") |  ForEach-Object {if (($_ -as [int] -eq $null )) {$_ }})))){
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
		ForEach($Appx in $AllProvisioned){
			$error.Clear()
			If(-Not $Keep.Contains([system.String]::Join(".", ($Appx.split(".") |  ForEach-Object {if (($_ -as [int] -eq $null )) {$_ }})))){
				Try{
					Get-ProvisionedAppxPackage -Online | Where-Object {$_.DisplayName -eq $Appx} | Remove-ProvisionedAppxPackage -Online | Out-Null
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
#endregion Main Local Machine Microsoft Store
#============================================================================
#============================================================================
#region Main Local Machine Remove OneDrive
#============================================================================
If (-Not $UserOnly) {
	$process = Start-Process -FilePath "taskkill" -ArgumentList @("/f","/im","OneDrive.exe")
	#https://social.technet.microsoft.com/Forums/ie/en-US/2eaa1b6a-c906-4161-b76c-370ac8910a11/windows-10-sysprep-issue-image-always-hangs-at-quotgetting-readyquot?forum=win10itprosetup
	If (Test-Path ($env:systemroot + "\SysWOW64\OneDriveSetup.exe")) {
		Write-Host "Removing OneDrive . . ."
		$process = Start-Process -FilePath ('"'+ $env:systemroot + "\SysWOW64\OneDriveSetup.exe" + '"') -ArgumentList @("/uninstall","/quiet") -PassThru -NoNewWindow -Wait
	}
	If (Test-Path ($env:systemroot + "\System32\OneDriveSetup.exe")) {
		Write-Host "Removing OneDrive . . ."
		$process = Start-Process -FilePath ('"'+ $env:systemroot + "\System32\OneDriveSetup.exe" + '"') -ArgumentList @("/uninstall","/quiet") -PassThru -NoNewWindow -Wait
	}
	Remove-Item -Recurse -Force -Path ($env:userprofile + "\OneDrive") -erroraction 'silentlycontinue'| out-null
	Remove-Item -Recurse -Force -Path ($env:localappdata + "\Microsoft\OneDrive") -erroraction 'silentlycontinue'| out-null
	Remove-Item -Recurse -Force -Path ($env:programdata + "\Microsoft\OneDrive") -erroraction 'silentlycontinue'| out-null
	Remove-Item -Recurse -Force -Path ("HKCR:\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}") -erroraction 'silentlycontinue'| out-null
	If ([Environment]::Is64BitOperatingSystem) {
		Remove-Item -Recurse -Force -Path ("HKCR:\Wow6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}") -erroraction 'silentlycontinue'| out-null
	}
	If (-Not (Test-Path ("HKLM:\SOFTWARE\Policies\Microsoft\Windows\Skydrive"))) {
		New-Item -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\' -Name 'Skydrive' | Out-Null
		New-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Skydrive' -Name 'DisableFileSync' -PropertyType DWORD -Value '1' | Out-Null
		New-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Skydrive' -Name 'DisableLibrariesDefaultSaveToSkyDrive' -PropertyType DWORD -Value '1' | Out-Null 
	}
	#Removes OneDrive from This PC
	write-host ("`tOneDrive from This PC ") -foregroundcolor "gray"
	If (Test-Path ("HKCR:\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}")) {
		Set-Reg "HKCR:\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}" "System.IsPinnedToNameSpaceTree" 0 "DWORD"
		If ([Environment]::Is64BitOperatingSystem) {
			Set-Reg "HKCR:\WOW6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}" "System.IsPinnedToNameSpaceTree" 0 "DWORD"
		}
	}
	If(Test-Path ($HKLWE + "\MyComputer\NameSpace\{018D5C66-4533-4307-9B53-224DE2ED1FE6}")) {
		Remove-Item ($HKLWE + "\MyComputer\NameSpace\{018D5C66-4533-4307-9B53-224DE2ED1FE6}") -Recurse | Out-Null
	}
}
#============================================================================
#endregion Main Local Machine Remove OneDrive
#============================================================================
#============================================================================
#region Main Local Machine Set OEM Info
#============================================================================
If (-Not $UserOnly) {
	If ($NoOEMInfo) {
		#$Bios_Info = Get-CimInstance -ClassName Win32_BIOS
		Write-Host "Setup System OEM Info . . ."
		If (-Not $IsVM) {
			Set-Reg "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation" "Manufacturer" ((Get-CimInstance -ClassName Win32_ComputerSystem).Manufacturer) "String"
			If ($OEMInfoAddSerial) {
				Set-Reg "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation" "Model" ((Get-CimInstance -ClassName Win32_ComputerSystem).model + " (Serial Number: " + (Get-CimInstance -ClassName Win32_BIOS).SerialNumber + ")") "String"
			}else{
				Set-Reg "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation" "Model" ((Get-CimInstance -ClassName Win32_ComputerSystem).model) "String"
			}
		}
		If (-Not (Test-Path ($env:windir + "\system32\oobe\info\"))) {
			New-Item -ItemType directory -Path ($env:windir + "\system32\oobe\info\") | out-null
		}
		Copy-Item  ($LICache + "\" + $Custom_Wallpaper_SubFolder + "\" + $Custom_OEM_Logo) -Destination ($env:windir + "\system32\oobe\info\" + $Custom_OEM_Logo ) -Recurse -Force
		Set-Reg "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation" "Logo" ($env:windir + "\system32\oobe\info\" + $Custom_OEM_Logo ) "String"
	}
}
#============================================================================
#endregion Main Local Machine Set OEM Info
#============================================================================
#============================================================================
#region Main Local Machine FortiClient
#============================================================================
If (-Not $UserOnly) {
	If (Test-Path ($LICache + "\RemoveFCTID.exe")) {
		Write-Host ("Setting up RemoveFCTID Shortcut")
		If ((Test-Path ($env:USERPROFILE + "\Desktop")) -and -Not (Test-Path($env:USERPROFILE + "\Desktop\RemoveFCTID.lnk"))){
			If (Test-Path ($LICache + "\RemoveFCTID.exe")) {				
				$ShortCut = $WScriptShell.CreateShortcut($env:USERPROFILE + "\Desktop\RemoveFCTID.lnk")
				$ShortCut.TargetPath=($LICache + "\RemoveFCTID.exe")
				$ShortCut.WorkingDirectory = ($env:ProgramFiles + "\Fortinet\FortiClient")
				$ShortCut.Hotkey = "CTRL+SHIFT+F"
				$ShortCut.IconLocation = "%SystemRoot%\System32\imageres.dll, 100"
				$ShortCut.Description = "Run Before Imaging"
				$ShortCut.Save()
				#Make ShortCut ran as admin https://stackoverflow.com/questions/28997799/how-to-create-a-run-as-administrator-shortcut-using-powershell
				$bytes = [System.IO.File]::ReadAllBytes($env:USERPROFILE + "\Desktop\RemoveFCTID.lnk")
				$bytes[0x15] = $bytes[0x15] -bor 0x20 #set byte 21 (0x15) bit 6 (0x20) ON
				[System.IO.File]::WriteAllBytes($env:USERPROFILE + "\Desktop\RemoveFCTID.lnk", $bytes)
			} else {
				Write-Warning "Copy failed please manually copy and create shortcut."
			}
		}
		If ((Test-Path ($UsersProfileFolder + "\administrator\Desktop")) -and -Not (Test-Path($UsersProfileFolder + "\administrator\Desktop\RemoveFCTID.lnk"))) {			If (Test-Path ($env:ProgramFiles + "\Fortinet\FortiClient\RemoveFCTID.exe")) {
				$ShortCut = $WScriptShell.CreateShortcut($UsersProfileFolder + "\administrator\Desktop\RemoveFCTID.lnk")
				$ShortCut.TargetPath=($LICache + "\RemoveFCTID.exe")
				$ShortCut.WorkingDirectory = ($env:ProgramFiles + "\Fortinet\FortiClient")
				$ShortCut.Hotkey = "CTRL+SHIFT+F"
				$ShortCut.IconLocation = "%SystemRoot%\System32\imageres.dll, 100"
				$ShortCut.Description = "Run Before Imaging"
				$ShortCut.Save()
				#Make ShortCut ran as admin https://stackoverflow.com/questions/28997799/how-to-create-a-run-as-administrator-shortcut-using-powershell
				$bytes = [System.IO.File]::ReadAllBytes($UsersProfileFolder + "\administrator\Desktop\RemoveFCTID.lnk")
				$bytes[0x15] = $bytes[0x15] -bor 0x20 #set byte 21 (0x15) bit 6 (0x20) ON
				[System.IO.File]::WriteAllBytes($UsersProfileFolder + "\administrator\Desktop\RemoveFCTID.lnk", $bytes)
			} else {
				Write-Warning "Copy failed please manually copy and create shortcut."
			}
		}
		If (Test-Path ($LICache + "\RemoveFCTID.exe")) {
			Write-Host "Running FortiClient ID Cleanup"
			$process = Start-Process -FilePath ('"' + $LICache + "\RemoveFCTID.exe" + '"') -PassThru -NoNewWindow -Wait
		}
	}
}
#============================================================================
#endregion Main Local Machine FortiClient
#============================================================================
#============================================================================
#region Main Local Machine Disable Netbios
#============================================================================
If (-Not $UserOnly) {
	#https://community.spiceworks.com/topic/2010972-disable-netbios-over-tcp-ip-using-gpo-in-ad-environment 
	Write-Host ("Disabling Netbios...") -foregroundcolor darkgray
	$key = "HKLM:SYSTEM\CurrentControlSet\services\NetBT\Parameters\Interfaces"
	Get-ChildItem $key |
	ForEach-Object { Set-ItemProperty -Path "$key\$($_.pschildname)" -Name NetbiosOptions -Value 2 -Verbose}
	If (-Not $IPv6) {
		Write-Host ("Disabling IPv6...") -foregroundcolor darkgray
		#https://directaccess.richardhicks.com/2013/08/27/disabling-unused-ipv6-transition-technologies-for-directaccess-clients/
		Set-Net6to4Configuration -State disabled
		Set-NetTeredoConfiguration -Type disabled
		Set-NetIsatapConfiguration -State disabled

		#Disabled IPv6 in all interfaces
		Get-NetAdapterBinding -DisplayName "Internet Protocol Version 6 (TCP/IPv6)" | Set-NetAdapterBinding -Enabled:$false
	}
}
#============================================================================
#endregion Main Local Machine Disable Netbios
#============================================================================
#============================================================================
#region SNMP Setup
#============================================================================
If ($Store) {
	Write-Host ("Setting up SNMP...") -foregroundcolor darkgray
	If (-Not (Get-Service -Name "SNMP" | Out-null)) {
		Write-Host ("`tInstalling up SNMP...") -foregroundcolor darkgray
		Enable-WindowsOptionalFeature -online -FeatureName "SNMP" -NoRestart
		Get-Service -Name "SNMP" | Set-Service -StartupType Disabled 
	}
	(get-item "HKLM:SYSTEM\CurrentControlSet\Services\SNMP\Parameters\ValidCommunities").property| ForEach-Object { Remove-ItemProperty -Name $_ -Path "HKLM:SYSTEM\CurrentControlSet\Services\SNMP\Parameters\ValidCommunities"}
	#Set Community
	Set-Reg "HKLM:SYSTEM\CurrentControlSet\Services\SNMP\Parameters\ValidCommunities" $SNMPValue 4 "DWORD"
	#Sets All info
	Set-Reg "HKLM:\SYSTEM\CurrentControlSet\Services\SNMP\Parameters\RFC1156Agent" "sysServices" 79 "DWORD"
	#Allows All hosts
	(get-item "HKLM:\SYSTEM\CurrentControlSet\Services\SNMP\Parameters\PermittedManagers").property| ForEach-Object { Remove-ItemProperty -Name $_ -Path "HKLM:\SYSTEM\CurrentControlSet\Services\SNMP\Parameters\PermittedManagers"}

}
#============================================================================
#endregion SNMP Setup
#============================================================================
#============================================================================
#region Set Registry Permissions
#============================================================================
Write-Host ("Setting Registry Permissions ... ") -foregroundcolor darkgray
	ForEach ( $item in $RegPerms) {
		# Source: https://social.technet.microsoft.com/Forums/en-US/1f082309-dc39-4c7e-ab45-b19094c21877/powershell-script-to-change-permission-of-hkcu-registry-and-make-it-read-only-permission-for-the?forum=winserverpowershell
		Write-Host ("`tUpdating: '" +  $rootKey + ":\" + $item.Key + "' for '" + $item.User + "' to '" + $item.Action + "' with '" + $item.Perm + "'")	
		switch -regex ($item.Hive) {
        'HKCU|HKEY_CURRENT_USER'    { $rootKey = 'HKCU' }
        'HKLM|HKEY_LOCAL_MACHINE'   { $rootKey = 'HKLM' }
        'HKCR|HKEY_CLASSES_ROOT'    { $rootKey = 'HKCR' }
        'HKCC|HKEY_CURRENT_CONFIG'  { $rootKey = 'HKCC' }
        'HKU|HKEY_USERS'            { $rootKey = 'HKU' }
		}
		$path = ($rootKey + ":\" + $item.Key)
		If(!(Test-Path $path)) {
			New-Item -Path $path -Force | Out-Null
		}
		Set-KeyOwnership $item.Hive $item.Key
		$Acl = Get-ACL $path
		$AccessRule= New-Object System.Security.AccessControl.RegistryAccessRule($item.User,$item.Perm,$item.Action)
		$Acl.SetAccessRule($AccessRule)
		Set-Acl $path $Acl
	}
#============================================================================
#endregion Set Registry Permissions
#============================================================================
#============================================================================
#region Main Local Machine Load Local GPO
#============================================================================
If (-Not $UserOnly) {
	Write-Host "Setting Machine Policy . . ."
	#Set User Account Lockout
	Set-Reg "HKLM:\SYSTEM\CurrentControlSet\Services\RemoteAccess\Parameters\AccountLockout" "MaxDenials" "6" "DWord"
	Set-Reg "HKLM:\SYSTEM\CurrentControlSet\Services\RemoteAccess\Parameters\AccountLockout" "ResetTime (mins)" "15" "DWord"
	#Wi-Fi Sense must be disabled.
	Set-Reg "HKLM:\Software\Microsoft\wcmsvc\wifinetworkmanager\config" "AutoConnectAllowedOEM" "0" "DWord"
	#Configure the default autorun behavior to prevent autorun commands.
	Set-Reg "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\CredUI" "EnumerateAdministrators" "0" "DWord"
	#Configure the default autorun behavior to prevent autorun commands.
	Set-Reg "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer" "NoAutorun" "1" "DWord"
	#Autoplay must be disabled for all drives.
	Set-Reg "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer" "NoDriveTypeAutoRun" "255" "DWord"
	#Disable Internet File Association Service.
	Set-Reg "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer" "NoInternetOpenWith" "1" "DWord"
	#The Order Prints Online wizard must be turned off.
	Set-Reg "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer" "NoOnlinePrintsWizard" "1" "DWord"
	#Turn off the "Publish to Web" task for files and folders
	Set-Reg "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer" "NoPublishingWizard" "1" "DWord"
	#Turn off Internet download for Web publishing and online ordering wizards
	Set-Reg "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer" "NoWebServices" "1" "DWord"
	#File Explorer shell protocol must run in protected mode.
	Set-Reg "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer" "PreXPSP2ShellProtocolBehavior" "0" "DWord"
	#Force script to run one at a time
	Set-Reg "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer" "AsyncRunOnce" "0" "DWord"
	#Add-on performance notifications must be disallowed.
	Set-Reg "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Ext" "DisableAddonLoadTimePerformanceNotifications" "1" "DWord"
	#disable add-on allowing prompting
	Set-Reg "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Ext" "IgnoreFrameApprovalCheck" "1" "DWord"
	#ActiveX opt-in prompt must be disallowed.
	Set-Reg "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Ext" "NoFirsttimeprompt" "1" "DWord"
	#Automatically signing in the last interactive user after a system-initiated restart must be disabled
	Set-Reg "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System" "DisableAutomaticRestartSignOn" "1" "DWord"
	#Turn off Windows Startup Sound
	Set-Reg "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System" "DisableStartupSound" "1" "DWord"
	#Disable First Time Sign-in Animation
	Set-Reg "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System" "EnableFirstLogonAnimation" "0" "DWord"
	#Local administrator accounts must have their privileged token filtered to prevent elevated privileges from being used over the network on domain systems.
	Set-Reg "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System" "LocalAccountTokenFilterPolicy" "0" "DWord"
	#The setting to allow Microsoft accounts to be optional for modern style apps must be enabled
	Set-Reg "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System" "MSAOptional" "1" "DWord"
	#The Welcome screen may be displayed for 30 seconds, and the logon script interacts with me when I try to log on
	Set-Reg "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System" "DelayedDesktopSwitchTimeout" "0" "DWord"
	#Command line data must be included in process creation events.
	Set-Reg "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System\Audit" "ProcessCreationIncludeCmdLine_Enabled" "1" "DWord"
	#Encryption Oracle Remediation - Force Updated Clients 
	#https://getadmx.com/?Category=Windows_10_2016&Policy=Microsoft.Policies.CredentialsSSP::AllowEncryptionOracle
	Set-Reg "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System\CredSSP\Parameters" "AllowEncryptionOracle" "0" "DWord"
	#Enhanced anti-spoofing for facial recognition must be enabled 
	Set-Reg "HKLM:\Software\Policies\Microsoft\Biometrics\FacialFeatures" "EnhancedAntiSpoofing" "1" "DWord"
	#Event Viewer Events.asp links must be turned off
	Set-Reg "HKLM:\Software\Policies\Microsoft\EventViewer" "MicrosoftEventVwrDisableLinks" "1" "DWord"
	#Microsoft services to provide enhanced suggestions as the user types in the Address bar must be disallowed 
	Set-Reg "HKLM:\Software\Policies\Microsoft\Internet Explorer" "AllowServicePoweredQSA" "0" "DWord"
	#Basic authentication for RSS feeds over HTTP must be turned off
	Set-Reg "HKLM:\Software\Policies\Microsoft\Internet Explorer\Feeds" "AllowBasicAuthInClear" "0" "DWord"
	#Prevent RSS attachment downloads
	Set-Reg "HKLM:\Software\Policies\Microsoft\Internet Explorer\Feeds" "DisableEnclosureDownload" "1" "DWord"
	#Force showing Link Bar in IE
	Set-Reg "HKLM:\Software\Policies\Microsoft\Internet Explorer\LinksBar" "Enabled" "1" "DWord"
	#Disable Internet Explorer First Run Welcome Screen
	Set-Reg "HKLM:\Software\Policies\Microsoft\Internet Explorer\Main" "DisableFirstRunCustomize" "1" "DWord"
	#Hide the status bar
	Set-Reg "HKLM:\Software\Policies\Microsoft\Internet Explorer\Main" "StatusBarWeb" "1" "DWord"
	#Turn off suggestions for all user-installed providers
	Set-Reg "HKLM:\Software\Policies\Microsoft\Internet Explorer\SearchScopes" "ShowSearchSuggestionsGlobal" "0" "DWord"
	#Disable security warning "Your current security settings put your computer at risk"
	Set-Reg "HKLM:\Software\Policies\Microsoft\Internet Explorer\Security" "DisableFixSecuritySettings" "1" "DWord"
	#Disable Suggested Sites
	Set-Reg "HKLM:\Software\Policies\Microsoft\Internet Explorer\Suggested Sites" "Enabled" "0" "DWord"
	#Turn off the Windows Messenger Customer Experience Improvement Program
	Set-Reg "HKLM:\Software\Policies\Microsoft\Messenger\Client" "CEIP" "2" "DWord"
	#Disable InPrivate browsing
	Set-Reg "HKLM:\Software\Policies\Microsoft\MicrosoftEdge\Main" "AllowInPrivate" "0" "DWord"
	#The password manager function in the Edge browser must be disabled.
	Set-Reg "HKLM:\Software\Policies\Microsoft\MicrosoftEdge\Main" "FormSuggest Passwords" "no" "String"
	#The SmartScreen filter for Microsoft Edge must be enabled.
	Set-Reg "HKLM:\Software\Policies\Microsoft\MicrosoftEdge\PhishingFilter" "EnabledV9" "1" "DWord"
	#Users must not be allowed to ignore SmartScreen filter warnings for malicious websites in Microsoft Edge
	Set-Reg "HKLM:\Software\Policies\Microsoft\MicrosoftEdge\PhishingFilter" "PreventOverride" "1" "DWord"
	#Users must not be allowed to ignore Windows Defender SmartScreen filter warnings for unverified files in Microsoft Edge.
	Set-Reg "HKLM:\Software\Policies\Microsoft\MicrosoftEdge\PhishingFilter" "PreventOverrideAppRepUnknown" "1" "DWord"
	#Windows 10 must be configured to prevent Microsoft Edge browser data from being cleared on exit.
	Set-Reg "HKLM:\Software\Policies\Microsoft\MicrosoftEdge\Privacy" "ClearBrowsingHistoryOnExit" "0" "DWord"
	#Windows 10 must be configured to require a minimum pin length of six characters or greater
	Set-Reg "HKLM:\Software\Policies\Microsoft\PassportForWork\PINComplexity" "MinimumPINLength" "6" "DWord"
	#The system must be configured to prevent automatic forwarding of error information.
	Set-Reg "HKLM:\Software\Policies\Microsoft\PCHealth\ErrorReporting" "DoReport" "0" "DWord"
	#Turn off Help and Support Center "Did you know?" content
	Set-Reg "HKLM:\Software\Policies\Microsoft\PCHealth\HelpSvc" "Headlines" "0" "DWord"
	#Turn off Help and Support Center Microsoft Knowledge Base search
	Set-Reg "HKLM:\Software\Policies\Microsoft\PCHealth\HelpSvc" "MicrosoftKBSearch" "0" "DWord"
	If ($store) {
		#Do not prompt for password after sleep
		Set-Reg "HKLM:\Software\Policies\Microsoft\Power\PowerSettings\0e796bdb-100d-47d6-a2d5-f7d2daa51f51" "ACSettingIndex" "0" "DWord"
		Set-Reg "HKLM:\Software\Policies\Microsoft\Power\PowerSettings\0e796bdb-100d-47d6-a2d5-f7d2daa51f51" "DCSettingIndex" "0" "DWord"
	}else{
		#Prompt for password after sleep
		Set-Reg "HKLM:\Software\Policies\Microsoft\Power\PowerSettings\0e796bdb-100d-47d6-a2d5-f7d2daa51f51" "ACSettingIndex" "1" "DWord"
		Set-Reg "HKLM:\Software\Policies\Microsoft\Power\PowerSettings\0e796bdb-100d-47d6-a2d5-f7d2daa51f51" "DCSettingIndex" "1" "DWord"		
	}
	#Search Companion prevented from automatically downloading content updates.
	Set-Reg "HKLM:\Software\Policies\Microsoft\SearchCompanion" "DisableContentFileUpdates" "1" "DWord"
	#Disable Windows Customer Experience Improvement Program
	Set-Reg "HKLM:\Software\Policies\Microsoft\SQMClient\Windows" "CEIPEnable" "0" "DWord"
	#Users must be prevented from making changes to Exploit Protection settings in the Windows Defender Security Center on Windows 10.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows Defender Security Center\App and Browser protection" "DisallowExploitProtectionOverride" "1" "DWord"
	#Printing over HTTP must be prevented.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows NT\Printers" "DisableHTTPPrinting" "1" "DWord"
	#Downloading print driver packages over HTTP must be prevented.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows NT\Printers" "DisableWebPnPDownload" "1" "DWord"
	#Unauthenticated RPC clients must be restricted from connecting to the RPC server.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows NT\Rpc" "RestrictRemoteClients" "1" "DWord"
	#Passwords must not be saved in the Remote Desktop Client.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows NT\Terminal Services" "DisablePasswordSaving" "1" "DWord"
	#Solicited Remote Assistance must not be allowed.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows NT\Terminal Services" "fAllowToGetHelp" "0" "DWord"
	#Local drives prevented from sharing with Terminal Servers (Terminal Server Role).
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows NT\Terminal Services" "fDisableCdm" "1" "DWord"
	#The Remote Desktop Session Host must require secure RPC communications.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows NT\Terminal Services" "fEncryptRPCTraffic" "1" "DWord"
	#Remote Desktop Services must be configured with the client connection encryption set to the required level.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows NT\Terminal Services" "MinEncryptionLevel" "3" "DWord"
	#Specifies that the Transport Layer Security (TLS) protocol is used by the server and the client for authentication before a remote desktop connection is established.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows NT\Terminal Services" "SecurityLayer" "2" "DWord"
	#Prevent the Application Compatibility Program Inventory from collecting data and sending the information to Microsoft.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\AppCompat" "DisableInventory" "1" "DWord"
	#Microsoft consumer experiences must be turned off.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\CloudContent" "DisableWindowsConsumerFeatures" "1" "DWord"
	#Do not show Windows Tips
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\CloudContent" "DisableSoftLanding" "1" "DWord"
	#Windows 10 must be configured to enable Remote host allows delegation of non-exportable credentials
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\CredentialsDelegation" "AllowProtectedCreds" "1" "DWord"
	#Set Telemetry to Security [Enterprise Only]
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\DataCollection" "AllowTelemetry" "0" "DWord"
	#Windows Update must not obtain updates from other PCs on the Internet.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\DeliveryOptimization" "DODownloadMode" "0" "DWord"
	#The Application event log size must be configured to 32,768 KB or greater.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\EventLog\Application" "MaxSize" "32768" "DWord"
	#The Security event log must be configured to a minimum size requirement.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\EventLog\Security" "MaxSize" "1024000" "DWord"
	#Windows event log sizes must meet minimum requirements.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\EventLog\Setup" "MaxSize" "32768" "DWord"
	#Windows event log sizes must meet minimum requirements.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\EventLog\System" "MaxSize" "32768" "DWord"
	#Turn off autoplay for non-volume devices.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\Explorer" "NoAutoplayfornonVolume" "1" "DWord"
	#Explorer Data Execution Prevention must be enabled.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\Explorer" "NoDataExecutionPrevention" "0" "DWord"
	#Disable heap termination on corruption in Windows Explorer.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\Explorer" "NoHeapTerminationOnCorruption" "0" "DWord"
	#Access to the Windows Store must be turned off.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\Explorer" "NoUseStoreOpenWith" "1" "DWord"
	#Disables "Recently added" Apps List on the Start Menu for All Users
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\Explorer" "HideRecentlyAddedApps" "1" "DWord"
	#Windows 10 must be configured to disable Windows Game Recording and Broadcasting.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\GameDVR" "AllowGameDVR" "0" "DWord"
	#Disable background processing of Registry Policy in Windows
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\Group Policy\{35378EAC-683F-11D2-A89A-00C04FBBCFA2}" "NoBackgroundPolicy" "0" "DWord"
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\Group Policy\{35378EAC-683F-11D2-A89A-00C04FBBCFA2}" "NoGPOListChanges" "0" "DWord"
	#Handwriting recognition error reports (Tablet PCs) are not sent to Microsoft.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\HandwritingErrorReports" "PreventHandwritingErrorReports" "1" "DWord"
	#The Windows Installer Always install with elevated privileges option must be disabled.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\Installer" "AlwaysInstallElevated" "0" "DWord"
	#Prevent users from changing Windows installer options.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\Installer" "EnableUserControl" "0" "DWord"
	#IE security prompt is enabled for web-based installations.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\Installer" "SafeForScripting" "0" "DWord"
	#The Internet Connection Wizard cannot download a list of ISPs from Microosft.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\Internet Connection Wizard" "ExitOnMSICW" "1" "DWord"
	#Insecure logons to an SMB server must be disabled.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\LanmanWorkstation" "AllowInsecureGuestAuth" "0" "DWord"
	#Prohibit Internet Connection Sharing
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\Network Connections" "NC_ShowSharedAccessUI" "0" "DWord"
	#Hardened UNC Paths must be defined to require mutual authentication and integrity for at least the \\*\SYSVOL and \\*\NETLOGON shares.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\NetworkProvider\HardenedPaths" "\\*\NETLOGON" "RequireMutualAuthentication=1 RequireIntegrity=1" "String"
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\NetworkProvider\HardenedPaths" "\\*\SYSVOL" "RequireMutualAuthentication=1 RequireIntegrity=1" "String"
	#The use of OneDrive for storage must be disabled.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\OneDrive" "DisableFileSyncNGSC" "1" "DWord"
	#Force a specific default lock screen and logon image
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\Personalization" "LockScreenImage" "C:\Windows\System32\oobe\info\backgrounds\background1920×1200.jpg" "String"
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\Personalization" "LockScreenOverlaysDisabled" "1" "DWord"
	#Disable Changing Lock Screen Background
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\Personalization" "NoChangingLockScreen" "1" "DWord"
	#Camera access from the lock screen must be disabled. 
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\Personalization" "NoLockScreenCamera" "1" "DWord"
	#The display of slide shows on the lock screen must be disabled. 
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\Personalization" "NoLockScreenSlideshow" "1" "DWord"	
	#PowerShell script block logging must be enabled.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging" "EnableScriptBlockLogging" "1" "DWord"
	#Windows Registration Wizard must be turned off.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\Registration Wizard Control" "NoRegistration" "1" "DWord"
	#Disable Allow users to select when a password is required when resuming from connected standby
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\System" "AllowDomainDelayLock" "0" "DWord"
	#Signing in using a PIN must be turned off.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\System" "AllowDomainPINLogon" "0" "DWord"
	#Disable Domain Users Sign-in using Picture Password 
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\System" "BlockDomainPicturePassword" "1" "DWord"
	#App notifications on the lock screen must be turned off.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\System" "DisableLockScreenAppNotifications" "1" "DWord"
	#The network selection user interface (UI) must not be displayed on the logon screen.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\System" "DontDisplayNetworkSelectionUI" "1" "DWord"
	#The Windows Defender SmartScreen for Explorer must be enabled.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\System" "EnableSmartScreen" "1" "DWord"
	#Local users on domain-joined computers must not be enumerated.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\System" "EnumerateLocalUsers" "0" "DWord"
	#The Windows Defender SmartScreen for Explorer must be enabled.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\System" "ShellSmartScreenLevel" "Block" "String"
	#Prevent handwriting personalization data sharing with Microsoft.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\TabletPC" "PreventHandwritingDataSharing" "1" "DWord"
	#Change Prohibit connection to roaming Mobile Broadband networks 
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\WcmSvc\GroupPolicy" "fBlockRoaming" "1" "DWord"
	#Simultaneous connections to the Internet or a Windows domain must be limited.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\WcmSvc\GroupPolicy" "fMinimizeConnections" "1" "DWord"
	#Turn off Windows Error Reporting to Microsoft.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\Windows Error Reporting" "Disabled" "1" "DWord"
	#Disable Cortana
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\Windows Search" "AllowCortana" "0" "DWord"
	#Disable Cortana on Lock Screen in Windows 10
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\Windows Search" "AllowCortanaAboveLock" "0" "DWord"
	#Disable Cortana Page in OOBE on an AAD account
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\Windows Search" "AllowCortanaInAAD" "0" "DWord"
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\Windows Search" "AllowCortanaInAADPathOOBE" "0" "DWord"
	#Indexing of encrypted files must be turned off.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\Windows Search" "AllowIndexingEncryptedStoresOrItems" "0" "DWord"
	#Disable OS Upgrade for  Windows 10
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate" "DisableOSUpgrade" "1" "DWord"
	Set-Reg "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\OSUpgrade" "AllowOSUpgrade" "0" "DWord"
	#Systems take Feature Updates from Semi-annual Channel 
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate" "BranchReadinessLevel" "32" "DWord"
	#Defer feature updates
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate" "DeferFeatureUpdates" "1" "DWord"
	# Wait a year after new feature comes out before installing it. 
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate" "DeferFeatureUpdatesPeriodInDays" "365" "DWord"
	# Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate" "DeferFeatureUpdatesPeriodInDays" "0" "DWord"
	#Defer Quality Updates
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate" "DeferQualityUpdates" "1" "DWord"
	# Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate" "DeferQualityUpdatesPeriodInDays" "30" "DWord"
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate" "DeferQualityUpdatesPeriodInDays" "0" "DWord"
	#Manage Preview Builds
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate" "ManagePreviewBuilds" "1" "DWord"
	#Disable Preview Updates
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate" "ManagePreviewBuildsPolicyValue" "0" "DWord"
	#Enable Microsoft Updates 
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate\AU" "AllowMUUpdateService" "1" "DWord"	
	#Notify before download
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate\AU" "AUOptions" "2" "DWord"
	#Enable Automatic Updates
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate\AU" "NoAutoUpdate" "0" "DWord"
	#Install Patches on Tuesday
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate\AU" "ScheduledInstallDay" "3" "DWord"
	#Install on the 2nd Tuesday
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate\AU" "ScheduledInstallSecondWeek" "1" "DWord"
	#Install on the 3rd Tuesday
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate\AU" "ScheduledInstallThirdWeek" "1" "DWord"
	#Install at 1am
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate\AU" "ScheduledInstallTime" "1" "DWord"
	#Prevents the upgrade to Windows 10
	Set-Reg "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate" "OSUpgrade" "0" "DWord"
	#Disable domain joined computers automatically and silently get registered as devices with Azure Active Directory
	# https://getadmx.com/?Category=Windows_10_2016&Policy=Microsoft.Policies.WorkplaceJoin::WJ_AutoJoin
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\WorkplaceJoin" "autoWorkplaceJoin" "0" "DWord"
	#Set Default to Microsoft Update instead of Windows Update
	If (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Services\7971f918-a847-4430-9279-4a52d1efe18d") {
		$regkeypath= "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Services"
		$value = "DefaultService"
		$test = (Get-ItemProperty $regkeypath).$value -eq "7971f918-a847-4430-9279-4a52d1efe18d" 
		If ($test -eq $False) {
			Set-Reg "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Services" "DefaultService" "7971f918-a847-4430-9279-4a52d1efe18d" "String"
		}
	}
	#The Windows Remote Management (WinRM) client must not use Basic authentication.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\WinRM\Client" "AllowBasic" "0" "DWord"
	#The Windows Remote Management (WinRM) client must not use Digest authentication.
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\WinRM\Client" "AllowDigest" "0" "DWord"
	#The Windows Remote Management (WinRM) client must not allow unencrypted traffic
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\WinRM\Client" "AllowUnencryptedTraffic" "0" "DWord"
	#The Windows Remote Management (WinRM) Service must not use Basic authentication
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\WinRM\Service" "AllowBasic" "0" "DWord"
	#The Windows Remote Management (WinRM) Service must not allow unencrypted traffic	
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\WinRM\Service" "AllowUnencryptedTraffic" "0" "DWord"
	#The Windows Remote Management (WinRM) service must not store RunAs credentials
	Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\WinRM\Service" "DisableRunAs" "1" "DWord"
	#Disable Crash Dumps
	Set-Reg "HKLM:\System\CurrentControlSet\Control\CrashControl" "CrashDumpEnabled" 0 "DWord"	
	#Disable Log about Dumps
	Set-Reg "HKLM:\System\CurrentControlSet\Control\CrashControl" "LogEvent" 0 "DWord"	
	#Disable Sending Alerts about Dumps
	Set-Reg "HKLM:\System\CurrentControlSet\Control\CrashControl" "SendAlert" 0 "DWord"	
	#Enable Automatic Crash Reboot
	Set-Reg "HKLM:\System\CurrentControlSet\Control\CrashControl" "AutoReboot" 1 "DWord"
	#NTFS Disable Last Access Update File Time Stamp 	
	Set-Reg "HKLM:\System\CurrentControlSet\Control\FileSystem" "NtfsDisableLastAccessUpdate" 1 "DWord"	
	#WDigest Authentication must be disabled
	Set-Reg "HKLM:\System\CurrentControlSet\Control\SecurityProviders\WDigest" "UseLogonCredential" "0" "DWord"
	#Structured Exception Handling Overwrite Protection (SEHOP) must be turned on
	Set-Reg "HKLM:\System\CurrentControlSet\Control\Session Manager\kernel" "DisableExceptionChainValidation" "0" "DWord"
	#Write Error to Log but do not display System Hard Error Message Dialog Boxes
	Set-Reg "HKLM:\System\CurrentControlSet\Control\Windows" "ErrorMode" 2 "DWord"
	#Disable Server SMB1 Protocol
	Set-Reg "HKLM:\System\CurrentControlSet\Services\LanmanServer\Parameters" "SMB1" "0" "DWord"
	Set-Reg "HKLM:\System\CurrentControlSet\Services\MrxSmb10" "Start" "4" "DWord"
	#The system will be configured to ignore NetBIOS name release requests except from WINS servers
	Set-Reg "HKLM:\System\CurrentControlSet\Services\Netbt\Parameters" "NoNameReleaseOnDemand" "1" "DWord"
	#The system must be configured to prevent IP source routing.
	Set-Reg "HKLM:\System\CurrentControlSet\Services\Tcpip6\Parameters" "DisableIPSourceRouting" "2" "DWord"
	Set-Reg "HKLM:\System\CurrentControlSet\Services\Tcpip\Parameters" "DisableIPSourceRouting" "2" "DWord"
	#The system will be configured to prevent ICMP redirects from overriding OSPF generated routes.
	Set-Reg "HKLM:\System\CurrentControlSet\Services\Tcpip\Parameters" "EnableICMPRedirect" "0" "DWord"
	#Enable PS/2 Mouse 
		#https://superuser.com/questions/996001/do-ps2-keyboards-work-on-windows-10
		#https://www.dell.com/community/Laptops-General-Read-Only/Mouse-not-working-Device-Manager-shows-PS-2-Compatible-Mouse/td-p/5087851
	Set-Reg "HKLM:\SYSTEM\CurrentControlSet\Services\i8042prt" "Start" 1 "DWord"
	Set-Reg "HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4d36e96f-e325-11ce-bfc1-08002be10318}" "UpperFilters" "mouclass" "MultiString"
	#Disable Use SNMP Legacy mode: http://blog.rtwilson.com/how-to-fix-a-network-printer-suddenly-showing-as-offline-in-windows-vista/
	Set-Reg "HKLM:\System\CurrentControlSet\Control\Print" "SNMPLegacy" 1 "DWord"
	#EMV Force COM5
	Set-Reg "HKLM:\SYSTEM\CurrentControlSet\Services\IngenicoVCOM\InstallConfig\R0" "COM" 5 "DWord"
	Set-Reg "HKLM:\SYSTEM\CurrentControlSet\Services\IngenicoVCOM\InstallConfig" "ForceComEnabled" 1 "DWord"
	#Prevent Downloading Printer Info and Icon
	#https://www.reddit.com/r/Windows10/comments/d3q45d/printer_and_amplifier_suddenly_showing_as/
	Set-Reg "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Device Metadata" "PreventDeviceMetadataFromNetwork " 1 "DWord"
	#Block from Switching to MS Account
	Set-Reg 'HKLM:\SOFTWARE\Microsoft\PolicyManager\default\Settings\AllowYourAccount' 'value' 0 "DWORD"
	#Block Microsoft Account
	Set-Reg 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' 'NoConnectedUser' 1 "DWORD"
	#Hide not-critical notifications
	Set-Reg 'HKLM:\SOFTWARE\Microsoft\Windows Defender Security Center\Notifications' 'DisableEnhancedNotifications' 1 "DWORD"
	Set-Reg 'HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender Security Center\Notifications' 'DisableEnhancedNotifications' 1 "DWORD"
	#Hide all notifications
	Set-Reg 'HKLM:\SOFTWARE\Microsoft\Windows Defender Security Center\Notifications' 'DisableNotifications' 1 "DWORD"
	Set-Reg 'HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender Security Center\Notifications' 'DisableNotifications' 1 "DWORD"
	#Disable Use SNMP Legacy mode
	Set-Reg 'HKLM:\System\CurrentControlSet\Control\Print' 'SNMPLegacy' 1 "DWORD"


	$regkeypath= "HKLM:\Software\Policies\Microsoft\Windows NT\Terminal Services"
	$value = "fAllowFullControl"
	$test = ([string]::IsNullOrEmpty((Get-ItemProperty $regkeypath).$value))
	If ($test -eq $False) {
		Remove-ItemProperty -path $regkeypath -name $value
	} 
	$regkeypath= "HKLM:\Software\Policies\Microsoft\Windows NT\Terminal Services"
	$value = "fUseMailto"
	$test = ([string]::IsNullOrEmpty((Get-ItemProperty $regkeypath).$value))
	If ($test -eq $False) {
		Remove-ItemProperty -path $regkeypath -name $value
	} 
	$regkeypath= "HKLM:\Software\Policies\Microsoft\Windows NT\Terminal Services"
	$value = "MaxTicketExpiry"
	$test = ([string]::IsNullOrEmpty((Get-ItemProperty $regkeypath).$value))
	If ($test -eq $False) {
		Remove-ItemProperty -path $regkeypath -name $value
	} 
	$regkeypath= "HKLM:\Software\Policies\Microsoft\Windows NT\Terminal Services"
	$value = "MaxTicketExpiryUnits"
	$test = ([string]::IsNullOrEmpty((Get-ItemProperty $regkeypath).$value))
	If ($test -eq $False) {
		Remove-ItemProperty -path $regkeypath -name $value
	}
	$regkeypath= "HKLM:\Software\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging"
	$value = "EnableScriptBlockInvocationLogging"
	$test = ([string]::IsNullOrEmpty((Get-ItemProperty $regkeypath).$value))
	If ($test -eq $False) {
		Remove-ItemProperty -path $regkeypath -name $value
	}
}

#============================================================================
#endregion Main Local Machine Load Local GPO
#============================================================================
#============================================================================
#region Main Local Machine VMWare Horzion Settings
#============================================================================
If ([Environment]::Is64BitOperatingSystem) {
	Set-Reg ($VMWare_Horzion_Key.replace("\Software\","\Software\Wow6432Node\") + "\Security") "AllowCmdLineCredentials" "0" "DWord"
	Set-Reg ($VMWare_Horzion_Key.replace("\Software\","\Software\Wow6432Node\") + "\Security") "CertCheckMode" "2" "DWord"
	Set-Reg ($VMWare_Horzion_Key.replace("\Software\","\Software\Wow6432Node\") + "\Security") "LogInAsCurrentUser" "true" "String"
	Set-Reg ($VMWare_Horzion_Key.replace("\Software\","\Software\Wow6432Node\") + "\Security") "LogInAsCurrentUser_Display" "true" "String"
	Set-Reg ($VMWare_Horzion_Key.replace("\Software\","\Software\Wow6432Node\") + "\Security") "SSLCipherList" $VMware_Horizon_SSLCipherList "String"
	Set-Reg ($VMWare_Horzion_Key.replace("\Software\","\Software\Wow6432Node\") + "\Security") "EnableTicketSSLAuth" 3 "DWORD"
	Set-Reg ($VMWare_Horzion_Key.replace("\Software\","\Software\Wow6432Node\")) "AutoUpdateAllowed" "false" "String"
	Set-Reg ($VMWare_Horzion_Key.replace("\Software\","\Software\Wow6432Node\")) "AllowDataSharing" "false" "String"
	Set-Reg ($VMWare_Horzion_Key.replace("\Software\","\Software\Wow6432Node\")) "IpProtocolUsage" "IPv4" "String"
	Set-Reg ($VMWare_Horzion_Key.replace("\Software\","\Software\Wow6432Node\")) "DomainName" $VMware_Horizon_NetBIOSDomain "String"
	Set-Reg ($VMWare_Horzion_Key.replace("\Software\","\Software\Wow6432Node\")) "ServerURL" $VMware_Horizon_Server "String"
	
} else {
	Set-Reg ($VMWare_Horzion_Key + "\Security") "AllowCmdLineCredentials" "0" "DWord"
	Set-Reg ($VMWare_Horzion_Key + "\Security") "CertCheckMode" "2" "DWord"
	Set-Reg ($VMWare_Horzion_Key + "\Security") "LogInAsCurrentUser" "true" "String"
	Set-Reg ($VMWare_Horzion_Key + "\Security") "LogInAsCurrentUser_Display" "true" "String"
	Set-Reg ($VMWare_Horzion_Key + "\Security") "SSLCipherList" $VMware_Horizon_SSLCipherList "String"
	Set-Reg ($VMWare_Horzion_Key + "\Security") "EnableTicketSSLAuth" 3 "DWORD"
	Set-Reg ($VMWare_Horzion_Key) "AutoUpdateAllowed" "false" "String"
	Set-Reg ($VMWare_Horzion_Key) "AllowDataSharing" "false" "String"
	Set-Reg ($VMWare_Horzion_Key) "IpProtocolUsage" "IPv4" "String"
	Set-Reg ($VMWare_Horzion_Key) "DomainName" $VMware_Horizon_NetBIOSDomain "String"
	Set-Reg ($VMWare_Horzion_Key) "ServerURL" $VMware_Horizon_Server "String"
}

#============================================================================
#endregion Main Local Machine VMWare Horzion Settings
#============================================================================
#============================================================================
#regionMain Local Machine Temp Cleanup
#============================================================================
If (get-ScheduledJob | Where-Object {$_.Name -eq "Clean-Temp-Folders"}) {
	get-ScheduledJob | Where-Object {$_.Name -eq "Clean-Temp-Folders"} | Unregister-ScheduledJob
}

$SchTrigger = New-JobTrigger -AtStartup
$SchJobOptions = New-ScheduledJobOption -RunElevated
Register-ScheduledJob -Name "Clean-Temp-Folders" -Trigger $SchTrigger -ScheduledJobOption $SchJobOptions  -ScriptBlock {
	If (Test-Path($env:systemdrive + "\temp")) {
		Get-ChildItem -Path ($env:systemdrive + "\temp") | Remove-Item -Recurse -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
	}
	If (Test-Path($env:systemroot+ "\temp")) {
		Get-ChildItem -Path ($env:systemroot + "\temp") | Remove-Item -Recurse -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
	}
	$UsersProfileFolders = @((Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\' -Name "ProfilesDirectory").ProfilesDirectory,($env:systemdrive + "\Users"))
	ForEach ($UserFolders in ($UsersProfileFolders | Select-Object -Unique)) {
		ForEach ($Folder in (Get-ChildItem -Directory -Path $UserFolders).fullname) {
			write-host ($Folder + "\AppData\Local\Temp")
			If (Test-Path($Folder + "\AppData\Local\Temp")) {
				Get-ChildItem -Path ($Folder + "\AppData\Local\Temp") | Remove-Item -Recurse -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
			}
		}
	}
}
#============================================================================
#endregionMain Local Machine Temp Cleanup
#============================================================================
#============================================================================
#regionMain Local Machine Cleanup
#============================================================================
#Recording Version of script
write-host ("Recording " + $ScriptVersionValue + ": " + $ScriptVersion + " in " + $ScriptVersionKey + " Key.") -foregroundcolor "Green"
Set-Reg ("HKLM:\Software\" + $ScriptVersionKey) $ScriptVersionValue  $ScriptVersion "String"
Set-Reg ("HKLM:\Software\" + $ScriptVersionKey) $ScriptDateValue  (Get-Date -format yyyyMMdd) "String"
write-host
#cleanup mapped drives
If (Test-Path "PSRemote:\") {
	Remove-PSDrive -Name "PSRemote"
}

$sw.Stop()
Write-Host ("Script took: " + (FormatElapsedTime($sw.Elapsed)) + " to run.")
#============================================================================
#endregion Main Local Machine Cleanup
#============================================================================
If (-Not [string]::IsNullOrEmpty($LogFile)) {
	Stop-Transcript
}
#############################################################################
#endregion Main
#############################################################################
