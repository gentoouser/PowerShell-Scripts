<# 
.SYNOPSIS
    Name: Windows_Hardending.ps1
    Hardens Fresh installs of Windows

.DESCRIPTION
	* Hardens c:\
	* Caches configureation files
	* Creates Store Users
	* Lock down users by loading registry and appling settings
	* Lock down users by appling GPO
	
.PARAMETER Config
	XML Configuration file with all the hardening settings.
.PARAMETER Profiles
	Array of users to lockdown. If Store is enabled all store Window users will be added to the array and locked down.	
.PARAMETER LICache
	Location of the local configureation files cache.	
.PARAMETER RemoteFiles
	Network path of configureation files are to copy down. 	
.PARAMETER User
	Username for network configureation share.
.PARAMETER Password
	Password that goes with useranme for network configureation share.
.PARAMETER ActiveUser
	User that will Actively logon to the computer every day. 
.PARAMETER Managers
	Enable relaxation of policies for Managers.
.PARAMETER Store
	Enables more locked down of users and creates store Local Windows accounts.
.PARAMETER LockedDown
	Lock down user accounts more
.PARAMETER UserOnly
	Sets user settings only and no machine settings
.PARAMETER NoCacheUpdate
	Skip updating local cache.
.PARAMETER AllowClientTLS1
	Enables Computer to go to TLS 1.0 sites.
.PARAMETER NoOEMInfo
	Keeps from reseting the OEM Info.
.PARAMETER OEMInfoAddSerial
	Added Serial number to the System Preferences.
.PARAMETER NoBgInfo
	Does not setup BGInfo to launch on logon.
.PARAMETER Wifi
	Enables services needed for WiFi
.PARAMETER IPv6
	Keeps IPv6 enabled; otherwise IPv6 will be disabled. 
.EXAMPLE
   & .\Windows_Hardending.ps1 -AllowClientTLS1
.EXAMPLE
	powershell -executionpolicy unrestricted -file .\Windows_Hardending.ps1 -Config .\Windows_Hardending.ps1.config -Store -AllowClientTLS1 -Profile Default,User 
.NOTES
 Author: Paul Fuller
 Changes:
 	* Version 3.00.00 - Switch to XML Config
	* Version 3.00.01 - Fixed Cipher issue where powershell could not handle "/"
	* Version 3.00.02 - Fixed VM Detection
	* Version 3.00.03 - Fixed Setting A Binary Registry key.
	* Version 3.00.04 - Use "Get-CimInstance" when avalible if not default to "Get-WmiObject". 
			    Also use Regex to detect Manager and Store PC's. 
			    Added AddressFilter to the Firewall to better control remote connections. 
			    Moved contol of ScheduledJob to xml; also if ScheduledJob is not avalible use ScheduledTask.
			    Fixed Issue with locking down Default user.
			    Using SID to find local profile.
			    Disable WiFi by default; use -Wifi to enable Wifi.
			    Updated RemoveFCTID Shortcut.
	* Version 3.00.05 - Fixed Bug with AllowClientTLS1 switch. Updated Get-MachineType.  	
			    Added more debugging to Deny files.
			    Added Get-envValueFromString Function to handle PowerShell $env: variables from XML.
	* Version 3.00.06 - Cleaned up messages when Deny file does not exist.
			    Cleaned up usage of -UserOnly
			    Create new if user exists and the profiles does not.
			    Apply certain keys only to Default User.
			    Fixed Password generation bug.
			    Fixed bug where account was not enabled when trying to recreate user profile.
	* Version 3.00.07 - Fixed "Import-StartLayout: Access to the path"
			    Fixed Active User bug.
	* Version 3.00.08 -  Reset Startmenu layout on Windows 10 1809+.
	* Version 3.00.09 - Hide errors for Start Menu 1809 reset.
	* Version 3.00.10 - Updated name of log file. Added function to install fonts. Import Security Template INI
	* Version 3.00.11 - Updated info for Taskmenu layout. Clear old StartMenu layout.
	#>
#Requires -Version 5.1 -PSEdition Desktop
#############################################################################
#region Parameter Config
#############################################################################
PARAM (
	[string]$Config 			= $null,
    [array]$Profiles  	  		= @("Default"),	
	[string]$LICache	  		= $null,
	[string]$RemoteFiles  		= (Split-Path -Parent -Path $MyInvocation.MyCommand.Definition),
	[String]$User		    	= $null,
	[String]$			= $null,
	[String]$ActiveUser	    	= $null,
	[string]$StartLayoutXML		= $null,
	[String]$BackgroundFolder 	= $null,
	[switch]$Manager	    	= $false,
	[switch]$Store	  	  		= $false,
	[switch]$LockedDown	  		= $false,
	[switch]$UserOnly			= $false,
	[switch]$NoCacheUpdate		= $false,
	[switch]$AllowClientTLS1	= $false,
	[switch]$Wifi				= $false,
	[switch]$NoOEMInfo			= $false,
	[switch]$OEMInfoAddSerial	= $false,
	[switch]$NoBgInfo			= $false,
	[switch]$IPv6				= $false
)
#############################################################################
#endregion Parameter Config
#############################################################################
#region Force Running Script as Admin
If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
	{   
		$arguments = "& '" + $myinvocation.mycommand.definition + "'"
		Start-Process powershell -Verb runAs -ArgumentList $arguments
		Break
	}
#endregion Force Running Script as Admin
#############################################################################
#region User Variables
#############################################################################
$ScriptVersion = "3.0.11"
$LogFile = ("\Logs\" + `
		   ($MyInvocation.MyCommand.Name -replace ".ps1","") + "_" + `
		   $env:computername + "_" + `
		   (Get-Date -format yyyyMMdd-hhmm) + ".log")
$sw = [Diagnostics.Stopwatch]::StartNew()
$IsVM = $False
$HKEY = "HKU\DEFAULTUSER"
$UsersProfileFolder = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\' -Name "ProfilesDirectory").ProfilesDirectory
$ProfileList =  New-Object System.Collections.ArrayList
$WScriptShell = New-Object -ComObject ("WScript.Shell")
$RegAddSCHANNEL = 'HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL'
#############################################################################
#endregion User Variables
#############################################################################

#############################################################################
#region Setup Sessions
#############################################################################
#region Import Modules
Add-Type -AssemblyName System.web
#endregion Import Modules

#region Import Config
If (-Not $Config) {
	If ( Test-Path ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\" + $MyInvocation.MyCommand.Name + ".config")) {
		$Config = ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\" + $MyInvocation.MyCommand.Name + ".config")
	}
}
If (Test-Path $Config) {
    # Import email settings from config file
    [xml]$ConfigFile = Get-Content $Config
	#Test Import	
    If (-Not $ConfigFile) {
        Write-Error -Message "Config Corrupted . . . Exiting"
        break
    }
} else {
	Write-Error -Message "No Config File Found . . . Exiting"
	break
}
#endregion Import Config
#region Update Local Cache variable
If (-Not $LICache) {
	If ($ConfigFile.Config.Company.LocalCache) {
		$LICache = $ConfigFile.Config.Company.LocalCache
		#Remove any trailing \
		If ($LICache.Substring($LICache.Length - 1) -eq "\") {
			$LICache = $LICache.Substring(0,$LICache.Length - 1)
		}

		Write-Host ("Local Cache: " + $LICache)
	} else {
		Write-Error -Message "No Local Cache Defined . . . Exiting"
		break		
	}	
} else {
	Write-Host ("Local Cache: " + $LICache)
}
#endregion Update Local Cache variable
#region Start logging.
If (-Not [string]::IsNullOrEmpty($LICache + $LogFile)) {
	If (-Not( Test-Path (Split-Path -Path ($LICache + $LogFile) -Parent))) {
		New-Item -ItemType directory -Path (Split-Path -Path ($LICache + $LogFile) -Parent) | Out-Null
		$Acl = Get-Acl (Split-Path -Path ($LICache + $LogFile) -Parent)
		$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule('Users', "FullControl", "ContainerInherit, ObjectInherit", "None", "Allow")
		$Acl.Setaccessrule($Ar)
		Set-Acl (Split-Path -Path ($LICache + $LogFile) -Parent) $Acl
	}
	try { 
		Start-Transcript -Path ($LICache + $LogFile) -Append
	} catch { 
		Stop-transcript
		Start-Transcript -Path ($LICache + $LogFile) -Append
	} 
	Write-Host ("Script Name   : " + $MyInvocation.MyCommand.Name)
	Write-Host ("Script Version: " + $ScriptVersion)
	Write-Host ("XML Version   : " + $ConfigFile.Config.Company.Version)
	Write-Host (" ")
}	
#endregion Start logging.
#region Store Setup
If ($Store) {
	$LockedDown = $True
}
#endregion Store Setup
#region Load Registry Hives
New-PSDrive -PSProvider Registry -Name HKU -Root HKEY_USERS -erroraction 'silentlycontinue' | Out-Null
New-PSDrive -PSProvider Registry -Name HKCR -Root HKEY_CLASSES_ROOT -erroraction 'silentlycontinue' | Out-Null
#endregion Load Registry Hives
#region Share Credential Setup
if ( $User -and $Password) {
	$Credential = New-Object System.Management.Automation.PSCredential ($User, (ConvertTo-SecureString $Password -AsPlainText -Force))
}
#endregion Share Credential Setup
#region ProfileList Setup
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
# Find Active User Using IP
If (-Not $ActiveUser) {
	If (Get-Command Get-CimInstance -errorAction SilentlyContinue) {
		$ActiveIP = (Get-CimInstance -Class Win32_NetworkAdapterConfiguration | Where-Object {$_.DefaultIPGateway -ne $null}).IPAddress | select-object -first 1
	} Else {
		$ActiveIP = (Get-WmiObject -Class Win32_NetworkAdapterConfiguration | Where-Object {$_.DefaultIPGateway -ne $null}).IPAddress | select-object -first 1
	} 
	$ActiveAN = ([int]$ActiveIP.substring($ActiveIP.Length -2) - 10)
	If ( $ActiveAN -notin ([int]$ConfigFile.Config.Company.UserRangeStart)..([int]$ConfigFile.Config.Company.UserRangeEnd) ) {
		$ActiveAN = ([int]$ConfigFile.Config.Company.UserRangeStart)
	}
	If ($Store) {
		$ActiveUser = ($ConfigFile.Config.Company.UserBaseName + $ActiveAN)
		Write-Host ('Active User: ' + $ActiveUser)
	}
}
#Find if Managers Computer
If ($Manager -eq $false -and $env:computername -match ($configfile.Config.Company.ManagerComputernameRegEx)) {
	$Manager = $true
}
#Find if Store Computer
If ($Store -eq $false -and $env:computername -match ($configfile.Config.Company.StoreComputernameRegEx)) {
	$Store = $true
}
#endregion ProfileList Setup

#############################################################################
#endregion Setup Sessions
#############################################################################

#############################################################################
#region Functions
#############################################################################
function FormatElapsedTime($ts) {
    #https://stackoverflow.com/questions/3513650/timing-a-commands-execution-in-powershell
	$elapsedTime = ""
    if ( $ts.Hours -gt 0 ) {
        $elapsedTime = [string]::Format( "{0:00} hours {1:00} min. {2:00}.{3:00} sec.", $ts.Hours, $ts.Minutes, $ts.Seconds, $ts.Milliseconds / 10 );
    } else {
        if ( $ts.Minutes -gt 0 ) {
            $elapsedTime = [string]::Format( "{0:00} min. {1:00}.{2:00} sec.", $ts.Minutes, $ts.Seconds, $ts.Milliseconds / 10 );
        } else {
            $elapsedTime = [string]::Format( "{0:00}.{1:00} sec.", $ts.Seconds, $ts.Milliseconds / 10 );
        }
        if ($ts.Hours -eq 0 -and $ts.Minutes -eq 0 -and $ts.Seconds -eq 0) {
            $elapsedTime = [string]::Format("{0:00} ms.", $ts.Milliseconds);
        }
        if ($ts.Milliseconds -eq 0) {
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
function Get-envValueFromString {
    <#
    .SYNOPSIS
    Enumerates string with $ENV: into  Enviromental variable value. 
    .DESCRIPTION
    Enumerates string with $ENV: into  Enviromental variable value. 
    .EXAMPLE
    $teststr="`$env:programdata\Microsoft\Windows\Start Menu\Programs\Administrative Tools"
    Get-ENVFromString -Path $teststr

    C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Administrative Tools
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0,HelpMessage ="String input required.")]
        [string]
        # String to Enumerate
        $Path
	)
    [Array]$PathArray = $null
    #cleanup brackets in input
    If ($Path -match "{" -or $Path -match "}") {
        $Path = $Path -replace "{","" -replace "}",""
    }
    If ($Path -match '\$env:') {
        #Loop though and test for var
        Foreach ($Folder in ($Path.Split("\"))) {
            If ($Folder -match '\$env:') {
                #Get value of matchin envoriment varible
                $PathArray += (Get-ChildItem Env: | Where-Object{ $_.Name -eq ($Folder.Replace("`$env:",""))}).value
            } else {
                $PathArray += $Folder
            }
        }
        return ( $PathArray -join "\")
    } else {
        return $Path
    }
}
function Write-Color {
    <#
 .SYNOPSIS
        Write-Color is a wrapper around Write-Host.
 
        It provides:
        - Easy manipulation of colors,
        - Logging output to file (log)
        - Nice formatting options out of the box.
 
 .DESCRIPTION
        Author: przemyslaw.klys at evotec.pl
        Project website: https://evotec.xyz/hub/scripts/write-color-ps1/
        Project support: https://github.com/EvotecIT/PSWriteColor
 
        Original idea: Josh (https://stackoverflow.com/users/81769/josh)
 
 .EXAMPLE
    Write-Color -Text "Red ", "Green ", "Yellow " -Color Red,Green,Yellow
 
    .EXAMPLE
 Write-Color -Text "This is text in Green ",
     "followed by red ",
     "and then we have Magenta... ",
     "isn't it fun? ",
     "Here goes DarkCyan" -Color Green,Red,Magenta,White,DarkCyan
 
    .EXAMPLE
 Write-Color -Text "This is text in Green ",
     "followed by red ",
     "and then we have Magenta... ",
     "isn't it fun? ",
                    "Here goes DarkCyan" -Color Green,Red,Magenta,White,DarkCyan -StartTab 3 -LinesBefore 1 -LinesAfter 1
 
    .EXAMPLE
 Write-Color "1. ", "Option 1" -Color Yellow, Green
 Write-Color "2. ", "Option 2" -Color Yellow, Green
 Write-Color "3. ", "Option 3" -Color Yellow, Green
 Write-Color "4. ", "Option 4" -Color Yellow, Green
 Write-Color "9. ", "Press 9 to exit" -Color Yellow, Gray -LinesBefore 1
 
    .EXAMPLE
 Write-Color -LinesBefore 2 -Text "This little ","message is ", "written to log ", "file as well." `
    -Color Yellow, White, Green, Red, Red -LogFile "C:\testing.txt" -TimeFormat "yyyy-MM-dd HH:mm:ss"
 Write-Color -Text "This can get ","handy if ", "want to display things, and log actions to file ", "at the same time." `
    -Color Yellow, White, Green, Red, Red -LogFile "C:\testing.txt"
 
    .EXAMPLE
    # Added in 0.5
    Write-Color -T "My text", " is ", "all colorful" -C Yellow, Red, Green -B Green, Green, Yellow
    wc -t "my text" -c yellow -b green
    wc -text "my text" -c red
 
    .NOTES
        CHANGELOG
 
        Version 0.5 (25th April 2018)
        -----------
        - Added backgroundcolor
        - Added aliases T/B/C to shorter code
        - Added alias to function (can be used with "WC")
        - Fixes to module publishing
 
        Version 0.4.0-0.4.9 (25th April 2018)
        -------------------
        - Published as module
        - Fixed small issues
 
        Version 0.31 (20th April 2018)
        ------------
        - Added Try/Catch for Write-Output (might need some additional work)
        - Small change to parameters
 
        Version 0.3 (9th April 2018)
        -----------
        - Added -ShowTime
        - Added -NoNewLine
        - Added function description
        - Changed some formatting
 
        Version 0.2
        -----------
        - Added logging to file
 
        Version 0.1
        -----------
        - First draft
 
        Additional Notes:
        - TimeFormat https://msdn.microsoft.com/en-us/library/8kb3ddd4.aspx
    #>
    [alias('Write-Colour')]
    [CmdletBinding()]
    param ([alias ('T')] [String[]]$Text,
        [alias ('C', 'ForegroundColor', 'FGC')] [ConsoleColor[]]$Color = [ConsoleColor]::White,
        [alias ('B', 'BGC')] [ConsoleColor[]]$BackGroundColor = $null,
        [alias ('Indent')][int] $StartTab = 0,
        [int] $LinesBefore = 0,
        [int] $LinesAfter = 0,
        [int] $StartSpaces = 0,
        [alias ('L')] [string] $LogFile = '',
        [Alias('DateFormat', 'TimeFormat')][string] $DateTimeFormat = 'yyyy-MM-dd HH:mm:ss',
        [alias ('LogTimeStamp')][bool] $LogTime = $true,
        [ValidateSet('unknown', 'string', 'unicode', 'bigendianunicode', 'utf8', 'utf7', 'utf32', 'ascii', 'default', 'oem')][string]$Encoding = 'Unicode',
        [switch] $ShowTime,
        [switch] $NoNewLine)
    $DefaultColor = $Color[0]
    if ($null -ne $BackGroundColor -and $BackGroundColor.Count -ne $Color.Count) { Write-Error "Colors, BackGroundColors parameters count doesn't match. Terminated."; return }
    if ($LinesBefore -ne 0) { for ($i = 0; $i -lt $LinesBefore; $i++) { Write-Host -Object "`n" -NoNewline } }
    if ($StartTab -ne 0) { for ($i = 0; $i -lt $StartTab; $i++) { Write-Host -Object "`t" -NoNewLine } }
    if ($StartSpaces -ne 0) { for ($i = 0; $i -lt $StartSpaces; $i++) { Write-Host -Object ' ' -NoNewLine } }
    if ($ShowTime) { Write-Host -Object "[$([datetime]::Now.ToString($DateTimeFormat))]" -NoNewline }
    if ($Text.Count -ne 0) {
        if ($Color.Count -ge $Text.Count) { if ($null -eq $BackGroundColor) { for ($i = 0; $i -lt $Text.Length; $i++) { Write-Host -Object $Text[$i] -ForegroundColor $Color[$i] -NoNewLine } } else { for ($i = 0; $i -lt $Text.Length; $i++) { Write-Host -Object $Text[$i] -ForegroundColor $Color[$i] -BackgroundColor $BackGroundColor[$i] -NoNewLine } } } else {
            if ($null -eq $BackGroundColor) {
                for ($i = 0; $i -lt $Color.Length; $i++) { Write-Host -Object $Text[$i] -ForegroundColor $Color[$i] -NoNewLine }
                for ($i = $Color.Length; $i -lt $Text.Length; $i++) { Write-Host -Object $Text[$i] -ForegroundColor $DefaultColor -NoNewLine }
            } else {
                for ($i = 0; $i -lt $Color.Length; $i++) { Write-Host -Object $Text[$i] -ForegroundColor $Color[$i] -BackgroundColor $BackGroundColor[$i] -NoNewLine }
                for ($i = $Color.Length; $i -lt $Text.Length; $i++) { Write-Host -Object $Text[$i] -ForegroundColor $DefaultColor -BackgroundColor $BackGroundColor[0] -NoNewLine }
            }
        }
    }
    if ($NoNewLine -eq $true) { Write-Host -NoNewline } else { Write-Host }
    if ($LinesAfter -ne 0) { for ($i = 0; $i -lt $LinesAfter; $i++) { Write-Host -Object "`n" -NoNewline } }
    if ($Text.Count -ne 0 -and $LogFile -ne "") {
        $TextToFile = ""
        for ($i = 0; $i -lt $Text.Length; $i++) { $TextToFile += $Text[$i] }
        try { if ($LogTime) { Write-Output -InputObject "[$([datetime]::Now.ToString($DateTimeFormat))]$TextToFile" | Out-File -FilePath $LogFile -Encoding $Encoding -Append } else { Write-Output -InputObject "$TextToFile" | Out-File -FilePath $LogFile -Encoding $Encoding -Append } } catch { $_.Exception }
    }
}
Function Install-Font {
    <#
        .Synopsis
        Installs one or more fonts.
        .Parameter FontPath
        The path to the font to be installed or a directory containing fonts to install.
        .Parameter Recurse
        Searches for fonts to install recursively when a path to a directory is provided.
        .Notes
        There's no checking if a given font is already installed. This is problematic as an existing
        installation will trigger a GUI dialogue requesting confirmation to overwrite the installed
		font, breaking unattended and CLI-only scenarios.
		.Source
		 https://www.powershellgallery.com/packages/PSWinGlue/0.3.3/Content/Functions%5CInstall-Font.ps1
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [String]$FontPath,

        [Switch]$Recurse
    )
	function Get-FontName
	{
	<#
	.SYNOPSIS
		Retrieves the font name from a TTF file.

	.DESCRIPTION
		This cmdlet does not install a font. It retrieves the font name from the given file, or the provided path

	.PARAMETER Path
		Specifies the path to the font files


	.PARAMETER Item
		Specifies the file item object. This is provided by Get-Item or Get-ChildItem.

	.EXAMPLE 
		Get-FontName -Path $myfontPath

	.EXAMPLE 
		Get-ChildItem -Path *.ttf | Get-FontName

	.NOTES
		Micky Balladelli
		Source: https://github.com/MickyBalladelli/Get-FontName/blob/master/Get-FontName.ps1
	#>
		[CmdletBinding()]
		PARAM(
			[Parameter(
				ParameterSetName='Path'
			)]
			[String]$Path,
	 
			[Parameter(
				ValueFromPipeline = $true,
				ParameterSetName='Item'
			)]
			[object[]]$Item

		)

		BEGIN
		{
			Add-Type -AssemblyName System.Drawing
			$ttfFiles = @()
			$fontCollection = new-object System.Drawing.Text.PrivateFontCollection
		}
		
		PROCESS
		{
			if ($Path -ne "")
			{
				$ttfFiles = Get-ChildItem $path
			}
			else
			{
				$ttfFiles += $Item
			}

		}

		END
		{
			$ttfFiles | ForEach-Object {
				$fontCollection.AddFontFile($_.fullname)
				$fontCollection.Families[-1].Name
			}
		}
	}
    $ErrorActionPreference = 'Stop'
    $ShellAppFontNamespace = 0x14
	#list Fonts
	[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
	$objFonts = New-Object System.Drawing.Text.InstalledFontCollection
	$colFonts = $objFonts.Families

    if (Test-Path -Path $FontPath) {
        $FontItem = Get-Item -Path $FontPath
        if ($FontItem -is [IO.DirectoryInfo]) {
            if ($Recurse) {
                $Fonts = Get-ChildItem -Path $FontItem -Include ('*.fon','*.otf','*.ttc','*.ttf') -Recurse
            } else {
                $Fonts = Get-ChildItem -Path "$FontItem\*" -Include ('*.fon','*.otf','*.ttc','*.ttf')
            }

            if (!$Fonts) {
                throw ('Unable to locate any fonts in provided directory: {0}' -f $FontItem.FullName)
            }
        } elseif ($FontItem -is [IO.FileInfo]) {
            if ($FontItem.Extension -notin ('.fon','.otf','.ttc','.ttf')) {
                throw ('Provided file does not appear to be a valid font: {0}' -f $FontItem.FullName)
            }

            $Fonts = $FontItem
        } else {
            throw ('Expected directory or file but received: {0}' -f $FontItem.GetType().Name)
        }
    } else {
        throw ('Provided font path does not appear to be valid: {0}' -f $FontPath)
    }

    $ShellApp = New-Object -ComObject Shell.Application
    $FontsFolder = $ShellApp.NameSpace($ShellAppFontNamespace)
    foreach ($Font in $Fonts) {
        If ((Get-FontName -Path $Font.FullName) -notin $colFonts){
            Write-Verbose -Message ('Installing font: {0}' -f $Font.BaseName)
            #Write-Host ('Installing font: {0}' -f (Get-FontName -Path $Font.FullName))
            $FontsFolder.CopyHere($Font.FullName)
        }else{
            Write-Verbose -Message ('Skipping font: {0}' -f (Get-FontName -Path $Font.FullName))
            #Write-Host ('Skipping font: {0}' -f (Get-FontName -Path $Font.FullName))
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
#region VM Test
#Get Where we are running
If ((Get-MachineType).type -eq "VM") {
	$IsVM = $True
	Write-Host ("Running in on Virtual Hardware")
}else{
	Write-Host ("Running in on Physical Hardware")
}
#endregion VM Test
#region Local Cache Update
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
		write-host ("Copying to Local Install cache: " + $LICache + " Please Wait.")
		$CurrentScriptUTC = $(Get-Item $MyInvocation.MyCommand.Definition).LastWriteTimeUtc		
		#Copy-Item  "PSRemote:\*" -Destination $LICache -Recurse -Force
		#Copy-Newer -Source "PSRemote:\" -Destination $LICache -Exclude @("logs") -Overwrite
		$UNCTemp = (Get-Item "PSRemote:\").FullName
		If ($UNCTemp.Substring($UNCTemp.Length - 1) -eq "\") {
			$UNCTemp = $UNCTemp.Substring(0,$UNCTemp.Length - 1)
		}
		$temp = @("/E",('"' + $UNCTemp + '"') ,('"' + $LICache + '"'))
		$temp += $ConfigFile.Config.RoboCopyOptions.Option
		$process = Start-Process -FilePath ("robocopy.exe") -ArgumentList $temp -PassThru -NoNewWindow -wait

		If (Test-Path ($LICache + "\" + $MyInvocation.MyCommand.Name)) {
			If ($(Get-Item ($LICache + "\" + $MyInvocation.MyCommand.Name)).LastWriteTimeUtc -gt $CurrentScriptUTC) {
				#write-host ("Starting newer copy of script...")
				#Stop-transcript
				#Need to fix getting the correct parameters sent to the new script instance
				#$&$MyInvocation.MyCommand.Definition  $MyInvocation.MyCommand.Parameters 
				#exit
			}
		}
	}
}
#endregion Local Cache Update
#region Harden Permission on the c:\
# Remove user the rights to create and modify data on the root of the c:\ drive.
If (-Not $UserOnly) {
	write-host ("Hardening Permissions: " + ($env:systemdrive + "\"))
	$acl = Get-Acl ($env:systemdrive + "\")
	If ($acl.Access | Where-Object { $_.IdentityReference -eq "NT AUTHORITY\Authenticated Users" }) {
		$usersid = New-Object System.Security.Principal.Ntaccount ("NT AUTHORITY\Authenticated Users")
		$acl.PurgeAccessRules($usersid) | Out-Null
		$acl | Set-Acl ($env:systemdrive + "\") | Out-Null
	}
	If (Test-Path $ConfigFile.Config.Company.SoftwarePath) {
		write-host ("Setting Permissions: " + $ConfigFile.Config.Company.SoftwarePath)
		$acl = Get-Acl $ConfigFile.Config.Company.SoftwarePath
		If (-Not ($acl.Access | Where-Object { $_.IdentityReference -eq "BUILTIN\Users" -and $_.FileSystemRights -eq "FullControl"})) {
			$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule('Users', "FullControl", "ContainerInherit, ObjectInherit", "None", "Allow")
			$Acl.Setaccessrule($Ar) | Out-Null
			Set-Acl $ConfigFile.Config.Company.SoftwarePath $Acl | Out-Null
		}
	}
}
#endregion Harden Permission on the c:\
#region Create Local Store users
If ($Store) {
	#Testing if we need to create any accounts
	Write-Host ('Testing for existing ' + $ConfigFile.Config.Company.UserBaseName + ' users.')
	$CreateUsers = $False
	$UserRange = ([int]$ConfigFile.Config.Company.UserRangeStart)..([int]$ConfigFile.Config.Company.UserRangeEnd)
	ForEach ( $i in $UserRange) {	
		If ($i) {
			If (-Not (Get-LocalUser -Name ($ConfigFile.Config.Company.UserBaseName + $i) -erroraction 'silentlycontinue')) {
				$CreateUsers = $True
			}else{
				If (Test-Path ($UsersProfileFolder + "\" + $ConfigFile.Config.Company.UserBaseName + $i)) {
					Write-Host ("`tAdding to Profile List: " + ($UsersProfileFolder + "\" + $ConfigFile.Config.Company.UserBaseName + $i))
					$ProfileList.Add(($ConfigFile.Config.Company.UserBaseName + $i).ToLower()) | Out-Null
					$HideAccounts += ($ConfigFile.Config.Company.UserBaseName + $i).ToLower()
				}else{
					$CreateUsers = $True
				}
			}
		}
	}
	#Disable Password Requirements for creating new accounts
	If ($CreateUsers) {
		ForEach ( $i in $UserRange) {	
			If ($i) {
				#Only create profile if user is a local user
				If (-Not (Get-LocalUser -Name ($ConfigFile.Config.Company.UserBaseName + $i) -erroraction SilentlyContinue)) {
					#Only create profile if no profile exists
					$CurrentUserSID = (Get-LocalUser -Name ($ConfigFile.Config.Company.UserBaseName + $i) -erroraction SilentlyContinue).SID
					If (Get-Command Get-CimInstance -errorAction SilentlyContinue) {
						$UserProfile = (Get-CimInstance Win32_UserProfile | Where-Object { $_.SID -eq $CurrentUserSID}).localpath
					} Else {
						$UserProfile = (Get-WmiObject Win32_UserProfile | Where-Object { $_.SID -eq $CurrentUserSID}).localpath
					} 
					If (-Not (Test-Path ($UserProfile + "\ntuser.dat"))) {
						write-host ("Creating User: " + ($ConfigFile.Config.Company.UserBaseName + $i))
						#Random 120 chr. password
						$TempPass= (ConvertTo-SecureString ([system.web.security.membership]::GeneratePassword(120,32)).tostring() -AsPlainText -Force)
						New-LocalUser -Name ($ConfigFile.Config.Company.UserBaseName + $i).ToLower() -Description "Store Window User" -FullName ($ConfigFile.Config.Company.UserBaseName + $i) -Password $TempPass -AccountNeverExpires -UserMayNotChangePassword -PasswordNeverExpires | Out-Null
						Add-LocalGroupMember -Name 'Administrators' -Member ($ConfigFile.Config.Company.UserBaseName + $i) | Out-Null
						Write-Host "`tWorking on Creating user profile: " ($ConfigFile.Config.Company.UserBaseName + $i)
						#launch process as user to create user profile
						# https://msdn.microsoft.com/en-us/library/system.diagnostics.processstartinfo(v=vs.110).aspx
						$processStartInfo = New-Object System.Diagnostics.ProcessStartInfo
						$processStartInfo.UserName = ($ConfigFile.Config.Company.UserBaseName + $i)
						$processStartInfo.Domain = "."
						$processStartInfo.Password = $TempPass
						$processStartInfo.FileName = "cmd"
						$processStartInfo.WorkingDirectory = $LICache
						$processStartInfo.Arguments = "/C echo . && echo %username% && echo ."
						$processStartInfo.LoadUserProfile = $true
						$processStartInfo.UseShellExecute = $false
						$processStartInfo.WindowStyle  = "minimized"
						$processStartInfo.RedirectStandardOutput = $false
						$process = [System.Diagnostics.Process]::Start($processStartInfo)
						$Process.WaitForExit()   
						#Add setup user to profiles created to allow registry to be created. 
						If (Test-Path ($UsersProfileFolder + "\Window" + $i) ) {
							$ProfileList.Add(($ConfigFile.Config.Company.UserBaseName + $i).ToLower()) | Out-Null
							$HideAccounts += ($ConfigFile.Config.Company.UserBaseName + $i).ToLower()
							#Grant Current user rights on new Profiles
							Write-Host ("`tUpdating ACLs and adding to Profile List: " + ($UsersProfileFolder + "\Window" + $i))
							$user_account=$env:username
							$Acl = Get-Acl ($UsersProfileFolder + "\Window" + $i)
							$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule($user_account, "FullControl", "ContainerInherit, ObjectInherit", "None", "Allow")
							$Acl.Setaccessrule($Ar)
							Set-Acl ($UsersProfileFolder + "\Window" + $i) $Acl
							#Disable User.
							Disable-LocalUser -Name ($ConfigFile.Config.Company.UserBaseName + $i).ToLower() -Confirm:$false
						}
					}
				} else {
					#Only create profile if no profile exists
					$CurrentUserSID = (Get-LocalUser -Name ($ConfigFile.Config.Company.UserBaseName + $i)).SID
					If (Get-Command Get-CimInstance -errorAction SilentlyContinue) {
						$UserProfile = (Get-CimInstance Win32_UserProfile | Where-Object { $_.SID -eq $CurrentUserSID}).localpath
					} Else {
						$UserProfile = (Get-WmiObject Win32_UserProfile | Where-Object { $_.SID -eq $CurrentUserSID}).localpath
					} 
					If (-Not (Test-Path ($UserProfile + "\ntuser.dat"))) {
						If ((Get-LocalUser -Name ($ConfigFile.Config.Company.UserBaseName + $i)).Enabled -eq $false) {
							Enable-LocalUser -Name ($ConfigFile.Config.Company.UserBaseName + $i).ToLower()
						}
						Write-Host "`tResetting password for profile: " ($ConfigFile.Config.Company.UserBaseName + $i)
						$TempPass= (ConvertTo-SecureString ([system.web.security.membership]::GeneratePassword(120,32)).tostring() -AsPlainText -Force)
						Set-LocalUser -Name ($ConfigFile.Config.Company.UserBaseName + $i).ToLower() -Password $TempPass
						Write-Host "`tWorking on Creating user profile: " ($ConfigFile.Config.Company.UserBaseName + $i)
						#launch process as user to create user profile
						# https://msdn.microsoft.com/en-us/library/system.diagnostics.processstartinfo(v=vs.110).aspx
						$processStartInfo = New-Object System.Diagnostics.ProcessStartInfo
						$processStartInfo.UserName = ($ConfigFile.Config.Company.UserBaseName + $i)
						$processStartInfo.Domain = "."
						$processStartInfo.Password = $TempPass
						$processStartInfo.FileName = "cmd"
						$processStartInfo.WorkingDirectory = $LICache
						$processStartInfo.Arguments = "/C echo . && echo %username% && echo ."
						$processStartInfo.LoadUserProfile = $true
						$processStartInfo.UseShellExecute = $false
						$processStartInfo.WindowStyle  = "minimized"
						$processStartInfo.RedirectStandardOutput = $false
						$process = [System.Diagnostics.Process]::Start($processStartInfo)
						$Process.WaitForExit()   
						#Add setup user to profiles created to allow registry to be created. 
						If (Test-Path ($UsersProfileFolder + "\Window" + $i) ) {
							$ProfileList.Add(($ConfigFile.Config.Company.UserBaseName + $i).ToLower()) | Out-Null
							$HideAccounts += ($ConfigFile.Config.Company.UserBaseName + $i).ToLower()
							#Grant Current user rights on new Profiles
							Write-Host ("`tUpdating ACLs and adding to Profile List: " + ($UsersProfileFolder + "\Window" + $i))
							$user_account=$env:username
							$Acl = Get-Acl ($UsersProfileFolder + "\Window" + $i)
							$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule($user_account, "FullControl", "ContainerInherit, ObjectInherit", "None", "Allow")
							$Acl.Setaccessrule($Ar)
							Set-Acl ($UsersProfileFolder + "\Window" + $i) $Acl
							#Disable User.
							Disable-LocalUser -Name ($ConfigFile.Config.Company.UserBaseName + $i).ToLower() -Confirm:$false
						}
					}
				}
				
			}
		}
	}
}
#endregion Create Local Store users
#region Disable Local Administrator
#If not logged in as administrator and administrators groups has more than one user set administrator account with random password.
If ($env:username -ne "Administrator") {
	If ((Get-LocalGroupMember -Name 'Administrators').count -gt 1) {
		#Sets Random 265 character password
		set-localuser -Name 'Administrator' -Password (ConvertTo-SecureString ([system.web.security.membership]::GeneratePassword(128,32) + [system.web.security.membership]::GeneratePassword(128,32)).tostring() -AsPlainText -Force )
		Disable-LocalUser -Name 'Administrator' -Confirm:$false
	}
}
#endregion Disable Local Administrator
#============================================================================
#endregion Main Setup
#============================================================================
#============================================================================
#region Main Local Start Menu and Taskbar Settings
#============================================================================
#Import Start Menu Layout
If (-Not $UserOnly) {
	If ([environment]::OSVersion.Version.Major -ge 10 -and (Get-Command Import-StartLayout -ErrorAction SilentlyContinue)) {
		ForEach ($ProfileLocation in (Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList" | Foreach-object { $_.GetValue("ProfileImagePath")})) {
			#Fix "Import-StartLayout : Access to the path" issue 
			If (Test-Path -Path ($ProfileLocation + "\AppData\Local\Microsoft\Windows\Shell\LayoutModification.xml")) {
				Remove-Item -Force -Path ($ProfileLocation + "\AppData\Local\Microsoft\Windows\Shell\LayoutModification.xml")
			}
		}
		#Fix "Import-StartLayout : Access to the path" issue Default user
		If (Test-Path -Path ((Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList" ).Default + "\AppData\Local\Microsoft\Windows\Shell\LayoutModification.xml")) {
			Remove-Item -Force -Path ((Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList" ).Default + "\AppData\Local\Microsoft\Windows\Shell\LayoutModification.xml")
		}
		#If no Command override use XML
		If (-Not $StartLayoutXML) {
			If ($Store) {
				$StartLayoutXML = ($ConfigFile.Config.WindowsSettings.StartMenuXml.ChildNodes | Where-Object {$_.store -eq "true"} | Select-Object -First 1).'#text'
			}else {
				$StartLayoutXML = ($ConfigFile.Config.WindowsSettings.StartMenuXml.ChildNodes | Where-Object {$_.store -eq "false"} | Select-Object -First 1).'#text'
			}
		}
		If (Test-Path ($LICache + "\" + $StartLayoutXML)) {
			write-host ("Setting Taskbar and Start Menu: " + ($LICache + "\" + $StartLayoutXML))
			Import-StartLayout -LayoutPath ($LICache + "\" + $StartLayoutXML) -MountPath ($env:systemdrive + "\") | Out-Null
			Copy-Item -Path ($LICache + "\" + $StartLayoutXML) -Destination 'C:\Windows\OEM\TaskbarLayoutModification.xml' -Force -Confirm:$false
			Copy-Item -Path ($LICache + "\" + $StartLayoutXML) -Destination 'C:\Recovery\AutoApply\TaskbarLayoutModification.xml' -Force -Confirm:$false
		}
	}
}
#============================================================================
#endregion Main Local Start Menu and Taskbar Settings
#============================================================================
#============================================================================
#region Main Set User Defaults 
#============================================================================
Write-Host ("-"*[console]::BufferWidth)
Write-Host ("Starting User Profile Setup")
Write-Host ("-"*[console]::BufferWidth)
$UserProgress = 1
ForEach ( $CurrentProfile in $ProfileList.ToArray() ) {
	Write-Progress -ID 0 -Activity 'Hardening User Profiles' -CurrentOperation $CurrentProfile -PercentComplete (($UserProgress / $ProfileList.count) * 100)
	# write-host ("Working with user: " + $CurrentProfile) -foregroundcolor "Magenta"
	Write-Color "Working with user: ",
				$CurrentProfile -Color White,Magenta
	$HKEY = ("HKU\H_" + $CurrentProfile)
	If (-Not (Test-Path $HKEY)) {
		#region Load User Regsitry
		# Default user need to be handled differently
		If ($CurrentProfile.ToUpper() -eq "DEFAULT") {
			$CurrentUserSID = $null
			#See if profile is in the default location
			If (Test-Path ($UsersProfileFolder + "\Default\ntuser.dat")) {
				[gc]::collect()
				$process = (REG LOAD  $HKEY ($UsersProfileFolder + "\Default\ntuser.dat"))
				If ($LASTEXITCODE -ne 0 ) {
					write-error ( "Cannot load profile for: " + ($UsersProfileFolder + "\" + $CurrentProfile + "\ntuser.dat") )
					continue
				}
			} else {
				#See if Profile is in the systemdrive instead.
				If (Test-Path ($UsersProfileFolder.Replace($UsersProfileFolder.Substring(0,1),($env:systemdrive).Substring(0,1)) + "\Default\ntuser.dat")) {
					[gc]::collect()
					$process = (REG LOAD $HKEY ($UsersProfileFolder.Replace($UsersProfileFolder.Substring(0,1),($env:systemdrive).Substring(0,1)) + "\Default\ntuser.dat"))
					If ($LASTEXITCODE -ne 0 ) {
						write-error ( "Cannot load profile for: " + ($UsersProfileFolder.Replace($UsersProfileFolder.Substring(0,1),($env:systemdrive).Substring(0,1)) + "\Default\ntuser.dat") )
						continue
					}
				}
			}
		}else{
			$CurrentUserSID = (Get-LocalUser -Name $CurrentProfile).SID
			If (Get-Command Get-CimInstance -errorAction SilentlyContinue) {
				$UserProfile = (Get-CimInstance Win32_UserProfile | Where-Object { $_.SID -eq $CurrentUserSID}).localpath
			} Else {
				$UserProfile = (Get-WmiObject Win32_UserProfile | Where-Object { $_.SID -eq $CurrentUserSID}).localpath
			} 
			#If profile is loaded modify loaded profile	
			If (Test-Path ("HKU:\" + $CurrentUserSID)) {
				$HKEY = ("HKU\" + $CurrentUserSID)
			} Else {
				#See if profile location is in the regsitry
				If (Test-Path ($UserProfile + "\ntuser.dat")) { 
					#Load User Hive
					#REG LOAD $HKEY ($UserProfile + "\ntuser.dat")
					[gc]::collect()
					$process = (REG LOAD  $HKEY ($UserProfile + "\ntuser.dat"))
					If ($LASTEXITCODE -ne 0 ) {
						write-error ( "Cannot load profile for: " + ($UsersProfileFolder + "\" + $CurrentProfile + "\ntuser.dat") )
						continue
					}
				}else{
					#See if the profile location is on the system drive
					Try {
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
					Catch {
						write-error ( "Cannot load profile for: " + ($UsersProfileFolder + "\" + $CurrentProfile + "\ntuser.dat") )
						continue
					}
				}
			}
			#Reset Start Menu Windows 1809+
			If ($StartLayoutXML) {
				Remove-Item ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\CloudStore\Store\Cache\DefaultAccount\*$start.tilegrid$windows.data.curatedtilecollection.tilecollection")  -Force -Recurse -ErrorAction SilentlyContinue | Out-Null
			}
			#File permission that would not work with default profile.
			If ($Store) {
				#Add AllowFolder ACL
				ForEach ( $file in (($ConfigFile.Config.Permissions.FileSystem.AllowFolder.Item | Where-Object {$_.store -eq 'true' -and $_.MinimumVersion -ge ([environment]::OSVersion.Version.Major)}).'#text')) {
					If (-Not (Test-Path (Get-envValueFromString -Path $file))){
						New-Item -Path (Get-envValueFromString -Path $file) -Force
					} 
					Write-Host ("`t`tAllowing: " + (Get-envValueFromString -Path $file))
					$Acl = Get-Acl(Get-envValueFromString -Path $file)
					$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule(($env:computer + "\" + $CurrentProfile), "Modify", "Allow")
					$Acl.Setaccessrule($Ar)
					Set-Acl (Get-envValueFromString -Path $file) $Acl
				}
				#Add Deny ACL
				ForEach ( $file in (($ConfigFile.Config.Permissions.FileSystem.Deny.Item | Where-Object {$_.store -eq 'true' -and $_.MinimumVersion -ge ([environment]::OSVersion.Version.Major)}).'#text')) {
					If (Test-Path (Get-envValueFromString -Path $file)) {
						Write-Host ("`t`tDenying: " + (Get-envValueFromString -Path $file))
						$Acl = Get-Acl (Get-envValueFromString -Path $file)
						$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule(($env:computer + "\" + $CurrentProfile), "ReadAndExecute", "Deny")
						$Acl.Setaccessrule($Ar)
						Set-Acl (Get-envValueFromString -Path $file) $Acl	
					} else {
						Write-Warning ("`t`tCannot find '" + (Get-envValueFromString -Path $file) + "' to deny access to.")
					}
				}
				#Add Deny ACL User Profile
				ForEach ( $file in (($ConfigFile.Config.Permissions.FileSystem.DenyUser.Item | Where-Object {$_.store -eq 'true' -and $_.MinimumVersion -ge ([environment]::OSVersion.Version.Major)}).'#text')) {
					If (Test-Path ($UserProfile + "\"+ $file)) {
						Write-Host ("`t`tDenying: " + $file)
						$Acl = Get-Acl ($UserProfile + "\"+ $file)
						$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule(($env:computer + "\" + $CurrentProfile), "ReadAndExecute", "Deny")
						$Acl.Setaccessrule($Ar)
						Set-Acl ($UserProfile + "\"+ $file) $Acl	
						Get-ChildItem -path ($UserProfile + "\"+ $file) -Recurse -Force | ForEach-Object {$_.attributes = "Hidden"}
					} else {
						Write-Warning ("`t`tCannot find '" + ($UserProfile + "\"+ $file) + "' to deny access to.")
					}
				}
				
			}else {
				#Add AllowFolder ACL
				ForEach ( $file in (($ConfigFile.Config.Permissions.FileSystem.AllowFolder.Item | Where-Object {$_.store -eq 'false' -and $_.MinimumVersion -ge ([environment]::OSVersion.Version.Major)}).'#text')) {
					If (-Not (Test-Path (Get-envValueFromString -Path $file))){
						New-Item -Path (Get-envValueFromString -Path $file) -Force
					} 
					Write-Host ("`t`tAllowing: " + (Get-envValueFromString -Path $file))
					$Acl = Get-Acl (Get-envValueFromString -Path $file)
					$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule(($env:computer + "\" + $CurrentProfile), "Modify", "Allow")
					$Acl.Setaccessrule($Ar)
					Set-Acl (Get-envValueFromString -Path $file) $Acl
				}
				#Add Deny ACL
				ForEach ( $file in (($ConfigFile.Config.Permissions.FileSystem.Deny.Item | Where-Object {$_.store -eq 'false' -and $_.MinimumVersion -ge ([environment]::OSVersion.Version.Major)}).'#text')) {
					If (Test-Path (Get-envValueFromString -Path $file)) {
						Write-Host ("`t`tDenying: " + (Get-envValueFromString -Path $file))
						$Acl = Get-Acl (Get-envValueFromString -Path $file)
						$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule(($env:computer + "\" + $CurrentProfile), "ReadAndExecute", "Deny")
						$Acl.Setaccessrule($Ar)
						Set-Acl (Get-envValueFromString -Path $file) $Acl	
					} else {
						Write-Warning ("`t`tCannot find '" + (Get-envValueFromString -Path $file) + "' to deny access to.")
					}
				}
				#Add Deny ACL User Profile
				ForEach ( $file in (($ConfigFile.Config.Permissions.FileSystem.DenyUser.Item | Where-Object {$_.store -eq 'false' -and $_.MinimumVersion -ge ([environment]::OSVersion.Version.Major)}).'#text')) {
					If (Test-Path ($UserProfile + "\"+ $file)) {
						Write-Host ("`t`tDenying: " + $file)
						$Acl = Get-Acl ($UserProfile + "\"+ $file)
						$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule(($env:computer + "\" + $CurrentProfile), "ReadAndExecute", "Deny")
						$Acl.Setaccessrule($Ar)
						Set-Acl ($UserProfile + "\"+ $file) $Acl	
						Get-ChildItem -path ($UserProfile + "\"+ $file) -Recurse -Force | ForEach-Object {$_.attributes = "Hidden"}
					} else {
						Write-Warning ("`t`tCannot find '" + ($UserProfile + "\"+ $file) + "' to deny access to.")
					}
				}
			}
		}		
		#endregion Load User Regsitry
		#region Set User 
			#region Registry Setup
				#Update/Add Items Values
				write-host ("`tUpdating Registry Settings:")
				Foreach ($key in ($ConfigFile.Config.UserSettings.UserRegistry.Item | Where-Object {$_.MinimumVersion -ge ([environment]::OSVersion.Version.Major)})) {
					If ($key.Default -eq 'true' -and $key.Store -eq 'false' -and $CurrentProfile.ToUpper() -eq "DEFAULT") {		
						Write-Color -Text "Default User:  ",
											$key.Comment -Color Blue,DarkGray -StartTab 2
						Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\" + $key.Key) $key.Value $key.Data $key.Type
					} ElseIf ($LockedDown -and $key.LockedDown -eq 'true' -and $key.Store -eq 'false' -and $CurrentProfile.ToUpper() -ne "DEFAULT") {		
						Write-Color -Text "LockedDown:  ",
											$key.Comment -Color Blue,DarkGray -StartTab 2
						Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\" + $key.Key) $key.Value $key.Data $key.Type
					} ElseIf ($Store -and $key.Store -eq 'true' -and $CurrentProfile.ToUpper() -ne "DEFAULT") {				
						Write-Color -Text "Store:       ",
											$key.Comment -Color Red,DarkGray -StartTab 2
						Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\" + $key.Key) $key.Value $key.Data $key.Type
					} ElseIf ($Manager -and $key.Manager -eq 'true' -and $CurrentProfile.ToUpper() -ne "DEFAULT") {				
						Write-Color -Text "Manager:     ",
											$key.Comment -Color Yellow,DarkGray -StartTab 2
						Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\" + $key.Key) $key.Value $key.Data $key.Type
					} ElseIf ( $key.Store -eq 'false'-and $key.LockedDown -eq 'false' -and $key.Manager -eq 'false' -and $key.Default -eq 'false' -and $CurrentProfile.ToUpper() -ne "DEFAULT") {
						Write-Color -Text "All:         ",
									$key.Comment -Color DarkGreen,DarkGray -StartTab 2
						Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\" + $key.Key) $key.Value $key.Data $key.Type	
					}
				}			
				If ($IsVM) {
					Foreach ($key in ($ConfigFile.Config.UserSettings.VM.UserRegistry.Item | Where-Object { $_.MinimumVersion -ge ([environment]::OSVersion.Version.Major)})) {
						If ($key.Default -eq 'true' -and $key.Store -eq 'false' -and $CurrentProfile.ToUpper() -eq "DEFAULT") {		
							Write-Color -Text "Default User:  ",
												$key.Comment -Color Blue,DarkGray -StartTab 2
							Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\" + $key.Key) $key.Value $key.Data $key.Type
						} ElseIf ($LockedDown -and $key.LockedDown -eq 'true' -and $key.Store -eq 'false' -and $CurrentProfile.ToUpper() -ne "DEFAULT") {		
							Write-Color -Text "LockedDown:  ",
												$key.Comment -Color Blue,DarkGray -StartTab 2
							Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\" + $key.Key) $key.Value $key.Data $key.Type
						} ElseIf ($Store -and $key.Store -eq 'true' -and $CurrentProfile.ToUpper() -ne "DEFAULT") {				
							Write-Color -Text "Store:       ",
												$key.Comment -Color Red,DarkGray -StartTab 2
							Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\" + $key.Key) $key.Value $key.Data $key.Type
						} ElseIf ($Manager -and $key.Manager -eq 'true' -and $CurrentProfile.ToUpper() -ne "DEFAULT") {				
							Write-Color -Text "Manager:     ",
												$key.Comment -Color Yellow,DarkGray -StartTab 2
							Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\" + $key.Key) $key.Value $key.Data $key.Type
						} ElseIf ( $key.Store -eq 'false'-and $key.LockedDown -eq 'false' -and $key.Manager -eq 'false' -and $key.Default -eq 'false' -and $CurrentProfile.ToUpper() -ne "DEFAULT") {
							Write-Color -Text "All:         ",
										$key.Comment -Color DarkGreen,DarkGray -StartTab 2
							Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\" + $key.Key) $key.Value $key.Data $key.Type	
						}
					}
				}
				#Remove Items
				Foreach ($key in ($ConfigFile.Config.UserSettings.UserRegistry.Remove | Where-Object {$_.MinimumVersion -ge ([environment]::OSVersion.Version.Major)})) {
					If ($LockedDown -and $key.LockedDown -eq 'true' -and $key.Store -eq 'false') {				
						Write-Color -Text "LockedDown:  ",
											$key.Comment -Color Blue,DarkGray -StartTab 2
						If ($key.Value) {
							Remove-ItemProperty -Path ($HKEY.replace("HKU\","HKU:\") + "\" + $key.Key) -Name $key.Value -Confirm:$False  -erroraction 'silentlycontinue'
						} else {
							Remove-Item -Path ($HKEY.replace("HKU\","HKU:\") + "\" + $key.Key)  -Confirm:$False  -erroraction 'silentlycontinue'
						}
					} ElseIf ($Store -and $key.Store -eq 'true' -and $CurrentProfile.ToUpper() -ne "DEFAULT") {				
						Write-Color -Text "Store:       ",
											$key.Comment -Color Red,DarkGray -StartTab 2
						If ($key.Value) {
							Remove-ItemProperty -Path ($HKEY.replace("HKU\","HKU:\") + "\" + $key.Key) -Name $key.Value -Confirm:$False  -erroraction 'silentlycontinue'
						} else {
							Remove-Item -Path ($HKEY.replace("HKU\","HKU:\") + "\" + $key.Key)  -Confirm:$False  -erroraction 'silentlycontinue'
						}
					} ElseIf ($Manager -and $key.Manager -eq 'true' -and $CurrentProfile.ToUpper() -ne "DEFAULT") {				
						Write-Color -Text "Manager:     ",
											$key.Comment -Color Yellow,DarkGray -StartTab 2
						If ($key.Value) {
							Remove-ItemProperty -Path ($HKEY.replace("HKU\","HKU:\") + "\" + $key.Key) -Name $key.Value -Confirm:$False  -erroraction 'silentlycontinue'
						} else {
							Remove-Item -Path ($HKEY.replace("HKU\","HKU:\") + "\" + $key.Key)  -Confirm:$False  -erroraction 'silentlycontinue'
						}
					} ElseIf ( $key.Store -eq 'false'-and $key.LockedDown -eq 'false' -and $key.Manager -eq 'false') {
						Write-Color -Text "All:         ",
									$key.Comment -Color DarkGreen,DarkGray -StartTab 2
						If ($key.Value) {
							Remove-ItemProperty -Path ($HKEY.replace("HKU\","HKU:\") + "\" + $key.Key) -Name $key.Value -Confirm:$False  -erroraction 'silentlycontinue'
						} else {
							Remove-Item -Path ($HKEY.replace("HKU\","HKU:\") + "\" + $key.Key)  -Confirm:$False  -erroraction 'silentlycontinue'
						}	
					}
				}
				#Add Keys
				Foreach ($key in ($ConfigFile.Config.UserSettings.UserRegistry.Add | Where-Object {$_.MinimumVersion -ge ([environment]::OSVersion.Version.Major)})) {
					If ($LockedDown -and $key.LockedDown -eq 'true' -and $key.Store -eq 'false') {				
						Write-Color -Text "LockedDown:  ",
											$key.Comment -Color Blue,DarkGray -StartTab 2
						If ($key.key -and -Not (Test-Path ($HKEY.replace("HKU\","HKU:\") + "\" + $key.Key))) {
							New-Item -Path ($HKEY.replace("HKU\","HKU:\") + "\" + $key.Key) -Force | Out-Null
						}
					} ElseIf ($Store -and $key.Store -eq 'true' -and $CurrentProfile.ToUpper() -ne "DEFAULT") {				
						Write-Color -Text "Store:       ",
											$key.Comment -Color Red,DarkGray -StartTab 2
						If ($key.key -and -Not (Test-Path ($HKEY.replace("HKU\","HKU:\") + "\" + $key.Key))) {
							New-Item -Path ($HKEY.replace("HKU\","HKU:\") + "\" + $key.Key) -Force | Out-Null
						}
					} ElseIf ($Manager -and $key.Manager -eq 'true' -and $CurrentProfile.ToUpper() -ne "DEFAULT") {				
						Write-Color -Text "Manager:     ",
											$key.Comment -Color Yellow,DarkGray -StartTab 2
						If ($key.key -and -Not (Test-Path ($HKEY.replace("HKU\","HKU:\") + "\" + $key.Key))) {
							New-Item -Path ($HKEY.replace("HKU\","HKU:\") + "\" + $key.Key) -Force | Out-Null
						}
					} ElseIf ( $key.Store -eq 'false'-and $key.LockedDown -eq 'false' -and $key.Manager -eq 'false') {
						Write-Color -Text "All:         ",
									$key.Comment -Color DarkGreen,DarkGray -StartTab 2
						If ($key.key -and -Not (Test-Path ($HKEY.replace("HKU\","HKU:\") + "\" + $key.Key))) {
							New-Item -Path ($HKEY.replace("HKU\","HKU:\") + "\" + $key.Key) -Force | Out-Null
						}
					}
				}
			#endregion Registry Setup 
			#region Set Non-Store 
				#region OneDrive
					If ($ConfigFile.Config.UserSettings.DisableOnedrive -eq "true") {
						write-host ("`tRemove OneDrive:")
						Remove-Itemproperty -Path ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Run") -name 'OneDriveSetup' -erroraction 'silentlycontinue'| out-null
					}
				#endregion OneDrive
				#region Windows Explorer, Force Enable Basic Settings
					#Show This PC
					Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel") "{20D04FE0-3AEA-1069-A2D8-08002B30309D}" 0 "DWORD"
					Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\ClassicStartMenu") "{20D04FE0-3AEA-1069-A2D8-08002B30309D}" 0 "DWORD"
					#Show Frequent Access
					Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Explorer") "ShowFrequent" 1 "DWORD"
					Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Explorer") "ShowRecent" 1 "DWORD"
					# Change Explorer home screen back to ""Quick Access"
					Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced") "LaunchTo" 2 "DWORD"	
				#endregion Windows Explorer, Force Enable Basic Settings
				#region Internet Explorer
					write-host ("`tSetting up Internet Explorer:")
					#Using Script block to allow for changing x86 and x64 IE registies
					[scriptblock]$IEScriptBlock = {
						#MigrateProxy
						If ($ConfigFile.Config.IE.AutoDetect) {
							Set-Reg $HKEYIS "AutoDetect" $ConfigFile.Config.IE.AutoDetect "DWORD" 
						}
						#ProxyEnable
						If ($ConfigFile.Config.IE.ProxyEnable) {
							Set-Reg $HKEYIS "ProxyEnable" $ConfigFile.Config.IE.ProxyEnable "DWORD" 
						}
						#CacheScripts
						If ($ConfigFile.Config.IE.EnableAutoProxyResultCache) {
							Set-Reg $HKEYIS "EnableAutoProxyResultCache" $ConfigFile.Config.IE.EnableAutoProxyResultCache "DWORD" 
						}
						#Set SSL Caching
						If ($ConfigFile.Config.IE.DisableCachingOfSSLPages) {
							Set-Reg $HKEYIS "DisableCachingOfSSLPages" $ConfigFile.Config.IE.DisableCachingOfSSLPages "DWORD" 
							Set-Reg "HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings" "DisableCachingOfSSLPages" $ConfigFile.Config.IE.DisableCachingOfSSLPages "DWORD"
						}
						#Enable changing Automatic Configuration settings
						If ($ConfigFile.Config.IE.EnableUserChangingProxySettings) {
							Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Control Panel") "Autoconfig" $ConfigFile.Config.IE.EnableUserChangingProxySettings "DWORD"
						}
						#AutoConfigProxy
						If ($ConfigFile.Config.IE.AutoConfigProxy) {
							#Enable Auto Config of Proxy
							$temp = (Get-ItemProperty -Path ($HKEYIS + "\Connections") -name "DefaultConnectionSettings" -erroraction 'silentlycontinue').DefaultConnectionSettings  | out-null
							if (!($temp)) {
								$temp = (70,0,0,0,3,0,0,0,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)
							} 
							$temp[8] = $ConfigFile.Config.IE.AutoConfigProxy
							Set-Reg ($HKEYIS + "\Connections") "DefaultConnectionSettings" $temp  "Binary"		
						}
						#PopupsUseNewWindow
						If ($ConfigFile.Config.IE.PopupsUseNewWindow) {
							Set-Reg ($HKEYIE + "\TabbedBrowsing") "PopupsUseNewWindow" $ConfigFile.Config.IE.PopupsUseNewWindow "DWORD"
						}
						#PhishingFilter
						If ($ConfigFile.Config.IE.PhishingFilter) {
							Set-Reg ($HKEYIE + "\PhishingFilter") "Enabled" $ConfigFile.Config.IE.PhishingFilter "DWORD"
						}
						#Enable AutoImageResize
						If ($ConfigFile.Config.IE.AutoImageResize) {
							Set-Reg ($HKEYIE + "\Main") "Enable AutoImageResize" $ConfigFile.Config.IE.AutoImageResize "String"
						}		
						#Homepage
						If ($ConfigFile.Config.Company.HomePage) {
							Set-Reg ($HKEYIE + "\Main") "Start Page" $ConfigFile.Config.Company.HomePage "String"
						}		
						#PageSetup header
						If ($ConfigFile.Config.IE.Header) {
							Set-Reg ($HKEYIE + "\PageSetup") "header" $ConfigFile.Config.IE.Header "String"
						} else {
							Set-Reg ($HKEYIE + "\PageSetup") "header" "" "String"
						}		
						#PageSetup footer
						If ($ConfigFile.Config.IE.footer) {
							Set-Reg ($HKEYIE + "\PageSetup") "footer" $ConfigFile.Config.IE.footer "String"
						} else {
							Set-Reg ($HKEYIE + "\PageSetup") "footer" "" "String"
						}		
						#PageSetup margin_bottom
						If ($ConfigFile.Config.IE.margin_bottom) {
							Set-Reg ($HKEYIE + "\PageSetup") "margin_bottom" $ConfigFile.Config.IE.margin_bottom "String"
						} else {
							Set-Reg ($HKEYIE + "\PageSetup") "margin_bottom" "" "String"
						}		
						#PageSetup margin_top
						If ($ConfigFile.Config.IE.margin_top) {
							Set-Reg ($HKEYIE + "\PageSetup") "margin_top" $ConfigFile.Config.IE.margin_top "String"
						} else {
							Set-Reg ($HKEYIE + "\PageSetup") "margin_top" "" "String"
						}		
						#PageSetup margin_left
						If ($ConfigFile.Config.IE.margin_left) {
							Set-Reg ($HKEYIE + "\PageSetup") "margin_left" $ConfigFile.Config.IE.margin_left "String"
						} else {
							Set-Reg ($HKEYIE + "\PageSetup") "margin_left" "" "String"
						}		
						#PageSetup margin_right
						If ($ConfigFile.Config.IE.margin_right) {
							Set-Reg ($HKEYIE + "\PageSetup") "margin_right" $ConfigFile.Config.IE.margin_right "String"
						} else {
							Set-Reg ($HKEYIE + "\PageSetup") "margin_right" "" "String"
						}	
						#CacheLimit in MB; will convirt to KB
						If ($ConfigFile.Config.IE.Cache_Size) {
							Set-Reg ($HKEYIS + "\5.0\Cache\Content\CacheLimit") "CacheLimit" ([int]([int]$ConfigFile.Config.IE.Cache_Size * 1024)) "DWORD"
							Set-Reg ($HKEYIS + "\Cache\Content\CacheLimit") "CacheLimit" ([int]([int]$ConfigFile.Config.IE.Cache_Size * 1024)) "DWORD"
						}
						#Cache Persistents
						If ($ConfigFile.Config.IE.Cache_Persistent) {
							Set-Reg ($HKEYIS + "\Cache") "CacheLimit" $ConfigFile.Config.IE.Cache_Persistent "DWORD"
						}
						#region Zones Setup
						If ($ConfigFile.Config.IE.ZoneMaps.ZoneMap.Count -gt 0) {
							#Zones Cleanup
							If (Test-Path ($HKEYIS + "\ZoneMap\Domains")) {
								Remove-Item ($HKEYIS + "\ZoneMap\Domains") -Recurse -Confirm:$false | out-null
							}
							If (Test-Path ($HKEYIS + "\ZoneMap\EscDomains")) {
								Remove-Item ($HKEYIS + "\ZoneMap\EscDomains") -Recurse -Confirm:$false | out-null
							}
							#Setup Zone  
							ForEach ( $item in $ConfigFile.Config.IE.ZoneMaps.ZoneMap) {
								Switch ($item.Zone) {
									0 {$IEZone = "My Computer"}
									1 {$IEZone = "Local Intranet Zone"}
									2 {$IEZone = "Trusted sites Zone"}
									3 {$IEZone = "Internet Zone"}
									4 {$IEZone = "Restricted Sites Zone"}
								}
								# write-host ("`t`tAdding Site: " + $item.Site + " to Zone: " + $IEZone + " for Protocol: " + $item.Protocol)
								write-color -Text   "Adding Site: ",
													$item.Site,
													" to Zone: ",
													$IEZone,
													" for Protocol: ",
													$item.Protocol -ForegroundColor  White,Red,White,Magenta,White,Cyan -StartTab 2
								Set-Reg ($HKEYIS + "\ZoneMap\Domains\" +  $item.Site) $item.Protocol $item.Zone "DWORD"
								Set-Reg ($HKEYIS + "\ZoneMap\EscDomains\" +  $item.Site) $item.Protocol $item.Zone "DWORD"
							}
						}
						#endregion Zones Setup
					}	
					#For x86			
					$HKEYIS = ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Internet Settings")
					$HKEYIE = ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Internet Explorer")
					Invoke-Command -ScriptBlock $IEScriptBlock
					#For X64
					If ([Environment]::Is64BitOperatingSystem) {
						$HKEYIS = $HKEYIS.replace("\Software\","\Software\Wow6432Node\")
						$HKEYIE = $HKEYIE.replace("\Software\","\Software\Wow6432Node\")
						Invoke-Command -ScriptBlock $IEScriptBlock
					}
				#endregion Internet Explorer
				#region Windows Media Player
					write-host ("`tSetting up Windows Media Player")
					#DesktopShortcut
					If ($ConfigFile.Config.Windows_Media_Player.DesktopShortcut) {
						Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\MediaPlayer\Setup\UserOptions") "DesktopShortcut" $ConfigFile.Config.Windows_Media_Player.DesktopShortcut "String" 
					}		
					#QuickLaunchShortcut
					If ($ConfigFile.Config.Windows_Media_Player.QuickLaunchShortcut) {
						Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\MediaPlayer\Setup\UserOptions") "QuickLaunchShortcut" $ConfigFile.Config.Windows_Media_Player.QuickLaunchShortcut "DWORD" 
					}		
					#AcceptedPrivacyStatement
					If ($ConfigFile.Config.Windows_Media_Player.AcceptedPrivacyStatement) {
						Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\MediaPlayer\Preferences") "AcceptedPrivacyStatement" $ConfigFile.Config.Windows_Media_Player.AcceptedPrivacyStatement "DWORD" 
					}		
					#FirstRun
					If ($ConfigFile.Config.Windows_Media_Player.FirstRun) {
						Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\MediaPlayer\Preferences") "FirstRun" $ConfigFile.Config.Windows_Media_Player.FirstRun "DWORD" 
					}		
					#DisableMRU
					If ($ConfigFile.Config.Windows_Media_Player.DisableMRU) {
						Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\MediaPlayer\Preferences") "DisableMRU" $ConfigFile.Config.Windows_Media_Player.DisableMRU "DWORD" 
					}		
					#AutoCopyCD
					If ($ConfigFile.Config.Windows_Media_Player.AutoCopyCD) {
						Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\MediaPlayer\Preferences") "AutoCopyCD" $ConfigFile.Config.Windows_Media_Player.AutoCopyCD "DWORD" 
					}		
				#endregion Windows Media Player
				#region Chrome Setup
					#region Remove Chrome Settings
					If (Test-Path ($UserProfile + "\AppData\Local\Google")) {
						Remove-Item -Path ($UserProfile + "\AppData\Local\Google") -Recurse -Confirm:$false | out-null
					}
					If (Test-Path ($HKEY.replace("HKU\","HKU:\") + "\SOFTWARE\Google")) {
						Remove-Item -Recurse -Confirm:$false -Path ($HKEY.replace("HKU\","HKU:\") + "\SOFTWARE\Google") -erroraction 'silentlycontinue'
					}
					#endregion Remove Chrome Settings
					#region Deploy Chrome Base Profile
					If ($ConfigFile.Config.Chrome.BaseZip) {
						If (Test-Path ($LICache + "\" + $ConfigFile.Config.Chrome.BaseZip)) {
							If (Test-Path ($UserProfile + "\AppData\Local")) {
								Write-Host ("`tSetting-up Chrome Base Settings")
								$ProgressPreference = 'SilentlyContinue'
								Expand-Archive -Path ($LICache + "\" + $ConfigFile.Config.Chrome.BaseZip) -DestinationPath ($UserProfile + "\AppData\Local") -Force 
								$ProgressPreference = "Continue"
							}
						}
					}
					#endregion Deploy Chrome Base Profile			
				#endregion Chrome Setup
			#endregion Set Non-Store
			#region Set LockedDown
				#Chrome
				#Disables all extensions
				If ($ConfigFile.Config.Chrome.DisablesExtensions -eq "True") {
					If (Test-Path ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome\ExtensionInstallBlacklist")) {
						Remove-Item ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome\ExtensionInstallBlacklist") -Recurse | out-null
					}
					Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome\ExtensionInstallBlacklist") 1 "*" "String"
				}
				#Sets Startup page
				If ($ConfigFile.Config.Company.HomePage){
					If (Test-Path ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome\RestoreOnStartupURLs")) {
						Remove-Item ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome\RestoreOnStartupURLs") -Recurse | out-null
					}
					Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome\RestoreOnStartupURLs") 1 $ConfigFile.Config.Company.HomePage "String"
				}			
				#$ChromeURLBlackList Stops local browsing
				If ($ConfigFile.Config.Chrome.BlackListURLs.URL.Count -ge 1) {
					If (Test-Path($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome\URLBlacklist")) {
						Remove-Item ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome\URLBlacklist") -Recurse
					}
					$i = 1
					ForEach ( $item in $ConfigFile.Config.Chrome.BlackListURLs.URL) {
						Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Google\Chrome\URLBlacklist") $i $item "String"
						$i++
					}	
				}	 
			#endregion Set LockedDown
			#region Set Store 
				If (($Store) -and ($CurrentProfile.ToUpper() -ne "DEFAULT" )) {
					#region Company Software Auto Start
					If ($ConfigFile.Config.Company.ExccutableName -and $ConfigFile.Config.Company.SoftwarePath) {
						If (-Not (Test-Path ($UsersProfileFolder + "\" + $CurrentProfile + "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\" + $ConfigFile.Config.Company.ExccutableName.Substring(0,$ConfigFile.Config.Company.ExccutableName.IndexOfAny(".")) + ".lnk"))) {
							If (-Not (Test-Path ($UsersProfileFolder + "\" + $CurrentProfile + "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"))) {
								New-Item -Force -ItemType "directory" -Path ($UsersProfileFolder + "\" + $CurrentProfile + "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup")
							}
							If (Test-Path ($ConfigFile.Config.Company.SoftwarePath + "\" + $ConfigFile.Config.Company.ExccutableName)) {
								$ShortCut = $WScriptShell.CreateShortcut($UsersProfileFolder + "\" + $CurrentProfile + "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\" + $ConfigFile.Config.Company.ExccutableName.Substring(0,$ConfigFile.Config.Company.ExccutableName.IndexOfAny(".")) + ".lnk")
								$ShortCut.TargetPath=($ConfigFile.Config.Company.SoftwarePath + "\" + $ConfigFile.Config.Company.ExccutableName)
								$ShortCut.WorkingDirectory = ($ConfigFile.Config.Company.SoftwarePath)
								$ShortCut.IconLocation = ( $ConfigFile.Config.Company.SoftwarePath + "\" + $ConfigFile.Config.Company.ExccutableName + ",0")
								$ShortCut.Save()
								#Make ShortCut ran as admin https://stackoverflow.com/questions/28997799/how-to-create-a-run-as-administrator-shortcut-using-powershell
								$bytes = [System.IO.File]::ReadAllBytes($UsersProfileFolder + "\" + $CurrentProfile + "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\" + $ConfigFile.Config.Company.ExccutableName.Substring(0,$ConfigFile.Config.Company.ExccutableName.IndexOfAny(".")) + ".lnk")
								$bytes[0x15] = $bytes[0x15] -bor 0x20 #set byte 21 (0x15) bit 6 (0x20) ON
								[System.IO.File]::WriteAllBytes($UsersProfileFolder + "\" + $CurrentProfile + "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\" + $ConfigFile.Config.Company.ExccutableName.Substring(0,$ConfigFile.Config.Company.ExccutableName.IndexOfAny(".")) + ".lnk", $bytes)
							}
						}
					}
					#endregion Company Software Auto Start
					#region Deny Programs to run
					If ($ConfigFile.Config.WindowsSettings.BlackListPrograms.Block.count -ge 1) {
						write-host ("`tSetting up Store settings Deny Programs")
						#Cleanup old
						If (Test-Path ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun")) {
							Remove-Item ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun") -Recurse | out-null
						}
						ForEach ( $Exe in $ConfigFile.Config.WindowsSettings.BlackListPrograms.Block) {
							# write-host ("`t`tBlackListing: " + $Exe)
							Write-Color -Text "BlackListing: ",
												$Exe -Color White,Red -StartTab 2
							Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun") $i $Exe "String"
							$i++
						}
					}
					#endregion Deny Programs to run		
				}
			#endregion Set Store
			#region Set Active User 
				If (($ActiveUser.ToUpper() -eq $CurrentProfile.ToUpper()) -and ($CurrentProfile.ToUpper() -ne "DEFAULT" )) {
					If (Get-LocalUser $CurrentProfile) {
						Write-Color -Text "Enabling User: ",
									$CurrentProfile -Color White,Green -StartTab 2
						Enable-LocalUser $CurrentProfile -Confirm:$false
					}
					If ($Manager) {

					}

				}
			#endregion Actvie User 
		#endregion Set User
	}#End if User Exsits

	#region Replace Favorites
	If (Test-Path ($LICache + "\Favorites")) {
		write-host ("`tSetting up Favorites")
		If ($CurrentProfile -eq "Default") {
			If (Test-Path ($UsersProfileFolder + "\Default\Favorites")) {
				Remove-Item -path ($UsersProfileFolder + "\Default\Favorites") -recurse -force
				Copy-Item  ($LICache + "\Favorites") -Destination ($UsersProfileFolder + "\Default\Favorites") -recurse -force
			}
		}else{
			$CurrentUserSID = (Get-LocalUser -Name $CurrentProfile).SID
			If (Get-Command Get-CimInstance -errorAction SilentlyContinue) {
				$UserProfile = (Get-CimInstance Win32_UserProfile | Where-Object { $_.SID -eq $CurrentUserSID}).localpath
			} Else {
				$UserProfile = (Get-WmiObject Win32_UserProfile | Where-Object { $_.SID -eq $CurrentUserSID}).localpath
			} 
			If (Test-Path ($UserProfile + "\Favorites")) {
				Remove-Item -path ($UserProfile + "\Favorites") -recurse -force
				Copy-Item  ($LICache + "\Favorites") -Destination ($UserProfile + "\Favorites") -recurse -force
			}
		}
	}
	#endregion Replace Favorites
	#region WinX	
	If ($Store -and $CurrentProfile.ToUpper() -ne "DEFAULT") {	
		$WinXZip = ($ConfigFile.Config.WindowsSettings.WinXZip.Item | Where-Object {$_.Store -eq 'true' -and $_.MinimumVersion -ge ([environment]::OSVersion.Version.Major)} | Select-Object -Last 1).'#text'
	
	}else {
		$WinXZip = ($ConfigFile.Config.WindowsSettings.WinXZip.Item | Where-Object {$_.Store -eq 'false' -and $_.MinimumVersion -ge ([environment]::OSVersion.Version.Major)} | Select-Object -Last 1).'#text'
	}
	If ($WinXZip) {
		If (Test-Path ($LICache + "\" + $WinXZip)) {
			If (Test-Path ($UserProfile + "\AppData\Local\Microsoft\Windows\WinX")) {
				Write-Host ("`tSetting-up Win+X Custom Settings")
				#Remove standard entries before adding customized ones. 
				Remove-Item -Recurse -Confirm:$false -Path ($UserProfile + "\AppData\Local\Microsoft\Windows\WinX") -erroraction 'silentlycontinue'
				Expand-Archive -Path ($LICache + "\" + $WinXZip) -DestinationPath ($UserProfile + "\AppData\Local\Microsoft\Windows") -Force
			}
		}
	}
	#endregion WinX		

	#Unload only if use is not logged in. 
	If (-Not (Test-Path ("HKU:\" + $CurrentUserSID))) {
		# Unload the User profile hive
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
	}
	[gc]::collect()
	$UserProgress++
}
Write-Host ("-"*[console]::BufferWidth)
Write-Host ("Ending User Profile Setup:")
Write-Progress -ID 0 -Completed -Activity "Hardening User Profiles Complete."
Write-Host ("-"*[console]::BufferWidth)
#============================================================================
#endregion Main Set User Defaults 
#============================================================================
#============================================================================
#region Import and Set Security Template INI
#============================================================================
If ($ConfigFile.Config.WindowsSettings.SecurityTemplateINI) {
	Write-Host "Importing Security Template Settings ..."
	[string[]]$InICollection = $null 
	#First Manitory Section Unicode
	[string]$LastSection = $null
	ForEach ($ini in ($ConfigFile.Config.WindowsSettings.SecurityTemplateINI.ini | Where-Object {$_.Section -eq "Unicode"}) ) {
		If ($LastSection -ne $ini.Section) {
			#Create new Section
			$InICollection += "[" + $ini.Section + "]"
			If ($ini.Value) {
				If ($ini.Data) {
					$InICollection += $ini.Value + "=" + $ini.Data
				}Else{
					$InICollection += $ini.Value + "="
				}
			}
			$LastSection = $ini.Section
		}Else{
			If ($ini.Value) {
				If ($ini.Data) {
					$InICollection += $ini.Value + "=" + $ini.Data
				}Else{
					$InICollection += $ini.Value + "="
				}
			}
		}
	}
	#Second Manitory Section Version
	[string]$LastSection = $null
	ForEach ($ini in ($ConfigFile.Config.WindowsSettings.SecurityTemplateINI.ini | Where-Object {$_.Section -eq "Version"}) ) {
		If ($LastSection -ne $ini.Section) {
			#Create new Section
			$InICollection += ""
			$InICollection += "[" + $ini.Section + "]"
			If ($ini.Value) {
				If ($ini.Data) {
					$InICollection += $ini.Value + "=" + $ini.Data
				}Else{
					$InICollection += $ini.Value + "="
				}
			}
			$LastSection = $ini.Section
		}Else{
			If ($ini.Value) {
				If ($ini.Data) {
					$InICollection += $ini.Value + "=" + $ini.Data
				}Else{
					$InICollection += $ini.Value + "="
				}
			}
		}
	}
	#All other Sections
	[string]$LastSection = $null
	ForEach ($ini in ($ConfigFile.Config.WindowsSettings.SecurityTemplateINI.ini | Where-Object {$_.Section -ne "Unicode" -and $_.Section -ne "Version"}) ) {
		If ($LastSection -ne $ini.Section) {
			#Create new Section
			$InICollection += ""
			$InICollection += "[" + $ini.Section + "]"
			If ($ini.Value) {
				If ($ini.Data) {
					$InICollection += $ini.Value + "=" + $ini.Data
				}Else{
					If ($ini.Section -eq "File Security") {
						$InICollection += $ini.Value
					}Else{
						$InICollection += $ini.Value + "="
					}
				}
			}
			$LastSection = $ini.Section
		}Else{
			If ($ini.Value) {
				If ($ini.Data) {
					$InICollection += $ini.Value + "=" + $ini.Data
				}Else{
					If ($ini.Section -eq "File Security") {
						$InICollection += $ini.Value
					}Else{
						$InICollection += $ini.Value + "="
					}
				}
			}
		}
	}
	#create file
	If (Test-Path -Path ($env:TMP + "\WHST.inf")) {
		Remove-Item -Force -Confirm:$false -Path ($env:TMP + "\WHST.inf")
	}
	If (Test-Path -Path ($env:TMP + "\WHST.sdb")) {
		Remove-Item -Force -Confirm:$false -Path ($env:TMP + "\WHST.sdb")
	}
	$InICollection | Out-File -FilePath ($env:TMP + "\WHST.inf") 
	#import file
	If (Test-Path -Path ($env:TMP + "\WHST.inf")) {
		Secedit /import /db ($env:TMP + "\WHST.sdb") /cfg ($env:TMP + "\WHST.inf") /quiet
	}
}
#============================================================================
#endregion Import and Set Security Template INI
#============================================================================
#============================================================================
#region Main Local Machine
#============================================================================
If (-Not $UserOnly) {
	#region Windows 10 Only
	If ([environment]::OSVersion.Version.Major -ge 10) {
		#region Windows Feature setup
		Write-Host "Disabling Windows Features:"
		ForEach ( $Feature in $ConfigFile.Config.WindowsSettings.RemoveWindowsFeatures.Remove ) {
			If (Get-Command Get-WindowsOptionalFeature -errorAction SilentlyContinue) {
				If ((Get-WindowsOptionalFeature -Online -FeatureName $Feature).state -eq "Enabled") {
					# Write-Host ("`t" + $Feature) -ForegroundColor gray
					Write-Color -Text "Disabling Windows Optional Feature: ",
										$Feature -Color DarkYellow,White -StartTab 1
					Disable-WindowsOptionalFeature -Online -FeatureName $Feature -NoRestart | out-null
				} else {
					If (Get-WindowsOptionalFeature -Online -FeatureName $Feature) {
						# Write-Host ("`tWindows Optional Feature: " + $Feature + " Already disabled.") -ForegroundColor green
						Write-Color -Text "Disabled Windows Optional Feature: ",
											$Feature -Color DarkGreen,White -StartTab 1
					}
				}
			}
			If (Get-Command Get-WindowsCapability -errorAction SilentlyContinue) {
				If ((Get-WindowsCapability -Online | Where-Object {$_.name -like ("*" + $Feature + "*") -and $_.state -eq "Installed"}).state) {
					# Write-Host ("`t" + $Feature) -ForegroundColor gray
					Write-Color -Text "Disabling Windows Capability: ",
										$Feature -Color DarkYellow,White -StartTab 1
					Get-WindowsCapability -Online | Where-Object {$_.name -like ("*" + $Feature + "*") -and $_.state -eq "Installed"} | Remove-WindowsCapability -online | out-null
				} else {
					If ((Get-WindowsCapability -Online -Name $Feature).Name) {
						# Write-Host ("`tWindows Capability: " + $Feature + " Already disabled.") -ForegroundColor green
						Write-Color -Text "Disabled Windows Capability: ",
										$Feature -Color DarkGreen,White -StartTab 1
					}
				}
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
		ForEach ($Account in $ConfigFile.Config.WindowsSettings.HideAccounts.User) {
			# Write-Host ("`tHiding: " + $Account) -foregroundcolor "gray"
			Write-Color "Hiding: ",
						$Account -Color White,Gray -StartTab 1
			Set-Reg "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\SpecialAccounts\UserList" $Account 0 "DWORD"
		}
		If ($CreateUsers) {
			ForEach ( $i in $UserRange) {	
				If ($i) {
					# Write-Host ("`tHiding: " + ($ConfigFile.Config.Company.UserBaseName + $i)) -foregroundcolor "gray"
					Write-Color "Hiding: ",
						($ConfigFile.Config.Company.UserBaseName + $i) -Color White,Gray -StartTab 1
					Set-Reg "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\SpecialAccounts\UserList" ($ConfigFile.Config.Company.UserBaseName + $i) 0 "DWORD"
				}
			}
		}
		#endregion Hiding Accounts		
	}
	#endregion Windows 10 Only

	#region Registry Setup
		#Update/Add Items Values
		write-host ("Updating Computer Registry Settings:")
		Foreach ($key in ($ConfigFile.Config.WindowsSettings.ComputerRegistry.Item | Where-Object {$_.MinimumVersion -ge ([environment]::OSVersion.Version.Major)})) {
			If ($LockedDown -and $key.LockedDown -eq 'true' -and $key.Store -eq 'false') {					
				Write-Color -Text "LockedDown:  ",
									$key.Comment -Color Blue,DarkGray -StartTab 1
				Set-Reg ("HKLM:\" + $key.Key) $key.Value $key.Data $key.Type
			} ElseIf ($Store -and $key.Store -eq 'true') {				
				Write-Color -Text "Store:       ",
									$key.Comment -Color Red,DarkGray -StartTab 1
				Set-Reg ("HKLM:\" + $key.Key) $key.Value $key.Data $key.Type
			} ElseIf ($Manager -and $key.Manager -eq 'true') {				
				Write-Color -Text "Manager:     ",
									$key.Comment -Color Yellow,DarkGray -StartTab 1
				Set-Reg ("HKLM:\" + $key.Key) $key.Value $key.Data $key.Type
			} ElseIf ( $key.Store -eq 'false'-and $key.LockedDown -eq 'false' -and $key.Manager -eq 'false') {
				Write-Color -Text "All:         ",
							$key.Comment -Color DarkGreen,DarkGray -StartTab 1
				Set-Reg ("HKLM:\" + $key.Key) $key.Value $key.Data $key.Type	
			}
		}			
		If ($IsVM) {
			Foreach ($key in ($ConfigFile.Config.WindowsSettings.VM.ComputerRegistry.Item| Where-Object { $_.MinimumVersion -ge ([environment]::OSVersion.Version.Major)})) {
				If ($LockedDown -and $key.LockedDown -eq 'true' -and $key.Store -eq 'false') {				
					Write-Color -Text "LockedDown:  ",
										$key.Comment -Color Blue,DarkGray -StartTab 1
					Set-Reg ("HKLM:\" + $key.Key) $key.Value $key.Data $key.Type
				} ElseIf ($Store -and $key.Store -eq 'true') {				
					Write-Color -Text "Store:       ",
										$key.Comment -Color Red,DarkGray -StartTab 1
					Set-Reg ("HKLM:\" + $key.Key) $key.Value $key.Data $key.Type
				} ElseIf ($Manager -and $key.Manager -eq 'true') {				
					Write-Color -Text "Manager:     ",
										$key.Comment -Color Yellow,DarkGray -StartTab 1
					Set-Reg ("HKLM:\" + $key.Key) $key.Value $key.Data $key.Type
				} ElseIf ( $key.Store -eq 'false'-and $key.LockedDown -eq 'false' -and $key.Manager -eq 'false') {
					Write-Color -Text "All:         ",
								$key.Comment -Color DarkGreen,DarkGray -StartTab 1
					Set-Reg ("HKLM:\" + $key.Key) $key.Value $key.Data $key.Type	
				}
			}
		}
		#Remove Items
		Foreach ($key in ($ConfigFile.Config.WindowsSettings.ComputerRegistry.Remove | Where-Object {$_.MinimumVersion -ge ([environment]::OSVersion.Version.Major)})) {
			If ($LockedDown -and $key.LockedDown -eq 'true' -and $key.Store -eq 'false') {				
				Write-Color -Text "LockedDown:  ",
									$key.Comment -Color Blue,DarkGray -StartTab 1
				If ($key.Value) {
					Remove-ItemProperty -Path ("HKLM:\" + $key.Key) -Name $key.Value -Confirm:$False  -erroraction 'silentlycontinue'
				} else {
					Remove-Item -Path ("HKLM:\" + $key.Key)  -Confirm:$False  -erroraction 'silentlycontinue'
				}
			} ElseIf ($Store -and $key.Store -eq 'true') {				
				Write-Color -Text "Store:       ",
									$key.Comment -Color Red,DarkGray -StartTab 1
				If ($key.Value) {
					Remove-ItemProperty -Path ("HKLM:\" + $key.Key) -Name $key.Value -Confirm:$False  -erroraction 'silentlycontinue'
				} else {
					Remove-Item -Path ("HKLM:\" + $key.Key)  -Confirm:$False  -erroraction 'silentlycontinue'
				}
			} ElseIf ($Manager -and $key.Manager -eq 'true') {				
				Write-Color -Text "Manager:     ",
									$key.Comment -Color Yellow,DarkGray -StartTab 1
				If ($key.Value) {
					Remove-ItemProperty -Path ("HKLM:\" + $key.Key) -Name $key.Value -Confirm:$False  -erroraction 'silentlycontinue'
				} else {
					Remove-Item -Path ("HKLM:\" + $key.Key)  -Confirm:$False  -erroraction 'silentlycontinue'
				}
			} ElseIf ( $key.Store -eq 'false'-and $key.LockedDown -eq 'false' -and $key.Manager -eq 'false') {
				Write-Color -Text "All:         ",
							$key.Comment -Color DarkGreen,DarkGray -StartTab 1
				If ($key.Value) {
					Remove-ItemProperty -Path ("HKLM:\" + $key.Key) -Name $key.Value -Confirm:$False  -erroraction 'silentlycontinue'
				} else {
					Remove-Item -Path ("HKLM:\" + $key.Key)  -Confirm:$False  -erroraction 'silentlycontinue'
				}	
			}
		}
		#Add
		Foreach ($key in ($ConfigFile.Config.UserSettings.UserRegistry.Add | Where-Object {$_.MinimumVersion -ge ([environment]::OSVersion.Version.Major)})) {
			If ($LockedDown -and $key.LockedDown -eq 'true' -and $key.Store -eq 'false') {				
				Write-Color -Text "LockedDown:  ",
									$key.Comment -Color Blue,DarkGray -StartTab 1
				If ($key.key -and -Not (Test-Path ("HKLM:\" + $key.Key))) {
					New-Item -Path ("HKLM:\" + $key.Key)  -Force | Out-Null
				}
			} ElseIf ($Store -and $key.Store -eq 'true') {				
				Write-Color -Text "Store:       ",
									$key.Comment -Color Red,DarkGray -StartTab 1
				If ($key.key -and -Not (Test-Path ("HKLM:\" + $key.Key) )) {
					New-Item -Path ("HKLM:\" + $key.Key)  -Force | Out-Null
				}
			} ElseIf ($Manager -and $key.Manager -eq 'true') {				
				Write-Color -Text "Manager:     ",
									$key.Comment -Color Yellow,DarkGray -StartTab 1
				If ($key.key -and -Not (Test-Path ("HKLM:\" + $key.Key) )) {
					New-Item -Path ("HKLM:\" + $key.Key)  -Force | Out-Null
				}
			} ElseIf ( $key.Store -eq 'false'-and $key.LockedDown -eq 'false' -and $key.Manager -eq 'false') {
				Write-Color -Text "All:         ",
							$key.Comment -Color DarkGreen,DarkGray -StartTab 1
				If ($key.key -and -Not (Test-Path ("HKLM:\" + $key.Key) )) {
					New-Item -Path ("HKLM:\" + $key.Key)  -Force | Out-Null
				}
			}
		}
	#endregion Registry Setup
	#region Remove from This PC
	write-host ("Remove Items from This PC")
	Foreach ($Item in $ConfigFile.Config.WindowsSettings.RemoveFromThisPC.item) {
		# If(!(Test-Path ("HKCR:\CLSID\" + $Item))) {
		# 	New-Item -Path ("HKCR:\CLSID\" + $Item) -Force | Out-Null
		# }
		If (Test-Path ("HKCR:\CLSID\" + $Item)) {
			Set-KeyOwnership "HKCR:\" ("CLSID\" + $Item)
			Set-Reg ("HKCR:\CLSID\" + $Item) "System.IsPinnedToNameSpaceTree"  0 "DWORD"
			If ([Environment]::Is64BitOperatingSystem) {
				# If(!(Test-Path ("HKCR:\WOW6432Node\CLSID\" + $Item))) {
				# 	New-Item -Path ("HKCR:\WOW6432Node\CLSID\" + $Item) -Force | Out-Null
				# }
				If (Test-Path ("HKCR:\WOW6432Node\CLSID\" + $Item)) {
					Set-KeyOwnership "HKCR:\" ("WOW6432Node\CLSID\" + $Item)
					Set-Reg ("HKCR:\WOW6432Node\CLSID\" + $Item) "System.IsPinnedToNameSpaceTree"  0 "DWORD"
				}
			}
		}		
	}
	#endregion Remove from This PC
	#region VM
	If ($IsVM) {
		If ($ConfigFile.Config.WindowsSettings.VM.DisableDiskTimeOut) {
			Write-Host "Disabling Hard Disk Timeouts..." -ForegroundColor Yellow
			POWERCFG /SETACVALUEINDEX 381b4222-f694-41f0-9685-ff5bb260df2e 0012ee47-9041-4b5d-9b77-535fba8b1442 6738e2c4-e8a5-4a42-b16a-e040e769756e ($ConfigFile.Config.WindowsSettings.VM.DisableDiskTimeOut)
			POWERCFG /SETDCVALUEINDEX 381b4222-f694-41f0-9685-ff5bb260df2e 0012ee47-9041-4b5d-9b77-535fba8b1442 6738e2c4-e8a5-4a42-b16a-e040e769756e ($ConfigFile.Config.WindowsSettings.VM.DisableDiskTimeOut)
		}
		If ($ConfigFile.Config.WindowsSettings.VM.ServiceStartupTimeout) {
			Write-Host ("Increasing Service Startup Timeout To " + $ConfigFile.Config.WindowsSettings.VM.ServiceStartupTimeout + " Seconds.") -ForegroundColor Yellow
			Try {
				Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control' -Name 'ServicesPipeTimeout' -Value ([int]$ConfigFile.Config.WindowsSettings.VM.ServiceStartupTimeout * 1000)
			}
			Catch {
				Write-Warning "Could Not Set Service Startup Timeout"
			}
		}
		If ($ConfigFile.Config.WindowsSettings.VM.DiskTimeOutValue) {
			Write-Host ("Increasing Disk I/O Timeout " + $ConfigFile.Config.WindowsSettings.VM.DiskTimeOutValue + " Seconds.") -ForegroundColor Yellow
			Try {
				Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control' -Name 'ServicesPipeTimeout' -Value ([int]$ConfigFile.Config.WindowsSettings.VM.DiskTimeOutValue)
			}
			Catch {
				Write-Warning "Could Not Set Increasing Disk I/O Timeout"
			}
		}
		If ($ConfigFile.Config.WindowsSettings.VM.DisablingHibernate -eq 'true' -or $ConfigFile.Config.WindowsSettings.VM.DisablingHibernate -eq 'yes') {
			Write-Host "Disabling Hibernate..." -ForegroundColor Green
			POWERCFG -h off
		}
		If ($ConfigFile.Config.WindowsSettings.VM.DisableSystemRestore -eq 'true' -or $ConfigFile.Config.WindowsSettings.VM.DisableSystemRestore -eq 'yes') {
			Write-Host "Disabling System Restore..." -ForegroundColor Green
			Disable-ComputerRestore -Drive "C:\"
		}
		If ($ConfigFile.Config.WindowsSettings.VM.NewNetworkWindowOff -eq 'true' -or $ConfigFile.Config.WindowsSettings.VM.NewNetworkWindowOff -eq 'yes') {
			If (-Not (Test-path -path 'HKLM:\SYSTEM\CurrentControlSet\Control\Network\NewNetworkWindowOff')) {			
				Write-Host "Disabling New Network Dialog..." -ForegroundColor Green
				New-Item -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Network' -Name 'NewNetworkWindowOff' | Out-Null
			}
		}
	}	
	#endregion VM
	#region Powerbutton
	If ($ConfigFile.Config.WindowsSettings.PowerButtionAction) {
		Switch ($ConfigFile.Config.WindowsSettings.PowerButtionAction) {
			0 {Write-Host 'Setting "Power Button" to "Do nothing"...' -ForegroundColor Green}
			0 {Write-Host 'Setting "Power Button" to "Do nothing"...' -ForegroundColor Green}
			1 {Write-Host 'Setting "Power Button" to "Sleep"...' -ForegroundColor Green}
			2 {Write-Host 'Setting "Power Button" to "Hibernate"...' -ForegroundColor Green}
			3 {Write-Host 'Setting "Power Button" to "Shut down"...' -ForegroundColor Green}
			4 {Write-Host 'Setting "Power Button" to "Turn off the display"...' -ForegroundColor Green}
		}
		powercfg /SETDCVALUEINDEX SCHEME_CURRENT 4f971e89-eebd-4455-a8de-9e59040e7347 7648efa3-dd9c-4e3e-b566-50f929386280 ($ConfigFile.Config.WindowsSettings.PowerButtionAction)
		powercfg /SETDCVALUEINDEX SCHEME_CURRENT 4f971e89-eebd-4455-a8de-9e59040e7347 7648efa3-dd9c-4e3e-b566-50f929386280 ($ConfigFile.Config.WindowsSettings.PowerButtionAction)
		powercfg -SetActive SCHEME_CURRENT	
	}
	#endregion Powerbutton
	#regon Install Font
	Foreach ($key in ($ConfigFile.Config.WindowsSettings.Fonts.Path)) {
		If (Test-Path -Path (Get-envValueFromString -Path $key.'#text')) {
			If ($key.Recurse -match "true") {
				Install-Font -FontPath (Get-envValueFromString -Path $key.'#text') -Recurse
			}else {
				Install-Font -FontPath (Get-envValueFromString -Path $key.'#text')
			}
		}
	}
	#endregon Install Font
}
#============================================================================
#endregion Main Local Machine
#============================================================================
#============================================================================
#region Main Local Machine Adobe
#============================================================================
If (-Not $UserOnly) {
	Write-Host "Setting up Adobe Policies"
	ForEach ( $CARV in $ConfigFile.Config.AdobeReader.Version ) {
		ForEach ( $item in $ConfigFile.Config.AdobeReader.Item ) {
			Set-Reg ("HKLM:\SOFTWARE\Policies\Adobe\Acrobat Reader\" + $CARV + "\" + $item.Key) $item.Value $item.Data $item.Type
			#Wow6432Node
			If ([Environment]::Is64BitOperatingSystem) {
				Set-Reg ("HKLM:\SOFTWARE\Policies\Adobe\Acrobat Reader\" + $CARV + "\" + $item.Key).replace("\Software\","\Software\Wow6432Node\") $item.Value $item.Data $item.Type			
			}
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
	Write-Host "Setting up Services: "
	# Source: https://github.com/W4RH4WK/Debloat-Windows-10/blob/master/scripts/disable-services.ps1
	#Services to Disable
	ForEach ($service in $configfile.Config.WindowsSettings.DisableServices.Service) {
		$ServiceObject = Get-Service -Name $service -erroraction 'silentlycontinue'
		If ($ServiceObject) {
			#Windows 10 and 2016 have hidden servcie bowser which will disalbe all SMB traffic if disabled.
			If (($ServiceObject).Name -ne "bowser") {
				# write-host ("`tDisabling: " + $ServiceObject.DisplayName ) -foregroundcolor green 
				Write-Color -Text "Disabling: ",
									$ServiceObject.DisplayName -Color White,DarkGreen -StartTab 1
				$ServiceObject | Stop-Service -ErrorAction SilentlyContinue | Out-Null
				$ServiceObject | Set-Service -StartupType Disabled -ErrorAction SilentlyContinue | Out-Null
			}
		}
	}
	#Services to set as Manual
	ForEach ($service in $configfile.Config.WindowsSettings.ManualServices.Service) {
		$ServiceObject = Get-Service -Name $service -erroraction 'silentlycontinue'
		If ( $ServiceObject) {
			# write-host ("`tManual Startup: " + $ServiceObject.DisplayName ) -foregroundcolor yellow 
			Write-Color -Text "Manual Startup: ",
							  $ServiceObject.DisplayName -Color White,DarkCyan -StartTab 1
			$ServiceObject | Stop-Service -ErrorAction SilentlyContinue | Out-Null
			$ServiceObject | Set-Service -StartupType Manual -ErrorAction SilentlyContinue | Out-Null
		}
	}
	#Services to set as Automatic
	ForEach ($service in $configfile.Config.WindowsSettings.AutomaticServices.Service) {
		$ServiceObject = Get-Service -Name $service -erroraction 'silentlycontinue'
		If ( $ServiceObject) {
			# write-host ("`tAutomatic Startup: " + $ServiceObject.DisplayName ) -foregroundcolor red 
			Write-Color -Text "Automatic Startup: ",
							  $ServiceObject.DisplayName -Color White,Red -StartTab 1
			$ServiceObject | Set-Service -StartupType Automatic -ErrorAction SilentlyContinue | Out-Null
			$ServiceObject | Start-Service -ErrorAction SilentlyContinue | Out-Null
		}
	}	
}
If ($Wifi) {
	ForEach ($service in $configfile.Config.WindowsSettings.WiFiServices.Service) {
		$ServiceObject = Get-Service -Name $service -erroraction 'silentlycontinue'
		If ( $ServiceObject) {
			# write-host ("`tAutomatic Startup: " + $ServiceObject.DisplayName ) -foregroundcolor red 
			Write-Color -Text "Automatic Startup: ",
							  $ServiceObject.DisplayName,
							  " Force for WiFi" -Color White,Red,Green -StartTab 1
			$ServiceObject | Set-Service -StartupType Automatic -ErrorAction SilentlyContinue | Out-Null
			$ServiceObject | Start-Service -ErrorAction SilentlyContinue | Out-Null
		}
	}	
}
#============================================================================
#endregion Main Local Machine Services
#============================================================================
#============================================================================
#region Main Local Machine Certs
#============================================================================
If (-Not $UserOnly) {
	If ([int]([environment]::OSVersion.Version.Major.tostring() + [environment]::OSVersion.Version.Minor.tostring()) -gt 61) {
		If (Get-Command Import-Certificate -errorAction SilentlyContinue) {
			Write-Host ("Setting up Certificates:")
			If (Test-Path ($LICache + "\" + $ConfigFile.Config.Company.PrivateCARoot)) {
				# Write-Host ("Importing Domain CA Root: " + $LICache + "\" + $ConfigFile.Config.Company.PrivateCARoot)
				Write-color	-Text "Importing Domain CA Root: ",
									($LICache + "\" + $ConfigFile.Config.Company.PrivateCARoot) -Color White,DarkGreen
				Import-Certificate -Filepath ($LICache + "\" + $ConfigFile.Config.Company.PrivateCARoot) -CertStoreLocation cert:\LocalMachine\Root | out-null
			}
			If (Test-Path ($LICache + "\" + $ConfigFile.Config.Company.PrivateCAIntermediate)) {
				# Write-Host ("Importing Domain CA Intermediate : " + $LICache + "\" + $ConfigFile.Config.Company.PrivateCAIntermediate)
				Write-color	-Text "Importing Domain CA Intermediate: ",
						($LICache + "\" + $ConfigFile.Config.Company.PrivateCAIntermediate) -Color White,DarkGreen
				Import-Certificate -Filepath ($LICache + "\" + $ConfigFile.Config.Company.PrivateCAIntermediate) -CertStoreLocation cert:\LocalMachine\CA | out-null
			}
			#Importing Code Signing Cert
			If (Test-Path ( $LICache + "\" + $ConfigFile.Config.Company.PrivateCACodeSigning )) {
				# Write-Host ("Importing Code Signing Cert : " + $LICache + "\" + $ConfigFile.Config.Company.PrivateCACodeSigning)
				Write-color	-Text "Importing Code Signing Cert: ",
						($LICache + "\" + $ConfigFile.Config.Company.PrivateCACodeSigning) -Color White,DarkGreen
				Import-Certificate -Filepath ($LICache + "\" + $ConfigFile.Config.Company.PrivateCACodeSigning) -CertStoreLocation cert:\LocalMachine\TrustedPublisher | out-null
			}
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
	Write-Host "Setting up SSL : "
	#http://www.vistax64.com/powershell/21794-creating-registry-keys-createsubkey-method.html
	#Set Ciphers
	#Need to go old school to set registry as powershell cannot handle keys with "/" in them. 
	Foreach ($Cipher in $ConfigFile.Config.WindowsSettings.Schannel.Cipher) {
		If ($Cipher.Status -eq "Enable" -or $Cipher.Status -eq "Enabled" -or $Cipher.Status -eq "on") {
			# Write-Host ("`t Enabling Cipher: " + $Cipher.'#text') -foregroundcolor Yellow
			Write-Color -Text "Enabling Cipher: ",
								$Cipher.'#text' -Color White,DarkYellow -StartTab 1
			reg add $('"' + $RegAddSCHANNEL + '\Ciphers\' + ($Cipher.'#text') + '"') /v Enabled /d 4294967295 /t REG_DWORD /f
			#Set-Reg ("HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\" + $Cipher.'#text' ) "Enabled" 4294967295 "DWORD"
		} else {
			# Write-Host ("`t Disabling Cipher: " + $Cipher.'#text') -foregroundcolor Green
			Write-Color -Text "Disabling Cipher: ",
								$Cipher.'#text' -Color White,DarkGreen -StartTab 1
			reg add $('"' + $RegAddSCHANNEL + '\Ciphers\' + ($Cipher.'#text') + '"') /v Enabled /d 0 /t REG_DWORD /f
			#Set-Reg ("HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\" + $Cipher.'#text' ) "Enabled" 0 "DWORD"
		}
	}
	#Set Hashes
	Foreach ($Hashe in $ConfigFile.Config.WindowsSettings.Schannel.Hashe) {
		If ($Hashe.Status -eq "Enable" -or $Hashe.Status -eq "Enabled" -or $Hashe.Status -eq "on") {
			# Write-Host ("`t Enabling Hashe: " + $Hashe.'#text') -foregroundcolor Yellow
			Write-Color -Text "Enabling Hashe: ",
								$Hashe.'#text' -Color White,DarkYellow -StartTab 1
			Set-Reg ("HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Hashes\" + $Hashe.'#text' ) "Enabled" 4294967295 "DWORD"
		} else {
			# Write-Host ("`t Disabling Hashe: " + $Hashe.'#text') -foregroundcolor Green
			Write-Color -Text "Disabling Hashe: ",
								$Hashe.'#text' -Color White,DarkGreen -StartTab 1
			Set-Reg ("HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Hashes\" + $Hashe.'#text' ) "Enabled" 0 "DWORD"
		}
	}
	If ($AllowClientTLS1) {
		Set-Reg ("HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Hashes\SHA") "Enabled" 4294967295 "DWORD"
		Write-Warning "Re-Enabling SHA" 
	}
	# Set KeyExchangeAlgorithms
	Foreach ($KeyExchangeAlgorithm in $ConfigFile.Config.WindowsSettings.Schannel.KeyExchangeAlgorithm) {
		If ($KeyExchangeAlgorithm.Status -eq "Enable" -or $KeyExchangeAlgorithm.Status -eq "Enabled" -or $KeyExchangeAlgorithm.Status -eq "on") {
			# Write-Host ("`t Enabling KeyExchangeAlgorithm: " + $KeyExchangeAlgorithm.'#text') -foregroundcolor Yellow
			Write-Color -Text "Enabling KeyExchangeAlgorithm: ",
								$KeyExchangeAlgorithm.'#text' -Color White,DarkYellow -StartTab 1
			Set-Reg ("HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\KeyExchangeAlgorithms\" + $KeyExchangeAlgorithm.'#text' ) "Enabled" 4294967295 "DWORD"
		} else {
			# Write-Host ("`t Disabling KeyExchangeAlgorithm: " + $KeyExchangeAlgorithm.'#text') -foregroundcolor Green
			Write-Color -Text "Disabling KeyExchangeAlgorithm: ",
								$KeyExchangeAlgorithm.'#text' -Color White,DarkGreen -StartTab 1
			Set-Reg ("HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\KeyExchangeAlgorithms\" + $KeyExchangeAlgorithm.'#text' ) "Enabled" 0 "DWORD"
		}
		If ($KeyExchangeAlgorithm.'#text' -eq "Diffie-Hellman" -and $KeyExchangeAlgorithm.ServerMinKeyBitLength) {
			Set-Reg ("HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\KeyExchangeAlgorithms\Diffie-Hellman") "ServerMinKeyBitLength" $KeyExchangeAlgorithm.ServerMinKeyBitLength "DWORD"
		}
	}
	#Set Protocols
	Foreach ($Protocol in $ConfigFile.Config.WindowsSettings.Schannel.Protocol) {
		#Server
		If ($Protocol.Server -eq "Enable" -or $Protocol.Server -eq "Enabled" -or $Protocol.Server -eq "on") {
			# Write-Host ("`t Enabling Server Protocol: " + $Protocol.'#text') -foregroundcolor Yellow
			Write-Color -Text "Enabling ",
								"Server ",
								"Protocol: ",
								$Protocol.'#text' -Color White,DarkCyan,White,DarkYellow -StartTab 1
			Set-Reg ("HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\" + $Protocol.'#text' + "\Server") "Enabled" 4294967295 "DWORD"
			Set-Reg ("HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\" + $Protocol.'#text' + "\Server") "DisabledByDefault" 0 "DWORD"
		} else {
			# Write-Host ("`t Disabling Server Protocol: " + $Protocol.'#text') -foregroundcolor Green
			Write-Color -Text "Disabling ",
								"Server ",
								"Protocol: ",
								$Protocol.'#text' -Color White,DarkCyan,White,DarkGreen -StartTab 1
			Set-Reg ("HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\" + $Protocol.'#text'  + "\Server" ) "Enabled" 0 "DWORD"
			Set-Reg ("HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\" + $Protocol.'#text' + "\Server") "DisabledByDefault" 1 "DWORD"
		}
		#Client
		If ($Protocol.Client -eq "Enable" -or $Protocol.Client -eq "Enabled" -or $Protocol.Client -eq "on") {
			# Write-Host ("`t Enabling Protocol: " + $Protocol.'#text') -foregroundcolor Yellow
			Write-Color -Text "Enabling ",
								"Client ",
								"Protocol: ",
								$Protocol.'#text' -Color White,DarkGray,White,DarkYellow -StartTab 1
			Set-Reg ("HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\" + $Protocol.'#text'  + "\Client") "Enabled" 4294967295 "DWORD"
			Set-Reg ("HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\" + $Protocol.'#text' + "\Client") "DisabledByDefault" 0 "DWORD"
		} else {
			# Write-Host ("`t Disabling Protocol: " + $Protocol.'#text') -foregroundcolor Green
			Write-Color -Text "Disabling ",
								"Client ",
								"Protocol: ",
								$Protocol.'#text' -Color White,DarkGray,White,DarkGreen -StartTab 1
			Set-Reg ("HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\" + $Protocol.'#text'  + "\Client") "Enabled" 0 "DWORD"
			Set-Reg ("HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\" + $Protocol.'#text' + "\Client") "DisabledByDefault" 1 "DWORD"
		}
	}
	If ($AllowClientTLS1) {
		Set-Reg ("HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0\Client") "Enabled" 4294967295 "DWORD"
		Set-Reg ("HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0\Client") "DisabledByDefault" 0 "DWORD"
		Set-Reg ("HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Client") "Enabled" 4294967295 "DWORD"
		Set-Reg ("HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Client") "DisabledByDefault" 0 "DWORD"
		Write-Warning "Re-Enabling Client TLS 1.0" 
	}
	#.Net TLS Settings
	If ($ConfigFile.Config.WindowsSettings.Schannel.DotNetTLS12 -eq 'true' -and $ConfigFile.Config.WindowsSettings.Schannel.DotNetTLS12 -eq 'yes' -and $ConfigFile.Config.WindowsSettings.Schannel.DotNetTLS12 -eq 'on') {
		Write-Host "Setting up .Net for TLS 1.2"
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
	}
	If ($ConfigFile.Config.WindowsSettings.Schannel.WinHttp) {
		#https://support.microsoft.com/en-us/help/3140245/update-to-enable-tls-1-1-and-tls-1-2-as-a-default-secure-protocols-in
		Set-Reg ("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\WinHttp") "DefaultSecureProtocols" $ConfigFile.Config.WindowsSettings.Schannel.WinHttp "DWORD"
		If ([Environment]::Is64BitOperatingSystem) {
			Set-Reg ("HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Internet Settings\WinHttp") "DefaultSecureProtocols" $ConfigFile.Config.WindowsSettings.Schannel.WinHttp "DWORD"
		}
		Set-Reg ("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\WinHttp") "SecureProtocols" $ConfigFile.Config.WindowsSettings.Schannel.WinHttp "DWORD"
		If ([Environment]::Is64BitOperatingSystem) {
			Set-Reg ("HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Internet Settings\WinHttp") "SecureProtocols" $ConfigFile.Config.WindowsSettings.Schannel.WinHttp "DWORD"
		}
	}
}
#============================================================================
#endregion Main Local Machine Schannel for PCI
#============================================================================
#============================================================================
#region Main Local Machine User Icons
#============================================================================
If (-Not $UserOnly) {
	If ([int]([environment]::OSVersion.Version.Major.tostring() + [environment]::OSVersion.Version.Minor.tostring()) -gt 61 -and $ConfigFile.Config.Company.UserPictures) {
		Write-Host "Setting up User Icons: "
		If (Test-Path ($LICache + "\" + $ConfigFile.Config.Company.UserPictures)) {
			copy-item ($LICache + "\" + $ConfigFile.Config.Company.UserPictures + "\*.*") -Destination ($env:programdata + "\Microsoft\User Account Pictures") -force
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
	Write-Host "Setting up Background: "
	#Set Default Picture
	Set-Owner -Path ($env:windir + "\Web\Wallpaper\Windows\img0.jpg")
	#Add Administrators with full control
	$Folderpath=($env:windir + "\Web\Wallpaper\Windows\img0.jpg")
	$user_account='Administrators'
	$Acl = Get-Acl $Folderpath
	$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule($user_account,"FullControl", "None", "None", "Allow")
	$Acl.Setaccessrule($Ar)
	Set-Acl $Folderpath $Acl
	If (-Not $BackgroundFolder -and $ConfigFile.Config.Company.BackgroundFolder) {
		$BackgroundFolder = $ConfigFile.Config.Company.BackgroundFolder
	}

	If (Test-Path ($LICache + "\" + $ConfigFile.Config.Company.WallPaperFolder + "\" + $BackgroundFolder + "\img0.jpg")) {	
		copy-item ($LICache + "\" + $ConfigFile.Config.Company.WallPaperFolder + "\" + $BackgroundFolder + "\img0.jpg") -Destination ($env:windir + "\Web\Wallpaper\Windows\img0.jpg") -force | out-null
	} else {
		Write-Warning ("Please make sure the following file exists: " +  ($LICache + "\" + $ConfigFile.Config.Company.WallPaperFolder + "\" + $BackgroundFolder + "\img0.jpg") )
	}
	If (Test-Path ($LICache + "\" + $ConfigFile.Config.Company.WallPaperFolder + "\" + $BackgroundFolder + "\Backgrounds")) {	
		If (-Not( Test-Path ($env:windir + "\System32\oobe\info\backgrounds\"))) {
			New-Item -ItemType directory -Path ($env:windir + "\system32\oobe\info\backgrounds") | out-null
		}
		copy-item ($LICache + "\" + $ConfigFile.Config.Company.WallPaperFolder + "\" + $BackgroundFolder + "\Backgrounds\*.*") -Destination ($env:windir + "\System32\oobe\info\backgrounds\") -force | out-null
	} else {
		Write-Warning ("Please make sure the following folder exists: " +  ($LICache + "\" + $ConfigFile.Config.Company.WallPaperFolder + "\" + $BackgroundFolder + "\Backgrounds") )
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
		If (Test-Path ($LICache + "\" + $ConfigFile.Config.Company.WallPaperFolder + "\" + $BackgroundFolder + "\4K\Wallpaper\Windows")) {	
			copy-item ($LICache + "\" + $ConfigFile.Config.Company.WallPaperFolder + "\" + $BackgroundFolder + "\4K\Wallpaper\Windows\*.*") -Destination ($env:windir + "\Web\4K\Wallpaper\Windows") -force
		} else {
			Write-Warning ("Please make sure the following folder exists: " +  ($LICache + "\" + $ConfigFile.Config.Company.WallPaperFolder + "\" + $BackgroundFolder + "\4K\Wallpaper\Windows") )
		}
	}
}
#============================================================================
#endregion Main Local Machine Background
#============================================================================
#============================================================================
#region Main Local Machine Setup Windows Time
#============================================================================
If (-Not $UserOnly -and $ConfigFile.config.WindowsSettings.NtpServer) {
	Write-Host "Setting up Time: "
	#Disable Clients being NTP Servers
	Set-Reg "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\TimeProviders\NtpServer" "Enabled" 0 "DWORD"
	If ($Store) {
		net stop w32time | out-null
		W32tm /config /syncfromflags:manual /manualpeerlist:($ConfigFile.config.WindowsSettings.NtpServer) | out-null
		w32tm /config /reliable:yes | out-null
		net start w32time | out-null
		w32tm /resync /rediscover | out-null
	} else {
		net stop w32time | out-null
		W32tm /config /syncfromflags:ALL /manualpeerlist:($ConfigFile.config.WindowsSettings.NtpServer) | out-null
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
		Write-Host "Setting up BGInfo: "
		If (Test-Path ($LICache + "\BgInfo")) {
			copy-item ($LICache + "\BgInfo") -Destination ($env:programfiles) -Force -Recurse
			Get-ChildItem ($env:programdata + "\Microsoft\Windows\Start Menu\Programs\StartUp") | Where-Object Name -Like "*bginfo*.lnk" | ForEach-Object { Remove-Item $_.fullname}
			If ($Store -or $IsVM) {
				copy-item ($env:programfiles + "\BgInfo\" + ($ConfigFile.Config.BGInfo.Store)) ($env:programdata + "\Microsoft\Windows\Start Menu\Programs\StartUp") -Force
			}else{
				copy-item ($env:programfiles + "\BgInfo\" + ($ConfigFile.Config.BGInfo.Default)) ($env:programdata + "\Microsoft\Windows\Start Menu\Programs\StartUp") -Force
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
If (-Not $UserOnly) {
	Write-Host "Setting up Firewall: "
	#region Custom Software Firewall
	If ([int]([environment]::OSVersion.Version.Major.tostring() + [environment]::OSVersion.Version.Minor.tostring()) -gt 61) {
		$ProgressPreference = "SilentlyContinue"
		Remove-NetFirewallRule -Group (split-path $ConfigFile.Config.Company.SoftwarePath -Leaf ) -erroraction 'silentlycontinue'
		$ProgressPreference = "Continue"
		If (Test-Path $ConfigFile.Config.Company.SoftwarePath) {
			# Write-Host ("`tAdding " + (split-path $ConfigFile.Config.Company.SoftwarePath -Leaf ) + " to Firewall...") -foregroundcolor darkgray
			Write-Color -Text "Adding: ",
							(split-path $ConfigFile.Config.Company.SoftwarePath -Leaf ),
							" to Firewall..." -Color White,DarkGray,White -StartTab 2	
			Get-ChildItem -Path $ConfigFile.Config.Company.SoftwarePath -Filter *.exe -Recurse| ForEach-Object {
				# Write-Host ("`t`t Adding rule for: " + $_.Name) -foregroundcolor yellow
				Write-Color -Text "Adding Rule for: ",
									$_.Name -Color White,Yellow -StartTab 2
				New-NetFirewallRule -DisplayName $_.Name -Direction Inbound -Program $_.VersionInfo.FileName -Group (split-path $ConfigFile.Config.Company.SoftwarePath -Leaf ) -Action Allow | out-null
			}
		}
	}
	#endregion Custom Software Firewall
	#region DisplayGroup
	If ([environment]::OSVersion.Version.Major -ge 10) {
		Write-Host "`tDisabling un-needed Firewall Rules: " -foregroundcolor darkgray
		If ($ConfigFile.Config.WindowsSettings.Firewall.DisplayGroup | Where-object {$_.action -eq "disable"}) {
			Foreach ($DisplayGroup in ($ConfigFile.Config.WindowsSettings.Firewall.DisplayGroup | Where-object {$_.action -eq "disable"}).'#text') {
				# Write-host ("`t`tDisabling DisplayGroup: " + $DisplayGroup)
				Write-Color -Text "Disabling DisplayGroup: ",
									$DisplayGroup -Color White,DarkGreen -StartTab 2
				$ProgressPreference = "SilentlyContinue"
				Disable-NetFirewallRule -DisplayGroup $DisplayGroup -erroraction 'silentlycontinue' | out-null
				$ProgressPreference = "Continue"
			}
		}
		If ($ConfigFile.Config.WindowsSettings.Firewall.DisplayGroup | Where-object {$_.action -eq "enable"}) {
			Foreach ($DisplayGroup in ($ConfigFile.Config.WindowsSettings.Firewall.DisplayGroup | Where-object {$_.action -eq "enable"})) {
				# Write-host ("`t`enable DisplayGroup: " + $DisplayGroup)
				Write-Color -Text "Enabling DisplayGroup: ",
								($DisplayGroup.'#text') -Color White,Yellow -StartTab 2
				$ProgressPreference = "SilentlyContinue"
				Enable-NetFirewallRule -DisplayGroup $DisplayGroup.'#text' -erroraction 'silentlycontinue' | out-null
				If ($DisplayGroup.AddressFilter) {
					Write-Color -Text "Alowing Remote Addresses: ",
										$DisplayGroup.AddressFilter -Color White,Red -StartTab 3
					Get-NetFirewallRule | Where-Object {$_.DisplayGroup -match ($DisplayGroup.'#text') } | Get-NetFirewallAddressFilter | Where-Object { $_.RemoteAddress -ne $DisplayGroup.AddressFilter} | Set-NetFirewallAddressFilter -RemoteAddress $DisplayGroup.AddressFilter
				}
				$ProgressPreference = "Continue"
			}
		}		
	}
	#endregion DisplayGroup
	#region New Rule
	Write-Host "`tAdding Firewall Rules: " -foregroundcolor darkgray
	If ($ConfigFile.Config.WindowsSettings.Firewall.Rule) {
		Foreach ($Rule in $ConfigFile.Config.WindowsSettings.Firewall.Rule) {
			# Write-Host ("`t`tAdding rule for: " + $Rule.DisplayName)
			Write-Color -Text "Adding rule for: ",
							$Rule.DisplayName -Color White,DarkGreen -StartTab 2
			If ($Rule.Protocol -match "ICMP") {
				New-NetFirewallRule -DisplayName $Rule.DisplayName -Direction $Rule.Direction -Protocol $Rule.Protocol -IcmpType $Rule.IcmpType  -Action $Rule.Action | out-null
			} else {
				If ($Rule.LocalPort){
					New-NetFirewallRule -DisplayName $Rule.DisplayName -Direction $Rule.Direction -Protocol $Rule.Protocol -LocalPort $Rule.LocalPort  -Action $Rule.Action | out-null
				}ElseIf ($Rule.RemotePort) {
					New-NetFirewallRule -DisplayName $Rule.DisplayName -Direction $Rule.Direction -Protocol $Rule.Protocol -LocalPort $Rule.LocalPort  -Action $Rule.Action | out-null
				}
			}
		}
	}
	#endregion New Rule
}
#============================================================================
#endregion Main Local Machine Firewall Setup
#============================================================================
#============================================================================
#region Main Local Machine All Users Desktop
#============================================================================
If (-Not $UserOnly) {
	If ($ConfigFile.Config.IE.AddIEtoAllUserDesktop -eq "true" -or $ConfigFile.Config.IE.AddIEtoAllUserDesktop -eq "yes") {
		If (-Not (Test-Path ($env:Public + "\Desktop\Internet Explorer.lnk"))) {
			If ( Test-Path ($env:appdata + "\Microsoft\Windows\Start Menu\Programs\Accessories\Internet Explorer.lnk")) {
				Write-Host "Adding Internet Explorer to All Users Desktop"
				copy-item ($env:appdata + "\Microsoft\Windows\Start Menu\Programs\Accessories\Internet Explorer.lnk") ($env:Public + "\Desktop\Internet Explorer.lnk")
			}
		}
	}
	#Add other Icons to all users desktop.
	If (Test-Path($LICache + "\" + $ConfigFile.Config.Company.CopyAllUserDesktopFolder)) {
		Copy-Item -Force -Recurse -Path ($LICache + "\" + $ConfigFile.Config.Company.CopyAllUserDesktopFolder) -Destination ($env:Public)
	}
	#Copy Custom Icons.
	If (Test-Path($LICache + "\" +  $ConfigFile.Config.Company.CopyIconsLocalFolder)) {
		If (-Not (Test-path($ConfigFile.Config.Company.IconPath))) {
			New-Item -ItemType Directory -Force -Path $ConfigFile.Config.Company.IconPath
		}
		Copy-Item -Force -Recurse -Path ($LICache + "\" +  $ConfigFile.Config.Company.CopyIconsLocalFolder + "\*") -Destination ($ConfigFile.Config.Company.IconPath)
	}
}
#============================================================================
#endregion Main Local Machine All Users Desktop
#============================================================================
#============================================================================
#region Main Local Machine RDP
#============================================================================
If (-Not $UserOnly) {
	If ($ConfigFile.Config.WindowsSettings.EnableRDP -eq "true" -or $ConfigFile.Config.WindowsSettings.EnableRDP -eq "yes") {
		Write-Host "Enabling RDP"
		Set-Reg "HKLM:\SYSTEM\CurrentControlSet\control\Terminal Server" "fDenyTSConnections " 0 "DWORD"
	} else {
		Write-Host "Disabling RDP"
		Set-Reg "HKLM:\SYSTEM\CurrentControlSet\control\Terminal Server" "fDenyTSConnections " 1 "DWORD"
	}
}
#============================================================================
#endregion Main Local Machine RDP
#============================================================================
#============================================================================
#region Main Local Machine Setup Screen Saver
#============================================================================
If (-Not $UserOnly) {
	If ($ConfigFile.Config.WindowsSettings.ScreenSave.Active -and $ConfigFile.Config.WindowsSettings.ScreenSave.Secure -and $ConfigFile.Config.WindowsSettings.ScreenSave.TimeOut -and $ConfigFile.Config.WindowsSettings.ScreenSave.ScreenSaver) {
		Write-Host "Setup Logon Screen Saver:"
		Set-Reg "HKU:\.DEFAULT\Control Panel\Desktop" "ScreenSaveActive" $ConfigFile.Config.WindowsSettings.ScreenSave.Active "STRING"
		Set-Reg "HKU:\.DEFAULT\Control Panel\Desktop" "ScreenSaverIsSecure" $ConfigFile.Config.WindowsSettings.ScreenSave.Secure "STRING"
		Set-Reg "HKU:\.DEFAULT\Control Panel\Desktop" "ScreenSaveTimeOut" $ConfigFile.Config.WindowsSettings.ScreenSave.TimeOut "STRING"
		Set-Reg "HKU:\.DEFAULT\Control Panel\Desktop" "SCRNSAVE.EXE" $ConfigFile.Config.WindowsSettings.ScreenSave.ScreenSaver "STRING"
	}
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
		Write-Host "Remove Microsoft Store Apps:"
		#region Remove Appx Packages
		If (Get-Command Get-AppxPackage -errorAction SilentlyContinue) {
			$AllInstalled = Get-AppxPackage -AllUsers | Where-Object {$_.NonRemovable -ne $True} | ForEach-Object {$_.Name}		
			#Turn off the progress bar
			$ProgressPreference = 'silentlyContinue'
			[array]$WhiteList = $ConfigFile.Config.WindowsSettings.MicrosoftStore.WhiteList
			ForEach($Appx in $AllInstalled){
				$error.Clear()
				If (-Not ([string]::IsNullOrEmpty($Appx))) {
					If ($Appx -match '(\{|\()?[A-Za-z0-9]{4}([A-Za-z0-9]{4}\-?){4}[A-Za-z0-9]{12}(\}|\()?' ) {
						$AppxClean = $Appx
					} else {
						$AppxClean = [String]($Appx -replace '\d+' -replace '\.\.')
					}
					If (-Not ($WhiteList.Contains($AppxClean))){
						Try{			
							Get-AppxPackage -Name $Appx | Remove-AppxPackage
						}
						Catch{
							$ErrorMessage = $_.Exception.Message
							$FailedItem = $_.Exception.ItemName
							Write-Host "There was an error removing Appx: $Appx"
							Write-Host $ErrorMessage
							Write-Host $FailedItem
						}
						If(!$error){
							# Write-Host "Removed Appx: $Appx" -ForegroundColor Green
							Write-Color -Text "Removed App: ",
										$Appx -Color White,DarkGreen -StartTab 1
						}
					}
					Else{
						# Write-Host "Appx Package is whitelisted: $Appx" -ForegroundColor DarkBlue
						Write-Color -Text "Whitelisted App: ",
										$Appx -Color White,DarkCyan -StartTab 1
					}
				}
			}
			#Turn on the progress bar
			$ProgressPreference = 'Continue'		
		}
		#endregion
		#region Remove Provisioned Appx Packages
		If (Get-Command Get-ProvisionedAppxPackage -errorAction SilentlyContinue) {
			$AllProvisioned = Get-ProvisionedAppxPackage -Online | Where-Object {$_.NonRemovable -ne $True}| ForEach-Object {$_.DisplayName}
			ForEach($Appx in $AllProvisioned){
				$error.Clear()
				If(-Not ([array]$ConfigFile.Config.WindowsSettings.MicrosoftStore.WhiteList).Contains([system.String]::Join(".", ($Appx.split(".") |  ForEach-Object {if (($_ -as [int] -eq $null )) {$_ }})))){
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
						# Write-Host "Removed Provisioned Appx: $Appx" -ForegroundColor Green
						Write-Color -Text "Removed App: ",
										$Appx -Color White,DarkGreen -StartTab 1
					}
				}
				Else{
					#Write-Host "Appx Package is whitelisted: $Appx" -ForegroundColor DarkBlue
					Write-Color -Text "Whitelisted App: ",
										$Appx -Color White,DarkCyan -StartTab 1
				}
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
If (-Not $UserOnly -and ($ConfigFile.Config.WindowsSettings.RemoveOneDrive -eq "true" -or $ConfigFile.Config.WindowsSettings.RemoveOneDrive -eq "yes")) {
	$process = Start-Process -FilePath "taskkill" -ArgumentList @("/f","/im","OneDrive.exe")
	#https://social.technet.microsoft.com/Forums/ie/en-US/2eaa1b6a-c906-4161-b76c-370ac8910a11/windows-10-sysprep-issue-image-always-hangs-at-quotgetting-readyquot?forum=win10itprosetup
	If (Test-Path ($env:systemroot + "\SysWOW64\OneDriveSetup.exe")) {
		Write-Host "Removing OneDrive:" -foregroundcolor darkgray
		$process = Start-Process -FilePath ('"'+ $env:systemroot + "\SysWOW64\OneDriveSetup.exe" + '"') -ArgumentList @("/uninstall","/quiet") -PassThru -NoNewWindow -Wait
	}
	If (Test-Path ($env:systemroot + "\System32\OneDriveSetup.exe")) {
		Write-Host "Removing OneDrive:" -foregroundcolor darkgray
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
		Write-Host "Setup System OEM Info:"
		If (-Not $IsVM) {
			Set-Reg "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation" "Manufacturer" ((Get-CimInstance -ClassName Win32_ComputerSystem).Manufacturer) "String"
			If ($OEMInfoAddSerial -or $ConfigFile.Config.Company.OEMInfoAddSerial -eq "true" -or $ConfigFile.Config.Company.OEMInfoAddSerial -eq "yes") {
				Set-Reg "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation" "Model" ((Get-CimInstance -ClassName Win32_ComputerSystem).model + " (Serial Number: " + (Get-CimInstance -ClassName Win32_BIOS).SerialNumber + ")") "String"
			}else{
				Set-Reg "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation" "Model" ((Get-CimInstance -ClassName Win32_ComputerSystem).model) "String"
			}
		}
		If (-Not (Test-Path ($env:windir + "\system32\oobe\info\"))) {
			New-Item -ItemType directory -Path ($env:windir + "\system32\oobe\info\") | out-null
		}
		Copy-Item  ($LICache + "\" + $ConfigFile.Config.Company.WallPaperFolder + "\" + $ConfigFile.Config.Company.OEMLogo) -Destination ($env:windir + "\system32\oobe\info\" + $ConfigFile.Config.Company.OEMLogo ) -Recurse -Force
		Set-Reg "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation" "Logo" ($env:windir + "\system32\oobe\info\" + $ConfigFile.Config.Company.OEMLogo ) "String"
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
	Write-Host ("Disabling Netbios") -foregroundcolor darkgray
	$key = "HKLM:SYSTEM\CurrentControlSet\services\NetBT\Parameters\Interfaces"
	Get-ChildItem $key |
	ForEach-Object { Set-ItemProperty -Path "$key\$($_.pschildname)" -Name NetbiosOptions -Value 2 }
	If (-Not $IPv6) {
		Write-Host ("Disabling IPv6") -foregroundcolor darkgray
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
If ($Store -and -Not $UserOnly) {
	Write-Host ("Setting up SNMP") -foregroundcolor darkgray
	If (-Not (Get-Service -Name "SNMP" -ErrorAction SilentlyContinue)) {
		If (Get-Command Get-WindowsOptionalFeature -errorAction SilentlyContinue) {
			If (Get-WindowsOptionalFeature -Online -FeatureName "SNMP" -ErrorAction SilentlyContinue) {
				Write-Host ("`tInstalling up SNMP") -foregroundcolor darkgray
				Enable-WindowsOptionalFeature -online -FeatureName "SNMP" -NoRestart | Out-Null
				Get-Service -Name "SNMP" -ErrorAction SilentlyContinue | Set-Service -StartupType Disabled 
			}
		}
		If (Get-Command Get-WindowsCapability -errorAction SilentlyContinue) {
			If ((Get-WindowsCapability -Online -Name "SNMP.Client*" -ErrorAction SilentlyContinue).name ) {
				Write-Host ("`tInstalling up SNMP") -foregroundcolor darkgray
				Add-WindowsCapability -Online -Name "SNMP.Client*" | Out-Null
				Get-Service -Name "SNMP" -ErrorAction SilentlyContinue | Set-Service -StartupType Disabled 
			}
		}
	}
	If (Test-Path -Path "HKLM:SYSTEM\CurrentControlSet\Services\SNMP\Parameters\ValidCommunities") {
		(get-item "HKLM:SYSTEM\CurrentControlSet\Services\SNMP\Parameters\ValidCommunities").property| ForEach-Object { Remove-ItemProperty -Name $_ -Path "HKLM:SYSTEM\CurrentControlSet\Services\SNMP\Parameters\ValidCommunities"}
	}
	#Set Community
	Set-Reg "HKLM:\SYSTEM\CurrentControlSet\Services\SNMP\Parameters\ValidCommunities" $ConfigFile.Config.Company.SNMP 4 "DWORD"
	#Sets All info
	Set-Reg "HKLM:\SYSTEM\CurrentControlSet\Services\SNMP\Parameters\RFC1156Agent" "sysServices" 79 "DWORD"
	#Allows All hosts
	If (Test-Path -Path "HKLM:\SYSTEM\CurrentControlSet\Services\SNMP\Parameters\PermittedManagers") {
		(get-item "HKLM:\SYSTEM\CurrentControlSet\Services\SNMP\Parameters\PermittedManagers").property| ForEach-Object { Remove-ItemProperty -Name $_ -Path "HKLM:\SYSTEM\CurrentControlSet\Services\SNMP\Parameters\PermittedManagers"}
	}
}
#============================================================================
#endregion SNMP Setup
#============================================================================
#============================================================================
#region Set Registry Permissions
#============================================================================
Write-Host ("Setting Registry Permissions ... ") -foregroundcolor darkgray
	ForEach ( $item in $ConfigFile.Config.Permissions.Registry.Item) {
		If ($item.User -and $item.Perm -and $item.Action) {
			# Source: https://social.technet.microsoft.com/Forums/en-US/1f082309-dc39-4c7e-ab45-b19094c21877/powershell-script-to-change-permission-of-hkcu-registry-and-make-it-read-only-permission-for-the?forum=winserverpowershell
			# Write-Host ("`tUpdating: '" +  $rootKey + ":\" + $item.Key + "' for '" + $item.User + "' to '" + $item.Action + "' with '" + $item.Perm + "'")	
			Write-Color -Text "Updating: ",
								($rootKey + ":\" + $item.Key),
								" for ",
								$item.key,
								" to ",
								$item.Action,
								" with ",
								$item.Perm -Color White,Red,White,DarkMagenta,White,DarkCyan,White,DarkBlue

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
	}
#============================================================================
#endregion Set Registry Permissions
#============================================================================
#============================================================================
#region Main Local Machine VMWare Horzion Settings
#============================================================================
If (-Not $UserOnly) {
	Write-Host ("VMWare Horzion Settings: ") -foregroundcolor darkgray
	If ([Environment]::Is64BitOperatingSystem) {
		If ($ConfigFile.Config.VMWare_Horizon.AllowCmdLineCredentials) {
			Set-Reg ($ConfigFile.Config.VMWare_Horizon.RegistryKey.replace("\Software\","\Software\Wow6432Node\") + "\Security") "AllowCmdLineCredentials" $ConfigFile.Config.VMWare_Horizon.AllowCmdLineCredentials "DWord"
		}
		If ($ConfigFile.Config.VMWare_Horizon.CertCheckMode) {
			Set-Reg ($ConfigFile.Config.VMWare_Horizon.RegistryKey.replace("\Software\","\Software\Wow6432Node\") + "\Security") "CertCheckMode" $ConfigFile.Config.VMWare_Horizon.CertCheckMode "DWord"
		}
		If ($ConfigFile.Config.VMWare_Horizon.LogInAsCurrentUser) {
			Set-Reg ($ConfigFile.Config.VMWare_Horizon.RegistryKey.replace("\Software\","\Software\Wow6432Node\") + "\Security") "LogInAsCurrentUser" $ConfigFile.Config.VMWare_Horizon.LogInAsCurrentUser "String"
		}
		If ($ConfigFile.Config.VMWare_Horizon.LogInAsCurrentUser_Display) {
			Set-Reg ($ConfigFile.Config.VMWare_Horizon.RegistryKey.replace("\Software\","\Software\Wow6432Node\") + "\Security") "LogInAsCurrentUser_Display" $ConfigFile.Config.VMWare_Horizon.LogInAsCurrentUser_Display "String" 
		}
		If ($ConfigFile.Config.VMWare_Horizon.SSLCipherList) {
			Set-Reg ($ConfigFile.Config.VMWare_Horizon.RegistryKey.replace("\Software\","\Software\Wow6432Node\") + "\Security") "SSLCipherList" $ConfigFile.Config.VMWare_Horizon.SSLCipherList "String"
		}
		If ($ConfigFile.Config.VMWare_Horizon.EnableTicketSSLAuth) {
			Set-Reg ($ConfigFile.Config.VMWare_Horizon.RegistryKey.replace("\Software\","\Software\Wow6432Node\") + "\Security") "EnableTicketSSLAuth" $ConfigFile.Config.VMWare_Horizon.EnableTicketSSLAuth "DWORD" 
		}
		If ($ConfigFile.Config.VMWare_Horizon.AutoUpdateAllowed) {
			Set-Reg ($ConfigFile.Config.VMWare_Horizon.RegistryKey.replace("\Software\","\Software\Wow6432Node\")) "AutoUpdateAllowed" $ConfigFile.Config.VMWare_Horizon.AutoUpdateAllowed "String"
		}
		If ($ConfigFile.Config.VMWare_Horizon.AllowDataSharing) {
			Set-Reg ($ConfigFile.Config.VMWare_Horizon.RegistryKey.replace("\Software\","\Software\Wow6432Node\")) "AllowDataSharing" $ConfigFile.Config.VMWare_Horizon.AllowDataSharing "String"
		}
		If ($ConfigFile.Config.VMWare_Horizon.IpProtocolUsage) {
			Set-Reg ($ConfigFile.Config.VMWare_Horizon.RegistryKey.replace("\Software\","\Software\Wow6432Node\")) "IpProtocolUsage" $ConfigFile.Config.VMWare_Horizon.IpProtocolUsage "String"
		}
		If ($ConfigFile.Config.VMWare_Horizon.NetBIOSDomain) {	
			Set-Reg ($ConfigFile.Config.VMWare_Horizon.RegistryKey.replace("\Software\","\Software\Wow6432Node\")) "DomainName" $ConfigFile.Config.VMWare_Horizon.NetBIOSDomain "String"
		}
		If ($ConfigFile.Config.VMWare_Horizon.Server) {
			Set-Reg ($ConfigFile.Config.VMWare_Horizon.RegistryKey.replace("\Software\","\Software\Wow6432Node\")) "ServerURL" $ConfigFile.Config.VMWare_Horizon.Server "String"
		}
	} else {
		If ($ConfigFile.Config.VMWare_Horizon.AllowCmdLineCredentials) {
			Set-Reg ($ConfigFile.Config.VMWare_Horizon.RegistryKey + "\Security") "AllowCmdLineCredentials" $ConfigFile.Config.VMWare_Horizon.AllowCmdLineCredentials "DWord"
		}
		If ($ConfigFile.Config.VMWare_Horizon.CertCheckMode) {
			Set-Reg ($ConfigFile.Config.VMWare_Horizon.RegistryKey + "\Security") "CertCheckMode" $ConfigFile.Config.VMWare_Horizon.CertCheckMode "DWord"
		}
		If ($ConfigFile.Config.VMWare_Horizon.LogInAsCurrentUser) {
			Set-Reg ($ConfigFile.Config.VMWare_Horizon.RegistryKey + "\Security") "LogInAsCurrentUser"$ConfigFile.Config.VMWare_Horizon.LogInAsCurrentUser "String"
		}
		If ($ConfigFile.Config.VMWare_Horizon.LogInAsCurrentUser_Display) {
			Set-Reg ($ConfigFile.Config.VMWare_Horizon.RegistryKey + "\Security") "LogInAsCurrentUser_Display" $ConfigFile.Config.VMWare_Horizon.LogInAsCurrentUser_Display "String"
		}
		If ($ConfigFile.Config.VMWare_Horizon.SSLCipherList) {
			Set-Reg ($ConfigFile.Config.VMWare_Horizon.RegistryKey + "\Security") "SSLCipherList" $ConfigFile.Config.VMWare_Horizon.SSLCipherList "String"
		}
		If ($ConfigFile.Config.VMWare_Horizon.EnableTicketSSLAuth){
			Set-Reg ($ConfigFile.Config.VMWare_Horizon.RegistryKey + "\Security") "EnableTicketSSLAuth" $ConfigFile.Config.VMWare_Horizon.EnableTicketSSLAuth "DWORD"
		}
		If ($ConfigFile.Config.VMWare_Horizon.AutoUpdateAllowed) {
			Set-Reg ($ConfigFile.Config.VMWare_Horizon.RegistryKey) "AutoUpdateAllowed" "false" "String"
		}
		If ($ConfigFile.Config.VMWare_Horizon.AllowDataSharing) {
			Set-Reg ($ConfigFile.Config.VMWare_Horizon.RegistryKey) "AllowDataSharing" $ConfigFile.Config.VMWare_Horizon.AllowDataSharing "String"
		}
		If ($ConfigFile.Config.VMWare_Horizon.IpProtocolUsage){
			Set-Reg ($ConfigFile.Config.VMWare_Horizon.RegistryKey) "IpProtocolUsage" $ConfigFile.Config.VMWare_Horizon.IpProtocolUsage "String"
		}
		If ($ConfigFile.Config.VMWare_Horizon.NetBIOSDomain){
			Set-Reg ($ConfigFile.Config.VMWare_Horizon.RegistryKey) "DomainName" $ConfigFile.Config.VMWare_Horizon.NetBIOSDomain "String"
		}
		If ($ConfigFile.Config.VMWare_Horizon.Server) {
			Set-Reg ($ConfigFile.Config.VMWare_Horizon.RegistryKey) "ServerURL" $VMware_Horizon_Server "String"
		}
	}
}
#============================================================================
#endregion Main Local Machine VMWare Horzion Settings
#============================================================================
#============================================================================
#region Main Local Machine Temp Cleanup
#============================================================================
If (-Not $UserOnly) {
	If ($configfile.Config.WindowsSettings.ScheduledJobs.Job) {
		If (Get-Command Get-ScheduledJob -errorAction SilentlyContinue) {
			ForEach ($Job in $configfile.Config.WindowsSettings.ScheduledJobs.Job) {
				#Clean Up Old Job
				If (get-ScheduledJob | Where-Object {$_.Name -eq $Job.Name}) {
					get-ScheduledJob | Where-Object {$_.Name -eq $Job.Name} | Unregister-ScheduledJob -Force
				}
				#
				$ArrayJTrigger = $Job.Trigger -split ";"
				$HashJTrigger = @{}
				$ArrayJTrigger | ForEach-Object { 
					If ($_ -match "=") {
						$tajt = $_ -split "="
						$HashJTrigger.Add($tajt[0],$tajt[1])
					} Else {
						$HashJTrigger.Add($_ ,"")
					}
				}
				$SchJobOptions = New-ScheduledJobOption -RunElevated
				Register-ScheduledJob -Name $Job.Name -Trigger $HashJTrigger -ScheduledJobOption $SchJobOptions  -ScriptBlock {$Job.'#text'}
			}
		} Else {
			ForEach ($Job in $configfile.Config.WindowsSettings.ScheduledJobs.Job) {
				If (Get-ScheduledTask | Where-Object { $_.TaskName -eq $Job.Name}) {
					Get-ScheduledTask | Where-Object { $_.TaskName -eq $Job.Name} | Unregister-ScheduledTask -Confirm:$false
				}
				$STT = $null
				$STTF = $null
				$STTR = $null
				$STTA = $null
				$STTDW = $null	
				$EAction = [convert]::ToBase64String([System.Text.encoding]::Unicode.GetBytes($Job.'#text')) 
				$ArrayJTrigger = $Job.Trigger -split ";"
				$ArrayJTrigger | ForEach-Object { 
					If ($_ -match "=") {
						$tajt = $_ -split "="
						If ($tajt[0] -eq "Frequency") {
							$STTF = $tajt[1]
						}
						If ($tajt[0] -eq "RandomDelay") {
							$STTR = [timespan]$tajt[1]
						}
						If ($tajt[0] -eq "At") {
							$STTA = [datetime]$tajt[1]
						}
						If ($tajt[0] -eq "DaysOfWeek") {
							$STTDW =$tajt[1]
						}
					} 
				}
				switch ($STTF) {
					Once {$STT = New-ScheduledTaskTrigger -Once -AT $STTA}
					Daily {
						If ($STTR) {
							$STT = New-ScheduledTaskTrigger -Daily -AT $STTA -RandomDelay $STTR
						} Else {
							$STT = New-ScheduledTaskTrigger -Daily -AT $STTA
						}
					}
					Weekly {$STT = New-ScheduledTaskTrigger -Weekly -AT $STTA -DaysOfWeek $STTDW}
					AtLogon {$STT = New-ScheduledTaskTrigger -AtLogOn}
					AtStartup {$STT = New-ScheduledTaskTrigger -AtStartup}
				}
				$action = New-ScheduledTaskAction -Execute 'Powershell.exe' -Argument ('-NoProfile -WindowStyle Hidden -EncodedCommand "' + $EAction + '"')
				Register-ScheduledTask -Action $action -Trigger $STT -TaskName $Job.Name -Description ("Created by: " +  $MyInvocation.MyCommand.Name + " Script Version: " + $ScriptVersion + " XML Version: " + $ConfigFile.Config.Company.Version) -TaskPath "\Microsoft\Windows\PowerShell\ScheduledJobs"
			}
		}
	}
}
#============================================================================
#endregion Main Local Machine Temp Cleanup
#============================================================================
#============================================================================
#region Main Local Machine Cleanup
#============================================================================
#Recording Version of script
write-host " "
if ($ConfigFile.Config.Company.ScriptVersionValue -and $ConfigFile.Config.Company.ScriptXMLVersionValue -and $ConfigFile.Config.Company.Version -and $ConfigFile.Config.Company.ScriptKey -and $ConfigFile.Config.Company.ScriptDateValue) {
	write-host ("Recording " + $ConfigFile.Config.Company.ScriptVersionValue + ": " + $ScriptVersion + " in " + $ConfigFile.Config.Company.ScriptVersionValue + " Key.") -foregroundcolor "Green"
	Set-Reg ("HKLM:\Software\" + $ConfigFile.Config.Company.ScriptKey) $ConfigFile.Config.Company.ScriptVersionValue  $ScriptVersion "String"
	write-host ("Recording " + $ConfigFile.Config.Company.ScriptXMLVersionValue + ": " + $ConfigFile.Config.Company.Version + " in " + $ConfigFile.Config.Company.ScriptXMLVersionValue + " Key.") -foregroundcolor "Green"
	Set-Reg ("HKLM:\Software\" + $ConfigFile.Config.Company.ScriptKey) $ConfigFile.Config.Company.ScriptXMLVersionValue  $ConfigFile.Config.Company.Version "String"
	Set-Reg ("HKLM:\Software\" + $ConfigFile.Config.Company.ScriptKey) $ConfigFile.Config.Company.ScriptDateValue  (Get-Date -format yyyyMMdd) "String"
}
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
If (-Not [string]::IsNullOrEmpty($LICache + $LogFile)) {
	Stop-Transcript
}
#############################################################################
#endregion Main
#############################################################################
