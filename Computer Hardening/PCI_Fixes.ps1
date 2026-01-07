<#
.SYNOPSIS
    Apply Fixes for PCI scans

  
.DESCRIPTION
    - Used to fix "Microsoft Windows Unquoted Service Path Enumeration" CVE-2013-1609 issue.
	- WinVerifyTrust Signature Validation Vulnerability
	- Encryption Oracle Remediation
	- Remove Microsoft Sliverlight
	- Remove SSPORT Driver
	- Remove SigPlus.ocx and SigSign.ocx
	- IE CVE-2017-8529
	- Windows Speculative Execution
	- Unquoted Services Path Enumeration
	- Removing vulnerable Appx packages
	- Remove Adobe Flash
	- Remove Log4j from Crystal Reports
	- Remove Log4j from customapp2
	- Enable MS Office Automatic Updates
	- Install Microsoft Updates
	- Fix Insecure Windows Services Permissions
	- Update Microsoft Defender Definitions
	- Remove Log4j from cssimpact
	- isusweb.dll removal
	- Microsoft XML Parser (MSXML) and XML Core Services Unsupported removal or update.
	- Sysmon update
	- Hardened UNC Paths
.RELATED LINKS
	https://isc.sans.edu/diary/Help+eliminate+unquoted+path+vulnerabilities/14464
    http://cwe.mitre.org/data/definitions/428.html
    http://www.ryanandjeffshow.com/blog/2013/04/11/powershell-fixing-unquoted-service-paths-complete/
        

.PARAMETER NoMSUpdate
	Do not run Microsoft Update
.PARAMETER NoOfficeUpdate
	Do not run Office Update
.PARAMETER ForceUpdate
	Force script to run even if already run.
.PARAMETER UpdateMSXML
	Force update of MSXML4
.NOTES
  Changes:
    1.0.0  - Draft Script
    1.0.1  - Added more tries and different ways to run commands. Fixed logging path, force script to run as admin, fix permission to force removal of flash.
    1.0.2  - Fixed script errors and msiexec issue. 
    1.0.3  - Fixed script issue where WinTrust registry is not created. Fixed issue with flash service and files not being removed.
    1.0.4  - Found way to query installed packages without WMI. 
    1.0.5  - Added fixes for SSPORT.SYS, SigPlus.ocx and SigSign.ocx issues. Attempt Defender definition update. Added Move-OnReboot to fix deleting file issue.
    1.0.6  - Clean up of and reorganize of fixes.
    1.0.7  - Added Encryption Oracle Remediation.
    1.0.8  - Created array of appx packages to remove.
	1.0.9  - Enable MS Office Updates in registry. Set Oracle remediation to force updated clients Option.
	1.0.10 - Added Http get to update Microsoft Defender update and install it. 
	1.0.11 - Fixed Issues with appx.
	1.0.12 - Add logic to remove Microsoft Sliverlight and Crystal report Log4j files. Also added call to run Microsoft Update.
	1.0.13 - Add more logging
	1.0.14 - Fixed issue with nuget. Force update of MS Office.
	1.0.15 - Added switches to turn off MS and Office updates. Added more logging for sliverlight removal. 
	1.0.16 - Test if Get-ProvisionedAppPackage is available to clean up logs.
	1.0.17 - Write script version to registry, to allow for better remediation tracking. Also script will not apply setting unless newer version.
	1.0.18 - Force network connection to be set to private. Added logic to remove Intel Wireless Network Drivers.
	1.0.19 - Fix Insecure Windows Services Permissions
	1.0.20 - Better logic for Defender update.
	1.0.21 - Better logic for Defender update. Fix issues with flash removal.
	1.0.22 - Remove Log4j from customapp2
	1.0.23 - Fix CVE-2017-8529 and Windows Speculative Execution
	1.0.24 - Fix for Microsoft XML Parser (MSXML) and XML Core Services Unsupported, Remove Log4j from cssimpact,  isusweb.dll removal
	1.0.25 - Remove Microsoft XML Parser (MSXML) by default, unless -UpdateMSXML is specified.
	1.0.26 - Disable 3DES.
	1.0.27 - 20250630 Silverlight deny creating files and folders. Also Disable Configuration Manager tasks.
	1.0.28 - 20251113 Sysmon update. Cleaned up errors with Sliverlight deny rules. and Reg key check for Sliverlight removal.
	1.0.29 - 20251114 Fixed Hardened UNC Paths
	1.0.30 - 20251222 ECDH public server param reuse fix
  Release Date: 07/17/2024
  Update Date:  11-13-2025
   
  Author: Paul Fuller

.EXAMPLE
    .\PCI_Fixes.ps1
	powershell.exe -executionpolicy bypass -file "C:\IT_Updates\PCI_Fixes.ps1"
.PARAMETER NoMSUpdate
	Do not run Microsoft Update
.PARAMETER NoOfficeUpdate
	Do not run Office Update

#>
#region Parameters
Param(
	[CmdletBinding()]
		[switch]$NoMSUpdate = $false,
		[switch]$NoOfficeUpdate = $false,
		[switch]$ForceUpdate = $false,
		[switch]$UpdateMSXML = $false
)
#endregion Parameters
#region Force Script to run as Admin 
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))  
{  
  $arguments = "& '" +$myinvocation.mycommand.definition + "'"
  $PS = Start-Process powershell -Verb runAs -ArgumentList $arguments
  Exit $PS
}
##Requires -RunAsAdministrator
#endregion Force Script to run as Admin 
#region Variables 
$ScriptVersion = "1.0.29"
$ScriptRecordVersionValue = "PCIScriptVersion"
$ScriptRecordDateValue = "PCIScriptRunDate"
$ScriptRecordkey = "Github"
$LocalLogs = "C:\IT_Updates\Logs\"
$WindowsDefenderUpdateHTTP="http://localhttpserver/repository/Packages/Microsoft%20Defender/mpam-fe.exe"
$MSXML4UpdateHTTP = "http://localhttpserver/repository/Packages/MSXML4/msxml.msi"
$SysmonUpdateHTTP = "https://download.sysinternals.com/files/Sysmon.zip"
$FileDate = (Get-Date -format yyyyMMdd-hhmm)
$sw = [Diagnostics.Stopwatch]::StartNew()
$CBS = @()
$Uninstall = @()
$Services = @()
$RebootRequired = $False
$RemoveAppx=@(
	"Microsoft.Microsoft3DViewer"
	"Microsoft.MSPaint"
	"Microsoft.BingWeather"
	"Microsoft.Office.OneNote"
)
$RemoveGroupsFromServices = @(
	"Users"
	"Everyone"
	"Authenticated Users"
	"Domain Users"
)
$RemovePermissionFromServices = @(
	"FullControl"
	"Modify"
	"Write"
)
$DisableScheduledTasks = @(
"Configuration Manager Health Evaluation"
"Configuration Manager Idle Detection"
"Configuration Manager Maintenance"
)
Add-Type -AssemblyName System.IO.Compression.FileSystem
#endregion Variables 
#region LogFile
If (-Not $LogFile) {
    If (Test-Path -Path $LocalLogs) {
		If (-Not $LocalLogs.EndsWith("\")) {
			$LocalLogs = $LocalLogs + "\"
		}
        $LogFile = ($LocalLogs + ($MyInvocation.MyCommand.Name -replace ".ps1","") + "_" + $env:COMPUTERNAME + "_" + $FileDate + ".log")
    }Else{
        $LogFile = ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\Logs\" + ($MyInvocation.MyCommand.Name -replace ".ps1","") + "_" + $env:COMPUTERNAME + "_" + $FileDate + ".log")
    }
}
If (-Not [string]::IsNullOrEmpty($LogFile)) {
	If (-Not( Test-Path (Split-Path -Path $LogFile -Parent))) {
		New-Item -ItemType directory -Path (Split-Path -Path $LogFile -Parent)
        $Acls = Get-Acl (Split-Path -Path $LogFile -Parent)
        $Ar = New-Object system.Security.AccessControl.FileSystemAccessRule('Users', "FullControl", 'ContainerInherit,ObjectInherit', "None", "Allow")
        $Acls.Setaccessrule($Ar)
        Set-Acl (Split-Path -Path $LogFile -Parent) $Acls
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
#endregion LogFile
#region Functions
function Set-Reg {
	[CmdletBinding()] 
Param 
( 
	[Parameter(Mandatory=$true,Position=1,HelpMessage="Path to Registry Key")][string]$regPath, 
	[Parameter(Mandatory=$true,Position=2,HelpMessage="Name of Value")][string]$name,
	[Parameter(Mandatory=$true,Position=3,HelpMessage="Data for Value")]$value,
	[Parameter(Mandatory=$true,Position=4,HelpMessage="Type of Value")][ValidateSet("String", "ExpandString","Binary","DWord","MultiString","Qword","Unknown",IgnoreCase =$true)][string]$type ,
	[Parameter(Mandatory=$false,Position=5,HelpMessage="Comment value")]$comment
) 
$key=$null
#Source: https://github.com/nichite/chill-out-windows-10/blob/master/chill-out-windows-10.ps1
# String: Specifies a null-terminated string. Equivalent to REG_SZ.
# ExpandString: Specifies a null-terminated string that contains unexpanded references to environment variables that are expanded when the value is retrieved. Equivalent to REG_EXPAND_SZ.
# Binary: Specifies binary data in any form. Equivalent to REG_BINARY.
# DWord: Specifies a 32-bit binary number. Equivalent to REG_DWORD.
# MultiString: Specifies an array of null-terminated strings terminated by two null characters. Equivalent to REG_MULTI_SZ.
# Qword: Specifies a 64-bit binary number. Equivalent to REG_QWORD.
# Unknown: Indicates an unsupported registry data type, such as REG_RESOURCE_LIST.

$key = $null
$regvalue = $null
$regname = $null
If(Test-Path $regPath) {
	$key = Get-Item -Path $regPath
}Else{
	Write-Host ("`tCreating Key:" + $regPath )
	New-Item -Path $regPath -Force | Out-Null
	$key = Get-Item -Path $regPath
}
If($type -eq "Binary" -and $value.GetType().Name -eq "String" -and $value -match ",") {
	$value = [byte[]]($value -split ",")
}
If ($key.Property.Equals($Name)){
	If($key.GetValue($Name) -eq $Value) {
		Write-Host ("`tSame:" + $regPath + "\" + $name + " = " + $value)
	}Else {
		If($null -eq $value){
			Write-Host ("`tCreating:" + $regPath + "\" + $name + " = " + $value)
		}Else {
			Write-Host ("`tUpdating:" + $regPath + "\" + $name + " = " + $value)
		}

		New-ItemProperty -Path $regPath -Name $name -Value $value -PropertyType $type -Force | Out-Null
	}
}Else{
	If($null -eq $value){
		Write-Host ("`tCreating:" + $regPath + "\" + $name + " = " + $value)
	}Else {
		Write-Host ("`tUpdating:" + $regPath + "\" + $name + " = " + $value)
	}
	New-ItemProperty -Path $regPath -Name $name -Value $value -PropertyType $type -Force | Out-Null
}
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
		$FileAdminAcl.SetAccessRule($AdminACL)
		$DirAdminAcl.SetAccessRule($AdminACL)
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

function Move-OnReboot {
	[CmdletBinding()] 
	Param ( 
		[parameter(Mandatory=$True)]$Path, 
		$Destination
	)
	Begin{
		try{
			[Microsoft.PowerShell.Commands.AddType.AutoGeneratedTypes.MoveFileUtils]|Out-Null
		}catch{
			$memberDefinition = @'
[DllImport("kernel32.dll", SetLastError=true, CharSet=CharSet.Auto)]
public static extern bool MoveFileEx(string lpExistingFileName, string lpNewFileName, int dwFlags);
'@
			Add-Type -Name MoveFileUtils -MemberDefinition $memberDefinition
		}
	}
	Process{
		$Path="$((Resolve-Path $Path).Path)"
		if ($Destination){
			$Destination = $executionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Destination)
			#Write-Output ("`tMove file " + $Path + " to " + $Destination + " on next reboot")
		}else{
			$Destination = [Management.Automation.Language.NullString]::Value
			#Write-Output ("`tDelete file " + $Path + " on next reboot")
		}
		$MOVEFILE_DELAY_UNTIL_REBOOT = 0x00000004
		[Microsoft.PowerShell.Commands.AddType.AutoGeneratedTypes.MoveFileUtils]::MoveFileEx($Path, $Destination, $MOVEFILE_DELAY_UNTIL_REBOOT)
		}
	End{}
}
#endregion Functions
If (Test-Path -Path ("HKLM:\Software\" + $ScriptRecordkey)) {
	Try{
		$RegistryScriptVersion = Get-ItemPropertyValue -Path ("HKLM:\Software\" + $ScriptRecordkey) -Name $ScriptRecordVersionValue -ErrorAction SilentlyContinue
	}Catch{
		$RegistryScriptVersion = $null
	}
}
If ($RegistryScriptVersion -lt $ScriptVersion -or $ForceUpdate) {
	#Older or now version run script.
	#region Fix WinVerifyTrust Signature Validation Vulnerability
	If ((Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Cryptography\Wintrust\Config" -ErrorAction SilentlyContinue)."EnableCertPaddingCheck" -ne 1 -and (Get-ItemProperty -Path "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Cryptography\Wintrust\Config" -ErrorAction SilentlyContinue)."EnableCertPaddingCheck" -ne 1) {
		Write-Host ("WinVerifyTrust Signature Validation Vulnerability")
		Set-Reg ("HKLM:\SOFTWARE\Microsoft\Cryptography\Wintrust\Config") "EnableCertPaddingCheck" 1 "DWORD"
		Set-Reg ("HKLM:\SOFTWARE\Wow6432Node\Microsoft\Cryptography\Wintrust\Config") "EnableCertPaddingCheck" 1 "DWORD"
	}
	#endregion Fix WinVerifyTrust Signature Validation Vulnerability
	#region Encryption Oracle Remediation
	If ((Get-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System\CredSSP\Parameters" -ErrorAction SilentlyContinue)."AllowEncryptionOracle" -le 1 ) {
		# Write-Host ("Fixing Encryption Oracle Remediation")
		Set-Reg -regPath ("HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System\CredSSP\Parameters") -name "AllowEncryptionOracle" -value 0 -type "DWORD" -comment "Fixing Encryption Oracle Remediation"
	}
	#endregion Encryption Oracle Remediation
	#region Scheduled Tasks
		Foreach ($Task in $DisableScheduledTasks) {
			Get-ScheduledTask -TaskName  $Task -ErrorAction SilentlyContinue | Where-Object {$_.starte -ne "Disabled"} | Disable-ScheduledTask -ErrorAction SilentlyContinue
		}
	#endregion Scheduled Takes
	#region Remove Microsoft Sliverlight
		$RemoveAppName = "Microsoft Silverlight"
		#region Remove Sliverlight by GUID
		Write-host "Trying to uninstall any installed $RemoveAppName"
		Try {
			# Write-host "Trying to uninstall any installed $RemoveAppName with Fallback"
			MsiExec.exe /X"{89F4137D-6C26-4A84-BDB8-2E5A4BB71E00}" /quiet
		}Catch{
			
		}
		Try {
			# Write-host "Trying to uninstall any installed $RemoveAppName with Fallback"
			MsiExec.exe /X"{83B900D2-51E8-4B67-BD75-643C8F14BBD8}" /quiet
		}Catch{		
		}
		Try {
			# Write-host "Trying to uninstall any installed $RemoveAppName with Fallback"
			MsiExec.exe /X"{24AB137D-C266-A844-DB8B-0A4E1B0BE572}" /quiet
		}Catch{
		}
		Try {
			# Write-host "Trying to uninstall any installed $RemoveAppName with Fallback"
			MsiExec.exe /X"{BB2A70DE-1FD4-4A3F-926E-D7F4FC23EADE}" /quiet
		}Catch{
		}

		#endregion Remove Sliverlight by GUID
		
		Write-host "Trying to delete program folder for $RemoveAppName"
		$Destination = ($env:ProgramFiles + "\" + $RemoveAppName)
		$Destination86 = (${env:ProgramFiles(x86)} + "\" + $RemoveAppName)
		$InstallerDestinations = @(
			'c:\Windows\Installer\{89F4137D-6C26-4A84-BDB8-2E5A4BB71E00}'
			'c:\Windows\Installer\13e31e30.msi'
		)
		#region remove Sliverlight Files
		If ((Test-Path -Path $Destination -ErrorAction SilentlyContinue) -OR (Test-Path -Path $Destination86 -ErrorAction SilentlyContinue)) {
			If (Test-Path -Path $Destination -ErrorAction SilentlyContinue) {
				Write-host ("`tTrying to remove: " + $Destination)
				If (Get-Command -Name Set-Owner -ErrorAction SilentlyContinue) {
					Set-Owner -Path $Destination -Account 'Administrators'
				}Else{
					$ACL = Get-ACL $Destination 
					$Group = New-Object System.Security.Principal.NTAccount("Builtin", "Administrators")
					$ACL.SetOwner($Group)
					Set-Acl -Path $Destination -AclObject $ACL	
				}
				$Acls = Get-Acl $Destination
				# Remove Deny ACLs
				foreach ($deny in $Acls.Access | Where-Object { $_.AccessControlType -eq "Deny" }) {
					$Acls.RemoveAccessRule($deny) | Out-Null
				}
				# Add Full Control for Administrators
				$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule("BUILTIN\Administrators", "FullControl", "Allow")
				$Acls.SetAccessRule($Ar)
				#Set modified ACLs
				Set-Acl $Destination $Acls -ErrorAction SilentlyContinue	
				Write-host ("`tTrying to remove: " + $Destination)
				Try{
					Remove-Item -Force -Recurse -Path $Destination 
					If(Test-Path -Path $Destination -ErrorAction SilentlyContinue) {
						Move-OnReboot -Path $Destination
						$RebootRequired = $True
					}
				}Catch {
					if ($_.Exception.Message -like "*Access*is denied*") {
						$Acls = Get-Acl $Destination
						If ($acls.Access.Where({$_.AccessControlType -eq "deny"}).count -ge 4) {
							Write-Host ("`t`tDeny ACLs Inplace: " + $Destination)
						}else {
							Write-Host ("`t`tPlease review ACLs: " + $Destination) -ForegroundColor Red
						}
					}Elseif ($_.Exception.Message -like "*being used by another process*") {
						Write-Host ("`t`tIn use by another process: " + $Destination) -ForegroundColor Yellow
					}else{
						Write-Host ("`t`tFailed to remove: " + $Destination) -ForegroundColor Red
					}
				}
			}
			If (Test-Path -Path $Destination86 -ErrorAction SilentlyContinue) {
				Write-host ("`tTrying to remove: " + $Destination86)
				If (Get-Command -Name Set-Owner -ErrorAction SilentlyContinue) {
					Set-Owner -Path $Destination86 -Account 'Administrators'
				}Else{
					$ACL = Get-ACL $Destination86 
					$Group = New-Object System.Security.Principal.NTAccount("Builtin", "Administrators")
					$ACL.SetOwner($Group)
					Set-Acl -Path $Destination86 -AclObject $ACL	
				}
				$Acls = Get-Acl $Destination86
				# Remove Deny ACLs
				foreach ($deny in $Acls.Access | Where-Object { $_.AccessControlType -eq "Deny" }) {
					$Acls.RemoveAccessRule($deny) | Out-Null
				}
				# Add Full Control for Administrators
				$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule("BUILTIN\Administrators", "FullControl", "Allow")
				$Acls.SetAccessRule($Ar)
				#Set modified ACLs
				Set-Acl $Destination86 $Acls -ErrorAction SilentlyContinue	
				Write-host ("`tTrying to remove: " + $Destination86)
				Try{
					Remove-Item -Force -Recurse -Path $Destination86 
					If(Test-Path -Path $Destination86 -ErrorAction SilentlyContinue) {
						Move-OnReboot -Path $Destination86
						$RebootRequired = $True
					}
				}Catch {
					if ($_.Exception.Message -like "*Access*is denied*") {
						$Acls = Get-Acl $Destination86
						If ($acls.Access.Where({$_.AccessControlType -eq "deny"}).count -ge 4) {
							Write-Host ("`t`tDeny ACLs Inplace: " + $Destination86)
						}else {
							Write-Host ("`t`tPlease review ACLs: " + $Destination86) -ForegroundColor Red
						}
					}else{
						Write-Host ("`t`tFailed to remove: " + $Destination86) -ForegroundColor Red
					}
				}
			}
			#endregion remove Sliverlight Files
			#region clean registry
			 	If (Test-Path -Path "HKLM:\Software\Microsoft\Silverlight" -ErrorAction SilentlyContinue) {
					reg delete 'HKLM\Software\Microsoft\Silverlight' /f
				}
				if (Test-Path "Registry::HKEY_CLASSES_ROOT\Installer\Products\D7314F9862C648A4DB8BE2A5B47BE100") {
					reg delete 'HKEY_CLASSES_ROOT\Installer\Products\D7314F9862C648A4DB8BE2A5B47BE100' /f
				}
				# reg delete 'HKEY_CLASSES_ROOT\Interface\{C4130531-1AF7-459E-A116-97AC497F414A}' /f
				# reg delete 'HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Interface\{C4130531-1AF7-459E-A116-97AC497F414A}' /f
				# reg delete 'HKEY_LOCAL_MACHINE\SOFTWARE\Classes\WOW6432Node\Interface\{C4130531-1AF7-459E-A116-97AC497F414A}' /f
				# reg delete 'HKEY_CLASSES_ROOT\Wow6432Node\Interface\{C4130531-1AF7-459E-A116-97AC497F414A}' /f
				If (Test-Path -Path "HKLM:\Software\Classes\Installer\Products\D7314F9862C648A4DB8BE2A5B47BE100" -ErrorAction SilentlyContinue) {
					reg delete 'HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Installer\Products\D7314F9862C648A4DB8BE2A5B47BE100' /f
				}
				If (Test-Path -Path "HKLM:\Software\Classes\WOW6432Node\Installer\Products\D7314F9862C648A4DB8BE2A5B47BE100" -ErrorAction SilentlyContinue) {
					reg delete 'HKEY_LOCAL_MACHINE\SOFTWARE\Classes\WOW6432Node\Installer\Products\D7314F9862C648A4DB8BE2A5B47BE100' /f
				}
				If (Test-Path -Path "HKLM:\Software\Classes\WOW6432Node\CLSID\{DFEAF541-F3E1-4c24-ACAC-99C30715084A}" -ErrorAction SilentlyContinue) {
					reg delete 'HKEY_LOCAL_MACHINE\SOFTWARE\Classes\WOW6432Node\CLSID\{DFEAF541-F3E1-4c24-ACAC-99C30715084A}' /f
				}
				If (Test-Path -Path "HKLM:\Software\Classes\WOW6432Node\TypeLib\{283C8576-0726-4DBC-9609-3F855162009A}" -ErrorAction SilentlyContinue) {
					reg delete 'HKEY_LOCAL_MACHINE\SOFTWARE\Classes\WOW6432Node\TypeLib\{283C8576-0726-4DBC-9609-3F855162009A}' /f
				}
				If (Test-Path -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\05A4B1AD0412E9149B7C44C8ED897F76" -ErrorAction SilentlyContinue) {
					reg delete 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\05A4B1AD0412E9149B7C44C8ED897F76' /f
				}
				If (Test-Path -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\07EF711D411E1BE4DB79C87C73FBE662" -ErrorAction SilentlyContinue) {
					reg delete 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\07EF711D411E1BE4DB79C87C73FBE662' /f
				}
				If (Test-Path -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\08196EFE47C51DE4885D641031114EBF" -ErrorAction SilentlyContinue) {
					reg delete 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\08196EFE47C51DE4885D641031114EBF' /f
				}
				If (Test-Path -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\088DAA370DECD5145A22F2BC95B32FB2" -ErrorAction SilentlyContinue) {
					reg delete 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\088DAA370DECD5145A22F2BC95B32FB2' /f
				}
				If (Test-Path -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\0AB480B8A63A9F94BBB1019D59C1E2B5" -ErrorAction SilentlyContinue) {
					reg delete 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\0AB480B8A63A9F94BBB1019D59C1E2B5' /f
				}
				If (Test-Path -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\0B9C15ACDD4DB3748844EE9D585A2E6F" -ErrorAction SilentlyContinue) {
					reg delete 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\0B9C15ACDD4DB3748844EE9D585A2E6F' /f
				}
				If (Test-Path -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\0C494A0C8886AA94D92579EFF5CC10FD" -ErrorAction SilentlyContinue) {
					reg delete 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\0C494A0C8886AA94D92579EFF5CC10FD' /f
				}
				If (Test-Path -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\0E6E033981A67A34387D3800D9E2B999" -ErrorAction SilentlyContinue) {
					reg delete 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\0E6E033981A67A34387D3800D9E2B999' /f
				}
				If (Test-Path -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\1091554B7FEF60443AAF37F0F52AC994" -ErrorAction SilentlyContinue) {
					reg delete 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\1091554B7FEF60443AAF37F0F52AC994' /f
				}
				If (Test-Path -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\1211261CEB465B04B8193520D9E42FEA" -ErrorAction SilentlyContinue) {
					reg delete 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\1211261CEB465B04B8193520D9E42FEA' /f
				}
				If (Test-Path -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\131AD5D1647C09D4FB9267032D9267E5" -ErrorAction SilentlyContinue) {
					reg delete 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\131AD5D1647C09D4FB9267032D9267E5' /f
				}
				If (Test-Path -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\1364D1438F5E7604393DDFA6AA4A4914" -ErrorAction SilentlyContinue) {
					reg delete 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\1364D1438F5E7604393DDFA6AA4A4914' /f
				}
				If (Test-Path -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\9274B1F3DD752DD46B78DE222629060E" -ErrorAction SilentlyContinue) {
					reg delete 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\9274B1F3DD752DD46B78DE222629060E' /f
				}
				If (Test-Path -Path "HKLM:\Software\Classes\TypeLib\{283C8576-0726-4DBC-9609-3F855162009A}" -ErrorAction SilentlyContinue) {
					reg delete 'HKEY_CLASSES_ROOT\TypeLib\{283C8576-0726-4DBC-9609-3F855162009A}' /f
				}
				If (Test-Path -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\App Paths\install.exe" -ErrorAction SilentlyContinue) {
					reg delete 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\install.exe' /f
				}
				If (Test-Path -Path "HKLM:\Software\Classes\AgControl.AgControl" -ErrorAction SilentlyContinue) {
					reg delete 'HKEY_CLASSES_ROOT\AgControl.AgControl' /f
				}
				If (Test-Path -Path "HKLM:\Software\Classes\AgControl.AgControl.5.1" -ErrorAction SilentlyContinue) {
					reg delete 'HKEY_CLASSES_ROOT\AgControl.AgControl.5.1' /f
				}
				If (Test-Path -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{89F4137D-6C26-4A84-BDB8-2E5A4BB71E00}" -ErrorAction SilentlyContinue) {
					reg delete 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{89F4137D-6C26-4A84-BDB8-2E5A4BB71E00}' /f
				}
				If (Test-Path -Path "HKLM:\Software\Classes\Installer\Patches\ED07A2BB4DF1F3A429E67D4FCF32AEED" -ErrorAction SilentlyContinue) {
					reg delete 'HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Installer\Patches\ED07A2BB4DF1F3A429E67D4FCF32AEED' /f
				}
				If (Test-Path -Path "HKLM:\Software\Classes\MIME\Database\Content Type\application/x-silverlight" -ErrorAction SilentlyContinue) {
					reg delete 'HKEY_LOCAL_MACHINE\SOFTWARE\Classes\MIME\Database\Content Type\application/x-silverlight' /f
				}
				If (Test-Path -Path "HKLM:\Software\Classes\MIME\Database\Content Type\application/x-silverlight-2" -ErrorAction SilentlyContinue) {
					reg delete 'HKEY_LOCAL_MACHINE\SOFTWARE\Classes\MIME\Database\Content Type\application/x-silverlight-2' /f
				}
				If (Test-Path -Path "HKLM:\Software\Microsoft\MATS\WindowsInstaller\{89F4137D-6C26-4A84-BDB8-2E5A4BB71E00}" -ErrorAction SilentlyContinue) {
					reg delete 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\MATS\WindowsInstaller\{89F4137D-6C26-4A84-BDB8-2E5A4BB71E00}' /f
				}
				If (Test-Path -Path "HKLM:\SYSTEM\ControlSet001\Control\Terminal Server\WinStations\RDP-Tcp\VideoRemotingWindowNames" -ErrorAction SilentlyContinue) {
					reg delete 'HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Control\Terminal Server\WinStations\RDP-Tcp\VideoRemotingWindowNames' /v "MicrosoftSilverlight" /f
				}
				If (Test-Path -Path "HKU:\.DEFAULT\Software\AppDataLow\Software\Microsoft\Silverlight" -ErrorAction SilentlyContinue) {
					reg delete 'HKEY_USERS\.DEFAULT\Software\AppDataLow\Software\Microsoft\Silverlight' /f
				}
				If (Test-Path -Path "HKU:\.DEFAULT\Software\Microsoft\Silverlight" -ErrorAction SilentlyContinue) {
					reg delete 'HKEY_USERS\.DEFAULT\Software\Microsoft\Silverlight' /f
				}
				
				$searchString = "Microsoft Silverlight"
				Get-ChildItem "HKLM:\SOFTWARE\Classes\Installer\Assemblies" | Where-Object {$_.name -match $searchString}| Remove-Item
				Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders" | Where-Object {$_.name -match $searchString}| Remove-Item
				Get-Item "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\UFH\SHC" | ForEach-Object {
						$props = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
						foreach ($property in $props.PSObject.Properties) {
							$Value
							if ($property.Value -is [string[]]) { # REG_MULTI_SZ is string[]
								foreach ($entry in $property.Value) {
									if ($entry -like "*$searchString*") {
										Remove-ItemProperty -Path $_.PSPath -Name $property.Name -ErrorAction SilentlyContinue | Out-Null
										# Write-Output "Found in $($_.PSPath) -> $($property.Name): $entry"
									}
								}
							}
						}
					}
			#endregion clean registry
		}
		#region Create empty Silverlight folder and set permissions to prevent future installations
			If (-Not (Test-Path -Path $Destination)) {
				New-Item -Path $Destination -ItemType Directory -Force | Out-Null
			}
				If (Get-Command -Name Set-Owner -ErrorAction SilentlyContinue) {
					Set-Owner -Path $Destination -Account 'Administrators'
				}Else{
					$ACL = Get-ACL $Destination 
					$Group = New-Object System.Security.Principal.NTAccount("Builtin", "Administrators")
					$ACL.SetOwner($Group)
					Set-Acl -Path $Destination -AclObject $ACL	
				}
				$Acls = Get-Acl $Destination
				$Acls.SetAccessRuleProtection($True, $False)
				$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule("NT AUTHORITY\SYSTEM", "FullControl", "Deny")
				$Acls.SetAccessRule($Ar)
				$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule("BUILTIN\Administrators", "FullControl", "Deny")
				$Acls.SetAccessRule($Ar)
				$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule("BUILTIN\Users", "FullControl", "Deny")
				$Acls.SetAccessRule($Ar)
				$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule("NT SERVICE\TrustedInstaller", "FullControl", "Deny")
				$Acls.SetAccessRule($Ar)
				# $Ar = New-Object system.Security.AccessControl.FileSystemAccessRule("APPLICATION PACKAGE AUTHORITY\ALL APPLICATION PACKAGES", "FullControl", "Deny")
				# $Acls.SetAccessRule($Ar)
				# $Ar = New-Object system.Security.AccessControl.FileSystemAccessRule("APPLICATION PACKAGE AUTHORITY\ALL RESTRICTED APPLICATION PACKAGES", "FullControl", "Deny")
				# $Acls.SetAccessRule($Ar)
				Set-Acl -Path $Destination -AclObject $Acls -ErrorAction SilentlyContinue

			If (-Not (Test-Path -Path $Destination86)) {
				New-Item -Path $Destination86 -ItemType Directory -Force | Out-Null
			}
				If (Get-Command -Name Set-Owner -ErrorAction SilentlyContinue) {
					Set-Owner -Path $Destination86 -Account 'Administrators'
				}Else{
					$ACL = Get-ACL $Destination86 
					$Group = New-Object System.Security.Principal.NTAccount("Builtin", "Administrators")
					$ACL.SetOwner($Group)
					Set-Acl -Path $Destination86 -AclObject $ACL	
				}
				$Acls = Get-Acl $Destination86
				$Acls.SetAccessRuleProtection($True, $False)
				$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule("NT AUTHORITY\SYSTEM", "FullControl", "Deny")
				$Acls.SetAccessRule($Ar)
				$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule("BUILTIN\Administrators", "FullControl", "Deny")
				$Acls.SetAccessRule($Ar)
				$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule("BUILTIN\Users", "FullControl", "Deny")
				$Acls.SetAccessRule($Ar)
				$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule("NT SERVICE\TrustedInstaller", "FullControl", "Deny")
				$Acls.SetAccessRule($Ar)
				# $Ar = New-Object system.Security.AccessControl.FileSystemAccessRule("APPLICATION PACKAGE AUTHORITY\ALL APPLICATION PACKAGES", "FullControl", "Deny")
				# $Acls.SetAccessRule($Ar)
				# $Ar = New-Object system.Security.AccessControl.FileSystemAccessRule("APPLICATION PACKAGE AUTHORITY\ALL RESTRICTED APPLICATION PACKAGES", "FullControl", "Deny")
				# $Acls.SetAccessRule($Ar)
				Set-Acl -Path $Destination86 -AclObject $Acls -ErrorAction SilentlyContinue
		#endregion Create empty Silverlight folder and set permissions to prevent future installations
		#region remove Installer Files
			ForEach ($InstallerDestination in $InstallerDestinations) {
				#Remove Installer Files
				If (Test-Path -Path $InstallerDestination -ErrorAction SilentlyContinue) {
					Write-host ("`tTrying to remove: " + $InstallerDestination)
					If (Get-Command -Name Set-Owner -ErrorAction SilentlyContinue) {
						Set-Owner -Path $InstallerDestination -Account 'Administrators'
					}Else{
						$ACL = Get-ACL $InstallerDestination 
						$Group = New-Object System.Security.Principal.NTAccount("Builtin", "Administrators")
						$ACL.SetOwner($Group)
						Set-Acl -Path $InstallerDestination -AclObject $ACL	
					}
					$Acls = Get-Acl $InstallerDestination
					$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule("BUILTIN\Administrators", "FullControl", "Allow")
					$Acls.SetAccessRule($Ar)
					Set-Acl $InstallerDestination $Acls -ErrorAction SilentlyContinue	
					Write-host ("`tTrying to remove: " + $InstallerDestination)
					Remove-Item -Force -Recurse -Path $InstallerDestination 
					If(Test-Path -Path $InstallerDestination -ErrorAction SilentlyContinue) {
						Move-OnReboot -Path $InstallerDestination
						$RebootRequired = $True
					}	
				}
			}
		#endregion remove Installer Files
		#region search for SilverLight MSI
			Get-ChildItem "C:\Windows\Installer" -Filter *.msi | 
				ForEach-Object {
					$msiPath = $_.FullName
					# $props = Get-WmiObject -Query "SELECT * FROM Win32_Product WHERE LocalPackage='$msiPath'"
					$installer = New-Object -ComObject WindowsInstaller.Installer
					Try{
						$database = $installer.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $installer, @($msiPath, 0))
						$view = $database.OpenView("SELECT Value FROM Property WHERE Property = 'ProductName'")
						$view.Execute()
						$record = $view.Fetch()
						if ($record) {
							$productName = $record.StringData(1)
						} else {
							$productName = $null
						}
						if ($productName -like "*Silverlight*") {
							Write-Output $msiPath

							Write-Host ("`tTrying to remove: " + $msiPath)
							If (Get-Command -Name Set-Owner -ErrorAction SilentlyContinue) {
								Set-Owner -Path $msiPath -Account 'Administrators'
							}Else{
								$ACL = Get-ACL $msiPath 
								$Group = New-Object System.Security.Principal.NTAccount("Builtin", "Administrators")
								$ACL.SetOwner($Group)		
								Set-Acl -Path $msiPath -AclObject $ACL
							}
							$Acls = Get-Acl $msiPath
							$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule("BUILTIN\Administrators", "FullControl", "Allow")
							$Acls.SetAccessRule($Ar)
							Set-Acl $msiPath $Acls -ErrorAction SilentlyContinue
							Remove-Item -Force -Path $msiPath -ErrorAction SilentlyContinue
							If(Test-Path -Path $msiPath -ErrorAction SilentlyContinue) {
								Move-OnReboot -Path $msiPath
								$RebootRequired = $True
							}
						}
					} catch {
						<#Do this if a terminating exception happens#>
					}
				}
		#endregion search for SilverLight MSI		
	#endregion Remove Microsoft Sliverlight
	#region SSPORT Driver
	If (Test-Path ($env:SystemRoot + "\System32\drivers\SSPORT.SYS")){
		$FileInfo = Get-Item -Path ($env:SystemRoot + "\System32\drivers\SSPORT.SYS")
		#If ($FileInfo.VersionInfo.ProductVersion -lt "1.0.0.1"){
			Write-Host "Removing SSPORT.SYS driver"
			If (Test-Path -Path "HKLM:\SYSTEM\CurrentControlSet\Services\SSPORT") {
				Remove-Item -Recurse -Force -Path "HKLM:\SYSTEM\CurrentControlSet\Services\SSPORT"
			}
			If (-Not (Test-Path -Path "HKLM:\SYSTEM\CurrentControlSet\Services\SSPORT")) {
				$Destination = ($env:SystemRoot + "\System32\drivers\SSPORT.cat")
				If (Test-Path $Destination){
					Set-Owner -Path $Destination -Account 'Administrators'
					$Acls = Get-Acl $Destination
					$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule("BUILTIN\Administrators", "FullControl", "Allow")
					$Acls.SetAccessRule($Ar)
					ForEach ($Ar in $Acls.Access) {
						# Remove any deny rules
						If ($Ar.AccessControlType -eq "Deny") {
							$Acls.RemoveAccessRule($Ar)
						}
					}
					Set-Acl $Destination $Acls -ErrorAction SilentlyContinue
					Write-host ("`tTrying to remove: " + $Destination)
					Remove-Item -Force -Path $Destination -ErrorAction SilentlyContinue
					If(Test-Path -Path $Destination) {
						Move-OnReboot -Path $Destination
						$RebootRequired = $True
					}
				}
				$Destination = ($env:SystemRoot + "\System32\CatRoot\{F750E6C3-38EE-11D1-85E5-00C04FC295EE}\SSPORT.cat")
				If (Test-Path $Destination){
					Set-Owner -Path $Destination -Account 'Administrators'
					$Acls = Get-Acl $Destination
					$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule("BUILTIN\Administrators", "FullControl", "Allow")
					$Acls.SetAccessRule($Ar)
					ForEach ($Ar in $Acls.Access) {
						# Remove any deny rules
						If ($Ar.AccessControlType -eq "Deny") {
							$Acls.RemoveAccessRule($Ar)
						}
					}
					Set-Acl $Destination $Acls -ErrorAction SilentlyContinue
					Write-host ("`tTrying to remove: " + $Destination)
					Remove-Item -Force -Path $Destination -ErrorAction SilentlyContinue
					If(Test-Path -Path $Destination) {
						Move-OnReboot -Path $Destination
						$RebootRequired = $True
					}
				}
				$Destination = ($env:SystemRoot + "\System32\drivers\SSPORT.sys")
				If (Test-Path $Destination){
					Set-Owner -Path $Destination -Account 'Administrators'
					$Acls = Get-Acl $Destination
					$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule("BUILTIN\Administrators", "FullControl", "Allow")
					$Acls.SetAccessRule($Ar)
					ForEach ($Ar in $Acls.Access) {
						# Remove any deny rules
						If ($Ar.AccessControlType -eq "Deny") {
							$Acls.RemoveAccessRule($Ar)
						}
					}
					Set-Acl $Destination $Acls -ErrorAction SilentlyContinue	
					Write-host ("`tTrying to remove: " + $Destination) 
					Remove-Item -Force -Path $Destination -ErrorAction SilentlyContinue
					If(Test-Path -Path $Destination) {
						Move-OnReboot -Path $Destination
						$RebootRequired = $True
					}
				}
			}
		#}
	}
	#endregion SSPORT Driver
	#region SigPlus.ocx
	If (Test-Path -Path ($env:SystemRoot + "\SysWOW64\SigPlus.ocx")){
		& ($env:SystemRoot + "\SysWOW64\regsvr32.exe" + " /s /u " + ($env:SystemRoot + "\SysWOW64\SigPlus.ocx"))
		$Destination = ($env:SystemRoot + "\SysWOW64\SigPlus.ocx")
		If (Test-Path $Destination){
			Set-Owner -Path $Destination -Account 'Administrators'
			$Acls = Get-Acl $Destination
			$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule("BUILTIN\Administrators", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
			$Acls.SetAccessRule($Ar)
			ForEach ($Ar in $Acls.Access) {
				# Remove any deny rules
				If ($Ar.AccessControlType -eq "Deny") {
					$Acls.RemoveAccessRule($Ar)
				}
			}
			Set-Acl $Destination $Acls -ErrorAction SilentlyContinue	
			Write-host ("`tTrying to remove: " + $Destination)
			Remove-Item -Force -Path $Destination -ErrorAction SilentlyContinue
			If(Test-Path -Path $Destination) {
				Move-OnReboot -Path $Destination
				$RebootRequired = $True
			}
		}
	}
	If (Test-Path -Path ($env:SystemRoot + "\SysWOW64\SigSign.ocx")){
		& ($env:SystemRoot + "\SysWOW64\regsvr32.exe" + " /s /u " + ($env:SystemRoot + "\SysWOW64\SigSign.ocx"))
		$Destination = ($env:SystemRoot + "\SysWOW64\SigSign.ocx")
		If (Test-Path $Destination){
			Set-Owner -Path $Destination -Account 'Administrators'
			$Acls = Get-Acl $Destination
			$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule("BUILTIN\Administrators", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
			$Acls.SetAccessRule($Ar)
			ForEach ($Ar in $Acls.Access) {
				# Remove any deny rules
				If ($Ar.AccessControlType -eq "Deny") {
					$Acls.RemoveAccessRule($Ar)
				}
			}
			Set-Acl $Destination $Acls -ErrorAction SilentlyContinue
			Write-host ("`tTrying to remove: " + $Destination)
			Remove-Item -Force -Path $Destination -ErrorAction SilentlyContinue
			If(Test-Path -Path $Destination) {
				Move-OnReboot -Path $Destination
				$RebootRequired = $True
			}
		}
	}
	#endregion SigPlus.ocx
	#region isusweb.dll removal
		If (Test-Path -Path ($env:SystemRoot + "\Downloaded Program Files\isusweb.dll")){
		& ($env:SystemRoot + "\SysWOW64\regsvr32.exe" + " /s /u " + ($env:SystemRoot + "\Downloaded Program Files\isusweb.dll"))
		$Destination = ($env:SystemRoot + "\Downloaded Program Files\isusweb.dll")
		If (Test-Path $Destination){
			Set-Owner -Path $Destination -Account 'Administrators'
			$Acls = Get-Acl $Destination
			$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule("BUILTIN\Administrators", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
			$Acls.SetAccessRule($Ar)
			ForEach ($Ar in $Acls.Access) {
				# Remove any deny rules
				If ($Ar.AccessControlType -eq "Deny") {
					$Acls.RemoveAccessRule($Ar)
				}
			}
			Set-Acl $Destination $Acls -ErrorAction SilentlyContinue
			Write-host ("`tTrying to remove: " + $Destination)
			Remove-Item -Force -Path $Destination -ErrorAction SilentlyContinue
			If(Test-Path -Path $Destination) {
				Move-OnReboot -Path $Destination
				$RebootRequired = $True
			}
		}
	}
	#endregion isusweb.dll removal
	#region CVE-2017-8529
	If ((Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Internet Explorer\Main\FeatureControl" -ErrorAction SilentlyContinue)."FEATURE_ENABLE_PRINT_INFO_DISCLOSURE_FIX" -ne 1 ) {
		Write-Host ("Fix CVE-2017-8529")
		Set-Reg ("HKLM:\SOFTWARE\Microsoft\Internet Explorer\Main\FeatureControl") "FEATURE_ENABLE_PRINT_INFO_DISCLOSURE_FIX" 1 "DWORD"

	}
	If ((Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Internet Explorer\Main\FeatureControl" -ErrorAction SilentlyContinue)."FEATURE_ENABLE_PRINT_INFO_DISCLOSURE_FIX" -ne 1 ) {
		Write-Host ("Fix CVE-2017-8529 X64")
		Set-Reg ("HKLM:\SOFTWARE\WOW6432Node\Microsoft\Internet Explorer\Main\FeatureControl") "FEATURE_ENABLE_PRINT_INFO_DISCLOSURE_FIX" 1 "DWORD"

	}
	#endregion CVE-2017-8529
	#region Windows Speculative Execution
	If ((Get-ItemProperty -Path "HKLM:SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management" -ErrorAction SilentlyContinue)."FeatureSettingsOverride" -ne 72 ) {
		Write-Host ("Windows Speculative Execution - FeatureSettingsOverride")
		Set-Reg ("HKLM:SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management") "FeatureSettingsOverride" 72 "DWORD"

	}
	If ((Get-ItemProperty -Path "HKLM:SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management" -ErrorAction SilentlyContinue)."FeatureSettingsOverrideMask" -ne 3 ) {
		Write-Host ("Windows Speculative Execution - FeatureSettingsOverrideMask")
		Set-Reg ("HKLM:SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management") "FeatureSettingsOverrideMask" 3 "DWORD"

	}
	#endregion Windows Speculative Execution
	#region Unquoted Services
	#Get All services that don't start with a double quote and have a space before the exe
	$Services = Get-ChildItem -Path "HKLM:\SYSTEM\CurrentControlSet\Services" | Where-Object {$_.GetValue("ImagePath") -match '^[^\"].*\s.*\.exe.*'}
	If($Services.Count -gt 0) {
		Write-Host (" ")
		Write-Host ("Fixing Services:")
		foreach ($SP in $Services) {
			$SPV = $SP.GetValue("ImagePath")
			Write-Host ("`t" +$SP.PSChildName + ":`n`t`tFrom:`t"+ $SPV + "`n`t`tTo:`t" + ($SPV -replace '^','"' -replace '\.exe','.exe"'))
			#Set fix ImagePath value with quotes before and after the exe path.
			Set-ItemProperty -Path $sp.PSPath -Name "ImagePath" -Value ($SPV -replace '^','"' -replace '\.exe','.exe"')
		}
		Write-Host (" ")
	}
	#endregion Unquoted Services
	#region remove appx
	Write-Host "Removing vulnerable Appx packages"
	$AppxPackages = Get-AppxPackage -AllUsers
	If (get-command Get-ProvisionedAppPackage -ErrorAction SilentlyContinue) {
		$ProvisionedAppx = Get-ProvisionedAppPackage -Online
	}
	ForEach ($RA in $RemoveAppx) {
		Foreach ($Package in ($AppxPackages.Where({$_.Name -match $RA}))) {
			write-host ("`tRemoving: " +  $Package.Name)
			Try {
				Remove-AppxPackage -Package $Package.PackageFullName -AllUsers | Out-Null
			}Catch {
				Remove-AppxPackage -Package $Package.PackageFullName | Out-Null
			}
		}
		If (get-command Get-ProvisionedAppPackage -ErrorAction SilentlyContinue) {
			Foreach ($Package in ($ProvisionedAppx.Where({$_.DisplayName -match $RA}))) {
				write-host ("`tRemoving: " +  $Package.DisplayName)
				Try{			
					$Package | Remove-ProvisionedAppxPackage -Online -AllUsers | Out-Null
				}Catch{
					$Package | Remove-ProvisionedAppxPackage -Online | Out-Null
				}
			}
		}	
		If (Test-Path -Path ($env:ProgramFiles + "\WindowsApps\" + $RA + "*")) {
			ForEach ($Folder in ((Get-ChildItem -Path $env:ProgramFiles + "\WindowsApps\").Where({$_.Name -match $RA}))) {
					Set-Owner -Path $Folder.FullName -Account 'Administrators'
					$Acls = Get-Acl $Folder.FullName
					$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule("BUILTIN\Administrators", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
					$Acls.SetAccessRule($Ar)
					Set-Acl $Folder.FullName $Acls -ErrorAction SilentlyContinue
					Write-host ("`tTrying to remove: " + $Folder.FullName)
					Remove-Item -Force -Recurse -ErrorAction SilentlyContinue -Path $Folder.FullName 
					If(Test-Path -Path $Folder.FullName) {
						Move-OnReboot -Path $Folder.FullName
						$RebootRequired = $True
					}
			}
		}
	}
	#endregion remove appx
	#region Remove Adobe Flash
		#region Prevent Adobe Flash Player from running
		If ((Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Internet Explorer\ActiveX Compatibility\{D27CDB6E-AE6D-11CF-96B8-444553540000}" -ErrorAction SilentlyContinue)."Compatibility Flags" -ne 1024 -and (Get-ItemProperty -Path "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Internet Explorer\ActiveX Compatibility\{D27CDB6E-AE6D-11CF-96B8-444553540000}" -ErrorAction SilentlyContinue)."Compatibility Flags" -ne 1024 -and (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office\Common\COM\Compatibility\{D27CDB6E-AE6D-11CF-96B8-444553540000}" -ErrorAction SilentlyContinue)."Compatibility Flags" -ne 1024) {
			Write-Host ("Prevent Adobe Flash Player from running")
			Set-Reg ("HKLM:\SOFTWARE\Microsoft\Internet Explorer\ActiveX Compatibility\{D27CDB6E-AE6D-11CF-96B8-444553540000}") "Compatibility Flags" 1024 "DWORD"
			Set-Reg ("HKLM:\SOFTWARE\Wow6432Node\Microsoft\Internet Explorer\ActiveX Compatibility\{D27CDB6E-AE6D-11CF-96B8-444553540000}") "Compatibility Flags" 1024 "DWORD"
			Set-Reg ("HKLM:\SOFTWARE\Microsoft\Office\Common\COM\Compatibility\{D27CDB6E-AE6D-11CF-96B8-444553540000}") "Compatibility Flags" 1024 "DWORD"
		}
		#endregion Prevent Adobe Flash Player from running
		#region Remove by Registry CR
		$RemoveAppName = "Adobe Flash"
		Write-host "Trying to uninstall any installed $RemoveAppName"
		Try{
			Write-host "`t with Registry"
			$Uninstall = Get-ChildItem -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
			$Uninstall += Get-ChildItem -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
			Foreach ($app in ($Uninstall.Where({ $_.GetValue("DisplayName") -match $RemoveAppName}))) {
				Invoke-Expression ("& " + $app.GetValue("UninstallString"))
			}
		}Catch{

		}
		#endregion Remove by Registry CR
		#region Remove by Windows Installer
		Try {
			Write-host "`t with WindowsInstaller"
			$Installer = New-Object -ComObject WindowsInstaller.Installer; 
			$InstalledProducts = ForEach($Product in ($Installer.ProductsEx("", "", 7))){
				[PSCustomObject]@{
				ProductCode = $Product.ProductCode(); 
				LocalPackage = $Product.InstallProperty("LocalPackage"); 
				VersionString = $Product.InstallProperty("VersionString"); 
				ProductPath = $Product.InstallProperty("ProductName")}
			} 
			ForEach ($WIRA in ($InstalledProducts.Where({$_.ProductPath -match $RemoveAppName }))) {
				Try {         
					Write-host ("`t`t" + $WIRA.ProductPath + " with MsiExec /X" + $WIRA.ProductCode)
					#MsiExec.exe /X"{${$WIRA.ProductCode}}" /quiet
					Start-Process -FilePath ($env:SystemRoot + "\System32\MsiExec.exe") -ArgumentList ('/X' + $WIRA.ProductCode + ' /qn') -NoNewWindow -Wait
				}Catch{
					
				}
			}
		}Catch{
			Write-host "`t with WMI"
			Get-WmiObject -Class Win32_Product -ErrorAction SilentlyContinue | Where-Object Name -Match $RemoveAppName | Foreach-Object { Write-Host ("Removing: " + $_.Name) ;$_.Uninstall()} | Out-Null
		}
		#endregion Remove by Windows Installer
		#region Remove by DISM
		Try {
			Write-host "`t with dism"
			dism /online /get-packages | Where-Object {$_ -match "Package Identity :" -and $_ -match "Adobe-Flash-For-Windows"} | ForEach-Object {
				Write-host ("`tRemoving Package: " + ($_ -replace "Package Identity :",""))
				dism /online /remove-package /packagename:($_ -replace "Package Identity :","")
			}
		}Catch{
			dism /online /remove-package /packagename:Adobe-Flash-For-Windows-Package~31bf3856ad364e35~amd64-10.0.17134.1
			dism /online /remove-package /packagename:Adobe-Flash-For-Windows-WOW64-Package~31bf3856ad364e35~amd64-10.0.17134.1
			dism /online /remove-package /packagename:Adobe-Flash-For-Windows-onecoreuap-Package~31bf3856ad364e35~amd64-10.0.17134.1
		}
		#endregion Remove by DISM
		#region Remove Flash Update Service
		Write-host "`tTrying to remove Flash update Service"
		Try {
			Get-Service | Where-Object {$_.Name -eq "AdobeFlashPlayerUpdateSvc"} | Remove-Service | Out-Null
		}Catch {
			#Start-Process -FilePath ($env:SystemRoot + "\System32\cmd.exe") -ArgumentList "/C sc STOP AdobeFlashPlayerUpdateSvc && sc DELETE AdobeFlashPlayerUpdateSvc" -NoNewWindow -PassThur
			#Start-Process -FilePath ($env:SystemRoot + "\System32\sc.exe") -ArgumentList "STOP AdobeFlashPlayerUpdateSvc" -NoNewWindow  | Out-Null
			#Start-Process -FilePath ($env:SystemRoot + "\System32\sc.exe") -ArgumentList "DELETE AdobeFlashPlayerUpdateSvc" -NoNewWindow -PassThru | Out-Null
		}
		#endregion Remove Flash Update Service
		#region Remove Flash Files
		Write-host "`tTrying to remove flash program Files"
		$Destination = ($env:SystemRoot + "\system32\Macromed")
		If (Test-Path -Path $Destination) {
			ForEach ($File in (Get-ChildItem -Path $Destination -Recurse  -File | Sort-Object -Property FullName -Descending)) {
				Write-Host ("`t`tTrying to remove: " + $File.FullName)
				Set-Owner -Path $File.FullName -Account 'Administrators'
				$Acls = Get-Acl $File.FullName
				$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule("BUILTIN\Administrators", "FullControl", "Allow")
				$Acls.SetAccessRule($Ar) 
				ForEach ($Ar in $Acls.Access) {
					# Remove any deny rules
					If ($Ar.AccessControlType -eq "Deny") {
						$Acls.RemoveAccessRule($Ar)
					}
				}
				Set-Acl $File.FullName $Acls -ErrorAction SilentlyContinue
				Write-host ("`t`tTrying to remove: " + $File.FullName)
				Remove-Item -Force -Path $File.FullName -ErrorAction SilentlyContinue
				If(Test-Path -Path $File.FullName) {
					Move-OnReboot -Path $File.FullName | Out-Null
					$RebootRequired = $True
				}
			}
		}
		$Destination = ($env:SystemRoot + "\SysWOW64\Macromed")
		If (Test-Path -Path $Destination) {
			ForEach ($File in (Get-ChildItem -Path $Destination -Recurse  -File | Sort-Object -Property FullName -Descending)) {
				Write-Host ("`t`tTrying to remove: " + $File.FullName)
				Set-Owner -Path $File.FullName -Account 'Administrators'
				$Acls = Get-Acl $File.FullName
				$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule("BUILTIN\Administrators", "FullControl", "Allow")
				$Acls.SetAccessRule($Ar) 
				ForEach ($Ar in $Acls.Access) {
					# Remove any deny rules
					If ($Ar.AccessControlType -eq "Deny") {
						$Acls.RemoveAccessRule($Ar)
					}
				}
				Set-Acl $File.FullName $Acls -ErrorAction SilentlyContinue
				Write-host ("`t`tTrying to remove: " + $File.FullName)
				Remove-Item -Force -Path $File.FullName -ErrorAction SilentlyContinue
				If(Test-Path -Path $File.FullName) {
					Move-OnReboot -Path $File.FullName | Out-Null
					$RebootRequired = $True
				}
			}
		}
		#endregion Remove Flash Files
		#region remove flash from users profiles
		$Profiles = (Get-ChildItem -Path ($env:SystemDrive + "\Users")).FullName
		#remove flash from user install
		ForEach ($Profile in $Profiles) {
			$Destination = ($Profile + "AppData\Roaming\Adobe\Flash Player")
			If (Test-Path -Path $Destination) {
				ForEach ($File in (Get-ChildItem -Path $Destination -Recurse  -File | Sort-Object -Property FullName -Descending)) {
					Write-Host ("`tTrying to remove: " + $File.FullName)
					Set-Owner -Path $File.FullName -Account 'Administrators'
					$Acls = Get-Acl $File.FullName
					$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule("BUILTIN\Administrators", "FullControl", "Allow")
					$Acls.SetAccessRule($Ar) 
					Set-Acl $File.FullName $Acls -ErrorAction SilentlyContinue
					Write-host ("`tTrying to remove: " + $File.FullName)
					Remove-Item -Force -Path $File.FullName -ErrorAction SilentlyContinue
					If(Test-Path -Path $File.FullName) {
						Move-OnReboot -Path $File.FullName
						$RebootRequired = $True
					}
					
				}
			}
			$Destination = ($Profile + "AppData\Roaming\Macromedia")
			If (Test-Path -Path $Destination) {
				If (Test-Path -Path $Destination) {
					ForEach ($File in (Get-ChildItem -Path $Destination -Recurse  -File | Sort-Object -Property FullName -Descending)) {
						Write-Host ("`tTrying to remove: " + $File.FullName)
						Set-Owner -Path $File.FullName -Account 'Administrators'
						$Acls = Get-Acl $File.FullName
						$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule("BUILTIN\Administrators", "FullControl", "Allow")
						$Acls.SetAccessRule($Ar) 
						Set-Acl $File.FullName $Acls -ErrorAction SilentlyContinue
						Write-host ("`tTrying to remove: " + $File.FullName)
						Remove-Item -Force -Path $File.FullName -ErrorAction SilentlyContinue
						If(Test-Path -Path $File.FullName) {
							Move-OnReboot -Path $File.FullName
							$RebootRequired = $True
						}
						
					}
				}
			}
		}
		#endregion remove flash from users profiles
	#endregion Remove Adobe Flash
	#region Remove Log4j from Crystal Reports
	Write-Host ("Crystal Reports Viewer log4j cleanup")
	If (Test-Path -Path ($env:ProgramFiles + "\Crystal Reports Viewer\")) {
		ForEach ($File in (Get-ChildItem -Recurse -File -Path ($env:ProgramFiles + "\Crystal Reports Viewer") -Include "log4j*.jar")) {
			Write-Host ("`tRemoving: " + $File.FullName)
			Set-Owner -Path $File.FullName -Account 'Administrators'
			$Acls = Get-Acl $File.FullName
			$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule("BUILTIN\Administrators", "FullControl", "Allow")
			$Acls.SetAccessRule($Ar)
			Set-Acl $File.FullName $Acls -ErrorAction SilentlyContinue
			Write-host ("`tTrying to remove: " + $File.FullName)
			Remove-Item -Force -Recurse -ErrorAction SilentlyContinue -Path $File.FullName 
			If(Test-Path -Path $File.FullName) {
				Move-OnReboot -Path $File.FullName
				$RebootRequired = $True
			}
		}
	}
	If (Test-Path -Path (${env:ProgramFiles(x86)} + "\Crystal Reports Viewer")) {	
		ForEach ($File in (Get-ChildItem -Recurse -File -Path (${env:ProgramFiles(x86)} + "\Crystal Reports Viewer") -Include "log4j*.jar")) {
			Write-Host ("`tRemoving: " + $File.FullName)
			Set-Owner -Path $File.FullName -Account 'Administrators'
			$Acls = Get-Acl $File.FullName
			$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule("BUILTIN\Administrators", "FullControl", "Allow")
			$Acls.SetAccessRule($Ar)
			Set-Acl $File.FullName $Acls -ErrorAction SilentlyContinue
			Write-host ("`tTrying to remove: " + $File.FullName)
			Remove-Item -Force -Recurse -ErrorAction SilentlyContinue -Path $File.FullName 
			If(Test-Path -Path $File.FullName) {
				Move-OnReboot -Path $File.FullName
				$RebootRequired = $True
			}
		}
	}
	#endregion Remove Log4j from Crystal Reports
	#region Remove Log4j from customapp2
	Write-Host ("customapp2 log4j cleanup")
	If (Test-Path -Path ($env:systemdrive + "\customapp2")) {
		ForEach ($File in (Get-ChildItem -Recurse -File -Path ($env:systemdrive + "\customapp2") -Include "log4j*.jar")) {
			Write-Host ("`tRemoving: " + $File.FullName)
			If (Get-Command -Name Set-Owner -ErrorAction SilentlyContinue) {
				Set-Owner -Path $File.FullName -Account 'Administrators'
			}Else{
				$ACL = Get-ACL $File.FullName
				$Group = New-Object System.Security.Principal.NTAccount("Builtin", "Administrators")
				$ACL.SetOwner($Group)
				Set-Acl -Path $File.FullName -AclObject $ACL	
			}
			$Acls = Get-Acl $File.FullName
			$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule("BUILTIN\Administrators", "FullControl", "Allow")
			$Acls.SetAccessRule($Ar)
			Set-Acl $File.FullName $Acls -ErrorAction SilentlyContinue
			Write-host ("`tTrying to remove: " + $File.FullName)
			Remove-Item -Force -Recurse -ErrorAction SilentlyContinue -Path $File.FullName 
			If(Test-Path -Path $File.FullName) {
				Move-OnReboot -Path $File.FullName
				$RebootRequired = $True
			}
		}
	}
	#endregion Remove Log4j from customapp2
	#region Remove Log4j from vendorapp
	Write-Host ("vendorapp log4j cleanup")
	If (Test-Path -Path ($env:ProgramFiles + "\vendorapp\thirdparty")) {
		ForEach ($File in (Get-ChildItem -Recurse -File -Path ($env:ProgramFiles + "\vendorapp\thirdparty") -Include "log4j*.jar")) {
			Write-Host ("`tRemoving: " + $File.FullName)
			If (Get-Command -Name Set-Owner -ErrorAction SilentlyContinue) {
				Set-Owner -Path $File.FullName -Account 'Administrators'
			}Else{
				$ACL = Get-ACL $File.FullName
				$Group = New-Object System.Security.Principal.NTAccount("Builtin", "Administrators")
				$ACL.SetOwner($Group)
				Set-Acl -Path $File.FullName -AclObject $ACL	
			}
			$Acls = Get-Acl $File.FullName
			$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule("BUILTIN\Administrators", "FullControl", "Allow")
			$Acls.SetAccessRule($Ar)
			Set-Acl $File.FullName $Acls -ErrorAction SilentlyContinue
			Write-host ("`tTrying to remove: " + $File.FullName)
			Remove-Item -Force -Recurse -ErrorAction SilentlyContinue -Path $File.FullName 
			If(Test-Path -Path $File.FullName) {
				Move-OnReboot -Path $File.FullName
				$RebootRequired = $True
			}
		}
	}
	#endregion Remove Log4j from vendorapp
	#region Enable MS Office Automatic Updates
	If ((Get-ItemProperty -Path "HKLM:\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate" -ErrorAction SilentlyContinue)."EnableAutomaticUpdates" -ne 1 ) {
		Write-Host ("Enabling MS Office Automatic Updates")	
		Set-Reg ("HKLM:\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate") "EnableAutomaticUpdates" 1 "DWORD"
	}	
	#endregion Enable MS Office Automatic Updates
	#region Force office update
	If ($NoOfficeUpdate = $false) {
		If (Test-Path -Path ($env:ProgramFiles + "\Common Files\microsoft shared\ClickToRun\OfficeC2RClient.exe")) {
			Start-Process -FilePath ($env:ProgramFiles + "\Common Files\microsoft shared\ClickToRun\OfficeC2RClient.exe") -ArgumentList "/update user forceappshutdown=true displaylevel=true"
		}
	}
	#endregion Force office update
	#region install Microsoft Updates
	If ( $NoMSUpdate = $false) {
		If (Get-Module -ListAvailable -Name "PSWindowsUpdate") {
			If (-Not (Get-Module "PSWindowsUpdate" -ErrorAction SilentlyContinue)) {
				Import-Module "PSWindowsUpdate"
			}
		}Else{
			[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
			Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
			
			Install-PackageProvider -Confirm:$False -Force -Scope AllUsers -Name NuGet -MinimumVersion 2.8.5.201 
			
			Install-Module -Force -Name PSWindowsUpdate -scope AllUsers
			
			Import-Module PSWindowsUpdate
		} 
		Get-WindowsUpdate -MicrosoftUpdate -Install -AcceptAll  -IgnoreReboot
	}
	#endregion install Microsoft Updates
	#region Fix Insecure Windows Services Permissions
	 Write-Host ("Fixing Insecure Windows Services Permissions")
		If (Get-Command "Get-CimInstance" -ErrorAction SilentlyContinue) {
			$Services = Get-CimInstance Win32_Service
		}ElseIf(Get-Command "Get-WmiObject" -ErrorAction SilentlyContinue){
			$Services = Get-WmiObject Win32_Service
		}
		ForEach ($Service in $Services) {
			If ($Service.PathName -ne $null) {
				$exepath = ((($Service.PathName -split ".exe")[0] + '.exe') -replace '"')
				$exeFolder = (Split-Path -Parent -Path $exepath)
				$RemoveACLS = @()
				If(Test-Path -path $exeFolder) {
					# Write-Host ("`tTesting Service: " + $Service.Caption)
					$ServiceACLs = Get-Acl -Path $exeFolder
					$ServiceACLsModifyed = $False
					$RemoveACL = $false
					ForEach ($ServiceACL in $ServiceACLs.Access) {
						ForEach ($RGFS in $RemoveGroupsFromServices) {
							If ($ServiceACL.IdentityReference -match $RGFS) {
								ForEach ($RPFS in $RemovePermissionFromServices) {
									If ($ServiceACL.FileSystemRights -match $RPFS) {
										write-host ("`tInsecure Permissions: "	+ $Service.Caption)
										$ServiceACLsModifyed = $True
										$ServiceACLs.RemoveAccessRule($ServiceACL) | Out-Null
									}
								}
							}
						}
					}
					If ($ServiceACLsModifyed) {
						Write-Host ("`t`tFixing Service: " + $Service.Caption)
						Set-Acl $exeFolder $ServiceACLs
					}
				}
			}
		}
		Write-Host ""
		Write-Host ""
	#endregion Fix Insecure Windows Services Permissions
	#region Microsoft Defender Definition Update
	If (Test-Path -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender"){

		Write-host ("Windows Defender is disabled by Group Policy, enabling it")
		Remove-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender" -Recurse -Force -ErrorAction SilentlyContinue
	}
	If (Test-Path -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender Security Center"){
		Write-host ("Windows Defender Security Center is disabled by Group Policy, enabling it")
		Remove-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender Security Center" -Recurse -Force -ErrorAction SilentlyContinue
	}
	$WinDefendService = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\WinDefend" -Name "Start" -ErrorAction SilentlyContinue
	If ($WinDefendService.Start -eq 4){
		Write-host ("Windows Defender service is disabled, enabling it")
		Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\WinDefend" -Name "Start" -Value 2
		Start-Service -Name WinDefend -ErrorAction SilentlyContinue
	}
	$WinWdNisSvcService = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\WdNisSvc" -Name "Start" -ErrorAction SilentlyContinue
	If ($WinWdNisSvcService.Start -eq 4){
		Write-host ("Windows Defender Network Inspection service is disabled, enabling it")
		Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\WdNisSvc" -Name "Start" -Value 2
		Start-Service -Name WdNisSvc -ErrorAction SilentlyContinue
	}
	$SenceService = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\Sense" -Name "Start" -ErrorAction SilentlyContinue
	If ($SenceService.Start -eq 4){
		Write-host ("Windows Defender Antivirus Network Inspection service is disabled, enabling it")
		Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\Sense" -Name "Start" -Value 2
		Start-Service -Name Sense -ErrorAction SilentlyContinue
	}
	
	If (Test-Path ($env:SystemRoot + "\System32\MpSigStub.exe")){
		Write-host ("Attempting to update Windows Defender")
		If (Test-Path ($env:ProgramFiles + "\Windows Defender\MpCmdRun.exe")){
			Try{
				Start-Process -FilePath ($env:ProgramFiles + "\Windows Defender\MpCmdRun.exe") -ArgumentList '-SignatureUpdate -http' -NoNewWindow -Wait
			}Catch {
			}

			$FileInfo = Get-Item -Path ($env:SystemRoot + "\System32\MpSigStub.exe")
			$FileVersion = [Version]$FileInfo.VersionInfo.ProductVersion
			If ($FileVersion -lt [Version]"1.1.16638.0"){
				Write-host ("Attempting to update Windows Defender from version: " + $FileInfo.VersionInfo.ProductVersion)
				If (Test-Path -Path ($env:ProgramFiles + "\Windows Defender\MpCmdRun.exe")){
				Start-Process -FilePath ($env:ProgramFiles + "\Windows Defender\MpCmdRun.exe") -ArgumentList '-removedefinitions -dynamicsignatures' -NoNewWindow -Wait
				Start-Process -FilePath ($env:ProgramFiles + "\Windows Defender\MpCmdRun.exe") -ArgumentList '-SignatureUpdate -http' -NoNewWindow -Wait
				}
				If (get-command "Update-MpSignature" -ErrorAction SilentlyContinue) {
					Update-MpSignature
				}
			}
			If (Test-Path -Path ($env:ProgramFiles + "\Windows Defender\MpCmdRun.exe")){
				Start-Process -FilePath ($env:ProgramFiles + "\Windows Defender\MpCmdRun.exe") -ArgumentList '-SignatureUpdate -MMPC' -NoNewWindow -Wait
			}
			$FileInfo = Get-Item -Path ($env:SystemRoot + "\System32\MpSigStub.exe")
			If ($FileInfo.VersionInfo.ProductVersion -lt [Version]"1.1.16638.0"){
				Invoke-WebRequest -Uri $WindowsDefenderUpdateHTTP -OutFile ($env:temp + "\" + (Split-Path -leaf -Path $WindowsDefenderUpdateHTTP))
				If (Test-Path -path ($env:temp + "\" + (Split-Path -leaf -Path $WindowsDefenderUpdateHTTP))) {
					Start-Process -FilePath ($env:temp + "\" + (Split-Path -leaf -Path $WindowsDefenderUpdateHTTP)) -NoNewWindow -Wait
					Start-Sleep -Milliseconds 5000
					Write-host ("`tTrying to remove: " + ($env:temp + "\" + (Split-Path -leaf -Path $WindowsDefenderUpdateHTTP)))
					Remove-Item -Path ($env:temp + "\" + (Split-Path -leaf -Path $WindowsDefenderUpdateHTTP)) -Force -ErrorAction SilentlyContinue
				}
			}
		}Else {
			Invoke-WebRequest -Uri $WindowsDefenderUpdateHTTP -OutFile ($env:temp + "\" + (Split-Path -leaf -Path $WindowsDefenderUpdateHTTP))
			If (Test-Path -path ($env:temp + "\" + (Split-Path -leaf -Path $WindowsDefenderUpdateHTTP))) {
				Start-Process -FilePath ($env:temp + "\" + (Split-Path -leaf -Path $WindowsDefenderUpdateHTTP)) -NoNewWindow -Wait
				Start-Sleep -Milliseconds 5000
				Write-host ("`tTrying to remove: " + ($env:temp + "\" + (Split-Path -leaf -Path $WindowsDefenderUpdateHTTP)))
				Remove-Item -Path ($env:temp + "\" + (Split-Path -leaf -Path $WindowsDefenderUpdateHTTP)) -Force -ErrorAction SilentlyContinue
			}
		}
	}
	#endregion Microsoft Defender Definition Update
	#region Microsoft XML Parser (MSXML) and XML Core Services Unsupported
	If (Test-Path -Path ($env:SystemRoot + "\SysWOW64\msxml4.dll")) {
		If ($UpdateMSXML) {
			$FileInfo = Get-Item -Path ($env:SystemRoot + "\SysWOW64\msxml4.dll")
			$FileVersion = $FileInfo.VersionInfo.ProductVersion
			If ($FileVersion -lt [Version]"4.30.2117.0"){
				Write-host ("Attempting to update MSXML4 from version: " + $FileInfo.VersionInfo.ProductVersion)
				Invoke-WebRequest -Uri $MSXML4UpdateHTTP -OutFile ($env:temp + "\" + (Split-Path -leaf -Path $MSXML4UpdateHTTP))
				If (Test-Path -path ($env:temp + "\" + (Split-Path -leaf -Path $MSXML4UpdateHTTP))) {
					Start-Process -FilePath ($env:temp + "\" + (Split-Path -leaf -Path $MSXML4UpdateHTTP)) -NoNewWindow -Wait -ArgumentList "/qb /norestart"
					Start-Sleep -Milliseconds 5000
					Write-host ("`tTrying to remove: " + ($env:temp + "\" + (Split-Path -leaf -Path $MSXML4UpdateHTTP)))
					Remove-Item -Path ($env:temp + "\" + (Split-Path -leaf -Path $MSXML4UpdateHTTP)) -Force -ErrorAction SilentlyContinue
				}
			}
		}Else{
			Write-host ("Removing MSXML4")
			Remove-Item -Path ($env:SystemRoot + "\SysWOW64\msxml4.dll") -Force -ErrorAction SilentlyContinue
		}
	}
	#endregion Microsoft XML Parser (MSXML) and XML Core Services Unsupported
	#region Disable 3DES
		If (Test-Path -Path ("HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\Triple DES 168") -ErrorAction SilentlyContinue) {
			If ((Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\Triple DES 168" -Name "Enabled" -ErrorAction SilentlyContinue).Enabled -ne 0) {
				Write-Host ("Disabling Triple DES 168")
				Set-Reg -regPath ("HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\Triple DES 168") -name "Enabled" -value 0 -type "DWord" -comment "Disabling Triple DES 168"
			}
		}
	#endregion Disable 3DES
	#region Sysmon Update
	If (Test-Path -Path ($env:SystemRoot + "\Sysmon.exe")) {
		If ($SysmonUpdateHTTP) {
			$FileInfo = Get-Item -Path ($env:SystemRoot + "\Sysmon.exe")
			$FileVersion = [Version]$FileInfo.VersionInfo.ProductVersion
			If ($FileVersion -lt [Version]"14.13"){
				Write-host ("Attempting to update Sysmon from version: " + $FileInfo.VersionInfo.ProductVersion)
				Invoke-WebRequest -Uri $SysmonUpdateHTTP -OutFile ($env:temp + "\" + (Split-Path -leaf -Path $SysmonUpdateHTTP))
				If (Test-Path -path ($env:temp + "\" + (Split-Path -leaf -Path $SysmonUpdateHTTP))) {
					# If you want to extract only one file:
					$zip = [System.IO.Compression.ZipFile]::OpenRead(($env:temp + "\" + (Split-Path -leaf -Path $SysmonUpdateHTTP)))
					$entry = $zip.Entries | Where-Object { $_.Name -eq "Sysmon.exe" }
					if ($entry) {
						$destination = Join-Path $env:SystemRoot $entry.Name
						# $entry.ExtractToFile($destination, $true)
						[System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $destination, $true)
					}
					$zip.Dispose()
					start-sleep -Milliseconds 500
					Write-host ("`tTrying to remove: " + ($env:temp + "\" + (Split-Path -leaf -Path $SysmonUpdateHTTP)))
					Remove-Item -Path ($env:temp + "\" + (Split-Path -leaf -Path $SysmonUpdateHTTP)) -Force -ErrorAction SilentlyContinue
				}
			}
		}else{
			Write-host ("Removing Sysmon")
			Remove-Item -Path ($env:SystemRoot + "\Sysmon.exe") -Force -ErrorAction SilentlyContinue
		}
	}
	If (Test-Path -Path ($env:SystemRoot + "\Sysmon64.exe")) {
		If ($SysmonUpdateHTTP) {
			$FileInfo = Get-Item -Path ($env:SystemRoot + "\Sysmon64.exe")
			$FileVersion = [Version]$FileInfo.VersionInfo.ProductVersion
			If ($FileVersion -lt [Version]"14.13"){
				Write-host ("Attempting to update Sysmon64 from version: " + $FileInfo.VersionInfo.ProductVersion)
				Invoke-WebRequest -Uri $SysmonUpdateHTTP -OutFile ($env:temp + "\" + (Split-Path -leaf -Path $SysmonUpdateHTTP))
				If (Test-Path -path ($env:temp + "\" + (Split-Path -leaf -Path $SysmonUpdateHTTP))) {
					# If you want to extract only one file:
					$zip = [System.IO.Compression.ZipFile]::OpenRead(($env:temp + "\" + (Split-Path -leaf -Path $SysmonUpdateHTTP)))
					$entry = $zip.Entries | Where-Object { $_.Name -eq "Sysmon64.exe" }
					if ($entry) {
						$destination = Join-Path $env:SystemRoot $entry.Name
						# $entry.ExtractToFile($destination, $true)
						[System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $destination, $true)
					}
					$zip.Dispose()
					start-sleep -Milliseconds 500
					Write-host ("`tTrying to remove: " + ($env:temp + "\" + (Split-Path -leaf -Path $SysmonUpdateHTTP)))
					Remove-Item -Path ($env:temp + "\" + (Split-Path -leaf -Path $SysmonUpdateHTTP)) -Force -ErrorAction SilentlyContinue
				}
			}
		}else{
			Write-host ("Removing Sysmon")
			Remove-Item -Path ($env:SystemRoot + "\Sysmon64.exe") -Force -ErrorAction SilentlyContinue
		}
	}
	If (Test-Path -Path ($env:SystemRoot + "\SysmonDrv.sys")) {
			Write-host ("Removing SysmonDrv.sys")
			Remove-Item -Path ($env:SystemRoot + "\SysmonDrv.sys") -Force -ErrorAction SilentlyContinue
	}
	#endregion Sysmon Update
	#region UNC Path Hardening
		Write-Host ("Applying UNC Hardening Settings")

		Set-Reg ("HKLM:\Software\Policies\Microsoft\Windows\NetworkProvider\HardenedPaths") "\\*\NETLOGON" "RequireMutualAuthentication=1, RequireIntegrity=1, RequirePrivacy=1" "String"
		Set-Reg ("HKLM:\Software\Policies\Microsoft\Windows\NetworkProvider\HardenedPaths") "\\*\SYSVOL" "RequireMutualAuthentication=1, RequireIntegrity=1, RequirePrivacy=1" "String"
		Set-Reg ("HKLM:\Software\Policies\Microsoft\Windows\NetworkProvider\HardenedPaths") "\\Github.com\users" "RequireMutualAuthentication=1, RequireIntegrity=1, RequirePrivacy=1" "String"
		Set-Reg ("HKLM:\Software\Policies\Microsoft\Windows\NetworkProvider\HardenedPaths") "\\Github.com\share" "RequireMutualAuthentication=1, RequireIntegrity=1, RequirePrivacy=1" "String"
		Set-Reg ("HKLM:\Software\Policies\Microsoft\Windows\NetworkProvider\HardenedPaths") "\\Github.com\*" "RequireMutualAuthentication=1, RequireIntegrity=1" "String"

	#endregion UNC Path Hardening
	#region Disable ECDH public server param reuse in the Windows registry
		#CVE-2022-35838
		#CVE-2016-10957
		Write-Host "Applying ECDH public server param reuse fix"
		if(!(Test-Path "hklm:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\KeyExchangeAlgorithms\ECDH")) {
		New-Item "hklm:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\KeyExchangeAlgorithms" -name "ECDH"
		}

		New-ItemProperty "hklm:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\KeyExchangeAlgorithms\ECDH" -Name EphemKeyReuseTime -Value 0 -PropertyType DWord
	#endregion Disable ECDH public server param reuse in the Windows registry
	#End of Actions
	If ($RebootRequired) {
		Write-Warning " "
		Write-Warning " "
		Write-Warning " "
		Write-Warning "Reboot Required to remove files"
	}
	#Record Script Version
	if ($ScriptRecordkey -and $ScriptVersion -and $ScriptRecordDateValue -and $ScriptRecordVersionValue -and $FileDate) {
		write-host ("Recording " + $ScriptRecordVersionValue + ": " + $ScriptVersion + " in " + $ScriptRecordkey + " Key.") -foregroundcolor "Green"
		Set-Reg ("HKLM:\Software\" + $ScriptRecordkey) $ScriptRecordVersionValue  $ScriptVersion "String" ($ScriptRecordVersionValue + "=" + $ScriptVersion)
		write-host ("Recording " + $ScriptRecordDateValue  + ": " +  $FileDate + " in " + $ScriptRecordkey + " Key.") -foregroundcolor "Green"
		Set-Reg ("HKLM:\Software\" + $ScriptRecordkey) $ScriptRecordDateValue  $FileDate "String" ($ScriptRecordDateValue + "=" + $FileDate)
	}
}Else {
	Write-Host ("Script already ran on system.")
}
Write-Host ("")
Write-Host ("Script took: " + (FormatElapsedTime($sw.Elapsed)) + " to run.")
Write-Host ("")
If (-Not [string]::IsNullOrEmpty($LogFile)) {
	Stop-Transcript
}
Exit 0
