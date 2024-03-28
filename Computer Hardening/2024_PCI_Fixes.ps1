<#
.SYNOPSIS
    Apply Fixes for 2024 PCI scans
	Used to fix "Microsoft Windows Unquoted Service Path Enumeration" CVE-2013-1609 issue.
  
.DESCRIPTION
    - Used to fix "Microsoft Windows Unquoted Service Path Enumeration" CVE-2013-1609 issue.
	- WinVerifyTrust Signature Validation Vulnerability
.RELATED LINKS
	https://isc.sans.edu/diary/Help+eliminate+unquoted+path+vulnerabilities/14464
    http://cwe.mitre.org/data/definitions/428.html
    http://www.ryanandjeffshow.com/blog/2013/04/11/powershell-fixing-unquoted-service-paths-complete/
        
.PARAMETER 

.NOTES
  Changes:
    1.0.0 - Draft Script
    1.0.1 - Added more tries and different ways to run commands. Fixed logging path, force script to run as admin, fix permission to force removal of flash.
    1.0.2 - Fixed script errors and msiexec issue. 
    1.0.3 - Fixed script issue where WinTrust registry is not created. Fixed issue with flash service and files not being removed.
    1.0.4 - Found way to query installed packages without WMI. 
    1.0.5 - Added fixes for SSPORT.SYS, SigPlus.ocx and SigSign.ocx issues. Attempt Defender definition update. Added Move-OnReboot to fix deleting file issue.
    1.0.6 - Clean up of and reorganize of fixes.
    1.0.7 - Added Encryption Oracle Remediation.

	
  Release Date: 03/27/2024
   
  Author: Paul Fuller

.EXAMPLE
    .\2024_PCI_Fixes.ps1

#>
#Force Script to run as Admin 
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))  
{  
  $arguments = "& '" +$myinvocation.mycommand.definition + "'"
  $PS = Start-Process powershell -Verb runAs -ArgumentList $arguments
  Exit $PS
}
##Requires -RunAsAdministrator
$ScriptVersion = "1.0.7"
$LocalLogs = "C:\IT_Updates\Logs\"
$FileDate = (Get-Date -format yyyyMMdd-hhmm)
$sw = [Diagnostics.Stopwatch]::StartNew()
$CBS = @()
$Uninstall = @()
$RebootRequired = $False
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
        $Acl = Get-Acl (Split-Path -Path $LogFile -Parent)
        $Ar = New-Object system.Security.AccessControl.FileSystemAccessRule('Users', "FullControl", 'ContainerInherit,ObjectInherit', "None", "Allow")
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
#endregion LogFile

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

#region Fix WinVerifyTrust Signature Validation Vulnerability
If ((Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Cryptography\Wintrust\Config" -ErrorAction SilentlyContinue)."EnableCertPaddingCheck" -ne 1 -and (Get-ItemProperty -Path "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Cryptography\Wintrust\Config" -ErrorAction SilentlyContinue)."EnableCertPaddingCheck" -ne 1) {
	Write-Host ("WinVerifyTrust Signature Validation Vulnerability")
	Set-Reg ("HKLM:\SOFTWARE\Microsoft\Cryptography\Wintrust\Config") "EnableCertPaddingCheck" 1 "DWORD"
	Set-Reg ("HKLM:\SOFTWARE\Wow6432Node\Microsoft\Cryptography\Wintrust\Config") "EnableCertPaddingCheck" 1 "DWORD"
}
#endregion Fix WinVerifyTrust Signature Validation Vulnerability
#region Encryption Oracle Remediation
If ((Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Cryptography\Wintrust\Config" -ErrorAction SilentlyContinue)."EnableCertPaddingCheck" -le 1 ) {
	Write-Host ("Fixing Encryption Oracle Remediation")
	Set-Reg ("HKLM:\SOFTWARE\Microsoft\Cryptography\Wintrust\Config") "EnableCertPaddingCheck" 1 "DWORD"
}
#endregion Encryption Oracle Remediation
#region Remove Microsoft Sliverlight
	$RemoveAppName = "Microsoft Silverlight"
	#region Remove by Windows Installer
	Try {
		Write-host "Trying to uninstall any installed $RemoveAppName with WindowsInstaller"
		$Installer = New-Object -ComObject WindowsInstaller.Installer; 
		$InstalledProducts = ForEach($Product in ($Installer.ProductsEx("", "", 7))){
			[PSCustomObject]@{
			ProductCode = $Product.ProductCode(); 
			LocalPackage = $Product.InstallProperty("LocalPackage"); 
			VersionString = $Product.InstallProperty("VersionString"); 
			ProductPath = $Product.InstallProperty("ProductName")}
		} 
		ForEach ($WIRA in ($InstalledProducts.Where({$_.ProductPath -match $RemoveAppName}))) {
			Try {         
				Write-host ("Trying to uninstall any installed " + $WIRA.ProductPath + " with MsiExec /X" + $WIRA.ProductCode)
				#MsiExec.exe /X"{${$WIRA.ProductCode}}" /quiet
				Start-Process -FilePath ($env:SystemRoot + "\Sysetm32\MsiExec.exe") -ArgumentList ('/X' + $WIRA.ProductCode + ' /quiet') -NoNewWindow -Wait
			}Catch{
				
			}
		}
	}Catch{
		Write-host "Trying to uninstall any installed $RemoveAppName with WMI"
		Get-WmiObject -Class Win32_Product -ErrorAction SilentlyContinue | Where-Object Name -Match $RemoveAppName | Foreach-Object { Write-Host ("Removing: " + $_.Name) ;$_.Uninstall()}
	}
	#endregion Remove by Windows Installer

	Try {
		Write-host "Trying to uninstall any installed $RemoveAppName with Fallback"
		MsiExec.exe /X"{89F4137D-6C26-4A84-BDB8-2E5A4BB71E00}" /quiet
	}Catch{
		
	}
#endregion Remove Microsoft Sliverlight
#region SSPORT Driver
If (Test-Path ($env:SystemRoot + "\System32\drivers\SSPORT.SYS")){
	$FileInfo = Get-Item -Path ($env:SystemRoot + "\System32\drivers\SSPORT.SYS")
	If ($FileInfo.VersionInfo.ProductVersion -lt "1.0.0.1"){
		Write-Host "Removing SSPORT.SYS driver"
		If (Test-Path -Path "HKLM:\SYSTEM\CurrentControlSet\Services\SSPORT") {
			Remove-Item -Recurse -Force -Path "HKLM:\SYSTEM\CurrentControlSet\Services\SSPORT"
		}
		If (-Not (Test-Path -Path "HKLM:\SYSTEM\CurrentControlSet\Services\SSPORT")) {
			$Destination = ($env:SystemRoot + "\System32\drivers\SSPORT.cat")
			If (Test-Path $Destination){
				Set-Owner -Path $Destination -Account 'Administrators'
				$Acls = Get-Acl $Destination
				$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule("BUILTIN\Administrators", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
				$Acls.SetAccessRule($Ar)
				Set-Acl $Destination $Acl -ErrorAction SilentlyContinue
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
				$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule("BUILTIN\Administrators", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
				$Acls.SetAccessRule($Ar)
				Set-Acl $Destination $Acl -ErrorAction SilentlyContinue
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
				Set-Acl $Destination $Acl -ErrorAction SilentlyContinue	 
				Remove-Item -Force -Path $Destination -ErrorAction SilentlyContinue
				If(Test-Path -Path $Destination) {
					Move-OnReboot -Path $Destination
					$RebootRequired = $True
				}
			}
		}
	}
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
		Set-Acl $Destination $Acl -ErrorAction SilentlyContinue	
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
		Set-Acl $Destination $Acl -ErrorAction SilentlyContinue
		Remove-Item -Force -Path $Destination -ErrorAction SilentlyContinue
		If(Test-Path -Path $Destination) {
			Move-OnReboot -Path $Destination
			$RebootRequired = $True
		}
	}
}
#endregion SigPlus.ocx
#region Microsoft Defender Definition Update
If (Test-Path ($env:SystemRoot + "\System32\MpSigStub.exe")){
	$FileInfo = Get-Item -Path ($env:SystemRoot + "\System32\MpSigStub.exe")
	If ($FileInfo.VersionInfo.ProductVersion -lt "1.1.16638.0"){
		Write-host ("Attempting to update Windows Defender from version: " + $FileInfo.VersionInfo.ProductVersion)
		If (Test-Path -Path ($env:ProgramFiles + "\Windows Defender\MpCmdRun.exe")){
		Start-Process -FilePath ($env:ProgramFiles + "\Windows Defender\MpCmdRun.exe") -ArgumentList '-removedefinitions -dynamicsignatures' -NoNewWindow -Wait
		Start-Process -FilePath ($env:ProgramFiles + "\Windows Defender\MpCmdRun.exe") -ArgumentList '-SignatureUpdate' -NoNewWindow -Wait
		}
	}
}
#endregion Microsoft Defender Definition Update
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

#region Remove Adobe Flash
	#Prevent Adobe Flash Player from running
	If ((Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Internet Explorer\ActiveX Compatibility\{D27CDB6E-AE6D-11CF-96B8-444553540000}" -ErrorAction SilentlyContinue)."Compatibility Flags" -ne 1024 -and (Get-ItemProperty -Path "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Internet Explorer\ActiveX Compatibility\{D27CDB6E-AE6D-11CF-96B8-444553540000}" -ErrorAction SilentlyContinue)."Compatibility Flags" -ne 1024 -and (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office\Common\COM\Compatibility\{D27CDB6E-AE6D-11CF-96B8-444553540000}" -ErrorAction SilentlyContinue)."Compatibility Flags" -ne 1024) {
		Write-Host ("Prevent Adobe Flash Player from running")
		Set-Reg ("HKLM:\SOFTWARE\Microsoft\Internet Explorer\ActiveX Compatibility\{D27CDB6E-AE6D-11CF-96B8-444553540000}") "Compatibility Flags" 1024 "DWORD"
		Set-Reg ("HKLM:\SOFTWARE\Wow6432Node\Microsoft\Internet Explorer\ActiveX Compatibility\{D27CDB6E-AE6D-11CF-96B8-444553540000}") "Compatibility Flags" 1024 "DWORD"
		Set-Reg ("HKLM:\SOFTWARE\Microsoft\Office\Common\COM\Compatibility\{D27CDB6E-AE6D-11CF-96B8-444553540000}") "Compatibility Flags" 1024 "DWORD"
	}
	$RemoveAppName = "Adobe Flash"
	#region Remove by Registry CR
	Try{
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
		Write-host "Trying to uninstall any installed $RemoveAppName with WindowsInstaller"
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
				Write-host ("Trying to uninstall any installed " + $WIRA.ProductPath + " with MsiExec /X" + $WIRA.ProductCode)
				#MsiExec.exe /X"{${$WIRA.ProductCode}}" /quiet
				Start-Process -FilePath ($env:SystemRoot + "\Sysetm32\MsiExec.exe") -ArgumentList ('/X' + $WIRA.ProductCode + ' /quiet') -NoNewWindow -Wait
			}Catch{
				
			}
		}
	}Catch{
		Write-host "Trying to uninstall any installed $RemoveAppName with WMI"
		Get-WmiObject -Class Win32_Product -ErrorAction SilentlyContinue | Where-Object Name -Match $RemoveAppName | Foreach-Object { Write-Host ("Removing: " + $_.Name) ;$_.Uninstall()}
	}
	#endregion Remove by Windows Installer

	#region Remove by DISM
	Try {
		Write-host "Getting dism packages listing to remove Flash"
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
	Write-host "Trying to remove Flash update Service"
	Try {
		Get-Service | Where-Object {$_.Name -eq "AdobeFlashPlayerUpdateSvc"} | Remove-Service
	}Catch {
		#Start-Process -FilePath ($env:SystemRoot + "\System32\cmd.exe") -ArgumentList "/C sc STOP AdobeFlashPlayerUpdateSvc && sc DELETE AdobeFlashPlayerUpdateSvc" -NoNewWindow -PassThur
		Start-Process -FilePath ($env:SystemRoot + "\System32\sc.exe") -ArgumentList "STOP AdobeFlashPlayerUpdateSvc" -NoNewWindow
		Start-Process -FilePath ($env:SystemRoot + "\System32\sc.exe") -ArgumentList "DELETE AdobeFlashPlayerUpdateSvc" -NoNewWindow -PassThru | Out-Null
		Write-host ""
	}
	#endregion Remove Flash Update Service

	
	#Remove Flash Files
	Write-host "Trying to remove flash program Files"
	$Destination = ($env:SystemRoot + "\system32\Macromed")
	If (Test-Path -Path $Destination) {
		ForEach ($File in (Get-ChildItem -Path $Destination -Recurse  -File | Sort-Object -Property FullName -Descending)) {
			Write-Host ("`tTrying to remove: " + $File.FullName)
			Set-Owner -Path $File.FullName -Account 'Administrators'
			$Acls = Get-Acl $File.FullName
			$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule("BUILTIN\Administrators", "FullControl", "Allow")
			$Acls.SetAccessRule($Ar) 
			Set-Acl $File.FullName $Acls -ErrorAction SilentlyContinue
			
			Remove-Item -Force -Path $File.FullName -ErrorAction SilentlyContinue
			If(Test-Path -Path $File.FullName) {
				Move-OnReboot -Path $File.FullName
				$RebootRequired = $True
			}
		}
	}
	$Destination = ($env:SystemRoot + "\SysWOW64\Macromed")
	If (Test-Path -Path $Destination) {
		ForEach ($File in (Get-ChildItem -Path $Destination -Recurse  -File | Sort-Object -Property FullName -Descending)) {
			Write-Host ("`tTrying to remove: " + $File.FullName)
			Set-Owner -Path $File.FullName -Account 'Administrators'
			$Acls = Get-Acl $File.FullName
			$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule("BUILTIN\Administrators", "FullControl", "Allow")
			$Acls.SetAccessRule($Ar) 
			Set-Acl $File.FullName $Acls -ErrorAction SilentlyContinue
			
			Remove-Item -Force -Path $File.FullName -ErrorAction SilentlyContinue
			If(Test-Path -Path $File.FullName) {
				Move-OnReboot -Path $File.FullName
				$RebootRequired = $True
			}
		}
	}
	
	$Profiles = (Get-ChildItem -Path ($env:SystemDrive + "\Users")).FullName
	#remove flash from user install
	ForEach ($Profile in $Profiles) {
		$Destination = ($Profile + "AppData\Roaming\Adobe\Flash Player")
		If (Test-Path -Path $Destination) {
			ForEach ($File in (Get-ChildItem -Path $Destination -Recurse  -File | Sort-Object -Property FullName -Descending)) {
				Write-Host ("`tTrying to remove: " + $File.FullName)
				Set-Owner -Path $File.FullName -Account 'Administrators'
				$Acls = Get-Acl $File.FullName
				$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule("BUILTIN\Administrators", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
				$Acls.SetAccessRule($Ar) 
				Set-Acl $File.FullName $Acls -ErrorAction SilentlyContinue
				
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
					$Ar = New-Object system.Security.AccessControl.FileSystemAccessRule("BUILTIN\Administrators", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
					$Acls.SetAccessRule($Ar) 
					Set-Acl $File.FullName $Acls -ErrorAction SilentlyContinue
					
					Remove-Item -Force -Path $File.FullName -ErrorAction SilentlyContinue
					If(Test-Path -Path $File.FullName) {
						Move-OnReboot -Path $File.FullName
						$RebootRequired = $True
					}
					
				}
			}
		}
	}
#endregion Remove Adobe Flash

If ($RebootRequired) {
	Write-Warning " "
	Write-Warning " "
	Write-Warning " "
	Write-Warning "Reboot Required to remove files"
}

Write-Host ("")
Write-Host ("Script took: " + (FormatElapsedTime($sw.Elapsed)) + " to run.")
Write-Host ("")
If (-Not [string]::IsNullOrEmpty($LogFile)) {
	Stop-Transcript
}
Exit 0
