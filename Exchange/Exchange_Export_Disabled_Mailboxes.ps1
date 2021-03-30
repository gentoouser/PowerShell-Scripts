<# 
.SYNOPSIS
    Name: Exchange_Export_Disabled_Mailboxes.ps1
    Disabled User MailBox Export

.DESCRIPTION
	*Find All Disabled Users in Active Directory
	*Filters for Mailboxes 
	*See if other users have full mailbox rights
    *Export mailbox to PST

    .DEPENDENCIES
    *Active Directory Module
    *Exchange Module

.PARAMETER Archive
    UNC path for where .pst will be exported to.
.PARAMETER Server
	FQDN of Exchange server.
.PARAMETER Mailbox
	Username to export one mailbox.
.PARAMETER NoDisable
	Do not Disable user in exchange after export.
.PARAMETER TestOnly
	Only shows who would be exported. 
.PARAMETER Wait
    Time to wait for PST export.
.EXAMPLE
   & Exchange_Export_Disabled_Mailboxes.ps1 -Archive \\remoteserver\share

.NOTES
 Author: Paul Fuller
 Changes:
    1.0.1 - Fixing Display issues. 
	1.0.2 - Fixed Display issue with Mailbox Permissions.
	1.0.3 - Updated Progress Display both Progress-bars and Create sub-folder called Logs for log files
	1.0.4 - Cleaned up Export-Mail function to make it more portable
	1.0.5 - Allow export of just one mailbox. Enable MAPI to export if needed. Added Switch to not-disable Exchange account after export. 
	1.0.6 - Create TestOnly that exports the users to CSV 
	1.0.7 - Updated the way Mailbox permission are evaluated.  
	1.0.8 - Allow to set the export Priority.
#>
PARAM (
    [Parameter(Mandatory=$false,HelpMessage="Folder path to Archived homedrive.")][string]$Archive ,
    [Parameter(Mandatory=$false,HelpMessage="Exchange Server.")][string]$Server,
    [Parameter(Mandatory=$false,HelpMessage="User to samAccountName to export.")][string]$MailBox,
    [Parameter(Mandatory=$false,HelpMessage="Leave mailbox active.")][switch]$NoDisable,
    [Parameter(Mandatory=$false,HelpMessage="Show only who would be exported.")][switch]$TestOnly,
    [Parameter(Mandatory=$false,HelpMessage="Seconds to wait before refreshing progress.")][int]$Wait = 120,
	[Parameter(Mandatory=$false,HelpMessage="Set the priority of export.")][ValidateSetAttribute("Lower","Low","Normal","High","Higher","Highest","Emergency")][string]$Priority="Normal",
	[Parameter(Mandatory=$false)][array]$ExcludeUsers=@(
		($env:USERDOMAIN + "\Domain Admins"),
		($env:USERDOMAIN + "\Enterprise Admins"),
		($env:USERDOMAIN + "\Organization Management"),
		($env:USERDOMAIN + "\Exchange Servers"),
		($env:USERDOMAIN + "\Exchange Domain Servers"),
		($env:USERDOMAIN + "\Administrators"),
		"NT AUTHORITY\SYSTEM",
		"NT AUTHORITY\SELF"
		)
)
$ScriptVersion = "1.0.8"
#Requires -Version 5.1 -PSEdition Desktop -Assembly System.DirectoryServices.AccountManagement

#relauch script if running with powershell lower then 5.1
# if ($PSVersionTable.PSVersion -lt [Version]"5.1") {
# 	# Re-launch as version 5 if we're not already
#     $params = ($PSBoundParameters.GetEnumerator() | ForEach-Object {"-{0} {1}" -f $_.Key,$_.Value}) -join " "
#     Start-Process -File PowerShell.exe -Argument "-Version 2 -noprofile -noexit -file $($myinvocation.mycommand.definition) $params"
#     Break
#   }

#############################################################################
#region User Variables
#############################################################################
$LogFile = ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\Logs\" + `
		($MyInvocation.MyCommand.Name -replace ".ps1","") + "_" + `
		(Get-Date -format yyyyMMdd-hhmm) + ".log")
$sw = [Diagnostics.Stopwatch]::StartNew()
$ID = 1
$DomainName = (Get-CimInstance Win32_NTDomain).DomainName
#############################################################################
#endregion User Variables$
#############################################################################

#############################################################################
#region Setup Sessions
#############################################################################
If (-Not [string]::IsNullOrEmpty($LogFile)) {
	If (-Not( Test-Path (Split-Path -Path $LogFile -Parent))) {
		New-Item -ItemType directory -Path (Split-Path -Path $LogFile -Parent)
	}
	try { 
		Stop-transcript -ErrorAction SilentlyContinue | Out-Null
	} catch { 
		#No transcript running
	} 
	try { 
		Start-Transcript -Path $LogFile -Append
	} catch { 
		Stop-transcript -ErrorAction SilentlyContinue | Out-Null
		Start-Transcript -Path $LogFile -Append
	} 	
	Write-Host ("Script: " + $MyInvocation.MyCommand.Name)
	Write-Host ("Version: " + $ScriptVersion)
	Write-Host (" ")
}		


#region ignore any SSL Warning 
## Choose to ignore any SSL Warning issues caused by Self Signed Certificates  

## Code From http://poshcode.org/624
## Create a compilation environment
$Provider=New-Object Microsoft.CSharp.CSharpCodeProvider
$Provider.CreateCompiler()
$Params=New-Object System.CodeDom.Compiler.CompilerParameters
$Params.GenerateExecutable=$False
$Params.GenerateInMemory=$True
$Params.IncludeDebugInformation=$False
$Params.ReferencedAssemblies.Add("System.DLL") | Out-Null

$TASource=@'
  namespace Local.ToolkitExtensions.Net.CertificatePolicy{
	public class TrustAll : System.Net.ICertificatePolicy {
	  public TrustAll() { 
	  }
	  public bool CheckValidationResult(System.Net.ServicePoint sp,
		System.Security.Cryptography.X509Certificates.X509Certificate cert, 
		System.Net.WebRequest req, int problem) {
		return true;
	  }
	}
  }
'@ 
$TAResults=$Provider.CompileAssemblyFromSource($Params,$TASource)
$TAAssembly=$TAResults.CompiledAssembly

## We now create an instance of the TrustAll and attach it to the ServicePointManager
$TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
[System.Net.ServicePointManager]::CertificatePolicy=$TrustAll

## end code from http://poshcode.org/624
#endregion ignore any SSL Warning 
#region Load Exchange Module
# Load All Exchange PSSnapins 
If ((Get-PSSession | Where-Object { $_.ConfigurationName -Match "Microsoft.Exchange" }).Count -eq 0 ) {
	Write-Host ("Loading Exchange Plugins") -foregroundcolor "Green"
	If ($([System.Net.Dns]::GetHostByName(($env:computerName))).hostname -eq $([System.Net.Dns]::GetHostByName(($Server))).hostname) {
		Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue
		. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
		Connect-ExchangeServer -auto -AllowClobber
	} else {
		$ERPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://$Server/PowerShell/ -Authentication Kerberos
		Import-PSSession $ERPSession -AllowClobber
	}
} Else {
	Write-Host ("Exchange Plug-ins Already Loaded") -foregroundcolor "Green"
}
#endregion Load Exchange Module
#Load .Net Assembly for AD
Add-Type -AssemblyName System.DirectoryServices.AccountManagement
$ContextType = [System.DirectoryServices.AccountManagement.ContextType]::Domain
$IdentityType = [System.DirectoryServices.AccountManagement.IdentityType]::SamAccountName

#Get mailbox
If ($MailBox) {
	Write-Host ("Getting Account. Please wait . . .")
	$DisabledAccounts = Get-User $MailBox
} else {
	#Get All Disabled accounts in Exchange
	Write-Host ("Getting Disabled Accounts. Please wait . . .")
	$DisabledAccounts = Get-User -RecipientTypeDetails UserMailbox -ResultSize Unlimited | Where-Object {$_.UseraccountControl -like "*accountdisabled*"}
}
#############################################################################
#endregion Setup Sessions
#############################################################################

#############################################################################
#region Functions
#############################################################################
Function FormatElapsedTime($ts) {
    #https://stackoverflow.com/questions/3513650/timing-a-commands-execution-in-powershell
	$elapsedTime = $null
    if ( $ts.Hours -gt 0 ){
        $elapsedTime = [string]::Format( "{0:00} hours {1:00} min. {2:00}.{3:00} sec.", $ts.Hours, $ts.Minutes, $ts.Seconds, $ts.Milliseconds / 10 );
    } else {
        if ( $ts.Minutes -gt 0 ){
            $elapsedTime = [string]::Format( "{0:00} min. {1:00}.{2:00} sec.", $ts.Minutes, $ts.Seconds, $ts.Milliseconds / 10 );
        } else {
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
Function Export-Mail() {
	[CmdletBinding()] 
	Param 
	( 
		[Parameter(Mandatory=$true,Position=1,HelpMessage="Username or Identity of user.")][string]$User, 
		[Parameter(Mandatory=$true,Position=2,HelpMessage="Path to archive PST to.")][string]$Archive,
		[Parameter(Mandatory=$false,Position=3,HelpMessage="Disable user in exchange after export.")][switch]$Disable,
		[Parameter(Mandatory=$false,Position=4,HelpMessage="Set the priority of export.")][ValidateSetAttribute("Lower","Low","Normal","High","Higher","Highest","Emergency")][string]$Priority="Normal"
	) 
	[bool]$MapiEnabled=$false
    #Get User Mailbox object
    $ObjUser = Get-User $User
	If ($ObjUser.RecipientType -eq "UserMailbox" ) {
		$CurrentMailBox = $ObjUser | Get-Mailbox
		#Testing to see if is in queue
		If ((Get-MailboxExportRequest | Where-Object { $_.Mailbox -eq $CurrentMailBox.Identity}).count -eq 0) {
			Write-Host ("`tExport Mail Name: " + $ObjUser.Name + " Alias: " + $ObjUser.SamAccountName + " Email: " + $ObjUser.WindowsEmailAddress)  -foregroundcolor "Cyan"
			#Create Archive path if not created
			if (-Not (Test-Path $Archive )) {
				New-Item -ItemType directory -Path $Archive | Out-Null
				If (-Not $?) {
						Write-Warning ("Path not valid: $Archive")
					Return
				}
			}
			#test to see if User has been exported
			if (Test-Path ($Archive + "\" + $($ObjUser.SamAccountName) + ".pst") ) {
					Write-Warning ("User: " + $ObjUser.SamAccountName + " already has been exported to: " + $($ObjUser.SamAccountName) + ".pst")
					Return
			}
			#Test to see of MAPI is enabled
			if (-Not (Get-CASMailbox -Identity $ObjUser.SamAccountName).MapiEnabled) {
				#Enable MAPI
				Set-CASMailbox -Identity $ObjUser.SamAccountName -MAPIEnabled $true
				[System.GC]::Collect()
				Start-Sleep -Seconds 5
			}else{ 
				$MapiEnabled=$true
			}
			#Export Mailbox to PST
			#$MER = New-MailboxExportRequest -Mailbox $ObjUser.SamAccountName -FilePath $($Archive + "\" + $($ObjUser.SamAccountName) + ".pst")
			New-MailboxExportRequest -Mailbox $ObjUser.SamAccountName -FilePath $($Archive + "\" + $($ObjUser.SamAccountName) + ".pst") -Priority $Priority | Out-Null
			If (-Not $?) {
				Return
			}
			Start-Sleep -Seconds 15
		} else {
			Write-Host ("`t`tUser " + $ObjUser.Name + " already submitted. ")
		}
		#Monitor Export	
		$ExportJobStatusName = $null
		$ExportJobStatusName = Get-MailboxExportRequest | Where-Object { $_.mailbox -eq $CurrentMailBox.Identity -and $_.status -ne 10 } | Select-Object -first 1 | Get-MailboxExportRequestStatistics 
		If ($null -ne $ExportJobStatusName) {
			#Write-Host ("`t`t`t Job Status loop: " + $ExportJobStatusName.status)
			while (($ExportJobStatusName.status -ne "Completed") -And ($ExportJobStatusName.status -ne "Failed")) {
				#View Status of Mailbox Export
				$ExportJobStatusName = Get-MailboxExportRequest | Where-Object { $_.mailbox -eq $CurrentMailBox.Identity -and $_.status -ne 10 } | Select-Object -first 1 | Get-MailboxExportRequestStatistics 
				Write-Progress -Id ($Id+1) -Activity $("Exporting user: " + $ExportJobStatusName.SourceAlias ) -status $("Export Percent Complete: " + $ExportJobStatusName.PercentComplete + " Copied " + $ExportJobStatusName.BytesTransferred + " out of " + $ExportJobStatusName.EstimatedTransferSize ) -percentComplete $ExportJobStatusName.PercentComplete
				Start-Sleep -Seconds 15
			}
		}

		#Check for Completion status
		$ExportMailBoxList = Get-MailboxExportRequest | Where-Object { $_.Mailbox -eq $CurrentMailBox.Identity -And ($_.status -ne "Completed" -Or $_.status -ne "Failed")}
			
		If ($ExportMailBoxList.status -eq "Completed") {
			#Remove Exchange account of PST was successful. 
			Write-Host ("`t`t Removing Mailbox from Exchange: " + $CurrentMailBox.Identity)
			#Disable MAPI unless it was already enabled
			Set-CASMailbox -Identity $ObjUser.SamAccountName -MAPIEnabled $MapiEnabled
			If ($Disable) {
				Disable-Mailbox -Identity $ObjUser.SamAccountName -confirm:$false
			}
			$ExportMailBoxList | Remove-MailboxExportRequest -Confirm:$false
		}
		#Stop if PST Export failed.
		If ($ExportMailBoxList.status -eq "Failed") {
			throw ("PST Export failed: " + $error[0].Exception)
			Break
		}
	}
}
#############################################################################
#endregion Functions
#############################################################################

#############################################################################
#region Main
#############################################################################
$AtE = 0
$NAtE = 0
If ($DisabledAccounts.count -ge 1) {
	$TotalUsers = $DisabledAccounts.count	
}else {
	$TotalUsers = 1
}
If ($TestOnly) {
	$CSVFile = $LogFile -replace "\.log",".csv"
	"Action,Name,Alias,Email,Size,Permissions" | Out-File -FilePath $CSVFile
}
ForEach ($DA in $DisabledAccounts) {
    $FixedDAMP = @{}
	$DAMP = $null
    Write-Progress -Id 0 -Activity $("Processing User: " + $DA.Name ) -status $("User: " + ($AtE + $NAtE +1) + " out of " + $TotalUsers ) -percentComplete ((($AtE + $NAtE + 1)/$TotalUsers)*100) 
    Write-Host ("Processing User: " + $DA.Name) -ForegroundColor DarkGray
	#Get Mailbox rights but exclude predefined users
	$DAMP = $DA | Get-MailboxPermission | Where-Object {($_.AccessRights -eq "FullAccess") -and ($_.User -notin $ExcludeUsers) -and ($_.User -notmatch "S-1-5-*" ) -and $_.IsInherited -eq $false} | Sort-Object -Unique -Property User
    
	If ($DAMP.count -gt 0 ) {
		#region filter Mailbox rights of users that are disabled too
        ForEach( $ACE in $DAMP) {
            [switch]$ADValid=$false
            If ($DomainName -eq (split-path -Path $ACE.User -Parent)) {
                $DomainObject = [DirectoryServices.AccountManagement.Principal]::FindByIdentity($ContextType,$IdentityType,(split-path -Path $ACE.User -Leaf))
                If ($DomainObject.StructuralObjectClass -eq  "user") {
                    If ($DomainObject.Enabled) {
                        $ADValid=$true
                    }
                }elseif ($DomainObject.StructuralObjectClass -eq  "group") {
                    $ADValid=$true
                }
            }else {
                $ADValid=$true
            }
            If ($ADValid) {
                If (-Not $FixedDAMP.ContainsKey($ACE.User)) {
                    $FixedDAMP.add($ACE.User,$ACE.AccessRights)
                }
            }
        }
		#endregion filter Mailbox rights of users that are disabled too
       If ($FixedDAMP.count -eq 0 ) {
			#Continue Export
			If ($TestOnly) {
				$MSize = ($DA | Get-MailboxStatistics| Select-Object @{name="Total Item Size (MB)"; expression={[math]::Round(($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),2)}})."Total Item Size (MB)"
				Write-Host ("`tExport, Mail Name: " + $DA.Name + " Alias: " +$DA.SamAccountName + " Email: " + $DA.WindowsEmailAddress + " Size (MB): " + ('{0:N0}' -f $MSize)) 
				("Export," + $DA.Name + "," + $DA.SamAccountName + "," +$DA.WindowsEmailAddress + ',"' + ('{0:N0}' -f $MSize) + '",' ) | Out-File -FilePath $CSVFile -Append
			}else{
				#Create Archive Folder
				if (-Not (Test-Path ($Archive + "\" + $DA.SamAccountName))) {
					New-Item -ItemType directory -Path ($Archive  + "\" + $DA.SamAccountName) | Out-Null
				}
				#Start Mail Export
				If ($NoDisable) {
					Export-Mail -User $DA.SamAccountName -Archive ($Archive  + "\" + $DA.SamAccountName) -Priority $Priority
				}else {
					Export-Mail -User $DA.SamAccountName -Archive ($Archive  + "\" + $DA.SamAccountName) -Disable -Priority $Priority
				}
			}
       }else {
			Write-Host ("`tSkipping, Mail Name: " + $DA.Name + " Alias: " +$DA.SamAccountName + " Email: " + $DA.WindowsEmailAddress + " Size (MB): " + ('{0:N0}' -f $MSize)) 

			Write-Host ("`tMailBox Permissions Count: " + $DAMP.count) -ForegroundColor Red
			Write-Host ("`tMailBox Permissions Fixed Count: " + $FixedDAMP.count) -ForegroundColor yellow
			$FixedDAMP | Format-Table
			$NAtE ++
			#Write Output to CSV
			If ($TestOnly) {	
				$CSVFixedDAMP = $null
				#Loop though all perms and add to string
				Foreach ($FDAMP in $FixedDAMP.GetEnumerator()) {
					If ($FDAMP.Key) {
						$CSVFixedDAMP = ($CSVFixedDAMP + ";" + $FDAMP.Key + " - " + (($FDAMP.Value | Out-String) -join " ") )
					}else{
						If ($FDAMP.Keys) {
							$CSVFixedDAMP = ($CSVFixedDAMP + ";" + $FDAMP.Keys + " - " + (($FDAMP.Values | Out-String) -join " ") )
						}
					}
				}
				#clean up first ; in string and remove all new lines in string
				If ($CSVFixedDAMP.substring(0,1) -eq ";") {
					$CSVFixedDAMP = $CSVFixedDAMP.substring(1,$CSVFixedDAMP.Length -1 ) -replace "`n|`r"
				}
				#write output to CSV
				("Skipping," + $DA.Name + "," + $DA.SamAccountName + "," +$DA.WindowsEmailAddress + ',"' + ('{0:N0}' -f $MSize) + '","' + $CSVFixedDAMP + '"') | Out-File -FilePath $CSVFile -Append
			}
       }
        
    }else {
        #Continue Export
		If ($TestOnly) {
			$MSize = ($DA | Get-MailboxStatistics| Select-Object @{name="Total Item Size (MB)"; expression={[math]::Round(($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),2)}})."Total Item Size (MB)"
			Write-Host ("`tExport, Mail Name: " + $DA.Name + " Alias: " +$DA.SamAccountName + " Email: " + $DA.WindowsEmailAddress + " Size (MB): " + ('{0:N0}' -f $MSize)) 
			("Export," + $DA.Name + "," + $DA.SamAccountName + "," +$DA.WindowsEmailAddress + ',"' + ('{0:N0}' -f $MSize) + '",' ) | Out-File -FilePath $CSVFile -Append
		}else{
			#Create Archive Folder
			if (-Not (Test-Path ($Archive + "\" + $DA.SamAccountName))) {
				New-Item -ItemType directory -Path ($Archive  + "\" + $DA.SamAccountName) | Out-Null
			}
			#Start Mail Export
			If ($NoDisable) {
				Export-Mail -User $DA.SamAccountName -Archive ($Archive  + "\" + $DA.SamAccountName) -Priority $Priority
			}else {
				Export-Mail -User $DA.SamAccountName -Archive ($Archive  + "\" + $DA.SamAccountName) -Disable -Priority $Priority
			}
		}
        $AtE ++
    }
}

$sw.Stop()
Write-Host ("Script took: " + (FormatElapsedTime($sw.Elapsed)) + " to run.")
#############################################################################
#endregion Main
#############################################################################
If (-Not [string]::IsNullOrEmpty($LogFile)) {
	Stop-Transcript
}
