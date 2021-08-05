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
.PARAMETER Disable
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
    1.0.01 - Fixing Display issues. 
	1.0.02 - Fixed Display issue with Mailbox Permissions.
	1.0.03 - Updated Progress Display both Progress-bars and Create sub-folder called Logs for log files
	1.0.04 - Cleaned up Export-Mail function to make it more portable
	1.0.05 - Allow export of just one mailbox. Enable MAPI to export if needed. Added Switch to not-disable Exchange account after export. 
	1.0.06 - Create TestOnly that exports the users to CSV 
	1.0.07 - Updated the way Mailbox permission are evaluated.  
	1.0.08 - Allow to set the export Priority.
	1.0.09 - Fix wait bug and size bug
	1.0.10 - Fix TestOnly output
	1.0.11 - Added option to split mailbox by year. 
	1.0.12 - Added added more infor to TestOnly and forced split mailboxes to happen if the mailbox is to large or has to may items in one folder. 
	1.0.13 - Allows for multiple Mailboxes using array or Comma delimited string.
	1.0.14 - Fixed restarting on split mailbox and added another progress bar.
	1.0.15 - Fixed issues with exporting from Exchange 2010. Also null status errors for write-progress
	1.0.16 - Fixed issue detecting already created PST files for mailbox. Added ablity to test PST export drive with different credentials.
	1.0.17 - Fixed Progress display
#>
PARAM (
    [Parameter(Mandatory=$false,HelpMessage="Folder path to Archived homedrive.")][string]$Archive = (Split-Path -Parent -Path $MyInvocation.MyCommand.Definition),
    [Parameter(Mandatory=$false,HelpMessage="Exchange Server.")][string]$Server = "mail",
    [Parameter(Mandatory=$false,HelpMessage="User to samAccountName to export.")][string[]]$MailBox,
    [Parameter(Mandatory=$false,HelpMessage="Disable mailbox. (Remove from Outlook)")][switch]$Disable,
    [Parameter(Mandatory=$false,HelpMessage="Show only who would be exported.")][switch]$TestOnly,
    [Parameter(Mandatory=$false,HelpMessage="Seconds to wait before refreshing progress.")][int]$Wait = 15,
	[Parameter(Mandatory=$false,HelpMessage="Set the priority of export.")][ValidateSet("Normal","High")][string]$Priority="Normal",
	[Parameter(Mandatory=$false,Position=5,HelpMessage="Split PST by Years.")][switch]$SplitYear, 
	[Parameter(Mandatory=$false,Position=5,HelpMessage="Sort user in reverse order.")][switch]$Reverse, 
	[Parameter(Mandatory=$false,Position=5,HelpMessage="Max items in one folder before splitting into yearly .pst files.")][int64]$MaxItems = 950000, 
	[Parameter(Mandatory=$false,Position=5,HelpMessage="Max Mailbox size before splitting into yearly .pst files.")][int64]$MaxMailboxSize=45GB, 
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
$ScriptVersion = "1.0.17"
#Requires -Version 5.1 -PSEdition Desktop -Assembly System.DirectoryServices.AccountManagement


#############################################################################
#region User Variables
#############################################################################
$LogFile = ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\Logs\" + `
			($MyInvocation.MyCommand.Name -replace ".ps1","") + "_" + `
			(Get-Date -format yyyyMMdd-hhmm) + ".log")
$CSVFile = ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\Logs\" + `
			($MyInvocation.MyCommand.Name -replace ".ps1","") + "_" + `
			(Get-Date -format yyyyMMdd-hhmm) + ".csv")
$sw = [Diagnostics.Stopwatch]::StartNew()
$ID = 1
$DomainName = (Get-CimInstance Win32_NTDomain).DomainName
$TestOnlyOut = @()
$DisabledAccounts = @()
$CurrentYear = (get-date).year
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
	#Handles Comma delimited list.
	If ($MailBox -contains ",") {
		$MailBox = ($MailBox.Split(","))
	}
	#Allows for multiple mailbox input
	ForEach ($CM in $Mailbox) {
		$DisabledAccounts += Get-User $CM
	}
} else {
	If ($Reverse) {
		#Get All Disabled accounts in Exchange
		Write-Host ("Getting Disabled in Reverse order Accounts. Please wait . . .")
		$DisabledAccounts += Get-User -RecipientTypeDetails UserMailbox -ResultSize Unlimited | Where-Object {$_.UseraccountControl -like "*accountdisabled*"} | Sort-Object -Descending -Property SamAccountName
	}Else{
		#Get All Disabled accounts in Exchange
		Write-Host ("Getting Disabled Accounts. Please wait . . .")
		$DisabledAccounts += Get-User -RecipientTypeDetails UserMailbox -ResultSize Unlimited | Where-Object {$_.UseraccountControl -like "*accountdisabled*"} | Sort-Object -Property SamAccountName
	}
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
Function Export-Mail {
	[CmdletBinding()] 
	Param 
	( 
		[Parameter(Mandatory=$true,Position=1,HelpMessage="Username or Identity of user.")][string]$User, 
		[Parameter(Mandatory=$true,Position=2,HelpMessage="Path to archive PST to.")][string]$Archive,
		[Parameter(Mandatory=$false,Position=3,HelpMessage="Disable user in exchange after export.")][switch]$Disable,
		[Parameter(Mandatory=$false,Position=4,HelpMessage="Set the priority of export.")][ValidateSet("Normal","High")][string]$Priority="Normal",
		[Parameter(Mandatory=$false,Position=5,HelpMessage="Export only specific year.")][int64]$Year,
		[Parameter(Mandatory=$false,Position=6,HelpMessage="Array of Active statuses to look for.")][array]$GoodStatuses = @("CompletionInProgress","InProgress","Queued","Retrying"), 
		[Parameter(Mandatory=$false,Position=7,HelpMessage="Array of InActive statuses to look for.")][array]$BadStatuses = @("AutoSuspended","CompletedWithWarning","Failed","Suspended","Synced"),
		[Parameter(Mandatory=$false,Position=8,HelpMessage="Specifies the parent activity of the current activity")][int]$ParentId = -1,
		[Parameter(Mandatory=$false,Position=9,HelpMessage="Use following Credential for commands (Get-Credential)")][pscredential]$Credential,
		[Parameter(Mandatory=$false,HelpMessage="Seconds to wait before refreshing progress.")][int]$Wait = 15
	)
	[bool]$MapiEnabled=$false
	$ExportJobStatusName = $null
	$ExportJobStatus = $null
	$CMExportBad = $false
	$JobComplete = 0
	$CopiedB = 0
	$CopiedBT = 0
	$CopiedIT = 0
	$CopiedI = 0
	$PST = $null
	If ($ParentId -ne -1) {
		$ID=$ParentId++
	}Else {
		$ID=2
	}
    #Get User Mailbox object
    $ObjUser = Get-User $User
	If ($ObjUser.RecipientType -eq "UserMailbox" ) {
		$CurrentMailBox = $ObjUser | Get-Mailbox
		$CMExport = (Get-MailboxExportRequest | Where-Object { $_.Mailbox -eq $CurrentMailBox.Identity})
		If ($year) {
			If ($CMExport | Where-Object { $_.Name -eq ("Export_" + $ObjUser.SamAccountName + "_" + $Year)}) {
				$CMExportBad = $true
			}else{
				$CMExport = $CMExport | Where-Object {$_.Name -notmatch ("Export_" + $ObjUser.SamAccountName + "_\d+")}
			}
		}
		#Testing to see if is in queue
		If ($CMExport.count -eq 0 -and $CMExportBad -eq $false) {
			If ($Year) {
				Write-Host ("`tExport Name: " + ("Export_" + $ObjUser.SamAccountName + "_" + $Year) + " Alias: " + $ObjUser.SamAccountName + " Email: " + $ObjUser.WindowsEmailAddress)  -foregroundcolor "Cyan"
			}Else{
				Write-Host ("`tExport Name: " + ("Export_" + $ObjUser.SamAccountName) + " Alias: " + $ObjUser.SamAccountName + " Email: " + $ObjUser.WindowsEmailAddress)  -foregroundcolor "Cyan"
			}
			#Test for exsiting psts using other Credentials
			If($Credential) {
				#Test Archive Folder
				If (-Not (Test-Path "PSHome:\") -or (Get-PSdrive -name "PSHome").root -ne $Archive) {
					If (Test-Path "PSHome:\") {
						Remove-PSDrive -Name "PSHome"
					}
					New-PSDrive -Name "PSHome" -PSProvider FileSystem -Root  $Archive  -Credential $Credential  -ErrorAction SilentlyContinue | Out-Null
					If (!(Test-Path "PSHome:\")) {
						#Try creating parent folder. 
						New-PSDrive -Name "PSHome" -PSProvider FileSystem -Root   (Split-Path -Path $Archive -Parent)  -Credential $Credential  -ErrorAction SilentlyContinue | Out-Null
						If (Test-Path "PSHome:\") {
							New-Item -ItemType Directory -Path ("PSHome:\" +(Split-Path -Path $Archive -Leaf))
							Remove-PSDrive -Name "PSHome"
							New-PSDrive -Name "PSHome" -PSProvider FileSystem -Root  $Archive  -Credential $Credential  -ErrorAction SilentlyContinue | Out-Null
						}
					}
					If (!(Test-Path "PSHome:\")) {
						Write-Warning ("Path not valid: $Archive")
						Return
					}
					$PST = Get-ChildItem -Path "PSHome:\" -Filter "*.pst" | Where-Object {$_.Name -eq ($ObjUser.SamAccountName + ".pst") -or $_.name -eq "$($ObjUser.SamAccountName)_$($Year).pst"}
					if ($PST) {
						Write-Warning ("User: " + $ObjUser.SamAccountName + " already has been exported to: " + ($PST | Select-object -First 1 Name).Name)
						Return
					}
				}
			}Else {
				#Use current Credential to test for PSTs.
				#Create Archive path if not created
				if (-Not (Test-Path $Archive )) {
					New-Item -ItemType directory -Path $Archive | Out-Null
					If (-Not $?) {
						Write-Warning ("Path not valid: $Archive")
						Return
					}
				}
				#Test to see if User has been exported
				$PST = Get-ChildItem -Path $Archive -Filter "*.pst" | Where-Object {$_.Name -eq ($ObjUser.SamAccountName + ".pst") -or $_.name -eq "$($ObjUser.SamAccountName)_$($Year).pst"}
				If ($PST) {
						Write-Warning ("User: " + $ObjUser.SamAccountName + " already has been exported to: " + ($PST | Select-object -First 1 Name).Name)
						Return
				}

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

			If ($Year) {
				#Export only on year of mail
				$ExportJobStatusName = New-MailboxExportRequest -Mailbox $ObjUser.SamAccountName -FilePath $($Archive + "\" + $($ObjUser.SamAccountName)  + "_" + $Year + ".pst") -Priority $Priority -ContentFilter "( Received -ge '01/01/$Year' -and Received -le '12/31/$Year')" -BatchName ("Export_" + $ObjUser.SamAccountName + "_" + $Year) -Name ("Export_" + $ObjUser.SamAccountName + "_" + $Year) | Out-Null

				If (-Not $?) {
					Return
				}
			}Else {
				#$MER = New-MailboxExportRequest -Mailbox $ObjUser.SamAccountName -FilePath $($Archive + "\" + $($ObjUser.SamAccountName) + ".pst")
				$ExportJobStatusName = New-MailboxExportRequest -Mailbox $ObjUser.SamAccountName -FilePath $($Archive + "\" + $($ObjUser.SamAccountName) + ".pst") -Priority $Priority -BatchName ("Export_" + $ObjUser.SamAccountName)  -Name ("Export_" + $ObjUser.SamAccountName)| Out-Null

				If (-Not $?) {
					Return
				}
			}
			Start-Sleep -Seconds $Wait
		} else {
			Write-Host ("`t`tUser " + $ObjUser.Name + " already submitted. ")
			$CMExport
		}
		#Try to get Job Name of we do not know it.
		#Try Just user Name
		If (!$ExportJobStatusName) {
			$ExportJobStatusName = Get-MailboxExportRequest -Name ("Export_" + $ObjUser.SamAccountName)
		}
		#Try Just user Name and Year
		If (!$ExportJobStatusName -and $Year) {
			$ExportJobStatusName = Get-MailboxExportRequest -Name ("Export_" + $ObjUser.SamAccountName + "_" + $Year)
		}		
		#Try find any active 
		If (!$ExportJobStatusName) {
			$ExportJobStatusName = Get-MailboxExportRequest -Status InProgress | Where-Object { $_.mailbox -eq $CurrentMailBox.Identity} | Select-Object -first 1 
		}
		#Try find any Job 
		If (!$ExportJobStatusName) {
			$ExportJobStatusName = Get-MailboxExportRequest -Status InProgress | Where-Object { $_.mailbox -eq $CurrentMailBox.Identity -and $_.status -in $GoodStatuses } | Select-Object -first 1 
		}
		#Get Job Stats
		$ExportJobStatus = $ExportJobStatusName | Get-MailboxExportRequestStatistics
		#Monitor Export	
		If ($ExportJobStatusName -and $ExportJobStatus) {
			#Write-Host ("`t`t`t Job Status loop: " + $ExportJobStatusName.status)
			while (($ExportJobStatus.status -ne "Completed") -And ($ExportJobStatus.status -ne "Failed")) {
				#View Status of Mailbox Export
				$ExportJobStatus = $ExportJobStatusName | Get-MailboxExportRequestStatistics 
				IF ($ExportJobStatus){
					If ($ExportJobStatus.PercentComplete -ge 1) {
						$JobComplete = $ExportJobStatus.PercentComplete
					}Else{
						$JobComplete = 0
					}
					If ($ExportJobStatus.BytesTransferred) {
						$CopiedB = [math]::round(($ExportJobStatus.BytesTransferred.ToString()).Split("(")[1].Split(" ")[0].Replace(",","")/1GB,2)
						if (-Not $CopiedB -gt 0) {
							$CopiedB = 0
						}
					}Else{
						$CopiedB = 0
					}
					$CopiedBT = [math]::round(($ExportJobStatus.EstimatedTransferSize.ToString()).Split("(")[1].Split(" ")[0].Replace(",","")/1GB,2)
					if (-Not $CopiedBT -gt 0) {
						$CopiedBT = 0
					}
					$CopiedI = $ExportJobStatus.ItemsTransferred
					if (-Not $CopiedI -gt 0) {
						$CopiedI = 0
					}
					$CopiedIT = $ExportJobStatus.EstimatedTransferItemCount
					if (-Not $CopiedIT -gt 0) {
						$CopiedIT = 0
					}
					Write-Progress -Id $ID -PercentComplete $JobComplete -Activity ("Exporting: " + $ExportJobStatus.Name + " Status Detail: " + $ExportJobStatus.StatusDetail) -status ("Export Percent Complete: " + $JobComplete + " Copied: " + $CopiedB  + " GB/" + $CopiedBT + " GB Items: " + ('{0:N0}' -f $CopiedI) + "/" + ('{0:N0}' -f $CopiedIT))  
				}
				Start-Sleep -Seconds $Wait
			}
		}Else{
			Write-Warning ("Can not find Export job to monitor")
			return
		}
		If ($ExportJobStatusName) {
			$ExportMailBoxList = $ExportJobStatusName | Get-MailboxExportRequest
		}
		If (!$ExportMailBoxList) {
			#Check for Completion status
			$ExportMailBoxList = Get-MailboxExportRequest | Where-Object { $_.Mailbox -eq $CurrentMailBox.Identity -And ($_.status -in $BadStatuses -or $_.Status -eq "Completed")}
		}
		If ($ExportMailBoxList.status -eq "Completed") {
			#Remove Exchange account of PST was successful. 
			
			#Disable MAPI unless it was already enabled
			Set-CASMailbox -Identity $ObjUser.SamAccountName -MAPIEnabled $MapiEnabled
			If ($Disable) {
				Write-Host ("`t`t Removing Mailbox from Exchange: " + $CurrentMailBox.Identity)
				Disable-Mailbox -Identity $ObjUser.SamAccountName -confirm:$false
			}
			Write-Host ("`t`t Removing MailboxExport job from Exchange: " + $CurrentMailBox.Identity)
			$ExportMailBoxList | Remove-MailboxExportRequest -Confirm:$false
		}
		#Stop if PST Export failed.
		If ($ExportMailBoxList.status -in $BadStatuses) {
			$ExportMailBoxList | Get-MailboxExportRequestStatistics | Format-list Message,*Failure*
			throw ("PST Export failed: " + ($ExportMailBoxList | Get-MailboxExportRequestStatistics | Select-Object Message).message)
			return
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
$MCount = 0
If ($DisabledAccounts.count -ge 1) {
	$TotalUsers = $DisabledAccounts.count	
}else {
	$TotalUsers = 1
}

ForEach ($DA in $DisabledAccounts) {
	$CMYearSplit = $SplitYear
    $FixedDAMP = @{}
	$DAMP = $null
	$MCount++
    Write-Progress -Id 0 -Activity $("Processing User: " + $DA.Name ) -status $("User: " + $MCount + " out of " + $TotalUsers ) -percentComplete (($MCount/$TotalUsers)*100) 
    Write-Host ("Processing User: " + $DA.Name) -ForegroundColor DarkGray
	#Get Mailbox Size
	$GSize = ($DA | Get-MailboxStatistics| Select-Object @{name="Total Item Size (GB)"; expression={[math]::Round(($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1GB),2)}})."Total Item Size (GB)"
	#Get the largest number it items in folder.
	$MailboxTopItemCount = ($DA | Get-MailboxFolderStatistics | Sort-Object -Property ItemsInFolder -Descending | Select-Object -First 1)
	#Test to see of over MaxItems and needs to be split
	If ($MailboxTopItemCount.ItemsInFolder -ge $MaxItems) {
		Write-host ("`tMailbox has more than " + ('{0:N0}' -f $MaxItems) + " in one folder. Forcing PST to be split by year.") -ForegroundColor Yellow
		$CMYearSplit = $true
	}
	#Test to see of over MaxMailbox size
	If ([math]::Round(($DA| Get-MailboxStatistics).TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")) -ge $MaxMailboxSize) {
		Write-host ("`tMailbox is larger than " + ('{0:N}' -f ([math]::Round($MaxMailboxSize/1GB))) + "GB. forcing PST to be split by year.") -ForegroundColor Yellow
		$CMYearSplit = $true
	}
	#Get Mailbox rights but exclude predefined users
	$DAMP = $DA | Get-MailboxPermission | Where-Object {($_.AccessRights -eq "FullAccess") -and ($_.User -notin $ExcludeUsers) -and ($_.User -notmatch "S-1-5-*" ) -and $_.IsInherited -eq $false} | Sort-Object -Unique -Property User
    
	If ($DAMP.count -gt 0 ) {
		#region filter Mailbox rights of users that are disabled too
        ForEach( $ACE in $DAMP) {
            [switch]$ADValid=$false
            If ($DomainName -eq (split-path -Path $ACE.User -Parent)) {
                $DomainObject = [DirectoryServices.AccountManagement.Principal]::FindByIdentity($ContextType,$IdentityType,(split-path -Path $ACE.User -Leaf))
				Foreach ($DO in $DomainObject) {
					If ($DO.StructuralObjectClass -eq  "user") {
						If ($DO.Enabled -and $DO.SamAccountName -eq (split-path -Path $ACE.User -Leaf)) {
							$ADValid=$true
						}
					}elseif ($DO.StructuralObjectClass -eq  "group") {
						$ADValid=$true
					}
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
				Write-Host ("`tExport, Email: " + $DA.WindowsEmailAddress + " Size (GB): " + ('{0:N}' -f $GSize)) 
				$TestOnlyOut += New-Object -TypeName PSObject -Property ([ordered]@{ 			
					"Name" = $DA.Name;
					"Alias" = $DA.SamAccountName;
					"Email" = $DA.WindowsEmailAddress;
					"Action" = "Export";
					"Size GB" = ('{0:N}' -f $GSize) ;
					"Highest Item Count" = ('{0:N}' -f ($MailboxTopItemCount.ItemsInFolder));
					"Split Year" = $CMYearSplit ;
					"Created" = $DA.WhenCreated.year;
					"Permissions" = ""
				})
			}else{
				#Create Archive Folder
				if (-Not (Test-Path ($Archive + "\" + $DA.SamAccountName))) {
					New-Item -ItemType directory -Path ($Archive  + "\" + $DA.SamAccountName) | Out-Null
				}
				If ($CMYearSplit) {
					$YearCounter = $DA.WhenCreated.year
					$YearCreated = $YearCounter
					while($YearCounter -lt $CurrentYear) {
						Write-Progress -Id 1 -Activity $("Processing Year: " + $YearCounter ) -status $("Split PST: " + ($YearCounter - $YearCreated) + " out of " + ($CurrentYear-$YearCreated)) -percentComplete ((($YearCounter - $YearCreated)/($CurrentYear-$YearCreated))*100) 
						If ($Disable) {
							Export-Mail -User $DA.SamAccountName -Archive ($Archive  + "\" + $DA.SamAccountName) -Disable -Priority $Priority -Wait $Wait -Year $YearCounter -ParentID 2
						}else {
							Export-Mail -User $DA.SamAccountName -Archive ($Archive  + "\" + $DA.SamAccountName) -Priority $Priority -Wait $Wait -Year $YearCounter -ParentID 2
						}
						$YearCounter++
					}
				}Else{
					#Start Mail Export
					If ($Disable) {
						Export-Mail -User $DA.SamAccountName -Archive ($Archive  + "\" + $DA.SamAccountName) -Disable -Priority $Priority -Wait $Wait -ParentID 1
					}else {
						Export-Mail -User $DA.SamAccountName -Archive ($Archive  + "\" + $DA.SamAccountName) -Priority $Priority -Wait $Wait -ParentID 1
					}
				}
			}
        }else {
			Write-Host ("`tMailBox Permissions Count: " + $DAMP.count) -ForegroundColor Red
			Write-Host ("`tMailBox Permissions Fixed Count: " + $FixedDAMP.count) -ForegroundColor yellow
			$FixedDAMP | Format-Table -AutoSize
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
				$TestOnlyOut += New-Object -TypeName PSObject -Property ([ordered]@{ 
					"Name" = $DA.Name;
					"Alias" = $DA.SamAccountName;
					"Email" = $DA.WindowsEmailAddress;
					"Action" = "Skipping";
					"Size GB" = ('{0:N}' -f $GSize) ;
					"Highest Item Count" = ('{0:N}' -f ($MailboxTopItemCount.ItemsInFolder));
					"Split Year" = $CMYearSplit ;
					"Created" = $DA.WhenCreated.year;
					"Permissions" = $CSVFixedDAMP
				})
			}
        }
        
    }else{
        #Continue Export
		If ($TestOnly) {
			Write-Host ("`tExport, Email: " + $DA.WindowsEmailAddress + " Size (GB): " + ('{0:N}' -f $GSize)) 
			$TestOnlyOut += New-Object -TypeName PSObject -Property ([ordered]@{ 
				"Name" = $DA.Name;
				"Alias" = $DA.SamAccountName;
				"Email" = $DA.WindowsEmailAddress;
				"Action" = "Export";
				"Size GB" = ('{0:N}' -f $GSize) ;
				"Highest Item Count" = ('{0:N}' -f ($MailboxTopItemCount.ItemsInFolder));
				"Split Year" = $CMYearSplit ;
				"Created" = $DA.WhenCreated.year;
				"Permissions" = ""
			})
		}else{
			#Create Archive Folder
			if (-Not (Test-Path ($Archive + "\" + $DA.SamAccountName))) {
				New-Item -ItemType directory -Path ($Archive  + "\" + $DA.SamAccountName) | Out-Null
			}

			If ($CMYearSplit) {
				$YearCounter = $DA.WhenCreated.year
				while($YearCounter -le ((get-date).year)) {
					If ($Disable) {
						Export-Mail -User $DA.SamAccountName -Archive ($Archive  + "\" + $DA.SamAccountName) -Disable -Priority $Priority -Wait $Wait -Year $YearCounter -ParentID 2
					}else {
						Export-Mail -User $DA.SamAccountName -Archive ($Archive  + "\" + $DA.SamAccountName) -Priority $Priority -Wait $Wait -Year $YearCounter -ParentID 2
					}
					$YearCounter++
				}
			}Else{
				#Start Mail Export
				If ($Disable) {
					Export-Mail -User $DA.SamAccountName -Archive ($Archive  + "\" + $DA.SamAccountName) -Disable -Priority $Priority -Wait $Wait -ParentID 1
				}else {
					Export-Mail -User $DA.SamAccountName -Archive ($Archive  + "\" + $DA.SamAccountName) -Priority $Priority -Wait $Wait -ParentID 1
				}
			}
		}
        $AtE ++
    }
}
If ($TestOnly -and $TestOnlyOut) {
	$TestOnlyOut | Export-csv -NoTypeInformation -Path $CSVFile
}
$sw.Stop()
Write-Host ("Script took: " + (FormatElapsedTime($sw.Elapsed)) + " to run.")
#############################################################################
#endregion Main
#############################################################################
If (-Not [string]::IsNullOrEmpty($LogFile)) {
	Stop-Transcript
}
