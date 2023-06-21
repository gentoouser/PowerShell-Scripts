<# 
.SYNOPSIS
    Name: Hybrid Mailbox and AD Info.ps1
    Creates CSV with all AD user and Exchange info

.DESCRIPTION
    *Dumps AD User info to CSV.

.EXAMPLE
   & .\Hybrid Mailbox and AD Info.ps1

.NOTES
 AUTHOR  : Victor Ashiedu
 WEBSITE : iTechguides.com
 BLOG    : iTechguides.com/blog-2/
 CREATED : 08-08-2014
 Updated By: Paul Fuller
 Changes:
    * Version 1.00.00 - 
    * Version 1.01.00 - Switch to using Classes
    * Version 1.01.01 - Fix errors, Cache All Mailboxes to increase speed and reduce errors.
    * Version 1.01.02 - Added What protocols are enabled for mailboxes.
    * Version 1.01.03 - Added proxyaddresses/email addresses.
    * Version 1.01.04 - Added DepartmentNumber and extensionAttribute2.
    * Version 1.02.01 - Switch script to cache more and filter from cache.
    

#>
#region Parameters
Param(
	[CmdletBinding()]
    $csvfile =  ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\Logs\" + `
                ($MyInvocation.MyCommand.Name -replace ".ps1","") + "_" + `
                $FileDate + ".csv"),
    $xlsxfile =  ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\Logs\" + `
                ($MyInvocation.MyCommand.Name -replace ".ps1","") + "_" + `
                $FileDate + ".xlsx"),
    $ExcludeUsers=@(
        ($env:USERDOMAIN + "\Domain Admins"),
        ($env:USERDOMAIN + "\Enterprise Admins"),
        ($env:USERDOMAIN + "\Organization Management"),
        ($env:USERDOMAIN + "\Exchange Servers"),
        ($env:USERDOMAIN + "\Exchange Domain Servers"),
        ($env:USERDOMAIN + "\Administrators"),
        "NT AUTHORITY\SYSTEM",
        "NT AUTHORITY\SELF"
    ),
    $ExchangeServer = "exchange.github.com"
)
#endregion Parameters
#region Variables 
$ScriptVersion = "1.2.01"
$output = @()
$FileDate = (Get-Date -format yyyyMMdd-hhmm)
$sw = [Diagnostics.Stopwatch]::StartNew()
#endregion Variables 

Function CleanDistinguishedName{
	[CmdletBinding()]
	Param(

		[Parameter(Mandatory = $true, ValueFromPipeline=$true)][string[]]$MemberOf
	)
	BEGIN {}
    PROCESS {
		ForEach ($Group in $MemberOf) {
			Return ($Group -split ",")[0] -replace "CN="
		}
	}
	END {}
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
Function MailboxGB {
	[CmdletBinding()]
	Param(
		[Parameter(Mandatory=$true,Position=1,HelpMessage="Mailbox Object that contains size")][Object]$MailboxSizeObject
	)
	$Return = 0 
	If ($MailboxSizeObject) {
		If($MailboxSizeObject.IsUnlimited) {
			$Return = "Unlimited"
		}Else{
			If(-Not [string]::IsNullOrWhiteSpace($MailboxSizeObject)){
				Try {
					$Return = $MailboxSizeObject.Value.ToGB()
				} Catch {
					$TempString = $MailboxSizeObject.ToString()
					If (-Not  [string]::IsNullOrWhiteSpace($TempString)) {
						If ($TempString -eq "Unlimited") {
							$Return = "Unlimited"
						}Else {
							Try {
								$TempSize = [long]($TempString.Split("(")[1].Split(" ")[0].Replace(",",""))
							}Catch {
								$Return = 0 
							}
						}
					}
					If ($TempSize -gt 0) {
						$Return = [math]::round($TempSize/1GB,2)
					}
				}
			}Else {

			}
		}
	}
	return $Return
}

Class ADExchangeOutput {
	${Logon Name}
	${Display Name}
	${Last Name}
	${Middle Name}
	${First Name}
	${Description}
	${Full address}
	${City}
	${State}
	${Postal Code}
	${Country-Region}
	${Job Title}
	${Company}
	${Department}
	${Department Number}
	${Employee Type}
	${Employee Number}
	${Office}
	${Phone}
	${Mobile Phone}
	${extensionAttribute2 CIFX Company}
	${Azure Licenses}
	${Azure Last Sync Time}
	${Group Membership}
	${Manager} 
	${Home Directory} 
	${Account Status} 
	${Password Never Expires} 
	${Password Not Required} 
	${Smartcard Logon Required} 
	${Last Log-On Date} 
	${Days Since Last Log-On} 
	${Creation Date} 
	${Days Since Creation} 
	${Last Password Change} 
	${Days from last password change} 
	${RDS CAL Expiration Date}
	${Email} 
	${Email Addresses} 
	${Mailbox Location}
	${Mailbox Server}
	${Mailbox Database}
	${Mailbox Permissions}
	${Mailbox Issue Warning Quota}
	${Mailbox Prohibit Send Quota}
	${Mailbox Prohibit Send Receive Quota}
	${Mailbox Use Database Quota Defaults}
	${Mailbox Storage Limit Status}
	${Mailbox Size} 
	${Mailbox Item Count} 
	${Mailbox Last Logged On User Account}
	${Mailbox Last Logon Time}
	${Mailbox Last Logoff Time}
	${OWA Enabled}
	${Mapi Enabled}
	${Active Sync Enabled}
	${IMAP Enabled}
	${POP Enabled}
	${Mailbox GUID}
	${Mailbox Creation Date}
	${Mailbox Archive Status}  
	${Mailbox Archive State}  
	${Mailbox Archive Domain}  
	${Mailbox Archive Location}  
	${Mailbox Archive Server}  
	${Mailbox Archive Database}  
	${Mailbox Archive Permissions}  
	${Mailbox Archive Issue Warning Quota} 
	${Mailbox Archive Prohibit Send Quota} 
	${Mailbox Archive Prohibit Send Receive Quota}  
	${Mailbox Archive Use Database Quota Defaults}  
	${Mailbox Archive Storage Limit Status} 
	${Mailbox Archive Size}
	${Mailbox Archive Item Count}
	${Mailbox Archive Last Logged On User Account}  
	${Mailbox Archive Last Logon Time}
	${Mailbox Archive Last Logoff Time}
	${Mailbox Archive GUID} 
	${Mailbox Archive Creation Date} 
	${Distinguished Name}
}
#region Import AD modules
If ((Get-Module | Where-Object {$_.Name -Match "ActiveDirectory"}).Count -eq 0 ) {
	Write-Host ("Loading Active Directory Plugins") -ForeGroundColor "Green"
	Import-Module "ActiveDirectory"  -ErrorAction SilentlyContinue
} Else {
	Write-Host ("Active Directory Plug-ins Already Loaded") -ForeGroundColor "Green"
}

If (-Not( Test-Path (Split-Path -Path $csvfile -Parent))) {
	New-Item -ItemType directory -Path (Split-Path -Path $csvfile -Parent) | Out-Null
}
#endregion Import AD modules
#region AzureAD
If (Get-Module -ListAvailable -Name "AzureAD") {
    If (-Not (Get-Module "AzureAD" -ErrorAction SilentlyContinue)) {
        Import-Module "AzureAD" -DisableNameChecking
    } Else {
        #write-host "AzureAD PowerShell Module Already Loaded"
    } 
} Else {
    Import-Module PackageManagement
    Import-Module PowerShellGet
    # Register-PSRepository -Name "PSGallery" –SourceLocation "https://www.powershellgallery.com/api/v2/" -InstallationPolicy Trusted
    If (-Not (Get-PSRepository -Name "PSGallery")) {
		Register-PSRepository -Default -InstallationPolicy Trusted 
	}
    Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted 
    Install-Module "AzureAD" -Force -Confirm:$false -Scope:CurrentUser -SkipPublisherCheck -AllowClobber
    Import-Module "AzureAD" -DisableNameChecking
    If (-Not (Get-Module "AzureAD" -ErrorAction SilentlyContinue)) {
        write-error ("Please install AzureAD Powershell Modules Error")
        exit
    }
}
#endregion AzureAD
#region Connect to Exchange Online
If (Get-Module -ListAvailable -Name "ExchangeOnlineManagement") {
    If (-Not (Get-Module "ExchangeOnlineManagement" -ErrorAction SilentlyContinue)) {
        Import-Module "ExchangeOnlineManagement" -DisableNameChecking
    } Else {
        #write-host "ExchangeOnlineManagement PowerShell Module Already Loaded"
    } 
} Else {
    Import-Module PackageManagement
    Import-Module PowerShellGet
    # Register-PSRepository -Name "PSGallery" –SourceLocation "https://www.powershellgallery.com/api/v2/" -InstallationPolicy Trusted
    Register-PSRepository -Default -InstallationPolicy Trusted 
    Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted 
    Install-Module "ExchangeOnlineManagement" -Force -Confirm:$false -Scope:CurrentUser -SkipPublisherCheck -AllowClobber
    Import-Module "ExchangeOnlineManagement" -DisableNameChecking
    If (-Not (Get-Module "ExchangeOnlineManagement" -ErrorAction SilentlyContinue)) {
        write-error ("Please install ExchangeOnlineManagement Powershell Modules Error")
        exit
    }
}
#Connect
If (-Not (Get-PSSession | Where-Object {$_.name -match "ExchangeOnline" -and $_.Availability -eq "Available"})) {
    Connect-ExchangeOnline -ShowBanner:$false -UserPrincipalName (([string]([ADSI]"LDAP://<SID=$([System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value)>").UserPrincipalName) ) -ShowProgress $true 
}
#endregion Connect to Exchange Online
#region Load Exchange commands EP prefix
If ((Get-PSSession | Where-Object { $_.ConfigurationName -Match "Microsoft.Exchange" -and $_.ComputerName -eq $ExchangeServer}).Count -eq 0 ) {
	Write-Host ("Loading Exchange Plugins") -ForeGroundColor "Green"
		#$sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck
		#$ERPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://$ExchangeServer/PowerShell/ -Authentication Kerberos -AllowRedirection -SessionOption $sessionOption
		$ERPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://$ExchangeServer/PowerShell/ -Authentication Negotiate -AllowRedirection 
		Import-PSSession $ERPSession -AllowClobber -DisableNameChecking -Prefix "EP"
		If ((Get-PSSession | Where-Object { $_.ConfigurationName -Match "Microsoft.Exchange" -and ($_.ComputerName -eq $ExchangeServer -or $_.ComputerName -eq "outlook.office365.com") } )) {
			$host.ui.RawUI.WindowTitle = "Hybrid Exchange Administrator Online and Local with EP Prefix: $env:USERDNSDOMAIN\$env:username on $ExchangeServer"
		}Else {
			$host.ui.RawUI.WindowTitle = "Local Exchange Administrator EP Prefix: $env:USERDNSDOMAIN\$env:username on $ExchangeServer"
		}
} Else {
	Write-Host ("Exchange Plug-ins Already Loaded") -ForeGroundColor "Green"
}
#endregion Load Exchange commands EP prefix

#region mailbox checking

Connect-AzureAD -AccountId ($env:USERNAME + "@" + $env:USERDNSDOMAIN )
#Load .Net Assembly for AD
Add-Type -AssemblyName System.DirectoryServices.AccountManagement
$ContextType = [System.DirectoryServices.AccountManagement.ContextType]::Domain
$IdentityType = [System.DirectoryServices.AccountManagement.IdentityType]::SamAccountName

#Sets the OU to do the base search for all user accounts, change as required.
$SearchBase = (Get-ADDomain).DistinguishedName

#Define variable for a server with AD web services installed
$ADServer = (Get-ADDomain).PDCEmulator

Write-Host ("Caching Objects please wait . . ." )
$swc = [Diagnostics.Stopwatch]::StartNew()
#Get All AD users
Write-Host "`tAD . . ." -ForegroundColor DarkGray
$ADUsers = Get-ADUser -server $ADServer -SearchBase $SearchBase -Filter * -Properties * 
#Get All Azure AD users
Write-Host "`tAzure AD . . ." -ForegroundColor DarkGray
$AzureUsers = Get-AzureADUser -All:$true
#Get All Exchange  Mailboxes
Write-Host "`tExchange Mailboxes . . ." -ForegroundColor DarkGray
$EPMailboxes = Get-EPMailbox -ResultSize unlimited
Write-Host "`tExchange Mailboxes Statistics . . ." -ForegroundColor DarkGray
$EPMailboxesStats = $EPMailboxes | Get-EPMailboxStatistics
Write-Host "`tExchange Mailboxes Permissions . . ." -ForegroundColor DarkGray
$EPMailboxesPerms = $EPMailboxes | Get-EPMailboxPermission
Write-Host "`tExchange Archive Mailboxes . . ." -ForegroundColor DarkGray
$EPMailboxesArchive = Get-EPMailbox -ResultSize unlimited -Archive
Write-Host "`tExchange Archive Mailboxes Statistics . . ." -ForegroundColor DarkGray
$EPMailboxesArchiveStats = $EPMailboxesArchive | Get-EPMailboxStatistics -Archive
Write-Host "`tExchange Archive Mailboxes Permission . . ." -ForegroundColor DarkGray
$EPMailboxesArchivePerms = $EPMailboxesArchive | Get-EPMailboxPermission
Write-Host "`tExchange Remove Mailboxes . . ." -ForegroundColor DarkGray
$EPRemoteMailboxes = Get-EPRemoteMailbox -ResultSize unlimited
$EPRemoteMailboxesArchive = Get-EPRemoteMailbox -ResultSize unlimited -Archive
#Get All Exchange Online Mailboxes
Write-Host "`tExchange Online Mailboxes . . ." -ForegroundColor DarkGray
$EXOMailboxes = Get-Mailbox -ResultSize unlimited
Write-Host "`tExchange Online Mailboxes Statistics . . ." -ForegroundColor DarkGray
$EXOMailboxesStats = $EXOMailboxes | Get-MailboxStatistics
Write-Host "`tExchange Online Mailboxes Permissions . . ." -ForegroundColor DarkGray
$EXOMailboxesPerms = $EXOMailboxes | Get-MailboxPermission
Write-Host "`tExchange Online Archive Mailboxes . . ." -ForegroundColor DarkGray
$EXOMailboxesArchive = Get-Mailbox -ResultSize unlimited -Archive
Write-Host "`tExchange Online Archive Mailboxes Statistics . . ." -ForegroundColor DarkGray
$EXOMailboxesArchiveStats = $EXOMailboxesArchive | Get-MailboxStatistics -Archive
Write-Host "`tExchange Online Archive Mailboxes Permission . . ." -ForegroundColor DarkGray
$EXOMailboxesArchivePerms = $EXOMailboxesArchive | Get-MailboxPermission

#Get All Exchange CAS Mailbox Settings
$EPCASM = Get-EPCASMailbox
$EPRemoteCASM = Get-CASMailbox

$swc.Stop()

Write-Host ("Caching Objects time: " + (FormatElapsedTime($swc.Elapsed)) + " to run. " + '{0:N0}' -f ($ADUsers.Count / $swc.Elapsed.TotalMinutes) + " Users's per Minute.")

$swp = [Diagnostics.Stopwatch]::StartNew()
Write-host "Processing Users Please Wait . .."
Write-Progress -Activity ("Creating " + ($MyInvocation.MyCommand.Name -replace ".ps1","") + " output") -Status ("Processing " + "[" + $output.Count + "/" + $ADUsers.count + "]") -percentComplete 0 -Id 0 
#Main loop
Foreach ($ADUser in $ADUsers) { 
	# Write-host ("`t" + $ADUser.name)
	#Update progress Source: https://stackoverflow.com/questions/67981500/making-an-powershell-progress-bar-more-efficient 
	$Script:WindowWidthChanged = $Script:WindowWidth -ne $Host.UI.RawUI.WindowSize.Width
	if ($Script:WindowWidthChanged) { $Script:WindowWidth = $Host.UI.RawUI.WindowSize.Width }
	$ProgressCompleted = [math]::floor($output.Count * $Script:WindowWidth / $ADUsers.count)
	if ($Script:WindowWidthChanged -or $ProgressCompleted -ne $Script:LastProgressCompleted) {
		Write-Progress -Activity ("Creating " + ($MyInvocation.MyCommand.Name -replace ".ps1","") + " output") -Status ("Processing " + $ADUser.name + "[" + $output.Count + "/" + $ADUsers.count + "]") -percentComplete (($output.Count / $ADUsers.count)  * 100) -Id 0
	}
	$Script:LastProgressCompleted = $ProgressCompleted
	#Clear Loop var
	$TotalItemSize = 0
	$TotalDeletedItemSize = 0

	$AzureADUser = $AzureUsers.Where({$_.UserPrincipalName -eq $ADUser.UserPrincipalName})
	$Record = [ADExchangeOutput]::new()
	$Record."Logon Name" = $ADUser.sAMAccountName
	$Record."Display Name" = $ADUser.DisplayName
	$Record."Last Name" = $ADUser.Surname
	$Record."Middle Name" = $ADUser.middleName
	$Record."First Name" = $ADUser.GivenName
	$Record."Full address" = $ADUser.StreetAddress
	$Record."City" = $ADUser.City
	$Record."State" = $ADUser.st
	$Record."Postal Code" = $ADUser.PostalCode
	$Record."Country-Region" = $ADUser.Country
	$Record."Job Title" = $ADUser.Title
	$Record."Company" = $ADUser.Company
	$Record."Description" = $ADUser.Description
	$Record."Department" = $ADUser.Department
	$Record."Department Number" = ($ADUser.DepartmentCode -join ", ")
	$Record."Employee Type" = $ADUser.employeeType
	$Record."Employee Number" = $ADUser.EmployeeNumber
	$Record."Office" = $ADUser.physicalDeliveryOfficeName
	$Record."Phone" = $ADUser.telephoneNumber
	$Record."Mobile Phone" = $ADUser.mobile
	$Record."extensionAttribute2 CIFX Company" = $ADUser.extensionAttribute2
	If($AzureADUser){
		$Record."Azure Licenses" = ((Get-AzureADUserLicenseDetail -ObjectId $AzureADUser.ObjectId | Select-Object SkuPartNumber).SkuPartNumber -join ",")
		$Record."Azure Last Sync Time" = $AzureADUser.LastDirSyncTime
	}
	$Record."Group Membership" = (($ADUser.MemberOf | CleanDistinguishedName) -join ",")
	$Record."Email" = $ADUser.Mail
	$Record."Email Addresses" = (($ADUser.proxyAddresses).Where({$_ -match "smtp"})) -replace "smtp:" -join ", "

	#region Manager name
	If($ADUser.Manager){
		$Record."Manager" = (($ADUsers).Where({$_.DistinguishedName -eq $ADUser.Manager})).DisplayName
	}
	#endregion Manager name
	$Record."Home Directory" = $ADUser.homeDirectory
	#region Account Status
	if ($ADUser.Enabled -eq 'TRUE') {
		$Record."Account Status" = 'Enabled'
	}Else{
		$Record."Account Status" = 'Disabled'
	}
	#endregion Account Status
	#region Password Never Expires
	if ($ADUser.passwordNeverExpires -eq 'TRUE') {
		$Record."Password Never Expires" = 'Enabled'
	}Else{
		$Record."Password Never Expires" = 'Disabled'
	}
	#endregion Password Never Expires
	#region Password Not Required
	if ($ADUser.PasswordNotRequired -eq 'TRUE') {
		$Record."Password Not Required" = 'Enabled'
	}Else{
		$Record."Password Not Required" = 'Disabled'
	}
	#endregion Password Not Required
	#region Smartcard Logon Required
	if ($ADUser.SmartcardLogonRequired -eq 'TRUE') {
		$Record."Smartcard Logon Required" = 'Enabled'
	}Else{
		$Record."Smartcard Logon Required" = 'Disabled'
	}
	#endregion Smartcard Logon Required
	$Record."Last Log-On Date" = ([DateTime]::FromFileTime($ADUser.lastLogon))
	$Record."Days Since Last Log-On" = ($((Get-Date)- ([DateTime]::FromFileTime($ADUser.lastLogon))).Days)
	$Record."Creation Date" = $ADUser.whencreated
	$Record."Days Since Creation" = ($((Get-Date) - ([DateTime]($ADUser.whencreated))).Days)
	$Record."Last Password Change" = ([DateTime]::FromFileTime($ADUser.pwdLastSet))
	$Record."Days from last password change" = ($((Get-Date) - ([DateTime]::FromFileTime($ADUser.pwdLastSet))).Days)
	$Record."RDS CAL Expiration Date" = ($ADUser.msTSExpireDate)
	$Record."Distinguished Name" = ($ADUser.distinguishedName)

	#region E-Mail Information
	#region Mailbox
	If($ADUser.msExchMailboxGuid){
		$Record."Mailbox Creation Date" = $ADUser.msExchWhenMailboxCreated
		$LM = $EPMailboxes.Where({ $_.UserPrincipalName -eq $ADUser.UserPrincipalName})
		If ($LM) {
			$Record."Mailbox Location" = "Local"
			$Record."Mailbox Server" = $LM.ServerName
			$Record."Mailbox Database" = $LM.Database
			$Record."Mailbox GUID" = $LM.ExchangeGuid
			$Record."Mailbox Use Database Quota Defaults" = $LM.UseDatabaseQuotaDefaults
			$Record."Mailbox Issue Warning Quota" = MailboxGB($LM.IssueWarningQuota)
			$Record."Mailbox Prohibit Send Quota" = MailboxGB($LM.ProhibitSendQuota)
			$Record."Mailbox Prohibit Send Receive Quota" = MailboxGB($LM.ProhibitSendReceiveQuota)

			#$LMS =  Get-EPMailboxStatistics -Identity $ADUser.UserPrincipalName -ErrorAction SilentlyContinue
			$LMS =  $EPMailboxesStats.Where({$_.Identity -eq $ADUser.UserPrincipalName})
			If($LMS){
				$Record."Mailbox Storage Limit Status" = $LMS.StorageLimitStatus
				$TotalItemSize = MailboxGB($LMS.TotalItemSize)
				$TotalDeletedItemSize = MailboxGB($LMS.TotalDeletedItemSize) 
				If ($TotalItemSize -ne "Unlimited" -and $TotalDeletedItemSize -ne "Unlimited" ) {
					$Record."Mailbox Size" = ($TotalItemSize + $TotalDeletedItemSize)
				}
				$Record."Mailbox Item Count" = $LMS.ItemCount
				$Record."Mailbox Last Logged On User Account" = $LMS.LastLoggedOnUserAccount
				$Record."Mailbox Last Logon Time" = $LMS.LastLogonTime
				$Record."Mailbox Last Logoff Time" = $LMS.LastLogoffTime
			}
			$LMCAS = $EPCASM.Where({ $_.SamAccountName.ToLower() -eq $ADUser.SamAccountName.ToLower()})
			If ($LMCAS) {
				$Record."OWA Enabled" = $LMCAS.OWAEnabled
				$Record."Mapi Enabled" = $LMCAS.ActiveSyncEnabled
				$Record."Active Sync Enabled" = $LMCAS.MapiEnabled
				$Record."IMAP Enabled" = $LMCAS.ImapEnabled
				$Record."POP Enabled" = $LMCAS.PopEnabled
			}

		}Else {
			$RM = $EXOMailboxes.Where({ $_.UserPrincipalName -eq $ADUser.UserPrincipalName})
			$LRM = $EPRemoteMailboxes.Where({ $_.UserPrincipalName -eq $ADUser.UserPrincipalName})
			If ($RM) {
				If ($RM -and $LRM) {
					If ($RM.ExchangeGuid -eq $LRM.ExchangeGuid) {
						$Record."Mailbox Location" = "Hybrid Remote"
					}Else {
						$Record."Mailbox Location" = "Hybrid Remote Broken"
						If (($LRM.ExchangeGuid -eq '00000000-0000-0000-0000-000000000000' -or $null -eq $LRM.ExchangeGuid)) {
							write-host ("`t Creating Remote Mailbox: " + $ADUser.ExchangeGuid) -ForeGroundColor red
							Enable-EPRemoteMailbox -Identity $ADUser.Alias -RemoteRoutingAddress ( $ADUser.Alias + "@github.mail.onmicrosoft.com")
							If (($LRM.ArchiveGuid -eq '00000000-0000-0000-0000-000000000000' -or $null -eq $LRM.ArchiveGuid) -and ($ADUser.ArchiveGuid -eq '00000000-0000-0000-0000-000000000000' -or $null -eq $ADUser.ArchiveGuid)) {
								write-host ("`t Setting ExchangeGUID: " + $ADUser.ExchangeGuid) -foregroundcolor yellow
								Set-EPRemoteMailbox -Identity $RM.Alias -ExchangeGUID $RM.ExchangeGuid
							}Else{
								write-host ("`t Setting ExchangeGUID: " + $ADUser.ExchangeGuid) -foregroundcolor yellow
								Set-EPRemoteMailbox -Identity $RM.Alias -ExchangeGUID $RM.ExchangeGuid -ArchiveGuid $RM.ArchiveGuid
							}
						}
					}
				}Else{
					$Record."Mailbox Location" = "Cloud Only"
				}
				$Record."Mailbox Server" = $RM.ServerName
				$Record."Mailbox Database" = $RM.Database
				$Record."Mailbox GUID" = $RM.ExchangeGuid
				$Record."Mailbox Use Database Quota Defaults" = $RM.UseDatabaseQuotaDefaults
				$Record."Mailbox Issue Warning Quota" = MailboxGB($RM.IssueWarningQuota) 
				$Record."Mailbox Prohibit Send Quota" = MailboxGB($RM.ProhibitSendQuota) 
				$Record."Mailbox Prohibit Send Receive Quota" = MailboxGB($RM.ProhibitSendReceiveQuota) 
				#$RMS =  Get-MailboxStatistics -Identity $ADUser.UserPrincipalName -ErrorAction SilentlyContinue
				$RMS =  $EXOMailboxesStats.Where({$_.Identity -eq $ADUser.UserPrincipalName})
				If($RMS){
					$Record."Mailbox Storage Limit Status" = $RMS.StorageLimitStatus
					$TotalItemSize = MailboxGB($RMS.TotalItemSize)
					$TotalDeletedItemSize = MailboxGB($RMS.TotalDeletedItemSize) 
					If ($TotalItemSize -ne "Unlimited" -and $TotalDeletedItemSize -ne "Unlimited" ) {
						$Record."Mailbox Size" = ($TotalItemSize + $TotalDeletedItemSize)
					}
					$Record."Mailbox Item Count" = $RMS.ItemCount
					$Record."Mailbox Last Logged On User Account" = $RMS.LastLoggedOnUserAccount
					$Record."Mailbox Last Logon Time" = $RMS.LastLogonTime
					$Record."Mailbox Last Logoff Time" = $RMS.LastLogoffTime
				}
				$RMCAS = $EPRemoteCASM.Where({ $_.SamAccountName.ToLower() -eq $ADUser.SamAccountName.ToLower()})
				If ($LMCAS) {
					$Record."OWA Enabled" = $RMCAS.OWAEnabled
					$Record."Mapi Enabled" = $RMCAS.ActiveSyncEnabled
					$Record."Active Sync Enabled" = $RMCAS.MapiEnabled
					$Record."IMAP Enabled" = $RMCAS.ImapEnabled
					$Record."POP Enabled" = $RMCAS.PopEnabled
				}
			}			
		}
		$DAMP = @()
		$CSVFixedDAMP = $null
		$FixedDAMP = @{}
		#region Mailbox Permission
		If ($Record."Mailbox Location" -eq "Local" -and $LM) {
			# $DAMP = Get-EPMailboxPermission  -Identity $ADUser.UserPrincipalName | Where-Object {($_.AccessRights -eq "FullAccess") -and ($_.User -notin $ExcludeUsers) -and ($_.User -notmatch "S-1-5-*" ) -and $_.IsInherited -eq $false} | Sort-Object -Unique -Property User
			$DAMP = ($EPMailboxesPerms.Where({$_.Identity -eq $ADUser.UserPrincipalName})).Where({($_.AccessRights -eq "FullAccess") -and ($_.User -notin $ExcludeUsers) -and ($_.User -notmatch "S-1-5-*" ) -and $_.IsInherited -eq $false}) | Sort-Object -Unique -Property User
		}
		If($Record."Mailbox Location" -eq "Cloud Only" -or $Record."Mailbox Location" -match "Hybrid Remote" -and $RM){
			# $DAMP =  Get-MailboxPermission -Identity $ADUser.UserPrincipalName | Where-Object {($_.AccessRights -eq "FullAccess") -and ($_.User -notin $ExcludeUsers) -and ($_.User -notmatch "S-1-5-*" ) -and $_.IsInherited -eq $false} | Sort-Object -Unique -Property User		
			$DAMP =  ($EXOMailboxesPerms.Where({$_.Identity -eq $ADUser.UserPrincipalName})).Where({($_.AccessRights -eq "FullAccess") -and ($_.User -notin $ExcludeUsers) -and ($_.User -notmatch "S-1-5-*" ) -and $_.IsInherited -eq $false}) | Sort-Object -Unique -Property User		
		}
		#Will remove disabled users with permissions
		If ($DAMP.count -gt 0 ) {
			ForEach( $ACE in $DAMP) {
				[switch]$ADValid=$false
				If (($env:USERDOMAIN).ToLower() -eq (split-path -Path $ACE.User -Parent).ToLower()) {
					Try{
						$DomainObject = [DirectoryServices.AccountManagement.Principal]::FindByIdentity($ContextType,$IdentityType,(split-path -Path $ACE.User -Leaf))
						Foreach ($DO in $DomainObject) {
							If ($DO.StructuralObjectClass -eq  "user") {
								If ($DO.Enabled -and $DO.SamAccountName.ToLower() -eq (split-path -Path $ACE.User -Leaf).ToLower()) {
									$ADValid=$true
								}
							}elseif ($DO.StructuralObjectClass -eq  "group") {
								$ADValid=$true
							}
						}
					}Catch{

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
		
			If ($FixedDAMP.count -gt 0) {
				#Loop though all perms and add to string
				Foreach ($FDAMP in $FixedDAMP.GetEnumerator()) {
					If ($FDAMP.Key) {
						$CSVFixedDAMP = ($CSVFixedDAMP + "," + $FDAMP.Key + " - " + (($FDAMP.Value | Out-String) -join " ") )
					}else{
						If ($FDAMP.Keys) {
							$CSVFixedDAMP = ($CSVFixedDAMP + "," + $FDAMP.Keys + " - " + (($FDAMP.Values | Out-String) -join " ") )
						}
					}
				}
				#clean up first ; in string and remove all new lines in string
				If ($CSVFixedDAMP) {
					If ($CSVFixedDAMP.substring(0,1) -eq ",") {
						$CSVFixedDAMP = $CSVFixedDAMP.substring(1,$CSVFixedDAMP.Length -1 ) -replace "`n|`r"
					}
					$Record."Mailbox Permissions" = $CSVFixedDAMP
				}
			}
		}
	}
		#endregion Mailbox Permission
	#endregion Mailbox	
	#region Mailbox Archive
	If($ADUser.msExchArchiveGUID){		
		$Record."Mailbox Archive Creation Date" = $ADUser.msExchWhenMailboxCreated
		$LM = $EPMailboxesArchive.Where({ $_.UserPrincipalName -eq $ADUser.UserPrincipalName})
		If ($LM) {
			$Record."Mailbox Archive Location" = "Local"
			$Record."Mailbox Archive Server" = $LM.ServerName
			$Record."Mailbox Archive Database" = $LM.Database
			$Record."Mailbox Archive GUID" = $LM.ExchangeGuid
			$Record."Mailbox Archive Use Database Quota Defaults" = $LM.UseDatabaseQuotaDefaults
			$Record."Mailbox Archive Issue Warning Quota" = MailboxGB($LM.IssueWarningQuota)
			$Record."Mailbox Archive Prohibit Send Quota" = MailboxGB($LM.ProhibitSendQuota)
			$Record."Mailbox Archive Prohibit Send Receive Quota" = MailboxGB($LM.ProhibitSendReceiveQuota)

			# $LMS =  Get-EPMailboxStatistics -Archive -Identity $ADUser.UserPrincipalName -ErrorAction SilentlyContinue
			$LMS =  $EPMailboxesArchiveStats.Where({$_.Identity -eq $ADUser.UserPrincipalName})
			If($LMS){
				$Record."Mailbox Archive Storage Limit Status" = $LMS.StorageLimitStatus
				$TotalItemSize = MailboxGB($LMS.TotalItemSize)
				$TotalDeletedItemSize = MailboxGB($LMS.TotalDeletedItemSize) 
				If ($TotalItemSize -ne "Unlimited" -and $TotalDeletedItemSize -ne "Unlimited" ) {
					$Record."Mailbox Archive Size" = ($TotalItemSize + $TotalDeletedItemSize)
				}

				$Record."Mailbox Archive Item Count" = $LMS.ItemCount
				$Record."Mailbox Archive Last Logged On User Account" = $LMS.LastLoggedOnUserAccount
				$Record."Mailbox Archive Last Logon Time" = $LMS.LastLogonTime
				$Record."Mailbox Archive Last Logoff Time" = $LMS.LastLogoffTime
			}
		}Else {
			$RM = $EXOMailboxesArchive.Where({ $_.UserPrincipalName -eq $ADUser.UserPrincipalName})
			$LRM = $EPRemoteMailboxesArchive.Where({ $_.UserPrincipalName -eq $ADUser.UserPrincipalName})
			If ($RM) {
				If ($RM -and $LRM) {
					If ($RM.ExchangeGuid -eq $LRM.ExchangeGuid) {
						$Record."Mailbox Archive Location" = "Hybrid Remote"
					}Else {
						$Record."Mailbox Archive Location" = "Hybrid Remote Broken"			
					}
				}Else{
					$Record."Mailbox Archive Location" = "Cloud Only"
				}
				$Record."Mailbox Archive Server" = $RM.ServerName
				$Record."Mailbox Archive Database" = $RM.Database
				$Record."Mailbox Archive GUID" = $RM.ExchangeGuid
				$Record."Mailbox Archive Use Database Quota Defaults" = $RM.UseDatabaseQuotaDefaults
				$Record."Mailbox Archive Issue Warning Quota" = MailboxGB($RM.IssueWarningQuota) 
				$Record."Mailbox Archive Prohibit Send Quota" = MailboxGB($RM.ProhibitSendQuota) 
				$Record."Mailbox Archive Prohibit Send Receive Quota" = MailboxGB($RM.ProhibitSendReceiveQuota) 
				# $RMS =  Get-MailboxStatistics -Archive -Identity $ADUser.UserPrincipalName -ErrorAction SilentlyContinue
				$RMS = $EXOMailboxesArchiveStats.Where({$_.Identity -eq $ADUser.UserPrincipalName})
				If($RMS){
					$Record."Mailbox Archive Storage Limit Status" = $RMS.StorageLimitStatus
					$TotalItemSize = MailboxGB($RMS.TotalItemSize)
					$TotalDeletedItemSize = MailboxGB($RMS.TotalDeletedItemSize) 
					If ($TotalItemSize -ne "Unlimited" -and $TotalDeletedItemSize -ne "Unlimited" ) {
						$Record."Mailbox Archive Size" = ($TotalItemSize + $TotalDeletedItemSize)
					}
					$Record."Mailbox Archive Item Count" = $RMS.ItemCount
					$Record."Mailbox Archive Last Logged On User Account" = $RMS.LastLoggedOnUserAccount
					$Record."Mailbox Archive Last Logon Time" = $RMS.LastLogonTime
					$Record."Mailbox Archive Last Logoff Time" = $RMS.LastLogoffTime
				}
			}			
		}
		$DAMP = $null
		$CSVFixedDAMP = $null
		$FixedDAMP = @{}
		#region Mailbox Permission
		If ($Record."Mailbox Archive Location" -eq "Local") {
			# $DAMP = $LM | Get-EPMailboxPermission | Where-Object {($_.AccessRights -eq "FullAccess") -and ($_.User -notin $ExcludeUsers) -and ($_.User -notmatch "S-1-5-*" ) -and $_.IsInherited -eq $false} | Sort-Object -Unique -Property User
			$DAMP = ($EPMailboxesArchivePerms.Where({$_.Identity -eq $LM.Identity})).Where({($_.AccessRights -eq "FullAccess") -and ($_.User -notin $ExcludeUsers) -and ($_.User -notmatch "S-1-5-*" ) -and $_.IsInherited -eq $false}) | Sort-Object -Unique -Property User
		}
		If($Record."Mailbox Archive Location" -eq "Cloud Only" -or $Record."Mailbox Location" -match "Hybrid Remote"){
			# $DAMP = $RM | Get-MailboxPermission | Where-Object {($_.AccessRights -eq "FullAccess") -and ($_.User -notin $ExcludeUsers) -and ($_.User -notmatch "S-1-5-*" ) -and $_.IsInherited -eq $false} | Sort-Object -Unique -Property User		
			$DAMP = ($EXOMailboxesArchivePerms.Where({$_.Identity -eq $RM.Identity})).Where({($_.AccessRights -eq "FullAccess") -and ($_.User -notin $ExcludeUsers) -and ($_.User -notmatch "S-1-5-*" ) -and $_.IsInherited -eq $false}) | Sort-Object -Unique -Property User		
		}
		#Will remove disabled users with permissions
		If ($DAMP.count -gt 0 ) {
			ForEach( $ACE in $DAMP) {
				[switch]$ADValid=$false
				If (($env:USERDOMAIN).ToLower() -eq (split-path -Path $ACE.User -Parent).ToLower()) {
					Try{
						$DomainObject = [DirectoryServices.AccountManagement.Principal]::FindByIdentity($ContextType,$IdentityType,(split-path -Path $ACE.User -Leaf))
						Foreach ($DO in $DomainObject) {
							If ($DO.StructuralObjectClass -eq  "user") {
								If ($DO.Enabled -and $DO.SamAccountName.ToLower() -eq (split-path -Path $ACE.User -Leaf).ToLower()) {
									$ADValid=$true
								}
							}elseif ($DO.StructuralObjectClass -eq  "group") {
								$ADValid=$true
							}
						}
					} Catch {

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
		
			If ($FixedDAMP.count -gt 0) {
				#Loop though all perms and add to string
				Foreach ($FDAMP in $FixedDAMP.GetEnumerator()) {
					If ($FDAMP.Key) {
						$CSVFixedDAMP = ($CSVFixedDAMP + "," + $FDAMP.Key + " - " + (($FDAMP.Value | Out-String) -join " ") )
					}else{
						If ($FDAMP.Keys) {
							$CSVFixedDAMP = ($CSVFixedDAMP + "," + $FDAMP.Keys + " - " + (($FDAMP.Values | Out-String) -join " ") )
						}
					}
				}
				#clean up first ; in string and remove all new lines in string
				If ($CSVFixedDAMP) {
					If ($CSVFixedDAMP.substring(0,1) -eq ",") {
						$CSVFixedDAMP = $CSVFixedDAMP.substring(1,$CSVFixedDAMP.Length -1 ) -replace "`n|`r"
					}
					$Record."Mailbox Archive Permissions" = $CSVFixedDAMP
				}
			}
		}
	}
	#endregion Mailbox Archive
	#endregion E-Mail Information

	$output += $Record
}

Write-Progress -Activity ("Creating " + ($MyInvocation.MyCommand.Name -replace ".ps1","") + " output") -Completed
$swp.Stop()
Write-Host ("Processing users time: " + (FormatElapsedTime($swp.Elapsed)) + " to run. " + '{0:N0}' -f ($ADUsers.Count / $swp.Elapsed.TotalMinutes) + " Users's per Minute.")

$swo = [Diagnostics.Stopwatch]::StartNew()
Write-Host "Saving Output"
$output | Export-Csv -Path $csvfile -NoTypeInformation


#region Load ImportExcel
If(-Not (Get-Module -Name ImportExcel -ListAvailable)){
	Install-Module -Name ImportExcel -Force -Confirm:$false
}
If (-Not (Get-Module "ImportExcel" -ErrorAction SilentlyContinue)) {
	Import-Module ImportExcel
}   
#endregion Load ImportExcel
#region Excel convert
$excel = $output | Export-Excel -Path $xlsxfile -ClearSheet -WorksheetName ("Hybrid_Info_" + $FileDate) -AutoFilter -AutoSize -FreezeTopRowFirstColumn -PassThru
$ws = $excel.Workbook.Worksheets[("Hybrid_Info_" + $FileDate)]
$LastRow = $ws.Dimension.End.Row
$LastColumn = $ws.Dimension.End.column

#Header Lookup
$htHeader =[ordered]@{}
for ($i = 1; $i -le  $LastColumn; $i++) {
	$htHeader.add(($ws.Cells[1,$i].value),$i)
}

#Setup
Set-ExcelRange -Worksheet $ws -Range ("A1:" + $LastColumn + $LastRow) -VerticalAlignment Top


#Days Since Last LogOn
# Add-ConditionalFormatting -WorkSheet $ws -address (($ws.Cells[1,$htHeader["Days Since Last LogOn"]].Address).Substring(0,($ws.Cells[1,$htHeader["Days Since Last LogOn"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Days Since Last LogOn"]].Address).Substring(0,($ws.Cells[1,$htHeader["Days Since Last LogOn"]].Address).Length-1) + $LastRow) -RuleType TwoColorScale

#Azure Licenses
$StrHeader= "Azure Licenses"
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader[$StrHeader]].Address).Substring(0,($ws.Cells[1,$htHeader[$StrHeader]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader[$StrHeader]].Address).Substring(0,($ws.Cells[1,$htHeader[$StrHeader]].Address).Length-1) + $LastRow) -WrapText 
#Replace ", " with "`r`n"
($ws.Cells[(($ws.Cells[1,$htHeader[$StrHeader]].Address).Substring(0,($ws.Cells[1,$htHeader[$StrHeader]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader[$StrHeader]].Address).Substring(0,($ws.Cells[1,$htHeader[$StrHeader]].Address).Length-1) + $LastRow)]).foreach({$_.Value = $_.Value -replace ",","`r`n"})

#Group Membership
$StrHeader= "Group Membership"
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader[$StrHeader]].Address).Substring(0,($ws.Cells[1,$htHeader[$StrHeader]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader[$StrHeader]].Address).Substring(0,($ws.Cells[1,$htHeader[$StrHeader]].Address).Length-1) + $LastRow) -WrapText 
#Replace ", " with "`r`n"
($ws.Cells[(($ws.Cells[1,$htHeader[$StrHeader]].Address).Substring(0,($ws.Cells[1,$htHeader[$StrHeader]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader[$StrHeader]].Address).Substring(0,($ws.Cells[1,$htHeader[$StrHeader]].Address).Length-1) + $LastRow)]).foreach({$_.Value = $_.Value -replace ",","`r`n"})
#Email Addresses
$StrHeader= "Email Addresses"
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader[$StrHeader]].Address).Substring(0,($ws.Cells[1,$htHeader[$StrHeader]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader[$StrHeader]].Address).Substring(0,($ws.Cells[1,$htHeader[$StrHeader]].Address).Length-1) + $LastRow) -WrapText 
#Replace ", " with "`r`n"
($ws.Cells[(($ws.Cells[1,$htHeader[$StrHeader]].Address).Substring(0,($ws.Cells[1,$htHeader[$StrHeader]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader[$StrHeader]].Address).Substring(0,($ws.Cells[1,$htHeader[$StrHeader]].Address).Length-1) + $LastRow)]).foreach({$_.Value = $_.Value -replace ",","`r`n"})

#Mailbox Permissions
$StrHeader= "Mailbox Permissions"
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader[$StrHeader]].Address).Substring(0,($ws.Cells[1,$htHeader[$StrHeader]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader[$StrHeader]].Address).Substring(0,($ws.Cells[1,$htHeader[$StrHeader]].Address).Length-1) + $LastRow) -WrapText 
#Replace ", " with "`r`n"
($ws.Cells[(($ws.Cells[1,$htHeader[$StrHeader]].Address).Substring(0,($ws.Cells[1,$htHeader[$StrHeader]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader[$StrHeader]].Address).Substring(0,($ws.Cells[1,$htHeader[$StrHeader]].Address).Length-1) + $LastRow)]).foreach({$_.Value = $_.Value -replace ",","`r`n"})

#Mailbox Archive Permissions
$StrHeader= "Mailbox Archive Permissions"
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader[$StrHeader]].Address).Substring(0,($ws.Cells[1,$htHeader[$StrHeader]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader[$StrHeader]].Address).Substring(0,($ws.Cells[1,$htHeader[$StrHeader]].Address).Length-1) + $LastRow) -WrapText 
#Replace ", " with "`r`n"
($ws.Cells[(($ws.Cells[1,$htHeader[$StrHeader]].Address).Substring(0,($ws.Cells[1,$htHeader[$StrHeader]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader[$StrHeader]].Address).Substring(0,($ws.Cells[1,$htHeader[$StrHeader]].Address).Length-1) + $LastRow)]).foreach({$_.Value = $_.Value -replace ",","`r`n"})


#Mailbox Item Count comma's
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["Days Since Last Log-On"]].Address).Substring(0,($ws.Cells[1,$htHeader["Days Since Last Log-On"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Days Since Last Log-On"]].Address).Substring(0,($ws.Cells[1,$htHeader["Days Since Last Log-On"]].Address).Length-1) + $LastRow) -NumberFormat '#,##0'

#Mailbox Size (GB) comma's
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["Mailbox Size"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Size"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Mailbox Size"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Size"]].Address).Length-1) + $LastRow) -NumberFormat '#,##0'
#Mailbox Item Count comma's
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["Mailbox Item Count"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Item Count"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Mailbox Item Count"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Item Count"]].Address).Length-1) + $LastRow) -NumberFormat '#,##0'
#Mailbox Last Logged On User Account Date
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["Mailbox Last Logged On User Account"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Last Logged On User Account"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Mailbox Last Logged On User Account"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Last Logged On User Account"]].Address).Length-1) + $LastRow) -NumberFormat 'Date-Time'
#Mailbox Last Logon Time Date
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["Mailbox Last Logon Time"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Last Logon Time"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Mailbox Last Logon Time"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Last Logon Time"]].Address).Length-1) + $LastRow) -NumberFormat 'Date-Time'
#Mailbox Last Logoff Time Date
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["Mailbox Last Logoff Time"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Last Logoff Time"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Mailbox Last Logoff Time"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Last Logoff Time"]].Address).Length-1) + $LastRow) -NumberFormat 'Date-Time'
#Mailbox Creation Date Date
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["Mailbox Creation Date"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Creation Date"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Mailbox Creation Date"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Creation Date"]].Address).Length-1) + $LastRow) -NumberFormat 'Date-Time'




#Mailbox Archive Size (GB) comma's
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["Mailbox Archive Size"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Archive Size"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Mailbox Archive Size"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Archive Size"]].Address).Length-1) + $LastRow) -NumberFormat '#,##0'
#Mailbox Archive Item Count comma's
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["Mailbox Archive Item Count"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Archive Item Count"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Mailbox Archive Item Count"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Archive Item Count"]].Address).Length-1) + $LastRow) -NumberFormat '#,##0'
#Mailbox Archive Last Logged On User Account Date
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["Mailbox Archive Last Logged On User Account"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Archive Last Logged On User Account"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Mailbox Archive Last Logged On User Account"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Archive Last Logged On User Account"]].Address).Length-1) + $LastRow) -NumberFormat 'Date-Time'
#Mailbox Archive Last Logon Time Date
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["Mailbox Archive Last Logon Time"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Archive Last Logon Time"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Mailbox Archive Last Logon Time"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Archive Last Logon Time"]].Address).Length-1) + $LastRow) -NumberFormat 'Date-Time'
#Mailbox Archive Last Logoff Time Date
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["Mailbox Archive Last Logoff Time"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Archive Last Logoff Time"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Mailbox Archive Last Logoff Time"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Archive Last Logoff Time"]].Address).Length-1) + $LastRow) -NumberFormat 'Date-Time'
#Mailbox Archive Creation Date Date
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["Mailbox Archive Creation Date"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Archive Creation Date"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Mailbox Archive Creation Date"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Archive Creation Date"]].Address).Length-1) + $Lastrow) -NumberFormat 'Date-Time'


#Last Log-On Date
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["Days Since Last Log-On"]].Address).Substring(0,($ws.Cells[1,$htHeader["Days Since Last Log-On"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Days Since Last Log-On"]].Address).Substring(0,($ws.Cells[1,$htHeader["Days Since Last Log-On"]].Address).Length-1) + $LastRow) -NumberFormat 'Date-Time'
#Days Since Last Log-On comma's
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["Days Since Last Log-On"]].Address).Substring(0,($ws.Cells[1,$htHeader["Days Since Last Log-On"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Days Since Last Log-On"]].Address).Substring(0,($ws.Cells[1,$htHeader["Days Since Last Log-On"]].Address).Length-1) + $LastRow) -NumberFormat '#,##0'
#Creation Date Date
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["Creation Date"]].Address).Substring(0,($ws.Cells[1,$htHeader["Creation Date"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Creation Date"]].Address).Substring(0,($ws.Cells[1,$htHeader["Creation Date"]].Address).Length-1) + $LastRow) -NumberFormat 'Date-Time'
#Days Since Creation comma's
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["Days Since Creation"]].Address).Substring(0,($ws.Cells[1,$htHeader["Days Since Creation"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Days Since Creation"]].Address).Substring(0,($ws.Cells[1,$htHeader["Days Since Creation"]].Address).Length-1) + $LastRow) -NumberFormat '#,##0'
#Last Password Change Date
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["Last Password Change"]].Address).Substring(0,($ws.Cells[1,$htHeader["Last Password Change"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Last Password Change"]].Address).Substring(0,($ws.Cells[1,$htHeader["Last Password Change"]].Address).Length-1) + $LastRow) -NumberFormat 'Date-Time'
#Days from last password change comma's
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["Days from last password change"]].Address).Substring(0,($ws.Cells[1,$htHeader["Days from last password change"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Days from last password change"]].Address).Substring(0,($ws.Cells[1,$htHeader["Days from last password change"]].Address).Length-1) + $LastRow) -NumberFormat '#,##0'
#RDS CAL Expiration Date Date
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["RDS CAL Expiration Date"]].Address).Substring(0,($ws.Cells[1,$htHeader["RDS CAL Expiration Date"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["RDS CAL Expiration Date"]].Address).Substring(0,($ws.Cells[1,$htHeader["RDS CAL Expiration Date"]].Address).Length-1) + $LastRow) -NumberFormat 'Date-Time'

#New worksheet that has created,pw change, last logon all over 60 days.
$output.Where({$_."Days Since Last LogOn" -gt 60 -and $_."Days Since Creation" -gt 60 -and $_."Days from last password change" -gt 60}) | Export-Excel -Path $xlsxfile -ClearSheet -WorksheetName "Look to Disable" -AutoFilter -AutoSize -FreezeTopRowFirstColumn


Close-ExcelPackage $excel
$swo.Stop()
Write-Host ("Saving output time: " + (FormatElapsedTime($swo.Elapsed)) + " to run. ")
Remove-Variable "output"
Remove-Variable "excel"

$sw.Stop()
Write-Host ("Script runtime: " + (FormatElapsedTime($sw.Elapsed)) + " to run. " + '{0:N0}' -f ($ADUsers.Count / $sw.Elapsed.TotalMinutes) + " Users's per Minute.")
#endregion Excel convert

