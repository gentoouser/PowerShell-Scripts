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
    * Version 1.03.00 - Setup script to do Parallel caching of objects.
    * Version 1.03.01 - Logon to Exchange Online using App ID.
    * Version 1.03.02 - Logon to Azure using App ID.
    * Version 1.03.03 - Fixed issues with permissions and caching.
    * Version 1.03.04 - Fixed issue with disabled users with permission not showing. Also fixed issue with Mailbox size not showing. Clean Mailbox permission when a user with full permission is gone.
    * Version 1.03.05 - Added parse ADFS for last logon time. Use GraphAPI to get Entra ID last logon Time too.
    * Version 1.03.06 - Fixes to create more worksheets if filtered data
    * Version 1.03.07 - Added Auto-Forward field. Added inbox forwarding rules. Fixed Progress bar for Monitoring Jobs.
    * Version 1.03.08 - Added Azure MFA info. Fixed issues with GraphAPI
    * Version 1.03.09 - Fixed issue with interactive logon and modules not being installed. 
    * Version 1.03.10 - Fixed issue with MSGraph not getting data from Azure and not saving data in output.
    

#>
#region Parameters
Param(
	[CmdletBinding()]
    $csvfile =  ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\Logs\" + `
                ($MyInvocation.MyCommand.Name -replace ".ps1","") + "_" + `
                (Get-Date -format yyyyMMdd-hhmm) + ".csv"),
    $xlsxfile =  ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\Logs\" + `
                ($MyInvocation.MyCommand.Name -replace ".ps1","") + "_" + `
                (Get-Date -format yyyyMMdd-hhmm) + ".xlsx"),
    $ExcludeUsers=@(
		($env:USERDOMAIN + "\Domain Admins"),
		($env:USERDOMAIN + "\Enterprise Admins"),
		($env:USERDOMAIN + "\Organization Management"),
		($env:USERDOMAIN + "\Exchange Servers"),
		($env:USERDOMAIN + "\Exchange Domain Servers"),
		($env:USERDOMAIN + "\Exchange Services"),
		($env:USERDOMAIN + "\Exchange Trusted Subsystem"),
		($env:USERDOMAIN + "\Administrators"),
		($env:USERDOMAIN + "\Public Folder Management"),
		($env:USERDOMAIN + "\Delegated Setup"),
		($env:USERDOMAIN + "\Managed Availability Servers"),
		"NT AUTHORITY\SYSTEM",
		"NT AUTHORITY\SELF",
		"NT AUTHORITY\NETWORK SERVICE"

	),
    $ExchangeServer                      = "exchange.github.com",
	[string]$AzureTenant                 = "",
    [string]$AZClientID                  = "",
    [string]$AZCertThumbprint            = "",
	[String]$AZOrg                       =	"github.onmicrosoft.com",
	[switch]$RemoveDisabledPerms         =	$False,
	[switch]$IgnoreJobs                  =	$False,
	[array]$ADFSServers=@(
		"PD01.github.com"
		"PD02.github.com"
		"PD03.github.com"
		"PD04.github.com"
	)
)
#endregion Parameters
#region Variables 
$ScriptVersion = "1.3.10"
$output = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
$FileDate = (Get-Date -format yyyyMMdd-hhmm)
$sw = [Diagnostics.Stopwatch]::StartNew()
$CacheRefresh = 4
$Jobs = @()

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
	${User Principal Name}
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
	${Azure MFA Status}
	${Azure MFA Preferred method}
	${Azure Phone Authentication}
	${Azure Authenticator App}
	${Azure Passwordless}
	${Azure Hello for Business}
	${Azure FIDO2 Security Key}
	${Azure Temporary Access Pass}
	${Azure Authenticator device}
	${Azure Licenses}
	${Azure Licenses Details}
	${Azure Last Sync Time}
	${Azure Last Sign-On}
	${Azure Last Sign-On Days}
	${Azure Non-Interactive Last Sign-On}
	${Azure Non-Interactive Last Sign-On Days}
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
	${ADFS Last Logon} 
	${ADFS Last Logon Days} 
	${ADFS Last Logon IP} 
	${ADFS Relying Party} 
	${ADFS Auth Protocol} 
	${ADFS Network Location} 
	${ADFS ADFS Server} 
	${ADFS User Agent String} 
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
	${Mailbox Forwarding Address}
	${Mailbox Forwarding Address SMTP}
	${Mailbox Forwarding Rules}
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
#endregion Import AD modules
#region Create CSV Path
If (-Not( Test-Path (Split-Path -Path $csvfile -Parent))) {
	New-Item -ItemType directory -Path (Split-Path -Path $csvfile -Parent) | Out-Null
}
#endregion Create CSV Path

#region mailbox checking

# If (Get-ChildItem "Cert:\LocalMachine\My\" | Where-Object {$_.Thumbprint -eq $AZCertThumbprint}) {
# 	(Connect-AzureAD -TenantId $AzureTenant -ApplicationId  $AZClientID -CertificateThumbprint $AZCertThumbprint) | Out-Null
# }Else{
# 	(Connect-AzureAD -AccountId ($User + "@" + $Domain )) | Out-Null
# }
#Load .Net Assembly for AD
Add-Type -AssemblyName System.DirectoryServices.AccountManagement
$ContextType = [System.DirectoryServices.AccountManagement.ContextType]::Domain
$IdentityType = [System.DirectoryServices.AccountManagement.IdentityType]::SamAccountName

#Sets the OU to do the base search for all user accounts, change as required.
$SearchBase = (Get-ADDomain).DistinguishedName

#Define variable for a server with AD web services installed
$ADServer = (Get-ADDomain).PDCEmulator

#region caching objects
If ($null -eq $HMADITimeStamp -or ($HMADITimeStamp -and (New-TimeSpan -Start $HMADITimeStamp -End (Get-Date)).Hours -ge $CacheRefresh )) {
	$ADUsers = @()
	# $AzureUsers = @()
	$EPMailboxes = @()
	$EPMailboxesArchive = @()
	$EPRemoteMailboxes = @()
	$EPRemoteMailboxesArchive  = @()
	$EXOMailboxes = @()
	$EXOMailboxesArchive = @()
	$EPRemoteCASM = @()
	
	$EPMailboxesStats = @()
	$EPMailboxesPerms = @()
	$EPMailboxesForwardingRules = @()
	$EPMailboxesArchiveStats = @()
	$EPMailboxesArchivePerms = @()
	$EXOMailboxesStats = @()
	$EXOMailboxesPerms = @()
	$EXOMailboxesArchiveStats = @()
	$EXOMailboxesArchivePerms = @()
	$EXOMailboxesForwardingRules = @()
	$EPCASM = @()
	$ADFSLogs = @()
	# $OutputGraphAPIU = @()
	# $OutputAZUL = @()
	$OutputMgUser
	Write-Host ("Caching Objects please wait . . ." )
	$swc = [Diagnostics.Stopwatch]::StartNew()

	#region stop jobs from other runs
	If ((get-job).count -gt 0 -and $IgnoreJobs -eq $False) {
		Write-warning ("Removing jobs before starting new jobs. Job count: " + (get-job).count)
		get-job| remove-job -force
	}
	#region stop jobs from other runs


	#region ADFS Event logs
	Write-Host "`tADFS . . ." -ForegroundColor DarkGray
	$ADFSScript = {
		param(
			$ADFS
		)
		Class ADFSEventRecord {
			${Date-Time}
			${ADFS Server}
			${User ID}
			${Relying Party}
			${Auth Protocol}
			${Network Location}
			${IP Address}
			${User Agent String}
		}
		$XPath = "*[System[Provider[@Name='AD FS Auditing']]]" 
		$Events= Get-WinEvent -LogName 'Security' -FilterXPath $XPath -ComputerName $ADFS | Where-Object {$_.id -in @(1200,1203)}
		$ADFSRecord = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
		$ADFSRecordOut = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
		$UserID = ""
		ForEach ($e in $Events) {

			$exml = [xml]$e.Message.Substring($e.Message.IndexOf("XML: ")+5)
			$Record = [ADFSEventRecord]::new()
			$Record."Date-Time" = $e.TimeCreated
			$Record."ADFS Server" = $ADFS
			$UserID = $exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "ResourceAuditComponent"}).UserID 
			If ($UserID -match '@') {
				$Record."User ID" = ($UserID -split "@")[0]
			}ElseIf ($UserID -match '\\'){
				$Record."User ID" = $Record."User ID" = ($UserID -split '\\')[1]
			}
			$Record."Relying Party" = $exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "ResourceAuditComponent"}).RelyingParty
			$Record."Auth Protocol" = $exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).AuthProtocol
			$Record."Network Location" = $exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).NetworkLocation
			$Record."IP Address" = $exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).IpAddress
			$Record."User Agent String" = $exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).UserAgentString

			$ADFSRecord += $Record
		}

		$ADFSRUsers = ($ADFSRecord | Select-Object -Unique -Property "User ID" | Sort-Object)."User ID"

		Foreach ($ADFSRUser in $ADFSRUsers) {
			$ADFSRecordOut += $ADFSRecord.Where({$_."User ID" -eq $ADFSRUser}) | Sort-Object -Descending -Property "Date-Time" | Select-Object -First 1
		}

		$ADFSRecordOut.GetEnumerator()
	}
	Foreach ($ADFS in $ADFSServers) {
		$Jobs += Start-Job -Name "ADFS Events on $ADFS" -ScriptBlock $ADFSScript -ArgumentList $ADFS
	}
	#endregion ADFS Event logs

	#Get All AD users
	Write-Host "`tAD . . ." -ForegroundColor DarkGray
	$ADScript = {
		param(
			$ADServer,
			$SearchBase
		)
		If ((Get-Module | Where-Object {$_.Name -Match "ActiveDirectory"}).Count -eq 0 ) {
			#Write-Host ("Loading Active Directory Plugins") -ForeGroundColor "Green"
			Import-Module "ActiveDirectory"  -ErrorAction SilentlyContinue
		}
		Get-ADUser -server $ADServer -SearchBase $SearchBase -Filter * -Properties * 
	}
	$Jobs += Start-Job -Name "AD User" -ScriptBlock $ADScript -ArgumentList $ADServer, $SearchBase

	#region Azure AD Users
	#Get All Azure AD users
	# Write-Host "`tAzure AD . . ." -ForegroundColor DarkGray
	# $AzueADScript = {
	# 	param(
	# 		$User,
	# 		$Domain,
	# 		$AZClientID,
	# 		$AZCertThumbprint,
	# 		$AzureTenant
	# 	)
	# 	#region AzureAD
	# 	If (Get-Module -ListAvailable -Name "AzureAD") {
	# 		If (-Not (Get-Module "AzureAD" -ErrorAction SilentlyContinue)) {
	# 			If ($PSVersionTable.PSVersion.Major -gt 5) {
	# 			Import-Module "AzureAD" -DisableNameChecking -UseWindowsPowerShell -SkipEditionCheck
	# 			}Else {
	# 				Import-Module "AzureAD" -DisableNameChecking
	# 			}
				
	# 		} Else {
	# 			#write-host "AzureAD PowerShell Module Already Loaded"
	# 		} 
	# 	} Else {
	# 		Import-Module PackageManagement
	# 		Import-Module PowerShellGet
	# 		# Register-PSRepository -Name "PSGallery" –SourceLocation "https://www.powershellgallery.com/api/v2/" -InstallationPolicy Trusted
	# 		If (-Not (Get-PSRepository -Name "PSGallery")) {
	# 			Register-PSRepository -Default -InstallationPolicy Trusted 
	# 		}
	# 		If ((Get-PSRepository -Name "PSGallery").InstallationPolicy -eq "Untrusted") {
	# 			Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
	# 		}
	# 		Install-Module "AzureAD" -Force -Confirm:$false -Scope:CurrentUser -SkipPublisherCheck -AllowClobber
	# 		If ($PSVersionTable.PSVersion.Major -gt 5) {
	# 			Import-Module "AzureAD" -DisableNameChecking -UseWindowsPowerShell -SkipEditionCheck
	# 		}Else {
	# 			Import-Module "AzureAD" -DisableNameChecking
	# 		}
	# 		If (-Not (Get-Module "AzureAD" -ErrorAction SilentlyContinue)) {
	# 			write-error ("Please install AzureAD Powershell Modules Error")
	# 			exit
	# 		}
	# 	}
	# 	#endregion AzureAD
	# 	If (Get-ChildItem "Cert:\LocalMachine\My\" | Where-Object {$_.Thumbprint -eq $AZCertThumbprint}) {
	# 		(Connect-AzureAD -TenantId $AzureTenant -ApplicationId  $AZClientID -CertificateThumbprint $AZCertThumbprint) | Out-Null
	# 	}Else{
	# 		(Connect-AzureAD -AccountId ($User + "@" + $Domain )) | Out-Null
	# 	}
		
	# 	Get-AzureADUser -All:$true
	# }
	# $Jobs += Start-Job -Name "AzureAD User" -ScriptBlock $AzueADScript -ArgumentList $env:USERNAME , $env:USERDNSDOMAIN, $AZClientID, $AZCertThumbprint, $AzureTenant
	#endregion Azure AD Users
	#region Get Licensing info for Azure AD Users
	# Write-Host "`tAzureAD User Licensing . . ." -ForegroundColor DarkGray
	# $AzueADLicensingScript = {
	# 	param(
	# 		$User,
	# 		$Domain,
	# 		$AZClientID,
	# 		$AZCertThumbprint,
	# 		$AzureTenant
	# 	)
	# 	$OutputAZUL = [ordered]@{}
	# 	$AZUL = $null
	# 	#region AzureAD
	# 	If (Get-Module -ListAvailable -Name "AzureAD") {
	# 		If (-Not (Get-Module "AzureAD" -ErrorAction SilentlyContinue)) {
	# 			If ($PSVersionTable.PSVersion.Major -gt 5) {
	# 				Import-Module "AzureAD" -DisableNameChecking -UseWindowsPowerShell
	# 			}Else {
	# 				Import-Module "AzureAD" -DisableNameChecking
	# 			}
	# 		} Else {
	# 			#write-host "AzureAD PowerShell Module Already Loaded"
	# 		} 
	# 	} Else {
	# 		Import-Module PackageManagement
	# 		Import-Module PowerShellGet
	# 		# Register-PSRepository -Name "PSGallery" –SourceLocation "https://www.powershellgallery.com/api/v2/" -InstallationPolicy Trusted
	# 		If (-Not (Get-PSRepository -Name "PSGallery")) {
	# 			Register-PSRepository -Default -InstallationPolicy Trusted 
	# 		}
	# 		If ((Get-PSRepository -Name "PSGallery").InstallationPolicy -eq "Untrusted") {
	# 			Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
	# 		}
	# 		Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted 
	# 		Install-Module "AzureAD" -Force -Confirm:$false -Scope:CurrentUser -SkipPublisherCheck -AllowClobber
	# 		If ($PSVersionTable.PSVersion.Major -gt 5) {
	# 			Import-Module "AzureAD" -DisableNameChecking -UseWindowsPowerShell
	# 		}Else {
	# 			Import-Module "AzureAD" -DisableNameChecking
	# 		}
	# 		If (-Not (Get-Module "AzureAD" -ErrorAction SilentlyContinue)) {
	# 			write-error ("Please install AzureAD Powershell Modules Error")
	# 			exit
	# 		}
	# 	}
	# 	#endregion AzureAD
	# 	If (Get-ChildItem "Cert:\LocalMachine\My\" | Where-Object {$_.Thumbprint -eq $AZCertThumbprint}) {
	# 		(Connect-AzureAD -TenantId $AzureTenant -ApplicationId  $AZClientID -CertificateThumbprint $AZCertThumbprint) | Out-Null
	# 	}Else{
	# 		(Connect-AzureAD -AccountId ($User + "@" + $Domain )) | Out-Null
	# 	}
	# 	Foreach ($AZU in (Get-AzureADUser).Where({$null -ne $_.ObjectId})) {
	# 		$AZUL = ((Get-AzureADUserLicenseDetail -ObjectId $AZU.ObjectId).SkuPartNumber -join ",")
	# 		If (-Not [string]::IsNullOrWhiteSpace($AZUL)) {
	# 			$OutputAZUL.Add($AZU.UserPrincipalName,$AZUL)
	# 		}
	# 	}  
	# 	$OutputAZUL
	# }
	# $Jobs += Start-Job -Name "AzureAD User Licensing" -ScriptBlock $AzueADLicensingScript -ArgumentList $env:USERNAME , $env:USERDNSDOMAIN, $AZClientID, $AZCertThumbprint, $AzureTenant
	#endregion Get Licensing info for Azure AD Users
	#Get All Exchange  Mailboxes
	Write-Host "`tExchange Mailboxes . . ." -ForegroundColor DarkGray
	$EPMailboxesScript = {
		param(
			$ExchangeServer
		)
		If ((Get-PSSession | Where-Object { $_.ConfigurationName -Match "Microsoft.Exchange" -and $_.ComputerName -eq $ExchangeServer}).Count -eq 0 ) {
			$ERPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://$ExchangeServer/PowerShell/ -Authentication Negotiate -AllowRedirection 
			Import-PSSession $ERPSession -AllowClobber -DisableNameChecking -Prefix "EP"
		}
		Get-EPMailbox -ResultSize unlimited
	}
	$Jobs += Start-Job -Name "Local Exchange Mailbox" -ScriptBlock $EPMailboxesScript -ArgumentList $ExchangeServer
	Write-Host "`tExchange Mailboxes Statistics . . ." -ForegroundColor DarkGray
	$EPMailboxesStatsScript = {
		param(
			$ExchangeServer						
			)
			If ((Get-PSSession | Where-Object { $_.ConfigurationName -Match "Microsoft.Exchange" -and $_.ComputerName -eq $ExchangeServer}).Count -eq 0 ) {
			$ERPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://$ExchangeServer/PowerShell/ -Authentication Negotiate -AllowRedirection 
			Import-PSSession $ERPSession -AllowClobber -DisableNameChecking -Prefix "EP"
		}
		Get-EPMailbox -ResultSize unlimited | Get-EPMailboxStatistics
	}
	$Jobs += Start-Job -Name "Local Exchange Mailbox Stats" -ScriptBlock $EPMailboxesStatsScript -ArgumentList $ExchangeServer
	Write-Host "`tExchange Mailboxes Permissions . . ." -ForegroundColor DarkGray
	$EPMailboxesPermissionsScript = {
		param(
			$ExchangeServer
		)
		$EPMailboxesPerms = [ordered]@{}
		If ((Get-PSSession | Where-Object { $_.ConfigurationName -Match "Microsoft.Exchange" -and $_.ComputerName -eq $ExchangeServer}).Count -eq 0 ) {
			$ERPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://$ExchangeServer/PowerShell/ -Authentication Negotiate -AllowRedirection 
			Import-PSSession $ERPSession -AllowClobber -DisableNameChecking -Prefix "EP"
		}
		Foreach ($Mailbox in (Get-EPMailbox -ResultSize unlimited)) {
			$EPMailboxesPerms.($Mailbox.UserPrincipalName) = $Mailbox |  Get-EPMailboxPermission
		}
		$EPMailboxesPerms	
	}
	$Jobs += Start-Job -Name "Local Exchange Mailbox Permissions" -ScriptBlock $EPMailboxesPermissionsScript -ArgumentList $ExchangeServer
	Write-Host "`tExchange Mailboxes Archive . . ." -ForegroundColor DarkGray
	$EPMailboxesArchiveScript = {
		param(
			$ExchangeServer
		)
		If ((Get-PSSession | Where-Object { $_.ConfigurationName -Match "Microsoft.Exchange" -and $_.ComputerName -eq $ExchangeServer}).Count -eq 0 ) {
			$ERPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://$ExchangeServer/PowerShell/ -Authentication Negotiate -AllowRedirection 
			Import-PSSession $ERPSession -AllowClobber -DisableNameChecking -Prefix "EP"
		}
		$EPMailboxesPerms = [ordered]@{}
		Foreach ($Mailbox in (Get-EPMailbox -ResultSize unlimited -Archive)) {
			$EPMailboxesPerms.($Mailbox.UserPrincipalName) = $Mailbox |  Get-EPMailboxPermission
		}
		$EPMailboxesPerms
	}
	$Jobs += Start-Job -Name "Local Exchange Mailbox Archive" -ScriptBlock $EPMailboxesArchiveScript -ArgumentList $ExchangeServer
	Write-Host "`tExchange Mailboxes Archive Statistics . . ." -ForegroundColor DarkGray
	$EPMailboxesArchiveStatsScript = {
		param(
			$ExchangeServer
			)
			If ((Get-PSSession | Where-Object { $_.ConfigurationName -Match "Microsoft.Exchange" -and $_.ComputerName -eq $ExchangeServer}).Count -eq 0 ) {
			$ERPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://$ExchangeServer/PowerShell/ -Authentication Negotiate -AllowRedirection 
			Import-PSSession $ERPSession -AllowClobber -DisableNameChecking -Prefix "EP"
		}
		Get-EPMailbox -ResultSize unlimited -Archive | Get-EPMailboxStatistics -Archive
	}
	$Jobs += Start-Job -Name "Local Exchange Archive Mailbox Stats" -ScriptBlock $EPMailboxesArchiveStatsScript -ArgumentList $ExchangeServer
	Write-Host "`tExchange Mailboxes Archive Permissions . . ." -ForegroundColor DarkGray
	$EPMailboxesArchivePermissionsScript = {
		param(
			$ExchangeServer
		)
		$EPMailboxesArchivePerms = [ordered]@{}
		If ((Get-PSSession | Where-Object { $_.ConfigurationName -Match "Microsoft.Exchange" -and $_.ComputerName -eq $ExchangeServer}).Count -eq 0 ) {
			$ERPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://$ExchangeServer/PowerShell/ -Authentication Negotiate -AllowRedirection 
			Import-PSSession $ERPSession -AllowClobber -DisableNameChecking -Prefix "EP"
		}
		Foreach ($Mailbox in (Get-EPMailbox -ResultSize unlimited -Archive)) {
			$EPMailboxesArchivePerms.($Mailbox.UserPrincipalName) = $Mailbox |  Get-EPMailboxPermission 
		}
		$EPMailboxesArchivePerms
	}
	$Jobs += Start-Job -Name "Local Exchange Archive Mailbox Permissions" -ScriptBlock $EPMailboxesArchivePermissionsScript -ArgumentList $ExchangeServer
	Write-Host "`tExchange Remote Mailboxes . . ." -ForegroundColor DarkGray
	$EPRemoteMailboxesScript = {
		param(
			$ExchangeServer
		)
		If ((Get-PSSession | Where-Object { $_.ConfigurationName -Match "Microsoft.Exchange" -and $_.ComputerName -eq $ExchangeServer}).Count -eq 0 ) {
			$ERPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://$ExchangeServer/PowerShell/ -Authentication Negotiate -AllowRedirection 
			Import-PSSession $ERPSession -AllowClobber -DisableNameChecking -Prefix "EP"
		}
		Get-EPRemoteMailbox -ResultSize unlimited
	}
	$Jobs += Start-Job -Name "Local Exchange Remote Mailbox" -ScriptBlock $EPRemoteMailboxesScript -ArgumentList $ExchangeServer
	$EPRemoteMailboxesArchiveScript = {
		param(
			$ExchangeServer
		)
		If ((Get-PSSession | Where-Object { $_.ConfigurationName -Match "Microsoft.Exchange" -and $_.ComputerName -eq $ExchangeServer}).Count -eq 0 ) {
			$ERPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://$ExchangeServer/PowerShell/ -Authentication Negotiate -AllowRedirection 
			Import-PSSession $ERPSession -AllowClobber -DisableNameChecking -Prefix "EP"
		}
		Get-EPRemoteMailbox -ResultSize unlimited -Archive
	}
	$Jobs += Start-Job -Name "Local Exchange Remote Mailbox Archive" -ScriptBlock $EPRemoteMailboxesArchiveScript -ArgumentList $ExchangeServer
	Write-Host "`tExchange Mailbox Forwarding Rules . . ." -ForegroundColor DarkGray
	$EPMailboxesForwardingRulesScript = {
		param(
			$ExchangeServer
		)
		If ((Get-PSSession | Where-Object { $_.ConfigurationName -Match "Microsoft.Exchange" -and $_.ComputerName -eq $ExchangeServer}).Count -eq 0 ) {
			$ERPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://$ExchangeServer/PowerShell/ -Authentication Negotiate -AllowRedirection 
			Import-PSSession $ERPSession -AllowClobber -DisableNameChecking -Prefix "EP"
		}
		$EPMailboxesForwardingRules = [ordered]@{}
		Foreach ($Mailbox in (Get-EPMailbox -ResultSize unlimited)) {
			$Rules = Get-EPInboxRule -Mailbox $Mailbox.DistinguishedName -WarningAction SilentlyContinue | Where-Object {-Not [string]::IsNullOrWhiteSpace($_.ForwardTo)}
			$ORules = @()
			ForEach($Rule in $Rules) {
				If ($_.ForwardTo -match '" \[EX:'){
					$ORules += ($_.Name + " = Mailbox: " + (($_.ForwardTo  -split "\[")[0] -replace '"'))
				}Else{
					$ORules += ($_.Name + " = " + ($_.ForwardTo -split "\[SMTP\:")[-1] -replace "]")
				} 
			}
			$EPMailboxesForwardingRules.($Mailbox.UserPrincipalName) = $ORules -join ","
		}
		$EPMailboxesForwardingRules
	}
	$Jobs += Start-Job -Name "Local Exchange Mailbox Forwarding Rules" -ScriptBlock $EPMailboxesForwardingRulesScript -ArgumentList $ExchangeServer
	#Get Local Exchange CAS Mailbox Settings
	$EPEPCASMailboxScript = {
		param(
			$ExchangeServer
		)
		If ((Get-PSSession | Where-Object { $_.ConfigurationName -Match "Microsoft.Exchange" -and $_.ComputerName -eq $ExchangeServer}).Count -eq 0 ) {
			$ERPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://$ExchangeServer/PowerShell/ -Authentication Negotiate -AllowRedirection 
			Import-PSSession $ERPSession -AllowClobber -DisableNameChecking -Prefix "EP"
		}
		Get-EPCASMailbox
	}
	$Jobs += Start-Job -Name "Local Exchange CAS Mailbox Settings" -ScriptBlock $EPEPCASMailboxScript -ArgumentList $ExchangeServer
	#Get All Exchange Online Mailboxes
	Write-Host "`tExchange Online Mailboxes . . ." -ForegroundColor DarkGray
	$EXOMailboxesScript = {
		param(
			$CUPN,
			$AZClientID,
			$AZCertThumbprint,
			$AZOrg
		)
		If (Get-Module -ListAvailable -Name "ExchangeOnlineManagement") {
			If (-Not (Get-Module "ExchangeOnlineManagement" -ErrorAction SilentlyContinue)) {
				Import-Module "ExchangeOnlineManagement" -DisableNameChecking
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
			If (Get-ChildItem "Cert:\LocalMachine\My\" | Where-Object {$_.Thumbprint -eq $AZCertThumbprint}) {
				Connect-ExchangeOnline -CertificateThumbPrint $AZCertThumbprint -AppID $AZClientID -Organization $AZOrg -ShowProgress:$false -ShowBanner:$false
			}Else{
				Connect-ExchangeOnline -ShowBanner:$false -UserPrincipalName $CUPN -ShowProgress:$false
			}
		}
		Get-EXOMailbox -ResultSize unlimited -PropertySets All
	}
	$Jobs += Start-Job -Name "Online Exchange Mailbox" -ScriptBlock $EXOMailboxesScript -ArgumentList ([string]([ADSI]"LDAP://<SID=$([System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value)>").UserPrincipalName,$AZClientID,$AZCertThumbprint,$AZOrg)
	Write-Host "`tExchange Online Mailboxes Statistics . . ." -ForegroundColor DarkGray
	$EXOMailboxesStatsScript = {
		param(
			$CUPN,
			$AZClientID,
			$AZCertThumbprint,
			$AZOrg
		)
		If (Get-Module -ListAvailable -Name "ExchangeOnlineManagement") {
			If (-Not (Get-Module "ExchangeOnlineManagement" -ErrorAction SilentlyContinue)) {
				Import-Module "ExchangeOnlineManagement" -DisableNameChecking
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
			If (Get-ChildItem "Cert:\LocalMachine\My\" | Where-Object {$_.Thumbprint -eq $AZCertThumbprint}) {
				Connect-ExchangeOnline -CertificateThumbPrint $AZCertThumbprint -AppID $AZClientID -Organization $AZOrg -ShowProgress:$false -ShowBanner:$false
			}Else{
				Connect-ExchangeOnline -ShowBanner:$false -UserPrincipalName $CUPN -ShowProgress:$false
			}
		}
		Get-EXOMailbox -ResultSize unlimited | Get-EXOMailboxStatistics
	}
	$Jobs += Start-Job -Name "Online Exchange Mailbox Stats" -ScriptBlock $EXOMailboxesStatsScript -ArgumentList ([string]([ADSI]"LDAP://<SID=$([System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value)>").UserPrincipalName,$AZClientID,$AZCertThumbprint,$AZOrg)

	Write-Host "`tExchange Online Mailboxes Permissions . . ." -ForegroundColor DarkGray
	$EXOMailboxesPermissionsScript = {
		param(
			$CUPN,
			$AZClientID,
			$AZCertThumbprint,
			$AZOrg
		)
		$EXOMailboxesPerms = [ordered]@{}
		If (Get-Module -ListAvailable -Name "ExchangeOnlineManagement") {
			If (-Not (Get-Module "ExchangeOnlineManagement" -ErrorAction SilentlyContinue)) {
				Import-Module "ExchangeOnlineManagement" -DisableNameChecking
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
			If (Get-ChildItem "Cert:\LocalMachine\My\" | Where-Object {$_.Thumbprint -eq $AZCertThumbprint}) {
				Connect-ExchangeOnline -CertificateThumbPrint $AZCertThumbprint -AppID $AZClientID -Organization $AZOrg -ShowProgress:$false -ShowBanner:$false
			}Else{
				Connect-ExchangeOnline -ShowBanner:$false -UserPrincipalName $CUPN -ShowProgress:$false
			}
		}
		Foreach ($Mailbox in (Get-EXOMailbox -ResultSize unlimited -PropertySets All)) {
			$EXOMailboxesPerms.($Mailbox.UserPrincipalName) = $Mailbox |  Get-EXOMailboxPermission -ResultSize Unlimited
		}
		$EXOMailboxesPerms
	}
	$Jobs += Start-Job -Name "Online Exchange Mailbox Permission" -ScriptBlock $EXOMailboxesPermissionsScript -ArgumentList ([string]([ADSI]"LDAP://<SID=$([System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value)>").UserPrincipalName,$AZClientID,$AZCertThumbprint,$AZOrg)	

	Write-Host "`tExchange Online Mailboxes Archive . . ." -ForegroundColor DarkGray
	$EXOMailboxesArchiveScript = {
		param(
			$CUPN,
			$AZClientID,
			$AZCertThumbprint,
			$AZOrg
		)
		If (Get-Module -ListAvailable -Name "ExchangeOnlineManagement") {
			If (-Not (Get-Module "ExchangeOnlineManagement" -ErrorAction SilentlyContinue)) {
				Import-Module "ExchangeOnlineManagement" -DisableNameChecking
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
			If (Get-ChildItem "Cert:\LocalMachine\My\" | Where-Object {$_.Thumbprint -eq $AZCertThumbprint}) {
				Connect-ExchangeOnline -CertificateThumbPrint $AZCertThumbprint -AppID $AZClientID -Organization $AZOrg -ShowProgress:$false -ShowBanner:$false
			}Else{
				Connect-ExchangeOnline -ShowBanner:$false -UserPrincipalName $CUPN -ShowProgress:$false
			}
		}
		Get-EXOMailbox -ResultSize unlimited -Archive -PropertySets All
	}
	$Jobs += Start-Job -Name "Online Exchange Mailbox Archive" -ScriptBlock $EXOMailboxesArchiveScript -ArgumentList ([string]([ADSI]"LDAP://<SID=$([System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value)>").UserPrincipalName,$AZClientID,$AZCertThumbprint,$AZOrg)
	Write-Host "`tExchange Online Mailboxes Archive Statistics . . ." -ForegroundColor DarkGray
	$EXOMailboxesArchiveStatsScript = {
		param(
			$CUPN,
			$AZClientID,
			$AZCertThumbprint,
			$AZOrg
			)
			If (Get-Module -ListAvailable -Name "ExchangeOnlineManagement") {
			If (-Not (Get-Module "ExchangeOnlineManagement" -ErrorAction SilentlyContinue)) {
				Import-Module "ExchangeOnlineManagement" -DisableNameChecking
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
			If (Get-ChildItem "Cert:\LocalMachine\My\" | Where-Object {$_.Thumbprint -eq $AZCertThumbprint}) {
				Connect-ExchangeOnline -CertificateThumbPrint $AZCertThumbprint -AppID $AZClientID -Organization $AZOrg -ShowProgress:$false -ShowBanner:$false
			}Else{
				Connect-ExchangeOnline -ShowBanner:$false -UserPrincipalName $CUPN -ShowProgress:$false
			}
		}
		Get-EXOMailbox -ResultSize unlimited -Archive | Get-EXOMailboxStatistics -Archive
	}
	$Jobs += Start-Job -Name "Online Exchange Mailbox Archive Stats" -ScriptBlock $EXOMailboxesArchiveStatsScript -ArgumentList ([string]([ADSI]"LDAP://<SID=$([System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value)>").UserPrincipalName,$AZClientID,$AZCertThumbprint,$AZOrg)
	Write-Host "`tExchange Online Mailboxes Archive Permissions . . ." -ForegroundColor DarkGray
	$EXOMailboxesArchivePermissionsScript = {
		param(
			$CUPN,
			$AZClientID,
			$AZCertThumbprint,
			$AZOrg
			)
		$EXOMailboxesArchivePerms = [ordered]@{}
		If (Get-Module -ListAvailable -Name "ExchangeOnlineManagement") {
			If (-Not (Get-Module "ExchangeOnlineManagement" -ErrorAction SilentlyContinue)) {
				Import-Module "ExchangeOnlineManagement" -DisableNameChecking
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
			If (Get-ChildItem "Cert:\LocalMachine\My\" | Where-Object {$_.Thumbprint -eq $AZCertThumbprint}) {
				Connect-ExchangeOnline -CertificateThumbPrint $AZCertThumbprint -AppID $AZClientID -Organization $AZOrg -ShowProgress:$false -ShowBanner:$false
			}Else{
				Connect-ExchangeOnline -ShowBanner:$false -UserPrincipalName $CUPN -ShowProgress:$false
			}
		}

		Foreach ($Mailbox in (Get-EXOMailbox -ResultSize unlimited -Archive -PropertySets All)) {
			$EXOMailboxesArchivePerms.($Mailbox.UserPrincipalName) = $Mailbox |  Get-EXOMailboxPermission -ResultSize Unlimited
		}
		$EXOMailboxesArchivePerms
	}
	$Jobs += Start-Job -Name "Online Exchange Mailbox Archive Permission" -ScriptBlock $EXOMailboxesArchivePermissionsScript -ArgumentList ([string]([ADSI]"LDAP://<SID=$([System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value)>").UserPrincipalName,$AZClientID,$AZCertThumbprint,$AZOrg)
	Write-Host "`tExchange Online Mailboxes Forward Rules . . ." -ForegroundColor DarkGray
	$EXOMailboxesForwardingRulesScript = {
		param(
			$CUPN,
			$AZClientID,
			$AZCertThumbprint,
			$AZOrg
			)
		$EXOMailboxesForwardingRules = [ordered]@{}
		If (Get-Module -ListAvailable -Name "ExchangeOnlineManagement") {
			If (-Not (Get-Module "ExchangeOnlineManagement" -ErrorAction SilentlyContinue)) {
				Import-Module "ExchangeOnlineManagement" -DisableNameChecking
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
			If (Get-ChildItem "Cert:\LocalMachine\My\" | Where-Object {$_.Thumbprint -eq $AZCertThumbprint}) {
				Connect-ExchangeOnline -CertificateThumbPrint $AZCertThumbprint -AppID $AZClientID -Organization $AZOrg -ShowProgress:$false -ShowBanner:$false
			}Else{
				Connect-ExchangeOnline -ShowBanner:$false -UserPrincipalName $CUPN -ShowProgress:$false
			}
		}

		Foreach ($Mailbox in (Get-EXOMailbox -ResultSize unlimited -PropertySets All)) {
			$Rules = Get-InboxRule -Mailbox $Mailbox.DistinguishedName -WarningAction SilentlyContinue | Where-Object {-Not [string]::IsNullOrWhiteSpace($_.ForwardTo)}
			$ORules = @()
			ForEach($Rule in $Rules) {
				If ($_.ForwardTo -match '" \[EX:'){
					$ORules += ($_.Name + " = Mailbox: " + (($_.ForwardTo  -split "\[")[0] -replace '"'))
				}Else{
					$ORules += ($_.Name + " = " + ($_.ForwardTo -split "\[SMTP\:")[-1] -replace "]")
				} 
			}	
			$EXOMailboxesForwardingRules.($Mailbox.UserPrincipalName) = $ORules -join ","
		}
		$EXOMailboxesForwardingRules
	}
	$Jobs += Start-Job -Name "Online Exchange Mailbox Forwarding Rules" -ScriptBlock $EXOMailboxesForwardingRulesScript -ArgumentList ([string]([ADSI]"LDAP://<SID=$([System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value)>").UserPrincipalName,$AZClientID,$AZCertThumbprint,$AZOrg)		
	#Get Online Exchange CAS Mailbox Settings
	$EXOCASMailboxScript = {
		param(
			$CUPN,
			$AZClientID,
			$AZCertThumbprint,
			$AZOrg
		)
		If (Get-Module -ListAvailable -Name "ExchangeOnlineManagement") {
			If (-Not (Get-Module "ExchangeOnlineManagement" -ErrorAction SilentlyContinue)) {
				Import-Module "ExchangeOnlineManagement" -DisableNameChecking
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
			If (Get-ChildItem "Cert:\LocalMachine\My\" | Where-Object {$_.Thumbprint -eq $AZCertThumbprint}) {
				Connect-ExchangeOnline -CertificateThumbPrint $AZCertThumbprint -AppID $AZClientID -Organization $AZOrg -ShowProgress:$false -ShowBanner:$false
			}Else{
				Connect-ExchangeOnline -ShowBanner:$false -UserPrincipalName $CUPN -ShowProgress:$false
			}
		}
		Get-EXOCasMailbox -PropertySets All
	}
	$Jobs += Start-Job -Name "Online Exchange CAS Mailbox Settings" -ScriptBlock $EXOCASMailboxScript -ArgumentList ([string]([ADSI]"LDAP://<SID=$([System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value)>").UserPrincipalName,$AZClientID,$AZCertThumbprint,$AZOrg)


	#region MgUsers Users
		Write-Host "`tMgUsers Entra ID Users . . ." -ForegroundColor DarkGray
		$MgUsers = {
			param(
				$AzureTenant,
				$AZClientID,
				$AZCertThumbprint
			)
			Class MSEnUsers {
				[string]${Entra ID}
				[string]${Display Name}
				[string]${User Principal Name}
				[string]${Sam Account Name}
				[datetime]${Entra Last Sync}
				[datetime]${Last Sign-In}
				[datetime]${Last Non-Interactive Sign-In}
				[string]${Licenses Part Number}
				[string]${Account Enabled}
				[string]${Licenses}
				[string]${User Type}
				[string]${ImmutableId}
				[string]${SecurityIdentifier}
				[string]${MFA Status}
				[string]${MFA Preferred method}
				[string]${Phone Authentication}
				[string]${Authenticator App}
				[string]${Passwordless}
				[string]${Hello for Business}
				[string]${FIDO2 Security Key}
				[string]${Temporary Access Pass}
				[string]${Authenticator device}
			}
			$ALRecords = New-Object -TypeName "System.Collections.ArrayList"
			#region Microsoft.Graph
			If($PSVersionTable.PSVersion.Major -eq 5){
				# Increase the Function Count
				$Global:MaximumFunctionCount = 8192
				
				# Increase the Variable Count
				$Global:MaximumVariableCount = 8192
			}
			If (Get-Module -ListAvailable -Name "Microsoft.Graph") {
				If (-Not (Get-Module "Microsoft.Graph" -ErrorAction SilentlyContinue)) {
					#Import-Module "Microsoft.Graph" 
				} Else {
					#write-host "Microsoft.Graph PowerShell Module Already Loaded"
				} 
			} Else {
				Import-Module PackageManagement
				Import-Module PowerShellGet
				# Register-PSRepository -Name "PSGallery" –SourceLocation "https://www.powershellgallery.com/api/v2/" -InstallationPolicy Trusted
				If (-Not (Get-PSRepository -Name "PSGallery")) {
					Register-PSRepository -Default -InstallationPolicy Trusted 
				}
				If ((Get-PSRepository -Name "PSGallery").InstallationPolicy -eq "Untrusted") {
					Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
				}
				Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted 
				Install-Module "Microsoft.Graph" -Force -Confirm:$false -Scope:CurrentUser -SkipPublisherCheck -AllowClobber
				#Import-Module "Microsoft.Graph" 
				If (-Not (Get-Module "Microsoft.Graph" -ErrorAction SilentlyContinue)) {
					write-error ("Please install MSGraph Powershell Modules Error")
					exit
				}
			}
	
			#region Get Access Token
			If($null -ne $AZCertThumbprint -and $AZCertThumbprint.Length -eq 40) {
				If (Get-ChildItem "Cert:\LocalMachine\My\" | Where-Object {$_.Thumbprint -eq $AZCertThumbprint}) {
					$ClientCertificate = Get-Item "Cert:\LocalMachine\My\$($AZCertThumbprint)"
					# $myAccessToken = Get-MsalToken -ClientId $AZClientID -TenantId $AzureTenant -ClientCertificate $ClientCertificate
					# $AccessToken  = $myAccessToken.AccessToken
				}
			}
			#endregion Get Access Token
	
			#Connect to MgGraph
			If ($null -ne $AzureTenant) {
				If($null -ne $ClientCertificate ) {
					Connect-MgGraph -ClientId $AZClientID -TenantId $AzureTenant -Certificate $ClientCertificate -NoWelcome 
				}ElseIf($null -ne $AZCertThumbprint -and $AZCertThumbprint.Length -eq 40) {
					Connect-MgGraph -ClientId $AZClientID -TenantId $AzureTenant -CertificateThumbprint $AZCertThumbprint -NoWelcome 
				}Else {
					Connect-MgGraph -NoWelcome -TenantId $AzureTenant -Scopes  "User.ReadBasic.All", "UserAuthenticationMethod.Read.All", "IdentityUserFlow.Read.All", "User.EnableDisableAccount.All", "User.EnableDisableAccount.All", "IdentityRiskyUser.Read.All"
				}
				$Users = Get-MgUser -all -Property Id, DisplayName, UserPrincipalName, AccountEnabled, SignInActivity, assignedLicenses, assignedPlans, licenseAssignmentStates, onPremisesSecurityIdentifier, onPremisesImmutableId, onPremisesLastSyncDateTime, onPremisesSamAccountName, passwordPolicies, passwordProfile, mail
		
				ForEach ($User in $Users ) {
					$Record = [MSEnUsers]::new()
					$Record."Entra ID" = $User.ID
					$Record."Display Name" = $User.displayName
					$Record."User Principal Name" = $User.userPrincipalName
					$Record."Sam Account Name" = $User.onPremisesSamAccountName
					$Record."Entra Last Sync" =  if(!([string]::IsNullOrWhiteSpace([DateTime]$User.OnPremisesLastSyncDateTime))) {[DateTime]$User.OnPremisesLastSyncDateTime} Else {$null}
					$Record."Last Sign-In" = if(!([string]::IsNullOrWhiteSpace([DateTime]$User.signInActivity.lastSignInDateTime))) { [DateTime]$User.signInActivity.lastSignInDateTime } Else {$null}
					$Record."ImmutableId" = $User.ImmutableId
					$Record."SecurityIdentifier" = $User.onPremisesSecurityIdentifier
					$Record."Last Non-Interactive Sign-In" = if(!([string]::IsNullOrWhiteSpace([DateTime]$User.signInActivity.lastNonInteractiveSignInDateTime))) { [DateTime]$User.signInActivity.lastNonInteractiveSignInDateTime } Else { $null }
					$Record."Account Enabled" = $User.AccountEnabled
		
					#Get License Info
					$MGULD = Get-MgUserLicenseDetail -UserId $User.userPrincipalName -All
					$Record."Licenses Part Number" = $MGULD.SkuPartNumber -join ","
					$Record."Licenses" = ($MGULD.ServicePlans.Where({$_.ProvisioningStatus -eq "Success"}).ServicePlanName) -join ","
							
					$ALRecords.Add($Record) | Out-Null
					
				}
				If ($ALRecords.Count -gt 0) {
					$ALRecords.GetEnumerator()
				}Else {
					throw "No GraphAPI Records!!"
				}
				Remove-Variable Users
				Remove-Variable Record
				Remove-Variable ALRecords
		}Else{
			throw "Missing Azure Tenant ID."
		}
	
		}
		$Jobs += Start-Job -Name "MgUsers Entra ID" -ScriptBlock $MgUsers -ArgumentList $AzureTenant, $AZClientID, $AZCertThumbprint
	#endregion MgUsers Users
	
	#region MgUsers Users MFA
		Write-Host "`tMgUsers MFA Entra ID Users . . ." -ForegroundColor DarkGray
		$MgUsersMFA = {
			param(
				$AzureTenant,
				$AZClientID,
				$AZCertThumbprint
			)
			Class MSEnUsersMFA {
				${Entra ID}
				${Display Name}
				${User Principal Name}
				${MFA Status}
				${MFA Preferred method}
				${Phone Authentication}
				${Authenticator App}
				${Passwordless}
				${Hello for Business}
				${FIDO2 Security Key}
				${Temporary Access Pass}
				${Authenticator device}
			}
			$ALRecords = New-Object -TypeName "System.Collections.ArrayList"
			#region Microsoft.Graph
			If($PSVersionTable.PSVersion.Major -eq 5){
				# Increase the Function Count
				$Global:MaximumFunctionCount = 8192
				
				# Increase the Variable Count
				$Global:MaximumVariableCount = 8192
			}
			If (Get-Module -ListAvailable -Name "Microsoft.Graph") {
				If (-Not (Get-Module "Microsoft.Graph" -ErrorAction SilentlyContinue)) {
					#Import-Module "Microsoft.Graph" 
				} Else {
					#write-host "Microsoft.Graph PowerShell Module Already Loaded"
				} 
			} Else {
				Import-Module PackageManagement
				Import-Module PowerShellGet
				# Register-PSRepository -Name "PSGallery" –SourceLocation "https://www.powershellgallery.com/api/v2/" -InstallationPolicy Trusted
				If (-Not (Get-PSRepository -Name "PSGallery")) {
					Register-PSRepository -Default -InstallationPolicy Trusted 
				}
				If ((Get-PSRepository -Name "PSGallery").InstallationPolicy -eq "Untrusted") {
					Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
				}
				Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted 
				Install-Module "Microsoft.Graph" -Force -Confirm:$false -Scope:CurrentUser -SkipPublisherCheck -AllowClobber
				#Import-Module "Microsoft.Graph" 
				If (-Not (Get-Module "Microsoft.Graph" -ErrorAction SilentlyContinue)) {
					write-error ("Please install MSGraph Powershell Modules Error")
					exit
				}
			}
	
			#region Get Access Token
			If($null -ne $AZCertThumbprint -and $AZCertThumbprint.Length -eq 40) {
				If (Get-ChildItem "Cert:\LocalMachine\My\" | Where-Object {$_.Thumbprint -eq $AZCertThumbprint}) {
					$ClientCertificate = Get-Item "Cert:\LocalMachine\My\$($AZCertThumbprint)"
					# $myAccessToken = Get-MsalToken -ClientId $AZClientID -TenantId $AzureTenant -ClientCertificate $ClientCertificate
					# $AccessToken  = $myAccessToken.AccessToken
				}
			}
			#endregion Get Access Token
	
			#Connect to MgGraph
			If ($null -ne $AzureTenant) {
				If($null -ne $ClientCertificate ) {
					Connect-MgGraph -ClientId $AZClientID -TenantId $AzureTenant -Certificate $ClientCertificate -NoWelcome 
				}ElseIf($null -ne $AZCertThumbprint -and $AZCertThumbprint.Length -eq 40) {
					Connect-MgGraph -ClientId $AZClientID -TenantId $AzureTenant -CertificateThumbprint $AZCertThumbprint -NoWelcome 
				}Else {
					Connect-MgGraph -NoWelcome -TenantId $AzureTenant -Scopes  "User.ReadBasic.All", "UserAuthenticationMethod.Read.All", "IdentityUserFlow.Read.All", "User.EnableDisableAccount.All", "User.EnableDisableAccount.All", "IdentityRiskyUser.Read.All"
				}
				$Users = Get-MgUser -all -Property Id, DisplayName, UserPrincipalName
		
				ForEach ($User in $Users ) {
					$Record = [MSEnUsersMFA]::new()
					$Record."Entra ID" = $User.ID
					$Record."Display Name" = $User.displayName
					$Record."User Principal Name" = $User.userPrincipalName
					$Record."MFA Status" = "Disabled"

					#Get MFA Info
					$MSMFA = (Get-MgUserAuthenticationMicrosoftAuthenticatorMethod -UserId $User.userPrincipalName -all | Select-Object -First 1).DisplayName
					If ($MSMFA) {
						$Record."MFA Status" = "Enable"
						$Record."Authenticator App" = "Microsoft Authenticator"
						$Record."Authenticator device" = $MSMFA
					}
					#Get Windows Hello For Business
					$MSAWHB = Get-MgUserAuthenticationWindowsHelloForBusinessMethod -UserId $User.userPrincipalName -All | Sort-Object -Property CreatedDateTime -Descending | select -First 1
					If ($MSAWHB) {
						$Record."MFA Status" = "Enable"
						If ([string]::IsNullOrWhiteSpace($MSAWHB.DisplayName)){
							$Record."Hello for Business" = "Enable"
						}Else{
							$Record."Hello for Business" = ("Enabled On: " + $MSAWHB.DisplayName )
						}
					}
					
					If (Get-MgUserAuthenticationTemporaryAccessPassMethod -UserId $User.userPrincipalName) {
						$Record."Temporary Access Pass" = "Enable"
					}
					#Fido2 Key
					$MGUAFido2 = Get-MgUserAuthenticationFido2Method -UserId $User.userPrincipalName -All
					If ($MGUAFido2) {
						If ([string]::IsNullOrWhiteSpace(($MGUAFido2| Sort-Object -Property CreatedDateTime | Select-Object -First 1).DisplayName)) {
							$Record."FIDO2 Security Key" = "Enable"
							$Record."MFA Status" = "Enable"
						}Else{
							$Record."FIDO2 Security Key" = ($MGUAFido2| Sort-Object -Property CreatedDateTime | Select-Object -First 1).DisplayName
						}
					}
					#Phone Auth
					$MGUPA = Get-MgUserAuthenticationPhoneMethod -UserId $User.userPrincipalName -All
					If ($MGUPA.PhoneNumber) {
						$Record."Phone Authentication" = $MGUPA.PhoneNumber
						$Record."MFA Status" = "Enable"
					}
					#${MFA Preferred method}
					# ${Passwordless}
		
					$ALRecords.Add($Record) | Out-Null
					
				}
				If ($ALRecords.Count -gt 0) {
					$ALRecords.GetEnumerator()
				}Else {
					throw "No GraphAPI Records!!"
				}
				Remove-Variable Users
				Remove-Variable Record
				Remove-Variable ALRecords

		}Else{
			throw "Missing Azure Tenant ID."
		}
	
		}
		$Jobs += Start-Job -Name "MgUsers MFA Entra ID" -ScriptBlock $MgUsers -ArgumentList $AzureTenant, $AZClientID, $AZCertThumbprint
	#endregion MgUsers Users MFA

	Write-host ("`tStart Monitoring Jobs . . .") -ForegroundColor DarkYellow
	#Main Loop to monitor jobs
	do {
		Foreach ($cjob in (Get-Job -State Completed)) {
			switch -Wildcard ($cjob.Name) {
				"AD User" { 
					$ADUsers = Receive-Job -Id ($cjob.Id) 
					If ($null -ne $ADUsers) {
						Remove-Job -Id ($cjob.Id)
						Write-Host "`t`tImporting AD User Results . . ." -ForegroundColor DarkGray
					}
				}
				"AzureAD User" { 
					$AzureUsers = Receive-Job -Id ($cjob.Id)
					If ($null -ne $AzureUsers) {
						Remove-Job -Id ($cjob.Id)
						Write-Host "`t`tImporting AzureAD User Results . . ." -ForegroundColor DarkGray
					} 
				}
				"Local Exchange Mailbox" { 
					$EPMailboxes = Receive-Job -Id ($cjob.Id)
					If ($null -ne $EPMailboxes) {
						Remove-Job -Id ($cjob.Id)
						Write-Host "`t`tImporting Local Exchange Mailbox Results . . ." -ForegroundColor DarkGray
					} 
				}
				"Local Exchange Mailbox Archive" { 
					$EPMailboxesArchive = Receive-Job -Id ($cjob.id)
					If ($null -ne $EPMailboxesArchive) {
						Remove-Job -Id ($cjob.id)
						Write-Host "`t`tImporting Local Exchange Mailbox Archive Results . . ." -ForegroundColor DarkGray
					}
				}
				"Local Exchange Remote Mailbox" { 
					$EPRemoteMailboxes = Receive-Job -Id ($cjob.id)
					If ($null -ne $EPRemoteMailboxes) {
						Remove-Job -Id ($cjob.id)
						Write-Host "`t`tImporting Local Exchange Remote Mailbox Results . . ." -ForegroundColor DarkGray
					}
				}
				"Local Exchange Remote Mailbox Archive" { 
					$EPRemoteMailboxesArchive = Receive-Job -Id ($cjob.id)
					If ($null -ne $EPRemoteMailboxesArchive) {
						Remove-Job -Id ($cjob.id)
						Write-Host "`t`tImporting Local Exchange Remote Mailbox Archive Results . . ." -ForegroundColor DarkGray
					} 
				}
				"Local Exchange Mailbox Stats" { 
					$EPMailboxesStats = Receive-Job -Id ($cjob.id)
					If ($null -ne $EPMailboxesStats) {
						Remove-Job -Id ($cjob.id)
						Write-Host "`t`tImporting Local Exchange Mailbox Stats Results . . ." -ForegroundColor DarkGray
					} 
				}
				"Local Exchange Mailbox Permissions" { 
					$EPMailboxesPerms = Receive-Job -Id ($cjob.id)
					If ($null -ne $EPMailboxesPerms) {
						Remove-Job -Id ($cjob.id)
						Write-Host "`t`tImporting Local Exchange Mailbox Permissions Results . . ." -ForegroundColor DarkGray
					} 	
				}
				"Local Exchange Mailbox Forwarding Rules" {
					$EPMailboxesForwardingRules = Receive-Job -Id ($cjob.id)
					If ($null -ne $EPMailboxesForwardingRules){
						Remove-Job -Id ($cjob.id)
						Write-Host "`t`tImporting Local Exchange Mailbox Forwarding Rules Results . . ." -ForegroundColor DarkGray
					}
				}
				"Local Exchange Archive Mailbox Stats" { 
					$EPMailboxesArchiveStats = Receive-Job -Id ($cjob.id)
					If ($null -ne $EPMailboxesArchiveStats) {
						Remove-Job -Id ($cjob.id)
						Write-Host "`t`tImporting Local Exchange Archive Mailbox Stats Results . . ." -ForegroundColor DarkGray
					} 
				}
				"Local Exchange Archive Mailbox Permissions" { 
					$EPMailboxesArchivePerms = Receive-Job -Id ($cjob.id)
					If ($null -ne $EPMailboxesArchivePerms) {
						Remove-Job -Id ($cjob.id)
						Write-Host "`t`tImporting Local Exchange Archive Mailbox Permissions Results . . ." -ForegroundColor DarkGray
					} 
				}
				"Local Exchange CAS Mailbox Settings" { 
					$EPCASM = Receive-Job -Id ($cjob.id)
					If ($null -ne $EPCASM) {
						Remove-Job -Id ($cjob.id)
						Write-Host "`t`tImporting Local Exchange CAS Mailbox Settings Results . . ." -ForegroundColor DarkGray
					}
				}
				"Online Exchange Mailbox" { 
					$EXOMailboxes = Receive-Job -Id ($cjob.id)
					If ($null -ne $EXOMailboxes) {
						Remove-Job -Id ($cjob.id)
						Write-Host "`t`tImporting Online Exchange Mailbox Results . . ." -ForegroundColor DarkGray
					} 
				}
				"Online Exchange Mailbox Archive" { 
					$EXOMailboxesArchive = Receive-Job -Id ($cjob.id)
					If ($null -ne $EXOMailboxesArchive) {
						Remove-Job -Id ($cjob.id)
						Write-Host "`t`tImporting Online Exchange Mailbox Archive Results . . ." -ForegroundColor DarkGray
					} 
				}
				"Online Exchange Mailbox Stats" { 
					$EXOMailboxesStats = Receive-Job -Id ($cjob.id)
					If ($null -ne $EXOMailboxesStats) {
						Remove-Job -Id ($cjob.id)
						Write-Host "`t`tImporting Online Exchange Mailbox Stats Results . . ." -ForegroundColor DarkGray
					} 
				}
				"Online Exchange Mailbox Permission" { 
					$EXOMailboxesPerms = Receive-Job -Id ($cjob.id)
					If ($null -ne $EXOMailboxesPerms) {
						Remove-Job -Id ($cjob.id)
						Write-Host "`t`tImporting Online Exchange Mailbox Permission Results . . ." -ForegroundColor DarkGray
					} 
				}
				"Online Exchange Mailbox Archive Stats" { 
					$EXOMailboxesArchiveStats = Receive-Job -Id ($cjob.id)
					If ($null -ne $EXOMailboxesArchiveStats) {
						Remove-Job -Id ($cjob.id)
						Write-Host "`t`tImporting Online Exchange Mailbox Archive Stats Results . . ." -ForegroundColor DarkGray
					} 
				}
				"Online Exchange Mailbox Archive Permission" { 
					$EXOMailboxesArchivePerms = Receive-Job -Id ($cjob.id)
					If ($null -ne $EXOMailboxesArchivePerms) {
						Remove-Job -Id ($cjob.id)
						Write-Host "`t`tImporting Online Exchange Mailbox Archive Permission Results . . ." -ForegroundColor DarkGray
					} 
				}
				"Online Exchange Mailbox Forwarding Rules" {
					$EXOMailboxesForwardingRules = Receive-Job -Id ($cjob.id)
					If ($null -ne $EXOMailboxesForwardingRules) {
						Remove-Job -Id ($cjob.id)
						Write-Host "`t`tImporting Online Exchange Mailbox Forwarding Rules Results . . ." -ForegroundColor DarkGray
					} 
				}
				"Online Exchange CAS Mailbox Settings" { 
					$EPRemoteCASM = Receive-Job -Id ($cjob.id)
					If ($null -ne $EPRemoteCASM) {
						Remove-Job -Id ($cjob.id)
						Write-Host "`t`tImporting Online Exchange CAS Mailbox Settings Results . . ." -ForegroundColor DarkGray
					} 
				}
				"AzureAD User Licensing" {
					$OutputAZUL = Receive-Job -Id ($cjob.id)
					If ($null -ne $OutputAZUL) {
						Remove-Job -Id ($cjob.id)
						Write-Host "`t`tImporting AzureAD User Licensing Results . . ." -ForegroundColor DarkGray	
					} 		
				}
				"ADFS Events on *" {
					$ADFSLogsTemp = Receive-Job -Id ($cjob.id)
					If ($null -ne $ADFSLogsTemp) {
						$ADFSLogs += $ADFSLogsTemp
						Remove-Job -Id ($cjob.id)
						Write-Host "`t`t$($cjob.Name) Results . . ." -ForegroundColor DarkGray	
					} 
				}
				"GraphAPI Entra ID User" {
					$OutputGraphAPIU = Receive-Job -Id ($cjob.id)
					If ($null -ne $OutputGraphAPIU) {
						Remove-Job -Id ($cjob.id)
						Write-Host "`t`tImporting GraphAPI Entra ID Results . . ." -ForegroundColor DarkGray	
					} 					
				}
				"MgUsers Entra ID" {
					$OutputMgUser = Receive-Job -Id ($cjob.id)
					If ($null -ne $OutputMgUser) {
						Remove-Job -Id ($cjob.id)
						Write-Host "`t`tMgUsers Entra ID Results . . ." -ForegroundColor DarkGray	
					} 					
				}
				"MgUsers MFA Entra ID" {
					$OutputMgUserMFA = Receive-Job -Id ($cjob.id)
					If ($null -ne $OutputMgUserMFA) {
						Remove-Job -Id ($cjob.id)
						Write-Host "`t`tMgUsers MFA Entra ID Results . . ." -ForegroundColor DarkGray	
					} 					
				}
				Default {}
			}
		}
		If (Get-Job -State Failed) {
			Write-Error "Following jobs having failed:"
			Get-Job -State Failed
			throw "Stopping script"
		}
		$RuningJobCount = (get-job).count
		$TotalJobCount = $Jobs.count
		If ($RuningJobCount -gt $TotalJobCount) {
			Write-Error "Jobs from last script still running. Please close window and relaunch"
			throw "Stopping script"
		}
		$intJobs = 0
		$intJobsComplete = 0
		If (($TotalJobCount - $RuningJobCount) -gt 0) {
			$intJobs = ($TotalJobCount - $RuningJobCount)
		}
		If ((($intJobs / $TotalJobCount)*100) -gt 0){
			$intJobsComplete = ( '{0:N0}' -f (($intJobs / $TotalJobCount)*100))
		}
		Write-Progress -Activity ("Monitoring Jobs " + ($MyInvocation.MyCommand.Name -replace ".ps1","") ) -Status ("Caching Objects " + "[" + $intJobs + "/" + $TotalJobCount + "]") -percentComplete $intJobsComplete -Id 1
		
		Start-sleep -Seconds 5
	}while((Get-Job).count -gt 0 -and (get-job -State Failed).count -le 0)


	Write-Progress -Id 1 -Completed -Activity ("Monitoring Jobs " + ($MyInvocation.MyCommand.Name -replace ".ps1","") )
	$swc.Stop()
	$HMADITimeStamp = get-date
	Write-Host ("Caching Objects time: " + (FormatElapsedTime($swc.Elapsed)) + " to run. " + '{0:N0}' -f ($ADUsers.Count / $swc.Elapsed.TotalMinutes) + " Users's per Minute.")
}Else{
	Write-host ("Using Cached Objects from: " + $HMADITimeStamp + " that is " + (New-TimeSpan -Start $HMADITimeStamp -End (Get-Date)).Hours + " hours old.")
}
#endregion caching objects
[gc]::collect()
$swp = [Diagnostics.Stopwatch]::StartNew()
Write-host "Processing Users Please Wait ..."
Write-Progress -Activity ("Creating " + ($MyInvocation.MyCommand.Name -replace ".ps1","") + " output") -Status ("Processing " + "[" + $output.Count + "/" + $ADUsers.count + "]") -percentComplete 0 -Id 0 
#Main loop
Foreach ($ADUser in $ADUsers) { 
	# Write-host ("`t" + $ADUser.name)
	#Update progress Source: https://stackoverflow.com/questions/67981500/making-an-powershell-progress-bar-more-efficient 
	$Script:WindowWidthChanged = $Script:WindowWidth -ne $Host.UI.RawUI.WindowSize.Width
	if ($Script:WindowWidthChanged) { $Script:WindowWidth = $Host.UI.RawUI.WindowSize.Width }
	$ProgressCompleted = [math]::floor($Script:output.Count * $Script:WindowWidth / $ADUsers.count)
	if ($Script:WindowWidthChanged -or $ProgressCompleted -ne $LastProgressCompleted) {
		Write-Progress -Activity ("Creating " + ($MyInvocation.MyCommand.Name -replace ".ps1","") + " output") -Status ("Processing " + $ADUser.name + "[" + $Script:output.Count + "/" + $Script:ADUsers.count + "]") -percentComplete (($Script:output.Count / $Script:ADUsers.count)  * 100) -Id 0
	}
	$LastProgressCompleted = $ProgressCompleted
	#Clear Loop var
	$TotalItemSize = 0
	$TotalDeletedItemSize = 0
	$CUPN = $ADUser.UserPrincipalName
	$Record = [ADExchangeOutput]::new()
	$Record."Logon Name" = $ADUser.sAMAccountName
	$Record."Display Name" = $ADUser.DisplayName
	$Record."Last Name" = $ADUser.Surname
	$Record."Middle Name" = $ADUser.middleName
	$Record."First Name" = $ADUser.GivenName
	$Record."User Principal Name" = $ADUser.userPrincipalName
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
	Try{
		$MgUser = $OutputMgUser.Where({$_."User Principal Name" -eq $CUPN}) | Select-Object -First 1
	}Catch{
		$MgUser =$null
	}
	Try{
		$MgUserMFA = $OutputMgUserMFA.Where({$_."User Principal Name" -eq $CUPN}) | Select-Object -First 1
	}Catch{
		$MgUserMFA =$null
	}
	If(!([string]::IsNullOrWhiteSpace($MgUser."MFA Status"))){
		$Record."Azure MFA Status" = [string]$MgUser."MFA Status"
		If ($MgUser."MFA Status" -ne "disabled") {
			$Record."Azure MFA Preferred method" = [string]$MgUser."MFA Preferred method"
			$Record."Azure Phone Authentication" = [string]$MgUser."Phone Authentication"
			$Record."Azure Authenticator App" = [string]$MgUser."Authenticator App"
			$Record."Azure Passwordless" = [string]$MgUser."Passwordless"
			$Record."Azure Hello for Business" = [string]$MgUser."Hello for Business"
			$Record."Azure FIDO2 Security Key" = [string]$MgUser."FIDO2 Security Key"
			$Record."Azure Temporary Access Pass" = [string]$MgUser."Temporary Access Pass"
			$Record."Azure Authenticator device" = [string]$MgUser."Authenticator device"
		}
	}
	If(!([string]::IsNullOrWhiteSpace($MgUser."Entra Last Sync"))){
		$Record."Azure Last Sync Time" = [DateTime]$MgUser."Entra Last Sync"
	}
	If(!([string]::IsNullOrWhiteSpace($MgUser."Licenses Part Number"))){
		$Record."Azure Licenses" = [string]$MgUser."Licenses Part Number"
	}

	If(!([string]::IsNullOrWhiteSpace($MgUser."Licenses"))){
		$Record."Azure Licenses Details" = [string]$MgUser."Licenses"
	}

	If(!([string]::IsNullOrWhiteSpace($MgUser."Last Sign-In"))){
		$Record."Azure Last Sign-On" = [DateTime]$MgUser."Last Sign-In"
		$Record."Azure Last Sign-On Days" = ($((Get-Date) - ([DateTime]$MgUser."Last Sign-In")).Days)
	}
	If(!([string]::IsNullOrWhiteSpace($MgUser."Last Non-Interactive Sign-In"))){
		Try{
			[DateTime]$LNIS = $MgUser."Last Non-Interactive Sign-In"
			$Record."Azure Non-Interactive Last Sign-On" = $LNIS
			$Record."Azure Non-Interactive Last Sign-On Days" = ($((Get-Date) - $LNIS).Days)
		}catch {
			$Record."Azure Non-Interactive Last Sign-On" = [DateTime]$MgUser."Last Non-Interactive Sign-In"
		}
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

	#region ADFS 
	$ADFSLog = $ADFSLogs.Where({$_."User ID" -eq ($ADUser.sAMAccountName)}) | Sort-Object -Property "Date-Time" -Descending | Select-Object -First 1
	If(-Not [string]::IsNullOrWhiteSpace($ADFSLog."Date-Time")){
		$Record."ADFS Last Logon" = $ADFSLog."Date-Time"
		$Record."ADFS Last Logon Days" = ($((Get-Date)- ([DateTime]($ADFSLog."Date-Time"))).Days)
	}
	$Record."ADFS Last Logon IP" = $ADFSLog."IP Address"
	$Record."ADFS Relying Party" = $ADFSLog."Relying Party"
	$Record."ADFS Auth Protocol" = $ADFSLog."Auth Protocol"
	$Record."ADFS Network Location" = $ADFSLog."Network Location"
	$Record."ADFS ADFS Server" = $ADFSLog."ADFS Server"
	$Record."ADFS User Agent String" = $ADFSLog."User Agent String"
	
	#endregion ADFS 
	$Record."RDS CAL Expiration Date" = ($ADUser.msTSExpireDate)
	$Record."Distinguished Name" = ($ADUser.distinguishedName)

	#region E-Mail Information
	#region Mailbox
	If($ADUser.msExchMailboxGuid){
		$Record."Mailbox Creation Date" = $ADUser.msExchWhenMailboxCreated
		$LM = $EPMailboxes.Where({ $_.UserPrincipalName -eq $CUPN})
		If ($LM) {
			$Record."Mailbox Location" = "Local"
			$Record."Mailbox Server" = $LM.ServerName
			$Record."Mailbox Database" = $LM.Database
			$Record."Mailbox GUID" = $LM.ExchangeGuid
			$Record."Mailbox Use Database Quota Defaults" = $LM.UseDatabaseQuotaDefaults
			$Record."Mailbox Issue Warning Quota" = MailboxGB($LM.IssueWarningQuota)
			$Record."Mailbox Prohibit Send Quota" = MailboxGB($LM.ProhibitSendQuota)
			$Record."Mailbox Prohibit Send Receive Quota" = MailboxGB($LM.ProhibitSendReceiveQuota)

			$LMS =  $EPMailboxesStats.Where({$_.MailboxGuid -eq $ad.msExchMailboxGuid})
			If($LMS.MailboxGuid){
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
				$Record."Mailbox Forwarding Address" = $LMS.ForwardingAddress
				$Record."Mailbox Forwarding Address SMTP" = $LMS.ForwardingSmtpAddress
			}
			$LMCAS = $EPCASM.Where({ If($_.SamAccountName) {$_.SamAccountName.ToLower() -eq $ADUser.SamAccountName.ToLower()}})
			If ($LMCAS) {
				$Record."OWA Enabled" = $LMCAS.OWAEnabled
				$Record."Mapi Enabled" = $LMCAS.ActiveSyncEnabled
				$Record."Active Sync Enabled" = $LMCAS.MapiEnabled
				$Record."IMAP Enabled" = $LMCAS.ImapEnabled
				$Record."POP Enabled" = $LMCAS.PopEnabled
			}
			$LMFR = $EPMailboxesForwardingRules.Where({ If($_.SamAccountName) {$_.SamAccountName.ToLower() -eq $ADUser.SamAccountName.ToLower()}})
			If ($LMFR) {
				$Record."Mailbox Forwarding Rules" = $LMFR -join ","
			}

		}Else {
			$RM = $EXOMailboxes.Where({ $_.UserPrincipalName -eq $CUPN})
			$LRM = $EPRemoteMailboxes.Where({ $_.UserPrincipalName -eq $CUPN})
			If ($RM) {
				If ($RM -and $LRM) {
					If ($RM.ExchangeGuid -eq $LRM.ExchangeGuid) {
						$Record."Mailbox Location" = "Hybrid Remote"
					}Else {
						$Record."Mailbox Location" = "Hybrid Remote Broken"
						If (($LRM.ExchangeGuid -eq '00000000-0000-0000-0000-000000000000' -or $null -eq $LRM.ExchangeGuid)) {
							write-host ("`t Creating Remote Mailbox: " + $ADUser.ExchangeGuid) -ForeGroundColor red
							Enable-EPRemoteMailbox -Identity $ADUser.Alias -RemoteRoutingAddress ( $ADUser.Alias + "@" + ($azorg -split "\.")[0] + ".mail.onmicrosoft.com")
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
				#$RMS =  Get-MailboxStatistics -Identity $CUPN -ErrorAction SilentlyContinue
				$RMS =  $EXOMailboxesStats.Where({$_.MailboxGuid -eq $ad.msExchMailboxGuid})
				If($RMS.MailboxGuid){
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
					If ([string]::IsNullOrWhiteSpace($RMS.ForwardingAddress)){
						$Record."Mailbox Forwarding Address" = $LRM.ForwardingAddress
					}Else{
						$Record."Mailbox Forwarding Address" = $RMS.ForwardingAddress
					}
					If ([string]::IsNullOrWhiteSpace($RMS.ForwardingSmtpAddress)){
						$Record."Mailbox Forwarding Address" = $LRM.ForwardingSmtpAddress
					}Else{
						$Record."Mailbox Forwarding Address" = $RMS.ForwardingSmtpAddress
					}
				}
				$RMCAS = $EPRemoteCASM.Where({ $_.guid -eq (New-Object -TypeName System.Guid -ArgumentList @(,$ADUser.msExchMailboxGuid)).ToString()})
				# $RMCAS = $EPRemoteCASM.Where({ $_.SamAccountName.ToLower() -eq $ADUser.SamAccountName.ToLower()})
				If ($RMCAS) {
					$Record."OWA Enabled" = $RMCAS.OWAEnabled
					$Record."Mapi Enabled" = $RMCAS.ActiveSyncEnabled
					$Record."Active Sync Enabled" = $RMCAS.MapiEnabled
					$Record."IMAP Enabled" = $RMCAS.ImapEnabled
					$Record."POP Enabled" = $RMCAS.PopEnabled
				}
				$RMFR = $EXOMailboxesForwardingRules.Where({ If($_.SamAccountName) {$_.SamAccountName.ToLower() -eq $ADUser.SamAccountName.ToLower()}})
				If ($RMFR) {
					$Record."Mailbox Forwarding Rules" = $RMFR -join ","
				}
			}			
		}
		$DAMP = @()
		$CSVFixedDAMP = $null
		$FixedDAMP = @{}
		#region Mailbox Permission
		If ($Record."Mailbox Location" -eq "Local" -and $LM) {
			Try {
				$DAMP = ($EPMailboxesPerms.($CUPN).Where({($_.AccessRights -eq "FullAccess") -and ($_.User -notin $ExcludeUsers) -and ($_.User -notmatch "S-1-5-*" ) -and $_.IsInherited -eq $false})) | Sort-Object -Unique -Property User
			}Catch{
				$DAMP = ($EPMailboxesPerms[1].($CUPN).Where({($_.AccessRights -eq "FullAccess") -and ($_.User -notin $ExcludeUsers) -and ($_.User -notmatch "S-1-5-*" ) -and $_.IsInherited -eq $false})) | Sort-Object -Unique -Property User
			}
		}
		If($Record."Mailbox Location" -eq "Cloud Only" -or $Record."Mailbox Location" -match "Hybrid Remote" -and $RM){
			Try {
				$DAMP =  ($EXOMailboxesPerms.($CUPN).Where({($_.AccessRights -eq "FullAccess") -and ($_.User -notin $ExcludeUsers) -and ($_.User -notmatch "S-1-5-*" ) -and $_.IsInherited -eq $false})) | Sort-Object -Unique -Property User		
			}Catch{
				$DAMP =  ($EXOMailboxesPerms[1].($CUPN).Where({($_.AccessRights -eq "FullAccess") -and ($_.User -notin $ExcludeUsers) -and ($_.User -notmatch "S-1-5-*" ) -and $_.IsInherited -eq $false})) | Sort-Object -Unique -Property User		
			}
		}
		#Will show disabled users with permissions
		If ($DAMP.count -gt 0 ) {
			ForEach( $ACE in $DAMP) {
				If (($env:USERDOMAIN).ToLower() -eq (split-path -Path $ACE.User -Parent).ToLower()) {
					Try{
						$DomainObject = [DirectoryServices.AccountManagement.Principal]::FindByIdentity($ContextType,$IdentityType,(split-path -Path $ACE.User -Leaf))
						Foreach ($DO in $DomainObject) {
							If ($DO.StructuralObjectClass -eq  "user") {
								If ($DO.Enabled -and $DO.SamAccountName.ToLower() -eq (split-path -Path $ACE.User -Leaf).ToLower()) {
									$FixedDAMP.add($ACE.User,$ACE.AccessRights)
								}Else{
									If ($RemoveDisabledPerms -and $ADUser.description -notmatch "Leave") {
										If ($Record."Mailbox Location" -eq "Cloud Only") {
											If (Get-Module -ListAvailable -Name "ExchangeOnlineManagement") {
												If (-Not (Get-Module "ExchangeOnlineManagement" -ErrorAction SilentlyContinue)) {
													Import-Module "ExchangeOnlineManagement" -DisableNameChecking
												} 
											} Else {
												Import-Module PackageManagement
												Import-Module PowerShellGet
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
												If (Get-ChildItem "Cert:\LocalMachine\My\" | Where-Object {$_.Thumbprint -eq $AZCertThumbprint}) {
													Connect-ExchangeOnline -CertificateThumbPrint $AZCertThumbprint -AppID $AZClientID -Organization $AZOrg -ShowProgress:$false -ShowBanner:$false
												}Else{
													Connect-ExchangeOnline -ShowBanner:$false -UserPrincipalName $CUPN -ShowProgress:$false
												}
											}
											#Remove Permissions
											$RM | Remove-MailboxPermission -Confirm:$false -User $ACE.User -AccessRights $ACE.AccessRights
											$FixedDAMP.add(($ACE.User + " - Disabled - Removed"),$ACE.AccessRights)
										}Else{
											If ((Get-PSSession | Where-Object { $_.ConfigurationName -Match "Microsoft.Exchange" -and $_.ComputerName -eq $ExchangeServer}).Count -eq 0 ) {
												$ERPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://$ExchangeServer/PowerShell/ -Authentication Negotiate -AllowRedirection 
												Import-PSSession $ERPSession -AllowClobber -DisableNameChecking -Prefix "EP"
											}
											#Remove Permissions
											$LM | Remove-EPMailboxPermission -Confirm:$false -User $ACE.User -AccessRights $ACE.AccessRights
											$FixedDAMP.add(($ACE.User + " - Disabled - Removed"),$ACE.AccessRights)
										}

									}Else{
										$FixedDAMP.add(($ACE.User + " - Disabled"),$ACE.AccessRights)
									}
								}
							}elseif ($DO.StructuralObjectClass -eq  "group") {
								$FixedDAMP.add($ACE.User,$ACE.AccessRights)
							}
						}
					}Catch{

					}
				}else {
					$FixedDAMP.add($ACE.User,$ACE.AccessRights)
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
		$LM = $EPMailboxesArchive.Where({ $_.UserPrincipalName -eq $CUPN})
		If ($LM) {
			$Record."Mailbox Archive Location" = "Local"
			$Record."Mailbox Archive Server" = $LM.ServerName
			$Record."Mailbox Archive Database" = $LM.Database
			$Record."Mailbox Archive GUID" = $LM.ExchangeGuid
			$Record."Mailbox Archive Use Database Quota Defaults" = $LM.UseDatabaseQuotaDefaults
			$Record."Mailbox Archive Issue Warning Quota" = MailboxGB($LM.IssueWarningQuota)
			$Record."Mailbox Archive Prohibit Send Quota" = MailboxGB($LM.ProhibitSendQuota)
			$Record."Mailbox Archive Prohibit Send Receive Quota" = MailboxGB($LM.ProhibitSendReceiveQuota)

			$LMS =  $EPMailboxesArchiveStats.Where({$_.MailboxGuid -eq $LM.ArchiveGuid})
			If($LMS.MailboxGuid){
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
			$RM = $EXOMailboxesArchive.Where({ $_.UserPrincipalName -eq $CUPN})
			$LRM = $EPRemoteMailboxesArchive.Where({ $_.UserPrincipalName -eq $CUPN})
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

				$RMS = $EXOMailboxesArchiveStats.Where({$_.MailboxGUID -eq $RM.ArchiveGuid})
				If($RMS.MailboxGuid){
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
			Try {
				$DAMP = ($EPMailboxesArchivePerms.($LM.Identity).Where({($_.AccessRights -eq "FullAccess") -and ($_.User -notin $ExcludeUsers) -and ($_.User -notmatch "S-1-5-*" ) -and $_.IsInherited -eq $false})) | Sort-Object -Unique -Property User
			}Catch{
				$DAMP = ($EPMailboxesArchivePerms[1].($LM.Identity).Where({($_.AccessRights -eq "FullAccess") -and ($_.User -notin $ExcludeUsers) -and ($_.User -notmatch "S-1-5-*" ) -and $_.IsInherited -eq $false})) | Sort-Object -Unique -Property User
			}
		}
		If($Record."Mailbox Archive Location" -eq "Cloud Only" -or $Record."Mailbox Location" -match "Hybrid Remote"){	
			Try {
				$DAMP = ($EXOMailboxesArchivePerms.($RM.Identity).Where({($_.AccessRights -eq "FullAccess") -and ($_.User -notin $ExcludeUsers) -and ($_.User -notmatch "S-1-5-*" ) -and $_.IsInherited -eq $false})) | Sort-Object -Unique -Property User
			}Catch{
				$DAMP = ($EXOMailboxesArchivePerms[1].($RM.Identity).Where({($_.AccessRights -eq "FullAccess") -and ($_.User -notin $ExcludeUsers) -and ($_.User -notmatch "S-1-5-*" ) -and $_.IsInherited -eq $false})) | Sort-Object -Unique -Property User
			}
		}
		#Will show disabled users with permissions
		If ($DAMP.count -gt 0 ) {
			ForEach( $ACE in $DAMP) {
				If (($env:USERDOMAIN).ToLower() -eq (split-path -Path $ACE.User -Parent).ToLower()) {
					Try{
						$DomainObject = [DirectoryServices.AccountManagement.Principal]::FindByIdentity($ContextType,$IdentityType,(split-path -Path $ACE.User -Leaf))
						Foreach ($DO in $DomainObject) {
							If ($DO.StructuralObjectClass -eq  "user") {
								If ($DO.Enabled -and $DO.SamAccountName.ToLower() -eq (split-path -Path $ACE.User -Leaf).ToLower()) {
									$FixedDAMP.add($ACE.User,$ACE.AccessRights)
								}Else{
									If ($RemoveDisabledPerms -and $ADUser.description -notmatch "Leave") {
										If ($Record."Mailbox Location" -eq "Cloud Only") {
											If (Get-Module -ListAvailable -Name "ExchangeOnlineManagement") {
												If (-Not (Get-Module "ExchangeOnlineManagement" -ErrorAction SilentlyContinue)) {
													Import-Module "ExchangeOnlineManagement" -DisableNameChecking
												} 
											} Else {
												Import-Module PackageManagement
												Import-Module PowerShellGet
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
												If (Get-ChildItem "Cert:\LocalMachine\My\" | Where-Object {$_.Thumbprint -eq $AZCertThumbprint}) {
													Connect-ExchangeOnline -CertificateThumbPrint $AZCertThumbprint -AppID $AZClientID -Organization $AZOrg -ShowProgress:$false -ShowBanner:$false
												}Else{
													Connect-ExchangeOnline -ShowBanner:$false -UserPrincipalName $CUPN -ShowProgress:$false
												}
											}

											#Remove Permissions
											$RM | Remove-MailboxPermission -Confirm:$false -User $ACE.User -AccessRights $ACE.AccessRights
											$FixedDAMP.add(($ACE.User + " - Disabled - Removed"),$ACE.AccessRights)
										}Else{
											If ((Get-PSSession | Where-Object { $_.ConfigurationName -Match "Microsoft.Exchange" -and $_.ComputerName -eq $ExchangeServer}).Count -eq 0 ) {
												$ERPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://$ExchangeServer/PowerShell/ -Authentication Negotiate -AllowRedirection 
												Import-PSSession $ERPSession -AllowClobber -DisableNameChecking -Prefix "EP"
											}
											#Remove Permissions
											$LM | Remove-EPMailboxPermission -Confirm:$false -User $ACE.User -AccessRights $ACE.AccessRights
											$FixedDAMP.add(($ACE.User + " - Disabled - Removed"),$ACE.AccessRights)
										}

									}Else{
										$FixedDAMP.add(($ACE.User + " - Disabled"),$ACE.AccessRights)
									}
								}
							}elseif ($DO.StructuralObjectClass -eq  "group") {
								$FixedDAMP.add($ACE.User,$ACE.AccessRights)
							}
						}
					} Catch {

					}
				}else {
					$FixedDAMP.add($ACE.User,$ACE.AccessRights)
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

If ($output.Length -gt 0) {
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
	$excel = $output | Export-Excel -Path $xlsxfile -ClearSheet -WorksheetName ("Hybrid_Info_" + $FileDate) -AutoFilter -FreezeTopRowFirstColumn -PassThru
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
	$DSCOutput = $output.Where({$_."Days Since Last LogOn" -gt 60 -and $_."Days Since Creation" -gt 60 -and $_."Days from last password change" -gt 60 -and $_."Account Status" -eq "Enabled" -and [string]::IsNullOrWhiteSpace($_."Employee Type")})

	$WorksheetName = "Look to Disable"
	If ($DSCOutput.count -gt 0) {
		$DSCOutput | Export-Excel -ExcelPackage $excel -ClearSheet -WorksheetName $WorksheetName -AutoFilter -AutoSize -FreezeTopRowFirstColumn

		$ws = $excel.Workbook.Worksheets[$WorksheetName]
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
	}

	Close-ExcelPackage $excel
	$swo.Stop()
	Write-Host ("Saving output time: " + (FormatElapsedTime($swo.Elapsed)) + " to run. ")
	Remove-Variable "output"
	Remove-Variable "excel"
}
$sw.Stop()
Write-Host ("Script runtime: " + (FormatElapsedTime($sw.Elapsed)) + " to run. " + '{0:N0}' -f ($ADUsers.Count / $sw.Elapsed.TotalMinutes) + " Users's per Minute.")
#endregion Excel convert

