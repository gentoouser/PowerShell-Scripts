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
    

#>

#region Variables 
$output = @()
$FileDate = (Get-Date -format yyyyMMdd-hhmm)
$csvfile =  ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\Logs\" + `
			($MyInvocation.MyCommand.Name -replace ".ps1","") + "_" + `
			$FileDate + ".csv")
$xlsxfile =  ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\Logs\" + `
			($MyInvocation.MyCommand.Name -replace ".ps1","") + "_" + `
			$FileDate + ".xlsx")
$ExcludeUsers=@(
	($env:USERDOMAIN + "\Domain Admins"),
	($env:USERDOMAIN + "\Enterprise Admins"),
	($env:USERDOMAIN + "\Organization Management"),
	($env:USERDOMAIN + "\Exchange Servers"),
	($env:USERDOMAIN + "\Exchange Domain Servers"),
	($env:USERDOMAIN + "\Administrators"),
	"NT AUTHORITY\SYSTEM",
	"NT AUTHORITY\SELF"
)
$ExchangeServer = ""

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
	${Employee Type}
	${Employee Number}
	${Office}
	${Phone}
	${Mobile Phone}
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
	Write-Host ("Loading Active Directory Plugins") -foregroundcolor "Green"
	Import-Module "ActiveDirectory"  -ErrorAction SilentlyContinue
} Else {
	Write-Host ("Active Directory Plug-ins Already Loaded") -foregroundcolor "Green"
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
	Write-Host ("Loading Exchange Plugins") -foregroundcolor "Green"
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
	Write-Host ("Exchange Plug-ins Already Loaded") -foregroundcolor "Green"
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


#Get All AD users
$ADUsers = Get-ADUser -server $ADServer -searchbase $SearchBase -Filter * -Properties * 
#Get All Azure AD users
$AzureUsers = Get-AzureADUser -All:$true

# update progress bar every 0.1 %
#$interval = $ADUsers.count / 1000

Write-host "Processing Users"
#Main loop
Foreach ($ADUser in $ADUsers) { 
	Write-host ("`t" + $ADUser.name)
	#Update progress
	If ($output.Count -gt 0) {
		# If ($output.Count % $interval -eq 0) {
		If ($output.Count % 2 -eq 0 -or $output.Count % 2 -eq 1) {
			Write-Progress -Activity ("Creating " + ($MyInvocation.MyCommand.Name -replace ".ps1","") + " output") -Status ("Processing " + $ADUser.name + "[" + $output.Count + "/" + $ADUsers.count + "]") -PercentComplete ($output.Count *  100 / $ADUsers.count)
		}
	}
	#Clear Loop var
	$TotalItemSize = 0
	$TotalDeletedItemSize = 0

	$AzureADUser = $AzureUsers | Where-Object {$_.UserPrincipalName -eq $ADUser.UserPrincipalName}
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
	$Record."Employee Type" = $ADUser.employeeType
	$Record."Employee Number" = $ADUser.EmployeeNumber
	$Record."Office" = $ADUser.physicalDeliveryOfficeName
	$Record."Phone" = $ADUser.telephoneNumber
	$Record."Mobile Phone" = $ADUser.mobile
	If($AzureADUser){
		$Record."Azure Licenses" = ((Get-AzureADUserLicenseDetail -ObjectId $AzureADUser.ObjectId | Select-Object SkuPartNumber).SkuPartNumber -join ",")
		$Record."Azure Last Sync Time" = $AzureADUser.LastDirSyncTime
	}
	$Record."Group Membership" = (($ADUser.MemberOf | CleanDistinguishedName) -join ",")
	$Record."Email" = $ADUser.Mail

	#region Manager name
	If($ADUser.Manager){
		$Record."Manager" = ($ADUsers | Where-Object {$_.DistinguishedName -eq $ADUser.Manager}).DisplayName
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
		$LM = Get-EPMailbox -Identity $ADUser.UserPrincipalName -ErrorAction SilentlyContinue
		If ($LM) {
			$Record."Mailbox Location" = "Local"
			$Record."Mailbox Server" = $LM.ServerName
			$Record."Mailbox Database" = $LM.Database
			$Record."Mailbox GUID" = $LM.ExchangeGuide
			$Record."Mailbox Use Database Quota Defaults" = $LM.UseDatabaseQuotaDefaults
			$Record."Mailbox Issue Warning Quota" = MailboxGB($LM.IssueWarningQuota)
			$Record."Mailbox Prohibit Send Quota" = MailboxGB($LM.ProhibitSendQuota)
			$Record."Mailbox Prohibit Send Receive Quota" = MailboxGB($LM.ProhibitSendReceiveQuota)

			$LMS =  Get-EPMailboxStatistics -Identity $ADUser.UserPrincipalName -ErrorAction SilentlyContinu
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
		}Else {
			$RM = Get-Mailbox -Identity $ADUser.UserPrincipalName -ErrorAction SilentlyContinue
			$LRM = Get-EPRemoteMailbox -Identity $ADUser.UserPrincipalName -ErrorAction SilentlyContinue
			If ($RM) {
				If ($RM -and $LRM) {
					If ($RM.ExchangeGuid -eq $LRM.ExchangeGuid) {
						$Record."Mailbox Location" = "Hybrid Remote"
					}Else {
						$Record."Mailbox Location" = "Hybrid Remote Broken"			
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
				$RMS =  Get-MailboxStatistics -Identity $ADUser.UserPrincipalName -ErrorAction SilentlyContinue
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
			}			
		}
		$DAMP = $null
		$CSVFixedDAMP = $null
		$FixedDAMP = @{}
		#region Mailbox Permission
		If ($Record."Mailbox Location" -eq "Local") {
			$DAMP = Get-EPMailboxPermission  -Identity $ADUser.UserPrincipalName | Where-Object {($_.AccessRights -eq "FullAccess") -and ($_.User -notin $ExcludeUsers) -and ($_.User -notmatch "S-1-5-*" ) -and $_.IsInherited -eq $false} | Sort-Object -Unique -Property User
		}
		If($Record."Mailbox Location" -eq "Cloud Only" -or $Record."Mailbox Location" -match "Hybrid Remote"){
			$DAMP = Get-MailboxPermission  -Identity $ADUser.UserPrincipalName | Where-Object {($_.AccessRights -eq "FullAccess") -and ($_.User -notin $ExcludeUsers) -and ($_.User -notmatch "S-1-5-*" ) -and $_.IsInherited -eq $false} | Sort-Object -Unique -Property User		
		}
		#Will remove disabled users with permissions
		If ($DAMP.count -gt 0 ) {
			ForEach( $ACE in $DAMP) {
				[switch]$ADValid=$false
				If (($env:USERDOMAIN).ToLower() -eq (split-path -Path $ACE.User -Parent).ToLower()) {
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
		$LM = Get-EPMailbox -Archive -Identity $ADUser.UserPrincipalName -ErrorAction SilentlyContinue
		If ($LM) {
			$Record."Mailbox Archive Location" = "Local"
			$Record."Mailbox Archive Server" = $LM.ServerName
			$Record."Mailbox Archive Database" = $LM.Database
			$Record."Mailbox Archive GUID" = $LM.ExchangeGuide
			$Record."Mailbox Archive Use Database Quota Defaults" = $LM.UseDatabaseQuotaDefaults
			$Record."Mailbox Archive Issue Warning Quota" = MailboxGB($LM.IssueWarningQuota)
			$Record."Mailbox Archive Prohibit Send Quota" = MailboxGB($LM.ProhibitSendQuota)
			$Record."Mailbox Archive Prohibit Send Receive Quota" = MailboxGB($LM.ProhibitSendReceiveQuota)

			$LMS =  Get-EPMailboxStatistics -Archive -Identity $ADUser.UserPrincipalName -ErrorAction SilentlyContinu
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
			$RM = Get-Mailbox -Archive -Identity $ADUser.UserPrincipalName -ErrorAction SilentlyContinue
			$LRM = Get-EPRemoteMailbox -Archive -Identity $ADUser.UserPrincipalName -ErrorAction SilentlyContinue
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
				$RMS =  Get-MailboxStatistics -Archive -Identity $ADUser.UserPrincipalName -ErrorAction SilentlyContinue
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
			$DAMP = $LM | Get-EPMailboxPermission | Where-Object {($_.AccessRights -eq "FullAccess") -and ($_.User -notin $ExcludeUsers) -and ($_.User -notmatch "S-1-5-*" ) -and $_.IsInherited -eq $false} | Sort-Object -Unique -Property User
		}
		If($Record."Mailbox Archive Location" -eq "Cloud Only" -or $Record."Mailbox Location" -match "Hybrid Remote"){
			$DAMP = $RM | Get-MailboxPermission | Where-Object {($_.AccessRights -eq "FullAccess") -and ($_.User -notin $ExcludeUsers) -and ($_.User -notmatch "S-1-5-*" ) -and $_.IsInherited -eq $false} | Sort-Object -Unique -Property User		
		}
		#Will remove disabled users with permissions
		If ($DAMP.count -gt 0 ) {
			ForEach( $ACE in $DAMP) {
				[switch]$ADValid=$false
				If (($env:USERDOMAIN).ToLower() -eq (split-path -Path $ACE.User -Parent).ToLower()) {
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

#Days Since Last LogOn
# Add-ConditionalFormatting -WorkSheet $ws -address (($ws.Cells[1,$htHeader["Days Since Last LogOn"]].Address).Substring(0,($ws.Cells[1,$htHeader["Days Since Last LogOn"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Days Since Last LogOn"]].Address).Substring(0,($ws.Cells[1,$htHeader["Days Since Last LogOn"]].Address).Length-1) + $Lastrow) -RuleType TwoColorScale

#Mailbox Item Count comma's
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["Days Since Last Log-On"]].Address).Substring(0,($ws.Cells[1,$htHeader["Days Since Last Log-On"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Days Since Last Log-On"]].Address).Substring(0,($ws.Cells[1,$htHeader["Days Since Last Log-On"]].Address).Length-1) + $Lastrow) -NumberFormat '#,##0'

#Mailbox Size (GB) comma's
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["Mailbox Size"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Size"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Mailbox Size"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Size"]].Address).Length-1) + $Lastrow) -NumberFormat '#,##0'
#Mailbox Item Count comma's
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["Mailbox Item Count"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Item Count"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Mailbox Item Count"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Item Count"]].Address).Length-1) + $Lastrow) -NumberFormat '#,##0'
#Mailbox Last Logged On User Account Date
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["Mailbox Last Logged On User Account"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Last Logged On User Account"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Mailbox Last Logged On User Account"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Last Logged On User Account"]].Address).Length-1) + $Lastrow) -NumberFormat 'Date-Time'
#Mailbox Last Logon Time Date
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["Mailbox Last Logon Time"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Last Logon Time"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Mailbox Last Logon Time"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Last Logon Time"]].Address).Length-1) + $Lastrow) -NumberFormat 'Date-Time'
#Mailbox Last Logoff Time Date
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["Mailbox Last Logoff Time"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Last Logoff Time"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Mailbox Last Logoff Time"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Last Logoff Time"]].Address).Length-1) + $Lastrow) -NumberFormat 'Date-Time'
#Mailbox Creation Date Date
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["Mailbox Creation Date"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Creation Date"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Mailbox Creation Date"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Creation Date"]].Address).Length-1) + $Lastrow) -NumberFormat 'Date-Time'




#Mailbox Archive Size (GB) comma's
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["Mailbox Archive Size"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Archive Size"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Mailbox Archive Size"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Archive Size"]].Address).Length-1) + $Lastrow) -NumberFormat '#,##0'
#Mailbox Archive Item Count comma's
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["Mailbox Archive Item Count"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Archive Item Count"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Mailbox Archive Item Count"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Archive Item Count"]].Address).Length-1) + $Lastrow) -NumberFormat '#,##0'
#Mailbox Archive Last Logged On User Account Date
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["Mailbox Archive Last Logged On User Account"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Archive Last Logged On User Account"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Mailbox Archive Last Logged On User Account"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Archive Last Logged On User Account"]].Address).Length-1) + $Lastrow) -NumberFormat 'Date-Time'
#Mailbox Archive Last Logon Time Date
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["Mailbox Archive Last Logon Time"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Archive Last Logon Time"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Mailbox Archive Last Logon Time"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Archive Last Logon Time"]].Address).Length-1) + $Lastrow) -NumberFormat 'Date-Time'
#Mailbox Archive Last Logoff Time Date
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["Mailbox Archive Last Logoff Time"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Archive Last Logoff Time"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Mailbox Archive Last Logoff Time"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Archive Last Logoff Time"]].Address).Length-1) + $Lastrow) -NumberFormat 'Date-Time'
#Mailbox Archive Creation Date Date
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["Mailbox Archive Creation Date"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Archive Creation Date"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Mailbox Archive Creation Date"]].Address).Substring(0,($ws.Cells[1,$htHeader["Mailbox Archive Creation Date"]].Address).Length-1) + $Lastrow) -NumberFormat 'Date-Time'


#Last Log-On Date
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["Days Since Last Log-On"]].Address).Substring(0,($ws.Cells[1,$htHeader["Days Since Last Log-On"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Days Since Last Log-On"]].Address).Substring(0,($ws.Cells[1,$htHeader["Days Since Last Log-On"]].Address).Length-1) + $Lastrow) -NumberFormat 'Date-Time'
#Days Since Last Log-On comma's
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["Days Since Last Log-On"]].Address).Substring(0,($ws.Cells[1,$htHeader["Days Since Last Log-On"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Days Since Last Log-On"]].Address).Substring(0,($ws.Cells[1,$htHeader["Days Since Last Log-On"]].Address).Length-1) + $Lastrow) -NumberFormat '#,##0'
#Creation Date Date
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["Creation Date"]].Address).Substring(0,($ws.Cells[1,$htHeader["Creation Date"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Creation Date"]].Address).Substring(0,($ws.Cells[1,$htHeader["Creation Date"]].Address).Length-1) + $Lastrow) -NumberFormat 'Date-Time'
#Days Since Creation comma's
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["Days Since Creation"]].Address).Substring(0,($ws.Cells[1,$htHeader["Days Since Creation"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Days Since Creation"]].Address).Substring(0,($ws.Cells[1,$htHeader["Days Since Creation"]].Address).Length-1) + $Lastrow) -NumberFormat '#,##0'
#Last Password Change Date
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["Last Password Change"]].Address).Substring(0,($ws.Cells[1,$htHeader["Last Password Change"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Last Password Change"]].Address).Substring(0,($ws.Cells[1,$htHeader["Last Password Change"]].Address).Length-1) + $Lastrow) -NumberFormat 'Date-Time'
#Days from last password changecomma's
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["Days from last password change"]].Address).Substring(0,($ws.Cells[1,$htHeader["Days from last password change"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Days from last password change"]].Address).Substring(0,($ws.Cells[1,$htHeader["Days from last password change"]].Address).Length-1) + $Lastrow) -NumberFormat '#,##0'
#RDS CAL Expiration Date Date
Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader["RDS CAL Expiration Date"]].Address).Substring(0,($ws.Cells[1,$htHeader["RDS CAL Expiration Date"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["RDS CAL Expiration Date"]].Address).Substring(0,($ws.Cells[1,$htHeader["RDS CAL Expiration Date"]].Address).Length-1) + $Lastrow) -NumberFormat 'Date-Time'

#New worksheet that has created,pw change, last logon all over 60 days.
$output | Where-Object {$_."Days Since Last LogOn" -gt 60 -and $_."Days Since Creation" -gt 60 -and $_."Days from last password change" -gt 60} | Export-Excel -Path $xlsxfile -ClearSheet -WorksheetName "Look to Disable" -AutoFilter -AutoSize -FreezeTopRowFirstColumn


Close-ExcelPackage $excel

Remove-Variable "output"
Remove-Variable "excel"
#endregion Excel convert
