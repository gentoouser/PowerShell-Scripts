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
    * Version 1.01.04 - Added departmentnumber and extensionAttribute2.
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
    * Version 1.03.11 - Fixed errors revolving around date being null
    * Version 1.03.12 - Added worksheet for AD Groups
    * Version 1.03.13 - Fix issue with Mailbox size showing.
    * Version 1.03.14 - Cleanup output from runspace. Fix MemberCount in AD Groups. 
    * Version 2.00.00 - Batched Main Loop.  
    * Version 2.00.01 - Batched Main loop working but now need to get data for other tabs.
    * Version 2.00.02 - Optimized Group output.
    * Version 2.00.03 - Fix Caching jobs.
    * Version 2.00.04 - Optimized AD caching.
    * Version 3.00.00 - Get all data from AD from .Net code. and then update Class Objects in Exchange and Entra ID jobs. Doing parallel tasks if running in powershell 7.
    * Version 3.00.01 - Fixed issue with Archive mailboxes not ignoring ExcludeUsers. Added column for "Password Change on Next Logon"
    * Version 3.00.02 - Fixed issue with "Password Change on Next Logon" and extra columns in excel. Cleaned up Excel formatting and made it easier update.
    * Version 3.00.03 - 20250418 - Added worksheet for Disabled users with mailboxes.
    * Version 3.00.04 - 20250501 - Add worksheet for large mailboxes. Fixed issues with filtered worksheets having issues with formatting. Added code to clean up old logs.
    
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
	$LogFile =  ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\Logs\" + `
                ($MyInvocation.MyCommand.Name -replace ".ps1","") + "_" + `
                (Get-Date -format yyyyMMdd-hhmm) + ".txt"),
    [int]$processbatchjobs = ([Environment]::ProcessorCount),
	[int]$CacheRefresh = 4,
    [int]$LogPrune = -90,
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
    [string]$ExchangeServer              = "exchane.github.com",
	[string]$AzureTenant                 = "uid",
    [string]$AZClientID                  = "uid",
    [string]$AZCertThumbprint            = "Thumbprint",
	[String]$AZOrg                       =	"github.onmicrosoft.com",
	[switch]$RemoveDisabledPerms         =	$False,
	[switch]$IgnoreJobs                  =	$False,
	# [array]$ADFSServers=@()
    [array]$ADFSServers=@(
        "ADFSPD01.github.com"
        "ADFSPD02.github.com"

    )
)
#endregion Parameters
#region Variables 
$ScriptVersion = "3.0.4"

$FileDate = (Get-Date -format yyyyMMdd-hhmm)
$swc = [Diagnostics.Stopwatch]::StartNew()
If(-Not (Get-variable -Name Configuration -ErrorAction SilentlyContinue)) {
    $Global:Configuration = [hashtable]::Synchronized(@{})
}
$Configuration.FileDate = $FileDate
$Configuration.csvfile = $csvfile
$Configuration.xlsxfile = $xlsxfile
$Configuration.LogFile = $LogFile
$Configuration.CacheRefresh = $CacheRefresh
$Configuration.ExcludeUsers = $ExcludeUsers
$Configuration.AzureTenant = $AzureTenant
$Configuration.AZClientID = $AZClientID
$Configuration.AZCertThumbprint = $AZCertThumbprint
$Configuration.AZOrg =	$AZOrg
$Configuration.ExchangeServer = $ExchangeServer
$Configuration.RemoveDisabledPerms = $RemoveDisabledPerms
$Configuration.IgnoreJobs = $IgnoreJobs
$Configuration.ADFSServers = $ADFSServers
$Configuration.Jobs = @{}
$Configuration.ScriptName = ($MyInvocation.MyCommand.Name -replace ".ps1","")
$Configuration.ScriptVersion = $ScriptVersion
$Configuration.ScriptPath =(Split-Path -Parent -Path $MyInvocation.MyCommand.Definition)
If($processbatchjobs -eq ([Environment]::ProcessorCount)){
	If (([math]::ceiling(([Environment]::ProcessorCount)/8)) -lt 2){
		$processbatchjobs = 2
	}Else{
		$processbatchjobs =([math]::ceiling(([Environment]::ProcessorCount)/4))
	}
}ElseIf($processbatchjobs -gt ([Environment]::ProcessorCount)){
	$processbatchjobs = ([Environment]::ProcessorCount)
}ElseIf($processbatchjobs -lt 2){
	$processbatchjobs = 2
}
$Configuration.ProcessBatchJobs = $processbatchjobs
#https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/manage/understand-security-groups#default-active-directory-security-groups
$Configuration.AD_Default_Groups = @(
	"Access Control Assistance Operators",
	"Account Operators",
	"Administrators",
	"Allowed RODC Password Replication",
	"Backup Operators",
	"Certificate Service DCOM Access",
	"Cert Publishers",
	"Cloneable Domain Controllers",
	"Cryptographic Operators",
	"Denied RODC Password Replication",
	"Device Owners",
	"DHCP Administrators",
	"DHCP Users",
	"Distributed COM Users",
	"DnsUpdateProxy",
	"DnsAdmins",
	"Domain Admins",
	"Domain Computers",
	"Domain Controllers",
	"Domain Guests",
	"Domain Users",
	"Enterprise Admins",
	"Enterprise Key Admins",
	"Enterprise Read-only Domain Controllers",
	"Event Log Readers",
	"Group Policy Creator Owners",
	"Guests",
	"Hyper-V Administrators",
	"IIS_IUSRS",
	"Incoming Forest Trust Builders",
	"Key Admins",
	"Network Configuration Operators",
	"Performance Log Users",
	"Performance Monitor Users",
	"Pre-Windows 2000 Compatible Access",
	"Print Operators",
	"Protected Users",
	"RAS and IAS Servers",
	"RDS Endpoint Servers",
	"RDS Management Servers",
	"RDS Remote Access Servers",
	"Read-only Domain Controllers",
	"Remote Desktop Users",
	"Remote Management Users",
	"Replicator",
	"Schema Admins",
	"Security Administrator",
	"Security Reader",
	"Server Operators",
	"Storage Replica Administrators",
	"System Managed Accounts",
	"Terminal Server License Servers",
	"Users",
	"Windows Authorization Access",
	"WinRMRemoteWMIUsers_"
)
#https://learn.microsoft.com/en-us/exchange/plan-and-deploy/active-directory/ad-changes?view=exchserver-2019#prepare-active-directory-containers-objects-and-other-items
$Configuration.Exchange_Security_Groups = @(
	"Compliance Management",
	"Compliance",
	"Delegated Setup",
	"Discovery Management",
	"Discovery",
	"Exchange All Hosted Organizations",
	"Exchange Servers",
	"Exchange Trusted Subsystem",
	"Exchange Windows Permissions",
	"ExchangeLegacyInterop",
	"Help Desk",
	"Hygiene Management",
	"Hygiene",
	"Managed Availability Servers",
	"Organization Management",
	"Organization",
	"Public Folder Management",
	"Public Folder",
	"Recipient Management",
	"Recipient",
	"Records Management",
	"Records",
	"Security Administrator",
	"Security Reader",
	"Server Management",
	"Server",
	"View-Only Organization"
)
$Configuration.CurrentUserUPN= [string]([ADSI]"LDAP://<SID=$([System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value)>").UserPrincipalName
$Configuration.CSVtoReturnHeaders = @(
    "Azure Licenses",
    "Azure Licenses Details",
    "Group Membership",
    "Email Addresses",
    "Exchange Mobile Devices",
    "Mailbox Permissions",
    "Mailbox Forwarding Rules",
    "Mailbox Archive Permissions"
)
$Configuration.CommaHeaders = @(
    "Mailbox Size GB",
    "Mailbox Item Count",
    "Mailbox Archive Item Count",
    "Mailbox Last Logged On User Account",
    "Days Since Last Log-On",
    "Days Since Creation",
    "Days from last password change"
)
$Configuration.DateTimeHeaders = @(
    "Last Log-On Date",
    "Creation Date",
    "Last Password Change",
    "Mailbox Last Logon Time",
    "Mailbox Last Logoff Time",
    "Mailbox Creation Date",
    "Mailbox Archive Last Logon Time",
    "Mailbox Archive Last Logoff Time",
    "Mailbox Archive Creation Date",
    "RDS CAL Expiration Date"
)
#Program Classes and Functions
$Configuration.ClassADExchangeOutput = @'
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
	${Azure Licenses}
	${Azure Licenses Details}
	${Azure Last Sync Time}
	${Azure Last Sign-On}
	${Azure Last Sign-On Days}
	${Azure Last Non-Interactive Sign-On}
	${Azure Last Non-Interactive Sign-On Days}
	${Azure User Type}
	${Azure Security Identifier}
	${Azure ImmutableId}
	${Azure ID}
	${Azure Account Enabled}
	${Group Membership}
	${Manager} 
	${Home Directory} 
	${Account Status} 
	${Password Never Expires} 
	${Password Change on Next Logon} 
	${Password Not Required} 
	${Smartcard Logon Required} 
	${Account Trusted for Delegation}
	${Last Log-On Date} 
	${Days Since Last Log-On} 
	${Creation Date} 
	${Days Since Creation} 
	${Last Password Change} 
	${Days from last password change} 
	${RDS CAL Expiration Date}
	${Email} 
	${Email Addresses} 
    ${Exchange Remote Routing Address} 
	${Exchange Recipient Type} 
	${Exchange Mobile Devices} 
	${Mailbox Location}
	${Mailbox Server}
	${Mailbox Database}
	${Mailbox Permissions}
    ${Mailbox Usage (%)}
	${Mailbox Issue Warning Quota}
	${Mailbox Prohibit Send Quota}
	${Mailbox Prohibit Send Receive Quota}
	${Mailbox Use Database Quota Defaults}
	${Mailbox Size GB} 
	${Mailbox Item Count} 
	${Mailbox Recoverable Items Size GB} 
	${Mailbox Last Logged On User Account}
	${Mailbox Last Logon Time}
	${Mailbox Last Logoff Time}
	${Mailbox Forwarding Address}
	${Mailbox Forwarding Address SMTP}
	${Mailbox Forwarding Rules}
    ${Mailbox Creation Date}
	${OWA Enabled}
	${Mapi Enabled}
	${Active Sync Enabled}
	${IMAP Enabled}
	${POP Enabled}
	${Mailbox GUID}
	${Online Mailbox GUID}
	${Online Mailbox Archive GUID}
    ${Mailbox Archive GUID} 
	${Mailbox Archive Status}  
	${Mailbox Archive State}  
	${Mailbox Archive Location}  
	${Mailbox Archive Server}  
	${Mailbox Archive Database}  
	${Mailbox Archive Permissions}  
	${Mailbox Archive Issue Warning Quota} 
	${Mailbox Archive Prohibit Send Quota} 
	${Mailbox Archive Prohibit Send Receive Quota}  
	${Mailbox Archive Use Database Quota Defaults}  
	${Mailbox Archive Storage Limit Status} 
	${Mailbox Archive Size GB}
	${Mailbox Archive Item Count}
	${Mailbox Archive Last Logged On User Account}  
	${Mailbox Archive Last Logon Time}
	${Mailbox Archive Last Logoff Time}
	${Mailbox Archive Forwarding Address}
	${Mailbox Archive Forwarding Address SMTP}
	${Mailbox Archive Creation Date}
	${Mailbox Archive Auto Expanding Archive}
	${ADFS Last Logon} 
	${ADFS Last Logon Days} 
	${ADFS Last Logon IP} 
	${ADFS Relying Party} 
	${ADFS Auth Protocol} 
	${ADFS Network Location} 
	${ADFS ADFS Server} 
	${ADFS User Agent String} 
	${Distinguished Name}
}
Class ADGroupOutputClass {
    ${Name}
    ${sAMAccountName}
    ${description}
    ${Security Group}
    ${type}
    ${mail}
    ${objectSid}
    ${distinguishedName}
    ${Member Count}
    ${Members}
}
Class ADFSEventRecord {
    ${Date-Time}
    ${ADFS Server}
    ${User ID}
    ${Relying Party}
    ${Auth Protocol}
    ${Network Location}
    ${IP Address}
    ${User Agent String}
    ${Comments}
}

Function Clean-DistinguishedName{
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
Function Format-ElapsedTime($ts) {
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
Function Format-MailboxGB {
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
Function Write-TMOutput {
	<#
	.SYNOPSIS
		Sends the specified objects to the next command in the pipeline. If the command is the last command in the pipeline, the objects are displayed in the console.
	 
	.DESCRIPTION
		Sends the specified objects to the next command in the pipeline. If the command is the last command in the pipeline, the objects are displayed in the console. This function is a wrapper for the Write-Output cmdlet with some additional features. It allows for changing the color of the text (foreground color), and the color behind the text (background color). It also allows for horizontal and vertical padding.
	 
	.PARAMETER InputObject
		Specifies the objects to send down the pipeline. Enter a variable that contains the objects, or type a command or expression that gets the objects.
	 
	.PARAMETER ForegroundColor
		Specifies the text color. The default is the current foreground color.
	 
	.PARAMETER BackgroundColor
		Specifies the background color. The default is the current background color.
	 
	.PARAMETER HorizontalPad
		Specifies the amount of space on each side of the included objects. If this value doubled and added to length of the object is greater than the width of the console, it will wrap and likely cause unintended results.
	 
	.PARAMETER VerticalPad
		Specifies the number of blank lines above and below the included objects.
	 
	.PARAMETER NoEnumerate
		By default, the Write-Output cmdlet always enumerates its output. The NoEnumerate parameter suppresses the default behavior, and prevents Write-Output from enumerating output. The NoEnumerate parameter has no effect on collections that were created by wrapping commands in parentheses, because the parentheses force enumeration. The Write-TMOutput function is a wrapper for Write-Output cmdlet.
	 
	.INPUTS
		System.Management.Automation.PSObject
		You can pipe objects to Write-TMOutput.
	 
	.OUTPUTS
		System.Management.Automation.PSObject
		Write-TMOutput returns the objects that are submitted as input.
	 
	.EXAMPLE
		Write-TMOutput -InputObject 'Testing 1, 2, 3'
		This example writes the text object to the console.
	 
	.EXAMPLE
		Write-TMOutput -InputObject 'Testing 1, 2, 3' -ForegroundColor Gray -BackgroundColor Black
		This example writes the text object to the console in a gray font on a black background.
	 
	.EXAMPLE
		Write-TMOutput -InputObject 'Testing 1, 2, 3' -ForegroundColor Gray -BackgroundColor Black -HorizontalPad 10 -VerticalPad 2
		This example writes the text object to the console in a gray font on a black background. It will also pad the area around the text object with the background color in both horizontal and vertical directions.
	 
	.NOTES
		NAME: Write-TMOutput
		AUTHOR: Tommy Maynard
		LASTEDIT: 05/12/2016
		VERSION 1.1: Removed ISE protection code; PowerShellHostName (currently) requires ConsoleHost only.
		PERSONAL WEBSITE POST: http://tommymaynard.com/quick-learn-write-output-gets-foreground-and-background-colors-and-more-2016
		Powershell Gallery: https://www.powershellgallery.com/packages/TMOutput/1.1
	#>
		[CmdletBinding()]
		Param (
			[Parameter(ValueFromPipeline=$true)]
			[psobject[]]$InputObject,
	
			[Parameter()]
			[ValidateSet('Black','Blue','Cyan','DarkBlue','DarkCyan',
						'DarkGray','DarkGreen','DarkMagenta','DarkRed','DarkYellow',
						'Gray','Green','Magenta','Red','White','Yellow')]      
			[string]$ForegroundColor = [System.Console]::ForegroundColor,
	
			[Parameter()]
			[ValidateSet('Black','Blue','Cyan','DarkBlue','DarkCyan',
						'DarkGray','DarkGreen','DarkMagenta','DarkRed','DarkYellow',
						'Gray','Green','Magenta','Red','White','Yellow')]    
			[string]$BackgroundColor = [System.Console]::BackgroundColor,
	
			[Parameter()]
			[int]$HorizontalPad = 0,
	
			[Parameter()]
			[int]$VerticalPad = 0,
	
			[Parameter()]
			[switch]$NoEnumerate
		)
	
		##### BEGIN.
		Begin {
			# Collect default foreground and background colors.
			$ResetColorCheck = [bool]([System.Console] | Get-Member -Static -MemberType Method -Name ResetColor)
			If (-Not $ResetColorCheck) {
				$DefaultConsoleForegroundColor = [System.Console]::ForegroundColor
				$DefaultConsoleBackgroundColor = [System.Console]::BackgroundColor
			}
		} # End Begin.
	
		##### PROCESS.
		Process {
			Foreach ($Object in $InputObject) {
				# Set foreground and background colors.
				[System.Console]::ForegroundColor = $ForeGroundColor
				[System.Console]::BackgroundColor = $BackGroundColor
	
				# *Possibly* pad left and right.
				If ($HorizontalPad -gt 0) {
					$Object = $Object.PadLeft($Object.Length + $HorizontalPad)
					$Object = $Object.PadRight($Object.Length + $HorizontalPad)
				}
	
				# *Possibly* pad top.
				If ($VerticalPad -gt 0) {
					$BlankLine = ' ' * $Object.Length
					1..$VerticalPad | ForEach-Object {
						Microsoft.PowerShell.Utility\Write-Output -InputObject $BlankLine
					}
				}
	
				# Write message.
				If ($PSBoundParameters.ContainsKey('NoEnumerate')) {
					Microsoft.PowerShell.Utility\Write-Output -InputObject $Object -NoEnumerate
				} Else {
					Microsoft.PowerShell.Utility\Write-Output -InputObject $Object
				}
	
				# *Possibly* pad bottom.
				If ($VerticalPad -gt 0) {
					1..$VerticalPad | ForEach-Object {
						Microsoft.PowerShell.Utility\Write-Output -InputObject $BlankLine
					}
				}
			} # End Foreach
		} # End Process
	
		##### END.
		End {
			# Reset Colors with ResetColor Method.
			If ($ResetColorCheck) {
				[System.Console]::ResetColor()
			# Reset Colors *without* ResetColor Method.
			} Else {
				[System.Console]::ForegroundColor = $DefaultConsoleForegroundColor
				[System.Console]::BackgroundColor = $DefaultConsoleBackgroundColor
			}
		} # End End.
	} # End Function: Write-TMOutput.
	
Function Show-TMOutputColor {
<#
.SYNOPSIS
	This function will display the available colors that can be used with the Write-TMOutput function.
	
.DESCRIPTION
	This function will display the available colors that can be used with the Write-TMOutput function. Colors can be used for either the ForegroundColor or BackgroundColor parameter of Write-TMOutput.
	
.NOTES
	NAME: Show-TMOutputColor
	AUTHOR: Tommy Maynard
	LASTEDIT: 05/12/2016
	VERSION: 1.0
	Powershell Gallery: https://www.powershellgallery.com/packages/TMOutput/1.1
#>
	If (Get-Command -Name Write-TMOutput) {
		[System.Enum]::GetNames([System.ConsoleColor]) | ForEach-Object {
			$_
			Write-TMOutput -InputObject (' ' * 20) -BackgroundColor $_
		}
	}
} # End Function: Show-TMOutputColor

'@
#Load Classes and Functions
Invoke-Expression $Configuration.ClassADExchangeOutput

#Create Output Objects if they do not exist
If ($null -eq $Configuration.ADUsers -or $Configuration.ADUsers.count -eq 0) {
    $Configuration.ADUsers = @{}
}
If ($null -eq $Configuration.ADGroups -or $Configuration.ADGroups.count -eq 0) {
    $Configuration.ADGroups = [System.Collections.ArrayList]::new()
}
If ($null -eq $Configuration.ADFSEvents -or $Configuration.ADFSEvents.count -eq 0) {
    $Configuration.ADFSEvents = @{}
}

Write-TMOutput -InputObject ("Processing. Please Wait . . ." )
#region Get AD Users
    Write-TMOutput -InputObject "`tAD Users . . ."   -ForegroundColor DarkGray
    $swad = [Diagnostics.Stopwatch]::StartNew()
    If($Configuration.ADUsers.count -eq 0) {
        #Load .Net Assembly for AD
        Add-Type -AssemblyName System.DirectoryServices.AccountManagement
        $ContextType = [System.DirectoryServices.AccountManagement.ContextType]::Domain
        # Define the search criteria using UserPrincipal
        $userPrincipal = New-Object System.DirectoryServices.AccountManagement.UserPrincipal($ContextType)
        $userPrincipal.Name = "*"
        # Create a PrincipalSearcher to perform the search
        $searcher = New-Object System.DirectoryServices.AccountManagement.PrincipalSearcher($userPrincipal)
        # Get the underlying DirectorySearcher object
        $directorySearcher = $searcher.GetUnderlyingSearcher()
        $directorySearcher.PageSize = 1000  # Enable paged search
        #region Specify additional properties to load
        $directorySearcher.PropertiesToLoad.Add("Company") | Out-Null
        $directorySearcher.PropertiesToLoad.Add("co") | Out-Null
        $directorySearcher.PropertiesToLoad.Add("Department") | Out-Null
        $directorySearcher.PropertiesToLoad.Add("departmentnumber") | Out-Null
        $directorySearcher.PropertiesToLoad.Add("division") | Out-Null
        $directorySearcher.PropertiesToLoad.Add("employeeNumber") | Out-Null
        $directorySearcher.PropertiesToLoad.Add("employeetype") | Out-Null
        $directorySearcher.PropertiesToLoad.Add("extensionAttribute2") | Out-Null
        $directorySearcher.PropertiesToLoad.Add("homedirectory") | Out-Null
        $directorySearcher.PropertiesToLoad.Add("lastlogon") | Out-Null
        $directorySearcher.PropertiesToLoad.Add("logonCount") | Out-Null
        $directorySearcher.PropertiesToLoad.Add("mail") | Out-Null
        $directorySearcher.PropertiesToLoad.Add("manager") | Out-Null
        $directorySearcher.PropertiesToLoad.Add("MemberOf") | Out-Null
        $directorySearcher.PropertiesToLoad.Add("middleName") | Out-Null
        $directorySearcher.PropertiesToLoad.Add("mobile") | Out-Null
        $directorySearcher.PropertiesToLoad.Add("msExchArchiveGUID") | Out-Null
        $directorySearcher.PropertiesToLoad.Add("msExchMailboxGuid") | Out-Null
        $directorySearcher.PropertiesToLoad.Add("msExchWhenMailboxCreated") | Out-Null
        $directorySearcher.PropertiesToLoad.Add("msTSExpireDate") | Out-Null
        $directorySearcher.PropertiesToLoad.Add("passwordNeverExpires") | Out-Null
        $directorySearcher.PropertiesToLoad.Add("PasswordNotRequired") | Out-Null
        $directorySearcher.PropertiesToLoad.Add("physicalDeliveryOfficeName") | Out-Null
        $directorySearcher.PropertiesToLoad.Add("PostalCode") | Out-Null
        $directorySearcher.PropertiesToLoad.Add("proxyAddresses") | Out-Null
        $directorySearcher.PropertiesToLoad.Add("SmartcardLogonRequired") | Out-Null
        $directorySearcher.PropertiesToLoad.Add("st") | Out-Null
        $directorySearcher.PropertiesToLoad.Add("StreetAddress") | Out-Null
        $directorySearcher.PropertiesToLoad.Add("targetaddress") | Out-Null
        $directorySearcher.PropertiesToLoad.Add("telephoneNumber") | Out-Null
        $directorySearcher.PropertiesToLoad.Add("Title") | Out-Null
        $directorySearcher.PropertiesToLoad.Add("whencreated") | Out-Null
        $directorySearcher.PropertiesToLoad.Add("l") | Out-Null
        #endregion Specify additional properties to load
        # Perform the search and retrieve all results
        $results = $directorySearcher.FindAll()
        $Configuration.ADUsersCount = $results.Count
        # Process the results
        [int]$swccounter = 0
        If ($PSVersionTable.PSVersion.Major -ge 7) {
            Write-TMOutput -InputObject ("`t`tUsing Parallel Processing AD Records . . .") -ForegroundColor DarkGray
            $results | ForEach-Object -Parallel {
                $Configuration = $using:Configuration
                $swad = $using:swad
                $item = $_
                $Record = $null
                $GM=$null
                $CUPN = $null
                #Load Class and Functions
			    Invoke-Expression $Configuration.ClassADExchangeOutput   
                #Load .Net Assembly for AD
                Add-Type -AssemblyName System.DirectoryServices.AccountManagement
                $ContextType = [System.DirectoryServices.AccountManagement.ContextType]::Domain
                If($item.Properties.userprincipalname) {
                    $CUPN = $item.Properties.userprincipalname[0]
                    if([string]::IsNullOrWhiteSpace($CUPN)) {
                        $CUPN = $item.Properties.userprincipalname
                    }
                    if(-Not [string]::IsNullOrWhiteSpace($CUPN)) {
                        if($Configuration.ADUsers.count -eq 0) {
                            $Record = [ADExchangeOutput]::new()
                            $Record."User Principal Name" =  $CUPN
                        }ElseIf ($Configuration.ADUsers.ContainsKey($CUPN)) {
                            $Record = $Configuration.ADUsers[$CUPN]
                        }Else{
                            $Record = [ADExchangeOutput]::new()
                            $Record."User Principal Name" =  $CUPN
                        }
                        If($item.Properties.samaccountname){
                            $Record.'Logon Name' =  $item.Properties.samaccountname[0]
                        }
                        If($item.Properties.displayname){
                            $Record."Display Name" =  $item.Properties.displayname[0]
                        }
                        If($item.Properties.givenname){
                            $Record."First Name" =  $item.Properties.givenname[0]
                        }
                        If($item.Properties.middlename){
                            $Record."Middle Name" =  $item.Properties.middlename[0]
                        }
                        If($item.Properties.sn){
                            $Record."Last Name" =  $item.Properties.sn[0]
                        }
                        If($item.Properties.distinguishedname -and $Record."Distinguished Name" -eq $null) {
                            $Record."Distinguished Name" =  $item.Properties.distinguishedname[0]
                        }
                        #region retrieve the useraccountcontrol attribute
                            # 0x0002: ACCOUNTDISABLE - The account is disabled.
                            $ACCOUNTDISABLE = 0x0002
                            # 0x0020: PASSWD_NOTREQD - No password is required
                            $PASSWD_NOTREQD = 0x0020
                            # 0x0200: LOCKOUT - The account is locked out.
                            $LOCKOUT = 0x0200
                            # 0x10000: DONT_EXPIRE_PASSWORD - The password never expires.
                            $DONT_EXPIRE_PASSWORD = 0x10000
                            # 0x20000: SMARTCARD_REQUIRED - Smartcard is required for login.
                            $SMARTCARD_REQUIRED = 0x20000
                            # 0x40000: TRUSTED_FOR_DELEGATION - The account is trusted for delegation.
                            $TRUSTED_FOR_DELEGATION = 0x40000
                            $userAccountControl = [int]$item.Properties["useraccountcontrol"][0]
                            if ( $userAccountControl -band $ACCOUNTDISABLE) {
                                $Record."Account Status" = "Disabled"
                            }Else{
                                $Record."Account Status" = "Enabled"
                            }
                            if ($userAccountControl -band $DONT_EXPIRE_PASSWORD) {
                                $Record."Password Never Expires" = "Enabled"
                            }else{
                                $Record."Password Never Expires" = "Disabled"
                            }
                            if ($userAccountControl -band $SMARTCARD_REQUIRED) {
                                $Record."Smartcard Logon Required" = "Enabled"
                            }else{
                                $Record."Smartcard Logon Required" = "Disabled"
                            }
                            if ($userAccountControl -band $PASSWD_NOTREQD ) {
                                $Record."Password Not Required" = "Enabled"
                            }else{
                                $Record."Password Not Required" = "Disabled"
                            }
                            
                            if ($userAccountControl -band $TRUSTED_FOR_DELEGATION) {
                                $Record."Account Trusted for Delegation" = "Enabled"
                            }else{
                                $Record."Account Trusted for Delegation" = "Disabled"
                            }
                        #endregion retrieve the useraccountcontrol attribute
                        if ( $null -eq $item.Properties.pwdlastset[0]) {
                            $Record."Password Change on Next Logon" =  "True"
                        }Else{
                            $Record."Password Change on Next Logon" =  "False"
                        }
                        if ( $item.Properties.streetaddress) {
                            $Record."Full address" =  $item.Properties.streetaddress[0]
                        }
                        if ( $item.Properties.l) {
                            $Record."City" =  $item.Properties.l[0]
                        }
                        if ( $item.Properties.st) {
                            $Record."State" =  $item.Properties.st[0]
                        }
                        if ( $item.Properties.postalcode) {
                            $Record."Postal Code" =  $item.Properties.postalcode[0]
                        }
                        if ( $item.Properties.co) {
                            $Record."Country-Region" =  $item.Properties.co[0]
                        }
                        if ( $item.Properties.title) {
                            $Record."Job Title" =  $item.Properties.title[0]
                        }
                        if ( $item.Properties.company) {
                            $Record."Company" =  $item.Properties.company[0]
                        }
                        if ( $item.Properties.description) {
                            $Record."Description" =  $item.Properties.description -join ", "
                        }
                        if ( $item.Properties.department) {
                            $Record."Department" =  $item.Properties.department[0]
                        }
                        if ( $item.Properties.departmentnumber) {
                            $Record."Department Number" = ( $item.Properties.departmentnumber -join ", ")
                        }
                        if ( $item.Properties.employeetype) {
                            $Record."Employee Type" =  $item.Properties.employeetype[0]
                        }
                        if ( $item.Properties.employeenumber) {
                            $Record."Employee Number" =  $item.Properties.employeenumber[0]
                        }
                        if ( $item.Properties.physicaldeliveryofficename) {
                            $Record."Office" =  $item.Properties.physicaldeliveryofficename[0]
                        }
                        if ( $item.Properties.telephonenumber) {
                            $Record."Phone" =  $item.Properties.telephonenumber[0]
                        }
                        if ( $item.Properties.mobile) {
                            $Record."Mobile Phone" =  $item.Properties.mobile[0]
                        }
                        if ( $item.Properties.extensionattribute2) {
                            $Record."extensionAttribute2 CIFX Company" =  $item.Properties.extensionattribute2[0]
                        }
                        if ( $item.Properties.manager) {
                            #Search for the manager's display name from AD
                            $userManagerPrincipal = New-Object System.DirectoryServices.AccountManagement.UserPrincipal($ContextType)
                            $userManagerPrincipal.Name = $item.Properties.manager[0]
                            $Managersearcher = New-Object System.DirectoryServices.AccountManagement.PrincipalSearcher($userManagerPrincipal)
                            $userManagerPrincipalObject = $Managersearcher.FindOne()
                            $Record."Manager" = $userManagerPrincipalObject.DisplayName
                            # $userManagerPrincipalObject.Dispose()
                            $Managersearcher.Dispose()
                            $userManagerPrincipal.Dispose()
                            $userManagerPrincipalObject = $null
                            $Managersearcher = $null
                            $userManagerPrincipal = $null
                            #Use Manager Object Name from AD if not found in the cache or by searching
                            If ($null -eq $Record."Manager"){
                                $Record."Manager" = ( $item.Properties.manager[0] -split ",")[0] -replace "CN="
                            }
                        }
                        if ( $item.Properties.homedirectory) {
                            $Record."Home Directory" =  $item.Properties.homedirectory[0]
                        }
                        if ($item.Properties.whencreated) {
                            $Record."Creation Date" = [DateTime]$item.Properties.whencreated[0]
                            $Record."Days Since Creation" =(New-TimeSpan -Start ([datetime]$item.Properties.whencreated[0]) -End (Get-Date)).Days
                        }
                        if ($item.Properties.lastlogon) {
                            $Record."Last Log-On Date" =[DateTime]::FromFileTime($item.Properties.lastlogon[0])
                            $Record."Days Since Last Log-On" = (New-TimeSpan -Start ([DateTime]::FromFileTime($item.Properties.lastlogon[0])) -End (Get-Date)).Days
                        }
                        if ($item.Properties.pwdlastset) {
                            $Record."Last Password Change" = [DateTime]::FromFileTime($item.Properties.pwdlastset[0])
                            $Record."Days from last password change" = (New-TimeSpan -Start ([DateTime]::FromFileTime($item.Properties.pwdlastset[0])) -End (Get-Date)).Days
                        }
                        if ($item.Properties.mstsexpiredate) {
                            $Record."RDS CAL Expiration Date" = [DateTime]$item.Properties.mstsexpiredate[0]
                        }
                        $GM=( $item.Properties.memberof | ForEach-Object {($_ -split ",")[0].Replace("CN=","")}) -join ", "
                        if ($GM) {
                            $Record."Group Membership" = $GM
                        }
                        if ( $item.Properties.proxyaddresses) {
                            Try{
                                $pa = ($item.Properties.proxyaddresses | Where-Object {$_ -match "SMTP"}).Replace("SMTP:","") -join ", "
                                If($null -ne $pa -and $pa -ne "") {
                                    $Record."Email Addresses" = $pa
                                }
                            }Catch{
                            }
                        }
                        if ( $item.Properties.mail) {
                            $Record."Email" =  $item.Properties.mail[0]
                        }
                        if ($item.Properties.msexchwhenmailboxcreated) {
                            $Record."Mailbox Creation Date" = [DateTime]$item.Properties.msexchwhenmailboxcreated[0]
                        }
                        if ( $item.Properties.msexchmailboxguid) {
                            $Record."Mailbox GUID" =  [System.Guid]::New($item.Properties.msexchmailboxguid[0]).guid
                        }
                        if( $item.Properties.msexcharchiveguid) {
                            $Record."Mailbox Archive GUID" =  [System.Guid]::New($item.Properties.msexcharchiveguid[0]).guid
                        }
                        if ( $item.Properties.targetaddress){
                            $Record."Exchange Remote Routing Address" = $item.Properties.targetaddress[0]
                        }
                        #Save the record
                        $Configuration.ADUsers[$CUPN] = $Record
                        If (($Configuration.ADUsersCount % $Configuration.ADUsers.Count -eq 0) -or ($swad.Elapsed.TotalSeconds -gt ($swccounter +2)) ) {
                            write-progress -Id 0 -Activity "Getting AD Records" -Status ("Processing Record " + $Configuration.ADUsers.Count + " of " + $Configuration.ADUsersCount) -PercentComplete (($Configuration.ADUsers.Count / $Configuration.ADUsersCount) * 100)
                            $swccounter = [int]$swad.Elapsed.TotalSeconds
                        }
                        If ($Configuration.ADUsers.Count -eq $Configuration.ADUsersCount) {
                            write-progress -Id 0 -Activity "Getting AD Records" -Status ("Done") -PercentComplete 100
                        }
                    }
                }
            } -ThrottleLimit $Configuration.ProcessBatchJobs
        } Else {
            ForEach ($item in $results ){
                $Record = $null
                $GM=$null
                $CUPN = $null
                If($item.Properties.userprincipalname) {
                    $CUPN = $item.Properties.userprincipalname[0]
                    if([string]::IsNullOrWhiteSpace($CUPN)) {
                        $CUPN = $item.Properties.userprincipalname
                    }
                    if(-Not [string]::IsNullOrWhiteSpace($CUPN)) {
                        if($Configuration.ADUsers.count -eq 0) {
                            $Record = [ADExchangeOutput]::new()
                            $Record."User Principal Name" =  $CUPN
                        }ElseIf ($Configuration.ADUsers.ContainsKey($CUPN)) {
                            $Record = $Configuration.ADUsers[$CUPN]
                        }Else{
                            $Record = [ADExchangeOutput]::new()
                            $Record."User Principal Name" =  $CUPN
                        }
                        If($item.Properties.samaccountname){
                            $Record.'Logon Name' =  $item.Properties.samaccountname[0]
                        }
                        If($item.Properties.displayname){
                            $Record."Display Name" =  $item.Properties.displayname[0]
                        }
                        If($item.Properties.givenname){
                            $Record."First Name" =  $item.Properties.givenname[0]
                        }
                        If($item.Properties.middlename){
                            $Record."Middle Name" =  $item.Properties.middlename[0]
                        }
                        If($item.Properties.sn){
                            $Record."Last Name" =  $item.Properties.sn[0]
                        }
                        If($item.Properties.distinguishedname -and $null -eq $Record."Distinguished Name") {
                            $Record."Distinguished Name" =  $item.Properties.distinguishedname[0]
                        }
                        #region retrieve the useraccountcontrol attribute
                            # 0x0002: ACCOUNTDISABLE - The account is disabled.
                            $ACCOUNTDISABLE = 0x0002
                            # 0x0020: PASSWD_NOTREQD - No password is required
                            $PASSWD_NOTREQD = 0x0020
                            # 0x0200: LOCKOUT - The account is locked out.
                            $LOCKOUT = 0x0200
                            # 0x10000: DONT_EXPIRE_PASSWORD - The password never expires.
                            $DONT_EXPIRE_PASSWORD = 0x10000
                            # 0x20000: SMARTCARD_REQUIRED - Smartcard is required for login.
                            $SMARTCARD_REQUIRED = 0x20000
                            # 0x40000: TRUSTED_FOR_DELEGATION - The account is trusted for delegation.
                            $TRUSTED_FOR_DELEGATION = 0x40000
                            $userAccountControl = [int]$item.Properties["useraccountcontrol"][0]
                            if ( $userAccountControl -band $ACCOUNTDISABLE) {
                                $Record."Account Status" = "Disabled"
                            }Else{
                                $Record."Account Status" = "Enabled"
                            }
                            if ($userAccountControl -band $DONT_EXPIRE_PASSWORD) {
                                $Record."Password Never Expires" = "Enabled"
                            }else{
                                $Record."Password Never Expires" = "Disabled"
                            }
                            if ($userAccountControl -band $SMARTCARD_REQUIRED) {
                                $Record."Smartcard Logon Required" = "Enabled"
                            }else{
                                $Record."Smartcard Logon Required" = "Disabled"
                            }
                            if ($userAccountControl -band $PASSWD_NOTREQD ) {
                                $Record."Password Not Required" = "Enabled"
                            }else{
                                $Record."Password Not Required" = "Disabled"
                            }
                            
                            if ($userAccountControl -band $TRUSTED_FOR_DELEGATION) {
                                $Record."Account Trusted for Delegation" = "Enabled"
                            }else{
                                $Record."Account Trusted for Delegation" = "Disabled"
                            }
                        #endregion retrieve the useraccountcontrol attribute
                        if ( $null -eq $item.Properties.pwdlastset[0]) {
                            $Record."Password Change on Next Logon" =  "True"
                        }Else{
                            $Record."Password Change on Next Logon" =  "False"
                        }
                        if ( $item.Properties.streetaddress) {
                            $Record."Full address" =  $item.Properties.streetaddress[0]
                        }
                        if ( $item.Properties.l) {
                            $Record."City" =  $item.Properties.l[0]
                        }
                        if ( $item.Properties.st) {
                            $Record."State" =  $item.Properties.st[0]
                        }
                        if ( $item.Properties.postalcode) {
                            $Record."Postal Code" =  $item.Properties.postalcode[0]
                        }
                        if ( $item.Properties.co) {
                            $Record."Country-Region" =  $item.Properties.co[0]
                        }
                        if ( $item.Properties.title) {
                            $Record."Job Title" =  $item.Properties.title[0]
                        }
                        if ( $item.Properties.company) {
                            $Record."Company" =  $item.Properties.company[0]
                        }
                        if ( $item.Properties.description) {
                            $Record."Description" =  $item.Properties.description -join ", "
                        }
                        if ( $item.Properties.department) {
                            $Record."Department" =  $item.Properties.department[0]
                        }
                        if ( $item.Properties.departmentnumber) {
                            $Record."Department Number" = ( $item.Properties.departmentnumber -join ", ")
                        }
                        if ( $item.Properties.employeetype) {
                            $Record."Employee Type" =  $item.Properties.employeetype[0]
                        }
                        if ( $item.Properties.employeenumber) {
                            $Record."Employee Number" =  $item.Properties.employeenumber[0]
                        }
                        if ( $item.Properties.physicaldeliveryofficename) {
                            $Record."Office" =  $item.Properties.physicaldeliveryofficename[0]
                        }
                        if ( $item.Properties.telephonenumber) {
                            $Record."Phone" =  $item.Properties.telephonenumber[0]
                        }
                        if ( $item.Properties.mobile) {
                            $Record."Mobile Phone" =  $item.Properties.mobile[0]
                        }
                        if ( $item.Properties.extensionattribute2) {
                            $Record."extensionAttribute2 CIFX Company" =  $item.Properties.extensionattribute2[0]
                        }
                        if ( $item.Properties.manager) {
                            #Search for the manager's display name in AD Cache
                            $Record."Manager" = (($Configuration.ADUsers.Values).Where({$item."Distinguished Name" -eq $_.Manager})).DisplayName
                            #Search for the manager's display name from AD
                            If ($null -eq $Record."Manager"){
                                $userManagerPrincipal = New-Object System.DirectoryServices.AccountManagement.UserPrincipal($ContextType)
                                $userManagerPrincipal.Name = $item.Properties.manager[0]
                                $Managersearcher = New-Object System.DirectoryServices.AccountManagement.PrincipalSearcher($userManagerPrincipal)
                                $userManagerPrincipalObject = $Managersearcher.FindOne()
                                $Record."Manager" = $userManagerPrincipalObject.DisplayName
                                # $userManagerPrincipalObject.Dispose()
                                $Managersearcher.Dispose()
                                $userManagerPrincipal.Dispose()
                                $userManagerPrincipalObject = $null
                                $Managersearcher = $null
                                $userManagerPrincipal = $null
                            }
                            #Use Manager Object Name from AD if not found in the cache or by searching
                            If ($null -eq $Record."Manager"){
                                $Record."Manager" = ( $item.Properties.manager[0] -split ",")[0] -replace "CN="
                            }
                        }
                        if ( $item.Properties.homedirectory) {
                            $Record."Home Directory" =  $item.Properties.homedirectory[0]
                        }
                        if ($item.Properties.whencreated) {
                            $Record."Creation Date" = [DateTime]$item.Properties.whencreated[0]
                            $Record."Days Since Creation" =(New-TimeSpan -Start ([datetime]$item.Properties.whencreated[0]) -End (Get-Date)).Days
                        }
                        if ($item.Properties.lastlogon) {
                            $Record."Last Log-On Date" = [DateTime]$item.Properties.lastlogon[0]
                            $Record."Days Since Last Log-On" = (New-TimeSpan -Start ([datetime]$item.Properties.lastlogon[0]) -End (Get-Date)).Days
                        }
                        if ($item.Properties.pwdlastset) {
                            $Record."Last Password Change" = [DateTime]$item.Properties.pwdlastset[0]
                            $Record."Days from last password change" = (New-TimeSpan -Start ([datetime]$item.Properties.pwdlastset[0]) -End (Get-Date)).Days
                        }
                        if ($item.Properties.mstsexpiredate) {
                            $Record."RDS CAL Expiration Date" = [DateTime]$item.Properties.mstsexpiredate[0]
                        }
                        $GM=( $item.Properties.memberof | ForEach-Object {($_ -split ",")[0].Replace("CN=","")}) -join ", "
                        if ($GM) {
                            $Record."Group Membership" = $GM
                        }
                        if ( $item.Properties.proxyaddresses) {
                            Try{
                                $pa = ($item.Properties.proxyaddresses | Where-Object {$_ -match "SMTP"}).Replace("SMTP:","") -join ", "
                                If($null -ne $pa -and $pa -ne "") {
                                    $Record."Email Addresses" = $pa
                                }
                            }Catch{
                            }
                        }
                        if ( $item.Properties.mail) {
                            $Record."Email" =  $item.Properties.mail[0]
                        }
                        if ($item.Properties.msexchwhenmailboxcreated) {
                            $Record."Mailbox Creation Date" = [DateTime]$item.Properties.msexchwhenmailboxcreated[0]
                        }
                        if ( $item.Properties.msexchmailboxguid) {
                            $Record."Mailbox GUID" =  [System.Guid]::New($item.Properties.msexchmailboxguid[0]).guid
                        }
                        if( $item.Properties.msexcharchiveguid) {
                            $Record."Mailbox Archive GUID" =  [System.Guid]::New($item.Properties.msexcharchiveguid[0]).guid
                        }
                        if ( $item.Properties.targetaddress){
                            $Record."Exchange Remote Routing Address" = $item.Properties.targetaddress[0]
                        }                      
                        #Save the record
                        $Configuration.ADUsers[$CUPN] = $Record
                        If (($Configuration.ADUsersCount % $Configuration.ADUsers.Count-eq 0) -or ($swad.Elapsed.TotalSeconds -gt ($swccounter +2))) {
                            write-progress -Id 0 -Activity "Getting AD Records" -Status ("Processing Record " + $Configuration.ADUsers.Count + " of " + $Configuration.ADUsersCount) -PercentComplete (($Configuration.ADUsers.Count / $Configuration.ADUsersCount) * 100)
                            $swccounter = [int]$swad.Elapsed.TotalSeconds
                        }
                    }
                }
            }
        }
        $Configuration.ADUsersUPNs = $Configuration.ADUsers.Keys
        # $Configuration.ADUsersCount = $Configuration.ADUsers.Count
        $Configuration.ADUsersUPNsWithEmails = ($Configuration.ADUsers.Values | Where-Object {$null -ne $_."Email"}).'User Principal Name'
        $Configuration.ADUsersUPNsWithERRA = ($Configuration.ADUsers.Values | Where-Object {$null -ne $_.'Exchange Remote Routing Address'}).'User Principal Name'
        $Configuration.ADUsersWithMailboxGUIDs = ($Configuration.ADUsers.Values | Where-Object {$null -ne $_."Mailbox GUID" -or $null -ne $_."Mailbox Archive GUID"}).'User Principal Name'
        $Configuration.ADUsersEnabled = ($Configuration.ADUsers.Values | Where-Object {$_."Account Status" -eq "Enabled"}).'User Principal Name'

        $userPrincipal.Dispose()
        $searcher.Dispose()
        $directorySearcher.Dispose()
        $results.Dispose()
        $Record = $null
        $GM = $null
        $CUPN = $null
        
        $swad.Stop()
        If($Configuration.ADUsers.Count -gt 0 -and $swad.Elapsed.TotalMinutes -gt 0) {
            Write-TMOutput -InputObject ("`tDone. Getting AD Records: " + (Format-ElapsedTime($swad.Elapsed)) + " to run. " + '{0:N0}' -f ($Configuration.ADUsers.Count / $swad.Elapsed.TotalMinutes) + " Users's per Minute.") -ForegroundColor DarkGray
        }
        write-progress -Id 0 -Activity "Getting AD Records" -Status "Done" -PercentComplete 100
    }
#endregion Get AD Users
#region RunSpace Setup
    $runspacepool = [runspacefactory]::CreateRunspacePool()
    $runspacepool.SetMinRunspaces(1) | Out-Null
    $runspacepool.SetMaxRunspaces($Configuration.ProcessBatchJobs) | Out-Null
    # $runspacepool.ThreadOptions = "ReuseThread"
    $runspacepool.Open() | Out-Null
#endregion RunSpace Setup
#region RunSpace Jobs
    $swrs = [Diagnostics.Stopwatch]::StartNew()
	#region get Groups
	# Write-TMOutput -InputObject "`tAD Groups . . ."  -ForegroundColor DarkGray
	$GroupsScript = {
        param(
            $Configuration
        )
        #Import Class ADGroupOutput
        Invoke-Expression $Configuration.ClassADExchangeOutput
        #Load .Net Assembly for AD
        Add-Type -AssemblyName System.DirectoryServices.AccountManagement
        $ContextType = [System.DirectoryServices.AccountManagement.ContextType]::Domain
		# Define a GroupPrincipal to specify the search criteria
        $DomainContext = New-Object System.DirectoryServices.AccountManagement.PrincipalContext($ContextType, $ENV:USERDNSDOMAIN)
        $groupPrincipal = New-Object System.DirectoryServices.AccountManagement.GroupPrincipal($DomainContext, "*")
		# Create a PrincipalSearcher to perform the search
		$searcher = New-Object System.DirectoryServices.AccountManagement.PrincipalSearcher($groupPrincipal)
        #Get all results
		$searchResultCollection = $searcher.FindAll()

        $directorySearcher = $searcher.GetUnderlyingSearcher()
        $directorySearcher.Filter = "(&(objectClass=group)(objectCategory=group)(mail=*))"
        $directorySearcher.PropertiesToLoad.Add("mail") | Out-Null
        # $directorySearcher.PropertiesToLoad.Add("name") | Out-Null
        $directorySearcher.PropertiesToLoad.Add("sAMAccountName") | Out-Null
        # $directorySearcher.PropertiesToLoad.Add("description") | Out-Null
        # $directorySearcher.PropertiesToLoad.Add("objectSid") | Out-Null
        # $directorySearcher.PropertiesToLoad.Add("distinguishedName") | Out-Null
        # $directorySearcher.PropertiesToLoad.Add("member") | Out-Null
        $directorySearcher.PageSize = 1000  # Enable paged search

        $searchResultCollectionEmail=@{}
        ForEach ($GroupEmail in $directorySearcher.FindAll()){
            $searchResultCollectionEmail.Add($GroupEmail.Properties.samaccountname[0],$GroupEmail.Properties.mail[0])
        }

        #Get the count of groups
        $Configuration.ADGroupsCount = ($searchResultCollection | Select-Object name).count
        # Loop through each group and collect the required information
		foreach ($Group in $searchResultCollection) {
            $Record = [ADGroupOutputClass]::new()
			$Record.Name = $Group.Name
			$Record.sAMAccountName = $Group.sAMAccountName
            if(-Not([string]::IsNullOrWhiteSpace($Group.Description))) {
			    $Record.description = $Group.Description
            }
            $Record.'Security Group' = $Group.IsSecurityGroup
			Try{
				If ($Group.Name -in $AD_Default_Groups -or $Group.Name -in $Exchange_Security_Groups) {
					$Record.type = "System"
				}
			} Catch {
			}
            If(-Not([string]::IsNullOrWhiteSpace($Group.GroupScope))) {
                $Record.GroupScope = $Group.GroupScope
            }
            if(-Not([string]::IsNullOrWhiteSpace($searchResultCollectionEmail[$Group.sAMAccountName]))) {
                $Record.mail = $searchResultCollectionEmail[$Group.sAMAccountName]
            }
			$Record.objectSid = $Group.Sid
			$Record.distinguishedName = $Group.distinguishedName
			$LocalMembers = $Group.Members.SamAccountName
            if($LocalMembers.count -gt 0) {
                $Record.'Member Count' =  $LocalMembers.count
                $Record.'Members' =  $LocalMembers -join ","
            }
            #Add the record to the collection
            $Configuration.ADGroups.Add($Record) | Out-Null
            #Update the progress bar
            If ($null -ne $Configuration.Jobs["AD Groups"]) {
                $Configuration.Jobs["AD Groups"].process.PercentComplete = (($Configuration.ADGroups.Count/$Configuration.ADGroupsCount)*100)
                $Configuration.Jobs["AD Groups"].process.Status = "Processing Record " + $Configuration.ADGroups.Count + " of " + $Configuration.ADGroupsCount
            }              
		}
		$searchResultCollection.Dispose()
		$searcher.Dispose()
        $groupPrincipal.Dispose()
        $Configuration.Jobs["AD Groups"].process.Completed = $true
        $Configuration.Jobs["AD Groups"].process.Status = "Completed"
        $Configuration.Jobs["AD Groups"].process.PercentComplete = 100
	}
    $powershell = [powershell]::Create()
    $powershell.RunspacePool = $runspacepool
    $powershell.AddScript($GroupsScript) | Out-Null
    $powershell.AddParameter('Configuration',$Configuration) | Out-Null
    $Configuration.Jobs["AD Groups"] =[PSCustomObject]@{ Pipe = $powershell; Status = $powershell.BeginInvoke(); Name= "AD Groups" ;Start = Get-Date ; End = ""; process  = [PSCustomObject]@{ Id= ($Configuration.Jobs.Count + 2); Activity = "Processing AD Groups"; Status = "Starting"; PercentComplete = 0;Completed =$false} }
	#endregion get Groups
    #region ADFS Event logs
        If ($Configuration.ADFSServers.Count -gt 0) {
            $ADFSScript = {
                param(
                    $ADFS,
                    $Configuration
                )
                #Set the XPath query to filter events with specific IDs
                #Update the progress bar
                If ( $null -ne $Configuration.Jobs["ADFS Events on $ADFS"]) {
                    $Configuration.Jobs["ADFS Events on $ADFS"].process.PercentComplete = 0
                    $Configuration.Jobs["ADFS Events on $ADFS"].process.Status = "Gathering ADFS Events on $ADFS"
                }
                $XPath = "*[System[Provider[@Name='AD FS Auditing']]]" 
                $Events= Get-WinEvent -LogName 'Security' -FilterXPath $XPath -ComputerName $ADFS | Where-Object {$_.id -in @(1200,1203)} | Sort-Object TimeCreated -Descending
                $EventCount = $Events.Count
                $ProgressCounter = 1
                $UserIDs = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
                If ( $null -ne $Configuration.Jobs["ADFS Events on $ADFS"]) {
                    $Configuration.Jobs["ADFS Events on $ADFS"].process.PercentComplete = 0
                    $Configuration.Jobs["ADFS Events on $ADFS"].process.Status = "Processing ADFS Events on $ADFS"
                }               
                If ($PSVersionTable.PSVersion.Major -ge 7) {
                    $Events | ForEach-Object -Parallel {
                        $e = $_
                        $Configuration = $using:Configuration
                        $ADFS = $using:ADFS
                        $EventCount = $using:EventCount
                        $ProgressCounter = $using:ProgressCounter
                        $UserIDs = $using:UserIDs
                        $Record = $null
                        $UserID = ""
                        Invoke-Expression $Configuration.ClassADExchangeOutput
                        #Load .Net Assembly for AD
                        Add-Type -AssemblyName System.DirectoryServices.AccountManagement
                        $ContextType = [System.DirectoryServices.AccountManagement.ContextType]::Domain
                        $IdentityType = [System.DirectoryServices.AccountManagement.IdentityType]::SamAccountName
                        $exml = [xml]$e.Message.Substring($e.Message.IndexOf("XML: ")+5)
                        #region Get User Name
                        If ($exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "ResourceAuditComponent"}).UserID -match '@') {
                            $UserID = ($exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "ResourceAuditComponent"}).UserID -split "@")[0]
                        }ElseIf ($exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "ResourceAuditComponent"}).UserID -match '\\'){
                            $UserID = ($exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "ResourceAuditComponent"}).UserID -split '\\')[1]
                        } 
                        # Test if $UserID is a number
                        if ($UserID -match '^\d+$') {
                            # Pad $UserID to be 7 characters long with leading zeros
                            $UserID = $UserID.PadLeft(7, '0')
                            # Write-Output "Padded UserID: $UserID"
                        }
                        #endregion Get User Name
                        If(-Not [string]::IsNullOrEmpty($e.TimeCreated) -and $UserID -and $UserIDs -notcontains $UserID) {
                            Try{
                                If ([datetime]$e.TimeCreated -is [datetime]) {
                                    #Get the User Principal Name
                                    $DomainObject = [DirectoryServices.AccountManagement.Principal]::FindByIdentity($ContextType,$IdentityType,$UserID)
                                    If ($DomainObject.UserPrincipalName) {
                                        $CurrentUsername = $DomainObject.UserPrincipalName
                                        If ($Configuration.ADUsers.ContainsKey($CurrentUsername)) {
                                            If([String]::IsNullOrWhiteSpace($Configuration.ADUsers[$CurrentUsername]."ADFS Last Logon") -or [DateTime]$e.TimeCreated -gt [DateTime]$Configuration.ADUsers[$CurrentUsername]."ADFS Last Logon") {
                                                $Configuration.ADUsers[$CurrentUsername]."ADFS Last Logon" = $e.TimeCreated
                                                $Configuration.ADUsers[$CurrentUsername]."ADFS Last Logon Days" = ($((Get-Date)- ([DateTime]$e.TimeCreated)).Days)
                                                If ($exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).IpAddress) {
                                                    $Configuration.ADUsers[$CurrentUsername]."ADFS Last Logon IP" = $exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).IpAddress
                                                }
                                                If($exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "ResourceAuditComponent"}).RelyingParty) {
                                                    $Configuration.ADUsers[$CurrentUsername]."ADFS Relying Party" = $exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "ResourceAuditComponent"}).RelyingParty
                                                }
                                                If($exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).AuthProtocol) {
                                                    $Configuration.ADUsers[$CurrentUsername]."ADFS Auth Protocol" = $exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).AuthProtocol
                                                }
                                                If($exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).NetworkLocation) {
                                                    $Configuration.ADUsers[$CurrentUsername]."ADFS Network Location" = $exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).NetworkLocation
                                                }
                                                If($exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).UserAgentString) {
                                                    $Configuration.ADUsers[$CurrentUsername]."ADFS User Agent String" = $exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).UserAgentString
                                                }
                                                $UserIDs.Add($UserID) | Out-Null
                                                # Write-TMOutput -InputObject ("`tDirect ADUser hashtable insert. " + [string]$UserID + " Last Logon: " + $e.TimeCreated) -ForegroundColor DarkGray
                                            }
                                        }
                                    }Else {
                                        #If the User Principal Name is not found, use the User ID
                                        $Record = [ADFSEventRecord]::new()   
                                        $Record."Date-Time" = $e.TimeCreated
                                        $Record."ADFS Server" = $ADFS
                                        $Record."User ID" =  $UserID
                                        If($exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "ResourceAuditComponent"}).RelyingParty) {
                                            $Record."Relying Party" = $exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "ResourceAuditComponent"}).RelyingParty
                                        }
                                        If($exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).AuthProtocol) {
                                            $Record."Auth Protocol" = $exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).AuthProtocol
                                        }
                                        If($exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).NetworkLocation) {
                                            $Record."Network Location" = $exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).NetworkLocation
                                        }
                                        If($exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).IpAddress) {
                                            $Record."IP Address" = $exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).IpAddress
                                        }
                                        If($exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).UserAgentString) {
                                            $Record."User Agent String" = $exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).UserAgentString
                                        }
                                        $Record."Comments" = "Non-Matching AD Username"
                                        $Configuration.ADFSEvents[$UserID] = $Record
                                        # Write-TMOutput -InputObject ("`tNon-Matching AD Username: " + [string]$UserID + " Last Logon: " + $e.TimeCreated) -ForegroundColor DarkGray
                                    }
                                }Else {
                                    #If the User Principal Name is not found, use the User ID
                                    $Record = [ADFSEventRecord]::new()
                                    $Record."ADFS Server" = $ADFS
                                    $Record."User ID" =  $UserID
                                    If ($e.TimeCreated) {
                                        $Record."Date-Time" = $e.TimeCreated
                                    }
                                    If($exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "ResourceAuditComponent"}).RelyingParty) {
                                        $Record."Relying Party" = $exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "ResourceAuditComponent"}).RelyingParty
                                    }
                                    If($exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).AuthProtocol) {
                                        $Record."Auth Protocol" = $exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).AuthProtocol
                                    }
                                    If($exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).NetworkLocation) {
                                        $Record."Network Location" = $exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).NetworkLocation
                                    }
                                    If($exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).IpAddress) {
                                        $Record."IP Address" = $exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).IpAddress
                                    }
                                    If($exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).UserAgentString) {
                                        $Record."User Agent String" = $exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).UserAgentString
                                    }
                                    $Record."Comments" = "Bad Date Time"
                                    $Configuration.ADFSEvents[$UserID] = $Record
                                    # Write-TMOutput -InputObject ("`tNon-Matching AD Username: " + [string]$UserID + " Last Logon: " + $e.TimeCreated) -ForegroundColor DarkGray
                                }
                            }Catch{
                                Write-TMOutput -InputObject ("Error Processing ADFS Event: " + $_.Exception.Message) -ForegroundColor Red
                            }
                        }
                        #Update the progress bar
                        If ($EventCount % $ProgressCounter -eq 0){
                            $Configuration.Jobs["ADFS Events on $ADFS"].process.PercentComplete = (($ProgressCounter/$EventCount)*100)
                            $Configuration.Jobs["ADFS Events on $ADFS"].process.Status = "Processing Record " + $ProgressCounter + " of " + $EventCount
                        }
                        # Write-Progress -Id 0 -Activity "Processing ADFS Events on $ADFS" -Status ("Processing Record " + $ProgressCounter + " of " + $EventCount) -PercentComplete (($ProgressCounter/$EventCount)*100)
                        $ProgressCounter++
                    } -ThrottleLimit $Configuration.ProcessBatchJobs
                }else{
                    #Import Class ADGroupOutput
                    Invoke-Expression $Configuration.ClassADExchangeOutput
                    #Load .Net Assembly for AD
                    Add-Type -AssemblyName System.DirectoryServices.AccountManagement
                    $ContextType = [System.DirectoryServices.AccountManagement.ContextType]::Domain
                    $IdentityType = [System.DirectoryServices.AccountManagement.IdentityType]::SamAccountName
                    ForEach ($e in $Events) {
                        $UserID = ""
                        $exml = [xml]$e.Message.Substring($e.Message.IndexOf("XML: ")+5)
                        #region Get User Name
                        If ($exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "ResourceAuditComponent"}).UserID -match '@') {
                            $UserID = ($exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "ResourceAuditComponent"}).UserID -split "@")[0]
                        }ElseIf ($exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "ResourceAuditComponent"}).UserID -match '\\'){
                            $UserID = ($exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "ResourceAuditComponent"}).UserID -split '\\')[1]
                        }  
                        # Test if $UserID is a number
                        if ($UserID -match '^\d+$') {
                            # Pad $UserID to be 7 characters long with leading zeros
                            $UserID = $UserID.PadLeft(7, '0')
                            # Write-Output "Padded UserID: $UserID"
                        }
                        #endregion Get User Name
                        If(-Not [string]::IsNullOrEmpty($e.TimeCreated) -and $UserID -and $UserIDs -notcontains $UserID) {
                            Try{
                                If ([datetime]$e.TimeCreated -is [datetime]) {
                                    #Get the User Principal Name
                                    $DomainObject = [DirectoryServices.AccountManagement.Principal]::FindByIdentity($ContextType,$IdentityType,$UserID)
                                    If ($DomainObject.UserPrincipalName) {
                                        $CurrentUsername = $DomainObject.UserPrincipalName
                                        If ($Configuration.ADUsers.ContainsKey($CurrentUsername)) {
                                            If([String]::IsNullOrWhiteSpace($Configuration.ADUsers[$CurrentUsername]."ADFS Last Logon") -or [DateTime]$e.TimeCreated -gt [DateTime]$Configuration.ADUsers[$CurrentUsername]."ADFS Last Logon") {
                                                $Configuration.ADUsers[$CurrentUsername]."ADFS Last Logon" = $e.TimeCreated
                                                $Configuration.ADUsers[$CurrentUsername]."ADFS Last Logon Days" = ($((Get-Date)- ([DateTime]$e.TimeCreated)).Days)
                                                If ($exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).IpAddress) {
                                                    $Configuration.ADUsers[$CurrentUsername]."ADFS Last Logon IP" = $exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).IpAddress
                                                }
                                                If($exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "ResourceAuditComponent"}).RelyingParty) {
                                                    $Configuration.ADUsers[$CurrentUsername]."ADFS Relying Party" = $exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "ResourceAuditComponent"}).RelyingParty
                                                }
                                                If($exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).AuthProtocol) {
                                                    $Configuration.ADUsers[$CurrentUsername]."ADFS Auth Protocol" = $exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).AuthProtocol
                                                }
                                                If($exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).NetworkLocation) {
                                                    $Configuration.ADUsers[$CurrentUsername]."ADFS Network Location" = $exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).NetworkLocation
                                                }
                                                If($exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).UserAgentString) {
                                                    $Configuration.ADUsers[$CurrentUsername]."ADFS User Agent String" = $exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).UserAgentString
                                                }
                                                $UserIDs.Add($UserID) | Out-Null
                                                # Write-TMOutput -InputObject ("`tDirect ADUser hashtable insert. " + [string]$UserID + " Last Logon: " + $e.TimeCreated) -ForegroundColor DarkGray
                                            }
                                        }
                                    }Else {
                                        #If the User Principal Name is not found, use the User ID
                                        $Record = [ADFSEventRecord]::new()   
                                        $Record."Date-Time" = $e.TimeCreated
                                        $Record."ADFS Server" = $ADFS
                                        $Record."User ID" =  $UserID
                                        If($exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "ResourceAuditComponent"}).RelyingParty) {
                                            $Record."Relying Party" = $exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "ResourceAuditComponent"}).RelyingParty
                                        }
                                        If($exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).AuthProtocol) {
                                            $Record."Auth Protocol" = $exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).AuthProtocol
                                        }
                                        If($exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).NetworkLocation) {
                                            $Record."Network Location" = $exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).NetworkLocation
                                        }
                                        If($exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).IpAddress) {
                                            $Record."IP Address" = $exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).IpAddress
                                        }
                                        If($exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).UserAgentString) {
                                            $Record."User Agent String" = $exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).UserAgentString
                                        }
                                        $Record."Comments" = "Non-Matching AD Username"
                                        $Configuration.ADFSEvents[$UserID] = $Record
                                        # Write-TMOutput -InputObject ("`tNon-Matching AD Username: " + [string]$UserID + " Last Logon: " + $e.TimeCreated) -ForegroundColor DarkGray
                                    }
                                }Else {
                                    #If the User Principal Name is not found, use the User ID
                                    $Record = [ADFSEventRecord]::new()
                                    $Record."ADFS Server" = $ADFS
                                    $Record."User ID" =  $UserID
                                    If ($e.TimeCreated) {
                                        $Record."Date-Time" = $e.TimeCreated
                                    }
                                    If($exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "ResourceAuditComponent"}).RelyingParty) {
                                        $Record."Relying Party" = $exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "ResourceAuditComponent"}).RelyingParty
                                    }
                                    If($exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).AuthProtocol) {
                                        $Record."Auth Protocol" = $exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).AuthProtocol
                                    }
                                    If($exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).NetworkLocation) {
                                        $Record."Network Location" = $exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).NetworkLocation
                                    }
                                    If($exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).IpAddress) {
                                        $Record."IP Address" = $exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).IpAddress
                                    }
                                    If($exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).UserAgentString) {
                                        $Record."User Agent String" = $exml.AuditBase.ContextComponents.Component.Where({$_."type" -eq "RequestAuditComponent"}).UserAgentString
                                    }
                                    $Record."Comments" = "Bad Date Time"
                                    $Configuration.ADFSEvents[$UserID] = $Record
                                    # Write-TMOutput -InputObject ("`tBad Date Time: " + [string]$UserID + " Last Logon: " + $e.TimeCreated) -ForegroundColor DarkGray
                                }
                            }Catch{
                                Write-TMOutput -InputObject ("Error Processing ADFS Event: " + $_.Exception.Message) -ForegroundColor Red
                            }
                        }
                        #Update the progress bar
                        If ($EventCount % $ProgressCounter -eq 0){
                            $Configuration.Jobs["ADFS Events on $ADFS"].process.PercentComplete = (($ProgressCounter/$EventCount)*100)
                            $Configuration.Jobs["ADFS Events on $ADFS"].process.Status = "Processing Record " + $ProgressCounter + " of " + $EventCount
                        }
                        # Write-Progress -Id 0 -Activity "Processing ADFS Events on $ADFS" -Status ("Processing Record " + $ProgressCounter + " of " + $EventCount) -PercentComplete (($ProgressCounter/$EventCount)*100)
                        $ProgressCounter++
                    }
                }
                
                $Events.Dispose()
                $Record = $null
                $exml = $null
                $UserID = $null
                $DomainObject = $null
                # Triggering garbage collection
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
                $Configuration.Jobs["ADFS Events on $ADFS"].process.Completed = $true
                $Configuration.Jobs["ADFS Events on $ADFS"].process.Status = "Completed"
                $Configuration.Jobs["ADFS Events on $ADFS"].process.PercentComplete = 100
                # $Configuration.Jobs["ADFS Events on $ADFS"].End = Get-Date
            }
            Foreach ($ADFS in $Configuration.ADFSServers) {
                $powershell = [powershell]::Create()
                $powershell.RunspacePool = $runspacepool
                $powershell.AddScript($ADFSScript) | Out-Null
                $powershell.AddParameter('ADFS',$ADFS) | Out-Null
                $powershell.AddParameter('Configuration',$Configuration) | Out-Null
                $Configuration.Jobs["ADFS Events on $ADFS"] = [PSCustomObject]@{ Pipe = $powershell; Status = $powershell.BeginInvoke(); Name= "ADFS Events on $ADFS" ;Start = Get-Date ; End = ""; process = [PSCustomObject]@{ Id= ($Configuration.Jobs.Count + 2); Activity = "Processing ADFS Events on $ADFS" ; Status = "Starting"; PercentComplete = 0;Completed =$false} }
            }
        }
    #endregion ADFS Event logs
	#region Entra ID Users
    # Write-TMOutput -InputObject "`tEntra ID Users . . ."   -ForegroundColor DarkGray
    $MgUsers = {
        param(
            $Configuration
        )
        #Import Class ADExchangeOutput
        Invoke-Expression $Configuration.ClassADExchangeOutput
        #region Graph API Setup and Connection
        If($PSVersionTable.PSVersion.Major -eq 5){
            # Increase the Function Count
            $Global:MaximumFunctionCount = 8192
            # Increase the Variable Count
            $Global:MaximumVariableCount = 8192
        }
        If (Get-Module -ListAvailable -Name "Microsoft.Graph") {
            If (-Not (Get-Module "Microsoft.Graph" -ErrorAction SilentlyContinue)) {
                #Import-Module "Microsoft.Graph" 
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
        If($null -ne $Configuration.AZCertThumbprint -and $Configuration.AZCertThumbprint.Length -eq 40) {
            If (Get-ChildItem "Cert:\LocalMachine\My\" | Where-Object {$_.Thumbprint -eq $Configuration.AZCertThumbprint}) {
                $ClientCertificate = Get-Item "Cert:\LocalMachine\My\$($Configuration.AZCertThumbprint)"
                # $myAccessToken = Get-MsalToken -ClientId $AZClientID -TenantId $AzureTenant -ClientCertificate $ClientCertificate
                # $AccessToken  = $myAccessToken.AccessToken
            }
        }
        #endregion Get Access Token

        #Connect to MgGraph
        If ($null -ne $Configuration.AzureTenant) {
            If($null -ne $ClientCertificate ) {
                Connect-MgGraph -ClientId $Configuration.AZClientID -TenantId $Configuration.AzureTenant -Certificate $ClientCertificate -NoWelcome 
            }ElseIf($null -ne $Configuration.AZCertThumbprint -and $Configuration.AZCertThumbprint.Length -eq 40) {
                Connect-MgGraph -ClientId $Configuration.AZClientID -TenantId $Configuration.AzureTenant -CertificateThumbprint $Configuration.AZCertThumbprint -NoWelcome 
            }Else {
                Connect-MgGraph -NoWelcome -TenantId $Configuration.AzureTenant -Scopes  "User.ReadBasic.All", "UserAuthenticationMethod.Read.All", "IdentityUserFlow.Read.All", "User.EnableDisableAccount.All", "User.EnableDisableAccount.All", "IdentityRiskyUser.Read.All", "Directory.Read.All", "AuditLog.Read.All"
            }
        #endregion Graph API Setup and Connection
            $Configuration.MgSubscribedSku  = Get-MgSubscribedSku -All
            $Users = Get-MgUser -all -Property UserPrincipalName,ID,OnPremisesImmutableId,OnPremisesLastSyncDateTime,OnPremisesSamAccountName,OnPremisesUserPrincipalName,onPremisesSecurityIdentifier,assignedLicenses,assignedPlans,UserPrincipalName,AccountEnabled,UserType,SignInActivity
            $ProgressCounter = 1
            ForEach ( $User in $Users ) {
                $Licenses = @()
                $LicenseDetails = @()
                If(-Not([string]::IsNullOrWhiteSpace($User.UserPrincipalName))) {
                    If($null -eq $Configuration.ADUsers[$User.UserPrincipalName]."User Principal Name") {
                        $Configuration.ADUsers[$User.UserPrincipalName] = [ADExchangeOutput]::new()
                        $Configuration.ADUsers[$User.UserPrincipalName]."User Principal Name" =  $User.UserPrincipalName
                    }
                    if(!([string]::IsNullOrWhiteSpace($User.OnPremisesLastSyncDateTime))) {
                        $Configuration.ADUsers[$User.UserPrincipalName]."Azure Last Sync Time" = [DateTime]$User.OnPremisesLastSyncDateTime
                    }
                    if(!([string]::IsNullOrWhiteSpace($User.signInActivity.lastSignInDateTime))) {
                        $Configuration.ADUsers[$User.UserPrincipalName]."Azure Last Sign-On" = [DateTime]$User.signInActivity.LastSignInDateTime
                        $Configuration.ADUsers[$User.UserPrincipalName]."Azure Last Sign-On Days" = ($((Get-Date) - ([DateTime]$User.signInActivity.lastSignInDateTime)).Days)
                    }
                    if(!([string]::IsNullOrWhiteSpace($User.signInActivity.lastNonInteractiveSignInDateTime))) {
                        $Configuration.ADUsers[$User.UserPrincipalName]."Azure Last Non-Interactive Sign-On" = [DateTime]$User.signInActivity.lastNonInteractiveSignInDateTime 
                        $Configuration.ADUsers[$User.UserPrincipalName]."Azure Last Non-Interactive Sign-On Days" = ($((Get-Date) - ([DateTime]$User.signInActivity.lastNonInteractiveSignInDateTime)).Days)
                    }
                    if(!([string]::IsNullOrWhiteSpace($User.onPremisesSecurityIdentifier))) {
                        $Configuration.ADUsers[$User.UserPrincipalName]."Azure Security Identifier" = [string]$User.onPremisesSecurityIdentifier
                    }
                    if(!([string]::IsNullOrWhiteSpace($User.onPremisesImmutableId))) {
                        $Configuration.ADUsers[$User.UserPrincipalName]."Azure ImmutableId" = $User.onPremisesImmutableId
                    }
                    if(!([string]::IsNullOrWhiteSpace($User.AccountEnabled))) {
                        $Configuration.ADUsers[$User.UserPrincipalName]."Azure Account Enabled" = $User.AccountEnabled
                    }
                    ForEach ($License in $User.AssignedLicenses) {
                        $Licenses += ($Configuration.MgSubscribedSku.Where({$_.SkuId -eq $License.SkuId}).SkuPartNumber)
                    }
                    if($Licenses.count -gt 0) {
                        $Configuration.ADUsers[$User.UserPrincipalName]."Azure Licenses" =  $Licenses -join ","
                    }
                    $LicenseDetails =  $user.AssignedPlans.Where({$_.CapabilityStatus -eq "Enabled"}).Service
                    if($LicenseDetails.count -gt 0) {
                        $Configuration.ADUsers[$User.UserPrincipalName]."Azure Licenses Details" = $LicenseDetails -join ","
                    }
                    If($User.UserType) {
                        $Configuration.ADUsers[$User.UserPrincipalName]."Azure User Type" = $User.UserType
                    }
                    If($User.Id) {
                        $Configuration.ADUsers[$User.UserPrincipalName]."Azure ID" = $User.Id
                    }
                    #Update Progress Bar
                    If ($null -ne $Configuration.Jobs["Entra ID"]) {
                        $Configuration.Jobs["Entra ID"].process.PercentComplete = (($ProgressCounter/$Users.Count)*100)
                        $Configuration.Jobs["Entra ID"].process.Status = "Processing Record " + $ProgressCounter + " of " + $Users.Count
                    }
                }	
                $ProgressCounter++				
            } 
        }Else{
            throw "Missing Azure Tenant ID. Entra ID"
        }
        $Configuration.Jobs["Entra ID"].process.Completed = $true
        $Configuration.Jobs["Entra ID"].process.Status = "Completed"
        $Configuration.Jobs["Entra ID"].process.PercentComplete = 100
    }
    $powershell = [powershell]::Create()
    $powershell.RunspacePool = $runspacepool
    $powershell.AddScript($MgUsers) | Out-Null
    $powershell.AddParameter('Configuration',$Configuration) | Out-Null
    $Configuration.Jobs["Entra ID"] = [PSCustomObject]@{ Pipe = $powershell; Status = $powershell.BeginInvoke(); Name= "Entra ID" ;Start = Get-Date ; End = ""; process  = [PSCustomObject]@{ Id= ($Configuration.Jobs.Count + 2); Activity = "Processing Entra ID Users" ; Status = "Starting"; PercentComplete = 0;Completed =$false} }
    #endregion Entra ID Users
    #region Local Exchange Info
        #region Local Exchange Mailbox
            # Write-TMOutput -InputObject "`tExchange Mailboxes . . ."  -ForegroundColor DarkGray
            $EPMailboxesScript = {
                param(
                    $Configuration
                )
                #Import Class ADExchangeOutput
                Invoke-Expression $Configuration.ClassADExchangeOutput
                
                If ((Get-PSSession | Where-Object { $_.ConfigurationName -Match "Microsoft.Exchange" -and $_.ComputerName -eq $Configuration.ExchangeServer}).Count -eq 0 ) {
                    $ERPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://$($Configuration.ExchangeServer)/PowerShell/ -Authentication Negotiate -AllowRedirection 
                    Import-PSSession $ERPSession -AllowClobber -DisableNameChecking -Prefix "EP" | Out-Null
                }
                $Mailboxes = Get-EPMailbox -ResultSize unlimited
                $MailboxesCount = $Mailboxes.Count
                If ($PSVersionTable.PSVersion.Major -ge 7) {
                    $ProgressCounter = 1
                    $Mailboxes | ForEach-Object -Parallel {
                        $Mailbox = $_
                        $MailboxesCount = $using:MailboxesCount
                        $CUPN = $Mailbox.UserPrincipalName
                        $ProgressCounter = $using:ProgressCounter
                        $Configuration = $using:Configuration

                        Invoke-Expression $Configuration.ClassADExchangeOutput

                        If((-Not([string]::IsNullOrWhiteSpace($CUPN))) -and (-Not([string]::IsNullOrWhiteSpace($Mailbox.ServerName)))) {

                            $Configuration.ADUsers[$CUPN]."Mailbox Location" = "Local"
                            $Configuration.ADUsers[$CUPN]."Exchange Recipient Type" = $Mailbox.RecipientType
                            $Configuration.ADUsers[$CUPN]."Mailbox Server" = $Mailbox.ServerName
                            $Configuration.ADUsers[$CUPN]."Mailbox Database" = $Mailbox.Database
                            If($Configuration.ADUsers[$CUPN]."Mailbox Creation Date" -ne $Mailbox.WhenMailboxCreated) {
                                $Configuration.ADUsers[$CUPN]."Mailbox Creation Date" = $Mailbox.WhenMailboxCreated
                            }
                            $Configuration.ADUsers[$CUPN]."Mailbox Use Database Quota Defaults" = $Mailbox.UseDatabaseQuotaDefaults
                            $Configuration.ADUsers[$CUPN]."Mailbox Prohibit Send Quota" = Format-MailboxGB($Mailbox.ProhibitSendQuota)
                            $Configuration.ADUsers[$CUPN]."Mailbox Prohibit Send Receive Quota" = Format-MailboxGB($Mailbox.ProhibitSendReceiveQuota)
                            $Configuration.ADUsers[$CUPN]."Mailbox Issue Warning Quota" = Format-MailboxGB($Mailbox.IssueWarningQuota)
                            If ($Configuration.ADUsers[$CUPN]."Mailbox GUID" -ne $Mailbox.ExchangeGuid -and (-Not ($Mailbox.ExchangeGuid -eq [System.Guid]::Empty))) {
                                $Configuration.ADUsers[$CUPN]."Mailbox GUID" = $Mailbox.ExchangeGuid
                            }
                            If ($Configuration.ADUsers[$CUPN]."Mailbox Archive GUID" -ne $Mailbox.ArchiveGuid -and (-Not ($Mailbox.ArchiveGuid -eq [System.Guid]::Empty))) {
                                $Configuration.ADUsers[$CUPN]."Mailbox Archive GUID" = $Mailbox.ArchiveGuid
                            }
                            $Configuration.ADUsers[$CUPN]."Mailbox Forwarding Address SMTP" = $Mailbox.ForwardingSmtpAddress
                            $Configuration.ADUsers[$CUPN]."Mailbox Forwarding Address" = $Mailbox.ForwardingAddress		
                            # $Configuration.ADUsers[$CUPN]."Mailbox Recoverable Items Quota" = $Mailbox.RecoverableItemsQuota
                            # $Configuration.ADUsers[$CUPN]."Mailbox Recoverable Items Warning Quota" = $Mailbox.RecoverableItemsWarningQuota
                            # $Configuration.ADUsers[$CUPN]."Mailbox Retain Deleted Items For" = $Mailbox.RetainDeletedItemsFor
                            # $Configuration.ADUsers[$CUPN]."Mailbox Retain Deleted Items Until" = $Mailbox.RetainDeletedItemsUntil
                            # $Configuration.ADUsers[$CUPN]."Mailbox Retain Deleted Items Until Date" = $Mailbox.RetainDeletedItemsUntilDate

                            #Update Record
                            # $Configuration.ADUsers[$Mailboxes.UserPrincipalName] = $Record
                            # $Configuration.LEMailbox[$User.UserPrincipalName]= $Record

                            #Update Progress
                            If ($null -ne $Configuration.Jobs["Local Exchange Mailbox"]) {
                                $Configuration.Jobs["Local Exchange Mailbox"].process.PercentComplete = (($ProgressCounter/$Mailboxes.Count)*100)
                                $Configuration.Jobs["Local Exchange Mailbox"].process.Status = "Processing Record " + $ProgressCounter + " of " + $Mailboxes.Count
                            }
                        }
                        $ProgressCounter++
                    } -ThrottleLimit $Configuration.processbatchjobs
                }Else{
                    ForEach ($Mailbox in $Mailboxes) {
                        $CUPN = $Mailbox.UserPrincipalName
                        If((-Not([string]::IsNullOrWhiteSpace($CUPN))) -and (-Not([string]::IsNullOrWhiteSpace($Mailbox.ServerName)))) {

                            $Configuration.ADUsers[$CUPN]."Mailbox Location" = "Local"
                            $Configuration.ADUsers[$CUPN]."Exchange Recipient Type" = $Mailbox.RecipientType
                            $Configuration.ADUsers[$CUPN]."Mailbox Server" = $Mailbox.ServerName
                            $Configuration.ADUsers[$CUPN]."Mailbox Database" = $Mailbox.Database
                            If($Configuration.ADUsers[$CUPN]."Mailbox Creation Date" -ne $Mailbox.WhenMailboxCreated) {
                                $Configuration.ADUsers[$CUPN]."Mailbox Creation Date" = $Mailbox.WhenMailboxCreated
                            }
                            $Configuration.ADUsers[$CUPN]."Mailbox Use Database Quota Defaults" = $Mailbox.UseDatabaseQuotaDefaults
                            $Configuration.ADUsers[$CUPN]."Mailbox Prohibit Send Quota" = Format-MailboxGB($Mailbox.ProhibitSendQuota)
                            $Configuration.ADUsers[$CUPN]."Mailbox Prohibit Send Receive Quota" = Format-MailboxGB($Mailbox.ProhibitSendReceiveQuota)
                            $Configuration.ADUsers[$CUPN]."Mailbox Issue Warning Quota" = Format-MailboxGB($Mailbox.IssueWarningQuota)
                            If ($Configuration.ADUsers[$CUPN]."Mailbox GUID" -ne $Mailbox.ExchangeGuid -and (-Not ($Mailbox.ExchangeGuid -eq [System.Guid]::Empty))) {
                                $Configuration.ADUsers[$CUPN]."Mailbox GUID" = $Mailbox.ExchangeGuid
                            }
                            If ($Configuration.ADUsers[$CUPN]."Mailbox Archive GUID" -ne $Mailbox.ArchiveGuid -and (-Not ($Mailbox.ArchiveGuid -eq [System.Guid]::Empty))) {
                                $Configuration.ADUsers[$CUPN]."Mailbox Archive GUID" = $Mailbox.ArchiveGuid
                            }			
                            # $Configuration.ADUsers[$CUPN]."Mailbox Recoverable Items Quota" = $Mailbox.RecoverableItemsQuota
                            # $Configuration.ADUsers[$CUPN]."Mailbox Recoverable Items Warning Quota" = $Mailbox.RecoverableItemsWarningQuota
                            # $Configuration.ADUsers[$CUPN]."Mailbox Retain Deleted Items For" = $Mailbox.RetainDeletedItemsFor
                            # $Configuration.ADUsers[$CUPN]."Mailbox Retain Deleted Items Until" = $Mailbox.RetainDeletedItemsUntil
                            # $Configuration.ADUsers[$CUPN]."Mailbox Retain Deleted Items Until Date" = $Mailbox.RetainDeletedItemsUntilDate
                            $Configuration.ADUsers[$CUPN]."Mailbox Forwarding Address SMTP" = $Mailbox.ForwardingSmtpAddress
                            $Configuration.ADUsers[$CUPN]."Mailbox Forwarding Address" = $Mailbox.ForwardingAddress		
                            #Update Progress
                            If ($null -ne $Configuration.Jobs["Local Exchange Mailbox"]) {
                                $Configuration.Jobs["Local Exchange Mailbox"].process.PercentComplete = (($ProgressCounter/$MailboxesCount)*100)
                                $Configuration.Jobs["Local Exchange Mailbox"].process.Status = "Processing Record " + $ProgressCounter + " of " + $MailboxesCount
                            }
                        }
                        $ProgressCounter++
                    }
                }
                $Configuration.Jobs["Local Exchange Mailbox"].process.Completed = $true
                $Configuration.Jobs["Local Exchange Mailbox"].process.Status = "Completed"
                $Configuration.Jobs["Local Exchange Mailbox"].process.PercentComplete = 100
            }
            $powershell = [powershell]::Create()
            $powershell.RunspacePool = $runspacepool
            $powershell.AddScript($EPMailboxesScript) | Out-Null
            $powershell.AddParameter('Configuration',$Configuration) | Out-Null
            $Configuration.Jobs["Local Exchange Mailbox"] = [PSCustomObject]@{ Pipe = $powershell; Status = $powershell.BeginInvoke(); Name= "Local Exchange Mailbox" ;Start = Get-Date ; End = ""; process  = [PSCustomObject]@{ Id= ($Configuration.Jobs.Count + 2); Activity = "Processing Local Exchange Mailbox" ; Status = "Starting"; PercentComplete = 0;Completed =$false} }
        #endregion Local Exchange Mailbox
        #region Local Exchange Mailboxes Statistics
            $EPMailboxesStatsScript = {
                param(
                    $Configuration
                )
                #Import Class ADExchangeOutput
                Invoke-Expression $Configuration.ClassADExchangeOutput
        
                If ((Get-PSSession | Where-Object { $_.ConfigurationName -Match "Microsoft.Exchange" -and $_.ComputerName -eq $Configuration.ExchangeServer}).Count -eq 0 ) {
                    $ERPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://$($Configuration.ExchangeServer)/PowerShell/ -Authentication Negotiate -AllowRedirection 
                    Import-PSSession $ERPSession -AllowClobber -DisableNameChecking -Prefix "EP" | Out-Null
                }
                # Get-EPMailbox -ResultSize unlimited | Get-EPMailboxStatistics
                $Mailboxes = Get-EPMailbox -ResultSize unlimited | Get-EPMailboxStatistics
                $ProgressCounter = 1
                If ($PSVersionTable.PSVersion.Major -ge 7) {
                    $Configuration.ADUsersWithMailboxGUIDs | ForEach-Object -Parallel {
                        $CUPN = $_
                        $Configuration = $using:Configuration
                        $ProgressCounter = $using:ProgressCounter
                        $MailboxStats = $Mailboxes.Where({$_.Identity.MailboxGuid.Guid -eq ($Configuration.ADUsers[$CUPN].'Mailbox GUID')})
                        If($MailboxStats.DisplayName -eq $null) {
                            $MailboxStats = $Mailboxes.Where({$_.DisplayName -eq ($Configuration.ADUsers[$CUPN].'Display Name')})
                        }

                        If ($Configuration.ADUsers[$CUPN].'Mailbox GUID'-and $MailboxStats.Identity.MailboxGuid.Guid) {

                            $TotalItemSize = Format-MailboxGB ($MailboxStats.TotalItemSize)
                            $TotalDeletedItemSize = Format-MailboxGB ($MailboxStats.TotalDeletedItemSize) 
                            If ($TotalItemSize -ne "Unlimited" -and $TotalDeletedItemSize -ne "Unlimited" ) {
                                $Configuration.ADUsers[$CUPN]."Mailbox Size GB" = ($TotalItemSize + $TotalDeletedItemSize)
                                If (-Not ([string]::IsNullOrWhiteSpace($Configuration.ADUsers[$CUPN].'Mailbox Prohibit Send Quota')) -and $Configuration.ADUsers[$CUPN].'Mailbox Prohibit Send Quota' -ne "Unlimited" -and $Configuration.ADUsers[$CUPN].'Mailbox Prohibit Send Quota' -ne 0) {
                                    $Configuration.ADUsers[$CUPN]."Mailbox Usage %" = [math]::Round((($TotalItemSize + $TotalDeletedItemSize)/$Configuration.ADUsers[$CUPN].'Mailbox Prohibit Send Quota')*100,2)
                                }
                            }
                            $Configuration.ADUsers[$CUPN]."Mailbox Item Count" = $MailboxStats.ItemCount
                            $Configuration.ADUsers[$CUPN]."Mailbox Last Logged On User Account" = $MailboxStats.LastLoggedOnUserAccount
                            $Configuration.ADUsers[$CUPN]."Mailbox Last Logon Time" = $MailboxStats.LastLogonTime
                            $Configuration.ADUsers[$CUPN]."Mailbox Last Logoff Time" = $MailboxStats.LastLogoffTime
                            
                            #Update Progress
                            If ($null -ne$Configuration.Jobs["Local Exchange Mailbox Statistics"]) {
                                $Configuration.Jobs["Local Exchange Mailbox Statistics"].process.PercentComplete = (($ProgressCounter/$Configuration.ADUsersWithMailboxGUIDs.Count)*100)
                                $Configuration.Jobs["Local Exchange Mailbox Statistics"].process.Status = "Processing Record " + $ProgressCounter + " of " + $Configuration.ADUsersWithMailboxGUIDs.Count
                            }
                        }
                        $ProgressCounter++
                    } -ThrottleLimit $Configuration.ProcessBatchJobs
                }Else {
                    ForEach ( $CUPN in $Configuration.ADUsersWithMailboxGUIDs ) {
                        $MailboxStats = $Mailboxes.Where({$_.Identity.MailboxGuid.Guid -eq ($Configuration.ADUsers[$CUPN].'Mailbox GUID')})
                        If($MailboxStats.DisplayName -eq $null) {
                            $MailboxStats = $Mailboxes.Where({$_.DisplayName -eq ($Configuration.ADUsers[$CUPN].'Display Name')})
                        }
                              
                        If ($Configuration.ADUsers[$CUPN].'Mailbox GUID' -and $MailboxStats.MailboxGuid) {
                            $TotalItemSize = Format-MailboxGB ($MailboxStats.TotalItemSize)
                            $TotalDeletedItemSize = Format-MailboxGB ($MailboxStats.TotalDeletedItemSize) 
                            If ($TotalItemSize -ne "Unlimited" -and $TotalDeletedItemSize -ne "Unlimited" ) {
                                $Configuration.ADUsers[$CUPN]."Mailbox Size GB" = ($TotalItemSize + $TotalDeletedItemSize)
                                If (-Not ([string]::IsNullOrWhiteSpace($Configuration.ADUsers[$CUPN].'Mailbox Prohibit Send Quota')) -and $Configuration.ADUsers[$CUPN].'Mailbox Prohibit Send Quota' -ne "Unlimited" -and $Configuration.ADUsers[$CUPN].'Mailbox Prohibit Send Quota' -ne 0) {
                                    $Configuration.ADUsers[$CUPN]."Mailbox Usage %" = [math]::Round((($TotalItemSize + $TotalDeletedItemSize)/$Configuration.ADUsers[$CUPN].'Mailbox Prohibit Send Quota')*100,2)
                                }
                            }
                            $Configuration.ADUsers[$CUPN]."Mailbox Item Count" = $MailboxStats.ItemCount
                            $Configuration.ADUsers[$CUPN]."Mailbox Last Logged On User Account" = $MailboxStats.LastLoggedOnUserAccount
                            $Configuration.ADUsers[$CUPN]."Mailbox Last Logon Time" = $MailboxStats.LastLogonTime
                            $Configuration.ADUsers[$CUPN]."Mailbox Last Logoff Time" = $MailboxStats.LastLogoffTime

                            #Update Progress
                            If ($null -ne $Configuration.Jobs["Local Exchange Mailbox Statistics"]) {
                                $Configuration.Jobs["Local Exchange Mailbox Statistics"].process.PercentComplete = (($ProgressCounter/$Configuration.ADUsersWithMailboxGUIDs.Count)*100)
                                $Configuration.Jobs["Local Exchange Mailbox Statistics"].process.Status = "Processing Record " + $ProgressCounter + " of " + $Configuration.ADUsersWithMailboxGUIDs.Count
                            }
                        }
                        $ProgressCounter++
                    }
                }
                $Configuration.Jobs["Local Exchange Mailbox Statistics"].process.Completed = $true
                $Configuration.Jobs["Local Exchange Mailbox Statistics"].process.Status = "Completed"
                $Configuration.Jobs["Local Exchange Mailbox Statistics"].process.PercentComplete = 100
            }
            $powershell = [powershell]::Create()
            $powershell.RunspacePool = $runspacepool
            $powershell.AddScript($EPMailboxesStatsScript) | Out-Null
            $powershell.AddParameter('Configuration',$Configuration) | Out-Null
            $Configuration.Jobs["Local Exchange Mailbox Statistics"] = [PSCustomObject]@{ Pipe = $powershell; Status = $powershell.BeginInvoke(); Name= "Local Exchange Mailbox Statistics" ;Start = Get-Date ; End = ""; process  = [PSCustomObject]@{ Id= ($Configuration.Jobs.Count + 2); Activity = "Processing Local Exchange Mailbox Statistics" ; Status = "Starting"; PercentComplete = 0;Completed =$false} }
        #endregion Local Exchange Mailboxes Statistics
        #region Local Exchange Mailboxes Permissions
            $EPMailboxesPermissionsScript = {
                param(
                    $Configuration
                )
                $ProgressCounter = 1
                #Import Class ADExchangeOutput
                Invoke-Expression $Configuration.ClassADExchangeOutput
                #Connect to Exchange
                If ((Get-PSSession | Where-Object { $_.ConfigurationName -Match "Microsoft.Exchange" -and $_.ComputerName -eq $Configuration.ExchangeServer}).Count -eq 0 ) {
                    $ERPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://$($Configuration.ExchangeServer)/PowerShell/ -Authentication Negotiate -AllowRedirection 
                    Import-PSSession $ERPSession -AllowClobber -DisableNameChecking -Prefix "EP" | Out-Null
                }
                #Load .Net Assembly for AD
                Add-Type -AssemblyName System.DirectoryServices.AccountManagement
                $ContextType = [System.DirectoryServices.AccountManagement.ContextType]::Domain
                $IdentityType = [System.DirectoryServices.AccountManagement.IdentityType]::SamAccountName
                ForEach ( $CUPN in $Configuration.ADUsersUPNsWithEmails ) {
                    $EPMailboxesPerms = Get-EPMailboxPermission -Identity $CUPN -erroraction SilentlyContinue
                    If($EPMailboxesPerms.count -gt 0) {
                        $EPMailboxesPerms = Get-EPMailboxPermission -Identity $Configuration.ADUsers[$CUPN]."Mailbox GUID" -erroraction SilentlyContinue
                    }
                    If($EPMailboxesPerms.count -gt 0) {
                        $EPMailboxesPerms = Get-EPMailboxPermission -Identity $Configuration.ADUsers[$CUPN]."Email" -erroraction SilentlyContinue
                    }
    
                    If($CUPN -and $EPMailboxesPerms.count -gt 0){
                        #region Mailbox Permission
                        $DAMP = @()
                        $FixedDAMP = @{}
                        $CSVFixedDAMP = ""
                        $DAMP = ($EPMailboxesPerms.Where({($_.AccessRights -eq "FullAccess") -and ($_.User -notin $Configuration.ExcludeUsers) -and ($_.User -notmatch "S-1-5-*" ) -and $_.IsInherited -eq $false})) | Sort-Object -Unique -Property User	
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
                                                    If ($ADUser.description -notmatch "Leave") {
                                                        $FixedDAMP.add(($ACE.User + " - Leave"),$ACE.AccessRights)
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
                                    $Configuration.ADUsers[$CUPN]."Mailbox Permissions" = $CSVFixedDAMP
                                    #Update Progress
                                    If ($null -ne $Configuration.Jobs["Local Exchange Mailboxes Permissions"]) {
                                        $Configuration.Jobs["Local Exchange Mailboxes Permissions"].process.PercentComplete = (($ProgressCounter/$Configuration.ADUsersUPNsWithEmails.Count)*100)
                                        $Configuration.Jobs["Local Exchange Mailboxes Permissions"].process.Status = "Processing Record " + $ProgressCounter + " of " + $Configuration.ADUsersUPNsWithEmails.Count
                                    }
                                }
                            }
                        }	
                        #endregion Mailbox Permission
                    }
                    $ProgressCounter++
                }
            
            }
            $powershell = [powershell]::Create()
            $powershell.RunspacePool = $runspacepool
            $powershell.AddScript($EPMailboxesPermissionsScript) | Out-Null
            $powershell.AddParameter('Configuration',$Configuration) | Out-Null
            $Configuration.Jobs["Local Exchange Mailboxes Permissions"] = [PSCustomObject]@{ Pipe = $powershell; Status = $powershell.BeginInvoke(); Name= "Local Exchange Mailboxes Permissions" ;Start = Get-Date ; End = ""; process  = [PSCustomObject]@{ Id= ($Configuration.Jobs.Count + 2); Activity = "Processing Local Exchange Mailboxes Permissions" ; Status = "Starting"; PercentComplete = 0;Completed =$false} }                
        #endregion Local Exchange Mailboxes Permissions
        #region Local Exchange Mailbox Archive
            $EPMailboxesArchiveScript = {
                param(
                    $Configuration
                )
                #Import Class ADExchangeOutput
                Invoke-Expression $Configuration.ClassADExchangeOutput
        
                If ((Get-PSSession | Where-Object { $_.ConfigurationName -Match "Microsoft.Exchange" -and $_.ComputerName -eq $Configuration.ExchangeServer}).Count -eq 0 ) {
                    $ERPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://$($Configuration.ExchangeServer)/PowerShell/ -Authentication Negotiate -AllowRedirection 
                    Import-PSSession $ERPSession -AllowClobber -DisableNameChecking -Prefix "EP" | Out-Null
                }
                # Get-EPMailbox -ResultSize unlimited -Archive | Where-Object {$_.ArchiveState -ne "HostedProvisioned"}
                $Mailboxes = Get-EPMailbox -ResultSize unlimited -Archive | Where-Object {$_.ArchiveState -ne "HostedProvisioned"}
                $MailboxesCount = $Mailboxes.Count
                $ProgressCounter = 1
                If ($PSVersionTable.PSVersion.Major -ge 7) {
                    $Mailboxes | ForEach-Object -Parallel {
                        $Mailbox = $_
                        $CUPN = $Mailbox.UserPrincipalName
                        $MailboxesCount = $using:MailboxesCount
                        $Configuration = $using:Configuration
                        $ProgressCounter = $using:ProgressCounter
                        #Import Class ADExchangeOutput
                        Invoke-Expression $Configuration.ClassADExchangeOutput

                        If($Configuration.ADUsers[$CUPN]."User Principal Name" -and $Mailbox.UserPrincipalName){
                            $Configuration.ADUsers[$CUPN]."Mailbox Archive Location" = "Local"
                            $Configuration.ADUsers[$CUPN]."Mailbox Archive State" = $Mailbox.ArchiveState
                            $Configuration.ADUsers[$CUPN]."Mailbox Archive Server" = $Mailbox.ServerName
                            $Configuration.ADUsers[$CUPN]."Mailbox Archive Database" = $Mailbox.ArchiveDatabase
                            If($Configuration.ADUsers[$CUPN]."Mailbox Archive Creation Date" -ne $Mailbox.WhenMailboxCreated) {
                                $Configuration.ADUsers[$CUPN]."Mailbox Archive Creation Date" = $Mailbox.WhenMailboxCreated
                            }
                            $Configuration.ADUsers[$CUPN]."Mailbox Archive Use Database Quota Defaults" = $Mailbox.UseDatabaseQuotaDefaults
                            $Configuration.ADUsers[$CUPN]."Mailbox Archive Prohibit Send Quota" = Format-MailboxGB($Mailbox.ProhibitSendQuota)
                            $Configuration.ADUsers[$CUPN]."Mailbox Archive Prohibit Send Receive Quota" = Format-MailboxGB($Mailbox.ProhibitSendReceiveQuota)
                            $Configuration.ADUsers[$CUPN]."Mailbox Archive Issue Warning Quota" = Format-MailboxGB($Mailbox.IssueWarningQuota)
                            If ($Configuration.ADUsers[$CUPN]."Mailbox Archive GUID" -ne $Mailbox.ArchiveGuid -and (-Not ($Mailbox.ArchiveGuid -eq [System.Guid]::Empty))) {
                                $Configuration.ADUsers[$CUPN]."Mailbox Archive GUID" = $Mailbox.ArchiveGuid
                            }
                            #Update Progress
                            If ($null -ne $Configuration.Jobs["Local Exchange Mailbox Archive"]) {
                                $Configuration.Jobs["Local Exchange Mailbox Archive"].process.PercentComplete = (($ProgressCounter/ $MailboxesCount)*100)
                                $Configuration.Jobs["Local Exchange Mailbox Archive"].process.Status = "Processing Record " + $ProgressCounter + " of " +  $MailboxesCount
                            }
                        }
                        $ProgressCounter++
                    } -ThrottleLimit $Configuration.ProcessBatchJobs
                }Else{
                    ForEach ( $Mailbox in $Mailboxes ) {
                        $CUPN = $Mailbox.UserPrincipalName

                        If($Configuration.ADUsers[$CUPN]."User Principal Name" -and $Mailbox.UserPrincipalName){
                            $Configuration.ADUsers[$CUPN]."Mailbox Archive Location" = "Local"
                            $Configuration.ADUsers[$CUPN]."Mailbox Archive State" = $Mailbox.ArchiveState
                            $Configuration.ADUsers[$CUPN]."Mailbox Archive Server" = $Mailbox.ServerName
                            $Configuration.ADUsers[$CUPN]."Mailbox Archive Database" = $Mailbox.ArchiveDatabase
                            If($Configuration.ADUsers[$CUPN]."Mailbox Archive Creation Date" -ne $Mailbox.WhenMailboxCreated) {
                                $Configuration.ADUsers[$CUPN]."Mailbox Archive Creation Date" = $Mailbox.WhenMailboxCreated
                            }
                            $Configuration.ADUsers[$CUPN]."Mailbox Archive Use Database Quota Defaults" = $Mailbox.UseDatabaseQuotaDefaults
                            $Configuration.ADUsers[$CUPN]."Mailbox Archive Prohibit Send Quota" = Format-MailboxGB($Mailbox.ProhibitSendQuota)
                            $Configuration.ADUsers[$CUPN]."Mailbox Archive Prohibit Send Receive Quota" = Format-MailboxGB($Mailbox.ProhibitSendReceiveQuota)
                            $Configuration.ADUsers[$CUPN]."Mailbox Archive Issue Warning Quota" = Format-MailboxGB($Mailbox.IssueWarningQuota)
                            If ($Configuration.ADUsers[$CUPN]."Mailbox Archive GUID" -ne $Mailbox.ArchiveGuid -and (-Not ($Mailbox.ArchiveGuid -eq [System.Guid]::Empty))) {
                                $Configuration.ADUsers[$CUPN]."Mailbox Archive GUID" = $Mailbox.ArchiveGuid
                            }
                             #Update Progress
                             If ($null -ne $Configuration.Jobs["Local Exchange Mailbox Archive"]) {
                                $Configuration.Jobs["Local Exchange Mailbox Archive"].process.PercentComplete = (($ProgressCounter/ $MailboxesCount)*100)
                                $Configuration.Jobs["Local Exchange Mailbox Archive"].process.Status = "Processing Record " + $ProgressCounter + " of " +  $MailboxesCount
                             }
                        }
                        $ProgressCounter++
                    }
                }
        
            }
            $powershell = [powershell]::Create()
            $powershell.RunspacePool = $runspacepool
            $powershell.AddScript($EPMailboxesArchiveScript) | Out-Null
            $powershell.AddParameter('Configuration',$Configuration) | Out-Null
            $Configuration.Jobs["Local Exchange Mailbox Archive"] = [PSCustomObject]@{ Pipe = $powershell; Status = $powershell.BeginInvoke(); Name= "Local Exchange Mailbox Archive" ;Start = Get-Date ; End = ""; process  = [PSCustomObject]@{ Id= ($Configuration.Jobs.Count + 2); Activity = "Processing Local Exchange Mailbox Archive" ; Status = "Starting"; PercentComplete = 0;Completed =$false} }      
        #endregion Local Exchange Mailbox Archive
        #region Local Exchange Mailboxes Archive Statistics
            $EPMailboxesArchiveStatsScript = {
                param(
                    $Configuration
                )
                #Import Class ADExchangeOutput
                Invoke-Expression $Configuration.ClassADExchangeOutput

                If ((Get-PSSession | Where-Object { $_.ConfigurationName -Match "Microsoft.Exchange" -and $_.ComputerName -eq $Configuration.ExchangeServer}).Count -eq 0 ) {
                    $ERPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://$($Configuration.ExchangeServer)/PowerShell/ -Authentication Negotiate -AllowRedirection 
                    Import-PSSession $ERPSession -AllowClobber -DisableNameChecking -Prefix "EP" | Out-Null
                }
                $ProgressCounter = 1
                $Mailboxes = Get-EPMailbox -ResultSize unlimited -Archive | Where-Object {$_.ArchiveState -ne "HostedProvisioned"}| Get-EPMailboxStatistics -Archive

                If ($PSVersionTable.PSVersion.Major -ge 7) {
                    $Configuration.ADUsersWithMailboxGUIDs | ForEach-Object -Parallel {
                        $CUPN = $_
                        $ProgressCounter = $using:ProgressCounter   
                        $Configuration = $using:Configuration
                        $Mailboxes = $using:Mailboxes
                        
                        $MailboxStats = $Mailboxes.Where({$_.Identity.MailboxGuid.Guid -eq ($Configuration.ADUsers[$CUPN].'Mailbox Archive GUID')})
                        If($MailboxStats.DisplayName -eq $null) {
                            $MailboxStats = $Mailboxes.Where({$_.DisplayName -eq ($Configuration.ADUsers[$CUPN].'Display Name')})
                        }

                        If( $Configuration.ADUsers[$CUPN]."Mailbox Archive GUID" -and $MailboxStats.Identity.MailboxGuid.Guid){
                             $Configuration.ADUsers[$CUPN]."Mailbox Archive Storage Limit Status" = $MailboxStats.StorageLimitStatus
                            $TotalItemSize = Format-MailboxGB ($MailboxStats.TotalItemSize)
                            $TotalDeletedItemSize = Format-MailboxGB ($MailboxStats.TotalDeletedItemSize) 
                            If ($TotalItemSize -ne "Unlimited" -and $TotalDeletedItemSize -ne "Unlimited" ) {
                                 $Configuration.ADUsers[$CUPN]."Mailbox Archive Size GB" = ($TotalItemSize + $TotalDeletedItemSize)
                            }
                             $Configuration.ADUsers[$CUPN]."Mailbox Archive Item Count" = $MailboxStats.ItemCount
                             $Configuration.ADUsers[$CUPN]."Mailbox Archive Last Logged On User Account" = $MailboxStats.LastLoggedOnUserAccount
                             $Configuration.ADUsers[$CUPN]."Mailbox Archive Last Logon Time" = $MailboxStats.LastLogonTime
                             $Configuration.ADUsers[$CUPN]."Mailbox Archive Last Logoff Time" = $MailboxStats.LastLogoffTime
                             $Configuration.ADUsers[$CUPN]."Mailbox Archive Forwarding Address" = $MailboxStats.ForwardingAddress
                             $Configuration.ADUsers[$CUPN]."Mailbox Archive Forwarding Address SMTP" = $MailboxStats.ForwardingSmtpAddress
                             #Update Progress
                             If ($null -ne $Configuration.Jobs["Local Exchange Mailboxes Archive Statistics"]) {
                                $Configuration.Jobs["Local Exchange Mailboxes Archive Statistics"].process.PercentComplete = (($ProgressCounter/$Configuration.ADUsersWithMailboxGUIDs.Count)*100)
                                $Configuration.Jobs["Local Exchange Mailboxes Archive Statistics"].process.Status = "Processing Record " + $ProgressCounter + " of " + $Configuration.ADUsersWithMailboxGUIDs.Count
                             }
                        }
                        $ProgressCounter++
                    } -ThrottleLimit $Configuration.processbatchjobs
                }Else{
                    ForEach ( $CUPN in $Configuration.ADUsersWithMailboxGUIDs ) {
                        $MailboxStats = $Mailboxes.Where({$_.Identity.MailboxGuid.Guid -eq ($Configuration.ADUsers[$CUPN].'Mailbox Archive GUID')})
                        If($MailboxStats.DisplayName -eq $null) {
                            $MailboxStats = $Mailboxes.Where({$_.DisplayName -eq ($Configuration.ADUsers[$CUPN].'Display Name')})
                        }
                        If( $Configuration.ADUsers[$CUPN]."Mailbox Archive GUID" -and $MailboxStats.Identity.MailboxGuid.Guid){
                             $Configuration.ADUsers[$CUPN]."Mailbox Archive Storage Limit Status" = $MailboxStats.StorageLimitStatus
                            $TotalItemSize = Format-MailboxGB ($MailboxStats.TotalItemSize)
                            $TotalDeletedItemSize = Format-MailboxGB ($MailboxStats.TotalDeletedItemSize) 
                            If ($TotalItemSize -ne "Unlimited" -and $TotalDeletedItemSize -ne "Unlimited" ) {
                                 $Configuration.ADUsers[$CUPN]."Mailbox Archive Size GB" = ($TotalItemSize + $TotalDeletedItemSize)
                            }
                             $Configuration.ADUsers[$CUPN]."Mailbox Archive Item Count" = $MailboxStats.ItemCount
                             $Configuration.ADUsers[$CUPN]."Mailbox Archive Last Logged On User Account" = $MailboxStats.LastLoggedOnUserAccount
                             $Configuration.ADUsers[$CUPN]."Mailbox Archive Last Logon Time" = $MailboxStats.LastLogonTime
                             $Configuration.ADUsers[$CUPN]."Mailbox Archive Last Logoff Time" = $MailboxStats.LastLogoffTime
                             $Configuration.ADUsers[$CUPN]."Mailbox Archive Forwarding Address" = $MailboxStats.ForwardingAddress
                             $Configuration.ADUsers[$CUPN]."Mailbox Archive Forwarding Address SMTP" = $MailboxStats.ForwardingSmtpAddress
                            #Update Progress
                            If ($null -ne $Configuration.Jobs["Local Exchange Mailboxes Archive Statistics"]) {
                                $Configuration.Jobs["Local Exchange Mailboxes Archive Statistics"].process.PercentComplete = (($ProgressCounter/$Configuration.ADUsersWithMailboxGUIDs.Count)*100)
                                $Configuration.Jobs["Local Exchange Mailboxes Archive Statistics"].process.Status = "Processing Record " + $ProgressCounter + " of " + $Configuration.ADUsersWithMailboxGUIDs.Count
                            }
                        }
                        $ProgressCounter++
                    }
                }
            }
            $powershell = [powershell]::Create()
            $powershell.RunspacePool = $runspacepool
            $powershell.AddScript($EPMailboxesArchiveStatsScript) | Out-Null
            $powershell.AddParameter('Configuration',$Configuration) | Out-Null
            $Configuration.Jobs["Local Exchange Mailboxes Archive Statistics"] = [PSCustomObject]@{ Pipe = $powershell; Status = $powershell.BeginInvoke(); Name= "Local Exchange Mailboxes Archive Statistics" ;Start = Get-Date ; End = ""; process  = [PSCustomObject]@{ Id= ($Configuration.Jobs.Count + 2); Activity = "Processing Local Exchange Mailboxes Archive Statistics" ; Status = "Starting"; PercentComplete = 0;Completed =$false} }                  
        #endregion Local Exchange Mailboxes Archive Statistics
        #region Local Exchange Archive Mailbox Permissions
            $EPMailboxesArchivePermissionsScript = {
                param(
                    $Configuration
                )
                #Import Class ADExchangeOutput
                Invoke-Expression $Configuration.ClassADExchangeOutput
        
                If ((Get-PSSession | Where-Object { $_.ConfigurationName -Match "Microsoft.Exchange" -and $_.ComputerName -eq $Configuration.ExchangeServer}).Count -eq 0 ) {
                    $ERPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://$($Configuration.ExchangeServer)/PowerShell/ -Authentication Negotiate -AllowRedirection 
                    Import-PSSession $ERPSession -AllowClobber -DisableNameChecking -Prefix "EP" | Out-Null
                }
                $ProgressCounter = 1
                #Load .Net Assembly for AD
                Add-Type -AssemblyName System.DirectoryServices.AccountManagement
                $ContextType = [System.DirectoryServices.AccountManagement.ContextType]::Domain
                $IdentityType = [System.DirectoryServices.AccountManagement.IdentityType]::SamAccountName

                $Mailboxes = (Get-EPMailbox -ResultSize unlimited -Archive| Where-Object {$_.ArchiveState -ne "HostedProvisioned"})
                $MailboxesCount = $Mailboxes.Count

                ForEach ( $Mailbox in $Mailboxes ) {
                    If( $Configuration.ADUsers[$CUPN]."User Principal Name" -and $Mailbox.UserPrincipalName){
                        $EPMailboxesPerms = $Mailbox |  Get-EPMailboxPermission
                        #region Mailbox Permission
                        $DAMP = @()
                        $FixedDAMP = @{}
                        $CSVFixedDAMP = ""
                        $DAMP = ($EPMailboxesPerms.Where({($_.AccessRights -eq "FullAccess") -and ($_.User -notin $Configuration.ExcludeUsers) -and ($_.User -notmatch "S-1-5-*" ) -and $_.IsInherited -eq $false})) | Sort-Object -Unique -Property User	
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
                                                    If ($ADUser.description -notmatch "Leave") {
                                                        $FixedDAMP.add(($ACE.User + " - Leave"),$ACE.AccessRights)
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
                                        $Configuration.ADUsers[$CUPN]."Mailbox Archive Permissions" = $CSVFixedDAMP
                                }
                            }
                        }	
                        #endregion Mailbox Permission
                            If ($null -ne $Configuration.Jobs["Local Exchange Mailboxes Archive Permissions"]) {
                            $Configuration.Jobs["Local Exchange Mailboxes Archive Permissions"].process.PercentComplete = (($ProgressCounter/$MailboxesCount)*100)
                            $Configuration.Jobs["Local Exchange Mailboxes Archive Permissions"].process.Status = "Processing Record " + $ProgressCounter + " of " + $MailboxesCount
                            }                            
                    }
                    $ProgressCounter++
                }
            
            }
            $powershell = [powershell]::Create()
            $powershell.RunspacePool = $runspacepool
            $powershell.AddScript($EPMailboxesArchivePermissionsScript) | Out-Null
            $powershell.AddParameter('Configuration',$Configuration) | Out-Null
            $Configuration.Jobs["Local Exchange Mailboxes Archive Permissions"] = [PSCustomObject]@{ Pipe = $powershell; Status = $powershell.BeginInvoke(); Name= "Local Exchange Mailboxes Archive Permissions" ;Start = Get-Date ; End = ""; process  = [PSCustomObject]@{ Id= ($Configuration.Jobs.Count + 2); Activity = "Processing Local Exchange Mailboxes Archive Permissions" ; Status = "Starting"; PercentComplete = 0;Completed =$false} }  
        #endregion Local Exchange Archive Mailbox Permissions
        #region Local Exchange Remote Mailbox
            $EPRemoteMailboxesScript = {
                param(
                    $Configuration
                )
                #Import Class ADExchangeOutput
                Invoke-Expression $Configuration.ClassADExchangeOutput
        
                If ((Get-PSSession | Where-Object { $_.ConfigurationName -Match "Microsoft.Exchange" -and $_.ComputerName -eq $Configuration.ExchangeServer}).Count -eq 0 ) {
                    $ERPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://$($Configuration.ExchangeServer)/PowerShell/ -Authentication Negotiate -AllowRedirection 
                    Import-PSSession $ERPSession -AllowClobber -DisableNameChecking -Prefix "EP" | Out-Null
                }
                $ProgressCounter = 1
                $Mailboxes = Get-EPRemoteMailbox -ResultSize unlimited
                $MailboxesCount = $Mailboxes.Count

                If ($PSVersionTable.PSVersion.Major -ge 7) {
                    $Mailboxes | ForEach-Object -Parallel {
                        $Mailbox = $_
                        $CUPN = $Mailbox.UserPrincipalName
                        $MailboxesCount = $using:MailboxesCount
                        $Configuration = $using:Configuration
                        $ProgressCounter = $using:ProgressCounter

                        #Import Class ADExchangeOutput
                        Invoke-Expression $Configuration.ClassADExchangeOutput

                        If($Mailbox.UserPrincipalName -and  $Configuration.ADUsers[$CUPN]."User Principal Name") {
                            If ( $Configuration.ADUsers[$CUPN]."Mailbox GUID" -ne $Mailbox.ExchangeGuid -and (-Not ($Mailbox.ExchangeGuid -eq [System.Guid]::Empty))) {
                                 $Configuration.ADUsers[$CUPN]."Mailbox GUID" = $Mailbox.ExchangeGuid
                            }
                            If ( $Configuration.ADUsers[$CUPN]."Mailbox Archive GUID" -ne $Mailbox.ArchiveGuid -and (-Not ($Mailbox.ArchiveGuid -eq [System.Guid]::Empty))) {
                                 $Configuration.ADUsers[$CUPN]."Mailbox Archive GUID" = $Mailbox.ArchiveGuid
                            }
                             $Configuration.ADUsers[$CUPN]."Exchange Remote Routing Address" = $Mailbox.RemoteRoutingAddress -replace "SMTP:"
                             $Configuration.ADUsers[$CUPN]."Mailbox Location" = "Hybrid Remote"
                             $Configuration.ADUsers[$CUPN]."Exchange Recipient Type" = $Mailbox.RecipientTypeDetails
                             $Configuration.ADUsers[$CUPN]."Mailbox Forwarding Address" = $Mailbox.ForwardingAddress
                             $Configuration.ADUsers[$CUPN]."Mailbox Forwarding Address SMTP" = $Mailbox.ForwardingSmtpAddress

                        }
                        #Update Progress
                        If ($null -ne $Configuration.Jobs["Local Exchange Remote Mailbox"]) {
                            $Configuration.Jobs["Local Exchange Remote Mailbox"].process.PercentComplete = (($ProgressCounter/$MailboxesCount)*100)
                            $Configuration.Jobs["Local Exchange Remote Mailbox"].process.Status = "Processing Record " + $ProgressCounter + " of " + $MailboxesCount
                        } 
                        $ProgressCounter++
                    } -ThrottleLimit $Configuration.ProcessBatchJobs
                }Else{
                    ForEach ( $Mailbox in $Mailboxes ) {
                        If($Mailbox.UserPrincipalName -and  $Configuration.ADUsers[$CUPN]."User Principal Name") {
                            If ( $Configuration.ADUsers[$CUPN]."Mailbox GUID" -ne $Mailbox.ExchangeGuid -and (-Not ($Mailbox.ExchangeGuid -eq [System.Guid]::Empty))) {
                                 $Configuration.ADUsers[$CUPN]."Mailbox GUID" = $Mailbox.ExchangeGuid
                            }
                            If ( $Configuration.ADUsers[$CUPN]."Mailbox Archive GUID" -ne $Mailbox.ArchiveGuid -and (-Not ($Mailbox.ArchiveGuid -eq [System.Guid]::Empty))) {
                                 $Configuration.ADUsers[$CUPN]."Mailbox Archive GUID" = $Mailbox.ArchiveGuid
                            }
                             $Configuration.ADUsers[$CUPN]."Exchange Remote Routing Address" = $Mailbox.RemoteRoutingAddress -replace "SMTP:"
                             $Configuration.ADUsers[$CUPN]."Mailbox Location" = "Hybrid Remote"
                             $Configuration.ADUsers[$CUPN]."Exchange Recipient Type" = $Mailbox.RecipientTypeDetails
                             $Configuration.ADUsers[$CUPN]."Mailbox Forwarding Address" = $Mailbox.ForwardingAddress
                             $Configuration.ADUsers[$CUPN]."Mailbox Forwarding Address SMTP" = $Mailbox.ForwardingSmtpAddress
                            #Update Progress
                            If ($null -ne $Configuration.Jobs["Local Exchange Remote Mailbox"]) {
                                $Configuration.Jobs["Local Exchange Remote Mailbox"].process.PercentComplete = (($ProgressCounter/$MailboxesCount)*100)
                                $Configuration.Jobs["Local Exchange Remote Mailbox"].process.Status = "Processing Record " + $ProgressCounter + " of " + $MailboxesCount
                            }  
                        }
                        $ProgressCounter++
                    }
                }
            }
            $powershell = [powershell]::Create()
            $powershell.RunspacePool = $runspacepool
            $powershell.AddScript($EPRemoteMailboxesScript) | Out-Null
            $powershell.AddParameter('Configuration',$Configuration) | Out-Null
            $Configuration.Jobs["Local Exchange Remote Mailbox"] = [PSCustomObject]@{ Pipe = $powershell; Status = $powershell.BeginInvoke(); Name= "Local Exchange Remote Mailbox" ;Start = Get-Date ; End = ""; process  = [PSCustomObject]@{ Id= ($Configuration.Jobs.Count + 2); Activity = "Processing Local Exchange Remote Mailbox" ; Status = "Starting"; PercentComplete = 0;Completed =$false} } 
        #endregion Local Exchange Remote Mailbox
        #region Local Exchange Remote Mailbox Archive
            $EPRemoteMailboxesArchiveScript = {
                param(
                    $Configuration
                )
                #Import Class ADExchangeOutput
                Invoke-Expression $Configuration.ClassADExchangeOutput
        
                If ((Get-PSSession | Where-Object { $_.ConfigurationName -Match "Microsoft.Exchange" -and $_.ComputerName -eq $Configuration.ExchangeServer}).Count -eq 0 ) {
                    $ERPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://$($Configuration.ExchangeServer)/PowerShell/ -Authentication Negotiate -AllowRedirection 
                    Import-PSSession $ERPSession -AllowClobber -DisableNameChecking -Prefix "EP" | Out-Null
                }
                $ProgressCounter = 1
                $Mailboxes = Get-EPRemoteMailbox -ResultSize unlimited -Archive
                $MailboxesCount = $Mailboxes.Count

                If ($PSVersionTable.PSVersion.Major -ge 7) {
                    $Mailboxes | ForEach-Object -Parallel {
                        $Mailbox = $_
                        $CUPN = $Mailbox.UserPrincipalName
                        $MailboxesCount = $using:MailboxesCount
                        $Configuration = $using:Configuration
                        $ProgressCounter = $using:ProgressCounter

                        #Import Class ADExchangeOutput
                        Invoke-Expression $Configuration.ClassADExchangeOutput

                        If($Mailbox.UserPrincipalName -and  $Configuration.ADUsers[$CUPN]."User Principal Name") {
                            If ( $Configuration.ADUsers[$CUPN]."Mailbox Archive GUID" -ne $Mailbox.ArchiveGuid -and (-Not ($Mailbox.ArchiveGuid -eq [System.Guid]::Empty))) {
                                 $Configuration.ADUsers[$CUPN]."Mailbox Archive GUID" = $Mailbox.ArchiveGuid
                            }
                             $Configuration.ADUsers[$CUPN]."Mailbox Archive State" = $Mailbox.ArchiveState
                             $Configuration.ADUsers[$CUPN]."Exchange Remote Routing Address" = $Mailbox.RemoteRoutingAddress -replace "SMTP:"
                             $Configuration.ADUsers[$CUPN]."Mailbox Archive Location" = "Hybrid Remote"
                            #Update Progress
                            If ( $null -ne $Configuration.Jobs["Local Exchange Remote Mailbox Archive"]) {
                                $Configuration.Jobs["Local Exchange Remote Mailbox Archive"].process.PercentComplete = (($ProgressCounter/$MailboxesCount)*100)
                                $Configuration.Jobs["Local Exchange Remote Mailbox Archive"].process.Status = "Processing Record " + $ProgressCounter + " of " + $MailboxesCount
                            }  
                        }
                        $ProgressCounter++
                    } -ThrottleLimit $Configuration.ProcessBatchJobs
                }Else{
                    ForEach ( $Mailbox in $Mailboxes ) {
                        If($Mailbox.UserPrincipalName -and  $Configuration.ADUsers[$CUPN]."User Principal Name") {
                            If ( $Configuration.ADUsers[$CUPN]."Mailbox Archive GUID" -ne $Mailbox.ArchiveGuid -and (-Not ($Mailbox.ArchiveGuid -eq [System.Guid]::Empty))) {
                                 $Configuration.ADUsers[$CUPN]."Mailbox Archive GUID" = $Mailbox.ArchiveGuid
                            }
                             $Configuration.ADUsers[$CUPN]."Mailbox Archive State" = $Mailbox.ArchiveState
                             $Configuration.ADUsers[$CUPN]."Exchange Remote Routing Address" = $Mailbox.RemoteRoutingAddress -replace "SMTP:"
                             $Configuration.ADUsers[$CUPN]."Mailbox Archive Location" = "Hybrid Remote"
                            #Update Progress
                            If ( $null -ne $Configuration.Jobs["Local Exchange Remote Mailbox Archive"]) {
                                $Configuration.Jobs["Local Exchange Remote Mailbox Archive"].process.PercentComplete = (($ProgressCounter/$MailboxesCount)*100)
                                $Configuration.Jobs["Local Exchange Remote Mailbox Archive"].process.Status = "Processing Record " + $ProgressCounter + " of " + $MailboxesCount
                            }  
                        }
                        $ProgressCounter++
                    }
                }
            }           
            $powershell = [powershell]::Create()
            $powershell.RunspacePool = $runspacepool
            $powershell.AddScript($EPRemoteMailboxesArchiveScript) | Out-Null
            $powershell.AddParameter('Configuration',$Configuration) | Out-Null
            $Configuration.Jobs["Local Exchange Remote Mailbox Archive"] = [PSCustomObject]@{ Pipe = $powershell; Status = $powershell.BeginInvoke(); Name = "Local Exchange Remote Mailbox Archive" ;Start = Get-Date ; End = ""; process  = [PSCustomObject]@{ Id= ($Configuration.Jobs.Count + 2); Activity = "Processing Local Exchange Remote Mailbox Archive" ; Status = "Starting"; PercentComplete = 0;Completed =$false} }        
        #endregion Local Exchange Remote Mailbox Archive
        #region Local Exchange Mailbox Forwarding Rules
            $EPMailboxesForwardingRulesScript = {
                param(
                    $Configuration
                )
                $ProgressCounter = 1
                #Import Class ADExchangeOutput
                Invoke-Expression $Configuration.ClassADExchangeOutput

                If ((Get-PSSession | Where-Object { $_.ConfigurationName -Match "Microsoft.Exchange" -and $_.ComputerName -eq $Configuration.ExchangeServer}).Count -eq 0 ) {
                    $ERPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://$($Configuration.ExchangeServer)/PowerShell/ -Authentication Negotiate -AllowRedirection 
                    Import-PSSession $ERPSession -AllowClobber -DisableNameChecking -Prefix "EP" | Out-Null
                }

                ForEach ( $CUPN in $Configuration.ADUsersUPNsWithEmails) {
                    If( $Configuration.ADUsers[$CUPN]."User Principal Name"){
                        $Rules = Get-EPInboxRule -Mailbox  $Configuration.ADUsers[$CUPN]."User Principal Name" -WarningAction SilentlyContinue | Where-Object {-Not [string]::IsNullOrWhiteSpace($_.ForwardTo)}
                        $ORules = @()
                        ForEach($Rule in $Rules) {
                            If ($Rule.ForwardTo -match '" \[EX:'){
                                $ORules += ($Rule.Name + " = Mailbox: " + (($Rule.ForwardTo  -split "\[")[0] -replace '"'))
                            }Else{
                                If ($Rule.Name + " = " + ($Rule.ForwardTo -split "\[SMTP\:")[-1] -replace "]" -eq " = ") {
                                    $ORules += ($Rule.Name + " = " + ($Rule.ForwardTo -split "\[SMTP\:")[-1] -replace "]")
                                }
                            } 
                        }	
                        If ($ORules.count -gt 0) {
                            $Configuration.ADUsers[$CUPN]."Mailbox Forwarding Rules" = $ORules -join ","
                        }
                        #Update Progress
                        If ( $null -ne $Configuration.Jobs["Local Exchange Mailbox Forwarding Rules"]) {
                            $Configuration.Jobs["Local Exchange Mailbox Forwarding Rules"].process.PercentComplete = (($ProgressCounter/$Configuration.ADUsersUPNsWithEmails.Count)*100)
                            $Configuration.Jobs["Local Exchange Mailbox Forwarding Rules"].process.Status = "Processing Record " + $ProgressCounter + " of " + $Configuration.ADUsersUPNsWithEmails.Count
                        } 
                    }
                    $ProgressCounter++
                }
            }
            $powershell = [powershell]::Create()
            $powershell.RunspacePool = $runspacepool
            $powershell.AddScript($EPMailboxesForwardingRulesScript) | Out-Null
            $powershell.AddParameter('Configuration',$Configuration) | Out-Null
            $Configuration.Jobs["Local Exchange Mailbox Forwarding Rules"] = [PSCustomObject]@{ Pipe = $powershell; Status = $powershell.BeginInvoke(); Name= "Local Exchange Mailbox Forwarding Rules" ;Start = Get-Date ; End = ""; process  = [PSCustomObject]@{ Id= ($Configuration.Jobs.Count + 2); Activity = "Processing Local Exchange Mailbox Forwarding Rules" ; Status = "Starting"; PercentComplete = 0;Completed =$false} }          
        #endregion Local Exchange Mailbox Forwarding Rules
        #region Local Exchange CAS Mailbox Settings
            $EPCASMailboxScript = {
                param(
                    $Configuration
                )
                #Import Class ADExchangeOutput
                Invoke-Expression $Configuration.ClassADExchangeOutput
        
                If ((Get-PSSession | Where-Object { $_.ConfigurationName -Match "Microsoft.Exchange" -and $_.ComputerName -eq $Configuration.ExchangeServer}).Count -eq 0 ) {
                    $ERPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://$($Configuration.ExchangeServer)/PowerShell/ -Authentication Negotiate -AllowRedirection 
                    Import-PSSession $ERPSession -AllowClobber -DisableNameChecking -Prefix "EP" | Out-Null
                }
                $ProgressCounter = 1
                $Mailboxes = Get-EPCASMailbox
                $MailboxesCount = $Mailboxes.Count

                If ($PSVersionTable.PSVersion.Major -ge 7) {
                    $Mailboxes | ForEach-Object -Parallel {
                        $Mailbox = $_
                        $CUPN = $Mailbox.UserPrincipalName
                        $Configuration = $using:Configuration
                        $ProgressCounter = $using:ProgressCounter
                        $MailboxesCount = $using:MailboxesCount
                        #Import Class ADExchangeOutput
                        Invoke-Expression $Configuration.ClassADExchangeOutput
                        
                        If ([string]::IsNullOrEmpty($CUPN)) {
                            #Load .Net Assembly for AD
                            Add-Type -AssemblyName System.DirectoryServices.AccountManagement
                            $ContextType = [System.DirectoryServices.AccountManagement.ContextType]::Domain
                            $IdentityType = [System.DirectoryServices.AccountManagement.IdentityType]::SamAccountName
 
                            $DomainObject = [DirectoryServices.AccountManagement.Principal]::FindByIdentity($ContextType,$IdentityType, $Mailbox.SamAccountName)
                            If ($DomainObject.StructuralObjectClass -eq  "user" -and $DomainObject.UserPrincipalName) {
                                $CUPN = $DomainObject.UserPrincipalName
                            }
                        }

                        If( $Configuration.ADUsers[$CUPN]."User Principal Name" -and $Mailbox.SamAccountName){
        
                            $Configuration.ADUsers[$CUPN]."OWA Enabled" = $Mailbox.OWAEnabled
                            $Configuration.ADUsers[$CUPN]."Mapi Enabled" = $Mailbox.MAPIEnabled
                            $Configuration.ADUsers[$CUPN]."IMAP Enabled" = $Mailbox.IMAPEnabled
                            $Configuration.ADUsers[$CUPN]."Active Sync Enabled" = $Mailbox.ActiveSyncEnabled
                            $Configuration.ADUsers[$CUPN]."POP Enabled" = $Mailbox.POPEnabled
        
                            #  $Configuration.ADUsers[$CUPN]."Mailbox CAS Server" = $Mailbox.ServerName
                            #  $Configuration.ADUsers[$CUPN]."Mailbox CAS Database" = $Mailbox.Database
                            #  $Configuration.ADUsers[$CUPN]."Mailbox CAS OWA for Devices Enabled" = $Mailbox.OWAforDevicesEnabled
                            #  $Configuration.ADUsers[$CUPN]."Mailbox CAS ECP Enabled" = $Mailbox.ECPEnabled
                            #  $Configuration.ADUsers[$CUPN]."Mailbox CAS EWS Enabled" = $Mailbox.EWSEnabled
                            #  $Configuration.ADUsers[$CUPN]."Mailbox CAS OAB Enabled" = $Mailbox.OABEnabled
                            #  $Configuration.ADUsers[$CUPN]."Mailbox CAS ActiveSync Version" = $Mailbox.ActiveSyncVersion
                            #  $Configuration.ADUsers[$CUPN]."Mailbox CAS OWA Version" = $Mailbox.OWAVersion
                            #  $Configuration.ADUsers[$CUPN]."Mailbox CAS ECP Version" = $Mailbox.ECPVersion
                            #  $Configuration.ADUsers[$CUPN]."Mailbox CAS EWS Version" = $Mailbox.EWSVersion
                            #  $Configuration.ADUsers[$CUPN]."Mailbox CAS IMAP Version" = $Mailbox.IMAPVersion
                            #  $Configuration.ADUsers[$CUPN]."Mailbox CAS MAPI Version" = $Mailbox.MAPIVersion
                            #  $Configuration.ADUsers[$CUPN]."Mailbox CAS OAB Version" = $Mailbox.OABVersion
                            #  $Configuration.ADUsers[$CUPN]."Mailbox CAS POP Version" = $Mailbox.POPVersion
                            #  $Configuration.ADUsers[$CUPN]."Mailbox CAS ActiveSync Device ID" = $Mailbox.ActiveSyncDeviceID
                            #  $Configuration.ADUsers[$CUPN]."Mailbox CAS ActiveSync Device Model" = $Mailbox.ActiveSyncDeviceModel
                            #  $Configuration.ADUsers[$CUPN]."Mailbox CAS ActiveSync Device Type" = $Mailbox.ActiveSyncDeviceType
                            #  $Configuration.ADUsers[$CUPN]."Mailbox CAS ActiveSync Device User Agent" = $Mailbox.ActiveSyncDeviceUserAgent
                            #  $Configuration.ADUsers[$CUPN]."Mailbox CAS ActiveSync Device Access State" = $Mailbox.ActiveSyncDeviceAccessState
                            
                            #Update Progress
                            $wprogress = ($Configuration.jobs.Where({ $_.Name -eq "Local Exchange CAS Mailbox Settings" })).process
                            If ($null -ne $wprogress) {
                                $wprogress.PercentComplete = (($ProgressCounter/$MailboxesCount)*100)
                                $wprogress.Status = "Processing Record " + $ProgressCounter + " of " + $MailboxesCount
                            }
                        }
                        $ProgressCounter++
                    } -ThrottleLimit $Configuration.ProcessBatchJobs
                }Else{
                    ForEach ( $Mailbox in $Mailboxes) {
                        $CUPN = $Mailbox.UserPrincipalName
                        If ([string]::IsNullOrEmpty($CUPN)) {
                            #Load .Net Assembly for AD
                            Add-Type -AssemblyName System.DirectoryServices.AccountManagement
                            $ContextType = [System.DirectoryServices.AccountManagement.ContextType]::Domain
                            $IdentityType = [System.DirectoryServices.AccountManagement.IdentityType]::SamAccountName
 
                            $DomainObject = [DirectoryServices.AccountManagement.Principal]::FindByIdentity($ContextType,$IdentityType, $Mailbox.SamAccountName)
                            If ($DomainObject.StructuralObjectClass -eq  "user" -and $DomainObject.UserPrincipalName) {
                                $CUPN = $DomainObject.UserPrincipalName
                            }
                        }
                        If( $Configuration.ADUsers[$CUPN]."User Principal Name" -and $Mailbox.SamAccountName){
        
                            $Configuration.ADUsers[$CUPN]."OWA Enabled" = $Mailbox.OWAEnabled
                            $Configuration.ADUsers[$CUPN]."Mapi Enabled" = $Mailbox.MAPIEnabled
                            $Configuration.ADUsers[$CUPN]."IMAP Enabled" = $Mailbox.IMAPEnabled
                            $Configuration.ADUsers[$CUPN]."Active Sync Enabled" = $Mailbox.ActiveSyncEnabled
                            $Configuration.ADUsers[$CUPN]."POP Enabled" = $Mailbox.POPEnabled
        
                            #  $Configuration.ADUsers[$CUPN]."Mailbox CAS Server" = $Mailbox.ServerName
                            #  $Configuration.ADUsers[$CUPN]."Mailbox CAS Database" = $Mailbox.Database
                            #  $Configuration.ADUsers[$CUPN]."Mailbox CAS OWA for Devices Enabled" = $Mailbox.OWAforDevicesEnabled
                            #  $Configuration.ADUsers[$CUPN]."Mailbox CAS ECP Enabled" = $Mailbox.ECPEnabled
                            #  $Configuration.ADUsers[$CUPN]."Mailbox CAS EWS Enabled" = $Mailbox.EWSEnabled
                            #  $Configuration.ADUsers[$CUPN]."Mailbox CAS OAB Enabled" = $Mailbox.OABEnabled
                            #  $Configuration.ADUsers[$CUPN]."Mailbox CAS ActiveSync Version" = $Mailbox.ActiveSyncVersion
                            #  $Configuration.ADUsers[$CUPN]."Mailbox CAS OWA Version" = $Mailbox.OWAVersion
                            #  $Configuration.ADUsers[$CUPN]."Mailbox CAS ECP Version" = $Mailbox.ECPVersion
                            #  $Configuration.ADUsers[$CUPN]."Mailbox CAS EWS Version" = $Mailbox.EWSVersion
                            #  $Configuration.ADUsers[$CUPN]."Mailbox CAS IMAP Version" = $Mailbox.IMAPVersion
                            #  $Configuration.ADUsers[$CUPN]."Mailbox CAS MAPI Version" = $Mailbox.MAPIVersion
                            #  $Configuration.ADUsers[$CUPN]."Mailbox CAS OAB Version" = $Mailbox.OABVersion
                            #  $Configuration.ADUsers[$CUPN]."Mailbox CAS POP Version" = $Mailbox.POPVersion
                            #  $Configuration.ADUsers[$CUPN]."Mailbox CAS ActiveSync Device ID" = $Mailbox.ActiveSyncDeviceID
                            #  $Configuration.ADUsers[$CUPN]."Mailbox CAS ActiveSync Device Model" = $Mailbox.ActiveSyncDeviceModel
                            #  $Configuration.ADUsers[$CUPN]."Mailbox CAS ActiveSync Device Type" = $Mailbox.ActiveSyncDeviceType
                            #  $Configuration.ADUsers[$CUPN]."Mailbox CAS ActiveSync Device User Agent" = $Mailbox.ActiveSyncDeviceUserAgent
                            #  $Configuration.ADUsers[$CUPN]."Mailbox CAS ActiveSync Device Access State" = $Mailbox.ActiveSyncDeviceAccessState

                            #Update Progress
                            If ( $null -ne $Configuration.Jobs["Local Exchange CAS Mailbox Settings"]) {
                                $Configuration.Jobs["Local Exchange CAS Mailbox Settings"].process.PercentComplete = (($ProgressCounter/$MailboxesCount)*100)
                                $Configuration.Jobs["Local Exchange CAS Mailbox Settings"].process.Status = "Processing Record " + $ProgressCounter + " of " + $MailboxesCount
                            }  
                        }
                        $ProgressCounter++
                    }
                }
            }
            $powershell = [powershell]::Create()
            $powershell.RunspacePool = $runspacepool
            $powershell.AddScript($EPCASMailboxScript) | Out-Null
            $powershell.AddParameter('Configuration',$Configuration) | Out-Null
            $Configuration.Jobs["Local Exchange CAS Mailbox Settings"] = [PSCustomObject]@{ Pipe = $powershell; Status = $powershell.BeginInvoke(); Name= "Local Exchange CAS Mailbox Settings" ;Start = Get-Date ; End = ""; process  = [PSCustomObject]@{ Id= ($Configuration.Jobs.Count + 2); Activity = "Processing Local Exchange CAS Mailbox Settings" ; Status = "Starting"; PercentComplete = 0;Completed =$false} }
        #endregion Local Exchange CAS Mailbox Settings
        #region Local Exchange Mobile Device Settings
            $EPMobileMailboxScript = {
                param(
                    $Configuration
                )
                #Import Class ADExchangeOutput
                Invoke-Expression $Configuration.ClassADExchangeOutput
                $ProgressCounter = 1
                #region Connect to Exchange
                If ((Get-PSSession | Where-Object { $_.ConfigurationName -Match "Microsoft.Exchange" -and $_.ComputerName -eq $Configuration.ExchangeServer}).Count -eq 0 ) {
                    $ERPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://$($Configuration.ExchangeServer)/PowerShell/ -Authentication Negotiate -AllowRedirection 
                    Import-PSSession $ERPSession -AllowClobber -DisableNameChecking -Prefix "EP" | Out-Null
                }
                #endregion Connect to Exchange
                ForEach ( $CUPN in $Configuration.ADUsersWithMailboxGUIDs ) {
                    $MobileDevices = Get-EPMobileDevice -Mailbox $Configuration.ADUsers[$CUPN]."Mailbox GUID" -ResultSize Unlimited | Where-Object {$_.DeviceAccessState -eq "Allowed"} | Sort-Object -Property WhenCreated -Descending

                    If( $Configuration.ADUsers[$CUPN]."User Principal Name" -and $MobileDevices.Count -gt 0){
                        $MMDClean = @()
                        ForEach ($Device in $MobileDevices) {
                            $DeviceString= ""
                            If (-Not [string]::IsNullOrWhiteSpace($Device.FriendlyName)) {
                                $DeviceString =  $Device.FriendlyName
                            }
                            If (-Not [string]::IsNullOrWhiteSpace($Device.DeviceModel)) {
                                If ($DeviceString) {
                                    $DeviceString = ($DeviceString + " - " + $Device.DeviceModel)
                                }else{
                                    $DeviceString = $Device.DeviceModel
                                }
                            }
                            If (-Not [string]::IsNullOrWhiteSpace($Device.DeviceId)) {
                                If ($DeviceString) {
                                    $DeviceString = ($DeviceString + " - " + $Device.DeviceId)
                                }else{
                                    $DeviceString = $Device.DeviceId
                                }
                            }
                            If ($DeviceString) {
                                $MMDClean += $DeviceString
                            }
                        }
                        If($MMDClean.count -gt 0) {
                                $Configuration.ADUsers[$CUPN]."Exchange Mobile Devices" = $MMDClean -join ","
                        }
                        #Update Progress
                        If ( $null -ne $Configuration.Jobs["Local Exchange Mobile Device Settings"]) {
                            $Configuration.Jobs["Local Exchange Mobile Device Settings"].process.PercentComplete = (($ProgressCounter/$Configuration.ADUsersWithMailboxGUIDs.Count)*100)
                            $Configuration.Jobs["Local Exchange Mobile Device Settings"].process.Status = "Processing Record " + $ProgressCounter + " of " + $Configuration.ADUsersWithMailboxGUIDs.Count
                        }
                    }
                    $ProgressCounter++
                }
            
                $Configuration.Jobs["Local Exchange Mobile Device Settings"].process.PercentComplete = 100
                $Configuration.Jobs["Local Exchange Mobile Device Settings"].Process.Status = "Completed"
                # $Configuration.Jobs["Local Exchange Mobile Device Settings"].End = Get-Date
                $Configuration.Jobs["Local Exchange Mobile Device Settings"].process.completed = $true
            }
            $powershell = [powershell]::Create()
            $powershell.RunspacePool = $runspacepool
            $powershell.AddScript($EPMobileMailboxScript) | Out-Null
            $powershell.AddParameter('Configuration',$Configuration) | Out-Null
            $Configuration.Jobs["Local Exchange Mobile Device Settings"] = [PSCustomObject]@{ Pipe = $powershell; Status = $powershell.BeginInvoke(); Name= "Local Exchange Mobile Device Settings" ;Start = Get-Date ; End = ""; process  = [PSCustomObject]@{ Id= ($Configuration.Jobs.Count + 2); Activity = "Processing Local Exchange Mobile Device Settings" ; Status = "Starting"; PercentComplete = 0;Completed =$false} }
        #endregion Local Exchange Mobile Device Settings
    #endregion Local Exchange Info
    #region Online Exchange Info
        #region Online Exchange Mailboxes
            $EXOMailboxesScript = {
                param(
                    $Configuration
                )
                Invoke-Expression $Configuration.ClassADExchangeOutput
                #region Connect to Exchange Online
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
                    If (Get-ChildItem "Cert:\LocalMachine\My\" | Where-Object {$_.Thumbprint -eq  $Configuration.AZCertThumbprint}) {
                        Connect-ExchangeOnline -CertificateThumbPrint  $Configuration.AZCertThumbprint -AppID  $Configuration.AZClientID -Organization  $Configuration.AZOrg -ShowProgress:$false -ShowBanner:$false
                    }Else{
                        Connect-ExchangeOnline -ShowBanner:$false -UserPrincipalName  $Configuration.CurrentUserUPN -ShowProgress:$false
                    }
                }
                #endregion Connect to Exchange Online
                $ProgressCounter = 1
                $Mailboxes = Get-EXOMailbox -ResultSize unlimited -PropertySets All
                $MailboxesCount = $Mailboxes.Count

                If ($PSVersionTable.PSVersion.Major -ge 7) {
                    $Mailboxes | ForEach-Object -Parallel {
                        $Mailbox = $_
                        $CUPN = $Mailbox.UserPrincipalName
                        $MailboxesCount = $using:MailboxesCount
                        $Configuration = $using:Configuration
                        $ProgressCounter = $using:ProgressCounter

                        #Import Class ADExchangeOutput
                        Invoke-Expression $Configuration.ClassADExchangeOutput

                        If( $Configuration.ADUsers[$CUPN]."User Principal Name" -and $Mailbox.UserPrincipalName){
                            If( $Configuration.ADUsers[$CUPN]."Online Mailbox GUID" -ne $Mailbox.ExchangeGuid -and (-Not ($Mailbox.ExchangeGuid -eq [System.Guid]::Empty))) {
                                 $Configuration.ADUsers[$CUPN]."Online Mailbox GUID" = $Mailbox.ExchangeGuid
                            }

                            If( $Configuration.ADUsers[$CUPN]."Mailbox Location" -ne "Hybrid Remote") {
                                 $Configuration.ADUsers[$CUPN]."Mailbox Location" = "Online"
                                If ( $Configuration.ADUsers[$CUPN]."Mailbox GUID" -ne $Mailbox.ExchangeGuid -and (-Not ($Mailbox.ExchangeGuid -eq [System.Guid]::Empty))) {
                                    Write-Warning ("Online Mailbox GUID Mismatch: " + $Mailbox.UserPrincipalName + " Online GUID: " + $Mailbox.ExchangeGuid + " Local GUID: " +  $Configuration.ADUsers[$CUPN]."Mailbox GUID")
                                }
                            }
                            #Use MailboxLocations to get Mailbox Server and Archive Server
                            $MailboxLocation = ($Mailbox.MailboxLocations).Split(";")
                            For ($i=0; $i -lt $MailboxLocation.Count; $i++) {
                                If ($MailboxLocation[$i] -match "Primary") {
                                    If($null -eq  $Configuration.ADUsers[$CUPN]."Mailbox Server") {
                                            $Configuration.ADUsers[$CUPN]."Mailbox Server" = $MailboxLocation[($i + 1)]
                                    }
                                    If($null -eq  $Configuration.ADUsers[$CUPN]."Online Mailbox GUID") {
                                        $Configuration.ADUsers[$CUPN]."Online Mailbox GUID" = $MailboxLocation[($i -1)]
                                    }
                                }
                                If($MailboxLocation[$i] -match "MainArchive") {
                                    If($null -eq  $Configuration.ADUsers[$CUPN]."Mailbox Archive Server") {
                                            $Configuration.ADUsers[$CUPN]."Mailbox Archive Server" = $MailboxLocation[($i + 1)]
                                    }
                                    If($null -eq  $Configuration.ADUsers[$CUPN]."Online Mailbox Archive GUID") {
                                        $Configuration.ADUsers[$CUPN]."Online Mailbox Archive GUID" = $MailboxLocation[($i -1)]
                                    }
                                    
                                }
                            }
                            
                            If($null -eq  $Configuration.ADUsers[$CUPN]."Mailbox Database") {
                                 $Configuration.ADUsers[$CUPN]."Mailbox Database" = $Mailbox.Database
                            }
                            
                             $Configuration.ADUsers[$CUPN]."Mailbox Issue Warning Quota" = Format-MailboxGB($Mailbox.IssueWarningQuota)
                             $Configuration.ADUsers[$CUPN]."Mailbox Prohibit Send Quota" = Format-MailboxGB($Mailbox.ProhibitSendQuota)
                             $Configuration.ADUsers[$CUPN]."Mailbox Prohibit Send Receive Quota" = Format-MailboxGB($Mailbox.ProhibitSendReceiveQuota)
                             $Configuration.ADUsers[$CUPN]."Mailbox Use Database Quota Defaults" = $Mailbox.UseDatabaseQuotaDefaults

                             $Configuration.ADUsers[$CUPN]."Mailbox Forwarding Address" = $Mailbox.ForwardingAddress
                             $Configuration.ADUsers[$CUPN]."Mailbox Forwarding Address SMTP" = $Mailbox.ForwardingSmtpAddress
                            
                            If(($Mailbox.RecipientTypeDetails -ne "UserMailbox" -or $Mailbox.RecipientTypeDetails -ne "SharedMailbox") -and ( $Configuration.ADUsers[$CUPN]."Exchange Recipient Type" -ne "RemoteUserMailbox" -or  $Configuration.ADUsers[$CUPN]."Exchange Recipient Type" -ne "Migrated")) {
                                 $Configuration.ADUsers[$CUPN]."Exchange Recipient Type" = $Mailbox.RemoteRecipientType 
                                Write-warning ("Online Recipient Type Mismatch: " + $Mailbox.UserPrincipalName + " Online Recipient Type: " + $Mailbox.RecipientTypeDetails + " Local Recipient Type: " +  $Configuration.ADUsers[$CUPN]."Exchange Recipient Type")
                            }

                            #Update Progress
                            If ($null -ne $Configuration.Jobs["Online Exchange Mailbox"]) {
                                $Configuration.Jobs["Online Exchange Mailbox"].process.PercentComplete = (($ProgressCounter/$MailboxesCount)*100)
                                $Configuration.Jobs["Online Exchange Mailbox"].process.Status = "Processing Record " + $ProgressCounter + " of " + $MailboxesCount
                            }
                        }
                        $ProgressCounter++
                    } -ThrottleLimit $Configuration.ProcessBatchJobs
                }Else{
                    ForEach ($Mailbox in $Mailboxes ) {
                        If( $Configuration.ADUsers[$CUPN]."User Principal Name" -and $Mailbox.UserPrincipalName){
                            If( $Configuration.ADUsers[$CUPN]."Online Mailbox GUID" -ne $Mailbox.ExchangeGuid -and (-Not ($Mailbox.ExchangeGuid -eq [System.Guid]::Empty))) {
                                 $Configuration.ADUsers[$CUPN]."Online Mailbox GUID" = $Mailbox.ExchangeGuid
                            }

                            If( $Configuration.ADUsers[$CUPN]."Mailbox Location" -ne "Hybrid Remote") {
                                 $Configuration.ADUsers[$CUPN]."Mailbox Location" = "Online"
                                If ( $Configuration.ADUsers[$CUPN]."Mailbox GUID" -ne $Mailbox.ExchangeGuid -and (-Not ($Mailbox.ExchangeGuid -eq [System.Guid]::Empty))) {
                                    Write-Warning ("Online Mailbox GUID Mismatch: " + $Mailbox.UserPrincipalName + " Online GUID: " + $Mailbox.ExchangeGuid + " Local GUID: " +  $Configuration.ADUsers[$CUPN]."Mailbox GUID")
                                }
                            }

                            #Use MailboxLocations to get Mailbox Server and Archive Server
                            $MailboxLocation = ($Mailbox.MailboxLocations).Split(";")
                            For ($i=0; $i -lt $MailboxLocation.Count; $i++) {
                                If ($MailboxLocation[$i] -match "Primary") {
                                    If($null -eq  $Configuration.ADUsers[$CUPN]."Mailbox Server") {
                                            $Configuration.ADUsers[$CUPN]."Mailbox Server" = $MailboxLocation[($i + 1)]
                                    }
                                    If($null -eq  $Configuration.ADUsers[$CUPN]."Online Mailbox GUID") {
                                        $Configuration.ADUsers[$CUPN]."Online Mailbox GUID" = $MailboxLocation[($i -1)]
                                    }
                                }
                                If($MailboxLocation[$i] -match "MainArchive") {
                                    If($null -eq  $Configuration.ADUsers[$CUPN]."Mailbox Archive Server") {
                                            $Configuration.ADUsers[$CUPN]."Mailbox Archive Server" = $MailboxLocation[($i + 1)]
                                    }
                                    If($null -eq  $Configuration.ADUsers[$CUPN]."Online Mailbox Archive GUID") {
                                        $Configuration.ADUsers[$CUPN]."Online Mailbox Archive GUID" = $MailboxLocation[($i -1)]
                                    }
                                    
                                }
                            }
                            If($null -eq  $Configuration.ADUsers[$CUPN]."Mailbox Database") {
                                 $Configuration.ADUsers[$CUPN]."Mailbox Database" = $Mailbox.Database
                            }
                            
                             $Configuration.ADUsers[$CUPN]."Mailbox Issue Warning Quota" = Format-MailboxGB($Mailbox.IssueWarningQuota)
                             $Configuration.ADUsers[$CUPN]."Mailbox Prohibit Send Quota" = Format-MailboxGB($Mailbox.ProhibitSendQuota)
                             $Configuration.ADUsers[$CUPN]."Mailbox Prohibit Send Receive Quota" = Format-MailboxGB($Mailbox.ProhibitSendReceiveQuota)
                             $Configuration.ADUsers[$CUPN]."Mailbox Use Database Quota Defaults" = $Mailbox.UseDatabaseQuotaDefaults

                             $Configuration.ADUsers[$CUPN]."Mailbox Forwarding Address" = $Mailbox.ForwardingAddress
                             $Configuration.ADUsers[$CUPN]."Mailbox Forwarding Address SMTP" = $Mailbox.ForwardingSmtpAddress
                            
                            If($Mailbox.RecipientTypeDetails -ne "UserMailbox" -and ( $Configuration.ADUsers[$CUPN]."Exchange Recipient Type" -ne "RemoteUserMailbox" -or  $Configuration.ADUsers[$CUPN]."Exchange Recipient Type" -ne "Migrated")) {
                                 $Configuration.ADUsers[$CUPN]."Exchange Recipient Type" = $Mailbox.RemoteRecipientType 
                                Write-warning ("Online Recipient Type Mismatch: " + $Mailbox.UserPrincipalName + " Online Recipient Type: " + $Mailbox.RecipientTypeDetails + " Local Recipient Type: " +  $Configuration.ADUsers[$CUPN]."Exchange Recipient Type")
                            }
                            #Update Progress
                            If ($null -ne $Configuration.Jobs["Online Exchange Mailbox"]) {
                                $Configuration.Jobs["Online Exchange Mailbox"].process.PercentComplete = (($ProgressCounter/$MailboxesCount)*100)
                                $Configuration.Jobs["Online Exchange Mailbox"].process.Status = "Processing Record " + $ProgressCounter + " of " + $MailboxesCount
                            }
                        }
                        $ProgressCounter++
                    }
                }
	        }
            $powershell = [powershell]::Create()
            $powershell.RunspacePool = $runspacepool
            $powershell.AddScript($EXOMailboxesScript) | Out-Null
            $powershell.AddParameter('Configuration',$Configuration) | Out-Null
            $Configuration.Jobs["Online Exchange Mailbox"] = [PSCustomObject]@{ Pipe = $powershell; Status = $powershell.BeginInvoke(); Name= "Online Exchange Mailbox" ;Start = Get-Date ; End = ""; process  = [PSCustomObject]@{ Id= ($Configuration.Jobs.Count + 2); Activity = "Processing Online Exchange Mailbox" ; Status = "Starting"; PercentComplete = 0;Completed =$false} }
        #endregion Online Exchange Mailboxes
        #region Online Exchange Mailbox Statistics
            $EXOMailboxesScriptStatistics = {
                param(
                    $Configuration
                )
                $ProgressCounter = 1
                #Import Class ADExchangeOutput
                Invoke-Expression $Configuration.ClassADExchangeOutput
                #region Connect to Exchange Online
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
                    If (Get-ChildItem "Cert:\LocalMachine\My\" | Where-Object {$_.Thumbprint -eq  $Configuration.AZCertThumbprint}) {
                        Connect-ExchangeOnline -CertificateThumbPrint  $Configuration.AZCertThumbprint -AppID  $Configuration.AZClientID -Organization  $Configuration.AZOrg -ShowProgress:$false -ShowBanner:$false
                    }Else{
                        Connect-ExchangeOnline -ShowBanner:$false -UserPrincipalName  $Configuration.CurrentUserUPN -ShowProgress:$false
                    }
                }
                #endregion Connect to Exchange Online

                $ProgressCounter = 1
                $Mailboxes = Get-EXOMailbox -ResultSize unlimited -PropertySets All
                $MailboxesCount = $Mailboxes.Count
                
                ForEach ( $Mailbox in $Mailboxes) {
                    $CUPN = $Mailbox.UserPrincipalName
                    $MailboxStats = $Mailbox | Get-EXOMailboxStatistics -Properties TotalItemSize,ItemCount,LastLoggedOnUserAccount,LastLogonTime,LastLogoffTime
                  
                    If($Configuration.ADUsers[$CUPN]."User Principal Name" -and $MailboxStats.LastLogonTime){
                        $Configuration.ADUsers[$CUPN]."Mailbox Size GB" = Format-MailboxGB ($MailboxStats.TotalItemSize)
                        $Configuration.ADUsers[$CUPN]."Mailbox Item Count" = $MailboxStats.ItemCount
                        $Configuration.ADUsers[$CUPN]."Mailbox Last Logged On User Account" = $MailboxStats.LastLoggedOnUserAccount
                        $Configuration.ADUsers[$CUPN]."Mailbox Last Logon Time" = $MailboxStats.LastLogonTime
                        $Configuration.ADUsers[$CUPN]."Mailbox Last Logoff Time" = $MailboxStats.LastLogoffTime
                        If (-Not ([string]::IsNullOrWhiteSpace($Configuration.ADUsers[$CUPN].'Mailbox Prohibit Send Quota')) -and $Configuration.ADUsers[$CUPN].'Mailbox Prohibit Send Quota' -ne "Unlimited" -and $Configuration.ADUsers[$CUPN].'Mailbox Prohibit Send Quota' -ne 0) {
                            $Configuration.ADUsers[$CUPN]."Mailbox Usage %" = [math]::Round((($TotalItemSize + $TotalDeletedItemSize)/$Configuration.ADUsers[$CUPN].'Mailbox Prohibit Send Quota')*100,2)
                        }
                    }
                   
                    #Update Progress
                    If ( $null -ne $Configuration.Jobs["Online Exchange Mailbox Statistics"]) {
                        $Configuration.Jobs["Online Exchange Mailbox Statistics"].process.PercentComplete = (($ProgressCounter/$MailboxesCount)*100)
                        $Configuration.Jobs["Online Exchange Mailbox Statistics"].process.Status = "Processing Record " + $ProgressCounter + " of " + $MailboxesCount
                    }
                    $ProgressCounter++
                }
            }
            $powershell = [powershell]::Create()
            $powershell.RunspacePool = $runspacepool
            $powershell.AddScript($EXOMailboxesScriptStatistics) | Out-Null
            $powershell.AddParameter('Configuration',$Configuration) | Out-Null
            $Configuration.Jobs["Online Exchange Mailbox Statistics"] = [PSCustomObject]@{ Pipe = $powershell; Status = $powershell.BeginInvoke(); Name= "Online Exchange Mailbox Statistics" ;Start = Get-Date ; End = ""; process  = [PSCustomObject]@{ Id= ($Configuration.Jobs.Count + 2); Activity = "Processing Online Exchange Mailbox Statistics" ; Status = "Starting"; PercentComplete = 0;Completed =$false} }   
        #endregion Online Exchange Mailbox Statistics
        #region Online Exchange Mailbox Folder Statistics
            $EXOMailboxesScriptFolderStatistics = {
                param(
                    $Configuration
                )
                $ProgressCounter = 1
                #Import Class ADExchangeOutput
                Invoke-Expression $Configuration.ClassADExchangeOutput
                #region Connect to Exchange Online
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
                    If (Get-ChildItem "Cert:\LocalMachine\My\" | Where-Object {$_.Thumbprint -eq  $Configuration.AZCertThumbprint}) {
                        Connect-ExchangeOnline -CertificateThumbPrint  $Configuration.AZCertThumbprint -AppID  $Configuration.AZClientID -Organization  $Configuration.AZOrg -ShowProgress:$false -ShowBanner:$false
                    }Else{
                        Connect-ExchangeOnline -ShowBanner:$false -UserPrincipalName  $Configuration.CurrentUserUPN -ShowProgress:$false
                    }
                }
                #endregion Connect to Exchange Online

                $ProgressCounter = 1
                $Mailboxes = Get-EXOMailbox -ResultSize unlimited -PropertySets All
                $MailboxesCount = $Mailboxes.Count
                
                ForEach ( $Mailbox in $Mailboxes) {
                    $CUPN = $Mailbox.UserPrincipalName
                    $RIMailboxStats = $Mailbox | Get-MailboxFolderStatistics -FolderScope RecoverableItems | Where-Object{$_.Name -eq "Recoverable Items"}
                
                    If($Configuration.ADUsers[$CUPN]."User Principal Name" -and $RIMailboxStats.FolderAndSubfolderSize){
                        $Configuration.ADUsers[$CUPN]."Mailbox Recoverable Items Size GB" = Format-MailboxGB ($RIMailboxStats.FolderAndSubfolderSize)
                       
                    }
                
                    #Update Progress
                    If ( $null -ne $Configuration.Jobs["Online Exchange Folder Mailbox Statistics"]) {
                        $Configuration.Jobs["Online Exchange Folder Mailbox Statistics"].process.PercentComplete = (($ProgressCounter/$MailboxesCount)*100)
                        $Configuration.Jobs["Online Exchange Folder Mailbox Statistics"].process.Status = "Processing Record " + $ProgressCounter + " of " + $MailboxesCount
                    }
                    $ProgressCounter++
                }
            }
            $powershell = [powershell]::Create()
            $powershell.RunspacePool = $runspacepool
            $powershell.AddScript($EXOMailboxesScriptStatistics) | Out-Null
            $powershell.AddParameter('Configuration',$Configuration) | Out-Null
            $Configuration.Jobs["Online Exchange Folder Mailbox Statistics"] = [PSCustomObject]@{ Pipe = $powershell; Status = $powershell.BeginInvoke(); Name= "Online Exchange Folder Mailbox Statistics" ;Start = Get-Date ; End = ""; process  = [PSCustomObject]@{ Id= ($Configuration.Jobs.Count + 2); Activity = "Processing Online Exchange Folder Mailbox Statistics" ; Status = "Starting"; PercentComplete = 0;Completed =$false} }   
        #endregion Online Exchange Mailbox Folder Statistics        
        #region Online Exchange Mailbox Permission
            $EXOMailboxesPermissionsScript = {
                param(
                    $Configuration
                )
                $ProgressCounter = 1
                #Import Class ADExchangeOutput
                Invoke-Expression $Configuration.ClassADExchangeOutput
                #region Connect to Exchange Online
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
                    If (Get-ChildItem "Cert:\LocalMachine\My\" | Where-Object {$_.Thumbprint -eq  $Configuration.AZCertThumbprint}) {
                        Connect-ExchangeOnline -CertificateThumbPrint  $Configuration.AZCertThumbprint -AppID  $Configuration.AZClientID -Organization  $Configuration.AZOrg -ShowProgress:$false -ShowBanner:$false
                    }Else{
                        Connect-ExchangeOnline -ShowBanner:$false -UserPrincipalName  $Configuration.CurrentUserUPN -ShowProgress:$false
                    }
                }
                #endregion Connect to Exchange Online
                #Load .Net Assembly for AD
                Add-Type -AssemblyName System.DirectoryServices.AccountManagement
                $ContextType = [System.DirectoryServices.AccountManagement.ContextType]::Domain
                $IdentityType = [System.DirectoryServices.AccountManagement.IdentityType]::SamAccountName
                ForEach ($CUPN in $Configuration.ADUsersUPNsWithERRA) {
                    $EPMailboxesPerms =  Get-EXOMailboxPermission -ResultSize Unlimited -identity $CUPN -ErrorAction SilentlyContinue
                    #region Mailbox Permission
                    $DAMP = @()
                    $FixedDAMP = @{}
                    $CSVFixedDAMP = ""
                    If ($EPMailboxesPerms.AccessRights -gt 0) {
                        $DAMP = ($EPMailboxesPerms.Where({($_.AccessRights -eq "FullAccess") -and ($_.User -notin $Configuration.ExcludeUsers) -and ($_.User -notmatch "S-1-5-*" ) -and $_.IsInherited -eq $false})) | Sort-Object -Unique -Property User	
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
                                                    If ($ADUser.description -notmatch "Leave") {
                                                        $FixedDAMP.add(($ACE.User + " - Leave"),$ACE.AccessRights)
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
                                    $Configuration.ADUsers[$CUPN]."Mailbox Permissions" = $CSVFixedDAMP
                                }
                            }
                        }	
                        #endregion Mailbox Permission
                        #Update Progress
                        If ($null -ne $Configuration.Jobs["Online Exchange Mailbox Permission"]) {
                            $Configuration.Jobs["Online Exchange Mailbox Permission"].process.PercentComplete = (($ProgressCounter/$Configuration.ADUsersUPNsWithERRA.Count)*100)
                            $Configuration.Jobs["Online Exchange Mailbox Permission"].process.Status = "Processing Record " + $ProgressCounter + " of " + $Configuration.ADUsersUPNsWithERRA.Count
                        }
                    }
                    $ProgressCounter++
                }
            
            }
            $powershell = [powershell]::Create()
            $powershell.RunspacePool = $runspacepool
            $powershell.AddScript($EXOMailboxesPermissionsScript) | Out-Null
            $powershell.AddParameter('Configuration',$Configuration) | Out-Null
            $Configuration.Jobs["Online Exchange Mailbox Permission"] = [PSCustomObject]@{ Pipe = $powershell; Status = $powershell.BeginInvoke(); Name= "Online Exchange Mailbox Permission" ;Start = Get-Date ; End = ""; process  = [PSCustomObject]@{ Id= ($Configuration.Jobs.Count + 2); Activity = "Processing Online Exchange Mailbox Permission" ; Status = "Starting"; PercentComplete = 0;Completed =$false} }   
        #endregion Online Exchange Mailbox Permission
        #region Online Exchange Mailbox Archive
            $EXOMailboxesArchiveScript = {
                param(
                    $Configuration
                )
                #Import Class ADExchangeOutput
                Invoke-Expression $Configuration.ClassADExchangeOutput
                #region Connect to Exchange Online
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
                    If (Get-ChildItem "Cert:\LocalMachine\My\" | Where-Object {$_.Thumbprint -eq  $Configuration.AZCertThumbprint}) {
                        Connect-ExchangeOnline -CertificateThumbPrint  $Configuration.AZCertThumbprint -AppID  $Configuration.AZClientID -Organization  $Configuration.AZOrg -ShowProgress:$false -ShowBanner:$false
                    }Else{
                        Connect-ExchangeOnline -ShowBanner:$false -UserPrincipalName  $Configuration.CurrentUserUPN -ShowProgress:$false
                    }
                }
                #endregion Connect to Exchange Online
                
                # Get-EXOMailbox -ResultSize unlimited -Archive -PropertySets All
                $Mailboxes = Get-EXOMailbox -ResultSize unlimited -Archive -PropertySets All
                $ProgressCounter = 1
                $MailboxesCount = $Mailboxes.Count
                If ($PSVersionTable.PSVersion.Major -ge 7) {
                    $Mailboxes | ForEach-Object -Parallel {
                        $Mailbox = $_
                        $CUPN = $Mailbox.UserPrincipalName
                        $Configuration = $using:Configuration
                        $ProgressCounter = $using:ProgressCounter
                        $MailboxesCount = $using:MailboxesCount
                        #Import Class ADExchangeOutput
                        Invoke-Expression $Configuration.ClassADExchangeOutput

                        If($Configuration.ADUsers[$CUPN]."User Principal Name" -and $Mailbox.UserPrincipalName) {
                            #Use MailboxLocations to get Mailbox Archive Server
                            $MailboxLocation = ($Mailbox.MailboxLocations).Split(";")
                            For ($i=0; $i -lt $MailboxLocation.Count; $i++) {
                                If($MailboxLocation[$i] -match "MainArchive") {
                                    If($null -eq  $Configuration.ADUsers[$CUPN]."Mailbox Archive Server") {
                                            $Configuration.ADUsers[$CUPN]."Mailbox Archive Server" = $MailboxLocation[($i + 1)]
                                    }
                                    If($null -eq  $Configuration.ADUsers[$CUPN]."Online Mailbox Archive GUID") {
                                        $Configuration.ADUsers[$CUPN]."Online Mailbox Archive GUID" = $MailboxLocation[($i -1)]
                                    }
                                }
                            }
                            $Configuration.ADUsers[$CUPN]."Mailbox Archive Database" = $Mailbox.Database
                            If($Configuration.ADUsers[$CUPN]."Mailbox Archive Location" -ne "Hybrid Remote") {

                                If($Configuration.ADUsers[$CUPN]."Mailbox Archive GUID" -ne $Mailbox.ArchiveGuid -and ($Mailbox.ArchiveGuid -ne [System.Guid]::Empty)) {
                                    Write-Warning ("Online Mailbox Archive GUID Mismatch: " + $Mailbox.UserPrincipalName + " Online GUID: " + $Mailbox.ArchiveGuid + " Local GUID: " + $Configuration.ADUsers[$CUPN]."Mailbox Archive GUID")
                                }
                            }
                            If($Mailbox.ArchiveGuid -and ($Mailbox.ArchiveGuid -ne [System.Guid]::Empty)) {
                                $Configuration.ADUsers[$CUPN]."Online Mailbox Archive GUID" = $Mailbox.ArchiveGuid
                            }
                            $Configuration.ADUsers[$CUPN]."Mailbox Archive Creation Date" = $Mailbox.WhenCreated
                            $Configuration.ADUsers[$CUPN]."Mailbox Archive Status" = $Mailbox.ArchiveStatus
                            $Configuration.ADUsers[$CUPN]."Mailbox Archive State" = $Mailbox.ArchiveState
                            If($Mailbox.AutoExpandingArchive) {
                                $Configuration.ADUsers[$CUPN]."Mailbox Archive Auto Expanding Archive" = $Mailbox.AutoExpandingArchive
                            }
                            $Configuration.ADUsers[$CUPN]."Mailbox Archive Issue Warning Quota" = Format-MailboxGB($Mailbox.IssueWarningQuota)
                            $Configuration.ADUsers[$CUPN]."Mailbox Archive Prohibit Send Quota" = Format-MailboxGB($Mailbox.ProhibitSendQuota)
                            $Configuration.ADUsers[$CUPN]."Mailbox Archive Prohibit Send Receive Quota" = Format-MailboxGB($Mailbox.ProhibitSendReceiveQuota)
                            $Configuration.ADUsers[$CUPN]."Mailbox Archive Use Database Quota Defaults" = $Mailbox.UseDatabaseQuotaDefaults
                            $Configuration.ADUsers[$CUPN]."Mailbox Archive Storage Limit Status" = $Mailbox.StorageLimitStatus
                            $Configuration.ADUsers[$CUPN]."Mailbox Archive Forwarding Address" = $Mailbox.ForwardingAddress
                            $Configuration.ADUsers[$CUPN]."Mailbox Archive Forwarding Address SMTP" = $Mailbox.ForwardingSmtpAddress
                            #Update Progress
                            If ( $null -ne $Configuration.Jobs["Online Exchange Mailbox Archive"]) {
                                $Configuration.Jobs["Online Exchange Mailbox Archive"].process.PercentComplete = (($ProgressCounter/$MailboxesCount)*100)
                                $Configuration.Jobs["Online Exchange Mailbox Archive"].process.Status = "Processing Record " + $ProgressCounter + " of " + $MailboxesCount
                            }
                        }
                        $ProgressCounter++
                    } -ThrottleLimit $Configuration.processbatchjobs
                }else{
                    ForEach ($Mailbox in $Mailboxes) {
                        $CUPN = $Mailbox.UserPrincipalName

                        If($Mailbox.UserPrincipalName -and $Configuration.ADUsers[$CUPN]."User Principal Name") {
                           #Use MailboxLocations to get Mailbox Archive Server
                           $MailboxLocation = ($Mailbox.MailboxLocations).Split(";")
                           For ($i=0; $i -lt $MailboxLocation.Count; $i++) {
                               If($MailboxLocation[$i] -match "MainArchive") {
                                   If($null -eq  $Configuration.ADUsers[$CUPN]."Mailbox Archive Server") {
                                           $Configuration.ADUsers[$CUPN]."Mailbox Archive Server" = $MailboxLocation[($i + 1)]
                                   }
                                   If($null -eq  $Configuration.ADUsers[$CUPN]."Online Mailbox Archive GUID") {
                                       $Configuration.ADUsers[$CUPN]."Online Mailbox Archive GUID" = $MailboxLocation[($i -1)]
                                   }
                               }
                           }
                            $Configuration.ADUsers[$CUPN]."Mailbox Archive Database" = $Mailbox.Database
                            If($Configuration.ADUsers[$CUPN]."Mailbox Archive Location" -ne "Hybrid Remote") {
                                If($Configuration.ADUsers[$CUPN]."Mailbox Archive GUID" -ne $Mailbox.ArchiveGuid -and ($Mailbox.ArchiveGuid -ne [System.Guid]::Empty)) {
                                    Write-Warning ("Online Mailbox Archive GUID Mismatch: " + $Mailbox.UserPrincipalName + " Online GUID: " + $Mailbox.ArchiveGuid + " Local GUID: " + $Configuration.ADUsers[$CUPN]."Mailbox Archive GUID")
                                }
                            }
                            If($Mailbox.ArchiveGuid -ne [System.Guid]::Empty) {
                                $Configuration.ADUsers[$CUPN]."Online Mailbox Archive GUID" = $Mailbox.ArchiveGuid
                            }
                            $Configuration.ADUsers[$CUPN]."Mailbox Archive Creation Date" = $Mailbox.WhenCreated
                            $Configuration.ADUsers[$CUPN]."Mailbox Archive Status" = $Mailbox.ArchiveStatus
                            $Configuration.ADUsers[$CUPN]."Mailbox Archive State" = $Mailbox.ArchiveState
                            If($Mailbox.AutoExpandingArchive) {
                                $Configuration.ADUsers[$CUPN]."Mailbox Archive Auto Expanding Archive" = $Mailbox.AutoExpandingArchive
                            }
                            $Configuration.ADUsers[$CUPN]."Mailbox Archive Issue Warning Quota" = Format-MailboxGB($Mailbox.IssueWarningQuota)
                            $Configuration.ADUsers[$CUPN]."Mailbox Archive Prohibit Send Quota" = Format-MailboxGB($Mailbox.ProhibitSendQuota)
                            $Configuration.ADUsers[$CUPN]."Mailbox Archive Prohibit Send Receive Quota" = Format-MailboxGB($Mailbox.ProhibitSendReceiveQuota)
                            $Configuration.ADUsers[$CUPN]."Mailbox Archive Use Database Quota Defaults" = $Mailbox.UseDatabaseQuotaDefaults
                            $Configuration.ADUsers[$CUPN]."Mailbox Archive Storage Limit Status" = $Mailbox.StorageLimitStatus
                            $Configuration.ADUsers[$CUPN]."Mailbox Archive Forwarding Address" = $Mailbox.ForwardingAddress
                            $Configuration.ADUsers[$CUPN]."Mailbox Archive Forwarding Address SMTP" = $Mailbox.ForwardingSmtpAddress
                            #Update Progress
                            If ($null -ne $Configuration.Jobs["Online Exchange Mailbox Archive"]) {
                                $Configuration.Jobs["Online Exchange Mailbox Archive"].process.PercentComplete = (($ProgressCounter/$MailboxesCount)*100)
                                $Configuration.Jobs["Online Exchange Mailbox Archive"].process.Status = "Processing Record " + $ProgressCounter + " of " + $MailboxesCount
                            }
                        }
                    }
                }
            }
            $powershell = [powershell]::Create()
            $powershell.RunspacePool = $runspacepool
            $powershell.AddScript($EXOMailboxesArchiveScript) | Out-Null
            $powershell.AddParameter('Configuration',$Configuration) | Out-Null
            $Configuration.Jobs["Online Exchange Mailbox Archive"] = [PSCustomObject]@{ Pipe = $powershell; Status = $powershell.BeginInvoke(); Name= "Online Exchange Mailbox Archive" ;Start = Get-Date ; End = ""; process  = [PSCustomObject]@{ Id= ($Configuration.Jobs.Count + 2); Activity = "Processing Online Exchange Mailbox Archive" ; Status = "Starting"; PercentComplete = 0;Completed =$false} }       
        #endregion Online Exchange Mailbox Archive
        #region Online Exchange Mailbox Archive Permission
            $EXOMailboxesArchivePermissionsScript = {
                param(
                    $Configuration
                )
                $ProgressCounter = 1              
                #Import Class ADExchangeOutput
                Invoke-Expression $Configuration.ClassADExchangeOutput
                #region Connect to Exchange Online
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
                    If (Get-ChildItem "Cert:\LocalMachine\My\" | Where-Object {$_.Thumbprint -eq  $Configuration.AZCertThumbprint}) {
                        Connect-ExchangeOnline -CertificateThumbPrint  $Configuration.AZCertThumbprint -AppID  $Configuration.AZClientID -Organization  $Configuration.AZOrg -ShowProgress:$false -ShowBanner:$false
                    }Else{
                        Connect-ExchangeOnline -ShowBanner:$false -UserPrincipalName  $Configuration.CurrentUserUPN -ShowProgress:$false
                    }
                }
                #endregion Connect to Exchange Online
                #Load .Net Assembly for AD
                Add-Type -AssemblyName System.DirectoryServices.AccountManagement
                $ContextType = [System.DirectoryServices.AccountManagement.ContextType]::Domain
                $IdentityType = [System.DirectoryServices.AccountManagement.IdentityType]::SamAccountName

                ForEach ($CUPN in ($Configuration.ADUsersUPNsWithERRA)) {
                    If($Configuration.ADUsers[$CUPN].'Online Mailbox Archive GUID') {
                        $EPMailboxesPerms = Get-EXOMailboxPermission -ResultSize Unlimited -Identity $Configuration.ADUsers[$CUPN].'Online Mailbox Archive GUID' -erroraction silentlycontinue
                    }
                    If($EPMailboxesPerms.AccessRights.count -gt 0 -and $Configuration.ADUsers[$CUPN].'Mailbox Archive GUID') {
                        $EPMailboxesPerms = Get-EXOMailboxPermission -ResultSize Unlimited -Identity $Configuration.ADUsers[$CUPN].'Mailbox Archive GUID' -erroraction silentlycontinue
                    }
                    #region Mailbox Permission
                    $DAMP = @()
                    $FixedDAMP = @{}
                    $CSVFixedDAMP = ""
                    If($EPMailboxesPerms.AccessRights.count -gt 0 -and $Configuration.ADUsers[$CUPN]."User Principal Name") {
                        $DAMP = ($EPMailboxesPerms.Where({($_.AccessRights -eq "FullAccess") -and ($_.User -notin $Configuration.ExcludeUsers) -and ($_.User -notmatch "S-1-5-*" ) -and $_.IsInherited -eq $false})) | Sort-Object -Unique -Property User	
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
                                                    If ($ADUser.description -notmatch "Leave") {
                                                        $FixedDAMP.add(($ACE.User + " - Leave"),$ACE.AccessRights)
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
                                    $Configuration.ADUsers[$CUPN]."Mailbox Archive Permissions" = $CSVFixedDAMP
                                }
                            }
                        }
                        #Update Progress
                        If ( $null -ne $Configuration.Jobs["Online Exchange Mailbox Archive Permission"]) {
                            $Configuration.Jobs["Online Exchange Mailbox Archive Permission"].process.PercentComplete = (($ProgressCounter/$Configuration.ADUsersUPNsWithERRA.Count)*100)
                            $Configuration.Jobs["Online Exchange Mailbox Archive Permission"].process.Status = "Processing Record " + $ProgressCounter + " of " + $Configuration.ADUsersUPNsWithERRA.Count
                        }
                    }	
                    #endregion Mailbox Permission
                }
            }
        
            $powershell = [powershell]::Create()
            $powershell.RunspacePool = $runspacepool
            $powershell.AddScript($EXOMailboxesArchivePermissionsScript) | Out-Null
            $powershell.AddParameter('Configuration',$Configuration) | Out-Null
            $Configuration.Jobs["Online Exchange Mailbox Archive Permission"] = [PSCustomObject]@{ Pipe = $powershell; Status = $powershell.BeginInvoke(); Name= "Online Exchange Mailbox Archive Permission" ;Start = Get-Date ; End = ""; process  = [PSCustomObject]@{ Id= ($Configuration.Jobs.Count + 2); Activity = "Processing Online Exchange Mailbox Archive Permission" ; Status = "Starting"; PercentComplete = 0;Completed =$false} }       
        #endregion Online Exchange Mailbox Archive Permission
        #region Online Exchange Mailbox Archive Statistics
            $EXOMailboxesScriptStatisticsArchive = {
                param(
                    $Configuration
                )
                $ProgressCounter = 1
                #Import Class ADExchangeOutput
                Invoke-Expression $Configuration.ClassADExchangeOutput
                #region Connect to Exchange Online
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
                    If (Get-ChildItem "Cert:\LocalMachine\My\" | Where-Object {$_.Thumbprint -eq  $Configuration.AZCertThumbprint}) {
                        Connect-ExchangeOnline -CertificateThumbPrint  $Configuration.AZCertThumbprint -AppID  $Configuration.AZClientID -Organization  $Configuration.AZOrg -ShowProgress:$false -ShowBanner:$false
                    }Else{
                        Connect-ExchangeOnline -ShowBanner:$false -UserPrincipalName  $Configuration.CurrentUserUPN -ShowProgress:$false
                    }
                }
                #endregion Connect to Exchange Online

                $ProgressCounter = 1
                $Mailboxes = Get-EXOMailbox -ResultSize unlimited -PropertySets All -Archive
                $MailboxesCount = $Mailboxes.Count
                
                ForEach ( $Mailbox in $Mailboxes) {
                    $CUPN = $Mailbox.UserPrincipalName
                    $MailboxStatsArchive =  $Mailbox | Get-EXOMailboxStatistics -Archive -Properties TotalItemSize,ItemCount,LastLoggedOnUserAccount,LastLogonTime,LastLogoffTime

                    If($Configuration.ADUsers[$CUPN]."User Principal Name" -and $MailboxStatsArchive.LastLogonTimee) {
                        $Configuration.ADUsers[$CUPN]."Mailbox Archive Size GB" = Format-MailboxGB ($MailboxStatsArchive.TotalItemSize)
                        $Configuration.ADUsers[$CUPN]."Mailbox Archive Item Count" = $MailboxStatsArchive.ItemCount
                        $Configuration.ADUsers[$CUPN]."Mailbox Archive Last Logged On User Account" = $MailboxStatsArchive.LastLoggedOnUserAccount
                        $Configuration.ADUsers[$CUPN]."Mailbox Archive Last Logon Time" = $MailboxStatsArchive.LastLogonTime
                        $Configuration.ADUsers[$CUPN]."Mailbox Archive Last Logoff Time" = $MailboxStatsArchive.LastLogoffTime
                    }
                    #Update Progress
                    If ( $null -ne $Configuration.Jobs["Online Exchange Mailbox Archive Statistics"]) {
                        $Configuration.Jobs["Online Exchange Mailbox Archive Statistics"].process.PercentComplete = (($ProgressCounter/$MailboxesCount)*100)
                        $Configuration.Jobs["Online Exchange Mailbox Archive Statistics"].process.Status = "Processing Record " + $ProgressCounter + " of " + $MailboxesCount
                    }
                    $ProgressCounter++
                }
            }
            $powershell = [powershell]::Create()
            $powershell.RunspacePool = $runspacepool
            $powershell.AddScript($EXOMailboxesScriptStatisticsArchive) | Out-Null
            $powershell.AddParameter('Configuration',$Configuration) | Out-Null
            $Configuration.Jobs["Online Exchange Mailbox Archive Statistics"] = [PSCustomObject]@{ Pipe = $powershell; Status = $powershell.BeginInvoke(); Name= "Online Exchange Mailbox Archive Statistics" ;Start = Get-Date ; End = ""; process  = [PSCustomObject]@{ Id= ($Configuration.Jobs.Count + 2); Activity = "Processing Online Exchange Mailbox Archive Statistics" ; Status = "Starting"; PercentComplete = 0;Completed =$false} }   
        #endregion Online Exchange Mailbox Archive Statistics
        #region Online Exchange Mailbox Forwarding Rules
            $EXOMailboxesForwardingRulesScript = {
                param(
                    $UPNs,
                    $Configuration,
                    $RunspaceName
                )
                $ProgressCounter = 1
                #Import Class ADExchangeOutput
                Invoke-Expression $Configuration.ClassADExchangeOutput
                #region Connect to Exchange Online
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
                    If (Get-ChildItem "Cert:\LocalMachine\My\" | Where-Object {$_.Thumbprint -eq  $Configuration.AZCertThumbprint}) {
                        Connect-ExchangeOnline -CertificateThumbPrint  $Configuration.AZCertThumbprint -AppID  $Configuration.AZClientID -Organization  $Configuration.AZOrg -ShowProgress:$false -ShowBanner:$false
                    }Else{
                        Connect-ExchangeOnline -ShowBanner:$false -UserPrincipalName  $Configuration.CurrentUserUPN -ShowProgress:$false
                    }
                }
                #endregion Connect to Exchange Online

                ForEach ($CUPN in ($UPNs)) {
                    $Rules = $null
                    If ($Configuration.ADUsers[$CUPN].'Online Mailbox GUID') {
                        $Rules = Get-InboxRule -Mailbox ([System.Guid]$Configuration.ADUsers[$CUPN].'Online Mailbox GUID').Guid -IncludeHidden \ -ErrorAction SilentlyContinue | Where-Object {-Not [string]::IsNullOrWhiteSpace($_.ForwardTo)}
                    }
                    If ($Configuration.ADUsers[$CUPN].'Mailbox GUID' -and $null -eq $Rules) {
                        $Rules = Get-InboxRule -Mailbox ([System.Guid]$Configuration.ADUsers[$CUPN].'Mailbox GUID').Guid -IncludeHidden -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | Where-Object {-Not [string]::IsNullOrWhiteSpace($_.ForwardTo)}
                        
                    }
                    If ($Rules) {
                        $ORules = @()
                        ForEach($Rule in $Rules) {
                            If ($Rule.ForwardTo -match '" \[EX:'){
                                $ORules += ($Rule.Name + " = Mailbox: " + (($Rule.ForwardTo  -split "\[")[0] -replace '"'))
                            }Else{
                                $ORules += ($Rule.Name + " = " + ($Rule.ForwardTo -split "\[SMTP\:")[-1] -replace "]")
                            } 
                        }	
                        $Configuration.ADUsers[$CUPN]."Mailbox Forwarding Rules" =  $ORules -join ","

                        #Update Progress
                        If ( $null -ne $Configuration.Jobs[$RunspaceName]) {
                            $Configuration.Jobs[$RunspaceName].process.PercentComplete = (($ProgressCounter/$UPNs.Count)*100)
                            $Configuration.Jobs[$RunspaceName].process.Status = "Processing Record " + $ProgressCounter + " of " + $UPNs.Count
                        }
                    }
                    # Write-Progress -Activity $RunspaceName -Status "Processing Record $ProgressCounter of $($UPNs.Count)" -PercentComplete (($ProgressCounter/$UPNs.Count)*100)
                    $ProgressCounter++
                }
            }  
             #Setup number of batches
            $ADUsersUPNsWithERRACount = $Configuration.ADUsersUPNsWithERRA.Count
            $PBatchJobs = [Math]::Round(($Configuration.processbatchjobs)/2)
            If ( $PBatchJobs -lt 1) {
                $PBatchJobs = 1
            }
            $BSize = [Math]::Round( $ADUsersUPNsWithERRACount / $PBatchJobs)
            $BStart = 1
            $BEnd = $BSize

            Do {               
                $powershell = [powershell]::Create()
                $powershell.RunspacePool = $runspacepool
                $powershell.AddScript($EXOMailboxesForwardingRulesScript) | Out-Null
                $powershell.AddParameter('$UPNs',$Configuration.ADUsersUPNsWithERRA[$BStart..$BEnd]) | Out-Null                
                $powershell.AddParameter('Configuration',$Configuration) | Out-Null
                $powershell.AddParameter('RunspaceName',("Online Exchange Mailbox Forwarding Rules Start: " + $BStart + " End: " + $BEnd )) | Out-Null
                $Configuration.Jobs[("Online Exchange Mailbox Forwarding Rules Start: " + $BStart + " End: " + $BEnd )] = [PSCustomObject]@{ Pipe = $powershell; Status = $powershell.BeginInvoke(); Name= ("Online Exchange Mailbox Forwarding Rules Start: " + $BStart + " End: " + $BEnd ) ;Start = Get-Date ; End = ""; process  = [PSCustomObject]@{ Id= ($Configuration.Jobs.Count + 2); Activity = ("Processing Online Exchange Mailbox Forwarding Rules Start: " + $BStart + " End: " + $BEnd ) ; Status = "Starting"; PercentComplete = 0;Completed =$false} }   
                #Increment  Batch
                $BStart = $BEnd + 1
                $BEnd = $BEnd + $BSize
                If ($BEnd -gt $ADUsersUPNsWithERRACount) {
                    $BEnd = $ADUsersUPNsWithERRACount
                }
            } while ($BStart -lt $ADUsersUPNsWithERRACount) 
        #endregion Online Exchange Mailbox Forwarding Rules
        #region Online Exchange CAS Mailbox Settings
            $EXOCASMailboxScript = {
                param(
                    $Configuration
                )
                #Import Class ADExchangeOutput
                Invoke-Expression $Configuration.ClassADExchangeOutput
                #region Connect to Exchange Online
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
                    If (Get-ChildItem "Cert:\LocalMachine\My\" | Where-Object {$_.Thumbprint -eq  $Configuration.AZCertThumbprint}) {
                        Connect-ExchangeOnline -CertificateThumbPrint  $Configuration.AZCertThumbprint -AppID  $Configuration.AZClientID -Organization  $Configuration.AZOrg -ShowProgress:$false -ShowBanner:$false
                    }Else{
                        Connect-ExchangeOnline -ShowBanner:$false -UserPrincipalName  $Configuration.CurrentUserUPN -ShowProgress:$false
                    }
                }
                #endregion Connect to Exchange Online
                $ProgressCounter = 1

                # Get-EXOCasMailbox -PropertySets All
                $Mailboxes = Get-EXOCasMailbox -PropertySets All
                $MailboxesCount = $Mailboxes.Count
                If ($PSVersionTable.PSVersion.Major -ge 7) {
                    $Mailboxes | ForEach-Object -Parallel {
                        $Mailbox = $_
                        $CUPN = $Mailbox.UserPrincipalName
                        $Configuration = $using:Configuration
                        $ProgressCounter = $using:ProgressCounter
                        $MailboxesCount = $using:MailboxesCount
                        #Import Class ADExchangeOutput
                        Invoke-Expression $Configuration.ClassADExchangeOutput
                        If (-Not $Configuration.ADUsers[$CUPN]."User Principal Name" -and $Mailbox.SamAccountName) {
                            $Configuration.ADUsers[$CUPN]."OWA Enabled" = $Mailbox.OWAEnabled
                            $Configuration.ADUsers[$CUPN]."Mapi Enabled" = $Mailbox.MAPIEnabled
                            $Configuration.ADUsers[$CUPN]."IMAP Enabled" = $Mailbox.IMAPEnabled
                            $Configuration.ADUsers[$CUPN]."Active Sync Enabled" = $Mailbox.ActiveSyncEnabled
                            $Configuration.ADUsers[$CUPN]."POP Enabled" = $Mailbox.POPEnabled
        
                            # $Configuration.ADUsers[$CUPN]."Mailbox CAS Server" = $Mailbox.ServerName
                            # $Configuration.ADUsers[$CUPN]."Mailbox CAS Database" = $Mailbox.Database
                            # $Configuration.ADUsers[$CUPN]."Mailbox CAS OWA for Devices Enabled" = $Mailbox.OWAforDevicesEnabled
                            # $Configuration.ADUsers[$CUPN]."Mailbox CAS ECP Enabled" = $Mailbox.ECPEnabled
                            # $Configuration.ADUsers[$CUPN]."Mailbox CAS EWS Enabled" = $Mailbox.EWSEnabled
                            # $Configuration.ADUsers[$CUPN]."Mailbox CAS OAB Enabled" = $Mailbox.OABEnabled
                            # $Configuration.ADUsers[$CUPN]."Mailbox CAS ActiveSync Version" = $Mailbox.ActiveSyncVersion
                            # $Configuration.ADUsers[$CUPN]."Mailbox CAS OWA Version" = $Mailbox.OWAVersion
                            # $Configuration.ADUsers[$CUPN]."Mailbox CAS ECP Version" = $Mailbox.ECPVersion
                            # $Configuration.ADUsers[$CUPN]."Mailbox CAS EWS Version" = $Mailbox.EWSVersion
                            # $Configuration.ADUsers[$CUPN]."Mailbox CAS IMAP Version" = $Mailbox.IMAPVersion
                            # $Configuration.ADUsers[$CUPN]."Mailbox CAS MAPI Version" = $Mailbox.MAPIVersion
                            # $Configuration.ADUsers[$CUPN]."Mailbox CAS OAB Version" = $Mailbox.OABVersion
                            # $Configuration.ADUsers[$CUPN]."Mailbox CAS POP Version" = $Mailbox.POPVersion
                            # $Configuration.ADUsers[$CUPN]."Mailbox CAS ActiveSync Device ID" = $Mailbox.ActiveSyncDeviceID
                            # $Configuration.ADUsers[$CUPN]."Mailbox CAS ActiveSync Device Model" = $Mailbox.ActiveSyncDeviceModel
                            # $Configuration.ADUsers[$CUPN]."Mailbox CAS ActiveSync Device Type" = $Mailbox.ActiveSyncDeviceType
                            # $Configuration.ADUsers[$CUPN]."Mailbox CAS ActiveSync Device User Agent" = $Mailbox.ActiveSyncDeviceUserAgent
                            # $Configuration.ADUsers[$CUPN]."Mailbox CAS ActiveSync Device Access State" = $Mailbox.ActiveSyncDeviceAccessState
        
                            #Update Progress
                            If ( $null -ne $Configuration.Jobs["Online Exchange CAS Mailbox Settings"]) {
                                $Configuration.Jobs["Online Exchange CAS Mailbox Settings"].process.PercentComplete = (($ProgressCounter/$MailboxesCount)*100)
                                $Configuration.Jobs["Online Exchange CAS Mailbox Settings"].process.Status = "Processing Record " + $ProgressCounter + " of " + $MailboxesCount
                            }
                        }
                        $ProgressCounter++
                    } -ThrottleLimit $Configuration.processbatchjobs
                }Else{
                    ForEach ($Mailbox in ($Mailboxes)) {
                        $CUPN = $Mailbox.UserPrincipalName
                        If (-Not $Configuration.ADUsers[$CUPN]."User Principal Name" -and $Mailbox.SamAccountName) {
                            $Configuration.ADUsers[$CUPN]."OWA Enabled" = $Mailbox.OWAEnabled
                            $Configuration.ADUsers[$CUPN]."Mapi Enabled" = $Mailbox.MAPIEnabled
                            $Configuration.ADUsers[$CUPN]."IMAP Enabled" = $Mailbox.IMAPEnabled
                            $Configuration.ADUsers[$CUPN]."Active Sync Enabled" = $Mailbox.ActiveSyncEnabled
                            $Configuration.ADUsers[$CUPN]."POP Enabled" = $Mailbox.POPEnabled
        
                            # $Configuration.ADUsers[$CUPN]."Mailbox CAS Server" = $Mailbox.ServerName
                            # $Configuration.ADUsers[$CUPN]."Mailbox CAS Database" = $Mailbox.Database
                            # $Configuration.ADUsers[$CUPN]."Mailbox CAS OWA for Devices Enabled" = $Mailbox.OWAforDevicesEnabled
                            # $Configuration.ADUsers[$CUPN]."Mailbox CAS ECP Enabled" = $Mailbox.ECPEnabled
                            # $Configuration.ADUsers[$CUPN]."Mailbox CAS EWS Enabled" = $Mailbox.EWSEnabled
                            # $Configuration.ADUsers[$CUPN]."Mailbox CAS OAB Enabled" = $Mailbox.OABEnabled
                            # $Configuration.ADUsers[$CUPN]."Mailbox CAS ActiveSync Version" = $Mailbox.ActiveSyncVersion
                            # $Configuration.ADUsers[$CUPN]."Mailbox CAS OWA Version" = $Mailbox.OWAVersion
                            # $Configuration.ADUsers[$CUPN]."Mailbox CAS ECP Version" = $Mailbox.ECPVersion
                            # $Configuration.ADUsers[$CUPN]."Mailbox CAS EWS Version" = $Mailbox.EWSVersion
                            # $Configuration.ADUsers[$CUPN]."Mailbox CAS IMAP Version" = $Mailbox.IMAPVersion
                            # $Configuration.ADUsers[$CUPN]."Mailbox CAS MAPI Version" = $Mailbox.MAPIVersion
                            # $Configuration.ADUsers[$CUPN]."Mailbox CAS OAB Version" = $Mailbox.OABVersion
                            # $Configuration.ADUsers[$CUPN]."Mailbox CAS POP Version" = $Mailbox.POPVersion
                            # $Configuration.ADUsers[$CUPN]."Mailbox CAS ActiveSync Device ID" = $Mailbox.ActiveSyncDeviceID
                            # $Configuration.ADUsers[$CUPN]."Mailbox CAS ActiveSync Device Model" = $Mailbox.ActiveSyncDeviceModel
                            # $Configuration.ADUsers[$CUPN]."Mailbox CAS ActiveSync Device Type" = $Mailbox.ActiveSyncDeviceType
                            # $Configuration.ADUsers[$CUPN]."Mailbox CAS ActiveSync Device User Agent" = $Mailbox.ActiveSyncDeviceUserAgent
                            # $Configuration.ADUsers[$CUPN]."Mailbox CAS ActiveSync Device Access State" = $Mailbox.ActiveSyncDeviceAccessState
        
                            #Update Progress
                            If ( $null -ne $Configuration.Jobs["Online Exchange CAS Mailbox Settings"]) {
                                $Configuration.Jobs["Online Exchange CAS Mailbox Settings"].process.PercentComplete = (($ProgressCounter/$MailboxesCount)*100)
                                $Configuration.Jobs["Online Exchange CAS Mailbox Settings"].process.Status = "Processing Record " + $ProgressCounter + " of " + $MailboxesCount
                            }
                        }
                    }
                }
            }               
            $powershell = [powershell]::Create()
            $powershell.RunspacePool = $runspacepool
            $powershell.AddScript($EXOCASMailboxScript) | Out-Null
            $powershell.AddParameter('Configuration',$Configuration) | Out-Null
            $Configuration.Jobs["Online Exchange CAS Mailbox Settings"] = [PSCustomObject]@{ Pipe = $powershell; Status = $powershell.BeginInvoke(); Name= "Online Exchange CAS Mailbox Settings" ;Start = Get-Date ; End = ""; process  = [PSCustomObject]@{ Id= ($Configuration.Jobs.Count + 2); Activity = "Processing Online Exchange CAS Mailbox Settings" ; Status = "Starting"; PercentComplete = 0;Completed =$false} }         
        #endregion Online Exchange CAS Mailbox Settings
        #region Online Exchange Mobile Mailbox Settings
            $EXOMobileMailboxScript = {
                param(
                    $Configuration
                )
                $ProgressCounter = 1
                #Import Class ADExchangeOutput
                Invoke-Expression $Configuration.ClassADExchangeOutput
                #region Connect to Exchange Online
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
                    If (Get-ChildItem "Cert:\LocalMachine\My\" | Where-Object {$_.Thumbprint -eq  $Configuration.AZCertThumbprint}) {
                        Connect-ExchangeOnline -CertificateThumbPrint  $Configuration.AZCertThumbprint -AppID  $Configuration.AZClientID -Organization  $Configuration.AZOrg -ShowProgress:$false -ShowBanner:$false
                    }Else{
                        Connect-ExchangeOnline -ShowBanner:$false -UserPrincipalName  $Configuration.CurrentUserUPN -ShowProgress:$false
                    }
                }
                #endregion Connect to Exchange Online
                ForEach ($CUPN in ($Configuration.ADUsersUPNsWithERRA)) {
                    if($Configuration.ADUsers[$CUPN]."Mailbox GUID") {
                        $MobileDevices = Get-MobileDevice -Mailbox $Configuration.ADUsers[$CUPN]."Mailbox GUID" -ResultSize Unlimited -ErrorAction silentlycontinue| Where-Object {$_.DeviceAccessState -eq "Allowed"} | Sort-Object -Property WhenCreated -Descending

                        If( $MobileDevices.count -gt 0) {
                            $MMDClean = @()
                            ForEach ($Device in $MobileDevices) {
                                $DeviceString= ""
                                If (-Not [string]::IsNullOrWhiteSpace($Device.FriendlyName)) {
                                    $DeviceString =  $Device.FriendlyName
                                }
                                If (-Not [string]::IsNullOrWhiteSpace($Device.DeviceModel)) {
                                    If ($DeviceString) {
                                        $DeviceString = ($DeviceString + " - " + $Device.DeviceModel)
                                    }else{
                                        $DeviceString = $Device.DeviceModel
                                    }
                                }
                                If (-Not [string]::IsNullOrWhiteSpace($Device.DeviceId)) {
                                    If ($DeviceString) {
                                        $DeviceString = ($DeviceString + " - " + $Device.DeviceId)
                                    }else{
                                        $DeviceString = $Device.DeviceId
                                    }
                                }
                                If ($DeviceString) {
                                    $MMDClean += $DeviceString
                                }
                            }
                            If($MMDClean.count -gt 0) {
                                $Configuration.ADUsers[$CUPN]."Exchange Mobile Devices" = $MMDClean -join ","
                            }
                        }
                        #Update Progress
                        If ( $null -ne $Configuration.Jobs["Online Exchange Mobile Mailbox Settings"]) {
                            $Configuration.Jobs["Online Exchange Mobile Mailbox Settings"].process.PercentComplete = (($ProgressCounter/$Configuration.ADUsersUPNsWithERRA.Count)*100)
                            $Configuration.Jobs["Online Exchange Mobile Mailbox Settings"].process.Status = "Processing Record " + $ProgressCounter + " of " + $Configuration.ADUsersUPNsWithERRA.Count
                        }
                    }
                    # Write-Progress -Activity "Processing Online Exchange Mobile Mailbox Settings" -Status ("Processing Record " + $ProgressCounter + " of " + $Configuration.ADUsersUPNsWithERRA.Count) -PercentComplete (($ProgressCounter/$Configuration.ADUsersUPNsWithERRA.Count)*100)
                    $ProgressCounter++
                }			
            }
            $powershell = [powershell]::Create()
            $powershell.RunspacePool = $runspacepool
            $powershell.AddScript($EXOMobileMailboxScript) | Out-Null
            $powershell.AddParameter('Configuration',$Configuration) | Out-Null
            $Configuration.Jobs["Online Exchange Mobile Mailbox Settings"] = [PSCustomObject]@{ Pipe = $powershell; Status = $powershell.BeginInvoke(); Name= "Online Exchange Mobile Mailbox Settings" ;Start = Get-Date ; End = ""; process  = [PSCustomObject]@{ Id= ($Configuration.Jobs.Count + 2); Activity = "Processing Online Exchange Mobile Mailbox Settings" ; Status = "Starting"; PercentComplete = 0;Completed =$false} }        
        #endregion Online Exchange Mobile Mailbox Settings
    #endregion Online Exchange Info
#endregion RunSpace Jobs
#region Main Loop to monitor jobs
Write-TMOutput -InputObject ("Start Monitoring Jobs . . .") -ForegroundColor DarkYellow
$swcr = [Diagnostics.Stopwatch]::StartNew()
do {
    $completedRunspaces = $Configuration.Jobs.Values | Where-Object { $_.Pipe.InvocationStateInfo.State -eq 'Completed' }
    $runningRunspaces = $Configuration.Jobs.Values | Where-Object { $_.Pipe.InvocationStateInfo.State -eq 'Running' }

    $totalRunspaces = $Configuration.Jobs.Count

    $completedCount = $completedRunspaces.Count
    $percentComplete = ($completedCount / $totalRunspaces) * 100

    Write-Progress -Activity "Monitoring Runspaces" -Status ("Processing Objects [" + $completedCount + "/" + $totalRunspaces + "] Script time: " + (Format-ElapsedTime($swc.Elapsed)) + " Jobs time: " + (Format-ElapsedTime($swrs.Elapsed))) -PercentComplete $percentComplete -Id 1
   
    # Iterate through the jobs and handle their completion
    ForEach ($Name in $Configuration.Jobs.Keys) {
        If($Configuration.Jobs[$Name].Status.IsCompleted) {
            try {
            # Ensure that the correct IAsyncResult object is passed to EndInvoke
            if ([string]::IsNullOrWhiteSpace($Configuration.Jobs[$Name].End)) {
                $Configuration.Jobs[$Name].End = Get-Date
                If (-Not ([string]::IsNullOrWhiteSpace($Configuration.Jobs[$Name].Start))) {
                    $elapsedTime = New-TimeSpan -Start ($Configuration.Jobs[$Name].Start) -End (Get-Date)
                    Write-TMOutput -InputObject ("`t" + $Name + " has completed in " + (Format-ElapsedTime($elapsedTime))) -ForegroundColor DarkGray
                }Else {
                    Write-TMOutput -InputObject ("`t" + $Name + " has completed Start: " + $Configuration.Jobs[$Name].Start + " End: " + (Get-Date)) -ForegroundColor DarkGray
                }
            }

            $Configuration.Jobs[$Name].Pipe.EndInvoke($Configuration.Jobs[$Name].Status)
            $Configuration.Jobs[$Name].process.Completed = $true
            $Configuration.Jobs[$Name].process.Status = "Completed"
            $Configuration.Jobs[$Name].process.PercentComplete = 100
            } catch {
                Write-Error "Failed to complete job: $($_.Exception.Message)"
            }
        }
        $PercentComplete = 0
        If ($Configuration.Jobs[$Name].process.PercentComplete -gt 100) {
            $PercentComplete = 100
        }Else{
            $PercentComplete = $Configuration.Jobs[$Name].process.PercentComplete
        }
        $param = @{
            ID = $Configuration.Jobs[$Name].process.Id
            Activity = $Configuration.Jobs[$Name].process.Activity
            PercentComplete = $PercentComplete 
            Status = $Configuration.Jobs[$Name].process.Status
            Completed = $Configuration.Jobs[$Name].process.Completed
            ParentId = 1
        }
        Write-Progress @param
        Start-Sleep -Milliseconds 500
    }
    Start-Sleep -Seconds 5
} while ($runningRunspaces.Count -gt 0)

# Wait for all runspaces to complete
$Configuration.Jobs.values | ForEach-Object {
    $_.Pipe.EndInvoke($_.Status)
    $_.Pipe.Dispose()
}
# Ensure all runspaces are closed and disposed of
if ($runspacepool.RunspaceStateInfo.State -eq 'Opened') {
    $runspacepool.Close()
}
$runspacepool.Dispose()
$runspacepool = $null
#endregion Main Loop to monitor job
#region CloseRunspace
    $swrs.Stop()
    Write-Progress -Activity "Monitoring Runspaces" -Status ("Processing Objects [" + $completedCount + "/" + $totalRunspaces + "] Run time: " + (Format-ElapsedTime($swrs.Elapsed))) -PercentComplete $percentComplete -Id 1 -Completed
    $powershell = $null
    $Configuration.jobs.Clear()
    $Configuration.jobs = $null
#endregion CloseRunspace
Write-Progress -Id 1 -Completed -Activity "Monitoring Runspaces" -Status "Completed" -PercentComplete 100

If($Configuration.ADUsers.Count -gt 0 -and $swc.Elapsed.TotalMinutes -gt 0) {
    Write-TMOutput -InputObject ("Runspaces Jobs time: " + (Format-ElapsedTime($swcr.Elapsed)) + " to run. " + '{0:N0}' -f ($Configuration.ADUsers.Count / $swc.Elapsed.TotalMinutes) + " Users's per Minute.") -foregroundColor DarkYellow
}
#region Export to CSV
if($Configuration.ADUsers.Count -gt 0 -and $csvfile) {
    Write-TMOutput -InputObject ("`tExporting to CSV . . .") -ForegroundColor DarkGray
    $Configuration.ADUsers.Values | Export-Csv -Path ($csvfile) -NoTypeInformation
}

if($Configuration.ADGroups.Count -gt 0 -and $csvfile) {
    Write-TMOutput -InputObject ("`tExporting AD Groups to CSV . . .") -ForegroundColor DarkGray
    $Configuration.ADGroups.GetEnumerator() | Export-Csv -Path ($csvfile -replace ".csv","-ADGroups.csv") -NoTypeInformation
}
#endregion Export to CSV

#region Load ImportExcel
    $swo = [Diagnostics.Stopwatch]::StartNew() 
	If(-Not (Get-Module -Name ImportExcel -ListAvailable)){
		Install-Module -Name ImportExcel -Force -Confirm:$false
	}
	If (-Not (Get-Module "ImportExcel" -ErrorAction SilentlyContinue)) {
		Import-Module ImportExcel
	}   
#endregion Load ImportExcel
#region Excel Export
    if($Configuration.ADUsers.Count -gt 0 -and $Configuration.xlsxfile) {
        $excel = $Configuration.ADUsers.Values | Export-Excel -Path $Configuration.xlsxfile -WorksheetName ("Hybrid_Info_" + $Configuration.FileDate) -AutoFilter -FreezeTopRowFirstColumn
        #region Disabled Users with Mailboxes
            #New worksheet that has created,pw change, last logon all over 60 days.
            $DSCOutput = ($Configuration.ADUsers.Values).Where({$_."Account Status" -eq "Disabled" -and $null -eq $_."Employee Type" -and $null -ne $_."Exchange Recipient Type" -and $_."Description" -notmatch "Leave" -and $null -eq $_."Mailbox Permissions"})| Select-Object "Logon Name","Display Name","Description","User Principal Name","Employee Number", "Azure Licenses","Azure Account Enabled","Group Membership","Account Status","Last Log-On Date","Days Since Last Log-On","Creation Date","Days Since Creation","Last Password Change","Days from last password change","Email","Exchange Remote Routing Address","Exchange Recipient Type","Exchange Mobile Devices","Mailbox Location","Mailbox Permissions","Mailbox Size GB","Mailbox Usage (%)","Mailbox GUID","Mailbox Archive GUID","Distinguished Name"
            $WorksheetName = "Disabled Users with Mailboxes"
            If ($DSCOutput.count -gt 0) {
                Write-Host "Saving $WorksheetName Worksheet"
                $excel =  $DSCOutput | Export-Excel -Path $Configuration.xlsxfile -WorksheetName $WorksheetName -AutoFilter -FreezeTopRowFirstColumn -passThru
                $ws = $excel.Workbook.Worksheets[($WorksheetName)]
                $LastRow = $ws.Dimension.End.Row
                $LastColumn = $ws.Dimension.End.column

                #Header Lookup
                $htHeader =[ordered]@{}
                for ($i = 1; $i -le  $LastColumn; $i++) {
                    $htHeader.add(($ws.Cells[1,$i].value),$i)
                }

                #Wrap Text on Header information
                Set-Format -Worksheet $ws -Range ("A1:" + ($ws.Cells[1,$lastcolumn]).Address) -WrapText -Bold 

                #region Wrap Text
                ForEach ($WTH in $Configuration.CSVtoReturnHeaders) {
                    If ($htHeader[$WTH]) {
                        Set-Format -Worksheet $ws -Range (($ws.Cells[1,$htHeader[$WTH]].Address).Substring(0,($ws.Cells[1,$htHeader[$WTH]].Address).Length-1) + "2:" + ($ws.Cells[1,$htHeader[$WTH]].Address).Substring(0,($ws.Cells[1,$htHeader[$WTH]].Address).Length-1) + $LastRow) -WrapText
                    }
                }
                
                #Changes comma demiliter to new line
                ###Causes file to open in recovery mode.###
                # for ($i = 2; $i -le  $LastRow; $i++) {
                #     ForEach ($CRH in $Configuration.CSVtoReturnHeaders) {
                #         If ($htHeader[$CRH]) {
                #             $ws.Cells[($ws.Cells[1,$htHeader[$CRH]].Address).Substring(0,($ws.Cells[1,$htHeader[$CRH]].Address).Length-1) + $i].Value = (($ws.Cells[($ws.Cells[1,$htHeader[$CRH]].Address).Substring(0,($ws.Cells[1,$htHeader[$CRH]].Address).Length-1) + $i].Value) -replace ",","`r`n")
                #         }
                #     }
                # }
                for ($row = 2; $row -le $LastRow; $row++) {
                    ForEach ($CRH in $Configuration.CSVtoReturnHeaders) {
                        If ($htHeader[$CRH]) {
                            $cell = $ws.Cells[$row, $htHeader[$CRH]] 
                            if ($cell.Value -and $cell.Value -is [string]) {
                                $cell.Value = $cell.Value -replace ",", "`r`n"
                            }
                        }
                    }
                }
                #endregion Wrap Text
                #region Format with Commas
                ForEach ($CH in $Configuration.CommaHeaders) {
                    If ($htHeader[$CH]) {
                        Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader[$CH]].Address).Substring(0,($ws.Cells[1,$htHeader[$CH]].Address).Length-1) + "2:" + ($ws.Cells[1,$htHeader[$CH]].Address).Substring(0,($ws.Cells[1,$htHeader[$CH]].Address).Length-1) + $LastRow) -NumberFormat '#,##0'
                    }
                }
                #endregion Format with Commas
                #region Format with Date-Time
                ForEach ($DTH in $Configuration.DateTimeHeaders) {
                    If ($htHeader[$DTH]) {
                    Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader[$DTH]].Address).Substring(0,($ws.Cells[1,$htHeader[$DTH]].Address).Length-1) + "2:" + ($ws.Cells[1,$htHeader[$DTH]].Address).Substring(0,($ws.Cells[1,$htHeader[$DTH]].Address).Length-1) + $LastRow) -NumberFormat 'Date-Time'
                    }
                }
                #endregion Format with Date-Time
                close-ExcelPackage $excel
            }
            $DSCOutput = $null
            Remove-Variable -Name DSCOutput
        #endregion Disabled Users with Mailboxes
        #region Mailboxes Larger than 90 %
            #New worksheet that has created,pw change, last logon all over 60 days.
            $DSCOutput = ($Configuration.ADUsers.Values).Where({$_."Account Status" -eq "Enabled" -and $null -ne $_."Exchange Recipient Type" -and $_."Mailbox Usage (%)" -gt 90})| Select-Object "Logon Name","Display Name","Description","User Principal Name","Employee Number", "Azure Licenses","Azure Account Enabled","Group Membership","Account Status","Last Log-On Date","Days Since Last Log-On","Creation Date","Days Since Creation","Last Password Change","Days from last password change","Email","Exchange Remote Routing Address","Exchange Recipient Type","Exchange Mobile Devices","Mailbox Location","Mailbox Permissions","Mailbox Size GB","Mailbox Usage (%)","Mailbox Recoverable Items Size GB","Mailbox GUID","Mailbox Archive GUID","Distinguished Name"
            $WorksheetName = "Mailboxes Larger than 90%"
            If ($DSCOutput.count -gt 0) {
                Write-Host "Saving $WorksheetName Worksheet"
                $excel =  $DSCOutput | Export-Excel -Path $Configuration.xlsxfile -WorksheetName $WorksheetName -AutoFilter -FreezeTopRowFirstColumn -passThru
                $ws = $excel.Workbook.Worksheets[($WorksheetName)]
                $LastRow = $ws.Dimension.End.Row
                $LastColumn = $ws.Dimension.End.column

                #Header Lookup
                $htHeader =[ordered]@{}
                for ($i = 1; $i -le  $LastColumn; $i++) {
                    $htHeader.add(($ws.Cells[1,$i].value),$i)
                }

                #Wrap Text on Header information
                Set-Format -Worksheet $ws -Range ("A1:" + ($ws.Cells[1,$lastcolumn]).Address) -WrapText -Bold 

                #region Wrap Text
                ForEach ($WTH in $Configuration.CSVtoReturnHeaders) {
                    If ($htHeader[$WTH]) {
                        Set-Format -Worksheet $ws -Range (($ws.Cells[1,$htHeader[$WTH]].Address).Substring(0,($ws.Cells[1,$htHeader[$WTH]].Address).Length-1) + "2:" + ($ws.Cells[1,$htHeader[$WTH]].Address).Substring(0,($ws.Cells[1,$htHeader[$WTH]].Address).Length-1) + $LastRow) -WrapText
                    }
                }
                #Changes comma demiliter to new line
                ###Causes file to open in recovery mode.###
                # for ($i = 2; $i -le  $LastRow; $i++) {
                #     ForEach ($CRH in $Configuration.CSVtoReturnHeaders) {
                #         If ($htHeader[$CRH]) {
                #             $ws.Cells[($ws.Cells[1,$htHeader[$CRH]].Address).Substring(0,($ws.Cells[1,$htHeader[$CRH]].Address).Length-1) + $i].Value = (($ws.Cells[($ws.Cells[1,$htHeader[$CRH]].Address).Substring(0,($ws.Cells[1,$htHeader[$CRH]].Address).Length-1) + $i].Value) -replace ",","`r`n")
                #         }
                #     }
                # }
                #endregion Wrap Text
                #region Format with Commas
                ForEach ($CH in $Configuration.CommaHeaders) {
                    If ($htHeader[$CH]) {
                        Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader[$CH]].Address).Substring(0,($ws.Cells[1,$htHeader[$CH]].Address).Length-1) + "2:" + ($ws.Cells[1,$htHeader[$CH]].Address).Substring(0,($ws.Cells[1,$htHeader[$CH]].Address).Length-1) + $LastRow) -NumberFormat '#,##0'
                    }
                }
                #endregion Format with Commas
                #region Format with Date-Time
                ForEach ($DTH in $Configuration.DateTimeHeaders) {
                    If ($htHeader[$DTH]) {
                        Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader[$DTH]].Address).Substring(0,($ws.Cells[1,$htHeader[$DTH]].Address).Length-1) + "2:" + ($ws.Cells[1,$htHeader[$DTH]].Address).Substring(0,($ws.Cells[1,$htHeader[$DTH]].Address).Length-1) + $LastRow) -NumberFormat 'Date-Time'
                    }
                }
                #endregion Format with Date-Time
                close-ExcelPackage $excel
            }
            $DSCOutput = $null
            Remove-Variable -Name DSCOutput
        #endregion Mailboxes Larger than 90%
        #region Service Look to Disable
            #New worksheet that has created,pw change, last logon all over 60 days.
            $DSCOutput = ($Configuration.ADUsers.Values).Where({$_."Days Since Creation" -gt 365 -and $_."Days from last password change" -gt 365 -and $_."Account Status" -eq "Enabled" -and $_."Employee Type"}) | Select-Object "Logon Name","Display Name","Description","User Principal Name","Employee Type","Employee Number", "Azure Licenses","Azure Licenses Details","Azure Last Sign-On","Azure Last Sign-On Days","Azure Last Non-Interactive Sign-On","Azure Last Non-Interactive Sign-On Days","Azure Account Enabled","Group Membership","Account Status","Password Never Expires","Password Change on Next Logon","Last Log-On Date","Days Since Last Log-On","Creation Date","Days Since Creation","Last Password Change","Days from last password change","Email","Exchange Remote Routing Address","Exchange Recipient Type","Exchange Mobile Devices","Mailbox Location","Mailbox GUID","Mailbox Archive GUID","ADFS Last Logon","ADFS Last Logon Days","ADFS Last Logon IP","ADFS Relying Party","ADFS Auth Protocol","ADFS Network Location","ADFS ADFS Server","ADFS User Agent String","Distinguished Name"
            $WorksheetName = "Service Accounts Review >=365"
            If ($DSCOutput.count -gt 0) {
                Write-Host "Saving $WorksheetName Worksheet"
                $excel =  $DSCOutput | Export-Excel -Path $Configuration.xlsxfile -WorksheetName $WorksheetName -AutoFilter -FreezeTopRowFirstColumn -passThru
                $ws = $excel.Workbook.Worksheets[($WorksheetName)]
                $LastRow = $ws.Dimension.End.Row
                $LastColumn = $ws.Dimension.End.column

                #Header Lookup
                $htHeader =[ordered]@{}
                for ($i = 1; $i -le  $LastColumn; $i++) {
                    $htHeader.add(($ws.Cells[1,$i].value),$i)
                }

                #Wrap Text on Header information
                Set-Format -Worksheet $ws -Range ("A1:" + ($ws.Cells[1,$lastcolumn]).Address) -WrapText -Bold 

                #region Wrap Text

                ForEach ($WTH in $Configuration.CSVtoReturnHeaders) {
                    If ($htHeader[$WTH]) {
                        Set-Format -Worksheet $ws -Range (($ws.Cells[1,$htHeader[$WTH]].Address).Substring(0,($ws.Cells[1,$htHeader[$WTH]].Address).Length-1) + "2:" + ($ws.Cells[1,$htHeader[$WTH]].Address).Substring(0,($ws.Cells[1,$htHeader[$WTH]].Address).Length-1) + $LastRow) -WrapText
                    }
                }
                #Changes comma demiliter to new line
                ###Causes file to open in recovery mode.###
                # for ($i = 2; $i -le  $LastRow; $i++) {
                #     ForEach ($CRH in $Configuration.CSVtoReturnHeaders) {
                #         If ($htHeader[$CRH]) {
                #             $ws.Cells[($ws.Cells[1,$htHeader[$CRH]].Address).Substring(0,($ws.Cells[1,$htHeader[$CRH]].Address).Length-1) + $i].Value = (($ws.Cells[($ws.Cells[1,$htHeader[$CRH]].Address).Substring(0,($ws.Cells[1,$htHeader[$CRH]].Address).Length-1) + $i].Value) -replace ",","`r`n")
                #         }
                #     }
                # }
                #endregion Wrap Text
                #region Format with Commas
                ForEach ($CH in $Configuration.CommaHeaders) {
                    If ($htHeader[$CH]) {
                        Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader[$CH]].Address).Substring(0,($ws.Cells[1,$htHeader[$CH]].Address).Length-1) + "2:" + ($ws.Cells[1,$htHeader[$CH]].Address).Substring(0,($ws.Cells[1,$htHeader[$CH]].Address).Length-1) + $LastRow) -NumberFormat '#,##0'
                    }
                }
                #endregion Format with Commas
                #region Format with Date-Time
                ForEach ($DTH in $Configuration.DateTimeHeaders) {
                    If ($htHeader[$DTH]) {
                        Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader[$DTH]].Address).Substring(0,($ws.Cells[1,$htHeader[$DTH]].Address).Length-1) + "2:" + ($ws.Cells[1,$htHeader[$DTH]].Address).Substring(0,($ws.Cells[1,$htHeader[$DTH]].Address).Length-1) + $LastRow) -NumberFormat 'Date-Time'
                    }
                }
                #endregion Format with Date-Time
                close-ExcelPackage $excel
            }
            $DSCOutput = $null
            Remove-Variable -Name DSCOutput
        #endregion Service Look to Disable
        #region User Look to Disable
            #New worksheet that has created,pw change, last logon all over 60 days.
            $DSCOutput = ($Configuration.ADUsers.Values).Where({$_."Days Since Creation" -gt 60 -and $_."Days from last password change" -gt 60 -and $_."Account Status" -eq "Enabled" -and $null -eq $_."Employee Type"})| Select-Object "Logon Name","Display Name","Description","User Principal Name","Employee Type","Employee Number", "Azure Licenses","Azure Licenses Details","Azure Last Sign-On","Azure Last Sign-On Days","Azure Last Non-Interactive Sign-On","Azure Last Non-Interactive Sign-On Days","Azure Account Enabled","Group Membership","Account Status","Password Never Expires","Password Change on Next Logon","Last Log-On Date","Days Since Last Log-On","Creation Date","Days Since Creation","Last Password Change","Days from last password change","Email","Exchange Remote Routing Address","Exchange Recipient Type","Exchange Mobile Devices","Mailbox Location","Mailbox GUID","Mailbox Archive GUID","ADFS Last Logon","ADFS Last Logon Days","ADFS Last Logon IP","ADFS Relying Party","ADFS Auth Protocol","ADFS Network Location","ADFS ADFS Server","ADFS User Agent String","Distinguished Name"
            $WorksheetName = "Users Review >=60"
            If ($DSCOutput.count -gt 0) {
                Write-Host "Saving $WorksheetName Worksheet"
                $excel =  $DSCOutput | Export-Excel -Path $Configuration.xlsxfile -WorksheetName $WorksheetName -AutoFilter -FreezeTopRowFirstColumn -passThru
                $ws = $excel.Workbook.Worksheets[($WorksheetName)]
                $LastRow = $ws.Dimension.End.Row
                $LastColumn = $ws.Dimension.End.column

                #Header Lookup
                $htHeader =[ordered]@{}
                for ($i = 1; $i -le  $LastColumn; $i++) {
                    $htHeader.add(($ws.Cells[1,$i].value),$i)
                }

                #Wrap Text on Header information
                Set-Format -Worksheet $ws -Range ("A1:" + ($ws.Cells[1,$lastcolumn]).Address) -WrapText -Bold 

                #region Wrap Text
                ForEach ($WTH in $Configuration.CSVtoReturnHeaders) {
                    If ($htHeader[$WTH]) {
                        Set-Format -Worksheet $ws -Range (($ws.Cells[1,$htHeader[$WTH]].Address).Substring(0,($ws.Cells[1,$htHeader[$WTH]].Address).Length-1) + "2:" + ($ws.Cells[1,$htHeader[$WTH]].Address).Substring(0,($ws.Cells[1,$htHeader[$WTH]].Address).Length-1) + $LastRow) -WrapText
                    }
                }
                #Changes comma demiliter to new line
                ###Causes file to open in recovery mode.###
                # for ($i = 2; $i -le  $LastRow; $i++) {
                #     ForEach ($CRH in $Configuration.CSVtoReturnHeaders) {
                #         If ($htHeader[$CRH]) {
                #             $ws.Cells[($ws.Cells[1,$htHeader[$CRH]].Address).Substring(0,($ws.Cells[1,$htHeader[$CRH]].Address).Length-1) + $i].Value = (($ws.Cells[($ws.Cells[1,$htHeader[$CRH]].Address).Substring(0,($ws.Cells[1,$htHeader[$CRH]].Address).Length-1) + $i].Value) -replace ",","`r`n")
                #         }
                #     }
                # }
                #endregion Wrap Text
                #region Format with Commas
                ForEach ($CH in $Configuration.CommaHeaders) {
                    If ($htHeader[$CH]) {
                        Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader[$CH]].Address).Substring(0,($ws.Cells[1,$htHeader[$CH]].Address).Length-1) + "2:" + ($ws.Cells[1,$htHeader[$CH]].Address).Substring(0,($ws.Cells[1,$htHeader[$CH]].Address).Length-1) + $LastRow) -NumberFormat '#,##0'
                    }
                }
                #endregion Format with Commas
                #region Format with Date-Time
                ForEach ($DTH in $Configuration.DateTimeHeaders) {
                    If ($htHeader[$DTH]) {
                        Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader[$DTH]].Address).Substring(0,($ws.Cells[1,$htHeader[$DTH]].Address).Length-1) + "2:" + ($ws.Cells[1,$htHeader[$DTH]].Address).Substring(0,($ws.Cells[1,$htHeader[$DTH]].Address).Length-1) + $LastRow) -NumberFormat 'Date-Time'
                    }
                }
                #endregion Format with Date-Time
                close-ExcelPackage $excel
            }
            $DSCOutput = $null
            Remove-Variable -Name DSCOutput
        #endregion UserLook to Disable
        #region User Look to Disable no logon
            #New worksheet that has created,pw change, last logon all over 60 days.
            $DSCOutput = ($Configuration.ADUsers.Values).Where({$_."Days Since Creation" -gt 90 -and $_."Days from last password change" -gt 90 -and $_."Account Status" -eq "Enabled" -and $null -eq $_."Employee Type" -and (($null -eq $_.'Azure Last Sign-On Days' -or $null -eq $_.'Azure Last Non-Interactive Sign-On Days') -or ($_.'Azure Last Sign-On Days' -ge 90 -and $_.'Azure Last Non-Interactive Sign-On Days' -ge 90))}) | Select-Object "Logon Name","Display Name","Description","User Principal Name","Employee Type","Employee Number", "Azure Licenses","Azure Licenses Details","Azure Last Sign-On","Azure Last Sign-On Days","Azure Last Non-Interactive Sign-On","Azure Last Non-Interactive Sign-On Days","Azure Account Enabled","Group Membership","Account Status","Password Never Expires","Password Change on Next Logon","Last Log-On Date","Days Since Last Log-On","Creation Date","Days Since Creation","Last Password Change","Days from last password change","Email","Exchange Remote Routing Address","Exchange Recipient Type","Exchange Mobile Devices","Mailbox Location","Mailbox GUID","Mailbox Archive GUID","ADFS Last Logon","ADFS Last Logon Days","ADFS Last Logon IP","ADFS Relying Party","ADFS Auth Protocol","ADFS Network Location","ADFS ADFS Server","ADFS User Agent String","Distinguished Name"
            $WorksheetName = "Users Not Logged on >=90"
            If ($DSCOutput.count -gt 0) {
                Write-Host "Saving $WorksheetName Worksheet"
                $excel =  $DSCOutput | Export-Excel -Path $Configuration.xlsxfile -WorksheetName $WorksheetName -AutoFilter -FreezeTopRowFirstColumn -passThru
                $ws = $excel.Workbook.Worksheets[($WorksheetName)]
                $LastRow = $ws.Dimension.End.Row
                $LastColumn = $ws.Dimension.End.column

                #Header Lookup
                $htHeader =[ordered]@{}
                for ($i = 1; $i -le  $LastColumn; $i++) {
                    $htHeader.add(($ws.Cells[1,$i].value),$i)
                }

                #Wrap Text on Header information
                Set-Format -Worksheet $ws -Range ("A1:" + ($ws.Cells[1,$lastcolumn]).Address) -WrapText -Bold 

                #region Wrap Text
                ForEach ($WTH in $Configuration.CSVtoReturnHeaders) {
                    If ($htHeader[$WTH]) {
                        Set-Format -Worksheet $ws -Range (($ws.Cells[1,$htHeader[$WTH]].Address).Substring(0,($ws.Cells[1,$htHeader[$WTH]].Address).Length-1) + "2:" + ($ws.Cells[1,$htHeader[$WTH]].Address).Substring(0,($ws.Cells[1,$htHeader[$WTH]].Address).Length-1) + $LastRow) -WrapText
                    }
                }
                #Changes comma demiliter to new line
                ###Causes file to open in recovery mode.###
                # for ($i = 2; $i -le  $LastRow; $i++) {
                #     ForEach ($CRH in $Configuration.CSVtoReturnHeaders) {
                #         If ($htHeader[$CRH]) {
                #             $ws.Cells[($ws.Cells[1,$htHeader[$CRH]].Address).Substring(0,($ws.Cells[1,$htHeader[$CRH]].Address).Length-1) + $i].Value = (($ws.Cells[($ws.Cells[1,$htHeader[$CRH]].Address).Substring(0,($ws.Cells[1,$htHeader[$CRH]].Address).Length-1) + $i].Value) -replace ",","`r`n")
                #         }
                #     }
                # }
                #endregion Wrap Text
                #region Format with Commas
                ForEach ($CH in $Configuration.CommaHeaders) {
                    If ($htHeader[$CH]) {
                        Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader[$CH]].Address).Substring(0,($ws.Cells[1,$htHeader[$CH]].Address).Length-1) + "2:" + ($ws.Cells[1,$htHeader[$CH]].Address).Substring(0,($ws.Cells[1,$htHeader[$CH]].Address).Length-1) + $LastRow) -NumberFormat '#,##0'
                    }
                }
                #endregion Format with Commas
                #region Format with Date-Time
                ForEach ($DTH in $Configuration.DateTimeHeaders) {
                    If ($htHeader[$DTH]) {
                        Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader[$DTH]].Address).Substring(0,($ws.Cells[1,$htHeader[$DTH]].Address).Length-1) + "2:" + ($ws.Cells[1,$htHeader[$DTH]].Address).Substring(0,($ws.Cells[1,$htHeader[$DTH]].Address).Length-1) + $LastRow) -NumberFormat 'Date-Time'
                    }
                }
                #endregion Format with Date-Time
                close-ExcelPackage $excel
            }
            $DSCOutput = $null
            Remove-Variable -Name DSCOutput
        #endregion UserLook to Disable no logon
        #region AD Groups
        $WorksheetName = "AD Groups"
        If ($Configuration.ADGroups.count -gt 0) {
            Write-Host "Saving AD Groups Output Worksheet"
            $excel = $Configuration.ADGroups | Export-Excel -Path $Configuration.xlsxfile -WorksheetName $WorksheetName -AutoFilter -FreezeTopRowFirstColumn -passThru
            $excel.save()
        
            Try {
                $ws = $excel.Workbook.Worksheets[($WorksheetName)]
                $LastRow = $ws.Dimension.End.Row
                $LastColumn = $ws.Dimension.End.column
                If ($ws -and $LastRow -and $LastColumn ) {
                    #Header Lookup
                    $htHeader =[ordered]@{}
                    for ($i = 1; $i -le  $LastColumn; $i++) {
                        $htHeader.add(($ws.Cells[1,$i].value),$i)
                    }
                    #Setup
                    Set-ExcelRange -Worksheet $ws -Range ("A1:" + $LastColumn + $LastRow) -VerticalAlignment Top
                    #Members
                    $StrHeader= "Members"
                    Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader[$StrHeader]].Address).Substring(0,($ws.Cells[1,$htHeader[$StrHeader]].Address).Length-1) + "2:" + ($ws.Cells[1,$htHeader[$StrHeader]].Address).Substring(0,($ws.Cells[1,$htHeader[$StrHeader]].Address).Length-1) + $LastRow) -WrapText 
                    #Replace ", " with "`r`n"
                    ($ws.Cells[(($ws.Cells[1,$htHeader[$StrHeader]].Address).Substring(0,($ws.Cells[1,$htHeader[$StrHeader]].Address).Length-1) + "2:" + ($ws.Cells[1,$htHeader[$StrHeader]].Address).Substring(0,($ws.Cells[1,$htHeader[$StrHeader]].Address).Length-1) + $LastRow)]).foreach({$_.Value = $_.Value -replace ",","`r`n"})
                }
                $excel.Save()
                Close-ExcelPackage $excel
            } Catch {
                Write-Host "Error: $($_.Exception.Message)"
            }
        }
        #endregion AD Groups
        #region ADFS Events
        $WorksheetName = "ADFS Events"
        If ($Configuration.ADFSEvents.count -gt 0 ) {
            Write-Host "Saving ADFS Events Output Worksheet"
            $excel =  $Configuration.ADFSEvents.Values | Export-Excel -Path $Configuration.xlsxfile -WorksheetName $WorksheetName -AutoFilter -FreezeTopRowFirstColumn -passThru
            Try {
                $ws = $excel.Workbook.Worksheets[($WorksheetName)]
                $LastRow = $ws.Dimension.End.Row
                $LastColumn = $ws.Dimension.End.column
                If ($ws -and $LastRow -and $LastColumn ) {
                    #Header Lookup
                    $htHeader =[ordered]@{}
                    for ($i = 1; $i -le  $LastColumn; $i++) {
                        $htHeader.add(($ws.Cells[1,$i].value),$i)
                    }
                    #Setup
                    # Set-ExcelRange -Worksheet $ws -Range ("A1:" + $LastColumn + $LastRow) -VerticalAlignment Top
                    Set-ExcelRange -Worksheet $ws -Range ("A1:" + ($ws.Cells[$lastrow,$lastcolumn]).Address) -VerticalAlignment Top -HorizontalAlignment Left  -AutoSize 
                }
                $excel.save()
                Close-ExcelPackage $excel
            } Catch {
                Write-Host "Error: $($_.Exception.Message)"
            }
        }
        #endregion ADFS Events
        #region Excel Formatting
         Write-Host "Formatting $("Hybrid_Info_" + $Configuration.FileDate) Worksheet"
        $excel = Open-ExcelPackage -Path $Configuration.xlsxfile
        If ($excel.Workbook.CalcMode) {
            $ws = $excel.Workbook.Worksheets[("Hybrid_Info_" + $Configuration.FileDate)]
            $LastRow = $ws.Dimension.End.Row
            # $LastRow = $Configuration.ADUsers.Count + 1
            $LastColumn = $ws.Dimension.End.column

            #Header Lookup
            $htHeader =[ordered]@{}
            for ($i = 1; $i -le  $LastColumn; $i++) {
                $htHeader.add(($ws.Cells[1,$i].value),$i)
            }

            #Wrap Text on Header information
            Set-Format -Worksheet $ws -Range ("A1:" + ($ws.Cells[1,$lastcolumn]).Address) -WrapText -Bold 

            #region Wrap Text
            ForEach ($WTH in $Configuration.CSVtoReturnHeaders) {
                If ($htHeader[$WTH]) {
                    Set-Format -Worksheet $ws -Range (($ws.Cells[1,$htHeader[$WTH]].Address).Substring(0,($ws.Cells[1,$htHeader[$WTH]].Address).Length-1) + "2:" + ($ws.Cells[1,$htHeader[$WTH]].Address).Substring(0,($ws.Cells[1,$htHeader[$WTH]].Address).Length-1) + $LastRow) -WrapText
                }
            }
            #Changes comma demiliter to new line
            ###Causes file to open in recovery mode.###
            # for ($i = 2; $i -le  $LastRow; $i++) {
            #     ForEach ($CRH in $Configuration.CSVtoReturnHeaders) {
            #         If ($htHeader[$CRH]) {
            #             $ws.Cells[($ws.Cells[1,$htHeader[$CRH]].Address).Substring(0,($ws.Cells[1,$htHeader[$CRH]].Address).Length-1) + $i].Value = (($ws.Cells[($ws.Cells[1,$htHeader[$CRH]].Address).Substring(0,($ws.Cells[1,$htHeader[$CRH]].Address).Length-1) + $i].Value) -replace ",","`r`n")
            #         }
            #     }
            # }
            #endregion Wrap Text
            #region Format with Commas
            ForEach ($CH in $Configuration.CommaHeaders) {
                If ($htHeader[$CH]) {
                    Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader[$CH]].Address).Substring(0,($ws.Cells[1,$htHeader[$CH]].Address).Length-1) + "2:" + ($ws.Cells[1,$htHeader[$CH]].Address).Substring(0,($ws.Cells[1,$htHeader[$CH]].Address).Length-1) + $LastRow) -NumberFormat '#,##0'
                }
            }
            #endregion Format with Commas
            #region Format with Date-Time
            ForEach ($DTH in $Configuration.DateTimeHeaders) {
                If ($htHeader[$DTH]) {
                    Set-ExcelRange -Worksheet $ws -Range (($ws.Cells[1,$htHeader[$DTH]].Address).Substring(0,($ws.Cells[1,$htHeader[$DTH]].Address).Length-1) + "2:" + ($ws.Cells[1,$htHeader[$DTH]].Address).Substring(0,($ws.Cells[1,$htHeader[$DTH]].Address).Length-1) + $LastRow) -NumberFormat 'Date-Time'
                }
            }
            #endregion Format with Date-Time

            #Setup
            Set-ExcelRange -Worksheet $ws -Range ("A1:" + ($ws.Cells[$lastrow,$lastcolumn]).Address) -VerticalAlignment Top -HorizontalAlignment Left  -AutoSize 

            #Save Excel
            If ($excel.Save()) {
                Close-ExcelPackage $excel
                $excel.Dispose()
                $excel = $null
            }
            # Triggering garbage collection
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
        #endregion Excel Formatting
    }
	$swo.Stop()
	Write-TMOutput -InputObject ("")
	Write-TMOutput -InputObject ("Saving output time: " + (Format-ElapsedTime($swo.Elapsed)) + " to run. ")
#endregion Excel Export
#region Log file cleanup
#endregion Date Range
Write-Host ("Cleaning up output files older than " + $LogPrune + " days:")
#region Prune output Files
((Get-ChildItem -Path (Split-Path -Path $csvfile -Parent) -Filter ("*" + (Split-Path -Path $csvfile -Leaf).Substring(0,(Split-Path -Path $csvfile -Leaf).Length-18) + "*")).Where({$_.LastWriteTime -lt ((Get-Date).AddDays($LogPrune))})) | ForEach-Object {
  write-host ("`t" + $_.Name)
  remove-item -force -Confirm:$false -path $_.fullname
}
#endregion Prune output Files
#endregion Log file cleanup
$swc.Stop()
If($Configuration.ADUsers.Count -gt 0 -and $swc.Elapsed.TotalMinutes -gt 0) {
    Write-TMOutput -InputObject ("") 
    Write-TMOutput -InputObject ("Script completed. Run time: " + (Format-ElapsedTime($swc.Elapsed)) + " to run. " + '{0:N0}' -f ($Configuration.ADUsers.Count / $swc.Elapsed.TotalMinutes) + " Users's per Minute.") -foregroundColor DarkYellow
}
#region Cleanup
    $Configuration.ADUsers.Clear()
    $Configuration.ADUsers = $null
    $Configuration.ADGroups.Clear()
    $Configuration.ADGroups = $null
    $Configuration = $null
    $swc = $null
    $swo = $null
    $swad = $null
#endregion Cleanup
