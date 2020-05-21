<# 
.SYNOPSIS
    Name: Export_AD_Users_to_CSV.ps1
    Creates CSV with all AD user info

.DESCRIPTION
    *Dumps AD User info to CSV.

.EXAMPLE
   & .\Export_AD_Users_to_CSV.ps1

.NOTES
 AUTHOR  : Victor Ashiedu
 WEBSITE : iTechguides.com
 BLOG    : iTechguides.com/blog-2/
 CREATED : 08-08-2014
 Updated By: Paul Fuller
 Changes:
    * Version 1.00.00 - First Release by Victor Ashiedu
    * Version 1.00.01 - Added more fields to export.
    * Version 1.00.02 - Changed how days are calulated to get positive number. Also mailbox location if in cloud or on On-Premise
    

#>

#Import AD modules
If ((Get-Module | Where-Object {$_.Name -Match "ActiveDirectory"}).Count -eq 0 ) {
	Write-Host ("Loading Active Directory Plugins") -foregroundcolor "Green"
	Import-Module "ActiveDirectory"  -ErrorAction SilentlyContinue
} Else {
	Write-Host ("Active Directory Plug-ins Already Loaded") -foregroundcolor "Green"
}

If (-Not [string]::IsNullOrEmpty($LogFile)) {
	If (-Not( Test-Path (Split-Path -Path $LogFile -Parent))) {
		New-Item -ItemType directory -Path (Split-Path -Path $LogFile -Parent)
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

#Define CSV and log file location variables
#they have to be on the same location as the script
$csvfile =  ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\Logs\" + `
			($MyInvocation.MyCommand.Name -replace ".ps1","") + "_" + `
			(Get-Date -format yyyyMMdd-hhmm) + ".csv")
If (-Not( Test-Path (Split-Path -Path $csvfile -Parent))) {
	New-Item -ItemType directory -Path (Split-Path -Path $csvfile -Parent) | Out-Null
}

#Sets the OU to do the base search for all user accounts, change as required.
$SearchBase = (Get-ADDomain).DistinguishedName

#Define variable for a server with AD web services installed
$ADServer = (Get-ADDomain).PDCEmulator

Get-ADUser -server $ADServer -searchbase $SearchBase -Filter * -Properties * | 
Select-Object @{Label = "Last Name";Expression = {$_.Surname}},
@{Label = "Middle Name";Expression = {$_.middleName}},
@{Label = "First Name";Expression = {$_.GivenName}},
@{Label = "Logon Name";Expression = {$_.sAMAccountName}},
@{Label = "Display Name";Expression = {$_.DisplayName}},
@{Label = "Full address";Expression = {$_.StreetAddress}},
@{Label = "City";Expression = {$_.City}},
@{Label = "State";Expression = {$_.st}},
@{Label = "Post Code";Expression = {$_.PostalCode}},
@{Label = "Country/Region";Expression = {if (($_.Country -eq 'GB')  ) {'United Kingdom'} Else {''}}},
@{Label = "Job Title";Expression = {$_.Title}},
@{Label = "Company";Expression = {$_.Company}},
@{Label = "Directorate";Expression = {$_.Description}},
@{Label = "Department";Expression = {$_.Department}},
@{Label = "Employee Type";Expression = {$_.employeeType}},
@{Label = "Office";Expression = {$_.physicalDeliveryOfficeName}},
@{Label = "Phone";Expression = {$_.telephoneNumber}},
@{Label = "Mobile Phone";Expression = {$_.mobile}},
@{Label = "Email";Expression = {$_.Mail}},
@{Label = "Mail Store";Expression = {($_.homeMDB).SubString(3,($_.homeMDB.Indexof(",")-3))}},
@{Label = "Mailbox Locaton";Expression = {If ($_.targetAddress -match "onmicrosoft.com" -and $_.Mail){"Online"}Else{"On-Premise"}}},
@{Label = "Manager";Expression = {ForEach-Object{(Get-AdUser $_.Manager -server $ADServer -Properties DisplayName).DisplayName}}},
@{Label = "Home Directory";Expression = {$_.homeDirectory}},
@{Label = "Account Status";Expression = {if (($_.Enabled -eq 'TRUE')  ) {'Enabled'} Else {'Disabled'}}}, # the 'if statement# replaces $_.Enabled
@{Label = "Password Never Expires";Expression = {if (($_.passwordNeverExpires -eq 'TRUE')  ) {'Enabled'} Else {'Disabled'}}}, # the 'if statement# replaces $_.Enabled
@{Label = "Password Not Required";Expression = {if (($_.PasswordNotRequired -eq 'TRUE')  ) {'Enabled'} Else {'Disabled'}}}, # the 'if statement# replaces $_.Enabled
@{Label = "Smartcard Logon Required";Expression = {if (($_.SmartcardLogonRequired -eq 'TRUE')  ) {'Enabled'} Else {'Disabled'}}}, # the 'if statement# replaces $_.Enabled
@{Label = "Last Log-On Date";Expression = {[DateTime]::FromFileTime($_.lastLogon)}},
@{Label = "Days Since Last LogOn";Expression = {$((Get-Date)- ([DateTime]::FromFileTime($_.lastLogon))).Days}},
@{Label = "Creation Date";Expression = {$_.whencreated}}, 
@{Label = "Days Since Creation";Expression = {$((Get-Date) - ([DateTime]($_.whencreated))).Days}}, 
@{Label = "Last Password Change";Expression = {[DateTime]::FromFileTime($_.pwdLastSet)}}, 
@{Label = "Days from last password change";Expression = {$((Get-Date) - ([DateTime]::FromFileTime($_.pwdLastSet))).Days}},
@{Label = "RDS CAL Expiration Date";Expression = {$_.msTSExpireDate}}, 
@{Label = "Days to RDS CAL Expiration";Expression = {$( (Get-Date) - ([DateTime]($_.msTSExpireDate))).Days}},
@{Label = "Exchange Mailbox Creation Date";Expression = {($_.msExchWhenMailboxCreated)}},
@{Label = "Exchange Mailbox GUID";Expression = {New-Object Guid (,$_.msExchMailboxGuid)}},
@{Label = "Distinguished Name";Expression = {($_.distinguishedName)}} | 
Export-Csv -Path $csvfile -NoTypeInformation


If (-Not [string]::IsNullOrEmpty($LogFile)) {
	Stop-Transcript
}
