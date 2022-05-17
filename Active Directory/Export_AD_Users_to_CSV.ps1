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
    * Version 1.00.02 - Changed how days are calculated to get positive number. Also mailbox location if in cloud or on On-Premise
    * Version 1.00.03 - Added Employee Number
    * Version 1.00.04 - Added export to xlsx too.
    * Version 1.00.05 - Added MemberOf
    

#>
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


#Import AD modules
If ((Get-Module | Where-Object {$_.Name -Match "ActiveDirectory"}).Count -eq 0 ) {
	Write-Host ("Loading Active Directory Plugins") -foregroundcolor "Green"
	Import-Module "ActiveDirectory"  -ErrorAction SilentlyContinue
} Else {
	Write-Host ("Active Directory Plug-ins Already Loaded") -foregroundcolor "Green"
}


#Define CSV and log file location variables
#they have to be on the same location as the script
$output = @()
$FileDate = (Get-Date -format yyyyMMdd-hhmm)
$csvfile =  ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\Logs\" + `
			($MyInvocation.MyCommand.Name -replace ".ps1","") + "_" + `
			$FileDate + ".csv")
$xlsxfile =  ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\Logs\" + `
			($MyInvocation.MyCommand.Name -replace ".ps1","") + "_" + `
			$FileDate + ".xlsx")
If (-Not( Test-Path (Split-Path -Path $csvfile -Parent))) {
	New-Item -ItemType directory -Path (Split-Path -Path $csvfile -Parent) | Out-Null
}

#Sets the OU to do the base search for all user accounts, change as required.
$SearchBase = (Get-ADDomain).DistinguishedName

#Define variable for a server with AD web services installed
$ADServer = (Get-ADDomain).PDCEmulator

$output = Get-ADUser -server $ADServer -searchbase $SearchBase -Filter * -Properties * | 
Select-Object @{Label = "Logon Name";Expression = {$_.sAMAccountName}},
@{Label = "Display Name";Expression = {$_.DisplayName}},
@{Label = "Last Name";Expression = {$_.Surname}},
@{Label = "Middle Name";Expression = {$_.middleName}},
@{Label = "First Name";Expression = {$_.GivenName}},
@{Label = "Full address";Expression = {$_.StreetAddress}},
@{Label = "City";Expression = {$_.City}},
@{Label = "State";Expression = {$_.st}},
@{Label = "Post Code";Expression = {$_.PostalCode}},
@{Label = "Country/Region";Expression = {if (($_.Country -eq 'GB')  ) {'United Kingdom'} Else {$_.Country}}},
@{Label = "Job Title";Expression = {$_.Title}},
@{Label = "Company";Expression = {$_.Company}},
@{Label = "Directorate";Expression = {$_.Description}},
@{Label = "Department";Expression = {$_.Department}},
@{Label = "Employee Type";Expression = {$_.employeeType}},
@{Label = "Employee Number";Expression = {$_.EmployeeNumber}},
@{Label = "Office";Expression = {$_.physicalDeliveryOfficeName}},
@{Label = "Phone";Expression = {$_.telephoneNumber}},
@{Label = "Mobile Phone";Expression = {$_.mobile}},
@{Label = "Group Membership";Expression = {($_.MemberOf | CleanDistinguishedName) -join ", "}},
@{Label = "Email";Expression = {$_.Mail}},
@{Label = "Mail Store";Expression = {($_.homeMDB).SubString(3,($_.homeMDB.Indexof(",")-3))}},
@{Label = "Mailbox Location";Expression = {If($_.msExchMailboxGuid){If ($_.targetAddress -match "onmicrosoft.com" -and $_.Mail){"Online"}Else{"On-Premise"}}}},
@{Label = "Manager";Expression = {(($_.Manager -split ",")[0] -replace "CN=","" )}},
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
@{Label = "Distinguished Name";Expression = {($_.distinguishedName)}} 
 
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
$excel = $output | Export-Excel -Path $xlsxfile -ClearSheet -WorksheetName ("AD_Users_" + $FileDate) -AutoFilter -AutoSize -FreezeTopRowFirstColumn -PassThru
$ws = $excel.Workbook.Worksheets[("AD_Users_" + $FileDate)]
$LastRow = $ws.Dimension.End.Row
$LastColumn = $ws.Dimension.End.column

#Header Lookup
$htHeader =[ordered]@{}
for ($i = 1; $i -le  $LastColumn; $i++) {
	$htHeader.add(($ws.Cells[1,$i].value),$i)
}

#Days Since Last LogOn
# Add-ConditionalFormatting -WorkSheet $ws -address (($ws.Cells[1,$htHeader["Days Since Last LogOn"]].Address).Substring(0,($ws.Cells[1,$htHeader["Days Since Last LogOn"]].Address).Length-1) + "1:" + ($ws.Cells[1,$htHeader["Days Since Last LogOn"]].Address).Substring(0,($ws.Cells[1,$htHeader["Days Since Last LogOn"]].Address).Length-1) + $Lastrow) -RuleType TwoColorScale


#New worksheet that has created,pw change, last logon all over 60 days.
$output | Where-Object {$_."Days Since Last LogOn" -gt 60 -and $_."Days Since Creation" -gt 60 -and $_."Days from last password change" -gt 60} | Export-Excel -Path $xlsxfile -ClearSheet -WorksheetName "Look to Disable" -AutoFilter -AutoSize -FreezeTopRowFirstColumn


Close-ExcelPackage $excel

Remove-Variable "output"
Remove-Variable "excel"
#endregion Excel convert
