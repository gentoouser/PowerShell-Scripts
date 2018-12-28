<# Exchange User Maintenance Script
Operations:
	*Find All Disabled Users in Active Directory
	*Copy Directory if Home Drive Exists 
	*Remove Directory if Copy is fine.
Dependencies for this script:
	*Active Directory Module
Changes:

#>
$ScriptVersion = "1.0.0"
#############################################################################
#region User Variables
#############################################################################
#User Home Drive Share
$HomeDriveShare = "\\unc"
$InactiveUsers = "\\unc"
$UserHomePath = $null
$CopyWhat =@("/COPYALL","/B")
$CopyOptions =@("/R:0","/W:0","/NFL","/NDL")
$BlackListACL=@("BUILTIN\Administrators")
$AdminAccountPrefixs= @("Administrator")

$LogFile = ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\" + `
		$MyInvocation.MyCommand.Name + "_" + `
		(Get-Date -format yyyyMMdd-hhmm) + ".log")
#############################################################################
#endregion User Variables
#############################################################################

#############################################################################
#region Setup Sessions
#############################################################################
If (-Not [string]::IsNullOrEmpty($LogFile)) {
	Start-Transcript -Path $LogFile -Append
	Write-Host ("Script: " + $MyInvocation.MyCommand.Name)
	Write-Host ("Version: " + $ScriptVersion)
	Write-Host (" ")
}		

##Load Active Directory Module
# Load AD PSSnapins
If ((Get-Module | Where-Object {$_.Name -Match "ActiveDirectory"}).Count -eq 0 ) {
	Write-Host ("Loading Active Directory Plugins") -foregroundcolor "Green"
	Import-Module "ActiveDirectory"  -ErrorAction SilentlyContinue
} Else {
	Write-Host ("Active Directory Plug-ins Already Loaded") -foregroundcolor "Green"
}

#Get All Disabled accounts in AD
$DisabledAccounts = Search-ADAccount -AccountDisabled
#Test User Home Share
If ($HomeDriveShare.Substring($HomeDriveShare.Length-1) -eq "\") {$HomeDriveShare = $HomeDriveShare.Substring(0,$HomeDriveShare.Length-1)}
If (-not (test-path $HomeDriveShare)) {throw "User share does not exists: $HomeDriveShare"}
#Test Inactive Users Home Share
If ($InactiveUsers.Substring($InactiveUsers.Length-1) -eq "\") {$InactiveUsers = $InactiveUsers.Substring(0,$InactiveUsers.Length-1)}
If (-not (test-path $InactiveUsers)) {throw "User share does not exists: $InactiveUsers"}
#Get Short Domain Name
$DomainName =  (gwmi Win32_NTDomain).DomainName
#############################################################################
#endregion Setup Sessions
#############################################################################

#############################################################################
#region Main
#############################################################################

#MainLoop of Disabled Users
ForEach ($DA in $DisabledAccounts) {
	
	#Setup Valid $UserHomePath
	If (test-path ($HomeDriveShare + "\" + $DA.SamAccountName)) {
		$UserHomePath = ($HomeDriveShare + "\" + $DA.SamAccountName)
		$UHPACLs = (Get-Item $UserHomePath | get-acl).Access | ? inheritanceflags -eq 'none'
	}elseif(test-path ($HomeDriveShare + "\" + $DA.SamAccountName + "$")){
		$UserHomePath = ($HomeDriveShare + "\" + $DA.SamAccountName + "$")
		$UHPACLs = (Get-Item $UserHomePath | get-acl).Access | ? inheritanceflags -eq 'none'
	}
	If ($UserHomePath) {
		Write-Host ("Cleaning user: $($DA.Name)") -foregroundcolor "Gray"
		#Do Cleanup work
		
		#Process ACLs
		ForEach ($UHPACL in $UHPACLs) {
			If ($UHPACL.IdentityReference.value -eq $($DomainName + "\" + $DA.SamAccountName)) {
				#Not counting Self for Non-Inherited ACLs
			}elseif ($BlackListACL.Contains($UHPACL.IdentityReference.value)) {
				#Not counting Black Listed ACLs for Non-Inherited ACLs
			}elseif ($AdminAccountPrefixs | where { $UHPACL.IdentityReference.value -like $($DomainName + "\" + $_ +"*") }) {
				#Not counting Administrative Accounts Listed ACLs for Non-Inherited ACLs
			}else {
				$NBLACL += $UHPACL
			}
		}
		#Test to see if non-black-listed ACL exists
		If ($NBLACL) {
			Write-Host ("`tSkiping copying: $UserHomePath") -foregroundcolor "yellow"
			ForEach ($UHPACL in $NBLACL) {
				Write-Host ("`t`tNon-Inherited ACL: $($UHPACL.IdentityReference) with $($UHPACL.FileSystemRights)") -foregroundcolor "red"
				
			}
		}else{
			#Robocopy
			Write-Host ("`tCopy to Inactive Users:  $UserHomePath to $InactiveUsers\$($DA.SamAccountName)") -foregroundcolor "Green"
			$tempargs = @($UserHomePath,$($InactiveUsers + "\" + $DA.SamAccountName))
			$tempargs += $CopyWhat
			$tempargs += $CopyOptions
			robocopy @($tempargs)
			If ($?) {
				Write-Host ("`t`tCopy Successful; Deleting: $UserHomePath") -foregroundcolor "green"
				Remove-Item $UserHomePath -Force -Recurse
			}
		}
		
		
	}
	
	#Reset $UserHomePath
	$UserHomePath = $null
	$UHPACL = $null
	$NBLACL = @()
}
#############################################################################
#endregion Main
#############################################################################
If (-Not [string]::IsNullOrEmpty($LogFile)) {
	Stop-Transcript
}
