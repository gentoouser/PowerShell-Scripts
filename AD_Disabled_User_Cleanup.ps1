# Version 2.0.5
# Use: Reset users profiles that are Disabled and Compress their Home Drives


#Force Starting of Powershell script as Administrator 
If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))

{   
$arguments = "& '" + $myinvocation.mycommand.definition + "'"
Start-Process powershell -Verb runAs -ArgumentList $arguments
Break
}


#Import AD modules
If ((Get-Module | Where-Object {$_.Name -Match "ActiveDirectory"}).Count -eq 0 ) {
	Write-Host ("Loading Active Directory Plugins") -foregroundcolor "Green"
	Import-Module "ActiveDirectory"  -ErrorAction SilentlyContinue
} Else {
	Write-Host ("Active Directory Plug-ins Already Loaded") -foregroundcolor "Green"
}
#############################################################################
# User Variables
#############################################################################
$ScriptVersion = "2.0.5"
#Number of days to wait
$OlderThan = 90
$CompressedWait = 30
#User Home Drive Share
$HomeDriveShare = ""
#User Roming Profile shares Array.
$RomingProfiles = "","",""
#Archived Users Ignore
$ArchiveIgnoreFolders = "Compressed","Disable Network Users","profiles","App"
#Archive for Users Compressed 
$strArchiveHome = ""
#Archive for Users Compressed 
$strArchiveHomeCompressed = ""
#Archive for Users RDS Profiles
$strArchiveHomeProfiles = ""
#Current Script location
$PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
#XenApp Servers to clean
$XenAppServers ="","","",""
#Gets short Domain Name
$domainName = ((gwmi Win32_ComputerSystem).Domain).Split(".")[0]
#Get Current Date
$StrDate = Get-Date -format yyyyMMdd
#Reset Profiles Folder
$strResetProfiles = "ResetProfiles"
#Reset Permissions
$binGetPers = $False
#Temp Folder
$Temp = [environment]::GetEnvironmentVariable("temp","user")
#Log filename
$LogFile = ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\" + `
		$MyInvocation.MyCommand.Name + "_" + `
		(Get-Date -format yyyyMMdd-hhmm) + ".log")

#############################################################################

## Start of Main script
If (-Not [string]::IsNullOrEmpty($LogFile)) {
	Start-Transcript -Path $LogFile -Append
	Write-Host ("Script: " + $MyInvocation.MyCommand.Name)
	Write-Host ("Version: " + $ScriptVersion)
	Write-Host (" ")
}

cd $strArchiveHome

#7-Zip Set-up
if (-not (test-path "$env:ProgramFiles\7-Zip\7z.exe")) {throw "$env:ProgramFiles\7-Zip\7z.exe needed"} 
set-alias s7zip "$env:ProgramFiles\7-Zip\7z.exe"  
#DelProf2 Set-up
if (-not (test-path "$env:ProgramFiles (x86)\SysinternalsSuite\DelProf2.exe")) {throw "$env:ProgramFiles(x86)\SysinternalsSuite\DelProf2.exe"} 
set-alias sdp2 "$env:ProgramFiles (x86)\SysinternalsSuite\DelProf2.exe"

#Validates strings
If (-Not $strArchiveHome.EndsWith("\")) { $strArchiveHome= $($strArchiveHome + "\")}
If (-Not $HomeDriveShare.EndsWith("\")) { $HomeDriveShare= $($HomeDriveShare + "\")}
If (-Not $strArchiveHomeCompressed.EndsWith("\")) { $strArchiveHomeCompressed= $($strArchiveHomeCompressed + "\")}
If (-Not $Temp.EndsWith("\")) { $Temp= $($Temp + "\")}

$ObjADUsers = Get-AdUser -Filter {(ObjectClass -eq "user")} -SearchBase "OU=Disabled Users,DC=wwtps,DC=com" -Properties DisplayName, EmailAddress, homeDirectory, homeDrive, cn
#Start processing CSV file name
:ADLoop Foreach ($ObjADUser in $ObjADUsers) {
	Write-Host ("Processing User: " + $ObjADUser.DisplayName) -foregroundcolor gray
	#Backup Users roaming profiles
	Foreach ($strRomingProfile in $RomingProfiles) {
		If (-Not $strRomingProfile.EndsWith("\")) { $strRomingProfile= $($strRomingProfile + "\")}
		If (Test-Path $($strRomingProfile + $ObjADUser.SamAccountName)) {
			Write-Host ("`tBacking up and removing User's Roaming Profile: " + $($strRomingProfile + $ObjADUser.SamAccountName))
			s7zip a -mx9 -sfx"$env:ProgramFiles\7-Zip\7z.sfx" -ssw -bd -t7z -mmt -ms=on  $('"' + $strArchiveHomeCompressed + $ObjADUser.SamAccountName + "_" + $StrDate + '.exe"') $('"' + $strRomingProfile + $ObjADUser.SamAccountName + '"')
			If ($LASTEXITCODE -eq 0) {
				Remove-Item $($strRomingProfile + $ObjADUser.SamAccountName) -Force -Recurse
			} else {
				exit
				Break ADLoop
			}
		}
		If (Test-Path $($strRomingProfile + $ObjADUser.SamAccountName + "." + $domainName + ".V2")) {
			Write-Host ("`tBacking up and removing User's Roaming Profile: " + $($strRomingProfile + $ObjADUser.SamAccountName + "." + $domainName + ".V2"))
			s7zip a -mx9 -sfx"$env:ProgramFiles\7-Zip\7z.sfx" -ssw -bd -t7z -mmt -ms=on  $('"' + $strArchiveHomeCompressed + $ObjADUser.SamAccountName + "_" + $StrDate + '.exe"') $('"' + $strRomingProfile + $ObjADUser.SamAccountName + "." + $domainName + '.V2"')
			If ($LASTEXITCODE -eq 0) {
				Remove-Item $($strRomingProfile + $ObjADUser.SamAccountName + "." + $domainName + '.V2') -Force -Recurse
			} else {
				exit
				Break ADLoop
			}
		}
	}
	#Backup Home Data Validate that nothing has changed in the last $OlderThan
	If (Test-Path $($HomeDriveShare + $ObjADUser.SamAccountName )) {
		If ( $(Get-Item $($HomeDriveShare + $ObjADUser.SamAccountName )).LastWriteTime -gt (get-date).AddDays(-$OlderThan)) {
			Write-Host ("`tBacking up User Home Drive: " + $('"' + $HomeDriveShare + $ObjADUser.SamAccountName + '"')) -foregroundcolor "green"
			s7zip a -mx9 -sfx"$env:ProgramFiles\7-Zip\7z.sfx" -ssw -bd -t7z -mmt -ms=on $('"' + $strArchiveHomeCompressed + $ObjADUser.SamAccountName + "_" + $StrDate + '.exe"') $('"' + $HomeDriveShare + $ObjADUser.SamAccountName + '"') 
		}
	}

	#Check for compressed file and see if home drive can be removed
	$objArchiveResults = $(gci $($strArchiveHomeCompressed + $ObjADUser.SamAccountName + "_*"))	
	If (Test-Path $($HomeDriveShare + $ObjADUser.SamAccountName )) {
		$BNRMHome = $False
		$ALLBNRMHome = $False

		Foreach ($objArchive in $objArchiveResults) {
			If ($(Get-Item $($HomeDriveShare + $ObjADUser.SamAccountName )).LastWriteTime -ge $objArchive.LastWriteTime.AddDays(-$CompressedWait)) {
				#In-case of multiple files archives
				$BNRMHome = $True
				$ALLBNRMHome = $True
			}
			$ALLBNRMHome = $False
		}
	}
	If (($BNRMHome -eq $True) -and ($ALLBNRMHome -eq $True) ){
		Write-Host ("`t`tRemoving User's Home Drive: " + $('"' + $HomeDriveShare + $ObjADUser.SamAccountName + '"')) -foregroundcolor "red"
		Remove-Item $( $HomeDriveShare + $ObjADUser.SamAccountName ) -Force -Recurse
	}
	#Combine all archive for user.
	If ($objArchiveResults.Count -gt 1) {
		If (-Not (Test-Path($Temp + $ObjADUser.SamAccountName))) {
			New-Item $($Temp + $ObjADUser.SamAccountName) -type directory
		}
		cd $($Temp + $ObjADUser.SamAccountName)
		Foreach ($objArchive in $objArchiveResults) {
			If ($objArchive.path -ne $($strArchiveHomeCompressed + $ObjADUser.SamAccountName + "_" + $StrDate + '.exe')) {
				Write-Host ("`t`t" + 'Extracting Archives: "' + $objArchive + '" to: "' + $Temp + $ObjADUser.SamAccountName + '"') -foregroundcolor "red"
				s7zip x -y -oua -r $('"'+$objArchive+'"')
				If ($LASTEXITCODE -eq 0) {
					Write-Host ("`t`tRemoveing Archive: " + $objArchive )  -foregroundcolor "red"
					Remove-Item $objArchive -Force -Recurse
				} else {
				exit
				Break ADLoop
				}
			}
		}
		cd $Temp
		Write-Host ("`t`tCompressing recombined files to: " + $('"' + $strArchiveHomeCompressed + $ObjADUser.SamAccountName + "_" + $StrDate + '.exe"')) -foregroundcolor "green"
		s7zip a -mx9 -sfx"$env:ProgramFiles\7-Zip\7z.sfx" -ssw -bd -t7z -mmt -ms=on -r $('"' + $strArchiveHomeCompressed + $ObjADUser.SamAccountName + "_" + $StrDate + '.exe"') $('"' + $Temp + $ObjADUser.SamAccountName + '"')
		If ($LASTEXITCODE -eq 0) {
			Write-Host ("`t`tRemove TEMP files: " + $('"' + $Temp + $ObjADUser.SamAccountName + '"')) -foregroundcolor "yellow"
			Remove-Item $( $Temp + $ObjADUser.SamAccountName ) -Force -Recurse
			If ($LASTEXITCODE -ne 0) {exit}
		} else {
			exit
			Break ADLoop
		}
		cd $strArchiveHome
	}
}
#Clean Cached profiles on XenApp Servers
Foreach ($XAServer in $XenAppServers) {
	#Set-up FQDN
	If ($XAServer.Contains($env:USERDNSDOMAIN)) {
		$XAServerFQDN = $XAServer
	} Else {
		$XAServerFQDN = $($XAServer + "." + $env:USERDNSDOMAIN)
	}
	#Run Delprof2
	sdp2 /u /i /ed:"all users" /ed:default /ed:"default user" /ed:*service /ed:ctx_* /ed:public /ed:*AppPool /ed:*guest* /ed:Anon* /ed:s.* /c:\\$XAServerFQDN

}

If (-Not [string]::IsNullOrEmpty($LogFile)) {
	Stop-Transcript
}
