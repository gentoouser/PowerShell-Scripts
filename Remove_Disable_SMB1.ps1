#region SMB1 remove
[scriptblock]$SMBScriptBlock = {
	If (Get-Command Get-SmbServerConfiguration -errorAction SilentlyContinue) {	
		If ((Get-SmbServerConfiguration).EnableSMB1Protocol ) {
			$SMB1Enabled = $false
			$SMB1Disabled = $true
		}else {
			$SMB1Enabled = $false
			$SMB1Disabled = $false
		}
	}Elseif ((Get-Item HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters | ForEach-Object {Get-ItemProperty $_.pspath}).SMB1 -eq 0) {
		$SMB1Enabled = $false
		$SMB1Disabled = $true
	}Else{
		$SMB1Enabled = $true
		$SMB1Disabled = $false
	}
	If (Get-Command Get-WindowsOptionalFeature -errorAction SilentlyContinue) {
		$WOF = Get-WindowsOptionalFeature -Online -FeatureName "SMB1*" -ErrorAction SilentlyContinue | Where-Object { $_.state -eq "Enabled"}
		If ($WOF) {
			$WOF | Disable-WindowsOptionalFeature -online -NoRestart -WarningAction SilentlyContinue | Out-Null
			$SMB1Disabled = $true
		}
		If (Get-WindowsOptionalFeature -Online -FeatureName "SMB1*" -ErrorAction SilentlyContinue | Where-Object { $_.state -eq "Disabled"}) {
			$SMB1Enabled = $false
		}
	}
	If (Get-Command Get-WindowsCapability -errorAction SilentlyContinue) {
		$WC = Get-WindowsCapability -Online -Name "SMB1*" -ErrorAction SilentlyContinue | Where-Object { $_.state -eq "Enabled"}
		If ($WC) {
			$WC | Remove-WindowsCapability -Online -WarningAction SilentlyContinue | Out-Null
			$SMB1Disabled = $true
		}
		If (Get-WindowsCapability -Online -Name "SMB1*" -ErrorAction SilentlyContinue | Where-Object { $_.state -eq "Disabled"}) {
			$SMB1Enabled = $false
		}
	}
	If (Get-Command Get-WindowsFeature -errorAction SilentlyContinue) {
		$WF = Get-WindowsFeature -Name "SMB1*" -ErrorAction SilentlyContinue | Where-Object { $_.state -eq "Enabled"}
		If ($WF ) {
			$WF | Uninstall-WindowsFeature -Restart:$false -WarningAction SilentlyContinue | Out-Null
			$SMB1Disabled = $true
		}
		If ( Get-WindowsFeature -Name "SMB1*" -ErrorAction SilentlyContinue | Where-Object { $_.state -eq "Disabled"}) {
			$SMB1Enabled = $false
		}
	}
	If (Get-Command Get-SmbServerConfiguration -errorAction SilentlyContinue) {	
		If ((Get-SmbServerConfiguration).EnableSMB1Protocol ) {
			Set-SmbServerConfiguration -EnableSMB1Protocol $false -confirm:$false | Out-Null
			$SMB1Disabled = $true
			$SMB1Enabled = $false
		}else {
			$SMB1Enabled = $false
		}
	}
	If ([Environment]::OSVersion.Version -le (new-object 'Version' 6,2)) {
		If ((Get-Item HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters | ForEach-Object {Get-ItemProperty $_.pspath}).SMB1 -ne 0) {
			Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters" SMB1 -Type DWORD -Value 0 –Force | Out-Null
			$SMB1Disabled = $true
		}
		If ((Get-Item HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters | ForEach-Object {Get-ItemProperty $_.pspath}).SMB1 -eq 0) {
			$SMB1Enabled = $false
		}
	}
	#See if reboot is needed
	If ((Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') `
		-or (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired')`
		-or (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name 'PendingFileRenameOperations' -ErrorAction Ignore)) {
		$RebootNeeded = $true
	}else {
		$RebootNeeded = $false
	}
	#Return Status
	If ( $SMB1Enabled = $false) {
		If ($SMB1Disabled){
			If ($RebootNeeded) {
				return ("Removed/Disabled SMB1, Reboot Needed")
			}else {
				return ("Removed/Disabled SMB1, No Reboot Needed")
			}
			
		}else {
			If ($RebootNeeded) {		
				return ("SMB1 Not Enabled, Reboot Needed")
			}Else{
				return ("SMB1 Not Enabled, No Reboot Needed")
			}
		}
	}else{
		If (Get-Command Get-SmbServerConfiguration -errorAction SilentlyContinue) {	
			If ((Get-SmbServerConfiguration).EnableSMB1Protocol ) {
				If ($RebootNeeded) {
					return ("SMB1 Enabled Needs Manual Fix, Reboot Needed")
				}else {
					return ("SMB1 Enabled Needs Manual Fix, No Reboot Needed")
				}
			}else {
				If ($RebootNeeded) {
					return ("SMB1 Not Enabled, Reboot Needed")
				}else {
					return ("SMB1 Not Enabled, No Reboot Needed")
				}
			}
		}elseIf ((Get-Item HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters | ForEach-Object {Get-ItemProperty $_.pspath}).SMB1 -eq 0) {
			If ($RebootNeeded) {
				return ("SMB1 Not Enabled, Reboot Needed")
			}else {
				return ("SMB1 Not Enabled, No Reboot Needed")
			}
		}else{
			If ($RebootNeeded) {
				return ("SMB1 Enabled Needs Manual Fix, Reboot Needed")
			}else {
				return ("SMB1 Enabled Needs Manual Fix, No Reboot Needed")
			}
		}
	}
}
#Array of IP to fix
$computersIP =@(
"",
""
)
#Runs code on remote computer
Foreach ($IP in $computersIP ) {
	# $output = Invoke-Command -errorAction SilentlyContinue -ComputerName ([System.Net.Dns]::GetHostByAddress($IP).HostName) -ScriptBlock $SMBScriptBlock
	$output = Invoke-Command -errorAction SilentlyContinue -ComputerName ([System.Net.DNS]::GetHostEntry($IP).HostName) -ScriptBlock $SMBScriptBlock
	write-output ( $IP + "," + $output)
}

#endregion SMB1 remove
