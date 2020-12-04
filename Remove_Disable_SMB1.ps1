#region SMB1 remove
[scriptblock]$SMBScriptBlock = {
	If (Get-Command Get-SmbServerConfiguration -errorAction SilentlyContinue) {	
		If ((Get-SmbServerConfiguration).EnableSMB1Protocol ) {
			$SMB1Disabled = $true
			$SMB1Enabled = $false
		}else {
			$SMB1Enabled = $true
			$SMB1Disabled = $false
		}
	}Else {
		$SMB1Disabled = $false
		$SMB1Enabled = $true
	}
	If (Get-Command Get-WindowsOptionalFeature -errorAction SilentlyContinue) {
		$WOF = Get-WindowsOptionalFeature -Online -FeatureName "SMB1*" -ErrorAction SilentlyContinue | Where-Object { $_.state -eq "Enabled"}
		If ($WOF) {
			Write-Host ("`tRemoving up SMB1") -foregroundcolor darkgray
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
			Write-Host ("`tRemoving up SMB1") -foregroundcolor darkgray
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
			Write-Host ("`tRemoving up SMB1") -foregroundcolor darkgray
			$WF | Uninstall-WindowsFeature -Restart:$false -WarningAction SilentlyContinue | Out-Null
			$SMB1Disabled = $true
		}
		If ( Get-WindowsFeature -Name "SMB1*" -ErrorAction SilentlyContinue | Where-Object { $_.state -eq "Disabled"}) {
			$SMB1Enabled = $false
		}
	}
	If (Get-Command Get-SmbServerConfiguration -errorAction SilentlyContinue) {	
		If ((Get-SmbServerConfiguration).EnableSMB1Protocol ) {
			Write-Host ("`Disable up SMB1") -foregroundcolor darkgray
			Set-SmbServerConfiguration -EnableSMB1Protocol $false -confirm:$false
			$SMB1Disabled = $true
			$SMB1Enabled = $false
		}else {
			$SMB1Enabled = $false
		}
	}
	If ($SMB1Disabled = $False -and [Environment]::OSVersion.Version -le (new-object 'Version' 6,1)) {
		If ((Get-Item HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters | ForEach-Object {Get-ItemProperty $_.pspath}).SMB1 -ne 0) {
			Write-Host ("`Disable up SMB1") -foregroundcolor darkgray
			Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters" SMB1 -Type DWORD -Value 0 –Force
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
		$true
	}else {
		$false
	}
	#Return Status
	If ( $SMB1Enabled = $false) {
		If ($SMB1Disabled){
			If ($RebootNeeded) {
				#write-output ($env:computerName + ": Removed/Disabled SMB1 Reboot Needed")
				return ($env:computerName + ": Removed/Disabled SMB1, Reboot Needed")
			}else {
				#write-output ($env:computerName + ": Removed/Disabled SMB1")
				return ($env:computerName + ": Removed/Disabled SMB1, No Reboot Needed")
			}
			
		}else {
			If ($RebootNeeded) {		
				#write-output ($env:computerName + ": SMB1 Not Enabled Reboot Needed")
				return ($env:computerName + ": SMB1 Not Enabled, Reboot Needed")
			}Else{
				#write-output ($env:computerName + ": SMB1 Not Enabled")
				return ($env:computerName + ": SMB1 Not Enabled, No Reboot Needed")
			}
		}
	}else{
		If (Get-Command Get-SmbServerConfiguration -errorAction SilentlyContinue) {	
			If ((Get-SmbServerConfiguration).EnableSMB1Protocol ) {
				If ($RebootNeeded) {
					#write-output ($env:computerName + ": SMB1 Enabled Needs Manual Fix")
					return ($env:computerName + ": SMB1 Enabled Needs Manual Fix, Reboot Needed")
				}else {
					return ($env:computerName + ": SMB1 Enabled Needs Manual Fix, No Reboot Needed")
				}
			}else {
				If ($RebootNeeded) {
					return ($env:computerName + ": SMB1 Not Enabled, Reboot Needed")
				}else {
					return ($env:computerName + ": SMB1 Not Enabled, No Reboot Needed")
				}
			}
		
		}else {
			If ($RebootNeeded) {
				return ($env:computerName + ": SMB1 Enabled Needs Manual Fix, Reboot Needed")
			}else {
				return ($env:computerName + ": SMB1 Enabled Needs Manual Fix, No Reboot Needed")
			}
		}
	}
}

#Array of IPs to fix
$computersIP =@(
"",
""
)
#Loop that invoke the code on remote computer. 
Foreach ($IP in $computersIP ) {
	$output = Invoke-Command -errorAction SilentlyContinue -ComputerName ([System.Net.Dns]::GetHostByAddress($IP).HostName) -ScriptBlock $SMBScriptBlock
	write-output ( $IP + "," + $output)
}

#endregion SMB1 remove
