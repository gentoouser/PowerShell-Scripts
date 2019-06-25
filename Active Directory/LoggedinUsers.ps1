#region AD log on users
function Get-UserLogon {	 
	[CmdletBinding()]	 
	param	 
	( 
	[Parameter ()]	[String]$Computer,
	[Parameter ()]	[String]$OU,	 
	[Parameter ()]	[Switch]$All,
	[Parameter ()]	[string]$PSExecPath  = "PsExec.exe",
	[Parameter ()]	[string]$Command = "quser",
	[Parameter ()]	[string]$Logoff,
	[Parameter ()]	[string]$Timeout = 15
	)	 
	#https://sid-500.com/2018/02/28/powershell-get-all-logged-on-users-per-computer-ou-domain-get-userlogon/
	$ErrorActionPreference="SilentlyContinue"	 
	#Runspace outline from https://gist.github.com/proxb/6bc718831422df3392c4
	# Create an empty array that we'll use later
	$RunspaceCollection = @()
	# This is the array we want to ultimately add our information to
	$Results=[System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
	$RunspaceResults = @()
	# Dynamicaly figure out how many threads to use. 
	$MaxThreads = (((Get-CimInstance -ClassName 'Win32_Processor' | Select-Object -Property "NumberOfCores").NumberOfCores | Measure-Object -sum).Sum * 10)
	# Create a Runspace Pool with a minimum and maximum number of run spaces. (http://msdn.microsoft.com/en-us/library/windows/desktop/dd324626(v=vs.85).aspx)
	$RunspacePool = [RunspaceFactory]::CreateRunspacePool(1,$MaxThreads)
	# Open the RunspacePool so we can use it
	$RunspacePool.Open()

	If ($Computer) {		 
		$qinfo = New-Object System.Diagnostics.ProcessStartInfo
		$qinfo.FileName = $PSExecPath
		$qinfo.RedirectStandardError = $true
		$qinfo.RedirectStandardOutput = $true
		$qinfo.UseShellExecute = $false
		$qinfo.Arguments =$("\\" + $Computer + " -h -accepteula -nobanner " + $Command)
		$p = New-Object System.Diagnostics.Process
		$p.StartInfo = $qinfo
		$p.Start() | Out-Null
		$p.WaitForExit()
		$COutput = $p.StandardOutput.ReadToEnd() 
		ForEach ($CRAW in ($COutput.Split("`r`n") | Select-Object -Skip 1)) {
			$CRAW = $CRAW.trim()
			$CRAW = $CRAW -replace "\s+",' '
			$CRAW = $CRAW -replace '>',''
			$CArray =$CRAW.Split(" ")
			# USERNAME              SESSIONNAME        ID  STATE   IDLE TIME  LOGON TIME
			If ($CArray[3] -eq "Disc") {
				$array= ([ordered]@{
					'User' = $CArray[1]
					'Computer' = $Computer
					'Date' = $CArray[5]
					'Time' = $CArray[6..7] -join ' '
					'SessionName' = "" #No Session name in Dissconnected sessions
					'SessionID' = $CArray[2]
					'State' = $CArray[4]
					'Idle' = "" #No Idle time in Dissconnected sessions
				})
			} Else {
				$array= ([ordered]@{
					'User' = $CArray[1]
					'Computer' = $Computer
					'Date' = $CArray[6]
					'Time' = $CArray[7..8] -join ' '
					'SessionName' = $CArray[2]
					'SessionID' = $CArray[3]
					'State' = $CArray[4]
					'Idle' = $CArray[5]
				})
			}
			$result+=New-Object -TypeName PSCustomObject -Property $array	
		}
	}	 
	If ($OU) {		 
		$Computers=Get-ADComputer -Filter * -SearchBase "$OU" -Properties operatingsystem		 
		$count=$Computers.count
	}
	If ($All) {	 
		$Computers=Get-ADComputer -Filter * -Properties operatingsystem		 
		$count=$Computers.count		 
	}
	If ($Computers) {
		If ($count -gt 20) {		 
			Write-Warning "Search $count computers. This may take some time ..."		 
		}
		$ScriptBlock = {
			PARAM ($Computer,$Results,$PSExecPath,$Command,$Timeout,$Logoff)
			If (Test-Connection -Cn $Computer -BufferSize 16 -Count 1 -ea 0 -quiet) {	 
				$qinfo = New-Object System.Diagnostics.ProcessStartInfo
				$qinfo.FileName = $PSExecPath
				$qinfo.RedirectStandardError = $true
				$qinfo.RedirectStandardOutput = $true
				$qinfo.UseShellExecute = $false
				$qinfo.Arguments =$("\\" + $Computer + " -h -accepteula -nobanner -n " + $Timeout + " " + $Command)
				$p = New-Object System.Diagnostics.Process
				$p.StartInfo = $qinfo
				$p.Start() | Out-Null
				$p.WaitForExit()

				#Get Output
				$COutput = $p.StandardOutput.ReadToEnd() 
				ForEach ($CRAW in ($COutput.Split("`r`n") | Select-Object -Skip 1)) {
					$CRAW = $CRAW.trim()
					$CRAW = $CRAW -replace "\s+",' '
					$CRAW = $CRAW -replace '>',''
					$CArray =$CRAW.Split(" ")
					$RU = $false
					# USERNAME              SESSIONNAME        ID  STATE   IDLE TIME  LOGON TIME
					If ($CArray[2] -eq "Disc") {
						$array= ([ordered]@{
							'User' = $CArray[0]
							'Computer' = $Computer
							'Date' = $CArray[4]
							'Time' = $CArray[5..6] -join ' '
							'SessionName' = "" #No Session name in Dissconnected sessions
							'SessionID' = $CArray[1]
							'State' = $CArray[2]
							'Idle' = "" #No Idle time in Dissconnected sessions
						})
						If ("" -ne $CArray[0] ) {
							$RU = $true
						}
					} Else {
						$array= ([ordered]@{
							'User' = $CArray[0]
							'Computer' = $Computer
							'Date' = $CArray[5]
							'Time' = $CArray[6..7] -join ' '
							'SessionName' = $CArray[1]
							'SessionID' = $CArray[2]
							'State' = $CArray[3]
							'Idle' = $CArray[4]
						})
						If ("" -ne $CArray[0] ) {
							$RU = $true
						}
					}
					If ($RU) {
						$TempPSObject = New-Object -TypeName PSCustomObject -Property $array
						If ($Logoff) {
							foreach ($LO in ($Logoff.Split(","))) {
								If (($CArray[0]).ToLower() -match $LO.ToLower()) {
									$Command = ("logoff " + $CArray[2])
									$qinfo = New-Object System.Diagnostics.ProcessStartInfo
									$qinfo.FileName = $PSExecPath
									$qinfo.RedirectStandardError = $true
									$qinfo.RedirectStandardOutput = $true
									$qinfo.UseShellExecute = $false
									$qinfo.Arguments =$("\\" + $Computer + " -h -accepteula -nobanner -n " + $Timeout + " " + $Command)
									$p = New-Object System.Diagnostics.Process
									$p.StartInfo = $qinfo
									$p.Start() | Out-Null
									$p.WaitForExit()
									#Get Output
									$TempPSObject.State = "Logged off"
								}
							}
						}
						[System.Threading.Monitor]::Enter($Results.syncroot)
						$Results.Add($TempPSObject) | Out-Null
						[System.Threading.Monitor]::Exit($Results.syncroot)									
					}
				}	
			}	 
		}	
		$CCount = $Computers.count
		$CCCount = 0
		Foreach ($Computer in ($Computers | Select-Object DNSHostName).DNSHostName) {
			Write-Progress -ID 0 -Activity "Queuing Computers" -Status $Computer -PercentComplete (($CCCount/$CCount)*100)
			# Create a PowerShell object to run add the script and argument.
			# We first create a Powershell object to use, and simualtaneously add our script block we made earlier, and add our arguement that we created earlier
			$Powershell = [PowerShell]::Create()
			$Powershell.AddScript($ScriptBlock) | Out-Null
			$Powershell.AddArgument($Computer) | Out-Null
			$Powershell.AddArgument($Results) | Out-Null
			$Powershell.AddArgument($PSExecPath) | Out-Null
			$Powershell.AddArgument($Command) | Out-Null
			$Powershell.AddArgument($Timeout) | Out-Null
			$Powershell.AddArgument($Logoff) | Out-Null
			# Specify runspace to use
			# This is what let's us run concurrent and simualtaneous sessions
			$Powershell.RunspacePool = $RunspacePool
			# Create Runspace collection
			# When we create the collection, we also define that each Runspace should begin running
			[Collections.Arraylist]$RunspaceCollection += New-Object -TypeName PSObject -Property @{
				"Runspace" = $PowerShell.BeginInvoke()
				"PowerShell" = $PowerShell  
				"Computer" = $Computer
			}
			$CCCount++
		}
		Write-Progress -ID 0 -Activity "Queuing Computers"  -Completed
		# Now we need to wait for everything to finish running, and when it does go collect our results and cleanup our run spaces
		# We just say that so long as we have anything in our RunspacePool to keep doing work. This works since we clean up each runspace as it completes.
		$RunSCMCount = $RunspaceCollection.Count
		$RSCount = 0
		While($RunspaceCollection) {		
			# Just a simple ForEach loop for each Runspace to get resolved
			Foreach ($Runspace in $RunspaceCollection.ToArray()) {				
				# Here's where we actually check if the Runspace has completed
				If ($Runspace.Runspace.IsCompleted) {		
					Write-Progress -ID 1 -Activity "Waiting for Computers Results" -Status $Runspace.Computer -PercentComplete (($RSCount / $RunSCMCount) * 100)
					# Since it's completed, we get our results here
					$RunspaceResults.Add($Runspace.PowerShell.EndInvoke($Runspace.Runspace)) | Out-Null			
					# Here's where we cleanup our Runspace
					$Runspace.PowerShell.Dispose() | Out-Null
					$RunspaceCollection.Remove($Runspace) | Out-Null	
					$RSCount++		
				} #/If
			} #/ForEach
		} #/While
	}	
	Write-Progress -ID 1 -Activity "Waiting for Computers Results"  -Completed
	# $FinalReturn =  ($Results | Where-object {"" -ne $_.User})
	return $Results
}

$CSVFile = ("login_" + (Get-Date -Format yyyyMMdd-hhmm) + ".csv" )
 Get-UserLogon -All" -Logoff "admin"| Export-Csv $CSVFile -NoTypeInformation
# Get-UserLogon -OU "DC=IT,DC=com" | Export-Csv $CSVFile -NoTypeInformation -Append

#endregion AD log on users
