<# 
.SYNOPSIS
    Name: Execute_Command_Remote.ps1
    Runs commands on remote computer using PSExec using CSV

.DESCRIPTION
  Runs commands on remote computer using PSExec using CSV


.PARAMETER SourceFiles


.EXAMPLE
   & Execute_Command_Remote.ps1 -Command "shutdown -r -t 00" -computer 127.0.0.1

.NOTES
 Author: Paul Fuller
 Dependencies for this script:
	* PSExec
Changes:
	Version 1.0.1 - Added Username and Password
	Version 1.0.2 - Fixed reverse logic for ALLCSV. 
	Version 1.0.4 - Update Logs to Create sub-folder called Logs for log files 
	Version 1.0.5 - Update to run multible Commands. Also Allow a program to be copied to remote computer
	Version 1.0.6 - Fixed Progress bars and other tweaks.
        Version 1.0.7 - Run remote command as admin 
#>

PARAM (
	[array]$Computers , 
	[string]$ComputerList,
	[string]$PSExecPath   = ".\PsExec.exe",
	[array]$Commands,
	[switch]$ALLCSV 	  = $true,
	[String]$csv_Name     = "Device name",
	[String]$csv_IP       = "IP address",
	[String]$User,
	[String]$Password ,
	[boolean]$Copy
)

$ScriptVersion = "1.0.7"

#############################################################################
#region User Variables
#############################################################################
$LogFile = ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\Logs\" + `
		   ($MyInvocation.MyCommand.Name -replace ".ps1","") + "_" + `
		   (Get-Date -format yyyyMMdd-hhmm) + ".log")
$sw = [Diagnostics.Stopwatch]::StartNew()
$count = 1
$maximumRuntimeSeconds = 30

#############################################################################
#endregion User Variables
#############################################################################

#############################################################################
#region Setup Sessions
#############################################################################
#Check to make sure we have a computer to update.
If ($ComputerList) {
	If ($Computers) {
		throw ("Need options -Computers or -ComputerList set")
	}
}

#Start logging.
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

If (-Not ([string]::IsNullOrEmpty($PSExecPath))) {
	If (-Not (Test-Path $PSExecPath -PathType Leaf)) {
		throw ("psexec.exe is not found at: " + $PSExecPath)
	}
}else{
	throw ("psexec.exe is not found at: " + $PSExecPath)
}

If ($ALLCSV) {
	If (!(Test-Path -Path ($LogFile + "_all.csv"))) {
		Add-Content ($LogFile + "_all.csv") ("Date,Computer,Command,Status")
	}
}

If ($ComputerList) {
	If ( Test-Path $ComputerList ) {
		$ObjCSV = Import-Csv  $ComputerList
		$outItems = New-Object System.Collections.Generic.List[System.Object]
		Foreach ($objline in $ObjCSV) {
			
			#Clean up IP 
            #Write-Host ("Cleaning IP: " + $objline.$csv_IP)
			If (!([string]::IsNullOrEmpty($objline.$csv_IP))) {
				$arrTempIP = ($objline.$csv_IP).Split(".")
				If ($arrTempIP.Count -eq 4) {
					$TempIP = ([String]([int16]$arrTempIP[0]) + "." + [String]([int16]$arrTempIP[1]) + "." + [String]([int16]$arrTempIP[2]) + "." + [String]([int16]$arrTempIP[3]))
				} else {
					$TempIP = $objline.$csv_IP
				}
			}
		    #Testing for duplicate host entries
            If (!($outItems.Contains($objline.$csv_Name))) {		
				$addresslist = $null
				try {
						$addresslist = [Net.DNS]::GetHostEntry($objline.$csv_Name)
						$addresslist = $addresslist.addresslist.ipaddresstostring
					} catch {
						$addresslist = $null
					}
					
				If (!($addresslist)) {
					$IP = $null
					$IP = [ipaddress]$TempIP
					If ($IP.IPAddressToString) {
						Write-Host ("`t Computer does not exists in DNS: " + $objline.$csv_Name + " using IP: " + $IP.IPAddressToString) 
						If ($ALLCSV) {
							If (Test-Path -Path ($LogFile + "_all.csv")) {
								#"Date,Computer,Command,Status"
								Add-Content ($LogFile + "_all.csv") ((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + "," + $Command + ",Does not exists in DNS")
							}
						}
						$outItems.Add($IP.IPAddressToString)
					}
				}else{
					# Write-Host ("`t Computer already exists in DNS: " + ([Net.DNS]::GetHostEntry($objline.$csv_Name).hostname)) 
					# If ($ALLCSV) {
						# If (!(Test-Path -Path ($LogFile + "_all.csv"))) {
							#"Date,Computer,Command,Status"
							# Add-Content ($LogFile + "_all.csv") ((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + "," + $Command + ",Exists in DNS")
						# }
					# }
					#Add to list
					$outItems.Add($objline.$csv_Name)

				}
				
            } else {
				 Write-Warning ("Duplicate entry " + $objline.$csv_Name + " with IP " + $TempIP)
				If ($ALLCSV) {
					If (Test-Path -Path ($LogFile + "_all.csv")) {
						#"Date,Computer,Command,Status"
						Add-Content ($LogFile + "_all.csv") ((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + "," + $Command + ",Duplicate Entry")
					}
				}
            }
		}
		$Computers += $outItems.ToArray()	
		$ObjCSV=$null
	}
}

#############################################################################
#endregion Setup Sessions
#############################################################################
#############################################################################
#region Functions
#############################################################################
Function FormatElapsedTime($ts) 
{
    #https://stackoverflow.com/questions/3513650/timing-a-commands-execution-in-powershell
	$elapsedTime = ""

    if ( $ts.Hours -gt 0 )
    {
        $elapsedTime = [string]::Format( "{0:00} hours {2:00} min. {3:00}.{4:00} sec.", $ts.Hours, $ts.Minutes, $ts.Seconds, $ts.Milliseconds / 10 );
    }else {
        if ( $ts.Minutes -gt 0 )
        {
            $elapsedTime = [string]::Format( "{0:00} min. {1:00}.{2:00} sec.", $ts.Minutes, $ts.Seconds, $ts.Milliseconds / 10 );
        }
        else
        {
            $elapsedTime = [string]::Format( "{0:00}.{1:00} sec.", $ts.Seconds, $ts.Milliseconds / 10 );
        }

        if ($ts.Hours -eq 0 -and $ts.Minutes -eq 0 -and $ts.Seconds -eq 0)
        {
            $elapsedTime = [string]::Format("{0:00} ms.", $ts.Milliseconds);
        }

        if ($ts.Milliseconds -eq 0)
        {
            $elapsedTime = [string]::Format("{0} ms", $ts.TotalMilliseconds);
        }
    }
    return $elapsedTime
}
Function Start-PSExec()
{
	param(
		[Parameter(Mandatory=$true)][string]$Computer,
		[Parameter(Mandatory=$true)][string]$Command,
		[Parameter(Mandatory=$true)][string]$PSExecPath,
		[Parameter(Mandatory=$false)]$maximumRuntimeSeconds = 30,
		[Parameter(Mandatory=$false)][string]$User,
		[Parameter(Mandatory=$false)][string]$Pass, 
		[Parameter(Mandatory=$false)][bool]$Copy

		)

	If ($Command) {
		Write-Host ("`t`t Running program: " + $Command)
		if ( $User -and $Pass) {
			If ($Copy) {
				$process = Start-Process -FilePath $PSExecPath -ArgumentList $("\\" + $Computer + " -h -c -v -i -accepteula -nobanner -u " + $User + " -p " + $Pass + " " + $Command) -PassThru -NoNewWindow
			} else {
				$process = Start-Process -FilePath $PSExecPath -ArgumentList $("\\" + $Computer + " -h -i -accepteula -nobanner -u " + $User + " -p " + $Pass + " " + $Command) -PassThru -NoNewWindow
			}
		}else{
			If ($Copy) {
				$process = Start-Process -FilePath $PSExecPath -ArgumentList $("\\" + $Computer + " -h -c -v -i -accepteula -nobanner " + $Command) -PassThru -NoNewWindow
			} else {
				$process = Start-Process -FilePath $PSExecPath -ArgumentList $("\\" + $Computer + " -h -i -accepteula -nobanner " + $Command) -PassThru -NoNewWindow
			}
		}
		try 
		{
			$process | Wait-Process -Timeout $maximumRuntimeSeconds -ErrorAction Stop 
			If ($process.ExitCode -le 0) {
				Write-Host ("`t`t PSExec successfully completed within timeout.")
				If ($ALLCSV) {
					If (Test-Path -Path ($LogFile + "_all.csv")) {
						#"Date,Computer,Command,Status"
						Add-Content ($LogFile + "_all.csv") ((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + "," + $Command + ",success")
					}
				}
			}else{
				Write-Warning -Message $('PSExec could not run command. Exit Code: ' + $process.ExitCode)
				If ($ALLCSV) {
					If (Test-Path -Path ($LogFile + "_all.csv")) {
						#"Date,Computer,Command,Status"
						Add-Content ($LogFile + "_all.csv") ((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + "," + $Command + ",Failed Error:" + $process.ExitCode)
					}
				}
				continue
			}
		}catch{
			Write-Warning -Message 'PSExec exceeded timeout, will be killed now.' 
			If ($ALLCSV) {
				If (Test-Path -Path ($LogFile + "_all.csv")) {
					#"Date,Computer,Command,Status"
					Add-Content ($LogFile + "_all.csv") ((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + "," + $Command + ",Timed Out")
				}
			}
			$process | Stop-Process -Force
			continue
		} 
	}else{
		Write-Warning -Message "`t`t NO Commands"
	}
	
}
#############################################################################
#endregion Functions
#############################################################################
#############################################################################
#region Main
#############################################################################
#Loop thru servers
Foreach ($CurrentComputer in $Computers) {
	Write-Progress  -Id 0 -Activity ("Processing Computer") -Status ("( " + $count + "\" + $Computers.count + "): " + $CurrentComputer) -percentComplete ($count / $Computers.count*100)
	Write-Host ("( " + $count + "\" + $Computers.count + ") Computer: " + $CurrentComputer)

	If (Test-Connection -Cn $CurrentComputer -BufferSize 16 -Count 1 -ea 0 -quiet) {
		Write-Host ("`t Host $CurrentComputer on.") -foregroundcolor green
		$countCommands = 1
		ForEach ($Command in $Commands) {
			Write-Progress  -Id 1 -Activity ("Processing Command") -Status ("( " + $countCommands + "\" + $Commands.count + "): " + $Command) -percentComplete ($countCommands / $Commands.count*100)
			if ( $User -and $Password) {
				if ($Copy) {
					Start-PSExec -Computer $CurrentComputer -Command $Command -PSExecPath $PSExecPath -maximumRuntimeSeconds $maximumRuntimeSeconds -User $User -Pass $Password -Copy
				} else {
					Start-PSExec -Computer $CurrentComputer -Command $Command -PSExecPath $PSExecPath -maximumRuntimeSeconds $maximumRuntimeSeconds -User $User -Pass $Password
				}
			}else{
				if ($Copy) {
					Start-PSExec -Computer $CurrentComputer -Command $Command -PSExecPath $PSExecPath -maximumRuntimeSeconds $maximumRuntimeSeconds -Copy
				} else {
					Start-PSExec -Computer $CurrentComputer -Command $Command -PSExecPath $PSExecPath -maximumRuntimeSeconds $maximumRuntimeSeconds
				}
			}
			$countCommands ++
		}
	} else {
		Write-Host ("`t Host $CurrentComputer off.") -foregroundcolor red
		If ($ALLCSV) {
			If (Test-Path -Path ($LogFile + "_all.csv")) {
				#"Date,Computer,Command,Status"
				Add-Content ($LogFile + "_all.csv") ((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + "," + $Command + ",Host Off")
			}
		}
	}
	$count++
 }
 
$sw.Stop()
Write-Host ("Script took: " + (FormatElapsedTime($sw.Elapsed)) + " to run.")

#############################################################################
#endregion Main
#############################################################################
If (-Not [string]::IsNullOrEmpty($LogFile)) {
	Stop-Transcript
}
