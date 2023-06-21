<# 
.SYNOPSIS
    Name: Execute_Command_Remote.ps1
    Runs commands on remote computer using PSExec using CSV

.DESCRIPTION
  Runs commands on remote computer using PSExec using CSV


.PARAMETER Computers
.PARAMETER ComputerList
.PARAMETER PSExecPath
.PARAMETER Commands
.PARAMETER ALLCSV
.PARAMETER csv_Name
.PARAMETER csv_IP
.PARAMETER User
.PARAMETER Password
.PARAMETER Copy


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
	Version 1.0.5 - Update to run multiple Commands. Also Allow a program to be copied to remote computer
	Version 1.0.6 - Fixed Progress bars and other tweaks.
    Version 1.0.7 - Run remote command as admin 
    Version 1.0.8 - Update to run faster. Use Class for Logging 
#>

PARAM (
	[array]$Computers, 
	[string]$ComputerList,
	[string]$PSExecPath   = "\\github.com\share\PSTools\PsExec.exe",
	[array]$Commands,
	[String]$csv_Name     = "Device name",
	[String]$csv_IP       = "IP address",
	[String]$User,
	[String]$Password,
	[int]$Timeout = 30,
	[switch]$Copy,
	[switch]$ErrorCSV
)

$ScriptVersion = "1.0.8"

#############################################################################
#region User Variables
#############################################################################
$LogFile = ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\Logs\" + `
		   ($MyInvocation.MyCommand.Name -replace ".ps1","") + "_" + `
		   (Get-Date -format yyyyMMdd-hhmm) + ".log")
$CSVLogFile = ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\Logs\" + `
			($MyInvocation.MyCommand.Name -replace ".ps1","") + "_" + `
			(Get-Date -format yyyyMMdd-hhmm) + ".csv")

$sw = [Diagnostics.Stopwatch]::StartNew()
$count = 1
$maximumRuntimeSeconds = 30

$Logs=New-Object System.Collections.ArrayList
Class CSVObject {
	[DateTime]${Date}
	[string]${Computer}
	[string]${Source File}
	[string]${Source File Version}
	[string]${Source File Date}
	[string]${Destination File}
	[string]${Destination File Version}
	[string]${Destination File Date}
	[string]${Log Level}
	[string]${Status}
	[string]${Command}
	[switch]${Copy File}
}
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

#region Start logging.
If (-Not [string]::IsNullOrEmpty($LogFile)) {
	If (-Not( Test-Path (Split-Path -Path $LogFile -Parent))) {
		New-Item -ItemType directory -Path (Split-Path -Path $LogFile -Parent)
        $Acl = Get-Acl (Split-Path -Path $LogFile -Parent)
        $Ar = New-Object system.Security.AccessControl.FileSystemAccessRule('Users', "FullControl", "ContainerInherit, ObjectInherit", "None", "Allow")
        $Acl.SetAccessRule($Ar)
        Set-Acl (Split-Path -Path $LogFile -Parent) $Acl
	}
	Try{
		#Try to make sure we are not logging already.
		Stop-transcript|out-null
	}Catch [System.InvalidOperationException]{
		#Do not care about errors
	} Finally {
		#Do not care about errors
	}
	Try { 
	    Start-Transcript -Path $LogFile -Append
	}Catch{ 
		Stop-transcript
		Start-Transcript -Path $LogFile -Append
	} 
}
#endregion Start logging.
#region Check for PSExec
If(Get-Command PSExec){
	$PSExecPath = (get-command PSExec).source
}Else{
	If (-Not (Test-Path $PSExecPath -PathType Leaf)) {
		throw ("PSExec.exe is not found at: " + $PSExecPath)
	}
}
#endregion Check for PSExec

#region Importing ComputerList
#Check which computer input to use to set $Computers
If ([string]::IsNullOrEmpty($ComputerList)) {
	If ([string]::IsNullOrEmpty($Computers)) {
		throw " -Computers or -ComputerList is required."
	}
}Else{
	If (($ComputerList.Substring($ComputerList.Length - 3)).ToLower() -eq "csv") {
		If ($ComputerList) {
			If (Test-Path -Path $ComputerList) {
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
						$AddressList = $null
						try {
								$AddressList = [Net.DNS]::GetHostEntry($objline.$csv_Name)
								$AddressList = $AddressList.AddressList.ipaddresstostring
						} catch {
								$AddressList = $null
						}
							
						If (!($AddressList)) {
							$IP = $null
							$IP = [IPAddress]$TempIP
							If ($IP.IPAddressToString) {
								Write-Verbose ("`t Computer does not exists in DNS: " + $objline.$csv_Name + " using IP: " + $IP.IPAddressToString) 
								$Log = [CSVObject]::new()
								$Log.Date = Get-Date
								$Log.Computer = $objline.$csv_Name
								$Log.Status = ("Computer does not exists in DNS: " + $objline.$csv_Name + " using IP: " + $IP.IPAddressToString)
								$Log."Log Level" = "Error"
								$Logs.Add($Log) | Out-Null
								$outItems.Add($IP.IPAddressToString)
							}
						}else{
							#Add to list
							$outItems.Add($objline.$csv_Name)
						}
					} else {
						Write-Warning ("Duplicate entry " + $objline.$csv_Name + " with IP " + $TempIP)
						$Log = [CSVObject]::new()
						$Log.Date = Get-Date
						$Log.Computer = $objline.$csv_Name
						$Log.Status = ("Duplicate entry " + $objline.$csv_Name + " with IP " + $TempIP)
						$Log."Log Level" = "Error"
						$Logs.Add($Log) | Out-Null
					}
				}
				$Computers += $outItems.ToArray()	
				$ObjCSV=$null
			}
		}
	}Else{
		[Array]$Computers += Get-Content -Path $ComputerList
	}
}
#endregion Importing ComputerList
#############################################################################
#endregion Setup Sessions
#############################################################################
#############################################################################
#region Functions
#############################################################################
Function FormatElapsedTime($ts) {
    #https://stackoverflow.com/questions/3513650/timing-a-commands-execution-in-powershell
	$elapsedTime = ""
    if ( $ts.Hours -gt 0 ) {
        $elapsedTime = [string]::Format( "{0:00} hours {1:00} min. {2:00}.{3:00} sec.", $ts.Hours, $ts.Minutes, $ts.Seconds, $ts.Milliseconds / 10 );
    }else {
        if ( $ts.Minutes -gt 0 ) {
            $elapsedTime = [string]::Format( "{0:00} min. {1:00}.{2:00} sec.", $ts.Minutes, $ts.Seconds, $ts.Milliseconds / 10 );
        }else{
            $elapsedTime = [string]::Format( "{0:00}.{1:00} sec.", $ts.Seconds, $ts.Milliseconds / 10 );
        }
        if ($ts.Hours -eq 0 -and $ts.Minutes -eq 0 -and $ts.Seconds -eq 0){
            $elapsedTime = [string]::Format("{0:00} ms.", $ts.Milliseconds);
        }
        if ($ts.Milliseconds -eq 0){
            $elapsedTime = [string]::Format("{0} ms", $ts.TotalMilliseconds);
        }
    }
    return $elapsedTime
}
Function Start-PSProgram {
	param(
		[Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0)][string]$Computer,
		[Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=1)][string]$Command,
		[Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=2)][string]$PSExecPath,
		[Parameter(Mandatory=$False,ValueFromPipeline=$true,Position=3)][int]$Timeout = 30,
		[Parameter(Mandatory=$False,ValueFromPipeline=$true,Position=4)][string]$User,
		[Parameter(Mandatory=$False,ValueFromPipeline=$true,Position=5)][string]$Pass,
		[Parameter(Mandatory=$False,ValueFromPipeline=$true,Position=6)][bool]$Copy
		)
		Class CSVObject {
			[DateTime]${Date}
			[string]${Computer}
			[string]${Source File}
			[string]${Source File Version}
			[string]${Source File Date}
			[string]${Destination File}
			[string]${Destination File Version}
			[string]${Destination File Date}
			[string]${Log Level}
			[string]${Status}
			[string]${Command}
			[switch]${Copy File}
		}
		$Logs=New-Object System.Collections.ArrayList
	If ($Command) {
		Write-Host ("`t`t Running program: " + $Command)
		try{
			If ( $User -and $Pass) {
				If ($Copy) {
					$process = Start-Process -FilePath $PSExecPath -ArgumentList $("\\" + $Computer + " -h -c -v -i -accepteula -nobanner -u " + $User + " -p " + $Pass + " " + $Command) -PassThru -NoNewWindow
				}Else{
					$process = Start-Process -FilePath $PSExecPath -ArgumentList $("\\" + $Computer + " -h -i -accepteula -nobanner -u " + $User + " -p " + $Pass + " " + $Command) -PassThru -NoNewWindow
				}
			}Else{
				If ($Copy) {
					$process = Start-Process -FilePath $PSExecPath -ArgumentList $("\\" + $Computer + " -h -c -v -i -accepteula -nobanner " + $Command) -PassThru -NoNewWindow
				}Else{
					$process = Start-Process -FilePath $PSExecPath -ArgumentList $("\\" + $Computer + " -h -i -accepteula -nobanner " + $Command) -PassThru -NoNewWindow
				}
			}
			$process | Wait-Process -Timeout $Timeout -ErrorAction Stop 
			If ($process.ExitCode -le 0) {
				Write-Host ("`t`t PSExec successfully completed within timeout.")
				$Log = [CSVObject]::new()
				$Log.Date = Get-Date
				$Log.Computer = $Computer
				$Log.Status = "Success"
				$Log.Command = $Command
				$Log."Log Level" = "Informational"
				$Logs.Add($Log) | Out-Null
			}else{
				Write-Warning -Message $('PSExec could not run command. Exit Code: ' + $process.ExitCode)
				$Log = [CSVObject]::new()
				$Log.Date = Get-Date
				$Log.Computer = $Computer
				$Log.Status = ("Failed Error: " + $process.ExitCode)
				$Log.Command = $Command
				$Log."Log Level" = "Error"
				$Logs.Add($Log) | Out-Null
				continue
			}
		}catch{
			Write-Warning -Message 'PSExec exceeded timeout, will be killed now.' 
			$Log = [CSVObject]::new()
			$Log.Date = Get-Date
			$Log.Computer = $Computer
			$Log.Status = ("Timed Out")
			$Log."Log Level" = "Error"
			$Log.Command = $Command
			$Logs.Add($Log) | Out-Null
			$process | Stop-Process -Force
			continue
		} 
	}else{
		Write-Warning -Message "`t`t NO Commands"
		$Log = [CSVObject]::new()
		$Log.Date = Get-Date
		$Log.Computer = $Computer
		$Log.Status = "No Commands Given"
		$Log."Log Level" = "Error"
		$Log.Command = $Command
		$Logs.Add($Log) | Out-Null
	}
	return $Logs
}
#############################################################################
#endregion Functions
#############################################################################
#############################################################################
#region Main
#############################################################################
#Loop thru servers
Foreach ($Computer in $Computers) {
	#Reset logging variables
	$NoChange = $false
	$UpdatesNeeded = $false
	$MissingFiles = $false
	$ComputerError = $false
	Write-Progress -ID 0 -Activity ("Resolving Computer Name") -Status ("(" + $count.ToString().PadLeft($Computers.count.ToString().Length - $count.ToString().Length) + "\" + $Computers.count + ") Computer: " + $Computer) -percentComplete ($count / $Computers.count*100)
	Write-Host ("(" + $count.ToString().PadLeft($Computers.count.ToString().Length - $count.ToString().Length) + "\" + $Computers.count + ") Computer: " + $Computer)
	#test for good IPs from host
	$GoodIPs=@()
	If ([bool]($Computer -as [ipaddress])) {
		If (Test-Connection -ComputerName $Computer -BufferSize 16 -Count 1 -Quiet) {
			Write-Host ("`t Is up with IP: " + $Computer) -ForegroundColor Gray
			$GoodIPs += $Computer
		}
	}else{
		Foreach ($IP in ((([Net.DNS]::GetHostEntry($Computer)).AddressList.ipaddresstostring))) {
			If ($IP -ne "127.0.0.1" -and $IP -ne "::1") {
				If (Test-Connection -ComputerName $IP -BufferSize 16 -Count 1 -Quiet) {
					Write-Host ("`t Responds with IP: " + $IP) -ForegroundColor Gray
					$GoodIPs += $IP
				}
			}
		}
	}
	If ($GoodIPs.count -gt 0)	{
		Foreach ($IP in $GoodIPs ) {
			$cCount = 1
			Foreach ($Command in $Commands) {
				If ($Command) {
					Write-Progress -Id 2 -Activity ("Running Commands") -Status ("(" + $cCount.ToString().PadLeft($Commands.count.ToString().Length - $cCount.ToString().Length)  + "\" + $Commands.count + ") Command: " + $Command ) -percentComplete ($cCount / ($Commands.count)*100)
					#Write-Host ("`t`t Running $Command on  $Computer.") -ForegroundColor green
					if ( $User -and $Password) {
						$FunctionOut = Start-PSProgram -Computer $Computer -Command $Command -PSExecPath $PSExecPath -Timeout $Timeout -User $User -Pass $Password -Copy:$Copy
						$Logs.Add($FunctionOut)
					}else{
						$FunctionOut = Start-PSProgram -Computer $Computer -Command $Command -PSExecPath $PSExecPath -Timeout $Timeout -Copy:$Copy
						$Logs.Add($FunctionOut)
					}		
				}
				$cCount++
			}
		}
	} else {
		#Error computer.
		Write-Warning -Message ("Error: $Computer has NO working IP Addresses")
		If (Test-Connection -ComputerName $Computer -Quiet){ 
			Write-Host ("`t`t Host is up") -ForegroundColor green
		}else{
			Write-Warning -Message ("`t`t Host is Down")
		}
		$Log = [CSVObject]::new()
		$Log.Date = Get-Date
		$Log.Computer = $Computer
		# $Log."Source File" = $SourceFileInfo.value.name
		# $Log."Source File Version" = $SourceFileInfo.value.VersionInfo.ProductVersion
		# $Log."Source File Date" = $SourceFileInfo.value.LastWriteTime
		# $Log."Destination File" = $DestinationFileInfo.name
		# $Log."Destination File Version" = $DestinationFileInfo.VersionInfo.ProductVersion
		# $Log."Destination File Date" = $DestinationFileInfo.LastWriteTime
		$Log.Status = ("Cannot access host")
		$Log."Copy File" = $False
		$Log."Log Level" = "Error"
		$Logs.Add($Log) | Out-Null
		$ComputerError = $true
	}
	$count++
}
 
 If ($Logs) {
	If ($ErrorCSV) {
		$Logs | Where-Object {$_."Log Level" -eq "Error"} | Export-csv -NoTypeInformation -Path ($CSVLogFile -replace ".csv","_Error.csv")
	}
	$Logs | Export-csv -NoTypeInformation -Path $CSVLogFile
}

$sw.Stop()
Write-Host ("Script took: " + (FormatElapsedTime($sw.Elapsed)) + " to run. Averaging " + '{0:N0}' -f ($count / $sw.Elapsed.TotalMinutes) + " Computers per Minute.")

#############################################################################
#endregion Main
#############################################################################
If (-Not [string]::IsNullOrEmpty($LogFile)) {
	Stop-Transcript
}
