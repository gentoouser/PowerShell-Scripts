<# 
.SYNOPSIS
    Name: File_Deployment.ps1
    Copies files and or newer files to remote computer

.DESCRIPTION
   Copies files and or newer files to remote computer


.PARAMETER SourceFiles
	Array of files to be copied to Destination folder
.PARAMETER Destination
	Path on remove computer to copy files to. Example C:\Program Files (x86)

.PARAMETER Computers
.PARAMETER ComputerList   
.PARAMETER PSKillPath
.PARAMETER PSServicePath
.PARAMETER PSExecPath
.PARAMETER Command
.PARAMETER Program
.PARAMETER Service
.PARAMETER VerboseLog
.PARAMETER ErrorCSV 
.PARAMETER ALLCSV
.PARAMETER UseDate
.PARAMETER Copy
.PARAMETER User
.PARAMETER Password


.EXAMPLE
   & File_Deployment.ps1

.NOTES
 Author: Paul Fuller
 Dependencies for this script:
	* PSKill
	* PSService
 Changes:
	Version 1.0.02 - Added pskill
	Version 1.0.03 - Fixed Issue with loading from file. 
	Version 1.0.04 - Added timeout for pskill. 
	Version 1.0.05 - Testing pskill exit code to see if process is killed.
	Version 1.0.06 - Updated how ComputerList is imported.
	Version 1.0.07 - Add more debugging to non zero pskill return. 
	Version 1.0.08 - Updated code to copy if no processes to kill.
	Version 1.0.09 - Update to log pskill computers in failed log. 
	Version 1.0.10 - Added Counter. 
	Version 1.0.11 - Added checking for if $Computers is a string. 
	Version 1.0.12 - Added trying IP address instead of DNS. 
	Version 1.0.13 - Updated progress bar info. 
	Version 1.1.00 - Loop thru all DNS IP Address's for host. 
	Version 1.2.00 - Added ability to copy multiple files and Stop a service. 
	Version 1.2.01 - Fixed File Looping issue. 
	Version 1.2.02 - Added Service Start after copy. 
	Version 1.2.03 - Fixed Renaming bug 
	Version 1.3.00 - Fixed Issue to allow other extension besides dll. 
	Version 1.3.01 - Fixed PSService start issue
	Version 1.3.02 - Replaced Resolve-DnsName with[Net.DNS]::GetHostEntry for older PS compatibility. 
	Version 1.3.03 - Added VerboseLog logging. 
	Version 1.4.00 - Added Copy option to allow for just testing the copy. 
				   - Cleaned up code by adding Stop-PSService,Start-PSService and Kill-PSProgram function. 
				   - Cleaned up code by testing computer is up before testing if file exists. 
				   - Added elapsedTime tracking.
	Version 1.4.01 - Fixed typo. 
	Version 1.4.02 - Added More Info to -Copy:$false 
	Version 1.4.03 - Fixed issue with Run time formatting 
	Version 1.5.00 - Added ErrorCSV logging 
				   - Added UseDate testing control
	Version 1.5.01 - Fixed formatting issues for console and new file name. 
				   - Added AllCSV logging 
	Version 1.5.03 - Fixed Issue calling functions 
	Version 1.5.04 - Fixed issue where computer name would not work in ComputerList. 
	Version 1.5.05 - Updated Path for default PSTools Apps 
	Version 1.5.06 - Update Logs to Create sub-folder called Logs for log files 
	Version 1.6.00 - Added the ability to run program on remote computer 
	Version 1.6.01 - Added Logging when admin share cannot be reached 
	Version 1.6.02 - Fixed issue where Command was not running for same or newer files 
	Version 1.6.03 - Fixed logging for Command 
	Version 1.6.04 - Using PSDrive to map remote computer UNC. *** Beta
	Version 1.6.05 - Fixed issue where Copy-Item  was not using correct variable. 
	Version 1.6.06 - Changed how logging is outputted to reduce errors. Updated better status. 
	Version 1.6.07 - Fix false error showing up. Clean up progress bars
	Version 1.6.08 - Fix Log writing errors
	Version 1.6.09 - Fixed issue With Computers Parameter.
	Version 1.6.10 - Changed Commands to and array to allow running multiple commands. 
	Version 1.6.11 - Change how commands are ran.
	Version 1.6.12 - Added logic to detect csv files. Added -Timeout Param
	Version 1.6.13 - Added logic to create destination folder. Also move where commands are ran to run only on hosts that are up.
	Version 1.6.14 - Cleaned up output.
	Version 1.6.15 - Fixed but about creating folder when credentials are not preset
	Version 1.6.16 - Force psexec to run commands as admin
	Version 1.6.17 - Fixed issue with PSKill not using user and password.
	Version 1.7.00 - Switch to Class for CSV output. 
	Version 1.7.01 - Fixed Calling Funcions and logging issues. 
	Version 1.7.02 - Fixed more bugs from updates.
	Version 1.7.03 - Make stopping Program and Service run once for all files. 
	Version 1.7.04 - Change input to list of strings. 
#>

PARAM (
    [Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0)][string[]]$SourceFiles,
    [Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=1)][string]$Destination, 
	[Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=2)][string[]]$Computers, 
	[Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=3)][string]$ComputerList,
	[Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=4)][string[]]$Commands,	
    [Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=5)][string]$Program,    
	[Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=6)][string]$Service,
	[Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=7)][String]$User,
	[Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=8)][String]$Password,
	[Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=9)][String]$csv_Name       = "Device name",
	[Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=10)][String]$csv_IP         = "IP address",
	[Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=11)][int]$Timeout 			= 30,
	[Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=12)][switch]$CommandForeachFile,
	[Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=13)][switch]$ErrorCSV,
	[Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=14)][switch]$UseDate 		= $true,
	[Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=15)][switch]$Copy 			= $true,
    [Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=16)][string]$PSKillPath 	= "\\github.com\PSTools\pskill.exe",    
    [Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=17)][string]$PSServicePath  = "\\github.com\PSTools\PsService.exe",   
	[Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=18)][string]$PSExecPath     = "\\github.com\PSTools\PsExec.exe"
)
$ScriptVersion = "1.7.04"
#############################################################################
#region User Variables
#############################################################################
$SourceFileObjects=@{}
$count = 1
$sw = [Diagnostics.Stopwatch]::StartNew()
$GoodIPs=@()
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
#region Setup logging
#Start logging.
If (-Not [string]::IsNullOrEmpty($ComputerList)) {
	$LogFile = ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\Logs\" + ($MyInvocation.MyCommand.Name -replace ".ps1","") + "_" + (Split-Path -leaf -Path $Destination) + "_" + (Split-Path -leaf -Path $ComputerList) + "_" + (Get-Date -format yyyyMMdd-hhmm) + ".log")
	$CSVLogFile = ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\Logs\" + ($MyInvocation.MyCommand.Name -replace ".ps1","") + "_" + (Split-Path -leaf -Path $Destination ) + "_" + (Split-Path -leaf -Path $ComputerList ) + "_" + (Get-Date -format yyyyMMdd-hhmm) + ".csv")
}Else{
	$LogFile = ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\Logs\" + ($MyInvocation.MyCommand.Name -replace ".ps1","") + "_" + (Split-Path -leaf -Path $Destination ) + "_" + (Get-Date -format yyyyMMdd-hhmm) + ".log")
	$CSVLogFile = ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\Logs\" + ($MyInvocation.MyCommand.Name -replace ".ps1","") + "_" + (Split-Path -leaf -Path $Destination ) + "_" + (Get-Date -format yyyyMMdd-hhmm) + ".csv")
}
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
#endregion Setup logging	
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
#region Getting Source File Info
Foreach ($SourceFile in $SourceFiles) {
	#Check $SourceFile
	If (Test-Path $SourceFile -PathType Leaf) {
		$SourceFileObjects.Add((Split-Path -leaf -Path $SourceFile),(Get-ChildItem $SourceFile))
	}Else{
		Write-Warning -Message  "-SourceFiles is not a valid file: $SourceFile"
	}
}
If ($SourceFileObjects.Count -le 0) {
	throw " No valid source files"
}
#endregion Getting Source File Info
#Use Local version of PSTools if avalible
#region Check for PSKill
If (-Not ([string]::IsNullOrEmpty($Program))) {
	If(Get-Command pskill){
		$PSKillPath = (get-command pskill).source
	}Else{
		If (-Not (Test-Path $PSKillPath -PathType Leaf)) {
			throw ("pskill.exe is not found at: " + $PSKillPath)
		}
	}
}
#endregion Check for PSKill
#region Check for PSService
If (-Not ([string]::IsNullOrEmpty($Service))) {
	If(Get-Command PSService){
		$PSServicePath = (get-command PSService).source
	}Else{
		If (-Not (Test-Path $PSServicePath -PathType Leaf)) {
			throw ("psservice.exe is not found at: " + $PSServicePath)
		}
	}
}
#endregion Check for PSService
#region Check for PSExec
	If(Get-Command PSExec){
		$PSExecPath = (get-command PSExec).source
	}Else{
		If (-Not (Test-Path $PSExecPath -PathType Leaf)) {
			throw ("PSExec.exe is not found at: " + $PSExecPath)
		}
	}
#endregion Check for PSExec
#region Prepare Username and Password for use
if ( $User -and $Password) {
	$Credential = New-Object System.Management.Automation.PSCredential ($User, (ConvertTo-SecureString $Password -AsPlainText -Force))
}
#endregion Prepare Username and Password for use
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
Function Stop-PSService{
	param(
		[Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0)][string]$Computer,
		[Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=1)][string]$Service,
		[Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=2)][string]$PSServicePath,
		[Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=3)][string]$PSKillPath,
		[Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=4)][int]$Timeout = 60,
		[Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=5)][string]$User,
		[Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=6)][string]$Password
	)
	If ($Null -eq $PSServicePath){
		If(Get-Command PSService){
			$PSServicePath = (get-command PSService).source
		}Else{
			If (-Not (Test-Path $PSServicePath -PathType Leaf)) {
				throw ("psservice.exe is not found at: " + $PSServicePath)
			}
		}
	}
	If ($Null -eq $PSKillPath){
		If(Get-Command pskill){
			$PSKillPath = (get-command pskill).source
		}Else{
			If (-Not (Test-Path $PSKillPath -PathType Leaf)) {
				throw ("pskill.exe is not found at: " + $PSKillPath)
			}
		}
	}
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
	#Stops and then kills service on remote computer
	If ($Service) {
		Write-Host ("`t`t Stopping Service: " + $Service)
		try{
			If ($User -and $Password) {
				$process = Start-Process -FilePath $PSServicePath -ArgumentList @("\\" + $Computer, "stop",'"' + $Service + '"' + " -u " + $User + " -p " + $Password) -PassThru -NoNewWindow
			}Else{
				$process = Start-Process -FilePath $PSServicePath -ArgumentList @("\\" + $Computer, "stop",'"' + $Service + '"') -PassThru -NoNewWindow
			}			
			$process | Wait-Process -Timeout $Timeout -ErrorAction Stop 
			If ($process.ExitCode -le 0) {
				Write-Host ("`t`tPSService successfully completed within timeout.")
				$Log = [CSVObject]::new()
				$Log.Date = Get-Date
				$Log.Computer = $Computer
				$Log.Status = "PSService successfully completed within timeout"
				$Log.Command = ("Stop Service: " + $Service)
				$Log."Log Level" = "Informational"
				$Logs.Add($Log) | Out-Null
			}else{
				Write-Warning -Message $('PSService could not kill process. Exit Code: ' + $process.ExitCode)
				$Log = [CSVObject]::new()
				$Log.Date = Get-Date
				$Log.Computer = $Computer
				$Log.Status = ('PSService could not kill process. Exit Code: ' + $process.ExitCode)
				$Log.Command = ("Stop Service: " + $Service)
				$Log."Log Level" = "Error"
				$Logs.Add($Log) | Out-Null
				continue
			}
		}catch{
			Write-Warning -Message "PSService exceeded timeout, will now be killed."		
			$process | Stop-Process -Force
			#Term Program
			If ($Program) {
				Write-Host ("`t`t killing program: " + $Program)
				try{
					If ($User -and $Password) {
						$process = Start-Process -FilePath $PSKillPath -ArgumentList @("-t -nobanner \\" + $Computer + " " + $Program + " -u " + $User + " -p " + $Password) -PassThru -NoNewWindow
					}Else{
						$process = Start-Process -FilePath $PSKillPath -ArgumentList @("-t -nobanner \\" + $Computer + " " + $Program) -PassThru -NoNewWindow
					}
					$process | Wait-Process -Timeout $Timeout -ErrorAction Stop 
					If ($process.ExitCode -le 0) {
						Write-Host ("`t`tPSKill successfully completed within timeout.")
						$Log = [CSVObject]::new()
						$Log.Date = Get-Date
						$Log.Computer = $Computer
						$Log.Status = "PSService Failed but PSKill completed within timeout"
						$Log.Command = ("Kill: " + $Program)
						$Log."Log Level" = "Informational"
						$Logs.Add($Log) | Out-Null
					}else{
						Write-Warning -Message $('PSKill could not kill process. Exit Code: ' + $process.ExitCode)
						$Log = [CSVObject]::new()
						$Log.Date = Get-Date
						$Log.Computer = $Computer
						$Log.Status = ("PSService Failed but PSKill Failed. Exit Code: " + $process.ExitCode)
						$Log.Command = ("Kill: " + $Program)
						$Log."Log Level" = "Error"
						$Logs.Add($Log) | Out-Null
						continue
					}
				}catch{
					Write-Warning -Message 'PSKill exceeded timeout, will be killed now.' 
					$process | Stop-Process -Force
					$Log = [CSVObject]::new()
					$Log.Date = Get-Date
					$Log.Computer = $Computer
					$Log.Status = ("PSService Failed but PSKill exceeded timeout")
					$Log.Command = ("Kill: " + $Program)
					$Log."Log Level" = "Error"
					$Logs.Add($Log) | Out-Null
					continue
				} 
			}Else{
				Write-Warning -Message 'PSKill exceeded timeout, will be killed now.' 
				$process | Stop-Process -Force
				$Log = [CSVObject]::new()
				$Log.Date = Get-Date
				$Log.Computer = $Computer
				$Log.Status = ("PSService Failed Exit Code: " + $process.ExitCode)
				$Log.Command = ("Stop Service: " + $Service)
				$Log."Log Level" = "Error"
				$Logs.Add($Log) | Out-Null
			}
		} 
	}
	return $Logs
}
Function Start-PSService{
	param(
		[Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0)][string]$Computer,
		[Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=1)][string]$Service,
		[Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=2)][string]$PSServicePath,
		[Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=4)][int]$Timeout = 60,
		[Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=5)][string]$User,
		[Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=6)][string]$Password
	)

	If ($Null -eq $PSServicePath){
		If(Get-Command PSService){
			$PSServicePath = (get-command PSService).source
		}Else{
			If (-Not (Test-Path $PSServicePath -PathType Leaf)) {
				throw ("psservice.exe is not found at: " + $PSServicePath)
			}
		}
	}

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
	#Starts service on remote computer
	If ($Service) {
		Write-Host ("`t`t Stopping Service: " + $Service)
		If ($User -and $Password) {
			$process = Start-Process -FilePath $PSServicePath -ArgumentList @("\\" + $Computer, "start",'"' + $Service + '"' + " -u " + $User + " -p " + $Password) -PassThru -NoNewWindow
		}Else{
			$process = Start-Process -FilePath $PSServicePath -ArgumentList @("\\" + $Computer, "start",'"' + $Service + '"') -PassThru -NoNewWindow
		}
		$process | Wait-Process -Timeout $Timeout -ErrorAction Stop 
		If ($process.ExitCode -le 0) {
			Write-Host ("`t`tPSService Start successfully completed within timeout.")
			$Log = [CSVObject]::new()
			$Log.Date = Get-Date
			$Log.Computer = $Computer
			$Log.Status = "PSService successfully started Service completed within timeout"
			$Log.Command = ("Start Service: " + $Service)
			$Log."Log Level" = "Informational"
			$Logs.Add($Log) | Out-Null
		}else{
			Write-Warning -Message $('PSService Failed to start service. Exit Code: ' + $process.ExitCode)
			$Log = [CSVObject]::new()
			$Log.Date = Get-Date
			$Log.Computer = $Computer
			$Log.Status = ('PSService Failed to start service. Exit Code: ' + $process.ExitCode)
			$Log.Command = ("Start Service: " + $Service)
			$Log."Log Level" = "Error"
			$Logs.Add($Log) | Out-Null
			continue
		}
	}
	Return $Logs
}
Function Kill-PSProgram{
	param(
		[Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0)][string]$Computer,
		[Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=1)][string]$Program,
		[Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=2)][string]$PSKillPath,
		[Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=4)][int]$Timeout = 60,
		[Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=5)][string]$User,
		[Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=6)][string]$Password
	)
	If ($Null -eq $PSKillPath){
		If(Get-Command pskill){
			$PSKillPath = (get-command pskill).source
		}Else{
			If (-Not (Test-Path $PSKillPath -PathType Leaf)) {
				throw ("pskill.exe is not found at: " + $PSKillPath)
			}
		}
	}
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
	#Kills Process on remote computer
	If ($Program) {
		Write-Host ("`t`t killing program: " + $Program)		
		try{
			If ($User -and $Password) {
				$process = Start-Process -FilePath $PSKillPath -ArgumentList $("-t -nobanner \\" + $Computer + " " + $Program + " -u " + $User + " -p " + $Password) -PassThru -NoNewWindow
			}Else{
				$process = Start-Process -FilePath $PSKillPath -ArgumentList $("-t -nobanner \\" + $Computer + " " + $Program) -PassThru -NoNewWindow
			}
			$process | Wait-Process -Timeout $Timeout -ErrorAction Stop 
			If ($process.ExitCode -le 0) {
				Write-Host ("`t`tPSKill successfully completed within timeout.")
				$Log = [CSVObject]::new()
				$Log.Date = Get-Date
				$Log.Computer = $Computer
				$Log.Status = "PSKill successfully completed within timeout"
				$Log.Command = ("Stop Program: " + $Program)
				$Log."Log Level" = "Informational"
				$Logs.Add($Log) | Out-Null
			}else{
				Write-Warning -Message $('PSKill could not kill process. Exit Code: ' + $process.ExitCode)
				$Log = [CSVObject]::new()
				$Log.Date = Get-Date
				$Log.Computer = $Computer
				$Log.Status = ('PSKill could not kill process. Exit Code: ' + $process.ExitCode)
				$Log.Command = ("Stop Program: " + $Program)
				$Log."Log Level" = "Error"
				$Logs.Add($Log) | Out-Null
				continue
			}
		}catch{
			Write-Warning -Message 'PSKill exceeded timeout, will be killed now.' 
			$process | Stop-Process -Force
			$Log = [CSVObject]::new()
			$Log.Date = Get-Date
			$Log.Computer = $Computer
			$Log.Status = 'PSKill exceeded timeout, will be killed now'
			$Log.Command = ("Stop Program: " + $Program)
			$Log."Log Level" = "Error"
			$Logs.Add($Log) | Out-Null
			continue
		} 
	}
	Return $Logs
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
Function New-RemoteFolder {
	param (
		[Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0)][string]$Destination,
		[Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=1)][PSCredential]$Credential
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

    #Make sure we are going to a share
    If($Destination.Substring(0,2) -eq "\\") {
        $ArrayPath = ($Destination -split "\\")
        $Computer = $ArrayPath[2]       
        #Remove old Mapping
        if (Test-Path ("PSFRF:\") -ErrorAction SilentlyContinue) {
            Remove-PSDrive -Name "PSFRF" -Force | Out-Null
        } 

        #Map to the share
        New-PSDrive -Name "PSFRF" -PSProvider "FileSystem" -Root ($ArrayPath[0..3] -Join "\") -Credential $Credential -ErrorAction SilentlyContinue | Out-Null
        If (Test-Path "PSFRF:\"  -ErrorAction SilentlyContinue) {
            $error = New-Item -ItemType "directory" -Path ("PSFRF:\" + ($ArrayPath[0..($ArrayPath.GetUpperBound(0))] -Join "\")) -Force -ErrorAction SilentlyContinue
            If(Test-Path -Path ("PSFRF:\" + ($ArrayPath[4..($ArrayPath.GetUpperBound(0))] -Join "\"))) {
                $Log = [CSVObject]::new()
                $Log.Date = Get-Date
                $Log.Computer = $Computer
                $Log.Status = "Success to Create Folder"
                $Log.Command = ("Create: " + $Destination)
                $Log."Log Level" = "Informational"
                $Logs.Add($Log) | Out-Null
                Remove-PSDrive -Name "PSFRF" -Force | Out-Null
            }Else{
                Write-Warning ("Error creating folder: " + $Destination) 
                $Log = [CSVObject]::new()
                $Log.Date = Get-Date
                $Log.Computer = $Computer
                $Log.Status = ("Failed to Create Folder: " + $error)
                $Log.Command = ("Create: " + $Destination)
                $Log."Log Level" = "Error"
                $Logs.Add($Log) | Out-Null
                Remove-PSDrive -Name "PSFRF" -Force | Out-Null
            }
        }Else{
            Write-verbose ("Error accessing folder: " + $Destination) 
            $Log = [CSVObject]::new()
            $Log.Date = Get-Date
            $Log.Computer = $Computer
            $Log.Status = ("Error Accessing Folder: " + $Destination) 
            $Log.Command = ("Create: " + $Destination)
            $Log."Log Level" = "Error"
            $Logs.Add($Log) | Out-Null
        }
    }Else{
        $Log = [CSVObject]::new()
        $Log.Date = Get-Date
        # $Log.Computer = $Computer
        $Log.Status = "UNC path not given"
        $Log.Command = ("Create: " + $Destination)
        $Log."Log Level" = "Error"
        $Logs.Add($Log) | Out-Null
        Write-Error ("UNC path not given: " + $Destination)
    }
	Return $Logs
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
			#Check to see of remove UNC is mapped
			if (Test-Path "PSRemote:\" -ErrorAction SilentlyContinue) {
				#Remove Existing Mapping 
				Remove-PSDrive -Name "PSRemote"
				If ($LASTEXITCODE) {
					#Remove Existing Mapping 
					if (Test-Path "PSRemote:\" -ErrorAction SilentlyContinue) {
						Remove-PSDrive -Name "PSRemote" -Force
					}
				}
			}
			#Map remote UNC using PSDrive to allow for Credential to be mapped.			
			If ($Credential) {
				#Create Directory if it does not exist
				If (-Not (Test-Path -Path ("\\" +  $IP + "\" + $Destination.replace(":","$")) -Credential $Credential -PathType Container)) {
					$FunctionOut= New-RemoteFolder -Destination ("\\" +  $IP + "\" + $Destination.replace(":","$")) -Credential $Credential
					$Logs.Add($FunctionOut)
				} 
				New-PSDrive -Name "PSRemote" -PSProvider "FileSystem" -Root ("\\" +  $IP + "\" + $Destination.replace(":","$")) -Credential $Credential -ErrorAction SilentlyContinue | out-null
			}else{
				#Create Directory if it does not exist
				If (-Not (Test-Path -Path ("\\" +  $IP + "\" + $Destination.replace(":","$")) -PathType Container)) {
					New-Item -Path ("\\" +  $IP + "\" + $Destination.replace(":","$")) -ItemType "directory" | Out-Null
				} 
				New-PSDrive -Name "PSRemote" -PSProvider "FileSystem" -Root ("\\" +  $IP + "\" + $Destination.replace(":","$")) -ErrorAction SilentlyContinue | out-null
			}		
			#Test Destination Path
			If (Test-Path "PSRemote:\"){
				$fCount = 1
				Write-Progress -ID 0 -Activity ("Copying Files") -Status ("(" + $count.ToString().PadLeft($Computers.count.ToString().Length - $count.ToString().Length) + "\" + $Computers.count + ") Computer: " + $Computer) -percentComplete ($count / $Computers.count*100)
				$RunProgram = $True
				$RunService = $True
				$FilesCopied = $False
				Foreach ($SourceFileInfo in $SourceFileObjects.GetEnumerator()) {
					$DestinationFileInfo = $null
					$NewName = $null
					Write-Progress -Id 1 -Activity ("Testing File") -Status ("(" + $fCount.ToString().PadLeft($SourceFileObjects.count.ToString().Length - $fcount.ToString().Length) + "\" + $SourceFileObjects.count + ") File: " + $SourceFileInfo.value.name ) -percentComplete ($FCount / ($SourceFileObjects.count)*100)
					#Test for File.
					If (Test-Path $("PSRemote:\" + $SourceFileInfo.value.name)) {
						Write-Host ("`tFound at destination: ") -NoNewline -ForegroundColor Green
						Write-Host($("\\" +  $IP + "\" + $Destination.replace(":","$") + "\" + $SourceFileInfo.value.name)) -ForegroundColor Gray
						$DestinationFileInfo = (Get-ChildItem $("PSRemote:\" + $SourceFileInfo.value.name))
						#Test for Version Differences 
						If ($SourceFileInfo.value.VersionInfo.ProductVersion -gt $DestinationFileInfo.VersionInfo.ProductVersion) {
							#Copy newer version
							#Term Service
							If ($Service -and $RunService) {
								Write-Progress -Id 1 -Activity ("Test if we need to Stop: " +$Service) -Status ("(" + $fCount.ToString().PadLeft($SourceFileObjects.count.ToString().Length - $fcount.ToString().Length)  + "\" + $SourceFileObjects.count + ") File: " + $SourceFileInfo.value.name ) -percentComplete ($FCount / ($SourceFileObjects.count)*100)
								Write-Host ("`t Trying to Stop: " + $Service)
								$FunctionOut = Stop-PSService $IP $Service $PSServicePath $PSKillPath $Timeout $User $Password
								$Logs.Add($FunctionOut)
								$RunService = $False
							}
							#Term Program
							If ($Program -and $RunProgram) {
								Write-Progress -Id 1 -Activity ("Test if we need to kill: " + $Program) -Status ("(" + $fCount.ToString().PadLeft($SourceFileObjects.count.ToString().Length - $fcount.ToString().Length)  + "\" + $SourceFileObjects.count + ") File: " + $SourceFileInfo.value.name ) -percentComplete ($FCount / ($SourceFileObjects.count)*100)
								Write-Host ("`t Trying to kill: " + $Program)
								$FunctionOut = Kill-PSProgram $IP $Program $PSKillPath $Timeout $User $Password
								$Logs.Add($FunctionOut)
								$RunProgram = $false
							}
							If ($Copy) {
								#Backup Old 
								Write-Host ("`t`t Renaming destination: " + $Destination + "\" + $DestinationFileInfo.name) -ForegroundColor yellow
								Rename-Item -Path (Get-ChildItem $("PSRemote:\" + $SourceFileInfo.value.name)) -NewName ($DestinationFileInfo.Name.replace("." + $DestinationFileInfo.Extension,"") + "_" + $DestinationFileInfo.VersionInfo.ProductVersion + $DestinationFileInfo.Extension)
								#copy 
								Write-Host ("`t`t Copying new file: ") -NoNewline -ForegroundColor Green
								Write-Host -ForegroundColor Gray ($SourceFileInfo.value.name + " to destination: " + $("\\" + $Destination.replace(":","$")))
								Copy-Item $SourceFileInfo.value -Destination "PSRemote:\"
								$FilesCopied = $true
							} else {
								# newer version
								Write-Host ("`t`t New version: " + $Destination + "\" + $DestinationFileInfo.name) -ForegroundColor green
								Write-Host ("`t`t`t Destination Modified: " + $DestinationFileInfo.LastWriteTime)
								Write-Host ("`t`t`t Destination Version: " + $DestinationFileInfo.VersionInfo.ProductVersion)
							}
							$Log = [CSVObject]::new()
							$Log.Date = Get-Date
							$Log.Computer = $Computer
							$Log."Source File" = $SourceFileInfo.value.name
							$Log."Source File Version" = $SourceFileInfo.value.VersionInfo.ProductVersion
							$Log."Source File Date" = $SourceFileInfo.value.LastWriteTime
							$Log."Destination File" = $DestinationFileInfo.name
							$Log."Destination File Version" = $DestinationFileInfo.VersionInfo.ProductVersion
							$Log."Destination File Date" = $DestinationFileInfo.LastWriteTime
							$Log.Status = "Source file is newer by version"
							$Log."Copy File" = $True
							$Log."Log Level" = "Informational"
							$Logs.Add($Log) | Out-Null
							$UpdatesNeeded = $true
							#Start Service
							if (-Not [string]::IsNullOrEmpty($Service)) {
								$FunctionOut = Start-PSService $IP $Service $PSServicePath $Timeout $User $Password
								$Logs.Add($FunctionOut)
							}
							If ($Commands -and $CommandForeachFile) {
								$cCount = 1
								Foreach ($Command in $Commands) {
									If ($Command) {
										Write-Progress -Id 2 -Activity ("Running Commands") -Status ("(" + $cCount.ToString().PadLeft($Commands.count.ToString().Length - $cCount.ToString().Length)  + "\" + $Commands.count + ") Command: " + $Command ) -percentComplete ($cCount / ($Commands.count)*100)
										if ( $User -and $Password) {
											$FunctionOut = Start-PSProgram -Computer $Computer -Command $Command -PSExecPath $PSExecPath -Timeout $Timeout -User $User -Pass $Password
											$Logs.Add($FunctionOut)
										}else{
											$FunctionOut = Start-PSProgram -Computer $Computer -Command $Command -PSExecPath $PSExecPath -Timeout $Timeout
											$Logs.Add($FunctionOut)
										}	
									}
								}
							}
						}Else{
							If ($UseDate) {
								If ($SourceFileInfo.value.LastWriteTime.ToString("yyyyMMddHHmmssffff") -gt $DestinationFileInfo.LastWriteTime.ToString("yyyyMMddHHmmssffff")) {
									#Term Service
									If ($Service -and $RunService) {
										Write-Progress -Id 1 -Activity ("Test if we need to Stop: " +$Service)  -Status ("(" + $fCount.ToString().PadLeft($SourceFileObjects.count.ToString().Length - $fcount.ToString().Length)  + "\" + $SourceFileObjects.count + ") File: " + $SourceFileInfo.value.name ) -percentComplete ($FCount / ($SourceFileObjects.count)*100)
										Write-Host ("`tTrying to Stop: " + $Service)
										$FunctionOut = Stop-PSService $IP $Service $PSServicePath $PSKillPath $Timeout $User $Password
										$Logs.Add($FunctionOut)
										$RunService = $false
									}
									#Term Program
									If ($Program -and $RunProgram) {
										Write-Progress -Id 1 -Activity ("Test if we need to kill: " + $Program)  -Status ("(" + $fCount.ToString().PadLeft($SourceFileObjects.count.ToString().Length - $fcount.ToString().Length)  + "\" + $SourceFileObjects.count + ") File: " + $SourceFileInfo.value.name ) -percentComplete ($FCount / ($SourceFileObjects.count)*100)
										Write-Host ("`tTrying to kill: " + $Program)
										$FunctionOut = Kill-PSProgram $IP $Program $PSKillPath $Timeout $User $Password
										$Logs.Add($FunctionOut)
										$RunProgram = $false
									}
									If ($Copy) {
										#Backup Old
										Write-Host ("`t`t Renaming destination: " + ($DestinationFileInfo.Name.replace("." + $DestinationFileInfo.Extension,"") + "_" + $DestinationFileInfo.LastWriteTime.ToString("yyyyMMddHHmmssffff") + $DestinationFileInfo.Extension)) -ForegroundColor yellow
										Rename-Item -Path (Get-ChildItem $("PSRemote:\" + $SourceFileInfo.value.name)) -NewName ($DestinationFileInfo.Name.replace("." + $DestinationFileInfo.Extension,"") + "_" + $DestinationFileInfo.LastWriteTime.ToString("yyyyMMddHHmmssffff") + $DestinationFileInfo.Extension)
										#copy 
										Write-Host ("`t`t Copying new file: ") -NoNewline -ForegroundColor Green
										Write-Host -ForegroundColor Gray ($SourceFileInfo.value.name + " to destination: " + $("\\" +  $IP + "\" + $Destination.replace(":","$")) + $SourceFileInfo.value.name)
										Copy-Item $SourceFileInfo.value -Destination "PSRemote:\"
										$FilesCopied = $true
									} else {
										# newer version
										Write-Host ("`t`t New version: " + $Destination + "\" + $DestinationFileInfo.name) -ForegroundColor green
										Write-Host ("`t`t`t Destination Modified: " + $DestinationFileInfo.LastWriteTime)
										Write-Host ("`t`t`t Destination Version: " + $DestinationFileInfo.VersionInfo.ProductVersion)
									}
									$Log = [CSVObject]::new()
									$Log.Date = Get-Date
									$Log.Computer = $Computer
									$Log."Source File" = $SourceFileInfo.value.name
									$Log."Source File Version" = $SourceFileInfo.value.VersionInfo.ProductVersion
									$Log."Source File Date" = $SourceFileInfo.value.LastWriteTime
									$Log."Destination File" = $DestinationFileInfo.name
									$Log."Destination File Version" = $DestinationFileInfo.VersionInfo.ProductVersion
									$Log."Destination File Date" = $DestinationFileInfo.LastWriteTime
									$Log.Status = "Source file is newer by date"
									$Log."Copy File" = $True
									$Log."Log Level" = "Informational"
									$Logs.Add($Log) | Out-Null
									$UpdatesNeeded = $true
									#Start Service
									if (-Not [string]::IsNullOrEmpty($Service)) {
										$FunctionOut = Start-PSService $IP $Service $PSServicePath $Timeout $User $Password
										$Logs.Add($FunctionOut)
									}
									If ($Commands -and $CommandForeachFile) {
										$cCount = 1
										Foreach ($Command in $Commands) {
											If ($Command) {
												Write-Progress -Id 2 -Activity ("Running Commands") -Status ("(" + $cCount.ToString().PadLeft($Commands.count.ToString().Length - $cCount.ToString().Length)  + "\" + $Commands.count + ") Command: " + $Command ) -percentComplete ($cCount / ($Commands.count)*100)
												if ( $User -and $Password) {
													$FunctionOut = Start-PSProgram -Computer $Computer -Command $Command -PSExecPath $PSExecPath -Timeout $Timeout -User $User -Pass $Password
													$Logs.Add($FunctionOut)
												}else{
													$FunctionOut = Start-PSProgram -Computer $Computer -Command $Command -PSExecPath $PSExecPath -Timeout $Timeout
													$Logs.Add($FunctionOut)
												}	
											}
										}
									}
								}else{
									# Older version or same version
									Write-Host ("`t`t Same or Older version: " + $Destination + "\" + $DestinationFileInfo.name) -ForegroundColor DarkGray
									Write-Host ("`t`t`t Destination Modified: " + $DestinationFileInfo.LastWriteTime) -ForegroundColor DarkGray
									Write-Host ("`t`t`t Destination Version:  " + $DestinationFileInfo.VersionInfo.ProductVersion) -ForegroundColor DarkGray
									$NoChange = $true
									$Log = [CSVObject]::new()
									$Log.Date = Get-Date
									$Log.Computer = $Computer
									$Log."Source File" = $SourceFileInfo.value.name
									$Log."Source File Version" = $SourceFileInfo.value.VersionInfo.ProductVersion
									$Log."Source File Date" = $SourceFileInfo.value.LastWriteTime
									$Log."Destination File" = $DestinationFileInfo.name
									$Log."Destination File Version" = $DestinationFileInfo.VersionInfo.ProductVersion
									$Log."Destination File Date" = $DestinationFileInfo.LastWriteTime
									$Log.Status = "Source file is Same or older by date"
									$Log."Copy File" = $False
									$Log."Log Level" = "Informational"
									$Logs.Add($Log) | Out-Null
									If ($Commands -and $CommandForeachFile) {
										$cCount = 1
										Foreach ($Command in $Commands) {
											If ($Command) {
												Write-Progress -Id 2 -Activity ("Running Commands") -Status ("(" + $cCount.ToString().PadLeft($Commands.count.ToString().Length - $cCount.ToString().Length)  + "\" + $Commands.count + ") Command: " + $Command ) -percentComplete ($cCount / ($Commands.count)*100)
												if ( $User -and $Password) {
													$FunctionOut = Start-PSProgram -Computer $Computer -Command $Command -PSExecPath $PSExecPath -Timeout $Timeout -User $User -Pass $Password
													$Logs.Add($FunctionOut)
												}else{
													$FunctionOut = Start-PSProgram -Computer $Computer -Command $Command -PSExecPath $PSExecPath -Timeout $Timeout
													$Logs.Add($FunctionOut)
												}	
											}
										}
									}	
							
								}
							}else{
								# Older version or same version
								Write-Host ("`t`t Same or Older version: " + $Destination + "\" + $DestinationFileInfo.name) -ForegroundColor DarkGray
								Write-Host ("`t`t`t Destination Modified: " + $DestinationFileInfo.LastWriteTime) -ForegroundColor DarkGray
								Write-Host ("`t`t`t Destination Version:  " + $DestinationFileInfo.VersionInfo.ProductVersion) -ForegroundColor DarkGray
								$NoChange = $true
								$Log = [CSVObject]::new()
								$Log.Date = Get-Date
								$Log.Computer = $Computer
								$Log."Source File" = $SourceFileInfo.value.name
								$Log."Source File Version" = $SourceFileInfo.value.VersionInfo.ProductVersion
								$Log."Source File Date" = $SourceFileInfo.value.LastWriteTime
								$Log."Destination File" = $DestinationFileInfo.name
								$Log."Destination File Version" = $DestinationFileInfo.VersionInfo.ProductVersion
								$Log."Destination File Date" = $DestinationFileInfo.LastWriteTime
								$Log.Status = "Source file is Same or older by version"
								$Log."Copy File" = $False
								$Log."Log Level" = "Informational"
								$Logs.Add($Log) | Out-Null
								If ($Commands -and $CommandForeachFile) {
									$cCount = 1
									Foreach ($Command in $Commands) {
										If ($Command) {
											Write-Progress -Id 2 -Activity ("Running Commands") -Status ("(" + $cCount.ToString().PadLeft($Commands.count.ToString().Length - $cCount.ToString().Length)  + "\" + $Commands.count + ") Command: " + $Command ) -percentComplete ($cCount / ($Commands.count)*100)
											if ( $User -and $Password) {
												$FunctionOut = Start-PSProgram -Computer $Computer -Command $Command -PSExecPath $PSExecPath -Timeout $Timeout -User $User -Pass $Password
												$Logs.Add($FunctionOut)
											}else{
												$FunctionOut = Start-PSProgram -Computer $Computer -Command $Command -PSExecPath $PSExecPath -Timeout $Timeout
												$Logs.Add($FunctionOut)
											}	
										}
									}
								}
							}
						}
					}Else{
						#Copy; Missing
						#Term Service
						If ($Service -and $RunService) {
							Write-Progress -ID 1 -Activity ("Test if we need to Stop: " +$Service)  -Status ("(" + $fCount + "\" + $SourceFileObjects.count + ") File: " + $SourceFileInfo.value.name ) -percentComplete ($FCount / ($SourceFileObjects.count)*100)
							Write-Host ("`t Test if we need to Stop: " + $Service)
							$FunctionOut = Stop-PSService $IP $Service $PSServicePath $PSKillPath $Timeout $User $Password
							$Logs.Add($FunctionOut)
							$RunService = $false
						}
						#Term Program
						If ($Program -and $RunProgram) {
							Write-Progress -ID 1 -Activity ("Test if we need to kill: " + $Program)  -Status ("(" + $fCount + "\" + $SourceFileObjects.count + ") File: " + $SourceFileInfo.value.name ) -percentComplete ($FCount / ($SourceFileObjects.count)*100)
							$FunctionOut = Kill-PSProgram $IP $Program $PSKillPath $Timeout $User $Password
							$Logs.Add($FunctionOut)
							$RunProgram = $false
						}
						If ($Copy) {
							Write-Host ("`t`t Copying missing " + $SourceFileInfo.value.name + " to destination: " + $("\\" +  $IP + "\" + $Destination.replace(":","$"))) -ForegroundColor green
							Copy-Item $SourceFileInfo.value -Destination "PSRemote:\"
							$FilesCopied = $true
						}
						$MissingFiles = $true
						$Log = [CSVObject]::new()
						$Log.Date = Get-Date
						$Log.Computer = $Computer
						$Log."Source File" = $SourceFileInfo.value.name
						$Log."Source File Version" = $SourceFileInfo.value.VersionInfo.ProductVersion
						$Log."Source File Date" = $SourceFileInfo.value.LastWriteTime
						# $Log."Destination File" = $DestinationFileInfo.name
						# $Log."Destination File Version" = $DestinationFileInfo.VersionInfo.ProductVersion
						# $Log."Destination File Date" = $DestinationFileInfo.LastWriteTime
						$Log.Status = "Destination File is missing"
						$Log."Copy File" = $True
						$Log."Log Level" = "Informational"
						$Logs.Add($Log) | Out-Null
						#Start Service
						$FunctionOut = Start-PSService $IP $Service $PSServicePath $Timeout $User $Password
						$Logs.Add($FunctionOut)
						#Run Program
						$cCount = 1
						If ($Commands -and $CommandForeachFile) {
							Foreach ($Command in $Commands) {
								If ($Command) {
									Write-Progress -Id 2 -Activity ("Running Commands") -Status ("(" + $cCount.ToString().PadLeft($Commands.count.ToString().Length - $cCount.ToString().Length)  + "\" + $Commands.count + ") Command: " + $Command ) -percentComplete ($cCount / ($Commands.count)*100)
									if ( $User -and $Password) {
										$FunctionOut = Start-PSProgram -Computer $Computer -Command $Command -PSExecPath $PSExecPath -Timeout $Timeout -User $User -Pass $Password
										$Logs.Add($FunctionOut)
									}else{
										$FunctionOut = Start-PSProgram -Computer $Computer -Command $Command -PSExecPath $PSExecPath -Timeout $Timeout
										$Logs.Add($FunctionOut)
									}		
								}
							}
						}
					}
					$fCount ++
				}
				#Run Program at the end
				If ($Commands -and $CommandForeachFile -eq $false -and $FilesCopied) {
					$cCount = 1
					Foreach ($Command in $Commands) {
						If ($Command) {
							Write-Progress -Id 2 -Activity ("Running Commands") -Status ("(" + $cCount.ToString().PadLeft($Commands.count.ToString().Length - $cCount.ToString().Length)  + "\" + $Commands.count + ") Command: " + $Command ) -percentComplete ($cCount / ($Commands.count)*100)
							#Write-Host ("`t`t Running $Command on  $Computer.") -ForegroundColor green
							if ( $User -and $Password) {
								$FunctionOut = Start-PSProgram -Computer $Computer -Command $Command -PSExecPath $PSExecPath -Timeout $Timeout -User $User -Pass $Password
								$Logs.Add($FunctionOut)
							}else{
								$FunctionOut = Start-PSProgram -Computer $Computer -Command $Command -PSExecPath $PSExecPath -Timeout $Timeout
								$Logs.Add($FunctionOut)
							}		
						}
					}
				}	
			} else {
				Write-Warning ("Unable to access admin share of: " + "\\" +  $IP + "\" + $Destination.replace(":","$"))
				$Log = [CSVObject]::new()
				$Log.Date = Get-Date
				$Log.Computer = $Computer
				# $Log."Source File" = $SourceFileInfo.value.name
				# $Log."Source File Version" = $SourceFileInfo.value.VersionInfo.ProductVersion
				# $Log."Source File Date" = $SourceFileInfo.value.LastWriteTime
				# $Log."Destination File" = $DestinationFileInfo.name
				# $Log."Destination File Version" = $DestinationFileInfo.VersionInfo.ProductVersion
				# $Log."Destination File Date" = $DestinationFileInfo.LastWriteTime
				$Log.Status = ("Unable to access admin share of: " + "\\" +  $IP + "\" + $Destination.replace(":","$"))
				$Log."Copy File" = $False
				$Log."Log Level" = "Error"
				$Logs.Add($Log) | Out-Null
			}
		}		
	}Else{
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
#####################################################################
#endregion Main
#############################################################################
If (-Not [string]::IsNullOrEmpty($LogFile)) {
	Stop-Transcript
}
 
