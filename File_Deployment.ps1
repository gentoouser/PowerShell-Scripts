<# 
.SYNOPSIS
    Name: File_Deployment.ps1
    Copies files and or newer files to remote computer

.DESCRIPTION
   Copies files and or newer files to remote computer


.PARAMETER SourceFiles
	Array of files to be copied to Destination folder
.PARAMETER Destination
	Path on remove computer to copy files to. Exmaple C:\Program Files (x86)

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
	Version 1.0.07 - Add more debuging to non zero pskill return. 
	Version 1.0.08 - Updated code to copy if no processes to kill.
	Version 1.0.09 - Update to log pskill computers in failed log. 
	Version 1.0.10 - Added Counter. 
	Version 1.0.11 - Added checking for if $Computers is a string. 
	Version 1.0.12 - Added trying IP address instead of dns. 
	Version 1.0.13 - Updated progress bar info. 
	Version 1.1.00 - Loop thru all DNS IP Addresss for host. 
	Version 1.2.00 - Added ablity to copy multible files and Stop a servcie. 
	Version 1.2.01 - Fixed File Looping issue. 
	Version 1.2.02 - Added Service Start after copy. 
	Version 1.2.03 - Fixed Renameing bug 
	Version 1.3.00 - Fixed Issue to allow other extension besides dll. 
	Version 1.3.01 - Fixed PSService start issue
	Version 1.3.02 - Replaced Resolve-DnsName with[Net.DNS]::GetHostEntry for older PS compatibility. 
	Version 1.3.03 - Added VerboseLog logging. 
	Version 1.4.00 - Added Copy option to allow for just testing the copy. 
				   - Cleaned up code by adding Stop-PSService,Start-PSService and Kill-PSProgram function. 
				   - Cleaned up code by testing computer is up before testing if file exsits. 
				   - Added elapsedTime tracking.
	Version 1.4.01 - Fixed typo. 
	Version 1.4.02 - Added More Info to -Copy:$false 
	Version 1.4.03 - Fixed issue with Run time formatting 
	Version 1.5.00 - Added ErrorCSV logging 
				   - Added UseDate testing control
	Version 1.5.01 - Fixed formating issues for console and new file name. 
				   - Added AllCSV logging 
	Version 1.5.03 - Fixed Issue calling functions 
	Version 1.5.04 - Fixed issue where computer name would not work in ComputerList. 
	Version 1.5.05 - Updated Path for default PSTools Apps 
	Version 1.5.06 - Update Logs to Create sub-folder called Logs for log files 
	Version 1.6.00 - Added the ablity to run program on remote computer 
	Version 1.6.01 - Added Logging when admin share cannot be reached 
	Version 1.6.02 - Fixed issue where Command was not running for same or newer files 
	Version 1.6.03 - Fixed logging for Command 
	Version 1.6.04 - Using PSDrive to mapp remote computer UNC. *** Beta
	Version 1.6.05 - Fixed issue where Copy-Item  was not using correct variable. 
	Version 1.6.06 - Changed how logging is outputted to reduce errors. Updated better status. 

#>

PARAM (
    [Parameter(Mandatory=$true)][Array]$SourceFiles,
    [Parameter(Mandatory=$true)][string]$Destination, 
	[Array]$Computers, 
    [string]$ComputerList,    
    [string]$PSKillPath 	= ".\pskill.exe",    
    [string]$PSServicePath  = ".\PsService.exe",   
	[string]$PSExecPath     = ".\PsExec.exe",	
	[string]$Command,
    [string]$Program,    
    [string]$Service,    
	[switch]$VerboseLog 	= $false,
	[switch]$ErrorCSV 		= $false,
	[switch]$ALLCSV 		= $true,
	[switch]$UseDate 		= $true,
	[switch]$Copy 			= $true,
	[String]$User		    = $null,
	[String]$Password	    = $null
	
)
$ScriptVersion = "1.6.06"
#############################################################################
#region User Variables
#############################################################################
$SourceFileObjects=@{}
$LogFile = ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\Logs\" + `
		   $MyInvocation.MyCommand.Name + "_" + `
		   (Split-Path -leaf -Path $Destination ) + "_" + `
		   (Get-Date -format yyyyMMdd-hhmm) + ".log")
$maximumRuntimeSeconds = 30
$count = 1
$sw = [Diagnostics.Stopwatch]::StartNew()
$GoodIPs=@()
#############################################################################
#endregion User Variables
#############################################################################

#############################################################################
#region Setup Sessions
#############################################################################
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
#Check which computer input to use to set $Computers
If ([string]::IsNullOrEmpty($ComputerList)) {
	If ([string]::IsNullOrEmpty($Computers)) {
		throw " -Computers or -ComputerList is required."
	}Else{
		# $Computers is already set.
		If (($Computers.GetType().BaseType).Name -eq "String") {
			$Computers = $Computers.split(" ")
		}
	}
}Else{
	[Array]$Computers += Get-Content -Path $ComputerList
}
#Getting Source File Info.
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
#Check for PSKill
If (-Not ([string]::IsNullOrEmpty($Program))) {
	If (-Not (Test-Path $PSKillPath -PathType Leaf)) {
		throw ("pskill.exe is not found at: " + $PSKillPath)
	}
}
#Check for PSService
If (-Not ([string]::IsNullOrEmpty($Service))) {
	If (-Not (Test-Path $PSServicePath -PathType Leaf)) {
		throw ("psservice.exe is not found at: " + $PSServicePath)
	}
}

If ($ErrorCSV) {
	If (!(Test-Path -Path ($LogFile + "_errors.csv"))) {
		# Add-Content ($LogFile + "_errors.csv") ("Date,Computer,Source File,Source File Version,Source file Date,Destination File,Destination File Version,Destination File Date,Error")
		("Date,Computer,Source File,Source File Version,Source file Date,Destination File,Destination File Version,Destination File Date,Error") | Out-File -FilePath ($LogFile + "_errors.csv") -Append
	}
}
If ($ALLCSV) {
	If (!(Test-Path -Path ($LogFile + "_all.csv"))) {
		# Add-Content ($LogFile + "_all.csv") ("Date,Computer,Source File,Source File Version,Source file Date,Destination File,Destination File Version,Destination File Date,Status")
		("Date,Computer,Source File,Source File Version,Source file Date,Destination File,Destination File Version,Destination File Date,Status") | Out-File -FilePath ($LogFile + "_all.csv") -Append
	}
}
if ( $User -and $Password) {
	$Credential = New-Object System.Management.Automation.PSCredential ($User, (ConvertTo-SecureString $Password -AsPlainText -Force))
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
        $elapsedTime = [string]::Format( "{0:00} hours {1:00} min. {2:00}.{3:00} sec.", $ts.Hours, $ts.Minutes, $ts.Seconds, $ts.Milliseconds / 10 );
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
Function Stop-PSService($Computer,$Service,$PSServicePath,$PSKillPath,$maximumRuntimeSeconds) 
{
	#Stops and then kills service on remote computer
	If ($Service) {
		Write-Host ("`t`t Stopping Service: " + $Service)
		$process = Start-Process -FilePath $PSServicePath -ArgumentList @("\\" + $Computer, "stop",'"' + $Service + '"') -PassThru -NoNewWindow
		try 
		{
			$process | Wait-Process -Timeout $maximumRuntimeSeconds -ErrorAction Stop 
			If ($process.ExitCode -le 0) {
				Write-Host ("`t`tPSService successfully completed within timeout.")
			}else{
				Write-Warning -Message $('PSService could not kill process. Exit Code: ' + $process.ExitCode)
				
				continue
			}
		}catch{
			Write-Warning -Message 'PSKill exceeded timeout, will be killed now.' 
			$process | Stop-Process -Force
			#Term Program
			If ($Program) {
				Write-Host ("`t`t killing program: " + $Program)
				$process = Start-Process -FilePath $PSKillPath -ArgumentList $("-t -nobanner \\" + $Computer + " " + $Program) -PassThru -NoNewWindow
				try 
				{
					$process | Wait-Process -Timeout $maximumRuntimeSeconds -ErrorAction Stop 
					If ($process.ExitCode -le 0) {
						Write-Host ("`t`tPSKill successfully completed within timeout.")
					}else{
						Write-Warning -Message $('PSKill could not kill process. Exit Code: ' + $process.ExitCode)
						continue
					}
				}catch{
					Write-Warning -Message 'PSKill exceeded timeout, will be killed now.' 
					$process | Stop-Process -Force
					continue
				} 
			}
		} 
	}
}
Function Start-PSService($Computer,$Service,$PSServicePath,$maximumRuntimeSeconds)
{
	#Starts service on remote computer
	If ($Service) {
		Write-Host ("`t`t Stopping Service: " + $Service)
		$process = Start-Process -FilePath $PSServicePath -ArgumentList @("\\" + $Computer, "start",'"' + $Service + '"') -PassThru -NoNewWindow

		$process | Wait-Process -Timeout $maximumRuntimeSeconds -ErrorAction Stop 
		If ($process.ExitCode -le 0) {
			Write-Host ("`t`tPSService Start successfully completed within timeout.")
		}else{
			Write-Warning -Message $('PSService Start could not kill process. Exit Code: ' + $process.ExitCode)
			
			continue
		}
	}
}
Function Kill-PSProgram($Computer,$Program,$PSKillPath,$maximumRuntimeSeconds)
{
	#Kills Process on remote computer
	If ($Program) {
		Write-Host ("`t`t killing program: " + $Program)
		$process = Start-Process -FilePath $PSKillPath -ArgumentList $("-t -nobanner \\" + $Computer + " " + $Program) -PassThru -NoNewWindow
		try 
		{
			$process | Wait-Process -Timeout $maximumRuntimeSeconds -ErrorAction Stop 
			If ($process.ExitCode -le 0) {
				Write-Host ("`t`tPSKill successfully completed within timeout.")
			}else{
				Write-Warning -Message $('PSKill could not kill process. Exit Code: ' + $process.ExitCode)
				continue
			}
		}catch{
			Write-Warning -Message 'PSKill exceeded timeout, will be killed now.' 
			$process | Stop-Process -Force
			continue
		} 
	}
}
Function Start-PSProgram()
{
	param(
		[Parameter(Mandatory=$true)][string]$Computer,
		[Parameter(Mandatory=$true)][string]$Command,
		[Parameter(Mandatory=$true)][string]$PSExecPath,
		[Parameter(Mandatory=$false)]$maximumRuntimeSeconds = 30,
		[Parameter(Mandatory=$false)][string]$User,
		[Parameter(Mandatory=$false)][string]$Pass
		)

	If ($Command) {
		Write-Host ("`t`t Running program: " + $Command)
		if ( $User -and $Pass) {
			$process = Start-Process -FilePath $PSExecPath -ArgumentList $("\\" + $Computer + " -i -accepteula -nobanner -u " + $User + " -p " + $Pass + " " + $Command) -PassThru -NoNewWindow
		}else{
			$process = Start-Process -FilePath $PSExecPath -ArgumentList $("\\" + $Computer + " -i -accepteula -nobanner " + $Command) -PassThru -NoNewWindow
		}
		try 
		{
			$process | Wait-Process -Timeout $maximumRuntimeSeconds -ErrorAction Stop 
			If ($process.ExitCode -le 0) {
				Write-Host ("`t`t PSExec successfully completed within timeout.")
				If ($ALLCSV) {
					If (Test-Path -Path ($LogFile + "_all.csv")) {
						#"Date,Computer,Source File,Source File Version,Source file Date,Destination File,Destination File Version,Destination File Date,Status")
						# Add-Content ($LogFile + "_all.csv") ((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + ",,,,,,," + $Command + ":success")
						((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + ",,,,,,," + $Command + ":success") | Out-File -Append -FilePath ($LogFile + "_all.csv")
					}
				}
			}else{
				Write-Warning -Message $('PSExec could not run command. Exit Code: ' + $process.ExitCode)
				If ($ALLCSV) {
					If (Test-Path -Path ($LogFile + "_all.csv")) {
						#"Date,Computer,Command,Status"
						# Add-Content ($LogFile + "_all.csv") ((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + ",,,,,,," + $Command + ":Failed Error:" + $process.ExitCode)
						((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + ",,,,,,," + $Command + ":Failed Error:" + $process.ExitCode) | Out-File -Append -FilePath ($LogFile + "_all.csv")
					}
				}
				continue
			}
		}catch{
			Write-Warning -Message 'PSExec exceeded timeout, will be killed now.' 
			If ($ALLCSV) {
				If (Test-Path -Path ($LogFile + "_all.csv")) {
					#"Date,Computer,Command,Status"
					# Add-Content ($LogFile + "_all.csv") ((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + ",,,,,,," + $Command + ":Timed Out")
					((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + ",,,,,,," + $Command + ":Timed Out") | Out-File -Append -FilePath ($LogFile + "_all.csv")
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
Foreach ($Computer in $Computers) {
	#Reset logging variables
	$NoChange = $false
	$UpdatesNeeded = $false
	$MissingFiles = $false
	$ComputerError = $false
	Write-Progress -Activity ("Resolving Computer Name") -Status ("( " + $count + "\" + $Computers.count + ") Computer: " + $Computer) -percentComplete ($count / $Computers.count*100)
	Write-Host ("( " + $count + "\" + $Computers.count + ") Computer: " + $Computer)
	#test for good IPs from host
	$GoodIPs=@()
	If ([bool]($Computer -as [ipaddress])) {
		If (Test-Connection -Cn $Computer -BufferSize 16 -Count 1 -ea 0 -quiet) {
			Write-Host ("`t Is up with IP: " + $Computer)
			$GoodIPs += $Computer
		}
	}else{
		Foreach ($IP in ((([Net.DNS]::GetHostEntry($Computer)).addresslist.ipaddresstostring))) {
			If ($IP -ne "127.0.0.1" -and $IP -ne "::1") {
				If (Test-Connection -Cn $IP -BufferSize 16 -Count 1 -ea 0 -quiet) {
					Write-Host ("`t Responds with IP: " + $IP)
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
					Remove-PSDrive -Name "PSRemote" -Force
				}
			}
			#Map remote UNC using PSDrive to allow for Credential to be mapped.
			If ($Credential) {
				New-PSDrive -Name "PSRemote" -PSProvider "FileSystem" -Root ("\\" +  $IP + "\" + $Destination.replace(":","$")) -Credential $Credential | out-null
			}else{
				New-PSDrive -Name "PSRemote" -PSProvider "FileSystem" -Root ("\\" +  $IP + "\" + $Destination.replace(":","$")) | out-null
			}		
			#Test Destination Path
			If (Test-Path "PSRemote:\"){
				Foreach ($SourceFileInfo in $SourceFileObjects.GetEnumerator()) {
					$DestinationFileInfo = $null
					$NewName = $null
					Write-Progress -Activity ("Testing File: " + (Split-Path -leaf -Path $SourceFileInfo.value.name )) -Status ("( " + $count + "\" + $Computers.count + ") Computer: " + $Computer + " IP: " + $IP + " Runtime: " + (FormatElapsedTime($sw.Elapsed))) -percentComplete ($count / $Computers.count*100)
					#Test for File.
					If (Test-Path $("PSRemote:\" + $SourceFileInfo.value.name)) {
						Write-Host ("`t Found at destination: " + $("\\" +  $IP + "\" + $Destination.replace(":","$") + "\" + $SourceFileInfo.value.name))
						$DestinationFileInfo = (Get-ChildItem $("PSRemote:\" + $SourceFileInfo.value.name))
						#Test for Version Differences 
						If ($SourceFileInfo.value.VersionInfo.ProductVersion -gt $DestinationFileInfo.VersionInfo.ProductVersion) {
							#Copy newer version
							#Term Service
							If ($Service) {
								Write-Progress -Activity ("Test if we need to Stop: " +$Service) -Status ("( " + $count + "\" + $Computers.count + ") Computer: " + $Computer + " IP: " + $IP + " Runtime: " + (FormatElapsedTime($sw.Elapsed))) -percentComplete ($count / $Computers.count*100)
								Write-Host ("`t Test if we need to Stop: " + $Service)
								Stop-PSService $IP $Service $PSServicePath $PSKillPath $maximumRuntimeSeconds
							}
							#Term Program
							If ($Program) {
								Write-Progress -Activity ("Test if we need to kill: " + $Program) -Status ("( " + $count + "\" + $Computers.count + ") Computer: " + $Computer + " IP: " + $IP + " Runtime: " + (FormatElapsedTime($sw.Elapsed))) -percentComplete ($count / $Computers.count*100)
								Write-Host ("`t Test if we need to kill: " + $Program)
								Kill-PSProgram $IP $Program $PSKillPath $maximumRuntimeSeconds
							}
							If ($Copy) {
								#Backup Old 
								Write-Host ("`t`t Renaming destination: " + ($DestinationFileInfo.Name.replace("." + $DestinationFileInfo.Extension,"") + "_" + $DestinationFileInfo.VersionInfo.ProductVersion + $DestinationFileInfo.Extension))
								Rename-Item -Path (Get-ChildItem $("PSRemote:\" + $SourceFileInfo.value.name)) -NewName ($DestinationFileInfo.Name.replace("." + $DestinationFileInfo.Extension,"") + "_" + $DestinationFileInfo.VersionInfo.ProductVersion + $DestinationFileInfo.Extension)
								#copy 
								Write-Host ("`t`t Copying new " + $SourceFileInfo.value.name + " to destination: " + $("\\" + $Destination.replace(":","$")))
								Copy-Item $SourceFileInfo.value -Destination "PSRemote:\"
							} else {
								# newer version
								Write-Host ("`t`t New version: " + ($DestinationFileInfo.Name.replace("." + $DestinationFileInfo.Extension,"") + "_" + $DestinationFileInfo.VersionInfo.ProductVersion + $DestinationFileInfo.Extension)) -foregroundcolor green
								Write-Host ("`t`t`t Destination Modified: " + $DestinationFileInfo.LastWriteTime)
								Write-Host ("`t`t`t Destination Version: " + $DestinationFileInfo.VersionInfo.ProductVersion)
							}
							If ($ErrorCSV) {
								#"Date,Computer,Source File, Source File Version, Source file Date,Destination File,Destination File Version,Destination File Date,Error"
								# Add-Content ($LogFile + "_errors.csv") ((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + "," + $SourceFileInfo.value.name + "," + $SourceFileInfo.value.VersionInfo.ProductVersion + "," + $SourceFileInfo.value.LastWriteTime + "," + $DestinationFileInfo.name + "," + $DestinationFileInfo.VersionInfo.ProductVersion + "," + $DestinationFileInfo.LastWriteTime + ",Source file is newer by version")
								((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + "," + $SourceFileInfo.value.name + "," + $SourceFileInfo.value.VersionInfo.ProductVersion + "," + $SourceFileInfo.value.LastWriteTime + "," + $DestinationFileInfo.name + "," + $DestinationFileInfo.VersionInfo.ProductVersion + "," + $DestinationFileInfo.LastWriteTime + ",Source file is newer by version") | Out-File -Append -FilePath ($LogFile + "_errors.csv")
							}
							If ($AllCSV) {
								#"Date,Computer,Source File, Source File Version, Source file Date,Destination File,Destination File Version,Destination File Date,Status"
								# Add-Content ($LogFile + "_all.csv") ((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + "," + $SourceFileInfo.value.name + "," + $SourceFileInfo.value.VersionInfo.ProductVersion + "," + $SourceFileInfo.value.LastWriteTime + "," + $DestinationFileInfo.name + "," + $DestinationFileInfo.VersionInfo.ProductVersion + "," + $DestinationFileInfo.LastWriteTime + ",Source file is newer by version; copying file")
								((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + "," + $SourceFileInfo.value.name + "," + $SourceFileInfo.value.VersionInfo.ProductVersion + "," + $SourceFileInfo.value.LastWriteTime + "," + $DestinationFileInfo.name + "," + $DestinationFileInfo.VersionInfo.ProductVersion + "," + $DestinationFileInfo.LastWriteTime + ",Source file is newer by version; copying file") | Out-File -Append -FilePath ($LogFile + "_all.csv")
							}
							$UpdatesNeeded = $true
							#Start Service
							Start-PSService $IP $Service $PSServicePath $maximumRuntimeSeconds
							If ($Command) {
								#Write-Host ("`t`t Running $Command on  $Computer.") -foregroundcolor green
								if ( $User -and $Password) {
									Start-PSProgram $Computer $Command $PSExecPath $maximumRuntimeSeconds $User $Password
								}else{
									Start-PSProgram $Computer $Command $PSExecPath $maximumRuntimeSeconds
								}	
							}
							
						}Else{
							If ($UseDate) {
								If ($SourceFileInfo.value.LastWriteTime.ToString("yyyyMMddHHmmssffff") -gt $DestinationFileInfo.LastWriteTime.ToString("yyyyMMddHHmmssffff")) {
									#Term Service
									If ($Service) {
										Write-Progress -Activity ("Test if we need to Stop: " +$Service) -Status ("( " + $count + "\" + $Computers.count + ") Computer: " + $Computer + " IP: " + $IP + " Runtime: " + (FormatElapsedTime($sw.Elapsed))) -percentComplete ($count / $Computers.count*100)
										Write-Host ("`t Test if we need to Stop: " + $Service)
										Stop-PSService $IP $Service $PSServicePath $PSKillPath $maximumRuntimeSeconds
									}
									#Term Program
									If ($Program) {
										Write-Progress -Activity ("Test if we need to kill: " + $Program) -Status ("( " + $count + "\" + $Computers.count + ") Computer: " + $Computer + " IP: " + $IP + " Runtime: " + (FormatElapsedTime($sw.Elapsed))) -percentComplete ($count / $Computers.count*100)
										Write-Host ("`t Test if we need to kill: " + $Program)
										Kill-PSProgram $IP $Program $PSKillPath $maximumRuntimeSeconds
									}
									If ($Copy) {
										#Backup Old
										Write-Host ("`t`t Renaming destination: " + ($DestinationFileInfo.Name.replace("." + $DestinationFileInfo.Extension,"") + "_" + $DestinationFileInfo.LastWriteTime.ToString("yyyyMMddHHmmssffff") + $DestinationFileInfo.Extension))
										Rename-Item -Path (Get-ChildItem $("PSRemote:\" + $SourceFileInfo.value.name)) -NewName ($DestinationFileInfo.Name.replace("." + $DestinationFileInfo.Extension,"") + "_" + $DestinationFileInfo.LastWriteTime.ToString("yyyyMMddHHmmssffff") + $DestinationFileInfo.Extension)
										#copy 
										Write-Host ("`t`t Copying new " + $SourceFileInfo.value.name + " to destination: " + $("\\" +  $IP + "\" + $Destination.replace(":","$")) + $SourceFileInfo.value.name)
										Copy-Item $SourceFileInfo.value -Destination "PSRemote:\"
									} else {
										# newer version
										Write-Host ("`t`t New version: " + ($DestinationFileInfo.Name.replace("." + $DestinationFileInfo.Extension,"") + "_" + $DestinationFileInfo.LastWriteTime.ToString("yyyyMMddHHmmssffff") + $DestinationFileInfo.Extension)) -foregroundcolor green
										Write-Host ("`t`t`t Destination Modified: " + $DestinationFileInfo.LastWriteTime)
										Write-Host ("`t`t`t Destination Version: " + $DestinationFileInfo.VersionInfo.ProductVersion)
									}
									If ($ErrorCSV) {
										#"Date,Computer,Source File, Source File Version, Source file Date,Destination File,Destination File Version,Destination File Date,Error"
										Add-Content ($LogFile + "_errors.csv") ((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + "," + $SourceFileInfo.value.name + "," + $SourceFileInfo.value.VersionInfo.ProductVersion + "," + $SourceFileInfo.value.LastWriteTime + "," + $DestinationFileInfo.name + "," + $DestinationFileInfo.VersionInfo.ProductVersion + "," + $DestinationFileInfo.LastWriteTime + ",Source file is newer by date")
									}
									If ($AllCSV) {
										#"Date,Computer,Source File, Source File Version, Source file Date,Destination File,Destination File Version,Destination File Date,Status"
										Add-Content ($LogFile + "_all.csv") ((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + "," + $SourceFileInfo.value.name + "," + $SourceFileInfo.value.VersionInfo.ProductVersion + "," + $SourceFileInfo.value.LastWriteTime + "," + $DestinationFileInfo.name + "," + $DestinationFileInfo.VersionInfo.ProductVersion + "," + $DestinationFileInfo.LastWriteTime + ",Source file is newer by date; copying file")
									}
									$UpdatesNeeded = $true
									#Start Service
									Start-PSService $IP $Service $PSServicePath $maximumRuntimeSeconds
									If ($Command) {
										#Write-Host ("`t`t Running $Command on  $Computer.") -foregroundcolor green
										if ( $User -and $Password) {
											Start-PSProgram $Computer $Command $PSExecPath $maximumRuntimeSeconds $User $Password
										}else{
											Start-PSProgram $Computer $Command $PSExecPath $maximumRuntimeSeconds
										}	
									}
								}else{
									# Older version or same version
									Write-Host ("`t`t Same or Older version: " + ($DestinationFileInfo.Name.replace("." + $DestinationFileInfo.Extension,"") + "_" + $DestinationFileInfo.LastWriteTime.ToString("yyyyMMddHHmmssffff") + $DestinationFileInfo.Extension)) -foregroundcolor darkgray
									Write-Host ("`t`t`t Destination Modified: " + $DestinationFileInfo.LastWriteTime) -foregroundcolor darkgray
									Write-Host ("`t`t`t Destination Version: " + $DestinationFileInfo.VersionInfo.ProductVersion) -foregroundcolor darkgray
									$NoChange = $true
									If ($AllCSV) {
										#"Date,Computer,Source File, Source File Version, Source file Date,Destination File,Destination File Version,Destination File Date,Status"
										# Add-Content ($LogFile + "_all.csv") ((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + "," + $SourceFileInfo.value.name + "," + $SourceFileInfo.value.VersionInfo.ProductVersion + "," + $SourceFileInfo.value.LastWriteTime + "," + $DestinationFileInfo.name + "," + $DestinationFileInfo.VersionInfo.ProductVersion + "," + $DestinationFileInfo.LastWriteTime + ",Source file is Same or older by date")
										((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + "," + $SourceFileInfo.value.name + "," + $SourceFileInfo.value.VersionInfo.ProductVersion + "," + $SourceFileInfo.value.LastWriteTime + "," + $DestinationFileInfo.name + "," + $DestinationFileInfo.VersionInfo.ProductVersion + "," + $DestinationFileInfo.LastWriteTime + ",Source file is Same or older by date") | Out-File -Append -FilePath ($LogFile + "_all.csv")
									}
                                    If ($Command) {
								        #Write-Host ("`t`t Running $Command on  $Computer.") -foregroundcolor green
								        if ( $User -and $Password) {
									        Start-PSProgram $Computer $Command $PSExecPath $maximumRuntimeSeconds $User $Password
								        }else{
									        Start-PSProgram $Computer $Command $PSExecPath $maximumRuntimeSeconds
                                        }
								    }	
							
								}
							}else{
								# Older version or same version
								Write-Host ("`t`t Same or Older version: " + ($DestinationFileInfo.Name.replace("." + $DestinationFileInfo.Extension,"") + "_" + $DestinationFileInfo.VersionInfo.ProductVersion + $DestinationFileInfo.Extension)) -foregroundcolor darkgray
								Write-Host ("`t`t`t Destination Modified: " + $DestinationFileInfo.LastWriteTime) -foregroundcolor darkgray
								Write-Host ("`t`t`t Destination Version: " + $DestinationFileInfo.VersionInfo.ProductVersion) -foregroundcolor darkgray
								$NoChange = $true
								If ($AllCSV) {
									#"Date,Computer,Source File, Source File Version, Source file Date,Destination File,Destination File Version,Destination File Date,Status"
									# Add-Content ($LogFile + "_all.csv") ((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + "," + $SourceFileInfo.value.name + "," + $SourceFileInfo.value.VersionInfo.ProductVersion + "," + $SourceFileInfo.value.LastWriteTime + "," + $DestinationFileInfo.name + "," + $DestinationFileInfo.VersionInfo.ProductVersion + "," + $DestinationFileInfo.LastWriteTime + ",Source file is Same or older by version")
									((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + "," + $SourceFileInfo.value.name + "," + $SourceFileInfo.value.VersionInfo.ProductVersion + "," + $SourceFileInfo.value.LastWriteTime + "," + $DestinationFileInfo.name + "," + $DestinationFileInfo.VersionInfo.ProductVersion + "," + $DestinationFileInfo.LastWriteTime + ",Source file is Same or older by version") | Out-File -Append -FilePath ($LogFile + "_all.csv")
								}
                                If ($Command) {
								    #Write-Host ("`t`t Running $Command on  $Computer.") -foregroundcolor green
								    if ( $User -and $Password) {
									    Start-PSProgram $Computer $Command $PSExecPath $maximumRuntimeSeconds $User $Password
								    }else{
									    Start-PSProgram $Computer $Command $PSExecPath $maximumRuntimeSeconds
                                    }
								}

							}
						}
					}Else{
						#Copy; Missing
						#Term Service
						If ($Service) {
							Write-Progress -Activity ("Test if we need to Stop: " +$Service) -Status ("( " + $count + "\" + $Computers.count + ") Computer: " + $Computer + " IP: " + $IP + " Runtime: " + (FormatElapsedTime($sw.Elapsed))) -percentComplete ($count / $Computers.count*100)
							Write-Host ("`t Test if we need to Stop: " + $Service)
							Stop-PSService $IP $Service $PSServicePath $PSKillPath $maximumRuntimeSeconds
						}
						#Term Program
						If ($Program) {
							Write-Progress -Activity ("Test if we need to kill: " + $Program) -Status ("( " + $count + "\" + $Computers.count + ") Computer: " + $Computer + " IP: " + $IP + " Runtime: " + (FormatElapsedTime($sw.Elapsed))) -percentComplete ($count / $Computers.count*100)
							Write-Host ("`t Test if we need to kill: " + $Program)
							Kill-PSProgram $IP $Program $PSKillPath $maximumRuntimeSeconds
						}
						If ($Copy) {
							Write-Host ("`t`t Copying missing " + $SourceFileInfo.value.name + " to destination: " + $("\\" +  $IP + "\" + $Destination.replace(":","$"))) -foregroundcolor green
							Copy-Item $SourceFileInfo.value -Destination "PSRemote:\"
						}
						$MissingFiles = $true
						If ($ErrorCSV) {
							#"Date,Computer,Source File, Source File Version, Source file Date,Destination File,Destination File Version,Destination File Date,Error"
							# Add-Content ($LogFile + "_errors.csv") ((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + "," + $SourceFileInfo.value.name + "," + $SourceFileInfo.value.VersionInfo.ProductVersion + "," + $SourceFileInfo.value.LastWriteTime + ",,,,Destination File is missing; copying file")
							((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + "," + $SourceFileInfo.value.name + "," + $SourceFileInfo.value.VersionInfo.ProductVersion + "," + $SourceFileInfo.value.LastWriteTime + ",,,,Destination File is missing; copying file") | Out-File -Append -FilePath ($LogFile + "_errors.csv")
						}
						If ($AllCSV) {
							#"Date,Computer,Source File, Source File Version, Source file Date,Destination File,Destination File Version,Destination File Date,Status"
							# Add-Content ($LogFile + "_all.csv") ((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + "," + $SourceFileInfo.value.name + "," + $SourceFileInfo.value.VersionInfo.ProductVersion + "," + $SourceFileInfo.value.LastWriteTime + ",,,,Destination File is missing; copying file")
							((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + "," + $SourceFileInfo.value.name + "," + $SourceFileInfo.value.VersionInfo.ProductVersion + "," + $SourceFileInfo.value.LastWriteTime + ",,,,Destination File is missing; copying file") | Out-File -Append -FilePath ($LogFile + "_all.csv")
						}
						#Start Service
						Start-PSService $IP $Service $PSServicePath $maximumRuntimeSeconds
						#Run Program
						If ($Command) {
							#Write-Host ("`t`t Running $Command on  $Computer.") -foregroundcolor green
							if ( $User -and $Password) {
								Start-PSProgram $Computer $Command $PSExecPath $maximumRuntimeSeconds $User $Password
							}else{
								Start-PSProgram $Computer $Command $PSExecPath $maximumRuntimeSeconds
							}	
						}
					}
				}
			}else {
			Write-Warning ("Unable to access admin are of: " + "\\" +  $IP + "\" + $Destination.replace(":","$"))
			If ($ErrorCSV) {
				#"Date,Computer,Source File, Source File Version, Source file Date,Destination File,Destination File Version,Destination File Date,Error"
				# Add-Content ($LogFile + "_errors.csv") ((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + ",,,,,,,"+ "Unable to access admin are of: " + "\\" +  $IP + "\" + $Destination.replace(":","$"))
				 ((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + ",,,,,,,"+ "Unable to access admin are of: " + "\\" +  $IP + "\" + $Destination.replace(":","$")) | Out-File -Append -FilePath ($LogFile + "_errors.csv")
			}
			If ($AllCSV) {
				#"Date,Computer,Source File, Source File Version, Source file Date,Destination File,Destination File Version,Destination File Date,Status"
				# Add-Content ($LogFile + "_all.csv") ((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + "," + $SourceFileInfo.value.name + "," + $SourceFileInfo.value.VersionInfo.ProductVersion + "," + $SourceFileInfo.value.LastWriteTime + ",,,," + "Unable to access admin are of: " + "\\" +  $IP + "\" + $Destination.replace(":","$"))
				((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + "," + $SourceFileInfo.value.name + "," + $SourceFileInfo.value.VersionInfo.ProductVersion + "," + $SourceFileInfo.value.LastWriteTime + ",,,," + "Unable to access admin are of: " + "\\" +  $IP + "\" + $Destination.replace(":","$")) | Out-File -Append -FilePath ($LogFile + "_all.csv")
			}
			
			}
		}
	}Else{
		#Error missing folder or computer.
		Write-Warning -Message ("Error: $Computer has NO working IP Addresses")
		If (Test-Connection -ComputerName $Computer -Quiet){ 
			Write-Host ("`t`t Host is up") -ForegroundColor green
		}else{
			Write-Warning -Message ("`t`t Host is Down")
		}
		If ($ErrorCSV) {
			#"Date,Computer,Source File, Source File Version, Source file Date,Destination File,Destination File Version,Destination File Date,Error"
			# Add-Content ($LogFile + "_errors.csv") ((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + "," + $SourceFileInfo.value.name + "," + $SourceFileInfo.value.VersionInfo.ProductVersion + "," + $SourceFileInfo.value.LastWriteTime + ",,,,Cannot access host")
			((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + "," + $SourceFileInfo.value.name + "," + $SourceFileInfo.value.VersionInfo.ProductVersion + "," + $SourceFileInfo.value.LastWriteTime + ",,,,Cannot access host") | Out-File -Append -FilePath ($LogFile + "_errors.csv")
		}
		If ($AllCSV) {
			#"Date,Computer,Source File, Source File Version, Source file Date,Destination File,Destination File Version,Destination File Date,Status"
			# Add-Content ($LogFile + "_all.csv") ((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + "," + $SourceFileInfo.value.name + "," + $SourceFileInfo.value.VersionInfo.ProductVersion + "," + $SourceFileInfo.value.LastWriteTime + ",,,,Cannot access host")
			Add-Content ($LogFile + "_all.csv") ((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + "," + $SourceFileInfo.value.name + "," + $SourceFileInfo.value.VersionInfo.ProductVersion + "," + $SourceFileInfo.value.LastWriteTime + ",,,,Cannot access host") | Out-File -Append -FilePath ($LogFile + "_all.csv")
		}
		$ComputerError = $true
	}
	#Extra logging
	If ($VerboseLog) {
		# If ($NoChange) {Add-Content ($LogFile + "_NoChanges.txt") ("$Computer")}
		If ($NoChange) {("$Computer") | Out-File -Append -FilePath ($LogFile + "_NoChanges.txt")}
		# If ($UpdatesNeeded) {Add-Content ($LogFile + "_UpdatesNeeded.txt") ("$Computer")}
		If ($UpdatesNeeded) { ("$Computer") | Out-File -Append -FilePath ($LogFile + "_UpdatesNeeded.txt")}
		# If ($MissingFiles) {Add-Content ($LogFile + "_MissingFiles.txt") ("$Computer")}
		If ($MissingFiles) {("$Computer") | Out-File -Append -FilePath ($LogFile + "_MissingFiles.txt")}
	}
	# IF ($ComputerError) {Add-Content ($LogFile + "_ErrorComputers.txt") ("$OldComputer")}
	IF ($ComputerError) { ("$OldComputer")  | Out-File -Append -FilePath ($LogFile + "_ErrorComputers.txt")}
	#Increase Progress counter
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
