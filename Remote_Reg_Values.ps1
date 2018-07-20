<# Remote Reg value
Operations:
	*Read input
	*Test if host is alive
	* Powershell or PSExec to get registry Value Data

Dependencies for this script:
	* PsExec
Changes:
	*Convert hex to decimal for REG_DWORD Version 1.0.1
    *Fixed -computer issue Version 1.0.3
    *Fixed issue where psexec has and issue Version 1.0.4
#>

PARAM (
    [string]$Computer = $null, 
    [string]$ComputerList = $null,
    [string]$Key = "SOFTWARE\\Microsoft\\NET Framework Setup\\NDP\\v4\\Full",
    [string]$Hive = "HKLM",
    [string]$Value = "Release",
    [switch]$ErrorCSV = $false,
    [switch]$ALLCSV = $true,
    [switch]$UsePSExec = $true,
    [string]$PSExecPath = ".\PsExec.exe"
)

$ScriptVersion = "1.0.4"

#############################################################################
#region User Variables
#############################################################################
$LogFile = ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\" + `
		   $MyInvocation.MyCommand.Name + "_" + `
		   (Get-Date -format yyyyMMdd-hhmm) + ".log")
$sw = [Diagnostics.Stopwatch]::StartNew()
$count = 1
$maximumRuntimeSeconds = 30
$GoodIPs=@()
#############################################################################
#endregion User Variables
#############################################################################

#############################################################################
#region Setup Sessions
#############################################################################
#Start logging.
If (-Not [string]::IsNullOrEmpty($LogFile)) {
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
	If ([string]::IsNullOrEmpty($Computer)) {
		throw " -Computer or -ComputerList is required."
	}Else{
		# $Computers is already set.
		If (($Computer.GetType().BaseType).Name -eq "String") {
			$Computers = $Computer.split(" ")
		}
	}
}Else{
	[Array]$Computers += Get-Content -Path $ComputerList
}

#Clean up \ in key varible
$tempstr = $null
foreach ($tempkey in $Key.Split("\",[System.StringSplitOptions]::RemoveEmptyEntries)) {
    $tempstr = ($tempstr + "\\" + $tempkey)
}
$Key = $tempstr

$tempstr = $null
foreach ($tempkey in $Key.Split("\",[System.StringSplitOptions]::RemoveEmptyEntries)) {
    $tempstr = ($tempstr + "\" + $tempkey)
}
$PSKey = $tempstr

#CSV Setup
If ($ErrorCSV) {
	If (!(Test-Path -Path ($LogFile + "_errors.csv"))) {
		Add-Content ($LogFile + "_errors.csv") ("Date,Computer,Key,Value,Value Data,Error")
	}
}
If ($ALLCSV) {
	If (!(Test-Path -Path ($LogFile + "_all.csv"))) {
		Add-Content ($LogFile + "_all.csv") ("Date,Computer,Key,Value,Value Data,Status")
	}
}

#Check for PSExec
If (-Not ($UsePSExec)) {
	If (-Not (Test-Path $PSExecPath -PathType Leaf)) {
		throw ("psexec.exe is not found at: " + $PSExecPath)
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
#############################################################################
#endregion Functions
#############################################################################

#############################################################################
#region Main
#############################################################################
#Loop thru servers
Foreach ($Computer in $Computers) {
$ValueData = $null
Write-Progress -Activity ("Resolving Computer Name") -Status ("( " + $count + "\" + $Computers.count + ") Computer: " + $Computer) -percentComplete ($count / $Computers.count*100)
	Write-Host ("( " + $count + "\" + $Computers.count + ") Computer: " + $Computer)
	#test for good IPs from host
	$GoodIPs=@()
	If ([ipaddress]$Computer) {
		If (Test-Connection -Cn $Computer -BufferSize 16 -Count 1 -ea 0 -quiet) {
			Write-Host ("`t Responds with IP: " + $Computer)
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
             #Main Code.
            if ($UsePSExec) {
                Write-Host ("`t`t Trying Remote Registry on: " + $IP)
                $ValueData = $null
                $arr = $null
                $stdout = $null
                $stderr = $null
                $pinfo = $null
                $process = $null
                try {		        
                    #$process = Start-Process -FilePath $PSExecPath -ArgumentList @("\\" + $IP, 'reg query "' + $Hive + $PSKey + '" /v ' + $Value) -PassThru -NoNewWindow
		            $pinfo = New-Object System.Diagnostics.ProcessStartInfo
                    $pinfo.FileName = $PSExecPath
                    $pinfo.RedirectStandardError = $true
                    $pinfo.RedirectStandardOutput = $true
                    $pinfo.UseShellExecute = $false
                    $pinfo.Arguments = ("\\" + $IP +' reg query "' + $Hive + $PSKey + '" /v ' + $Value)
                    $process = New-Object System.Diagnostics.Process
                    $process.StartInfo = $pinfo
                    $process.Start() | Out-Null
                    $process.WaitForExit()
                    $stdout = $process.StandardOutput.ReadToEnd()
                    $stderr = $process.StandardError.ReadToEnd()
                    $arr = $stdout.Split([Environment]::NewLine,[System.StringSplitOptions]::RemoveEmptyEntries)
                    $arr =  ($arr[$arr.Count -1]).Split(" ",[System.StringSplitOptions]::RemoveEmptyEntries)
                    #Convert hex to dec
                    If ( ($arr[$arr.Count -2 ]) -eq "REG_DWORD") {
                        $ValueData = [Convert]::ToInt64($arr[$arr.Count -1],16)
                    }else{
                        $ValueData = $arr[$arr.Count -1]
                        If ($ValueData -eq "www.sysinternals.com") {
                            $arr = $stderr.Split([Environment]::NewLine,[System.StringSplitOptions]::RemoveEmptyEntries)
                            $ValueData = $arr.Count[$arr.Count -1]
                        }
                    }
                    #Write-Host $stdout
                    #Write-Host $stderr
                    Write-Host ("`t`t Value Data: " + $ValueData + " Key: " + $Key + " Value: " + $Value)
                    If ($ALLCSV) {
					   #"Date,Computer,Key,Value,Value Data,Status"
					   Add-Content ($LogFile + "_all.csv") ((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + "," + $key + "," + $Value + "," + $ValueData + ",Success")
					}
                } catch {
					Write-Host ("`t`t Cannot connect to remote registry.") -ForegroundColor red
					If ($ErrorCSV) {
						#Date,Computer,Key,Value,Value Data,Error
						Add-Content ($LogFile + "_errors.csv") ((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + "," + $key + "," + $Value + "," + $ValueData + ", Cannot read remove registry")
					}
                }
            }else{  
                $ValueData = $null
                try {
                    $Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($Hive, $IP)
                    $RegKey= $Reg.OpenSubKey($Key)
                    $ValueData = $RegKey.GetValue($Value)
                    Write-Host ("`t`t Value Data: " + $ValueData + " Key: " + $Key + " Value: " + $Value)
                    If ($ALLCSV) {
                       #"Date,Computer,Key,Value,Value Data,Status"
		               Add-Content ($LogFile + "_all.csv") ((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + "," + $key + "," + $Value + "," + $ValueData + ",Success")
                    }
                } catch {
                    Write-Host ("`t`t Cannot connect to remote registry.") -ForegroundColor red
                    If ($ErrorCSV) {
                        #Date,Computer,Key,Value,Value Data,Error
		                Add-Content ($LogFile + "_errors.csv") ((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + "," + $key + "," + $Value + "," + $ValueData + ", Cannot read remove registry")
                    }
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
			#Date,Computer,Key,Value,Value Data,Error
			Add-Content ($LogFile + "_errors.csv") ((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + "," + $key + "," + $Value + "," + $ValueData + ",NO working IP Addresses")
		}
        If ($ALLCSV) {
            #"Date,Computer,Key,Value,Value Data,Status"
		    Add-Content ($LogFile + "_all.csv") ((Get-Date -format yyyyMMdd-hhmm) + "," + $Computer + "," + $key + "," + $Value + "," + $ValueData + ",NO working IP Addresses")
        }
	}

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
