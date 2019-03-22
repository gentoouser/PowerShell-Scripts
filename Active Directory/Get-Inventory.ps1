<#
.SYNOPSIS
  Name: Get-Inventory.ps1
  The purpose of this script is to create a simple inventory.
  
.DESCRIPTION
  This is a simple script to retrieve all computer objects in Active Directory and then connect
  to each one and gather basic hardware information using Cim. The information includes Manufacturer,
  Model,Serial Number, CPU, RAM, Disks, Operating System, Sound Deivces and Graphics Card Controller.

.RELATED LINKS
  https://www.sconstantinou.com
  https://gallery.technet.microsoft.com/scriptcenter/PowerShell-Hardware-f99336f6

.NOTES
  Version 1.2

  Updated:      01-06-2018        - Replaced Get-WmiObject cmdlet with Get-CimInstance
                                  - Added Serial Number Information
                                  - Added Sound Device Information
                                  - Added Video Controller Information
                                  - Added option to send CSV file through email
                                  - Added parameters to enable email function option
                03/21/2019        - Added IP address
                                  - Added Software
                                  - Added Installed Roles
                                  - Added Installed Features
  Release Date: 10-02-2018
   
  Author: Stephanos Constantinou

.EXAMPLES
  Get-Inventory.ps1
  Find the output under C:\Scripts_Output

  Get-Inventory.ps1 -Email -Recipients user1@domain.com
  Find the output under C:\Scripts_Output and an email will be sent
  also to user1@domain.com

  Get-Inventory.ps1 -Email -Recipients user1@domain.com,user2@domain.com
  Find the output under C:\Scripts+Output and an email will be sent
  also to user1@domain.com and user2@domain.com
#>

Param(
    [switch]$Email = $false,
    [string]$Recipients = $null,
    [String]$OutputFolder = ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\Logs")
)

$LogFile = ($OutputFolder + "\" + `
		   $MyInvocation.MyCommand.Name + "_" + `
		   $env:computername + "_" + `
       (Get-Date -format yyyyMMdd-hhmm) + ".log")
$ScriptVersion = "1.2.0"
$sw = [Diagnostics.Stopwatch]::StartNew()
function FormatElapsedTime($ts) {
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


If (-Not [string]::IsNullOrEmpty($LogFile)) {
  If (-Not( Test-Path (Split-Path -Path $LogFile -Parent))) {
    New-Item -ItemType directory -Path (Split-Path -Path $LogFile -Parent) | Out-Null
    $Acl = Get-Acl (Split-Path -Path $LogFile -Parent)
    $Ar = New-Object system.Security.AccessControl.FileSystemAccessRule('Users', "FullControl", "ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.Setaccessrule($Ar)
    Set-Acl (Split-Path -Path $LogFile -Parent) $Acl
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
$Inventory = New-Object System.Collections.ArrayList

if ($Email -eq $true){

    $EmailCredentials = $host.ui.PromptForCredential("Need email credentials", "Provide the user that will be used to send the email.","","")
    $To  = @(($Recipients) -split ',')
    $Attachement = ($OutputFolder + "\Inventory_" + (Get-Date -format yyyyMMdd-hhmm) + ".csv")
    $From = $EmailCredentials.UserName

    $EmailParameters = @{
        To = $To
        Subject = "Inventory"
        Body = "Please find attached the inventory that you have requested."
        Attachments = $Attachement
        UseSsl = $True
        Port = "587"
        SmtpServer = "smtp.office365.com"
        Credential = $EmailCredentials
        From = $From}
}

$AllComputers = Get-ADComputer -Filter * -Properties Name
$AllComputersNames = $AllComputers.Name

Foreach ($ComputerName in $AllComputersNames) {
  Write-Host ("Recording Computer Info: " + $ComputerName)
  $Connection = Test-Connection $ComputerName -Count 1 -Quiet

  $ComputerInfo = New-Object System.Object

  $ComputerOS = Get-ADComputer $ComputerName -Properties OperatingSystem,OperatingSystemServicePack,OperatingSystemVersion,LastLogonDate

  $ComputerInfo | Add-Member  -Force -MemberType NoteProperty -Name "Name" -Value "$ComputerName" 
  $ComputerInfo | Add-Member  -Force -MemberType NoteProperty -Name "OperatingSystem" -Value $ComputerOS.OperatingSystem
  $ComputerInfo | Add-Member  -Force -MemberType NoteProperty -Name "OperatingSystemVersion" -Value $ComputerOS.OperatingSystemVersion
  $ComputerInfo | Add-Member  -Force -MemberType NoteProperty -Name "ServicePack" -Value $ComputerOS.OperatingSystemServicePack
  $ComputerInfo | Add-Member  -Force -MemberType NoteProperty -Name "LastLogonDate" -Value $ComputerOS.LastLogonDate

  If ($Connection -eq "True") {
    Write-Host ("`tHost is up")
    $ArrComputerIP= (Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $ComputerName -Filter 'IPEnabled = True' -ErrorAction SilentlyContinue | Where-object {$_.IPaddress -notlike "169.254.*" -and $_.IPAddress -ne "127.0.0.1"} | Select-Object IPAddress,Description)

    If ($ArrComputerIP) {    
      If (($ArrComputerIP | Measure-Object).count -gt 1) {
        $ComputerIP = ('"' + (($ArrComputerIP | ForEach-Object { ($_.IPAddress + " - " + $_.Description) -join " "}) -Join ",") + '"')
      } else {
        $ComputerIP = ('"' + ($ArrComputerIP | ForEach-Object { ($_.IPAddress + " - " + $_.Description) -join " "})  + '"')
      }
      Write-Host ("`t`tHost WMI reachable with IPs: " + $ComputerIP)
      $ComputerHW = Get-CimInstance -Class Win32_ComputerSystem -ComputerName $ComputerName |
          Select-Object Manufacturer,Model,NumberOfProcessors,@{Expression={($_.TotalPhysicalMemory / 1GB).ToString("#.##")};Label="TotalPhysicalMemoryGB"}

      $ComputerCPU = Get-CimInstance win32_processor -ComputerName $ComputerName |
          Select-Object DeviceID,Name,Manufacturer,NumberOfCores,NumberOfLogicalProcessors

      $ComputerDisks = Get-CimInstance -Class Win32_LogicalDisk -Filter "DriveType=3" -ComputerName $ComputerName |
          Select-Object DeviceID,VolumeName,@{Expression={($_.Size / 1GB).ToString("#.##")};Label="SizeGB"}

      $ComputerSerial = (Get-CimInstance Win32_Bios -ComputerName $ComputerName).SerialNumber

      $ComputerGraphics = Get-CimInstance -Class Win32_VideoController |Select-Object Name,@{Expression={($_.AdapterRAM / 1GB).ToString("#.##")};Label="GraphicsRAM"}

      $ComputerSoundDevices = ('"' + ((Get-CimInstance -Class Win32_SoundDevice).Name  -join ", ") + '"')
            
      $ComputerInfoManufacturer = $ComputerHW.Manufacturer
      $ComputerInfoModel = $ComputerHW.Model
      $ComputerInfoNumberOfProcessors = $ComputerHW.NumberOfProcessors
      $ComputerInfoProcessorID = $ComputerCPU.DeviceID | Select-Object -Unique
      $ComputerInfoProcessorManufacturer = $ComputerCPU.Manufacturer | Select-Object -Unique
      $ComputerInfoProcessorName = $ComputerCPU.Name | Select-Object -Unique
      $sum = 0
      $ComputerCPU.NumberOfCores  | ForEach-Object { $sum += $_}
      $ComputerInfoNumberOfCores = $sum
      $sum = 0
      $ComputerCPU.NumberOfLogicalProcessors | ForEach-Object { $sum += $_}
      $ComputerInfoNumberOfLogicalProcessors = $sum
      $ComputerInfoRAM = $ComputerHW.TotalPhysicalMemoryGB
      $ComputerInfoDiskDrive = ('"' + ($ComputerDisks.DeviceID  -join ", ") + '"')
      $ComputerInfoDriveName = ('"' + ($ComputerDisks.VolumeName  -join ", ") + '"')
      $ComputerInfoSize = ('"' + ($ComputerDisks.SizeGB   -join ", ") + '"')
      $ComputerInfoGraphicsName = $ComputerGraphics.Name
      $ComputerInfoGraphicsRAM = $ComputerGraphics.GraphicsRAM

      $ComputerFeatures = ('"' + ((Get-WmiObject -ComputerName $ComputerName -query "select Name from win32_optionalfeature where installstate= 1").name  -join ", ") + '"')
      If (($ComputerOS.OperatingSystem).Contains("Server")) {
        $ComputerRolesTemp = Invoke-Command -ComputerName $ComputerName -Verbose -ScriptBlock { Import-Module servermanager;get-windowsfeature | Where-Object { $_.installed -eq $true -and $_.featuretype -eq 'Role'} | Select-Object name}
        $ComputerRoles = ('"' + (($ComputerRolesTemp.Name) -join ", ") + '"')
      }
      
      $ArrComputerSoftware = Get-WmiObject -Class Win32_Product -ComputerName $ComputerName
      If (($ArrComputerSoftware | Measure-Object).count -gt 1) {
        $ComputerSoftware = ('"' + (($ArrComputerSoftware | ForEach-Object {($_.Name + " - " + $_.Version)})  -join ", ") + '"')
      } else {
        $ComputerSoftware = ('"' + ($ArrComputerSoftware.Name + " - " + $ArrComputerSoftware.Version)  + '"')
      }
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "ComputerIPs" -Value $ComputerIP -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "Manufacturer" -Value "$ComputerInfoManufacturer" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "Model" -Value "$ComputerInfoModel" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "Serial" -Value "$ComputerSerial" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "NumberOfProcessors" -Value "$ComputerInfoNumberOfProcessors" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "ProcessorID" -Value "$ComputerInfoProcessorID" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "ProcessorManufacturer" -Value "$ComputerInfoProcessorManufacturer" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "ProcessorName" -Value "$ComputerInfoProcessorName" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "NumberOfCores" -Value "$ComputerInfoNumberOfCores" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "NumberOfLogicalProcessors" -Value "$ComputerInfoNumberOfLogicalProcessors" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "RAM" -Value "$ComputerInfoRAM" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "DiskDrive" -Value "$ComputerInfoDiskDrive" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "DriveName" -Value "$ComputerInfoDriveName" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "Size" -Value "$ComputerInfoSize" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "Graphics" -Value "$ComputerInfoGraphicsName" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "GraphicsRAM" -Value "$ComputerInfoGraphicsRAM" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "SoundDevices" -Value "$ComputerSoundDevices" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "WindowsFeatures" -Value "$ComputerFeatures" -Force
      If (($ComputerOS.OperatingSystem).Contains("Server")) {
        $ComputerInfo | Add-Member -MemberType NoteProperty -Name "WindowsRoles" -Value "$ComputerRoles" -Force
      } else {
        $ComputerInfo | Add-Member -MemberType NoteProperty -Name "WindowsRoles" -Value "" -Force
      }
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "Software" -Value "$ComputerSoftware" -Force
    } else {
      Write-Host ("`t`tHost WMI Not Reachable for computer: " + $ComputerName) -ForegroundColor Red
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "ComputerIPs" -Value ""
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "Manufacturer" -Value "" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "Model" -Value "" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "Serial" -Value "" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "NumberOfProcessors" -Value "" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "ProcessorID" -Value "" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "ProcessorManufacturer" -Value "" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "ProcessorName" -Value "" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "NumberOfCores" -Value "" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "NumberOfLogicalProcessors" -Value "" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "RAM" -Value "" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "DiskDrive" -Value "" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "DriveName" -Value "" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "Size" -Value ""-Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "Graphics" -Value "" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "GraphicsRAM" -Value "" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "SoundDevices" -Value "" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "WindowsFeatures" -Value "" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "WindowsRoles" -Value ""-Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "Software" -Value "" -Force
    }
  }

  If ($ComputerInfo) {
      $Inventory.Add($ComputerInfo) | Out-Null
  }
   $ComputerHW = ""
   $ComputerCPU = ""
   $ComputerDisks = ""
   $ComputerSerial = ""
   $ComputerGraphics = ""
   $ComputerSoundDevices = ""
}

$Inventory | Export-Csv -NoTypeInformation -Path ($OutputFolder + "\Inventory_" + (Get-Date -format yyyyMMdd-hhmm) + ".csv")

if ($Email -eq $true){send-mailmessage @EmailParameters}

$sw.Stop()
Write-Host ("Script took: " + (FormatElapsedTime($sw.Elapsed)) + " to run.")
#============================================================================
#endregion Main Local Machine Cleanup
#============================================================================
If (-Not [string]::IsNullOrEmpty($LogFile)) {
	Stop-Transcript
}