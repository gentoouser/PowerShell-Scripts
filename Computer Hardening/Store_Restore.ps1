<# 
.SYNOPSIS
    Name: Store_Restore.ps1
    Import Store Settings to Zip file

.DESCRIPTION
    Restoreuser data and custom app settings.


.PARAMETER 


.EXAMPLE
   & Store_Restore.ps1

.NOTES
 Author: Paul Fuller
 Changes:
    1.0.00 - Basic script functioning and can work on Windows 7
    1.0.01 - Updated Manager registry settings
    1.0.02 - Added code to hide Console
    1.0.03 - Auto-Logon Fix, Fixes or unlocking Managers. 
    1.0.05 - Removed WindowLogonUser and Added logic to update logon based on machine name.
    1.0.07 - Fix IP update and computer rename.
    1.0.09 - Fix bug with disabling all accounts. Added prompt about disableing Auto Logon. Fixed issues with setting Manager settings. 
    1.0.10 - Updated to deal with zip file from powershell 2.0
    1.0.11 - Manager's registy keys bug. Get-CimInstance testing. Start Office 2019 install. 
    1.0.12 - Fixed issue with renaming machine.
#>
#Requires -Version 5.1 -PSEdition Desktop
#Force Starting of Powershell script as Administrator 
If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {   
    $arguments = "& '" + $myinvocation.mycommand.definition + "'"
    Start-Process powershell -Verb runAs -ArgumentList $arguments
    Break
 }
# .Net methods for hiding/showing the console in the background
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
#############################################################################
#region User Variables
#############################################################################
$Settings =[hashtable]::Synchronized(@{})
# $Settings =@{}
$SettingsOutput =[hashtable]::Synchronized(@{})
# $SettingsOutput =@{}
$Settings.Version = "1.0.12"
$Settings.WindowTitle = ("Store Restore Version: " + $Settings.Version)
$Settings.tempfolder = ""
$Settings.CustomAppFolder = "github\app"
$Settings.CustomAppRegKey = "github\app"
$Settings.CustomAppName = "app"
$Settings.DNS = @("1.1.1.1","8.8.8.8")
$Settings.Subnet = "255.255.255.0"
$Settings.OfficeSubFolder = ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\Microsoft Office 2019")
$Settings.OfficeActivationScript = ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "Office_2019_Activate.bat")
$settings.Admin = "admin"
$Settings.AccountBlacklist = @(
    "Administrator"
    "ASPNET"
    "DefaultAccount"
    "Guest"
    "WDAGUtilityAccount"  
)
$Settings.AccountDisableBlacklist = @(
        "ASPNET"
)
$Settings.DEVCNames = @(
    "DEV"
    "TST"
    "QA"
)
$Settings.WindowLogonUserRegString = "hkcu:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
$Settings.WindowLogonUserReg = (Get-ItemProperty -path $Settings.WindowLogonUserRegString)
$Settings.USF = (Get-ItemProperty -path "hkcu:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders")
$Settings.UsersProfileFolder = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\' -Name "ProfilesDirectory").ProfilesDirectory

$Settings.BackupFolders =@()
If (Test-Path (${env:ProgramFiles(x86)} + "\" + $Settings.CustomAppFolder)) {
    $Settings.BackupFolders += (${env:ProgramFiles(x86)} + "\" + $Settings.CustomAppFolder)
}
If (Test-Path (${env:ProgramFiles} + "\" + $Settings.CustomAppFolder)) {
    $Settings.BackupFolders += (${env:ProgramFiles} + "\" + $Settings.CustomAppFolder)
} 
$Settings.BackupFolders += $([string]$Settings.USF.Desktop)
$Settings.BackupFolders += $([string]$Settings.USF.Favorites)
$Settings.BackupFolders += $([string]$Settings.USF."My Pictures")
$Settings.BackupFolders += $([string]$Settings.USF."{374DE290-123F-4565-9164-39C4925E467B}") #Downloads
$Settings.BackupFolders += $([string]$Settings.USF.Personal) #My Documents

#region Icon
$iconBase64 =''
#endregion Icon
#############################################################################
#endregion User Variables
#############################################################################
#############################################################################
#region Functions
#############################################################################
function FormatElapsedTime($ts) {   	   	 
    #https://stackoverflow.com/questions/3513650/timing-a-commands-execution-in-powershell
	$elapsedTime = ""
    if ( $ts.Hours -gt 0 ) {	 
        $elapsedTime = [string]::Format( "{0:00} hours {1:00} min. {2:00}.{3:00} sec.", $ts.Hours, $ts.Minutes, $ts.Seconds, $ts.Milliseconds / 10 );
    } else {
        if ( $ts.Minutes -gt 0 ) {		 
            $elapsedTime = [string]::Format( "{0:00} min. {1:00}.{2:00} sec.", $ts.Minutes, $ts.Seconds, $ts.Milliseconds / 10 );
        } else {		 
            $elapsedTime = [string]::Format( "{0:00}.{1:00} sec.", $ts.Seconds, $ts.Milliseconds / 10 );
        }
        if ($ts.Hours -eq 0 -and $ts.Minutes -eq 0 -and $ts.Seconds -eq 0) {
            $elapsedTime = [string]::Format("{0:00} ms.", $ts.Milliseconds);
        }
        if ($ts.Milliseconds -eq 0) {		 
            $elapsedTime = [string]::Format("{0} ms", $ts.TotalMilliseconds);
        }
    }
    return $elapsedTime
}
function Set-Reg {
    [CmdletBinding()] 
    Param 
    ( 
        [Parameter(Mandatory=$true,Position=1,HelpMessage="Path to Registry Key")][string]$regPath, 
        [Parameter(Mandatory=$true,Position=2,HelpMessage="Name of Value")][string]$name,
        [Parameter(Mandatory=$true,Position=3,HelpMessage="Data for Value")]$value,
        [Parameter(Mandatory=$true,Position=4,HelpMessage="Type of Value")][ValidateSet("String", "ExpandString","Binary","DWord","MultiString","Qword","Unknown",IgnoreCase =$true)][string]$type 
    ) 
    #Source: https://github.com/nichite/chill-out-windows-10/blob/master/chill-out-windows-10.ps1
    # String: Specifies a null-terminated string. Equivalent to REG_SZ.
    # ExpandString: Specifies a null-terminated string that contains unexpanded references to environment variables that are expanded when the value is retrieved. Equivalent to REG_EXPAND_SZ.
    # Binary: Specifies binary data in any form. Equivalent to REG_BINARY.
    # DWord: Specifies a 32-bit binary number. Equivalent to REG_DWORD.
    # MultiString: Specifies an array of null-terminated strings terminated by two null characters. Equivalent to REG_MULTI_SZ.
    # Qword: Specifies a 64-bit binary number. Equivalent to REG_QWORD.
    If(!(Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }
    If($type -eq "Binary" -and $value.GetType().Name -eq "String" -and $value -match ",") {
        $value = [byte[]]($value -split ",")
    }
    New-ItemProperty -Path $regPath -Name $name -Value $value -PropertyType $type -Force | Out-Null
}
function Show-Console
{
    $consolePtr = [Console.Window]::GetConsoleWindow()
    # Hide = 0,
    # ShowNormal = 1,
    # ShowMinimized = 2,
    # ShowMaximized = 3,
    # Maximize = 3,
    # ShowNormalNoActivate = 4,
    # Show = 5,
    # Minimize = 6,
    # ShowMinNoActivate = 7,
    # ShowNoActivate = 8,
    # Restore = 9,
    # ShowDefault = 10,
    # ForceMinimized = 11
    [Console.Window]::ShowWindow($consolePtr, 4)
}
function Hide-Console
{
    $consolePtr = [Console.Window]::GetConsoleWindow()
    #0 hide
    [Console.Window]::ShowWindow($consolePtr, 0)
}
function Browse_File {
    param (
      
    )
	$Settings.Store_Setup.text = (" Opening Archive . . . Please wait.")
    $Settings.Stop.Text = "Stop"
    $Settings.Browse.Enabled = $false
    $Settings.Restore_Backup.Enabled = $false
    $Settings.IP_Address.Enabled = $false
    $Settings.Machine_Name.Enabled = $false
    $Settings.Start.Enabled = $false
    $Settings.CABackup.Enabled = $false
    $Settings.UserFilesBackup.Enabled = $false
    $Settings.FBackup.Enabled = $false
    $Settings.Network_Adapter.Enabled = $false
    $Settings.WindowLogonUser.Enabled = $false
    $Settings.Manager.Enabled = $false
    $Settings.OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    #$Settings.OpenFileDialog.initialDirectory = (Split-Path -Path $MyInvocation.MyCommand.Definition) 
    $Settings.OpenFileDialog.filter = "ZIP Archive Files|*.zip|All Files|*.*" 
    $Settings.OpenFileDialog.ShowDialog() | Out-Null
    $Settings.Restore_Backup.Text = $Settings.OpenFileDialog.filename    
    $Settings.tempfolder = ($env:temp + "\" + [io.path]::GetFileNameWithoutExtension($Settings.OpenFileDialog.filename))

      $BrowseRunspace =[runspacefactory]::CreateRunspace()
      $BrowseRunspace.ApartmentState = "STA"
      $BrowseRunspace.ThreadOptions = "ReuseThread"     
      $BrowseRunspace.Open()
      $BrowseRunspace.name = "Browse"
	  $BrowseRunspace.SessionStateProxy.SetVariable("Settings",$Settings)  
      $BrowseRunspace.SessionStateProxy.SetVariable("SettingsOutput",$SettingsOutput)    

      $BrowsepsCmd = "" | Select-Object PowerShell,Handle
      $BrowsepsCmd.PowerShell = [PowerShell]::Create().AddScript({ 
        Expand-Archive -Path $Settings.Restore_Backup.Text -DestinationPath $env:temp  -Force
     })
     $BrowsepsCmd.Powershell.Runspace = $BrowseRunspace
     $BrowsepsCmd.Handle = $BrowsepsCmd.Powershell.BeginInvoke()
         #Wait for code to complete and keep GUI responsive
    do {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 1
    } while ($BrowsepsCmd.Handle.IsCompleted -eq $false)

     If (Test-Path ($Settings.tempfolder + "\settings.xml")) {
        $SettingsOutput = Import-Clixml -Path ($Settings.tempfolder + "\settings.xml")
    }elseIf (Test-Path ($Settings.tempfolder + (Split-Path -Path $Settings.tempfolder -Leaf) + "\settings.xml")) {
        $SettingsOutput = Import-Clixml -Path ($Settings.tempfolder + (Split-Path -Path $Settings.tempfolder -Leaf) + "\settings.xml")
    }

    $Settings.Machine_Name.text               = $SettingsOutput.MachineName
    $Settings.IP_Address.Text                 = $SettingsOutput.IPAddress

    $Settings.Store_Setup.text = ( $Settings.WindowTitle)
    $Settings.Browse.Enabled = $true
    $Settings.Restore_Backup.Enabled = $true
    $Settings.IP_Address.Enabled = $true
    $Settings.Machine_Name.Enabled = $true
    $Settings.Start.Enabled = $true
	$Settings.CABackup.Enabled  = $true
    $Settings.UserFilesBackup.Enabled  = $True
    $Settings.FBackup.Enabled = $True
    $Settings.Network_Adapter.Enabled = $True
    $Settings.WindowLogonUser.Enabled = $True
    $Settings.Manager.Enabled = $True
    $Settings.Start.text = "Restore"
}
function Start_Work {
    param (
        
    )
    If (($Settings.Machine_Name.Text -split "-")[0] -eq "HP") {
       [System.Windows.MessageBox]::Show(('Invalid Machine Name: ' + $Settings.Machine_Name.Text + " Please fix. . ." ),('Invalid Machine Name: ' + $Settings.Machine_Name.Text + " Please fix. . ." ),'OK','Hand') 
    }else{
        $Settings.sw = [Diagnostics.Stopwatch]::StartNew()
        $Settings.Store_Setup.text  = ( $Settings.WindowTitle + " Working . . . Please wait.")
        $Settings.Stop.Text = "Stop"
        $Settings.Browse.Enabled = $false
        $Settings.Restore_Backup.Enabled = $false
        $Settings.IP_Address.Enabled = $false
        $Settings.Machine_Name.Enabled = $false
        $Settings.Start.Enabled = $false
        $Settings.CABackup.Enabled = $false
        $Settings.UserFilesBackup.Enabled = $false
        $Settings.FBackup.Enabled = $false
        $Settings.Network_Adapter.Enabled = $false
        $Settings.WindowLogonUser.Enabled = $false
        $Settings.Manager.Enabled = $false
        #region Main thread Start
        $MainRunspace =[runspacefactory]::CreateRunspace()      
        $MainRunspace.Open()
        $MainRunspace.SessionStateProxy.SetVariable("Settings",$Settings)  
        $MainRunspace.SessionStateProxy.SetVariable("SettingsOutput",$SettingsOutput)    

        $MainpsCmd = "" | Select-Object PowerShell,Handle
        $MainpsCmd.PowerShell = [PowerShell]::Create().AddScript({ 
            #region Thread Functions
            function Get-UInt32FromIPAddress {
                [CmdletBinding()]
                param ([Parameter(Mandatory=$true)][ipaddress]$IPAddress)
            
                $bytes = $IPAddress.GetAddressBytes()
                if ([BitConverter]::IsLittleEndian) {
                    [Array]::Reverse($bytes)
                }
                return [BitConverter]::ToUInt32($bytes, 0)								  
            }
            function Get-IPAddressFromUInt32 {
                [CmdletBinding()]
                param ([Parameter(Mandatory=$true)][UInt32]$UInt32)
                $bytes = [BitConverter]::GetBytes($UInt32)
                        
                if ([BitConverter]::IsLittleEndian)	{
                    [Array]::Reverse($bytes)
                }
                return New-Object ipaddress(,$bytes)
            }
            #endregion Thread Functions
            #region CustomAppReg Reg Import
            If (Test-Path (${env:ProgramFiles(x86)} + "\" + $Settings.CustomAppFolder)) {
                $Settings.CustomAppFullPath = (${env:ProgramFiles(x86)} + "\" + $Settings.CustomAppFolder)           
            }
            If (Test-Path (${env:ProgramFiles} + "\" + $Settings.CustomAppFolder)) {
                $Settings.CustomAppFullPath =  (${env:ProgramFiles} + "\" + $Settings.CustomAppFolder)
            } 
            if ($Settings.CustomAppFullPath) {
                If ($SettingsOutput) {
                    if ($SettingsOutput.CustomAppRegUser) {
                        #reg import ($env:temp + "\" + $Settings.tempfolder + "\" + $Settings.CustomAppName + "_User.reg") /y
                        $SettingsOutput.CustomAppRegUser | Set-ItemProperty
                    }  
                    if ($SettingsOutput.CustomAppRegx64) {
                        $SettingsOutput.CustomAppRegx64 | Set-ItemProperty
                    }
                    if ($SettingsOutput.CustomAppRegx86) {
                        # Going from x86 to x64 computer; need to convert reg path
                        If ( $SettingsOutput.CustomAppFullPath -contains ${env:ProgramFiles(x86)}) {
                            $SettingsOutput.CustomAppRegx64 = $SettingsOutput.CustomAppRegx86.PSPath.replace("\SOFTWARE","\SOFTWARE\WOW6432Node")
                            $SettingsOutput.CustomAppRegx64 = $SettingsOutput.CustomAppRegx86.PSPath.replace("\SOFTWARE","\SOFTWARE\WOW6432Node")
                            $SettingsOutput.CustomAppRegx64 | Set-ItemProperty
                        }
                        $SettingsOutput.CustomAppRegx86 | Set-ItemProperty
                    }  
                }
            }
            #endregion CustomAppReg Reg Import  
            #region Account setup
            #region Disable
                #Disable all accounts not Admin, User or Blacklist
                #Remove black listed accounts
                If ($Settings.AccountDisableBlacklist -notcontains $LocalUser) {
                    #Disable Accounts.
                    # write-output ("Disabled Non-Window " + $LocalUser.Name + " account . . .")
                    Disable-LocalUser -Name $LocalUser -Confirm:$false

                } 
   
            #endregion Disable and update password for user
            #Enable selected account
                If ($Settings.WindowLogonUser.SelectedItem.ToString()) {
                    Enable-LocalUser -Name ($Settings.WindowLogonUser.SelectedItem.ToString()) -Confirm:$false
                }
            #endregion Account setup
            #region Printers
            If ($SettingsOutput) {
                If ($SettingsOutput.Printers) {
                    If (Get-Command Get-CimInstance -errorAction SilentlyContinue) {
                        $CurrentPrinters = (Get-CimInstance Win32_Printer | Select-Object *)
                    } Else {
                        $CurrentPrinters = (Get-WmiObject Win32_Printer | Select-Object *)
                    } 
                    If (Get-Command Get-CimInstance -errorAction SilentlyContinue) {
                        $CurrentPrinterPorts = (Get-CimInstance win32_tcpipprinterport | Select-Object *)
                    } Else {
                        $CurrentPrinterPorts = (Get-WmiObject win32_tcpipprinterport | Select-Object *)
                    } 
                    If (Get-Command Get-CimInstance -errorAction SilentlyContinue) {
                        $CurrentPrinterDrivers = (Get-CimInstance Win32_PrinterDriver | Select-Object *)
                    } Else {
                        $CurrentPrinterDrivers = (Get-WmiObject Win32_PrinterDriver | Select-Object *)
                    } 
                    $UCPD = $CurrentPrinterDrivers | Where-Object { $_.name -match "Universal"} | Select-Object Name
                    ForEach ($Printer in $SettingsOutput.Printers) {
                        If ($CurrentPrinters | Where-Object {$_.name -eq $Printer.Printer_Name}) {
                            #Write-Host ("Already Mapped Printer: " + $Printer.Printer_Name)
                        } Else {
                            If ($Printer.Printer_Port_Type) {
                                Write-Host ("Mapping Network Printer: " + $Printer.Printer_Name)
                                If ($CurrentPrinterPorts | Where-Object {$_.Name -eq $Printer.Printer_Port_Name}) {
                                    Write-Host ("`tAlready Created Network Printer Port: " + $Printer.Printer_Port_Name)
                                } Else {
                                    If ($Printer.Printer_Port_Queue) {
                                        Write-Host ("`t`tCreating LPR Printer Port")
                                        Add-PrinterPort -Name $Printer.Printer_Port_Name -LprHostAddress $Printer.Printer_Port_IP -LprQueueName $Printer.Printer_Port_Queue
                                        #CreatePrinterPort -PrinterIP $PrinterIP -PrinterPort $PrinterPort -PrinterPortName $PrinterPortName -Computer $Computer
                                    } Else {                     
                                        If ($Printer.Printer_Port_SNMPCommunity) {
                                            Write-Host ("`t`tCreating Raw Printer Port with SNMP")
                                            Add-PrinterPort -Name $Printer.Printer_Port_Name -PrinterHostAddress $Printer.Printer_Port_IP -SNMPCommunity $Printer.Printer_Port_SNMPCommunity -SNMP:$Printer.Printer_Port_SNMPEnabled
                                        } Else {
                                            Write-Host ("`t`tCreating Raw Printer Port")
                                            #Add-PrinterPort -Name $Printer.Printer_Port_Name -PrinterHostAddress $Printer.Printer_Port_IP
                                            New-PrinterPort -PrinterIP $Printer.Printer_Port_IP -PrinterPort $PrinterPort -PrinterPortName $Printer.Printer_Port_Name -Computer $Computer
                                        }
                                    }
                                }
                                If ($CurrentPrinterDrivers | Where-Object { $_.Name -eq $Printer.Printer_DriverName}) {
                                    Write-Host ("`tCreating Network Printer: " + $Printer.Printer_Name)
                                    #Add-Printer -Name $printer.Printer_Name -PortName $Printer.Printer_Port_Name -DriverName $Printer.Printer_DriverName
                                    New-Printer -PrinterPortName $Printer.Printer_Port_Nam -DriverName $Printer.Printer_DriverName -PrinterCaption $printer.Printer_Name -Computer $Computer
                                } Else {
                                    Switch -Wildcard ($Printer.Printer_DriverName) {
                                        "*HP*" {
                                            If (($UCPD | Where-Object {$_.name -match "HP"}).name) {
                                                Write-Host ("Re-Mapping Printer Driver with : " + ($UCPD | Where-Object {$_.name -match "HP"}).name)
                                                #Add-Printer -Name $printer.Printer_Name -PortName $Printer.Printer_Port_Name -DriverName ($UCPD | Where-Object {$_.name -match "HP"}).name
                                                New-Printer -PrinterPortName $Printer.Printer_Port_Nam -DriverName ($UCPD | Where-Object {$_.name -match "HP"}).name -PrinterCaption $printer.Printer_Name -Computer $Computer
                                            }
                                            break
                                        }
                                        "*Samsung*" {
                                            If (($UCPD | Where-Object {$_.name -match "HP"}).name) {
                                                Write-Host ("Re-Mapping Printer Driver with : " + ($UCPD | Where-Object {$_.name -match "HP"}).name)
                                                #Add-Printer -Name $printer.Printer_Name -PortName $Printer.Printer_Port_Name -DriverName ($UCPD | Where-Object {$_.name -match "HP"}).name
                                                New-Printer -PrinterPortName $Printer.Printer_Port_Nam -DriverName ($UCPD | Where-Object {$_.name -match "HP"}).name -PrinterCaption $printer.Printer_Name -Computer $Computer
                                            }
                                            break
                                        }
                                        "*KONICA MINOLTA*" {
                                            
                                            If (($UCPD | Where-Object {$_.name -match "KONICA MINOLTA"}).name) {
                                                Write-Host ("Re-Mapping Printer Driver with : " + ($UCPD | Where-Object {$_.name -match "KONICA MINOLTA"}).name)
                                                #Add-Printer -Name $printer.Printer_Name -PortName $Printer.Printer_Port_Name -DriverName ($UCPD | Where-Object {$_.name -match "KONICA MINOLTA"}).name
                                                New-Printer -PrinterPortName $Printer.Printer_Port_Nam -DriverName ($UCPD | Where-Object {$_.name -match "KONICA MINOLTA"}).name -PrinterCaption $printer.Printer_Name -Computer $Computer
                                            }
                                            break
                                        }
                                        default {
                                            Write-Host ("`tCould not re-map driver!")
                                            break
                                        }
                                    }
                                                    
                                }
                            }
                            If ($Printer.Printer_ServerName) {
                                Write-Host ("Mapping Shared Printer: " + $Printer.Printer_Name)
                                #Add-Printer -ConnectionName $Printer.Printer_Name
                                (New-Object -ComObject WScript.Network).AddWindowsPrinterConnection($Printer.Printer_Name)
                            }
                        }
                    }
                    #Default
                    If (Get-Command Get-CimInstance -errorAction SilentlyContinue) {
                        $CurentDefault = (Get-CimInstance -Query " Select name FROM Win32_Printer WHERE Default=$true").Name
                    } Else {
                        $CurentDefault = (Get-WmiObject -Query " Select name FROM Win32_Printer WHERE Default=$true").Name
                    } 
                    $OldDefault = ($ImportCVS | Where-Object {$_.Printer_Default -eq $true}).Printer_Name
                    If ($CurrentDefault -ne $OldDefault) {
                        (New-Object -ComObject WScript.Network).SetDefaultPrinter($OldDefault)
                    }

                }
            }
            #endregion Printers
            #region Restore files
            If ($Settings.Restore_Backup.Text) {			
                ForEach ($Restore in $Settings.BackupFolders) {
                    $CFN = Split-Path -Leaf $Restore
                    #Create Folder for restored folder
                    If (!(Test-Path($Restore))) {
                        New-Item -ItemType Directory -Path ($Restore)
                    }
                    #Powershell 5.1 and newer zip
                    If (Test-Path($env:temp + "\" + $Settings.tempfolder + "\" + $CFN)) {
                        If ($CFN -eq $Settings.CustomAppName) {
                            robocopy /e ($env:temp + "\" + $Settings.tempfolder + "\" + $CFN) $Restore /w:3 /r:3 /XD *Image* /XD Export /XD Logs /XD Reports /XD test /XD ENV /XD Update* /XF ( $Settings.CustomAppName + " release notes*.pdf") /XF *.BAK /XF *.log /XF *text*.txt /XF *temp*.* /XF _* /XF ~* /XF Thumbs.db | Out-Null
                        } else {
                            robocopy /e ($env:temp + "\" + $Settings.tempfolder + "\" + $CFN) $Restore /w:3 /r:3 /XF ~* /XF Thumbs.db | Out-Null
                        }
                    }
                    #Powershell 2.0 Zip
                    If (Test-Path($env:temp + "\" + $Settings.tempfolder + "\" + $Settings.tempfolder + "\" + $CFN)) {
                        If ($CFN -eq $Settings.CustomAppName) {
                            robocopy /e ($env:temp + "\" + $Settings.tempfolder + "\" + $Settings.tempfolder + "\" + $CFN) $Restore /w:3 /r:3 /XD *Image* /XD Export /XD Logs /XD Reports /XD test /XD ENV /XD Update* /XF ( $Settings.CustomAppName + " release notes*.pdf") /XF *.BAK /XF *.log /XF *text*.txt /XF *temp*.* /XF _* /XF ~* /XF Thumbs.db | Out-Null
                        } else {
                            robocopy /e ($env:temp + "\" + $Settings.tempfolder + "\" + $Settings.tempfolder + "\" + $CFN) $Restore /w:3 /r:3 /XF ~* /XF Thumbs.db | Out-Null
                        }
                    }
                }
            }
            #endregion Restore files
            #region Set Machine IP
            If ($Settings.IP_Address.Text) {
                If ($Settings.Network_Adapter.SelectedItem.ToString()) {
                    $Settings.NetworkAddress = [IPAddress] (([IPAddress]$Settings.IP_Address.Text ).Address -band ([IPAddress] $Settings.Subnet).Address)
                    If (Get-Command Get-CimInstance -errorAction SilentlyContinue) {
                        $wmi = Get-CimInstance win32_networkadapterconfiguration -filter ("Description = '" + $Settings.Network_Adapter.SelectedItem.ToString() + "'")
                    } Else {
                        $wmi = Get-WmiObject win32_networkadapterconfiguration -filter ("Description = '" + $Settings.Network_Adapter.SelectedItem.ToString() + "'")
                    } 
                    #Only change IP if it is different.
                    If (( $wmi.ipaddress | Where-object {$_.IPaddress -notlike "169.254.*" -and $_.IPAddress -ne "127.0.0.1"}) -ne $Settings.IP_Address.Text) {
                        $wmi.EnableStatic($Settings.IP_Address.Text, $Settings.Subnet)              
                        $wmi.SetGateways((Get-IPAddressFromUInt32 -UInt32 ((Get-UInt32FromIPAddress -IPAddress $Settings.NetworkAddress.IpAddressToString) +1)).IPAddressToString, 1)        
                        $wmi.SetDNSServerSearchOrder($Settings.DNS)
                    }
                }
            }
		 
            #endregion Set Machine IP
            #region Set Machine Name
            If ($Settings.Machine_Name.Text) {
                If (Get-Command Rename-computer -errorAction SilentlyContinue) {
                      Rename-computer -NewName $Settings.Machine_Name.Text  -force 
                } Else {
                    If (Get-Command Get-CimInstance -errorAction SilentlyContinue) {
                        $ComputerInfo = Get-CimInstance -Class Win32_ComputerSystem
                    } Else {
                        $ComputerInfo = Get-WmiObject -Class Win32_ComputerSystem
                    } 
                    #Only change Name if it is different.
                    If ($ComputerInfo.Name.ToLower() -ne $Settings.Machine_Name.Text.ToLower()) {
                        $ComputerInfo.Rename($Settings.Machine_Name.Text)
                    }
                }
            }
		 
            #endregion Set Machine Name
            #region Managers 
            If ($Settings.Manager.Checked) {
                #Mounted User Hive Location
                $HKEY = ("HKU\H_" + $Settings.WindowLogonUser.SelectedItem.ToString())           
                New-PSDrive -PSProvider Registry -Name HKU -Root HKEY_USERS -erroraction 'silentlycontinue' | Out-Null
                #Get Hive file location
                $CurrentUserSID = (Get-LocalUser -Name $Settings.WindowLogonUser.SelectedItem.ToString()).SID
                If (Get-Command Get-CimInstance -errorAction SilentlyContinue) {
                    $UserProfile = (Get-CimInstance Win32_UserProfile | Where-Object { $_.SID -eq $CurrentUserSID}).localpath
                } Else {
                    $UserProfile = (Get-WmiObject Win32_UserProfile | Where-Object { $_.SID -eq $CurrentUserSID}).localpath
                }
                #Add Current user to Hive.
                $user_account=$env:username
                $Acl = Get-Acl $UserProfile
                $Ar = New-Object system.Security.AccessControl.FileSystemAccessRule($user_account, "FullControl", "ContainerInherit, ObjectInherit", "None", "Allow")
                $Acl.Setaccessrule($Ar)
                Set-Acl $UserProfile $Acl
                #Mount user Hive 
                If (Test-Path ($UserProfile + "\ntuser.dat")) { 
                    [gc]::collect()
                    $process = (REG LOAD  $HKEY ($UserProfile + "\ntuser.dat"))
                    If ($LASTEXITCODE -ne 0 ) {
                        write-error ( "Cannot load profile for: " + ($UserProfile + "\ntuser.dat") )
                        continue
                    }
                }else{
                    If (Test-Path $UserProfile.Replace($UserProfile.Substring(0,1),($env:systemdrive).Substring(0,1))) {
                        # REG LOAD $HKEY ($UserProfile + "\ntuser.dat")
                        [gc]::collect()
                        $process = (REG LOAD  $HKEY ($UserProfile + "\ntuser.dat"))
                        If ($LASTEXITCODE -ne 0 ) {
                            write-error ( "Cannot load profile for: " + ($Settings.UsersProfileFolder + "\" + $Settings.WindowLogonUser.SelectedIndex.ToString() + "\ntuser.dat") )
                            continue
                        }		
                    }else{
                        write-error ( "Cannot load profile for: " + ($Settings.UsersProfileFolder + "\" + $Settings.WindowLogonUser.SelectedIndex.ToString() + "\ntuser.dat") )
                        continue
                    }
                }
                #region Start Relaxing Setting for Managers #
                If (-Not (Test-Path -Path $HKEY.replace("HKU\","HKU:\"))) {
                    [System.Windows.MessageBox]::Show("Error: Loading Managers Registry. Please manually edit Manager's." ,'Error: Loading Managers Registry','OK','Error')
                }
                #Shows Run in Start and allows UNC paths.	
                Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoRun" 0 "DWORD"
                #Show all drives in Windows Explorer	
                Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoDrives" 0 "DWORD"
                #Enable user to using My Computer to gain access to the content of selected drives. 
                Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoViewOnDrive" 0 "DWORD"
                #Enable Context-sensitive menus .
                Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoTrayContextMenu" 0 "DWORD"
                #Enable right-click on Desktop and Windows Explorer
                Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoViewContextMenu" 0 "DWORD"
                #Enable right-click on Start Menu
                Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "DisableContextMenusInStart" 0 "DWORD"
                #Enable Context Menus in the Start Menu in Windows 10
                Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "DisableContextMenusInStart" 0 "DWORD"
                #Shows "This PC" in Windows Explorer
                Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\NonEnum") "{20D04FE0-3AEA-1069-A2D8-08002B30309D}" 0 "DWORD"
                #Enable  Right Click in Internet Explorer
                Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Internet Explorer\Restrictions") "NoBrowserContextMenu" 0 "DWORD"
                # Adds Desktop from This PC 
                #write-host ("`tDesktop folder from This PC ") -foregroundcolor "gray"
                If(-Not (Test-Path ("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}"))) {
                    New-Item  ("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}")
                }
                If(-Not (Test-Path ("HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}"))) {
                    New-Item  ("HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}")
                }
                # Adds Documents from This PC 
                #write-host ("`tDocuments folder from This PC ") -foregroundcolor "gray"
                If(-Not (Test-Path ("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{A8CDFF1C-4878-43be-B5FD-F8091C1C60D0}"))) {
                    New-Item ("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{A8CDFF1C-4878-43be-B5FD-F8091C1C60D0}")
                }
                If(-Not (Test-Path ("HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{A8CDFF1C-4878-43be-B5FD-F8091C1C60D0}"))) {
                    New-Item  ("HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{A8CDFF1C-4878-43be-B5FD-F8091C1C60D0}")
                }
                If(-Not (Test-Path ("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{d3162b92-9365-467a-956b-92703aca08af}"))) {
                    New-Item  ("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{d3162b92-9365-467a-956b-92703aca08af}")
                }
                If(-Not (Test-Path ("HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{d3162b92-9365-467a-956b-92703aca08af}"))) {
                    New-Item  ("HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{d3162b92-9365-467a-956b-92703aca08af}")
                }
                # Adds Downloads from This PC 
                #write-host ("`tDownloads folder from This PC ") -foregroundcolor "gray"
                If(-Not (Test-Path ("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{088e3905-0323-4b02-9826-5d99428e115f}"))) {
                    New-Item  ("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{088e3905-0323-4b02-9826-5d99428e115f}") 
                }
                If(-Not (Test-Path ("HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{088e3905-0323-4b02-9826-5d99428e115f}"))) {
                    New-Item  ("HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{088e3905-0323-4b02-9826-5d99428e115f}") 
                }
                If(-Not (Test-Path ("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{374DE290-123F-4565-9164-39C4925E467B}"))) {
                    New-Item  ("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{374DE290-123F-4565-9164-39C4925E467B}") 
                }
                If(-Not (Test-Path ("HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{374DE290-123F-4565-9164-39C4925E467B}"))) {
                    New-Item  ("HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{374DE290-123F-4565-9164-39C4925E467B}") 
                }
                #Adds Pictures (folder) from This PC 
                #write-host ("`tPictures folder from This PC ")  -foregroundcolor "gray"
                If(-Not (Test-Path ("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{3ADD1653-EB32-4cb0-BBD7-DFA0ABB5ACCA}"))) {
                    New-Item  ("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{3ADD1653-EB32-4cb0-BBD7-DFA0ABB5ACCA}") 
                }
                If(-Not (Test-Path ("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{3ADD1653-EB32-4cb0-BBD7-DFA0ABB5ACCA}"))) {
                    New-Item  ("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{3ADD1653-EB32-4cb0-BBD7-DFA0ABB5ACCA}") 
                }
                If(-Not (Test-Path ("HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{3ADD1653-EB32-4cb0-BBD7-DFA0ABB5ACCA}"))) {
                    New-Item  ("HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{3ADD1653-EB32-4cb0-BBD7-DFA0ABB5ACCA}") 
                }
                If(-Not (Test-Path ("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{24ad3ad4-a569-4530-98e1-ab02f9417aa8}"))) {
                    New-Item  ("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{24ad3ad4-a569-4530-98e1-ab02f9417aa8}") 
                    Set-Reg "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer" "{24AD3AD4-A569-4530-98E1-AB02F9417AA8}" 1 "DWORD"
                }
                If(-Not (Test-Path ("HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{24ad3ad4-a569-4530-98e1-ab02f9417aa8}"))) {
                    New-Item  ("HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{24ad3ad4-a569-4530-98e1-ab02f9417aa8}") 
                    Set-Reg "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer" "{24AD3AD4-A569-4530-98E1-AB02F9417AA8}" 1 "DWORD"
                }
                
                #Unload Manager
                [gc]::collect()
                $process = (REG UNLOAD $HKEY)
                If ($LASTEXITCODE -ne 0 ) {
                    [gc]::collect()
                    Start-Sleep 3
                    $process = (REG UNLOAD $HKEY)
                    If ($LASTEXITCODE -ne 0 ) {
                        write-error ("`t" + $UserProfile + ": Can not unload user registry!")
                    }
                }
                #endregion Start Relaxing Setting for Managers #
                #region Ask about office install
                    If (Test-Path -Path ($Settings.OfficeSubFolder + "\setup.exe") -ErrorAction SilentlyContinue) {
                        $OfficeConfig = (Get-ChildItem -Path ($Settings.OfficeSubFolder + "\config\*.xml") | Select-Object -First 1)
                        If ($OfficeConfig.FullName) {
                            If ([System.Windows.MessageBox]::Show(('Would you like to Install Office 2019?'),'Install Office?','YesNo','Question') -eq "Yes") {
                                Start-Process -FilePath ($Settings.OfficeSubFolder + "\setup.exe") -ArgumentList "/configure",('"' + $OfficeConfig + '"') -Wait
                                If (Test-Path -Path $Settings.OfficeActivationScript) {
                                    Start-Process -FilePath $Settings.OfficeActivationScript -Wait
                                } Else {
                                    [System.Windows.MessageBox]::Show("Error: Activating Office. Please manually Activate Office." ,'Error: Activating Office','OK','Error')
                                }

                            }
                        }
                    }
                #endregion Ask about office install
            }
            #endregion Managers 
            #region Stop Autologon
                #Ask to set Local Admin.
                If([System.Windows.MessageBox]::Show(('Would you like to disable auto logon?'),('Disable auto logon?'),'YesNo','Question') -eq "Yes" -and $SetPassAdmin) {
                    If ((Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name 'AutoAdminLogon' -ErrorAction SilentlyContinue).AutoAdminLogon -ne 0) {
                        Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name 'AutoAdminLogon' -Value '0'
                    }
                    If ((Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name 'DefaultUserName' -ErrorAction SilentlyContinue).DefaultUserName) {  
                        Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name 'DefaultUserName' -Value ''
                    }
                    If ((Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name 'DefaultPassword' -ErrorAction SilentlyContinue).DefaultPassword) {  
                        Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name 'DefaultPassword' -Value ''
                    }
                }
            #endregion Stop Autologon

            #Reboot after all done.
            If([System.Windows.MessageBox]::Show(('Would you like to Reboot?'),'System Reboot','YesNo','Question') -eq "Yes") {
                    Restart-Computer
            }  

        })
        $MainpsCmd.Powershell.Runspace = $MainRunspace
        $MainpsCmd.Handle = $MainpsCmd.Powershell.BeginInvoke()
        
        While ($MainpsCmd.Handle.IsCompleted -ne $true) {
            Start-Sleep -Milliseconds 100
            [gc]::collect()
											 
        }

        [gc]::collect()
        $Settings.sw.Stop()
        [gc]::collect()
        $Settings.Store_Setup.text = ( $Settings.WindowTitle + " Done. Time: " + (FormatElapsedTime($Settings.sw.Elapsed)) ) 
        $MainpsCmd.Powershell.EndInvoke($MainpsCmd.Handle)
        [gc]::collect()
        $Settings.Stop.Text = "Exit"
        [gc]::collect()
        #$Settings.Store_Setup.Close()
        #[void]$Settings.Store_Setup.Close()
        #endregion Main thread End
    }
}
function Stop_Work {
    param (
        
    )
    If( $Settings.Stop.Text -eq "Stop") {
        If ($BrowsepsCmd) {
            $BrowsepsCmd.Stop()
        }
        If ($MainpsCmd) {
            $MainpsCmd.Stop()
        }
        $Settings.Store_Setup.text                = ( $Settings.WindowTitle + " Cleaning Up. Please Wait . . ." )
        Start-Sleep -Seconds 5
        if (Test-Path ($Settings.tempfolder)) {
            Remove-Item -Path ($Settings.tempfolder) -Force -Recurse
        }
        [void]$Settings.Store_Setup.Close()

        #Exit
    } Else {
        [void]$Settings.Store_Setup.Close()
    }
}
#############################################################################
#endregion Functions
#############################################################################
#############################################################################
#region Setup Sessions
#############################################################################
Hide-Console
#Load .Net Classes
#Popup
Add-Type -AssemblyName PresentationCore,PresentationFramework
#Form
Add-Type -AssemblyName System.Windows.Forms
#Password Generation
Add-Type -AssemblyName System.web

[System.Windows.Forms.Application]::EnableVisualStyles()

$Settings.Store_Setup                     = New-Object system.Windows.Forms.Form
# $Store_Setup.ClientSize          = '400,500'
$Settings.Store_Setup.ClientSize          = '400,300'
$Settings.Store_Setup.text                = $Settings.WindowTitle
$Settings.Store_Setup.TopMost             = $false
#Show Icon https://stackoverflow.com/questions/53376491/powershell-how-to-embed-icon-in-powershell-gui-exe
If ($iconBase64) {
    $iconBytes       = [Convert]::FromBase64String($iconBase64)
    $stream          = New-Object IO.MemoryStream($iconBytes, 0, $iconBytes.Length)
    $stream.Write($iconBytes, 0, $iconBytes.Length);
    $Settings.Store_Setup.icon       = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -Argument $stream).GetHIcon())
}

$Settings.Restore_Backup_Label               = New-Object system.Windows.Forms.Label
$Settings.Restore_Backup_Label.text          = "Restore Backup:"
$Settings.Restore_Backup_Label.AutoSize      = $true
$Settings.Restore_Backup_Label.width         = 25
$Settings.Restore_Backup_Label.height        = 10
$Settings.Restore_Backup_Label.location      = New-Object System.Drawing.Point(14,10)
$Settings.Restore_Backup_Label.Font          = 'Microsoft Sans Serif,10'

$Settings.Restore_Backup                     = New-Object system.Windows.Forms.TextBox
$Settings.Restore_Backup.multiline           = $false
$Settings.Restore_Backup.width               = 194
$Settings.Restore_Backup.height              = 20
$Settings.Restore_Backup.location            = New-Object System.Drawing.Point(125,10)
$Settings.Restore_Backup.Font                = 'Microsoft Sans Serif,10'
$Settings.Restore_Backup.Enabled             = $false

$Settings.Browse                          = New-Object system.Windows.Forms.Button
$Settings.Browse.text                     = "Browse..."
$Settings.Browse.width                    = 70
$Settings.Browse.height                   = 25

$Settings.Browse.location                 = New-Object System.Drawing.Point(320,10)
$Settings.Browse.Font                     = 'Microsoft Sans Serif,10'
# If (Test-Path (Split-Path -Parent -Path $MyInvocation.MyCommand.Definition)) {
#     $Settings.Restore_Backup.Text = ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\" + $Settings.tempfolder + ".zip" )  
# }

$Settings.Machine_Name_Label              = New-Object system.Windows.Forms.Label
$Settings.Machine_Name_Label.text         = "Machine Name:"
$Settings.Machine_Name_Label.AutoSize     = $true
$Settings.Machine_Name_Label.width        = 25
$Settings.Machine_Name_Label.height       = 10
$Settings.Machine_Name_Label.location     = New-Object System.Drawing.Point(10,40)
$Settings.Machine_Name_Label.Font         = 'Microsoft Sans Serif,10'

$Settings.Machine_Name                    = New-Object system.Windows.Forms.TextBox
$Settings.Machine_Name.multiline          = $false
$Settings.Machine_Name.width              = 180
$Settings.Machine_Name.height             = 20
$Settings.Machine_Name.location           = New-Object System.Drawing.Point(125,40)
$Settings.Machine_Name.Font               = 'Microsoft Sans Serif,10'
# $Settings.Machine_Name.Enabled            = $false
$Settings.Machine_Name.text               = $env:computername


If (Get-Command Get-CimInstance -errorAction SilentlyContinue) {
    $Settings.Network_Adapter_List = Get-CimInstance -Class Win32_NetworkAdapterConfiguration -Filter 'IPEnabled = True' | Where-object {$_.IPaddress -notlike "169.254.*" -and $_.IPAddress -ne "127.0.0.1"} | Select-Object Description,IPAddress,DefaultIPGateway,IPSubnet,DNSServerSearchOrder
} Else {
    $Settings.Network_Adapter_List = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter 'IPEnabled = True' | Where-object {$_.IPaddress -notlike "169.254.*" -and $_.IPAddress -ne "127.0.0.1"} | Select-Object Description,IPAddress,DefaultIPGateway,IPSubnet,DNSServerSearchOrder
} 

$Settings.IP_Address_Label                = New-Object system.Windows.Forms.Label
$Settings.IP_Address_Label.text           = "IP Address:"
$Settings.IP_Address_Label.AutoSize       = $true
$Settings.IP_Address_Label.width          = 25
$Settings.IP_Address_Label.height         = 10
$Settings.IP_Address_Label.location       = New-Object System.Drawing.Point(10,65)
$Settings.IP_Address_Label.Font           = 'Microsoft Sans Serif,10'

$Settings.IP_Address                      = New-Object system.Windows.Forms.TextBox
$Settings.IP_Address.multiline            = $false
$Settings.IP_Address.width                = 180
$Settings.IP_Address.height               = 20
$Settings.IP_Address.location             = New-Object System.Drawing.Point(125,65)
$Settings.IP_Address.Font                 = 'Microsoft Sans Serif,10'
$Settings.IP_Address.Text                 = ($Settings.Network_Adapter_List | Select-Object -first 1).IPAddress | Where-Object {$_ -notlike '*:*'}
# $Settings.IP_Address.Enabled              = $false

$Settings.Network_Adapter_Label                = New-Object system.Windows.Forms.Label
$Settings.Network_Adapter_Label.text           = "Network Adapter:"
$Settings.Network_Adapter_Label.AutoSize       = $true
$Settings.Network_Adapter_Label.width          = 25
$Settings.Network_Adapter_Label.height         = 10
$Settings.Network_Adapter_Label.location       = New-Object System.Drawing.Point(10,95)
$Settings.Network_Adapter_Label.Font           = 'Microsoft Sans Serif,10'

$Settings.Network_Adapter                       = New-Object system.Windows.Forms.ComboBox
#$Settings.Network_Adapter.text                  = " "
$Settings.Network_Adapter.width                 = 265
$Settings.Network_Adapter.height                = 20
$Settings.Network_Adapter.location              = New-Object System.Drawing.Point(125,95)
$Settings.Network_Adapter.Font                  = 'Microsoft Sans Serif,10'


$Settings.WindowLogonUser_Label                = New-Object system.Windows.Forms.Label
$Settings.WindowLogonUser_Label.text           = "Window User:"
$Settings.WindowLogonUser_Label.AutoSize       = $true
$Settings.WindowLogonUser_Label.width          = 25
$Settings.WindowLogonUser_Label.height         = 10
$Settings.WindowLogonUser_Label.location       = New-Object System.Drawing.Point(10,125)
$Settings.WindowLogonUser_Label.Font           = 'Microsoft Sans Serif,10'

$Settings.WindowLogonUser                       = New-Object system.Windows.Forms.ComboBox
$Settings.WindowLogonUser.text                  = " "
$Settings.WindowLogonUser.width                 = 265
$Settings.WindowLogonUser.height                = 20
$Settings.WindowLogonUser.location              = New-Object System.Drawing.Point(125,125)
$Settings.WindowLogonUser.Font                  = 'Microsoft Sans Serif,10'


$Settings.Manager                      = New-Object System.Windows.Forms.Checkbox 
$Settings.Manager.Text                 = "Manager"
$Settings.Manager.width                = 180
$Settings.Manager.height               = 20
$Settings.Manager.Location             = New-Object System.Drawing.Size(125,150) 
$Settings.Manager.Font                 = 'Microsoft Sans Serif,10'
$Settings.Manager.Checked              = $False


$Settings.FBackup = New-Object System.Windows.Forms.GroupBox #create the group box
$Settings.FBackup.Location = New-Object System.Drawing.Size(10,170) #location of the group box (px) in relation to the primary window's edges (length, height)
$Settings.FBackup.size = New-Object System.Drawing.Size(375,70) #the size in px of the group box (length, height)
$Settings.FBackup.text = "Restore:" #labeling the box
$Settings.FBackup.Enabled = $false

$Settings.CABackup                      = New-Object System.Windows.Forms.Checkbox 
$Settings.CABackup.Text                 = $Settings.CustomAppName
$Settings.CABackup.width                = 180
$Settings.CABackup.height               = 20
# $Settings.CABackup.Location             = New-Object System.Drawing.Size(115,65) 
$Settings.CABackup.Location             = New-Object System.Drawing.Size(10,15) 
$Settings.CABackup.Font                 = 'Microsoft Sans Serif,10'
$Settings.CABackup.Checked              = $true
$Settings.CABackup.Enabled              = $false

$Settings.UserFilesBackup                      = New-Object System.Windows.Forms.Checkbox 
$Settings.UserFilesBackup.Text                 = "User Files"
$Settings.UserFilesBackup.width                = 180
$Settings.UserFilesBackup.height               = 20
# $Settings.UserFilesBackup.Location             = New-Object System.Drawing.Size(115,85) 
$Settings.UserFilesBackup.Location             = New-Object System.Drawing.Size(10,40) 
$Settings.UserFilesBackup.Font                 = 'Microsoft Sans Serif,10'
$Settings.UserFilesBackup.Checked              = $true
$Settings.UserFilesBackup.Enabled              = $false

$Settings.FBackup.Controls.AddRange(@($Settings.CABackup,$Settings.UserFilesBackup)) #activate the inside the group box


$Settings.Stop                         = New-Object system.Windows.Forms.Button
$Settings.Stop.text                    = "Exit"
$Settings.Stop.width                   = 70
$Settings.Stop.height                  = 25
$Settings.Stop.location                = New-Object System.Drawing.Point(250,270)
$Settings.Stop.Font                    = 'Microsoft Sans Serif,10'

$Settings.Start                         = New-Object system.Windows.Forms.Button
$Settings.Start.text                    = "Update"
$Settings.Start.width                   = 70
$Settings.Start.height                  = 25
$Settings.Start.location                = New-Object System.Drawing.Point(320,270)
$Settings.Start.Font                    = 'Microsoft Sans Serif,10'
# $Settings.Start.Enabled                 = $false

$Settings.Store_Setup.controls.AddRange(@($Settings.Machine_Name_Label,$Settings.IP_Address_Label,$Settings.Machine_Name,$Settings.IP_Address,$Settings.Network_Adapter_Label,$Settings.Network_Adapter,$Settings.Restore_Backup,$Settings.Start,$Settings.Stop,$Settings.Restore_Backup_Label,$Settings.Browse,$Settings.WindowLogonUser_Label,$Settings.WindowLogonUser,$Settings.Manager,$Settings.FBackup))


#############################################################################
#endregion Setup Sessions
#############################################################################
#############################################################################
#region Main 
#############################################################################

$Settings.Browse.Add_Click({ Browse_File })
$Settings.Start.Add_Click({ Start_Work })
$Settings.Stop.Add_Click({ Stop_Work })
#$Settings.WindowLogonUser.Add_SelectedIndexChanged({  })

ForEach ( $LocalUser in ((Get-LocalUser).name | Sort-Object {"$_" -replace '\d',''},{("$_" -replace '\D','') -as [int]})) {
    If (-Not ($Settings.AccountBlacklist.contains($LocalUser))) {
        $Settings.WindowLogonUser.Items.Add($LocalUser)
    }
}

If ($Settings.WindowLogonUserReg.DefaultUserName) {
    $Settings.WindowLogonUser.SelectedItem = $Settings.WindowLogonUserReg.DefaultUserName
}else {
    # If ($Settings.WindowLogonUser.SelectionLength -ge 0) {
    #     $Settings.WindowLogonUser.SelectedIndex = 0
    # }
}

ForEach ( $NIC in $Settings.Network_Adapter_List) {
    If ($NIC.InterfaceAlias) {
        $Settings.Network_Adapter.Items.Add($NIC.InterfaceAlias)
    } else {
        $Settings.Network_Adapter.Items.Add($NIC.Description)
    }
}

If ($Settings.Network_Adapter.SelectionLength -ge 0) {
    $Settings.Network_Adapter.SelectedIndex = 0
}


[void]$Settings.Store_Setup.ShowDialog()

#############################################################################
#endregion Main
#############################################################################
