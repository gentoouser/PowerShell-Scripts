<# 
.SYNOPSIS
    Name: Store_Backup.ps1
    Exports Store Settings to Zip file

.DESCRIPTION
    Backups up user data and custom app settings.


.PARAMETER


.EXAMPLE
   & Store_Backup.ps1

.NOTES
 Author: Paul Fuller
 Changes:
    1.0.1 - Added Installed Program to output.
    1.0.1 - Combined Printer and Printer port information into one object
    1.0.2 - Fixed folder redirect backup. Cleaned up interface.

#>
#Force Starting of Powershell script as Administrator 
If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {   
	$arguments = "& '" + $myinvocation.mycommand.definition + "'"
	Start-Process powershell -Verb runAs -ArgumentList $arguments
	Break
}

#############################################################################
#region User Variables
#############################################################################
$ScriptVersion = "1.0.2"
$Settings =[hashtable]::Synchronized(@{})
#$Settings =@{}
$SettingsOutput =[hashtable]::Synchronized(@{})
#$SettingsOutput =@{}

$Settings.WindowTitle = " Store Backup"
$Settings.tempfolder = ($env:computername + "_" + (Get-Date -format yyyyMMdd-hhmm))
$Settings.CustomAppFolder = "CustomApp"
$Settings.CustomAppRegKey = "CustomApp"
$Settings.CustomAppName = "CustomApp"
$USF = (Get-ItemProperty -path "hkcu:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders")

$Settings.BackupFolders =@()
If (Test-Path (${env:ProgramFiles(x86)} + "\" + $Settings.CustomAppFolder)) {
    $Settings.BackupFolders += (${env:ProgramFiles(x86)} + "\" + $Settings.CustomAppFolder)
}
If (Test-Path (${env:ProgramFiles} + "\" + $Settings.CustomAppFolder)) {
    $Settings.BackupFolders += (${env:ProgramFiles} + "\" + $Settings.CustomAppFolder)
} 
$Settings.BackupFolders += $([string]$USF.Desktop)
$Settings.BackupFolders += $([string]$USF.Favorites)
$Settings.BackupFolders += $([string]$USF."My Pictures")
$Settings.BackupFolders += $([string]$USF."{374DE290-123F-4565-9164-39C4925E467B}") #Downloads
$Settings.BackupFolders += $([string]$USF.Personal) #My Documents

$iconBase64 ='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'

#############################################################################
#endregion User Variables
#############################################################################
#############################################################################
#region Functions
#############################################################################
function FormatElapsedTime {
    param (
        $ts
    )
    #https://stackoverflow.com/questions/3513650/timing-a-commands-execution-in-powershell
	$elapsedTime = $null

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
function Browse_File {
    param (
        
    )
    $Settings.SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $Settings.SaveFileDialog.initialDirectory = (Split-Path -Parent -Path $Settings.Save_Backup.Text) 
    $Settings.SaveFileDialog.filename =  [io.path]::GetFileNameWithoutExtension($Settings.Save_Backup.Text)
    $Settings.SaveFileDialog.filter = "ZIP Archive Files|*.zip|All Files|*.*" 
    If ($Settings.SaveFileDialog.ShowDialog() -eq 'OK' ) {
       $Settings.Save_Backup.Text = $Settings.SaveFileDialog.filename
    }
}
function Start_Work {
    param (
        
    )
	$Settings.Stop.text = "Stop"
    $Settings.Start.Enabled = $false 
    $Settings.Machine_Name.Enabled = $false 
    $Settings.IP_Address.Enabled = $false 
    $Settings.UserFilesBackup.Enabled = $false
    $Settings.UserFilesBackup.Controls | ForEach-Object { $_.Enabled = $false}
    $Settings.CABackup.Enabled = $false
    $Settings.UserFilesBackup.Enabled = $false
    $Settings.Browse.Enabled = $false
    $Settings.sw = [Diagnostics.Stopwatch]::StartNew()
    $Settings.Store_Setup.text                = ( $Settings.WindowTitle + " Working . . . Please wait.")
    $Settings.Browse.Enabled = $false
    $Settings.Save_Backup.Enabled = $false
    $Settings.IP_Address.Enabled = $false
    $Settings.Machine_Name.Enabled = $false
    $Settings.Start.Enabled = $false

    #region Main thread Start
     $MainRunspace =[runspacefactory]::CreateRunspace()      
     $MainRunspace.Open()
     $MainRunspace.SessionStateProxy.SetVariable("Settings",$Settings)  
     $MainRunspace.SessionStateProxy.SetVariable("SettingsOutput",$SettingsOutput)    

     $MainpsCmd = "" | Select-Object PowerShell,Handle
     $MainpsCmd.PowerShell = [PowerShell]::Create().AddScript({ 
        #create temp . . .
        if (Test-Path ($env:temp + "\" + $Settings.tempfolder)) {
            Remove-Item -Path ($env:temp + "\" + $Settings.tempfolder) -Force -Recurse -Confirm:$false
            if (Test-Path ($env:temp + "\" + $Settings.tempfolder)) {
                Get-ChildItem -path ($env:temp + "\" + $Settings.tempfolder) | Remove-Item -Force -confirm:$false
            }
        }
        if (-Not (Test-Path ($env:temp + "\" + $Settings.tempfolder))) {
            New-Item -ItemType Directory -Path ($env:temp + "\" + $Settings.tempfolder)
        }
        #set-location -Path ($env:temp + "\" + $Settings.tempfolder)
        #Save Settings
        $SettingsOutput.MachineName = $Settings.Machine_Name.text
        $SettingsOutput.IPAddress = $Settings.IP_Address.Text
        $SettingsOutput.Network_Adapter_List = $Settings.Network_Adapter_List
        $SettingsOutput.Username = $env:USERNAME
        $Settings.AutoLogonRegString = "hkcu:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
        $SettingsOutput.AutoLogonReg = (Get-ItemProperty -path $Settings.AutoLogonRegString)
        #CustomAppReg Reg Export
        if (Test-Path ("HKCU:\SOFTWARE\" + $Settings.CustomAppRegKey)) {
            reg export ("HKCU\SOFTWARE\" + $Settings.CustomAppRegKey) ($env:temp + "\" + $Settings.tempfolder + "\" + $Settings.CustomAppName + "_User.reg") /y
            $SettingsOutput.CustomAppRegUser = Get-ItemProperty -Path ("HKCU:\SOFTWARE\" + $Settings.CustomAppRegKey)
        }    
        if (Test-Path ("HKLM:\SOFTWARE\" + $Settings.CustomAppRegKey)) {
            reg export ("HKLM\SOFTWARE\" + $Settings.CustomAppRegKey) ($env:temp + "\" + $Settings.tempfolder + "\" + $Settings.CustomAppName + "_x86.reg") /y
            $SettingsOutput.CustomAppRegx86 = Get-ItemProperty -Path ("HKLM:\SOFTWARE\" + $Settings.CustomAppRegKey)
        }       
        if (Test-Path ("HKLM:\SOFTWARE\WOW6432Node\" + $Settings.CustomAppRegKey)) {
            reg export ("HKLM\SOFTWARE\WOW6432Node\" + $Settings.CustomAppRegKey) ($env:temp + "\" + $Settings.tempfolder + "\" + $Settings.CustomAppName + "x64.reg") /y
            $SettingsOutput.CustomAppRegx64 = Get-ItemProperty -Path ("HKLM:\SOFTWARE\WOW6432Node\" + $Settings.CustomAppRegKey)
        }
        #Printers
        #$SettingsOutput.Printers = (Get-WMIObject -Class Win32_Printer)
        #$SettingsOutput.PrinterPorts = (Get-WmiObject win32_tcpipprinterport)
        $SettingsOutput.Printers = @()
        ForEach ($Printer in (Get-WmiObject Win32_Printer | Select-Object *)) {
            $Port = $null
            $Port = Get-WmiObject win32_tcpipprinterport -Filter ('name = "' + $Printer.PortName + '"') -ErrorAction SilentlyContinue
            $SettingsOutput.Printers += New-Object -TypeName psobject -Property @{      
                Printer_Name = $Printer.Name;
                Printer_Sharename = $Printer.ShareName ;
                Printer_SystemName = $Printer.SystemName;
                Printer_ServerName = $Printer.ServerName;
                Printer_DriverName = $Printer.DriverName;
                Printer_Default = $Printer.Default;
                Printer_Attributes = $Printer.Attributes;
                Printer_PrintProcessor = $Printer.PrintProcessor;
                Printer_Port_Name = $Printer.PortName;
                Printer_Port_Type = $Port.Description ;
                Printer_Port_Protocol = $Port.Protocol ;
                Printer_Port_IP  = $Port.HostAddress ;
                Printer_Port_Port = $Port.PortNumber ;
                Printer_Port_Queue = $Port.Queue ;
                Printer_Port_SNMPCommunity = $Port.SNMPCommunity ;
                Printer_Port_SNMPEnabled = $Port.SNMPEnabled ;
            }
        }
		#Installed Programs
        If (Test-Path("HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall")) {
            $SettingsOutput.Programs = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* )
        }
        If (Test-Path("HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall")) {
            $SettingsOutput.Programs += (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* )
        }
        #Backup files
        ForEach ($Backup in $Settings.BackupFolders) {
            If ($Backup.Substring(0,2) -ne "\\") {
                $CFN = Split-Path -Leaf $Backup
                #Write-Host ("Backing up: " + $CFN)
                New-Item -ItemType Directory -Path ($env:temp + "\" + $Settings.tempfolder + "\" + $CFN)
                If ($CFN -eq $Settings.CustomAppName) {
                    If ($Settings.CABackup.Checked) {
                        robocopy /e $Backup ($env:temp + "\" + $Settings.tempfolder + "\" + $CFN) /w:3 /r:3 /XD *Image* /XD Export /XD Logs /XD Reports /XD test /XD ENV /XD Update* /XF ( $Settings.CustomAppName + " release notes*.pdf") /XF *.BAK /XF *.log /XF *text*.txt /XF *temp*.* /XF _* /XF ~* /XF Thumbs.db | Out-Null
                    }
                } else {
                    If ($Settings.UserFilesBackup.Checked) {
                        robocopy /e $Backup ($env:temp + "\" + $Settings.tempfolder + "\" + $CFN) /w:3 /r:3 /XF ~* /XF Thumbs.db | Out-Null
                    }
                }
            }
        }
		#Export Settings file
        $SettingsOutput | Export-Clixml -Path ($env:temp + "\" + $Settings.tempfolder + "\settings.xml")

        #Create Archive
        Compress-Archive -Path ($env:temp + "\" + $Settings.tempfolder )  -DestinationPath $Settings.Save_Backup.Text 
        if (Test-Path ($Settings.Save_Backup.Text)) {
            Remove-Item -Path ($env:temp + "\" + $Settings.tempfolder) -Force -Recurse
        }
        
        $Settings.sw.Stop()
        $Settings.Browse.Enabled = $true
        $Settings.Save_Backup.Enabled = $true
        $Settings.IP_Address.Enabled = $true
        $Settings.Machine_Name.Enabled = $true
        $Settings.Start.Enabled = $true
        $Settings.Store_Setup.text = ( $Settings.WindowTitle + " Done. Time: " + (FormatElapsedTime($Settings.sw.Elapsed)) )
        $Settings.Stop.text = "Exit"
     })
     $MainpsCmd.Powershell.Runspace = $MainRunspace
     $MainpsCmd.Handle = $MainpsCmd.Powershell.BeginInvoke()
     #[void]$Settings.Store_Setup.Close()
    #endregion Main thread End
}
function Stop_Work {
    param (
        
    )
    If ($Settings.Stop.text -eq "Exit") {
        [void]$Settings.Store_Setup.Close()
    }Else {
        $MainpsCmd.Stop()
        Start-Sleep -Seconds 5
        if (Test-Path ($Settings.Save_Backup.Text)) {
            Remove-Item -Path ($env:temp + "\" + $Settings.tempfolder) -Force -Recurse
        }
        [void]$Settings.Store_Setup.Close()

        #Exit
    }
}
#############################################################################
#endregion Functions
#############################################################################
#############################################################################
#region Setup Sessions
#############################################################################
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Settings.Store_Setup                     = New-Object system.Windows.Forms.Form
# $Store_Setup.ClientSize          = '400,500'
$Settings.Store_Setup.ClientSize          = '400,200'
$Settings.Store_Setup.text                = $Settings.WindowTitle
$Settings.Store_Setup.TopMost             = $false
#Show Icon https://stackoverflow.com/questions/53376491/powershell-how-to-embed-icon-in-powershell-gui-exe
If ($iconBase64) {
    $iconBytes       = [Convert]::FromBase64String($iconBase64)
    $stream          = New-Object IO.MemoryStream($iconBytes, 0, $iconBytes.Length)
    $stream.Write($iconBytes, 0, $iconBytes.Length);
    $iconImage       = [System.Drawing.Image]::FromStream($stream, $true)
    $Settings.Store_Setup.icon       = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -Argument $stream).GetHIcon())
}
$Settings.Machine_Name_Label              = New-Object system.Windows.Forms.Label
$Settings.Machine_Name_Label.text         = "Machine Name:"
$Settings.Machine_Name_Label.AutoSize     = $true
$Settings.Machine_Name_Label.width        = 25
$Settings.Machine_Name_Label.height       = 10
$Settings.Machine_Name_Label.location     = New-Object System.Drawing.Point(10,10)
$Settings.Machine_Name_Label.Font         = 'Microsoft Sans Serif,10'

$Settings.Machine_Name                    = New-Object system.Windows.Forms.TextBox
$Settings.Machine_Name.multiline          = $false
$Settings.Machine_Name.width              = 180
$Settings.Machine_Name.height             = 20
$Settings.Machine_Name.location           = New-Object System.Drawing.Point(115,10)
$Settings.Machine_Name.Font               = 'Microsoft Sans Serif,10'
$Settings.Machine_Name.text               = $env:computername


$Settings.Network_Adapter_List = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter 'IPEnabled = True' | Where-object {$_.IPaddress -notlike "169.254.*" -and $_.IPAddress -ne "127.0.0.1"} | Select-Object Description,IPAddress,DefaultIPGateway,IPSubnet,DNSServerSearchOrder

$Settings.IP_Address_Label                = New-Object system.Windows.Forms.Label
$Settings.IP_Address_Label.text           = "IP Address:"
$Settings.IP_Address_Label.AutoSize       = $true
$Settings.IP_Address_Label.width          = 25
$Settings.IP_Address_Label.height         = 10
$Settings.IP_Address_Label.location       = New-Object System.Drawing.Point(10,35)
$Settings.IP_Address_Label.Font           = 'Microsoft Sans Serif,10'

$Settings.IP_Address                      = New-Object system.Windows.Forms.TextBox
$Settings.IP_Address.multiline            = $false
$Settings.IP_Address.width                = 180
$Settings.IP_Address.height               = 20
$Settings.IP_Address.location             = New-Object System.Drawing.Point(115,35)
$Settings.IP_Address.Font                 = 'Microsoft Sans Serif,10'
$Settings.IP_Address.Text                 = ($Settings.Network_Adapter_List | Select-Object -first 1).IPAddress

$Settings.FBackup = New-Object System.Windows.Forms.GroupBox #create the group box
$Settings.FBackup.Location = New-Object System.Drawing.Size(10,60) #location of the group box (px) in relation to the primary window's edges (length, height)
$Settings.FBackup.size = New-Object System.Drawing.Size(375,70) #the size in px of the group box (length, height)
$Settings.FBackup.text = "Backup:" #labeling the box


$Settings.CABackup                      = New-Object System.Windows.Forms.Checkbox 
$Settings.CABackup.Text                 = $Settings.CustomAppName
$Settings.CABackup.width                = 180
$Settings.CABackup.height               = 20
# $Settings.CABackup.Location             = New-Object System.Drawing.Size(115,65) 
$Settings.CABackup.Location             = New-Object System.Drawing.Size(10,15) 
$Settings.CABackup.Font                 = 'Microsoft Sans Serif,10'
$Settings.CABackup.Checked              = $true


$Settings.UserFilesBackup                      = New-Object System.Windows.Forms.Checkbox 
$Settings.UserFilesBackup.Text                 = "User Files"
$Settings.UserFilesBackup.width                = 180
$Settings.UserFilesBackup.height               = 20
# $Settings.UserFilesBackup.Location             = New-Object System.Drawing.Size(115,85) 
$Settings.UserFilesBackup.Location             = New-Object System.Drawing.Size(10,40) 
$Settings.UserFilesBackup.Font                 = 'Microsoft Sans Serif,10'
$Settings.UserFilesBackup.Checked              = $true

$Settings.FBackup.Controls.AddRange(@($Settings.CABackup,$Settings.UserFilesBackup)) #activate the inside the group box

$Settings.Save_Backup_Label               = New-Object system.Windows.Forms.Label
$Settings.Save_Backup_Label.text          = "Save Backup"
$Settings.Save_Backup_Label.AutoSize      = $true
$Settings.Save_Backup_Label.width         = 25
$Settings.Save_Backup_Label.height        = 10
$Settings.Save_Backup_Label.location      = New-Object System.Drawing.Point(14,135)
$Settings.Save_Backup_Label.Font          = 'Microsoft Sans Serif,10'

$Settings.Save_Backup                     = New-Object system.Windows.Forms.TextBox
$Settings.Save_Backup.multiline           = $false
$Settings.Save_Backup.width               = 194
$Settings.Save_Backup.height              = 20

$Settings.Save_Backup.location            = New-Object System.Drawing.Point(103,135)
$Settings.Save_Backup.Font                = 'Microsoft Sans Serif,10'

$Settings.Browse                          = New-Object system.Windows.Forms.Button
$Settings.Browse.text                     = "Browse..."
$Settings.Browse.width                    = 70
$Settings.Browse.height                   = 25

$Settings.Browse.location                 = New-Object System.Drawing.Point(304,135)
$Settings.Browse.Font                     = 'Microsoft Sans Serif,10'
If (Test-Path (Split-Path -Parent -Path $MyInvocation.MyCommand.Definition)) {
    $Settings.Save_Backup.Text = ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\" + $Settings.tempfolder + ".zip" )  
}
$Settings.Stop                         = New-Object system.Windows.Forms.Button
$Settings.Stop.text                    = "Exit"
$Settings.Stop.width                   = 70
$Settings.Stop.height                  = 25
$Settings.Stop.location                = New-Object System.Drawing.Point(229,165)
$Settings.Stop.Font                    = 'Microsoft Sans Serif,10'

$Settings.Start                         = New-Object system.Windows.Forms.Button
$Settings.Start.text                    = "Start"
$Settings.Start.width                   = 70
$Settings.Start.height                  = 25
$Settings.Start.location                = New-Object System.Drawing.Point(304,165)
$Settings.Start.Font                    = 'Microsoft Sans Serif,10'


$Settings.Store_Setup.controls.AddRange(@($Settings.Machine_Name_Label,$Settings.IP_Address_Label,$Settings.Machine_Name,$Settings.IP_Address,$Settings.FBackup,$Settings.Save_Backup,$Settings.Start,$Settings.Stop,$Settings.Save_Backup_Label,$Settings.Browse))

#############################################################################
#endregion Setup Sessions
#############################################################################
#############################################################################
#region Main 
#############################################################################

$Settings.Browse.Add_Click({ Browse_File })
$Settings.Start.Add_Click({ Start_Work })
$Settings.Stop.Add_Click({ Stop_Work })


[void]$Settings.Store_Setup.ShowDialog()
#############################################################################
#endregion Main
#############################################################################
