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
    1.0.0 - Basic script functioning and can work on Windows 7

#>
#Force Starting of Powershell script as Administrator 
If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))

{   
$arguments = "& '" + $myinvocation.mycommand.definition + "'"
Start-Process powershell -Verb runAs -ArgumentList $arguments
Break
}

#############################################################################
#region User Variables
#############################################################################
#$Settings =[hashtable]::Synchronized(@{})
$Settings =@{}
#$SettingsOutput =[hashtable]::Synchronized(@{})
$SettingsOutput =@{}

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
    $Settings.SaveFileDialog.initialDirectory = (Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) 
    $Settings.SaveFileDialog.filter = "ZIP Archive Files|*.zip|All Files|*.*" 
    $Settings.SaveFileDialog.ShowDialog() | Out-Null
    $Settings.Save_Backup.Text = $Settings.SaveFileDialog.filename
}
function Start_Work {
    param (
        
    )
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
        $SettingsOutput.Printers = (Get-WMIObject -Class Win32_Printer)
        $SettingsOutput.PrinterPorts = (Get-WmiObject win32_tcpipprinterport)
        #Backup files
        ForEach ($Backup in $Settings.BackupFolders) {
            $CFN = Split-Path -Leaf $Backup
            #Write-Host ("Backing up: " + $CFN)
            New-Item -ItemType Directory -Path ($env:temp + "\" + $Settings.tempfolder + "\" + $CFN)
            If ($CFN -eq $Settings.CustomAppName) {
                robocopy /e $Backup ($env:temp + "\" + $Settings.tempfolder + "\" + $CFN) /w:3 /r:3 /XD *Image* /XD Export /XD Logs /XD Reports /XD test /XD ENV /XD Update* /XF ( $Settings.CustomAppName + " release notes*.pdf") /XF *.BAK /XF *.log /XF *text*.txt /XF *temp*.* /XF _* /XF ~* /XF Thumbs.db | Out-Null
            } else {
                robocopy /e $Backup ($env:temp + "\" + $Settings.tempfolder + "\" + $CFN) /w:3 /r:3 /XF ~* /XF Thumbs.db | Out-Null
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
        $Settings.Store_Setup.text                = ( $Settings.WindowTitle + " Done. Time: " + (FormatElapsedTime($Settings.sw.Elapsed)) )
    
        $Settings.Browse.Enabled = $true
        $Settings.Save_Backup.Enabled = $true
        $Settings.IP_Address.Enabled = $true
        $Settings.Machine_Name.Enabled = $true
        $Settings.Start.Enabled = $true
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
    $MainpsCmd.Stop()
    Start-Sleep -Seconds 5
    if (Test-Path ($Settings.Save_Backup.Text)) {
        Remove-Item -Path ($env:temp + "\" + $Settings.tempfolder) -Force -Recurse
    }
    [void]$Settings.Store_Setup.Close()

    #Exit
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
if (Test-Path ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\2Comps.ico" ) ) {
    $Settings.Store_Setup.icon                = ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\2Comps.ico" )
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


$Settings.Network_Adapter_List = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter 'IPEnabled = True' | Where-object {$_.IPaddress -notlike "169.254.*" -and $_.IPAddress -ne "127.0.0.1"} | Select-Object InterfaceAlias,IPAddress

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

$Settings.Network_Adapter                 = New-Object system.Windows.Forms.Button
$Settings.Network_Adapter.text            = "Adapter"
$Settings.Network_Adapter.width           = 70
$Settings.Network_Adapter.height          = 25
$Settings.Network_Adapter.location        = New-Object System.Drawing.Point(304,35)
$Settings.Network_Adapter.Font            = 'Microsoft Sans Serif,10'
$Settings.Network_Adapter.Hide()


$Settings.Save_Backup_Label               = New-Object system.Windows.Forms.Label
$Settings.Save_Backup_Label.text          = "Save Backup"
$Settings.Save_Backup_Label.AutoSize      = $true
$Settings.Save_Backup_Label.width         = 25
$Settings.Save_Backup_Label.height        = 10
$Settings.Save_Backup_Label.location      = New-Object System.Drawing.Point(14,125)
$Settings.Save_Backup_Label.Font          = 'Microsoft Sans Serif,10'

$Settings.Save_Backup                     = New-Object system.Windows.Forms.TextBox
$Settings.Save_Backup.multiline           = $false
$Settings.Save_Backup.width               = 194
$Settings.Save_Backup.height              = 20

$Settings.Save_Backup.location            = New-Object System.Drawing.Point(103,125)
$Settings.Save_Backup.Font                = 'Microsoft Sans Serif,10'

$Settings.Browse                          = New-Object system.Windows.Forms.Button
$Settings.Browse.text                     = "Browse..."
$Settings.Browse.width                    = 70
$Settings.Browse.height                   = 25

$Settings.Browse.location                 = New-Object System.Drawing.Point(304,125)
$Settings.Browse.Font                     = 'Microsoft Sans Serif,10'
If (Test-Path (Split-Path -Parent -Path $MyInvocation.MyCommand.Definition)) {
    $Settings.Save_Backup.Text = ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\" + $Settings.tempfolder + ".zip" )  
}
$Settings.Stop                         = New-Object system.Windows.Forms.Button
$Settings.Stop.text                    = "Stop"
$Settings.Stop.width                   = 70
$Settings.Stop.height                  = 25

$Settings.Stop.location                = New-Object System.Drawing.Point(229,150)
$Settings.Stop.Font                    = 'Microsoft Sans Serif,10'

$Settings.Start                         = New-Object system.Windows.Forms.Button
$Settings.Start.text                    = "Start"
$Settings.Start.width                   = 70
$Settings.Start.height                  = 25

$Settings.Start.location                = New-Object System.Drawing.Point(304,150)
$Settings.Start.Font                    = 'Microsoft Sans Serif,10'


$Settings.Store_Setup.controls.AddRange(@($Settings.Machine_Name_Label,$Settings.IP_Address_Label,$Settings.Machine_Name,$Settings.IP_Address,$Settings.Network_Adapter,$Settings.Save_Backup,$Settings.Start,$Settings.Stop,$Settings.Save_Backup_Label,$Settings.Browse))

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
