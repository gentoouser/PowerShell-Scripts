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
    1.0.0 - Basic script functioning and can work on Windows 10

#>
#Force Starting of Powershell script as Administrator 
# If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))

# {   
# $arguments = "& '" + $myinvocation.mycommand.definition + "'"
# Start-Process powershell -Verb runAs -ArgumentList $arguments
# Break
# }

#############################################################################
#region User Variables
#############################################################################
$Settings =[hashtable]::Synchronized(@{})
# $Settings =@{}
$SettingsOutput =[hashtable]::Synchronized(@{})
# $SettingsOutput =@{}

$Settings.WindowTitle = " Store Restore"
$Settings.tempfolder = ""
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
    $Settings.Store_Setup.text                = ( $Settings.WindowTitle + " Working . . . Please wait.")
    $Settings.Browse.Enabled = $false


    $Settings.OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    #$Settings.OpenFileDialog.initialDirectory = (Split-Path -Path $MyInvocation.MyCommand.Definition) 
    $Settings.OpenFileDialog.filter = "ZIP Archive Files|*.zip|All Files|*.*" 
    $Settings.OpenFileDialog.ShowDialog() | Out-Null
    $Settings.Restore_Backup.Text = $Settings.OpenFileDialog.filename    

    $Settings.tempfolder = ($env:temp + "\" + [io.path]::GetFileNameWithoutExtension($Settings.OpenFileDialog.filename))

      $BrowseRunspace =[runspacefactory]::CreateRunspace()      
      $BrowseRunspace.Open()
      $BrowseRunspace.SessionStateProxy.SetVariable("Settings",$Settings)  
      $BrowseRunspace.SessionStateProxy.SetVariable("SettingsOutput",$SettingsOutput)    

      $BrowsepsCmd = "" | Select-Object PowerShell,Handle
      $BrowsepsCmd.PowerShell = [PowerShell]::Create().AddScript({ 

        Expand-Archive -Path $Settings.Restore_Backup.Text -DestinationPath $env:temp  -Force



     })
     $BrowsepsCmd.Powershell.Runspace = $BrowseRunspace
     $BrowsepsCmd.Handle = $BrowsepsCmd.Powershell.BeginInvoke()

     If (Test-Path ($Settings.tempfolder + "\settings.xml")) {
        $SettingsOutput = Import-Clixml -Path ($Settings.tempfolder + "\settings.xml")
    }
        #region Printer_Group

        #https://stackoverflow.com/questions/32278589/how-do-i-dynamically-create-check-boxes-from-an-array-using-the-form
        # Keep track of number of checkboxes
        $CheckBoxCounter = 1

        # When we create a new textbox, we add it to an array for easy reference later
        $Settings.Printers = @()
        foreach($Label in ($SettingsOutput.Printers).name) {
            $CheckBox = New-Object System.Windows.Forms.CheckBox        
            $CheckBox.UseVisualStyleBackColor = $True
            $System_Drawing_Size = New-Object System.Drawing.Size
            $System_Drawing_Size.Width = 300
            $System_Drawing_Size.Height = 24
            $CheckBox.Size = $System_Drawing_Size
            $CheckBox.TabIndex = 2

            # Assign text based on the input
            $CheckBox.Text = $Label

            $System_Drawing_Point = New-Object System.Drawing.Point
            $System_Drawing_Point.X = 27
            # Make sure to vertically space them dynamically, counter comes in handy
            $System_Drawing_Point.Y = 13 + (($CheckBoxCounter - 1) * 21)
            $CheckBox.Location = $System_Drawing_Point
            $CheckBox.DataBindings.DefaultDataSourceUpdateMode = 0

            # Give it a unique name based on our counter
            $CheckBox.Name = "CheckBox$CheckBoxCounter"
            $Settings.Printers."CheckBox$CheckBoxCounter" = $CheckBox
            # Add it to the form
            $Settings.Printer_Group.controls.Add($Settings.$CheckBox)
            # return object ref to array
            $CheckBox
            # increment our counter
            $CheckBoxCounter++
        }
        #endregion Printer_Group

        $Settings.Machine_Name.text               = $SettingsOutput.MachineName
        $Settings.IP_Address.Text                 = $SettingsOutput.IPAddress

    $Settings.Store_Setup.text                = ( $Settings.WindowTitle)
    $Settings.Browse.Enabled = $true
    $Settings.Restore_Backup.Enabled = $true
    $Settings.IP_Address.Enabled = $true
    $Settings.Machine_Name.Enabled = $true
    $Settings.Start.Enabled = $true
}
function Start_Work {
    param (
        
    )
    $Settings.sw = [Diagnostics.Stopwatch]::StartNew()
    $Settings.Store_Setup.text                = ( $Settings.WindowTitle + " Working . . . Please wait.")
    $Settings.Browse.Enabled = $false
    $Settings.Restore_Backup.Enabled = $false
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

      
        #CustomAppReg Reg Import
        If (Test-Path (${env:ProgramFiles(x86)} + "\" + $Settings.CustomAppFolder)) {
            $Settings.CustomAppFullPath = (${env:ProgramFiles(x86)} + "\" + $Settings.CustomAppFolder)           
        }
        If (Test-Path (${env:ProgramFiles} + "\" + $Settings.CustomAppFolder)) {
            $Settings.CustomAppFullPath =  (${env:ProgramFiles} + "\" + $Settings.CustomAppFolder)
        } 
        if ($SettingsOutput.CustomAppRegUser) {
            reg import ($env:temp + "\" + $Settings.tempfolder + "\" + $Settings.CustomAppName + "_User.reg") /y
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

        #Printers

        #Restore files
        ForEach ($Restore in $Settings.BackupFolders) {
            $CFN = Split-Path -Leaf $Backup
            #Write-Host ("Backing up: " + $CFN)
            New-Item -ItemType Directory -Path ($env:temp + "\" + $Settings.tempfolder + "\" + $CFN)
            If ($CFN -eq $Settings.CustomAppName) {
                robocopy /e $Backup ($env:temp + "\" + $Settings.tempfolder + "\" + $CFN) /w:3 /r:3 /XD *Image* /XD Export /XD Logs /XD Reports /XD test /XD ENV /XD Update* /XF ( $Settings.CustomAppName + " release notes*.pdf") /XF *.BAK /XF *.log /XF *text*.txt /XF *temp*.* /XF _* /XF ~* /XF Thumbs.db | Out-Null
            } else {
                robocopy /e $Backup ($env:temp + "\" + $Settings.tempfolder + "\" + $CFN) /w:3 /r:3 /XF ~* /XF Thumbs.db | Out-Null
            }
        }


        
        $Settings.sw.Stop()
        $Settings.Store_Setup.text                = ( $Settings.WindowTitle + " Done. Time: " + (FormatElapsedTime($Settings.sw.Elapsed)) )
    
        $Settings.Browse.Enabled = $true
        $Settings.Restore_Backup.Enabled = $true
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
$Settings.Store_Setup.ClientSize          = '400,600'
$Settings.Store_Setup.text                = $Settings.WindowTitle
$Settings.Store_Setup.TopMost             = $false
if (Test-Path ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\2Comps.ico" ) ) {
    $Settings.Store_Setup.icon                = ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\2Comps.ico" )
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

$Settings.Restore_Backup.location            = New-Object System.Drawing.Point(120,10)
$Settings.Restore_Backup.Font                = 'Microsoft Sans Serif,10'

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
$Settings.Machine_Name.location           = New-Object System.Drawing.Point(115,40)
$Settings.Machine_Name.Font               = 'Microsoft Sans Serif,10'
$Settings.Machine_Name.Enabled            = $false
#$Settings.Machine_Name.text               = $env:computername


$Settings.Network_Adapter_List = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter 'IPEnabled = True' | Where-object { $_.IPAddress -ne "127.0.0.1"} | Select-Object InterfaceAlias,IPAddress

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
$Settings.IP_Address.location             = New-Object System.Drawing.Point(115,65)
$Settings.IP_Address.Font                 = 'Microsoft Sans Serif,10'
#$Settings.IP_Address.Text                 = ($Settings.Network_Adapter_List | Select-Object -first 1).IPAddress
$Settings.IP_Address.Enabled              = $false

$Settings.Network_Adapter                 = New-Object system.Windows.Forms.Button
$Settings.Network_Adapter.text            = "Adapter"
$Settings.Network_Adapter.width           = 70
$Settings.Network_Adapter.height          = 25
$Settings.Network_Adapter.location        = New-Object System.Drawing.Point(304,65)
$Settings.Network_Adapter.Font            = 'Microsoft Sans Serif,10'
#$Settings.Network_Adapter.Hide()
$Settings.Network_Adapter.Enabled         = $false

$Settings.Printer_Group                   = New-Object system.Windows.Forms.Groupbox
$Settings.Printer_Group.height            = 360
$Settings.Printer_Group.width             = 380
$Settings.Printer_Group.text              = "Printers"
$Settings.Printer_Group.location          = New-Object System.Drawing.Point(10,100)


$Settings.Stop                         = New-Object system.Windows.Forms.Button
$Settings.Stop.text                    = "Stop"
$Settings.Stop.width                   = 70
$Settings.Stop.height                  = 25

$Settings.Stop.location                = New-Object System.Drawing.Point(229,565)
$Settings.Stop.Font                    = 'Microsoft Sans Serif,10'

$Settings.Start                         = New-Object system.Windows.Forms.Button
$Settings.Start.text                    = "Start"
$Settings.Start.width                   = 70
$Settings.Start.height                  = 25

$Settings.Start.location                = New-Object System.Drawing.Point(304,565)
$Settings.Start.Font                    = 'Microsoft Sans Serif,10'
$Settings.Start.Enabled                 = $false

$Settings.Store_Setup.controls.AddRange(@($Settings.Machine_Name_Label,$Settings.IP_Address_Label,$Settings.Machine_Name,$Settings.IP_Address,$Settings.Network_Adapter,$Settings.Printer_Group,$Settings.Restore_Backup,$Settings.Start,$Settings.Stop,$Settings.Restore_Backup_Label,$Settings.Browse))


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
