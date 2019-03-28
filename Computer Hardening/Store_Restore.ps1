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
    1.0.0 - Basic script functioning and can work on Windows 7

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
$Settings =[hashtable]::Synchronized(@{})
# $Settings =@{}
$SettingsOutput =[hashtable]::Synchronized(@{})
# $SettingsOutput =@{}

$Settings.WindowTitle = "Store Restore"
$Settings.tempfolder = ""
$Settings.CustomAppFolder = "CustomApp"
$Settings.CustomAppRegKey = "CustomApp"
$Settings.CustomAppName = "CustomApp"
$Settings.AccountBlacklist = @(
    "Administrator"
    "ASPNET"
    "DefaultAccount"
    "Guest"
    "WDAGUtilityAccount"
)
$Settings.AutoLogonRegString = "hkcu:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
$Settings.AutoLogonReg = (Get-ItemProperty -path $Settings.AutoLogonRegString)
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
$iconBase64 ='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'
#endregion Icon
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
function Set-Reg {
	#Source: https://github.com/nichite/chill-out-windows-10/blob/master/chill-out-windows-10.ps1
	# String: Specifies a null-terminated string. Equivalent to REG_SZ.
	# ExpandString: Specifies a null-terminated string that contains unexpanded references to environment variables that are expanded when the value is retrieved. Equivalent to REG_EXPAND_SZ.
	# Binary: Specifies binary data in any form. Equivalent to REG_BINARY.
	# DWord: Specifies a 32-bit binary number. Equivalent to REG_DWORD.
	# MultiString: Specifies an array of null-terminated strings terminated by two null characters. Equivalent to REG_MULTI_SZ.
	# Qword: Specifies a 64-bit binary number. Equivalent to REG_QWORD.
	# Unknown: Indicates an unsupported registry data type, such as REG_RESOURCE_LIST.
    Param (
        [parameter()][string]$Path,
        [parameter()][string]$Name,
        [parameter()][string]$Value,
        [parameter()][string]$Type
    )	
	If(!(Test-Path $Path)) {
        New-Item -Path $Path -Force | Out-Null
    }
    New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType $Type -Force | Out-Null
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
    $Settings.AutoLogon.Enabled = $false
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
    $Settings.AutoLogon.Enabled = $True
    $Settings.Manager.Enabled = $True
    $Settings.Start.text = "Restore"
}
function Start_Work {
    param (
        
    )
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
    $Settings.AutoLogon.Enabled = $false
    $Settings.Manager.Enabled = $false
    #region Main thread Start
      $MainRunspace =[runspacefactory]::CreateRunspace()      
      $MainRunspace.Open()
      $MainRunspace.SessionStateProxy.SetVariable("Settings",$Settings)  
      $MainRunspace.SessionStateProxy.SetVariable("SettingsOutput",$SettingsOutput)    

      $MainpsCmd = "" | Select-Object PowerShell,Handle
      $MainpsCmd.PowerShell = [PowerShell]::Create().AddScript({ 

        #region CustomAppReg Reg Import
        If (Test-Path (${env:ProgramFiles(x86)} + "\" + $Settings.CustomAppFolder)) {
            $Settings.CustomAppFullPath = (${env:ProgramFiles(x86)} + "\" + $Settings.CustomAppFolder)           
        }
        If (Test-Path (${env:ProgramFiles} + "\" + $Settings.CustomAppFolder)) {
            $Settings.CustomAppFullPath =  (${env:ProgramFiles} + "\" + $Settings.CustomAppFolder)
        } 
        if ($Settings.CustomAppFullPath) {
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
        #endregion CustomAppReg Reg Import  
        #region Autologon
        #http://get-cmd.com/?p=4679
        If (-Not ($Settings.AutoLogon.SelectedIndex) -and ($Settings.AutoLogon.SelectedIndex -ne $Settings.AutoLogonReg.DefaultUserName)) {
                Set-Reg -Path $Settings.AutoLogonRegString -Name "AutoAdminLogon" -Value "1" -Type String
                Set-Reg -Path $Settings.AutoLogonRegString -Name "DefaultUsername" -Value "$Settings.AutoLogon.SelectedIndex.ToString()" -Type String
                Set-Reg -Path $Settings.AutoLogonRegString -Name "DefaultPassword" -Value "$Settings.AutoLogon.SelectedIndex.ToString()" -Type String
            # If ($Settings.AutoLogonReg.DefaultUserName) {
            #     Set-ItemProperty $Settings.AutoLogonRegString "AutoAdminLogon" -Value "1" -type String 
            #     Set-ItemProperty $Settings.AutoLogonRegString "DefaultUsername" -Value "$Settings.AutoLogon.SelectedIndex.ToString()" -type String 
            #     Set-ItemProperty $Settings.AutoLogonRegString "DefaultPassword" -Value "$Settings.AutoLogon.SelectedIndex.ToString()" -type String
            # }else{                   
            #     New-ItemProperty $Settings.AutoLogonRegString "AutoAdminLogon" -Value "1" -type String 
            #     New-ItemProperty $Settings.AutoLogonRegString "DefaultUsername" -Value "$Settings.AutoLogon.SelectedIndex.ToString()" -type String 
            #     New-ItemProperty $Settings.AutoLogonRegString "DefaultPassword" -Value "$Settings.AutoLogon.SelectedIndex.ToString()" -type String
            # }
        }
        Enable-LocalUser -Name ($Settings.AutoLogon.SelectedIndex.ToString()) -Confirm:$false
        #endregion Autologon
        #region Managers 
        If ($Settings.Manager.Checked) {
            #Mounted User Hive Location
            New-PSDrive -PSProvider Registry -Name HKU -Root HKEY_USERS -erroraction 'silentlycontinue' | Out-Null
            $HKEY = ("HKU\H_" + $Settings.AutoLogon.SelectedItem.ToString())
            #Get Hive file location
            $UserProfile = (Get-WmiObject Win32_UserProfile |Where-Object { (Split-Path -leaf -Path ($_.LocalPath)) -eq $Settings.AutoLogon.SelectedItem.ToString()} |Select-Object Localpath).localpath 
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
						write-error ( "Cannot load profile for: " + ($Settings.UsersProfileFolder + "\" + $Settings.AutoLogon.SelectedIndex.ToString() + "\ntuser.dat") )
						continue
					}		
				}else{
					write-error ( "Cannot load profile for: " + ($Settings.UsersProfileFolder + "\" + $Settings.AutoLogon.SelectedIndex.ToString() + "\ntuser.dat") )
					continue
				}
            }
            #region Start Relaxing Setting for Managers #
            #Show all drives in Windows Explorer	
            Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoDrives" 0 "DWORD"
 			#Enable user to using My Computer to gain access to the content of selected drives. 
            Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoViewOnDrive" 0 "DWORD"
            #Enable Context-sensitive menus .
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoTrayContextMenu" 0 "DWORD"
            #Enable right-click on Desktop and Windows Explorer
            Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") "NoViewContextMenu" 0 "DWORD"
            #Enable Context Menus in the Start Menu in Windows 10
			Set-Reg ($HKEY.replace("HKU\","HKU:\") + "\Software\Policies\Microsoft\Windows\Explorer") "DisableContextMenusInStart" 0 "DWORD"

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
        }
        #endregion Managers 
        #region Printers
        If ($SettingsOutput.Printers) {
            $CurrentPrinters = (Get-WmiObject Win32_Printer | Select-Object *)
            $CurrentPrinterPorts = (Get-WmiObject win32_tcpipprinterport | Select-Object *)
            $CurrentPrinterDrivers = (Get-WmiObject Win32_PrinterDriver | Select-Object *)
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
            $CurentDefault = (Get-WmiObject -Query " Select name FROM Win32_Printer WHERE Default=$true").Name
            $OldDefault = ($ImportCVS | Where-Object {$_.Printer_Default -eq $true}).Printer_Name
            If ($CurrentDefault -ne $OldDefault) {
                (New-Object -ComObject WScript.Network).SetDefaultPrinter($OldDefault)
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
                If (Test-Path($env:temp + "\" + $Settings.tempfolder + "\" + $CFN)) {
                    If ($CFN -eq $Settings.CustomAppName) {
                        robocopy /e ($env:temp + "\" + $Settings.tempfolder + "\" + $CFN) $Restore /w:3 /r:3 /XD *Image* /XD Export /XD Logs /XD Reports /XD test /XD ENV /XD Update* /XF ( $Settings.CustomAppName + " release notes*.pdf") /XF *.BAK /XF *.log /XF *text*.txt /XF *temp*.* /XF _* /XF ~* /XF Thumbs.db | Out-Null
                    } else {
                        robocopy /e ($env:temp + "\" + $Settings.tempfolder + "\" + $CFN) $Restore /w:3 /r:3 /XF ~* /XF Thumbs.db | Out-Null
                    }
                }
            }
        }
        #endregion Restore files
        #region Set Machine IP
        If ($Settings.IP_Address.Text) {
            $Network = $SettingsOutput.Network_Adapter_List | Where-Object {$_.IPAddress -eq $Settings.IP_Address.Text}
            If ($Network) {
                $wmi = Get-WmiObject win32_networkadapterconfiguration -filter ("Description = '" + $Settings.Network_Adapter.SelectedItem.ToString() + "'")
                $wmi.EnableStatic($Network.IPAddress, $Network.IPSubnet)              
                $wmi.SetGateways($Network.DefaultIPGateway, 1)        
                $wmi.SetDNSServerSearchOrder($Network.DNSServerSearchOrder)
            }
        }
        #endregion Set Machine IP
        #region Set Machine Name
        If ($Settings.Machine_Name.Text) {
            Rename-computer –NewName $Settings.Machine_Name.Text  –force 
        }
        #endregion Set Machine Name
        
        $Settings.sw.Stop()
        $Settings.Store_Setup.text = ( $Settings.WindowTitle + " Done. Time: " + (FormatElapsedTime($Settings.sw.Elapsed)) )
        
        #Stop Auto Launch 
        $TempLnk = ($env:appdata + "\Microsoft\Windows\Start Menu\Programs\Startup\" + [io.path]::GetFileNameWithoutExtension($myinvocation.mycommand.definition) + ".lnk")
        If(Test-Path($TempLnk)) {
            Remove-Item -Force -Path $TempLnk
        }
        
      })
      $MainpsCmd.Powershell.Runspace = $MainRunspace
      $MainpsCmd.Handle = $MainpsCmd.Powershell.BeginInvoke()

     #[void]$Settings.Store_Setup.Close()
    #endregion Main thread End
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
Add-Type -AssemblyName System.Windows.Forms
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
    $iconImage       = [System.Drawing.Image]::FromStream($stream, $true)
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
#$Settings.Machine_Name.text               = $env:computername


$Settings.Network_Adapter_List = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter 'IPEnabled = True' | Where-object { $_.IPAddress -ne "127.0.0.1"} | Select-Object InterfaceAlias,IPAddress,Description

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
#$Settings.IP_Address.Text                 = ($Settings.Network_Adapter_List | Select-Object -first 1).IPAddress
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


$Settings.AutoLogon_Label                = New-Object system.Windows.Forms.Label
$Settings.AutoLogon_Label.text           = "Auto Logon User:"
$Settings.AutoLogon_Label.AutoSize       = $true
$Settings.AutoLogon_Label.width          = 25
$Settings.AutoLogon_Label.height         = 10
$Settings.AutoLogon_Label.location       = New-Object System.Drawing.Point(10,125)
$Settings.AutoLogon_Label.Font           = 'Microsoft Sans Serif,10'

$Settings.AutoLogon                       = New-Object system.Windows.Forms.ComboBox
$Settings.AutoLogon.text                  = " "
$Settings.AutoLogon.width                 = 265
$Settings.AutoLogon.height                = 20
$Settings.AutoLogon.location              = New-Object System.Drawing.Point(125,125)
$Settings.AutoLogon.Font                  = 'Microsoft Sans Serif,10'


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

$Settings.Store_Setup.controls.AddRange(@($Settings.Machine_Name_Label,$Settings.IP_Address_Label,$Settings.Machine_Name,$Settings.IP_Address,$Settings.Network_Adapter_Label,$Settings.Network_Adapter,$Settings.Restore_Backup,$Settings.Start,$Settings.Stop,$Settings.Restore_Backup_Label,$Settings.Browse,$Settings.AutoLogon_Label,$Settings.AutoLogon,$Settings.Manager,$Settings.FBackup))


#############################################################################
#endregion Setup Sessions
#############################################################################
#############################################################################
#region Main 
#############################################################################

$Settings.Browse.Add_Click({ Browse_File })
$Settings.Start.Add_Click({ Start_Work })
$Settings.Stop.Add_Click({ Stop_Work })
#$Settings.AutoLogon.Add_SelectedIndexChanged({  })

ForEach ( $LocalUser in ((Get-LocalUser).name | Sort-Object {"$_" -replace '\d',''},{("$_" -replace '\D','') -as [int]})) {
    If (-Not ($Settings.AccountBlacklist.contains($LocalUser))) {
        $Settings.AutoLogon.Items.Add($LocalUser)
    }
}

If ($Settings.AutoLogonReg.DefaultUserName) {
    $Settings.AutoLogon.SelectedItem = $Settings.AutoLogonReg.DefaultUserName
}else {
    # If ($Settings.AutoLogon.SelectionLength -ge 0) {
    #     $Settings.AutoLogon.SelectedIndex = 0
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
