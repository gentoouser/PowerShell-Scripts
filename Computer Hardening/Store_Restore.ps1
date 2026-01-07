<# 
.SYNOPSIS
    Name: Store_Restore.ps1
    Import Store Settings to Zip file, Changes hostname, and IP address and enables windows account.

.DESCRIPTION
    Import Store Settings to Zip file, Changes hostname, and IP address and enables windows account.
    This script will store the settings of the machine to a zip file. This will include the following:
    - Desktop
    - Favorites
    - My Pictures
    - customapp settings
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
    1.0.04 - Added Auto Update for BIOS and Auto Install for LanDesk.
    1.0.05 - Removed WindowLogonUser and Added logic to update logon based on machine name.
    1.0.06 - Updated Passwords based on PCI 2019 Store standard.
    1.0.07 - Fix IP update and computer rename.
    1.0.08 - Create StaffingModel folder.
    1.0.09 - Fix bug with disabling all accounts. Added prompt about disabling Auto Logon. Fixed issues with setting Manager settings. Fixed issues with not updating admin autologon to new password. 
    1.0.10 - Updated to deal with zip file from powershell 2.0
    1.0.11 - Manager's registry keys bug. Get-CimInstance testing. Start Office 2019 install. 
    1.0.12 - Machine factory name password issue. Fixed issue with renaming machine.
    1.0.13 - Disable RDP after LanDesk agent is installed.
    1.0.14 - Fixed issue with the Adapter selection list.
    1.0.15 - Add Logic to install newer Ivanti EPM agent.
    1.0.16 - Remove Old Agent
    1.0.17 - Add Logic for RDM Installer
    1.0.18 - 20250424 - Look at for Ivanti services to see if EPM agent is installed.
    1.0.19 - 20250603 - Disable reboot prompt.
    1.0.20 - 20251021 - Fixed issue with Ivanti EPM installation detection.
    1.0.21 - 20251210 - Major re-work to set IP Address. Added monitoring of EPM Agent install. 
    1.0.22 - 20251216 - Readded prompt to install EPM Agent.
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
$Settings.Version = "1.0.21"
$Settings.WindowTitle = ("Store Restore Version: " + $Settings.Version)
$Settings.tempfolder = ""
$Settings.EPMFolderItemCount = 25
$Settings.EPMAgentPath = "${env:ProgramFiles(x86)}\Ivanti\EPM Agent"
$Settings.CustomAppFolder = "Github\customapp"
$Settings.CustomAppRegKey = "Github\customapp"
$Settings.CustomAppName = "customapp"
$Settings.DNS = @("1.1.1.1","8.8.8.8")
$Settings.Subnet = "255.255.255.128"
$Settings.OfficeSubFolder = ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\NEW PC\Microsoft Office 2019 x64")
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
$Settings.UPrePass = ""
$Settings.UPostPass = ""
$Settings.DefaultState = ""
$Settings.APrePass = ""
$Settings.APostPass = ""
$Settings.WindowLogonUserRegString = "hkcu:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
$Settings.WindowLogonUserReg = (Get-ItemProperty -path $Settings.WindowLogonUserRegString)
$Settings.USF = (Get-ItemProperty -path "hkcu:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders")
$Settings.UsersProfileFolder = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\' -Name "ProfilesDirectory").ProfilesDirectory
$Settings.RDMEPMURL = "https://localwebserver/repository/Packages/RDM_AIO_Install/"

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
$iconBase64 = ""
#endregion Icon
#############################################################################
#endregion User Variables
#############################################################################
#############################################################################
#region Functions
#############################################################################
function Convert-PrefixLengthToSubnetMask {
    param (
    [Parameter(Mandatory = $true)]
    [ValidateRange(0, 32)]
    [int]$PrefixLength
    )
    # Calculate the subnet mask
    $mask = ([math]::Pow(2, $PrefixLength) - 1) * [math]::Pow(2, (32 - $PrefixLength))
    $bytes = [BitConverter]::GetBytes([UInt32]$mask)
    (($bytes.Count - 1)..0 | ForEach-Object { [String]$bytes[$_] }) -join "."
}
function Format-ElapsedTime($ts) {
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
	<# 
	.SYNOPSIS
	Set-Reg is a function to set a registry key and value.

	.DESCRIPTION

	.PARAMETER regPath
	The path to the registry key.
	.PARAMETER name
	The name of the registry value.
	.PARAMETER value
	The data for the registry value.
	.PARAMETER type
	The type of the registry value.
	Valid values are:
	 String: Specifies a null-terminated string. Equivalent to REG_SZ.
	 ExpandString: Specifies a null-terminated string that contains unexpanded references to environment variables that are expanded when the value is retrieved. Equivalent to REG_EXPAND_SZ.
	 Binary: Specifies binary data in any form. Equivalent to REG_BINARY.
	 DWord: Specifies a 32-bit binary number. Equivalent to REG_DWORD.
	 MultiString: Specifies an array of null-terminated strings terminated by two null characters. Equivalent to REG_MULTI_SZ.
	 Qword: Specifies a 64-bit binary number. Equivalent to REG_QWORD.
	 Unknown: Indicates an unsupported registry data type, such as REG_RESOURCE_LIST.
	.PARAMETER comment
	A comment for the registry value.

	.EXAMPLE
	Set-Reg -regPath "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" -name "Test" -value "C:\test.exe" -type "String" -comment "This is a

	.NOTES
	Source: https://github.com/nichite/chill-out-windows-10/blob/master/chill-out-windows-10.ps1
	Modifier: Paul Fuller 
	Changes:
		* Version 1.01.00 - Fix $RegBackup null issue and start to track changes.

	#>

	[CmdletBinding()] 
	Param 
	( 
		[Parameter(Mandatory=$true,Position=1,HelpMessage="Path to Registry Key")][string]$regPath, 
		[Parameter(Mandatory=$true,Position=2,HelpMessage="Name of Value")][string]$name,
		[Parameter(Mandatory=$true,Position=3,HelpMessage="Data for Value")]$value,
		[Parameter(Mandatory=$true,Position=4,HelpMessage="Type of Value")][ValidateSet("String", "ExpandString","Binary","DWord","MultiString","Qword","Unknown",IgnoreCase =$true)][string]$type ,
		[Parameter(Mandatory=$false,Position=5,HelpMessage="Comment value")]$comment
	) 
	$key=$null
	Class BackupRegistry{
		[String]$Path
		[String]$Name
		[String]$Value
		[String]$Type
		[String]$Comment
	}
	$key = $null
	$regvalue = $null
	$regname = $null
	If(Test-Path $regPath -ErrorAction SilentlyContinue) {
		$key = Get-Item -Path $regPath
	}Else{
		New-Item -Path $regPath -Force | Out-Null
		$key = Get-Item -Path $regPath
	}
	If($type -eq "Binary" -and $value.GetType().Name -eq "String" -and $value -match ",") {
		$value = [byte[]]($value -split ",")
	}
	If ($key.Property.Equals($Name)){
		If($key.GetValue($Name) -eq $Value) {
			Write-Verbose ("`Same:" + $regPath + "\" + $name + " = " + $value)
		}Else {
			$BackupReg = [BackupRegistry]::new()
			$BackupReg.Path = $regPath
			$BackupReg.Name = $name
			$BackupReg.Value = $value
			$BackupReg.Type = $type
			$BackupReg.Comment = $comment
			If($null -eq $value){
				Write-Verbose ("`Creating:" + $regPath + "\" + $name + " = " + $value)
			}Else {
				Write-Verbose ("`Updating:" + $regPath + "\" + $name + " = " + $value)
			}
			$Script:RegBackup.Add($BackupReg) | out-null
			New-ItemProperty -Path $regPath -Name $name -Value $value -PropertyType $type -Force | Out-Null
		}
	}Else{
		$BackupReg = [BackupRegistry]::new()
		$BackupReg.Path = $regPath
		$BackupReg.Name = $name
		$BackupReg.Value = $value
		$BackupReg.Type = $type
		$BackupReg.Comment = $comment
		If($null -eq $value){
			Write-Verbose ("`Creating:" + $regPath + "\" + $name + " = " + $value)
		}Else {
			Write-Verbose ("`Updating:" + $regPath + "\" + $name + " = " + $value)
		}
		if (Get-Variable Regbackup -ErrorAction SilentlyContinue){
			$Script:RegBackup.Add($BackupReg) | out-null
		}Else{
			$Script:RegBackup = New-Object System.Collections.ArrayList
			$Script:RegBackup.Add($BackupReg) | out-null
		}

		New-ItemProperty -Path $regPath -Name $name -Value $value -PropertyType $type -Force | Out-Null
	}
}
function Show-Console {
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
function Hide-Console {
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
Function Copy-WebFolder {
    <#
    .LINK
    https://stackoverflow.com/questions/11436694/how-to-download-a-whole-folder-of-files-subfolders-from-the-web-in-powershell

    .SYNOPSIS
     This function copies a folder (and optionally, its subfolders)
    .NOTES
    When copying subfolders it calls itself recursively
    .COMPONENT
    Requires WebClient object $webClient defined, e.g. $webClient = New-Object System.Net.WebClient
    .PARAMETER source
        The url of folder to copy, with trailing /, e.g. http://website/folder/structure/
    .PARAMETER destination
        The folder to copy $source to, with trailing \ e.g. D:\CopyOfStructure\
    .PARAMETER recursive
        True if subfolders of $source are also to be copied or False to ignore subfolders

    #>
    [CmdletBinding()] 
    Param 
    ( 
        [Parameter(Mandatory=$true,Position=1,HelpMessage="URL")][string]$source, 
        [Parameter(Mandatory=$true,Position=2,HelpMessage="Destination")][string]$destination,
        [Parameter(Mandatory=$true,Position=3,HelpMessage="Recursive")][switch]$recursive 
    ) 
    if (!$(Test-Path($destination))) {
        New-Item $destination -type directory -Force
    }
    If ($destination -notmatch "\\$") {
        $destination = $destination + "\"
    }   
    If ($source -notmatch "/$") {
        $source = $source + "/"
    }
    # Create a new WebClient object to download the files
    $webClient = New-Object System.Net.WebClient
    # Get the file list from the web page
    $webString = $webClient.DownloadString($source)
    $lines = [Regex]::Split($webString, "<br>")
    # Parse each line, looking for files and folders
    foreach ($line in $lines) {
        if ($line.ToUpper().Contains("HREF")) {
            # File or Folder
            if (!$line.ToUpper().Contains("[TO PARENT DIRECTORY]")) {
                # Not Parent Folder entry
                $items =[Regex]::Split($line, """")
                $items = [Regex]::Split($items[2], "(>|<)")
                $item = $items[2]
                if ($line.ToLower().Contains("&lt;dir&gt")) {
                    # Folder
                    if ($recursive) {
                        # Subfolder copy required
                        Copy-WebFolder "$source$item/" "$destination$item/" $recursive
                    } else {
                        # Subfolder copy not required
                    }
                } else {
                    # File
                    $webClient.DownloadFile("$source$item", "$destination$item")
                }
            }
        }
    }
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
            #region Disable and update password for user
                #Update Password using new computer name
                [boolean]$SetPassAdmin = $false
                If ($Settings.Machine_Name.Text) {
                    $arrComName = $Settings.Machine_Name.Text -split "-"
                    #State
                    #Deal with Dev boxes.
                    Try{
                        $State = $arrComName[0].ToUpper()
                        If ($State.Length -ne 2) {
                            $UPrePass = [System.Text.Encoding]::Unicode.GetString([System.Convert]::FromBase64String($Settings.UPrePass))
                            $UPostPass = [System.Text.Encoding]::Unicode.GetString([System.Convert]::FromBase64String($Settings.UPostPass))
                            $State = [System.Text.Encoding]::Unicode.GetString([System.Convert]::FromBase64String($Settings.DefaultState))
                        }
                    }
                    Catch {
                        $State = [System.Text.Encoding]::Unicode.GetString([System.Convert]::FromBase64String($Settings.DefaultState))
                    }
                    #Store Number
                    Try{
                        $Store = $arrComName[1].ToLower()
                    }
                    Catch {
                        $Store = $null
                    }
                    #Get Window
                    Try{
                        $Window = $arrComName[2] -replace  "\D+" 
                        If (-Not $Window) {
                            $Window = 01
                        }
                    }
                    Catch {
                        $Window = 01
                    }
                    #Get Window Type
                    Try{
                        $WindowType = ($arrComName[2] -replace '[0-9]').ToLower()
                        If ($WindowType.Length -ne 1) {
                            $WindowType = "w"
                        }
                    }
                    Catch {
                        $WindowType = "w"
                    }
                    Try {
                        If ($arrComName[2].ToLower() -match "m") {
                            $MWT = $Settings.IP_Address.Text
                            $ManaberWindow = ([int]$MWT.substring( $MWT.Length -2) - 10)
                            If ($ManaberWindow.Length -ne 2) {
                                $ManaberWindow = ("0" + $ManaberWindow)
                            }
                            #write-output ("Manager PC Setting window to: " + $ManaberWindow)
                            $Window = $ManaberWindow
                        }
                    }
                    Catch {

                    }

                    #Force Dev Password based on name
                    foreach ($Name in $Settings.DEVCNames) {
                        If ($env:COMPUTERNAME -contains $Name ) {
                            $UPrePass = [System.Text.Encoding]::Unicode.GetString([System.Convert]::FromBase64String($Settings.UPrePass))
                            $UPostPass = [System.Text.Encoding]::Unicode.GetString([System.Convert]::FromBase64String($Settings.UPostPass))
                            $State = [System.Text.Encoding]::Unicode.GetString([System.Convert]::FromBase64String($Settings.DefaultState))
                        }
                    }
                    #Loop and Set Passsword for users
                    ForEach ( $LocalUser in ((Get-LocalUser).name | Sort-Object {"$_" -replace '\d',''},{("$_" -replace '\D','') -as [int]})) {
                        If (-Not ($Settings.AccountBlacklist.contains($LocalUser))) {
                            #Get Window number from username
                            Try{
                                $UserWindow = $LocalUser -replace  "\D+" -replace '\b0*\B',''
                                If (-Not $UserWindow) {
                                    $UserWindow = 01
                                }
                                If ($UserWindow.Length -ne 2) {
                                    $UserWindow = ("0" + $UserWindow)
                                }

                            }
                            Catch {
                                $UserWindow = 01
                            }
                            #Set New Password
                            $TempPass = -join ($UPrePass,$State.ToUpper(),$Store,$WindowType,$UserWindow,$UPostPass)
                            If (-Not [string]::IsNullOrEmpty($TempPass)) {
                                Set-LocalUser -Name ($LocalUser) -PasswordNeverExpires:$true -Password (ConvertTo-SecureString $TempPass -AsPlainText -Force) | Out-Null
                            }
                        }
                        If ($LocalUser.ToUpper() -eq $settings.Admin.ToUpper()) {
                            If ($Settings.Machine_Name.Text) {
                                $APrePass = [System.Text.Encoding]::Unicode.GetString([System.Convert]::FromBase64String($Settings.APrePass))
                                $APostPass = [System.Text.Encoding]::Unicode.GetString([System.Convert]::FromBase64String($Settings.APostPass))
                                $arrComName = $Settings.Machine_Name.Text -split "-"
                                #Deal with Dev boxes.
                                Try{
                                    $State = $arrComName[0].ToUpper()
                                    If ($State.Length -ne 2) {
                                        $UPrePass = [System.Text.Encoding]::Unicode.GetString([System.Convert]::FromBase64String($Settings.UPrePass))
                                        $UPostPass = [System.Text.Encoding]::Unicode.GetString([System.Convert]::FromBase64String($Settings.UPostPass))
                                        $State = [System.Text.Encoding]::Unicode.GetString([System.Convert]::FromBase64String($Settings.DefaultState))
                                    }
                                }
                                Catch {
                                    $State = [System.Text.Encoding]::Unicode.GetString([System.Convert]::FromBase64String($Settings.DefaultState))
                                }
                                $TempPass = -join ($APrePass,$State.ToLower(),$APostPass)
                                If (-Not [string]::IsNullOrEmpty($TempPass)) {
                                    #Set Local Admin.
                                    $SetPassAdmin = $true
                                    Set-LocalUser -Name ($LocalUser) -PasswordNeverExpires:$true -Password (ConvertTo-SecureString $TempPass -AsPlainText -Force) | Out-Null
                                    #Update password for autologon
                                    If ((Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name 'AutoAdminLogon' -ErrorAction SilentlyContinue).AutoAdminLogon -ne 0) {
                                        If ((Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name 'DefaultPassword' -ErrorAction SilentlyContinue).DefaultPassword) {  
                                            Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name 'DefaultPassword' -Value $TempPass  -Type String
                                        }
                                    }
                                }   
                            }
                        }
                        #Disable all accounts not Admin, User or Blacklist
                        #Remove black listed accounts
                        If ($Settings.AccountDisableBlacklist -notcontains $LocalUser) {
                            #Disable Accounts.
                            # write-output ("Disabled Non-Window " + $LocalUser.Name + " account . . .")
                            Disable-LocalUser -Name $LocalUser -Confirm:$false

                        } 
                    }
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
            If ([string]::IsNullOrEmpty($Settings.IP_Address.Text) -eq $false) {
                If ([string]::IsNullOrEmpty($Settings.Network_Adapter.SelectedItem.ToString()) -eq $false) {
                    #Get NetAdapter info
                    $CurrentNic = $settings.Network_Adapter_List |Where-Object {$_.Description -eq $Settings.Network_Adapter.SelectedItem.ToString()}
                    #Calculate Network Address and Gateway if Needed
                    If ([string]::IsNullOrEmpty($CurrentNic.DefaultIPGateway)) {
                        $Settings.NetworkAddress = [IPAddress] (([IPAddress]$Settings.IP_Address.Text ).Address -band ([IPAddress] $Settings.IP_Subnet.Text  ).Address)
                        If ([string]::IsNullOrEmpty($Settings.NetworkAddress.IPAddressToString) -eq $false) {
                            $Settings.NetworkAddressGateway = (Get-IPAddressFromUInt32 -UInt32 ((Get-UInt32FromIPAddress -IPAddress $Settings.NetworkAddress.IpAddressToString) +1)).IPAddressToString
                        }
                    }Else {
                        $Settings.NetworkAddressGateway = [string]$CurrentNic.DefaultIPGateway
                    }
                    #Prompt for Gateway if not set
                    if ([string]::IsNullOrEmpty($Settings.NetworkAddressGateway) -eq $True) {
                        Add-Type -AssemblyName System.Windows.Forms

                        $form = New-Object System.Windows.Forms.Form
                        $form.Text = "Enter Network Address Gateway"
                        $form.Size = New-Object System.Drawing.Size(350,150)
                        $form.StartPosition = "CenterScreen"

                        $label = New-Object System.Windows.Forms.Label
                        $label.Text = "Network Address Gateway:"
                        $label.AutoSize = $true
                        $label.Location = New-Object System.Drawing.Point(10,20)
                        $form.Controls.Add($label)

                        $textBox = New-Object System.Windows.Forms.TextBox
                        $textBox.Size = New-Object System.Drawing.Size(200,20)
                        $textBox.Location = New-Object System.Drawing.Point(160,18)
                        $textBox.Text = $Settings.NetworkAddressGateway
                        $form.Controls.Add($textBox)

                        $okButton = New-Object System.Windows.Forms.Button
                        $okButton.Text = "OK"
                        $okButton.Location = New-Object System.Drawing.Point(100,60)
                        $okButton.Add_Click({
                            $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
                            $form.Close()
                        })
                        $form.Controls.Add($okButton)

                        $cancelButton = New-Object System.Windows.Forms.Button
                        $cancelButton.Text = "Cancel"
                        $cancelButton.Location = New-Object System.Drawing.Point(180,60)
                        $cancelButton.Add_Click({
                            $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                            $form.Close()
                        })
                        $form.Controls.Add($cancelButton)

                        if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                            $Settings.NetworkAddressGateway = $textBox.Text
                        }
                    }

                    #Set IP Address,Gateway,DNS
                    If (Get-Command New-NetIPAddress -errorAction SilentlyContinue) {
                        If ($CurrentNic.DHCPEnabled) {
                            New-NetIPAddress -IPAddress $Settings.IP_Address.Text -PrefixLength (Convert-SubnetMask $Settings.IP_Subnet.Text) -DefaultGateway $Settings.NetworkAddressGateway -InterfaceIndex $CurrentNic.InterfaceIndex
                            Set-DnsClientServerAddress -InterfaceIndex $CurrentNic.InterfaceIndex -ServerAddresses ([string]$Settings.DNS -join ",")
                        }Else {
                            Set-NetIPAddress -IPAddress $Settings.IP_Address.Text -PrefixLength (Convert-SubnetMask $Settings.IP_Subnet.Text) -DefaultGateway $Settings.NetworkAddressGateway -InterfaceIndex $CurrentNic.InterfaceIndex
                        }

                    }Else{
                        If (Get-Command Get-CimInstance -errorAction SilentlyContinue) {
                            $wmi = Get-CimInstance win32_networkadapterconfiguration -filter ("Description = '" + $Settings.Network_Adapter.SelectedItem.ToString() + "'")
                        } Else {
                            $wmi = Get-WmiObject win32_networkadapterconfiguration -filter ("Description = '" + $Settings.Network_Adapter.SelectedItem.ToString() + "'")
                        } 
                        #Only change IP if it is different.
                        If (( $wmi.ipaddress | Where-object {$_.IPaddress -notlike "169.254.*" -and $_.IPAddress -ne "127.0.0.1"}) -ne $Settings.IP_Address.Text) {
                            $wmi.EnableStatic($Settings.IP_Address.Text, $Settings.IP_Subnet.Text )              
                            $wmi.SetGateways($Settings.NetworkAddressGateway, 1)        
                            $wmi.SetDNSServerSearchOrder($Settings.DNS)
                        }
                    }
                    #Test IP Set
                    Start-Sleep -Seconds 5
                    If (Get-Command Get-CimInstance -errorAction SilentlyContinue) {
                        $wmiTest = Get-CimInstance win32_networkadapterconfiguration -filter ("Description = '" + $Settings.Network_Adapter.SelectedItem.ToString() + "'")
                    } Else {
                        $wmiTest = Get-WmiObject win32_networkadapterconfiguration -filter ("Description = '" + $Settings.Network_Adapter.SelectedItem.ToString() + "'")
                    }
                    If (( $wmiTest.ipaddress | Where-object {$_.IPaddress -notlike "169.254.*" -and $_.IPAddress -ne "127.0.0.1"}) -ne $Settings.IP_Address.Text) {
                        [System.Windows.MessageBox]::Show(('Failed to set IP Address to: ' + $Settings.IP_Address.Text + [Environment]::NewLine + 
                        'Current IP set to: ' + $wmiTest.IPAddress + [Environment]::NewLine +
                        'DHCP Enabled: ' + $wmiTest.DHCPEnabled + [Environment]::NewLine +
                        "Please check settings. . . " ),('Failed to set IP Address to: ' + $Settings.IP_Address.Text + " Please check settings. . ." ),'OK','Hand') 
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

            #region Stop Auto Launch 
            If (Get-Service -DisplayName Ivanti*) {
                try {
                    $TaskService = New-Object -ComObject "Schedule.Service"
                }
                catch [Management.Automation.PSArgumentException] {
                    throw $_
                }
                try {
                    $TaskService.Connect()
                }
                catch [Management.Automation.MethodInvocationException] {
                    Write-Error "Error connecting to '$ComputerName' - '$_'"
                    return
                }
                $rootFolder = $TaskService.GetFolder("\")
                try {
                    $Task=($rootFolder.GetTasks(0)| Where-Object  {([xml]$_.xml).Task.Actions.Exec.Command -contains $MyInvocation.MyCommand.Name})
                    #$taskDefinition = ( $rootFolder.GetTasks(0)| Where-Object  {$_.name -eq $taskName} ).Definition
                    Disable-ScheduledTask -TaskName $Task.Name
                }
                catch [Management.Automation.MethodInvocationException] {
                    Write-Error "Scheduled task '$task.Name' not found on '$computerName'."
                    return
                }
            }
            #endregion Stop Auto Launch 
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
        $Settings.Store_Setup.text = ( $Settings.WindowTitle + " Done. Time: " + (Format-ElapsedTime($Settings.sw.Elapsed)) ) 
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

function Update-IPForm {
    $Settings.IP_Address.Text = ($Settings.Network_Adapter_List | Where-object {$_.Description -eq $Settings.Network_Adapter.SelectedItem}).IPAddress
    $Settings.IP_Address.Refresh()    
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
#region BIOS update
    #Get Model Info
    $SysInfo = Get-CimInstance -ClassName Win32_ComputerSystem
    #Get BIOS Info
    $SysBIOS = Get-CimInstance -ClassName win32_bios
    $Temp = $SysBIOS.SMBIOSBIOSVersion -split " "
    $BIOSSystemVersion = [int64]($Temp[$Temp.count - 1] -replace  "\D+")
    #Enum Model
    Switch -Wildcard ($SysInfo.Model) {
        "HP t620*" {
            $BIOSFolder = (Get-ChildItem -Directory -Path (Split-Path -Parent -Path $MyInvocation.MyCommand.Definition)  -Filter "*BIOS*HP t620*" | Select-Object -Last 1)
        }
        "HP t630*" {
            $BIOSFolder = (Get-ChildItem -Directory -Path (Split-Path -Parent -Path $MyInvocation.MyCommand.Definition)  -Filter "*BIOS*HP t630*" | Select-Object -Last 1) 
        } 
        Default {
            #"No matches"
        }       
    }

    #Collect Info
    If ($BIOSFolder) {
        $BIOSInstallers = $BIOSFolder| Get-ChildItem -File -Filter *.exe
        $BIOSBin = $BIOSFolder| Get-ChildItem -File -Filter *.bin | Select-Object -Last 1
        $BIOSVersion = [int]($BIOSBin.Name -split "_" -replace ".bin" -replace  "\D+" | Select-Object -Last 1)
        If ($BIOSSystemVersion -lt $BIOSVersion) {
            #Installer type
            If ($SysInfo.SystemType -match "x64") {
                $BIOSInstaller = $BIOSInstallers | Where-Object {$_.Name -match "X64"} | Select-Object -Last 1
            } else {
                $BIOSInstaller = $BIOSInstallers | Where-Object {$_.Name -notmatch "X64"} | Select-Object -Last 1
            }
            #Ask to Install update.
            If ($BIOSInstaller) {
                If([System.Windows.MessageBox]::Show(('Would you like to update BIOS from: '+ $BIOSSystemVersion + " to: " + $BIOSVersion + "?"),'BIOS Update','YesNo','Question') -eq "Yes") {
                    #Update BIOS
                    Start-Process -FilePath ($BIOSInstaller.FullName) -ArgumentList $BIOSBin.Name -WorkingDirectory (Split-Path -Parent -Path $BIOSBin.FullName) -Wait
                    #Reboot Maching after install is done.
                    If([System.Windows.MessageBox]::Show(('Would you like to Reboot?'),'System Reboot','YesNo','Question') -eq "Yes") {
                        Restart-Computer
                    }
                }
            }
        }
    }
#endregion BIOS update

#region Landesk Agent Install
    #Test for LanDesk Servcies
    If (-Not (Get-Service -DisplayName LANDesk*,Managed*,Ivanti* -ErrorAction SilentlyContinue)) {
        #Find LanDesk Installer
        $LanDeskInstaller = Get-ChildItem -Path (Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) -Recurse -Filter "EPMAgentInstaller.exe" | Select-Object -First 1
        If ($LanDeskInstaller) {
             If([System.Windows.MessageBox]::Show(('Would you like install: '+ (Split-Path -Leaf -Path (Split-Path -Parent -Path $LanDeskInstaller.FullName)) + "?"),'Ivanti Agent Install','YesNo','Question') -eq "Yes") {
                #Install Agent
                Start-Process -FilePath ($LanDeskInstaller.FullName) -Wait
            
                #region Monitor for agent install
                    $EPMAgentSW = [Diagnostics.Stopwatch]::StartNew()
                    Add-Type -AssemblyName System.Windows.Forms

                    $form = New-Object System.Windows.Forms.Form
                    $form.Text = "Folder Monitor"
                    $form.Size = New-Object System.Drawing.Size(300,200)
                    $form.StartPosition = "CenterScreen"

                    $label = New-Object System.Windows.Forms.Label
                    $label.Location = New-Object System.Drawing.Point(20,10)
                    $label.Size = New-Object System.Drawing.Size(250,80)
                    $form.Controls.Add($label)

                    # Create the progress bar
                    $progressBar = New-Object System.Windows.Forms.ProgressBar
                    $progressBar.Location = New-Object System.Drawing.Point(20,110)
                    $progressBar.Size = New-Object System.Drawing.Size(250,20)
                    $progressBar.Minimum = 0
                    $progressBar.Maximum = $Settings.EPMFolderItemCount
                    $form.Controls.Add($progressBar)

                    function Update-Count {
                        $count = (Get-ChildItem -Path $Settings.EPMAgentPath -Force | Measure-Object).Count
                        $label.Text = ("Please wait, operation in progress... " + [System.Environment]::NewLine + "EPM is intalled when " + $Settings.EPMFolderItemCount + " items are in the folder." + [System.Environment]::NewLine + [System.Environment]::NewLine + "EPM Agent Folder Item Count: $count" + [System.Environment]::NewLine + "Elapsed Time: " +  (Format-ElapsedTime($EPMAgentSW.Elapsed)))
                        $progressBar.Value = [Math]::Min($count, $Settings.EPMFolderItemCount)
                        If ($count -ge $Settings.EPMFolderItemCount) { $form.Close() }
                    }

                    $timer = New-Object System.Windows.Forms.Timer
                    $timer.Interval = 3000
                    $timer.Add_Tick({ Update-Count })
                    $timer.Start()

                    Update-Count

                    $form.Add_FormClosed({ $timer.Stop() })
                    if ($form.IsDisposed -eq $false) {
                        [void]$form.ShowDialog()
                    }
                    $form.Dispose()
                    $timer.Dispose()
                #endregion Monitor for agent install
             }
        }
    }Elseif((Get-ChildItem $Settings.EPMAgentPath).count -lt $Settings.EPMFolderItemCount) {
                        #region Monitor for agent install
                $EPMAgentSW = [Diagnostics.Stopwatch]::StartNew()
                Add-Type -AssemblyName System.Windows.Forms

                $form = New-Object System.Windows.Forms.Form
                $form.Text = "Folder Monitor"
                $form.Size = New-Object System.Drawing.Size(300,200)
                $form.StartPosition = "CenterScreen"

                $label = New-Object System.Windows.Forms.Label
                $label.Location = New-Object System.Drawing.Point(20,10)
                $label.Size = New-Object System.Drawing.Size(250,80)
                $form.Controls.Add($label)

                # Create the progress bar
                $progressBar = New-Object System.Windows.Forms.ProgressBar
                $progressBar.Location = New-Object System.Drawing.Point(20,110)
                $progressBar.Size = New-Object System.Drawing.Size(250,20)
                $progressBar.Minimum = 0
                $progressBar.Maximum = $Settings.EPMFolderItemCount
                $form.Controls.Add($progressBar)

                function Update-Count {
                    $count = (Get-ChildItem -Path $Settings.EPMAgentPath -Force | Measure-Object).Count
                    $label.Text = ("Please wait, operation in progress... " + [System.Environment]::NewLine + "EPM is intalled when " + $Settings.EPMFolderItemCount + " items are in the folder." + [System.Environment]::NewLine + [System.Environment]::NewLine + "EPM Agent Folder Item Count: $count" + [System.Environment]::NewLine + "Elapsed Time: " +  (Format-ElapsedTime($EPMAgentSW.Elapsed)))
                    $progressBar.Value = [Math]::Min($count, $Settings.EPMFolderItemCount)
                    If ($count -ge $Settings.EPMFolderItemCount) { $form.Close() }
                }

                $timer = New-Object System.Windows.Forms.Timer
                $timer.Interval = 3000
                $timer.Add_Tick({ Update-Count })
                $timer.Start()

                Update-Count

                $form.Add_FormClosed({ $timer.Stop() })
                if ($form.IsDisposed -eq $false) {
                    [void]$form.ShowDialog()
                }
                $form.Dispose()
                $timer.Dispose()
            #endregion Monitor for agent install
    }
#endregion Landesk Agent Install
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
$Settings.Machine_Name.text               = $env:computername


If (Get-Command Get-CimInstance -errorAction SilentlyContinue) {
    $Settings.Network_Adapter_List = Get-CimInstance -Class Win32_NetworkAdapterConfiguration -Filter 'IPEnabled = True' | Where-object {$_.IPaddress -notlike "169.254.*" -and $_.IPAddress -ne "127.0.0.1" -and $_.IPaddress -notlike '*:*' -and $_.Description -notmatch "Hyper-V|VMnet1|VMnet8" } | Select-Object InterfaceAlias,InterfaceIndex,Description,IPAddress,DefaultIPGateway,IPSubnet,DNSServerSearchOrder,DHCPEnabled
} Else {
    $Settings.Network_Adapter_List = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter 'IPEnabled = True' | Where-object {$_.IPaddress -notlike "169.254.*" -and $_.IPAddress -ne "127.0.0.1" -and $_.IPaddress -notlike '*:*' -and $_.Description -notmatch "Hyper-V|VMnet1|VMnet8"} | Select-Object InterfaceAlias,InterfaceIndex,Description,IPAddress,DefaultIPGateway,IPSubnet,DNSServerSearchOrder,DHCPEnabled
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
$Settings.IP_Address.Text                 = ($Settings.Network_Adapter_List | Select-Object -first 1).IPAddress
# $Settings.IP_Address.Enabled              = $false

$Settings.IP_Subnet_Label                = New-Object system.Windows.Forms.Label
$Settings.IP_Subnet_Label.text           = "IP Subnet mask:"
$Settings.IP_Subnet_Label.AutoSize       = $true
$Settings.IP_Subnet_Label.width          = 25
$Settings.IP_Subnet_Label.height         = 10
$Settings.IP_Subnet_Label.location       = New-Object System.Drawing.Point(10,90)
$Settings.IP_Subnet_Label.Font           = 'Microsoft Sans Serif,10'

$Settings.IP_Subnet                      = New-Object system.Windows.Forms.TextBox
$Settings.IP_Subnet.multiline            = $false
$Settings.IP_Subnet.width                = 180
$Settings.IP_Subnet.height               = 20
$Settings.IP_Subnet.location             = New-Object System.Drawing.Point(125,90)
$Settings.IP_Subnet.Font                 = 'Microsoft Sans Serif,10'
if (($Settings.Network_Adapter_List | Select-Object -first 1).IPSubnet){
    $Settings.IP_Subnet.Text                 = ($Settings.Network_Adapter_List | Select-Object -first 1).IPSubnet
}Else{
    $Settings.IP_Subnet.Text                 = $Settings.Subnet
}

$Settings.Network_Adapter_Label                = New-Object system.Windows.Forms.Label
$Settings.Network_Adapter_Label.text           = "Network Adapter:"
$Settings.Network_Adapter_Label.AutoSize       = $true
$Settings.Network_Adapter_Label.width          = 25
$Settings.Network_Adapter_Label.height         = 10
$Settings.Network_Adapter_Label.location       = New-Object System.Drawing.Point(10,120)
$Settings.Network_Adapter_Label.Font           = 'Microsoft Sans Serif,10'

$Settings.Network_Adapter                       = New-Object system.Windows.Forms.ComboBox
#$Settings.Network_Adapter.text                  = " "
$Settings.Network_Adapter.width                 = 265
$Settings.Network_Adapter.height                = 20
$Settings.Network_Adapter.location              = New-Object System.Drawing.Point(125,120)
$Settings.Network_Adapter.Font                  = 'Microsoft Sans Serif,10'


$Settings.WindowLogonUser_Label                = New-Object system.Windows.Forms.Label
$Settings.WindowLogonUser_Label.text           = "Window User:"
$Settings.WindowLogonUser_Label.AutoSize       = $true
$Settings.WindowLogonUser_Label.width          = 25
$Settings.WindowLogonUser_Label.height         = 10
$Settings.WindowLogonUser_Label.location       = New-Object System.Drawing.Point(10,145)
$Settings.WindowLogonUser_Label.Font           = 'Microsoft Sans Serif,10'

$Settings.WindowLogonUser                       = New-Object system.Windows.Forms.ComboBox
$Settings.WindowLogonUser.text                  = " "
$Settings.WindowLogonUser.width                 = 265
$Settings.WindowLogonUser.height                = 20
$Settings.WindowLogonUser.location              = New-Object System.Drawing.Point(125,145)
$Settings.WindowLogonUser.Font                  = 'Microsoft Sans Serif,10'


$Settings.Manager                      = New-Object System.Windows.Forms.Checkbox 
$Settings.Manager.Text                 = "Manager"
$Settings.Manager.width                = 180
$Settings.Manager.height               = 20
$Settings.Manager.Location             = New-Object System.Drawing.Size(125,170) 
$Settings.Manager.Font                 = 'Microsoft Sans Serif,10'
$Settings.Manager.Checked              = $False


$Settings.FBackup = New-Object System.Windows.Forms.GroupBox #create the group box
$Settings.FBackup.Location = New-Object System.Drawing.Size(10,185) #location of the group box (px) in relation to the primary window's edges (length, height)
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

$Settings.Store_Setup.controls.AddRange(@($Settings.Machine_Name_Label,$Settings.IP_Address_Label,$Settings.IP_Subnet_Label,$Settings.IP_Subnet,$Settings.Machine_Name,$Settings.IP_Address,$Settings.Network_Adapter_Label,$Settings.Network_Adapter,$Settings.Restore_Backup,$Settings.Start,$Settings.Stop,$Settings.Restore_Backup_Label,$Settings.Browse,$Settings.WindowLogonUser_Label,$Settings.WindowLogonUser,$Settings.Manager,$Settings.FBackup))


#############################################################################
#endregion Setup Sessions
#############################################################################
#############################################################################
#region Main 
#############################################################################

$Settings.Browse.Add_Click({ Browse_File })
$Settings.Start.Add_Click({ Start_Work })
$Settings.Stop.Add_Click({ Stop_Work })
$Settings.Network_Adapter.Add_SelectedIndexChanged({ Update-IPForm })
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

If ($Settings.Network_Adapter.Items.Count -ge 0) {
    If (($Settings.Network_Adapter_List | Where-object {$_.IPAddress -eq $Settings.IP_Address.Text}).InterfaceAlias){
        $Settings.Network_Adapter.SelectedItem =  ($Settings.Network_Adapter_List | Where-object {$_.IPAddress -eq $Settings.IP_Address.Text}).InterfaceAlias
    }Else{
        $Settings.Network_Adapter.SelectedItem =  ($Settings.Network_Adapter_List | Where-object {$_.IPAddress -eq $Settings.IP_Address.Text}).Description
    }
}

[void]$Settings.Store_Setup.ShowDialog()

#############################################################################
#endregion Main
#############################################################################
