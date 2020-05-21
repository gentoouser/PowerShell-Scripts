<# 
.SYNOPSIS
    Name: DisableAccount.ps1
    Disables users in AD

.DESCRIPTION
  Disables users in AD

.EXAMPLE
   & DisableAccount.ps1 

.NOTES
 Changes:
    1.0.1 - Added more validations and save to SharePoint
    1.0.2 - Added Get-Credential to make easier to run. 
    1.0.3 - Updated SharePoint Location.
    1.0.4 - InputBox Function
    1.0.5 - Upload to SharePoint using PNP. Backup Home Drive. Export Mailbox.
    1.5.0 - Updated to form for all the options
    1.5.1 - Alot of bug fixes and Disabled thread. 
    1.5.2 - Exclude Groups for Azure Mail. 
    1.5.3 - Fixed bad - and "Use 'Register-PSRepository -Default' to register the PSGallery repository." issue.
    1.5.4 - Reduce Errors and warnings. Fig logic errors.
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
$Settings.Version = "1.5.4"
$Settings.WindowTitle = ("Disable User - Version: " + $Settings.Version)
$Settings.AdminCred = $null
$Settings.Users = @()
$Settings.CSV = $null
$Settings.ExchangeServer = "Exchange.github.com"
$Settings.SPSite = "https://github.sharepoint.com"
$Settings.SPWebFolder = "Shared Documents/Disabled Users"
$Settings.SPWebURL = "https://github.sharepoint.com/Shared Documents/Disabled Users"
$Settings.DisabledOU = "OU=Disabled Users,DC=Github,DC=com"
$Settings.DisabledPrepend = ('Disabled: ' + (Get-Date -UFormat %Y-%m%d))
$Settings.Desktop = [Environment]::GetFolderPath("Desktop")
$Settings.HomeDrive = $null
$Settings.HomeDriveArchive = "\\coldstorage.github.com\Disabled_Users"
$settings.Admin = "administrator"
$Settings.HomeDriveServers = @(
    New-Object PSObject -Property @{Server = "cluster.github.com";  LocalPath = "F:"; AdminShare = "F$"}
    New-Object PSObject -Property @{Server = "standalone.github.com";  LocalPath = "L:"; AdminShare = "L$"}
    New-Object PSObject -Property @{Server = "user.github.com";  LocalPath = "D:\Users"; AdminShare = "D$\Users"}
)
$Settings.ExcludeGroups = @(
    "Azure AD Connect Group"
)
$Settings.LogFile = ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\Logs\" + `
                    ($MyInvocation.MyCommand.Name -replace ".ps1","") + "_" + `
                    (Get-Date -format yyyyMMdd-hhmm) + ".log")

#region Icon
$iconBase64 ='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6e3u8fPy8/Px7+3s7/Hw7/Dv8e3u7+3u7u3u8O/t7PLy7+3s7O3s6+vr6+vo5gfn6+7w8vPz8/Lx7+3t8vTy7+/v8/Dt7u3t7ero7e3r6evs6uvv8PHy8fLx7uno5ubn6/Dz8/Pz9PPx8O7u8vTy7+/w8vHu7e3t7ujk6evr6unr6ffv8/T09PMA8uvr9+z47fDx9PT09fTy8fDv8fPx7u/x8fHw7O3t7ujk6Onr6+zv7enu9PT09PP18er47vHu7e3u7/Dx8vTz8vPx8PLx7Ozy8fHw7u3t7ugH5+UH9+fq7u3u9PT09PT07eru8vDx7+zs7Pjq7O7v8fT08vPy7ezy8PHv8O3t7ujnB+jp+Ov36Ovu8/X09PXw6+7x7+7v8O7s6+nn9+vq6/D09PTy7+/w8PHu7e3t7ujo7vHy8vLv7er48/X1APHr7vHw7ezs+Pjr6Ofn6Ovt+Ovt8fHv7fHw8e7r6+3t7ufs9PT09PT09fTt7vT17uns7/Dx8/Ly8O/t6+jnB+fq7Oz46+rt7PHw8Ovr6+3t7unw9PPz9PT09PTx6vfn5ejt8PP09PP09fT09fPt5+Tk6u3qB+nu7PDv8Orr6+7u7uvw9PP09PT09PT17ebl6Ovv8/T09PT09PT09fX07eUH+Orn6e/u7+/w7erq6u7u7+zs9fTz9PT09PT18uvo9+zx9PT09PT09PT08/T07ufo6efn6+747+/y+Ovq+O7t7u3r8/X09PT09PT18en37PD09PT09PT09PT08/Tz7+vo5+f37O3q7u7x6ur46+3t7e746/X1APX19PTy7+v47/Lz9PT09fT09PT08/Ty7/fn5+jr7e3r7fDu6urr6+3t7u7u5Pfw8/Pz8/Hv7fj48PLy9PT09fT09PT19fH45wfn5+nq6/jr7vHs6uvr6+3u7e7u6OPn6On47ezt6+rp7O3x8/X08/T19PT08erl5Ofo6Ovq6uv47/Hq6uv46+3t7u7u6+Pl5uvw6+np6Pf39/fq7O/v7e/y8u7qB+Tj5Ojo9+vq+Orr8O/p6+v46+3t7u3u7eLj5+nq5uTl5eXlB+bk5OTo8Obq6vflUuLiB+n36+rq+Pft8e3p+Ovu7e3t7e7u7+f/5OXl5OPj4+Pj4+Ti4v/jB+LkBwfl5OLl9/f37Orr6uju8evq7fjv7+3u7u7t8O724uTk5OXk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'
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
    $Settings.OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    #$Settings.OpenFileDialog.initialDirectory = (Split-Path -Path $MyInvocation.MyCommand.Definition) 
    $Settings.OpenFileDialog.filter = "Comma Separated Values (CSV)|*.CSV|All Files|*.*" 
    $Settings.OpenFileDialog.ShowDialog() | Out-Null
   
    Function Invoke-InputBox {
 
        [cmdletbinding(DefaultParameterSetName="plain")]
        [OutputType([system.string],ParameterSetName='plain')]
        [OutputType([system.security.securestring],ParameterSetName='secure')]
     
        Param(
            [Parameter(ParameterSetName="secure")]
            [Parameter(HelpMessage = "Enter the title for the input box. No more than 25 characters.",
            ParameterSetName="plain")]        
     
            [ValidateNotNullorEmpty()]
            [ValidateScript({$_.length -le 25})]
            [string]$Title = "User Input",
     
            [Parameter(ParameterSetName="secure")]        
            [Parameter(HelpMessage = "Enter a prompt. No more than 50 characters.",ParameterSetName="plain")]
            [ValidateNotNullorEmpty()]
            [ValidateScript({$_.length -le 50})]
            [string]$Prompt = "Please enter a value:",
            
            [Parameter(HelpMessage = "Use to mask the entry and return a secure string.",
            ParameterSetName="secure")]
            [switch]$AsSecureString
        )
        #Source: https://jdhitsolutions.com/blog/powershell/5816/a-powershell-input-tool/
        if ($PSEdition -eq 'Core') {
            Write-Warning "Sorry. This command will not run on PowerShell Core."
            #bail out
            Return
        }
     
        Add-Type -AssemblyName PresentationFramework
        Add-Type –assemblyName PresentationCore
        Add-Type –assemblyName WindowsBase
     
        #remove the variable because it might get cached in the ISE or VS Code
        Remove-Variable -Name myInput -Scope script -ErrorAction SilentlyContinue
     
        $form = New-Object System.Windows.Window
        $stack = New-object System.Windows.Controls.StackPanel
     
        #define what it looks like
        $form.Title = $title
        $form.Height = 150
        $form.Width = 350
     
        $label = New-Object System.Windows.Controls.Label
        $label.Content = "    $Prompt"
        $label.HorizontalAlignment = "left"
        $stack.AddChild($label)
     
        if ($AsSecureString) {
            $inputbox = New-Object System.Windows.Controls.PasswordBox
        }
        else {
            $inputbox = New-Object System.Windows.Controls.TextBox
        }
     
        $inputbox.Width = 300
        $inputbox.HorizontalAlignment = "center"
     
        $stack.AddChild($inputbox)
     
        $space = new-object System.Windows.Controls.Label
        $space.Height = 10
        $stack.AddChild($space)
     
        $btn = New-Object System.Windows.Controls.Button
        $btn.Content = "_OK"
     
        $btn.Width = 65
        $btn.HorizontalAlignment = "center"
        $btn.VerticalAlignment = "bottom"
     
        #add an event handler
        $btn.Add_click( {
                if ($AsSecureString) {
                    $script:myInput = $inputbox.SecurePassword
                }
                else {
                    $script:myInput = $inputbox.text
                }
                $form.Close()
            })
     
        $stack.AddChild($btn)
        $space2 = new-object System.Windows.Controls.Label
        $space2.Height = 10
        $stack.AddChild($space2)
     
        $btn2 = New-Object System.Windows.Controls.Button
        $btn2.Content = "_Cancel"
     
        $btn2.Width = 65
        $btn2.HorizontalAlignment = "center"
        $btn2.VerticalAlignment = "bottom"
     
        #add an event handler
        $btn2.Add_click( {
                $form.Close()
            })
     
        $stack.AddChild($btn2)
     
        #add the stack to the form
        $form.AddChild($stack)
     
        #show the form
        $inputbox.Focus() | Out-Null
        $form.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterScreen
     
        $form.ShowDialog() | out-null
     
        #write the result from the input box back to the pipeline
        $script:myInput
     
    }
    #
    $Settings.CSVHeader = Invoke-InputBox -Title "CSV Header" -Prompt "Please Type in the CSV Username Header"
    $Settings.CSV = import-csv -path $Settings.OpenFileDialog.filename    
    $Settings.Users = ($Settings.CSV | Select-Object $Settings.CSVHeader).($Settings.CSVHeader)
}
function Start_Work {
    param (
        
    )
        #Check to make sure user is found
        If (-Not $Settings.Disable_User.text -and -not $Settings.Users) {
            [System.Windows.Forms.MessageBox]::Show("No AD User to Find" , "Error")
            exit 
        }

        $Settings.sw = [Diagnostics.Stopwatch]::StartNew()
        $Settings.Form.text  = ( $Settings.WindowTitle + " Working . . . Please wait.")
        $Settings.Stop.Text = "Stop"

        #Test Creds
        $Root = "LDAP://" + ([ADSI]'').distinguishedName
        $Domain = New-Object System.DirectoryServices.DirectoryEntry($Root,$Settings.Admin_User.Text,$Settings.Admin_Pass.Text)
        
        If(!$Domain){
            Write-Warning "Something went wrong"
            exit
        }Else{
            If ($null -eq $Domain.name ) {
                [System.Windows.Forms.MessageBox]::Show("Bad Username or Password" , "Error")
                $Settings.Form.text  = ( $Settings.WindowTitle )
                $Settings.Stop.Text = "Disable"
                exit
            }
        }
         #Setup Credentials
         $Settings.AdminCred = New-Object System.Management.Automation.PSCredential  ($Settings.Admin_User.Text,(ConvertTo-SecureString $Settings.Admin_Pass.Text -AsPlainText -Force))
        If (-Not [PSCredential] $Settings.AdminCred) {
            [System.Windows.Forms.MessageBox]::Show("Bad Username or Password" , "Error")
            $Settings.Form.text  = ( $Settings.WindowTitle )
            $Settings.Stop.Text = "Disable"
            exit
        }
        #region Main thread Start
            # $MainRunspace =[runspacefactory]::CreateRunspace()      
            # $MainRunspace.Open()
            # $MainRunspace.SessionStateProxy.SetVariable("Settings",$Settings)  
            # $MainRunspace.SessionStateProxy.SetVariable("SettingsOutput",$SettingsOutput)    

            # $MainpsCmd = "" | Select-Object PowerShell,Handle
            # $MainpsCmd.PowerShell = [PowerShell]::Create().AddScript({ 
            
            #Setup Credentials
           write-host $Settings.Admin_User.Text
            $Settings.AdminCred = New-Object System.Management.Automation.PSCredential  ($Settings.Admin_User.Text,(ConvertTo-SecureString $Settings.Admin_Pass.Text -AsPlainText -Force))
            #region Load Modules
            #Form Modules
            [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
            #Load AD Module
            If (Get-Module -ListAvailable -Name "ActiveDirectory") {
                If (-Not (Get-Module "ActiveDirectory" -ErrorAction SilentlyContinue)) {
                    Import-Module "ActiveDirectory"
                } Else {
                    #write-host "ActiveDirectory PowerShell Module Already Loaded"
                }
            } Else {
                If (Get-WindowsCapability -Name "RSAT.ActiveDirectory*" -Online | Where-Object {$_.State -ne "Installed"}){
                    #write-host "Installing ActiveDirectory PowerShell Module!"
                    Get-WindowsCapability -Name "RSAT.ActiveDirectory*" -Online | Where-Object {$_.State -ne "Installed"} | Add-WindowsCapability -Online 
                    Import-Module "ActiveDirectory"
                } Else {
                    If (Get-WindowsFeature "RSAT-AD-PowerShell" -ErrorAction SilentlyContinue) {
                    #write-host "Installing ActiveDirectory PowerShell Module!"
                    Install-WindowsFeature "RSAT-AD-PowerShell"
                    Import-Module "ActiveDirectory"
                    }Else{
                        [System.Windows.Forms.MessageBox]::Show("Please install ActiveDirectory Powershell Modules" , "Error")
                        exit
                    }
                }
            }
            #Load SharePoint Module
            If (Get-Module -ListAvailable -Name "SharePointPnPPowerShellOnline") {
                If (-Not (Get-Module "SharePointPnPPowerShellOnline" -ErrorAction SilentlyContinue)) {
                    Import-Module "SharePointPnPPowerShellOnline" -DisableNameChecking
                } Else {
                    #write-host "SharePointPnPPowerShellOnline PowerShell Module Already Loaded"
                } 
            } Else {
                Import-Module PackageManagement
                Import-Module PowerShellGet
                If ((Get-PSRepository -name PSGallery).InstallationPolicy -ne "Trusted") {
                    If ((Get-PSRepository -name PSGallery).InstallationPolicy -eq "Untrusted") {
                        Set-PSRepository -InstallationPolicy Trusted -Name "PSGallery"
                    }Else{
                        Register-PSRepository -Default -InstallationPolicy Trusted
                        Set-PSRepository -InstallationPolicy Trusted -Name "PSGallery"
                    }
                }
                Install-Module "SharePointPnPPowerShellOnline"  -Force -Confirm:$false -Scope:CurrentUser -SkipPublisherCheck -AllowClobber
                Import-Module "SharePointPnPPowerShellOnline" -DisableNameChecking
                If (-Not (Get-Module "SharePointPnPPowerShellOnline" -ErrorAction SilentlyContinue)) {
                    [System.Windows.Forms.MessageBox]::Show("Please install SharePointPnPPowerShellOnline Powershell Modules" , "Error")
                    exit
                }
            }
            #Load 7zip Module
            If (Get-Module -ListAvailable -Name "7Zip4PowerShell") {
                If (-Not (Get-Module "7Zip4PowerShell" -ErrorAction SilentlyContinue)) {
                    Import-Module "7Zip4PowerShell"
                } Else {
                    #write-host "7Zip4PowerShell PowerShell Module Already Loaded"
                } 
            } Else {
                Import-Module PackageManagement
                Import-Module PowerShellGet
                If ((Get-PSRepository -name PSGallery).InstallationPolicy -ne "Trusted") {
                    If ((Get-PSRepository -name PSGallery).InstallationPolicy -eq "Untrusted") {
                        Set-PSRepository -InstallationPolicy Trusted -Name "PSGallery"
                    }Else{
                        Register-PSRepository -Default -InstallationPolicy Trusted
                        Set-PSRepository -InstallationPolicy Trusted -Name "PSGallery"
                    }
                }
                Install-Module "7Zip4PowerShell"  -Force -Confirm:$false -Scope:CurrentUser -SkipPublisherCheck -AllowClobber
                Import-Module "7Zip4PowerShell"
                If (-Not (Get-Module "7Zip4PowerShell" -ErrorAction SilentlyContinue)) {
                    [System.Windows.Forms.MessageBox]::Show("Please install 7Zip4PowerShell Powershell Modules" , "Error")
                    exit
                }
            }
            # Load All Exchange PSSnapins 
            If ((Get-PSSession | Where-Object { $_.ConfigurationName -Match "Microsoft.Exchange" }).Count -eq 0 ) {
               write-host ("Loading Exchange Plugins") -foregroundcolor "Green"
                $ERPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri ("http://" + $Settings.ExchangeServer + "/PowerShell/") -Credential  $Settings.AdminCred -Authentication Kerberos
                Import-PSSession $ERPSession -Prefix EP -AllowClobber -DisableNameChecking
                If ((Get-PSSession | Where-Object { $_.ConfigurationName -Match "Microsoft.Exchange" }).Count -eq 0 ) {
                    [System.Windows.Forms.MessageBox]::Show("Troble connecting to " + "http://" + $Settings.ExchangeServer + "/PowerShell/" + "." , "Error")
                    exit
                }
            
            } Else {
                #write-host ("Exchange Plug-ins Already Loaded") -foregroundcolor "Green"
            }
            ## Choose to ignore any SSL Warning issues caused by Self Signed Certificates  

            ## Code From http://poshcode.org/624
            ## Create a compilation environment
            $Provider=New-Object Microsoft.CSharp.CSharpCodeProvider
            $Provider.CreateCompiler()
            $Params=New-Object System.CodeDom.Compiler.CompilerParameters
            $Params.GenerateExecutable=$False
            $Params.GenerateInMemory=$True
            $Params.IncludeDebugInformation=$False
            $Params.ReferencedAssemblies.Add("System.DLL") | Out-Null

            $TASource=@'
namespace Local.ToolkitExtensions.Net.CertificatePolicy{
    public class TrustAll : System.Net.ICertificatePolicy {
    public TrustAll() { 
    }
    public bool CheckValidationResult(System.Net.ServicePoint sp,
        System.Security.Cryptography.X509Certificates.X509Certificate cert, 
        System.Net.WebRequest req, int problem) {
        return true;
    }
    }
}
'@ 
            $TAResults=$Provider.CompileAssemblyFromSource($Params,$TASource)
            $TAAssembly=$TAResults.CompiledAssembly

            ## We now create an instance of the TrustAll and attach it to the ServicePointManager
            $TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
            [System.Net.ServicePointManager]::CertificatePolicy=$TrustAll

            ## end code from http://poshcode.org/624
            #region Exchange online Modules
            #Load ExchangeOnlineManagement Module
            If (Get-Module -ListAvailable -Name "ExchangeOnlineManagement") {
                If (-Not (Get-Module "ExchangeOnlineManagement" -ErrorAction SilentlyContinue)) {
                    Import-Module "ExchangeOnlineManagement" -DisableNameChecking
                } Else {
                    #write-host "ExchangeOnlineManagement PowerShell Module Already Loaded"
                } 
            } Else {
                Import-Module PackageManagement
                Import-Module PowerShellGet
                Import-Module PackageManagement
                Import-Module PowerShellGet
                If ((Get-PSRepository -name PSGallery).InstallationPolicy -ne "Trusted") {
                    If ((Get-PSRepository -name PSGallery).InstallationPolicy -eq "Untrusted") {
                        Set-PSRepository -InstallationPolicy Trusted -Name "PSGallery"
                    }Else{
                        Register-PSRepository -Default -InstallationPolicy Trusted
                        Set-PSRepository -InstallationPolicy Trusted -Name "PSGallery"
                    }
                }
                Install-Module "ExchangeOnlineManagement" -Force -Confirm:$false -Scope:CurrentUser -SkipPublisherCheck -AllowClobber
                Import-Module "ExchangeOnlineManagement" -DisableNameChecking
                If (-Not (Get-Module "ExchangeOnlineManagement" -ErrorAction SilentlyContinue)) {
                    write-error ("Please install ExchangeOnlineManagement Powershell Modules" , "Error")
                    exit
                }
            }
            #Connect
            If (-Not (Get-PSSession | Where-Object {$_.name -match "ExchangeOnline" -and $_.Availability -eq "Available"})) {
                Connect-ExchangeOnline -ShowBanner:$false -UserPrincipalName (([string]([ADSI]"LDAP://<SID=$([System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value)>").UserPrincipalName) -replace "da-" ) -ShowProgress $true
            }
            #endregion Exchange online Modules
            #endregion Load Modules
            #region Thread Functions
            Function Remove-InvalidFileNameChars {
                param(
                  [Parameter(Mandatory=$true,
                    Position=0,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true)]
                  [String]$Name
                )
              
                $invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''
                $re = "[{0}]" -f [RegEx]::Escape($invalidChars)
                return ($Name -replace $re)
            }
            Function Export-Mail() {
                  [CmdletBinding()] 
                  Param 
                  ( 
                      [Parameter(Mandatory=$true,Position=0,HelpMessage="Username or Identity of user.")][string]$User, 
                      [Parameter(Mandatory=$true,Position=1,HelpMessage="Path to archive PST to.")][string]$Archive,
                      [Parameter(Mandatory=$false,Position=2,HelpMessage="Disable user in exchange after export.")][switch]$Disable,
                      [Parameter(Mandatory=$false,Position=3,HelpMessage="Admin Credentials for File Tests.")][pscredential]$Credential
                  ) 
                [bool]$MapiEnabled=$false
                [bool]$O365=$false
                
                #Get User Mailbox object
                $ObjUser = Get-User $User
                If ($ObjUser.RecipientType -ne "UserMailbox" ) {
                    $ObjUser = Get-EPUser $User
                }Else {
                    $O365=$true
                }
                $PSFile = ($Archive + "\" + $($ObjUser.SamAccountName)  + "_MailExport_" + (Get-Date -format yyyyMMdd)+ ".pst")
                #Test Archive Folder
                If (-Not (Test-Path "PSHome:\") -or (Get-PSdrive -name "PSHome").root -ne $Archive) {
                    If (Test-Path "PSHome:\") {
                        Remove-PSDrive -Name "PSHome"
                    }
                    New-PSDrive -Name "PSHome" -PSProvider FileSystem -Root  $Archive  -Credential $Settings.AdminCred  -ErrorAction SilentlyContinue | Out-Null
                    If (-Not $?) {
                        Write-Warning ("Path not valid: $Archive")
                        Return
                    }
                }
                If ($ObjUser.RecipientType -eq "UserMailbox" ) {
                    If ($O365){
                        $CurrentMailBox = $ObjUser | Get-Mailbox
                        #Testing to see if is in queue
                        If ((Get-MailboxExportRequest | Where-Object { $_.Mailbox -eq $CurrentMailBox.Identity}).count -eq 0) {
                            write-host ("`tExport Mail Name: " + $ObjUser.Name + " Alias: " + $ObjUser.SamAccountName + " Email: " + $ObjUser.WindowsEmailAddress)  -foregroundcolor "Cyan"
                            #test to see if User has been exported
                            if (Test-Path ($PSFile) -Credential $Credential) {
                                    Write-Warning ("User: " + $ObjUser.SamAccountName + " already has been exported to: " + $PSFile)
                                    Return
                            }
                            #Test to see of MAPI is enabled
                            if (-Not (Get-CASMailbox -Identity $ObjUser.SamAccountName).MapiEnabled) {
                                #Enable MAPI
                                Set-CASMailbox -Identity $ObjUser.SamAccountName -MAPIEnabled $true
                                [System.GC]::Collect()
                                Start-Sleep -Seconds 5
                            }else{ 
                                $MapiEnabled=$true
                            }
                            #Export Mailbox to PST
                            New-MailboxExportRequest -Mailbox $ObjUser.SamAccountName -FilePath $PSFile | Out-Null
                            If (-Not $?) {
                                Return
                            }
                            Start-Sleep -Seconds 15
                        } else {
                            write-host ("`t`tUser " + $ObjUser.Name + " already submitted. ")
                        }
                        #Monitor Export	
                        $ExportJobStatusName = $null
                        $ExportJobStatusName = Get-MailboxExportRequest | Where-Object { $_.mailbox -eq $CurrentMailBox.Identity -and $_.status -ne 10 } | Select-Object -first 1 | Get-MailboxExportRequestStatistics 
                        If ($null -ne $ExportJobStatusName) {
                            #write-output ("`t`t`t Job Status loop: " + $ExportJobStatusName.status)
                            while (($ExportJobStatusName.status -ne "Completed") -And ($ExportJobStatusName.status -ne "Failed")) {
                                #View Status of Mailbox Export
                                $ExportJobStatusName = Get-MailboxExportRequest | Where-Object { $_.mailbox -eq $CurrentMailBox.Identity -and $_.status -ne 10 } | Select-Object -first 1 | Get-MailboxExportRequestStatistics 
                                Write-Progress -Id ($Id+1) -Activity $("Exporting user: " + $ExportJobStatusName.SourceAlias ) -status $("Export Percent Complete: " + $ExportJobStatusName.PercentComplete + " Copied " + $ExportJobStatusName.BytesTransferred + " out of " + $ExportJobStatusName.EstimatedTransferSize ) -percentComplete $ExportJobStatusName.PercentComplete
                                Start-Sleep -Seconds 15
                            }
                        }
                
                        #Check for Completion status
                        $ExportMailBoxList = Get-MailboxExportRequest | Where-Object { $_.Mailbox -eq $CurrentMailBox.Identity -And ($_.status -ne "Completed" -Or $_.status -ne "Failed")}
                            
                        If ($ExportMailBoxList.status -eq "Completed") {
                            #Remove Exchange account of PST was successful. 
                            write-host ("`t`t Removing Mailbox from Exchange: " + $CurrentMailBox.Identity)
                            #Disable MAPI unless it was already enabled
                            Set-CASMailbox -Identity $ObjUser.SamAccountName -MAPIEnabled $MapiEnabled
                            If ($Disable) {
                                Disable-Mailbox -Identity $ObjUser.SamAccountName -confirm:$false
                            }
                            $ExportMailBoxList | Remove-MailboxExportRequest -Confirm:$false
                        }
                        #Stop if PST Export failed.
                        If ($ExportMailBoxList.status -eq "Failed") {
                            throw ("PST Export failed: " + $error[0].Exception)
                            Break
                        }
                    }Else{
                        $CurrentMailBox = $ObjUser | Get-EPMailbox
                        #Testing to see if is in queue
                        If ((Get-EPMailboxExportRequest | Where-Object { $_.Mailbox -eq $CurrentMailBox.Identity}).count -eq 0) {
                            write-host ("`tExport Mail Name: " + $ObjUser.Name + " Alias: " + $ObjUser.SamAccountName + " Email: " + $ObjUser.WindowsEmailAddress)  -foregroundcolor "Cyan"

                            #test to see if User has been exported
                            if (Test-Path ($PSFile)) {
                                    Write-Warning ("User: " + $ObjUser.SamAccountName + " already has been exported to: " + $PSFile)
                                    Return
                            }
                            #Test to see of MAPI is enabled
                            if (-Not (Get-EPCASMailbox -Identity $ObjUser.SamAccountName).MapiEnabled) {
                                #Enable MAPI
                                Set-EPCASMailbox -Identity $ObjUser.SamAccountName -MAPIEnabled $true
                                [System.GC]::Collect()
                                Start-Sleep -Seconds 5
                            }else{ 
                                $MapiEnabled=$true
                            }
                            #Export Mailbox to PST
                            New-EPMailboxExportRequest -Mailbox $ObjUser.SamAccountName -FilePath $PSFile | Out-Null
                            If (-Not $?) {
                                Return
                            }
                            Start-Sleep -Seconds 15
                        } else {
                            write-host ("`t`tUser " + $ObjUser.Name + " already submitted. ")
                        }
                        #Monitor Export	
                        $ExportJobStatusName = $null
                        $ExportJobStatusName = Get-EPMailboxExportRequest | Where-Object { $_.mailbox -eq $CurrentMailBox.Identity -and $_.status -ne 10 } | Select-Object -first 1 | Get-EPMailboxExportRequestStatistics 
                        If ($null -ne $ExportJobStatusName) {
                            #write-output ("`t`t`t Job Status loop: " + $ExportJobStatusName.status)
                            while (($ExportJobStatusName.status -ne "Completed") -And ($ExportJobStatusName.status -ne "Failed")) {
                                #View Status of Mailbox Export
                                $ExportJobStatusName = Get-EPMailboxExportRequest | Where-Object { $_.mailbox -eq $CurrentMailBox.Identity -and $_.status -ne 10 } | Select-Object -first 1 | Get-EPMailboxExportRequestStatistics 
                                Write-Progress -Id ($Id+1) -Activity $("Exporting user: " + $ExportJobStatusName.SourceAlias ) -status $("Export Percent Complete: " + $ExportJobStatusName.PercentComplete + " Copied " + $ExportJobStatusName.BytesTransferred + " out of " + $ExportJobStatusName.EstimatedTransferSize ) -percentComplete $ExportJobStatusName.PercentComplete
                                Start-Sleep -Seconds 15
                            }
                        }
                
                        #Check for Completion status
                        $ExportMailBoxList = Get-EPMailboxExportRequest | Where-Object { $_.Mailbox -eq $CurrentMailBox.Identity -And ($_.status -ne "Completed" -Or $_.status -ne "Failed")}
                            
                        If ($ExportMailBoxList.status -eq "Completed") {
                            #Remove Exchange account of PST was successful. 
                            write-host ("`t`t Removing Mailbox from Exchange: " + $CurrentMailBox.Identity)
                            #Disable MAPI unless it was already enabled
                            Set-EPCASMailbox -Identity $ObjUser.SamAccountName -MAPIEnabled $MapiEnabled
                            If ($Disable) {
                                Disable-EPMailbox -Identity $ObjUser.SamAccountName -confirm:$false
                            }
                            $ExportMailBoxList | Remove-EPMailboxExportRequest -Confirm:$false
                        }
                        #Stop if PST Export failed.
                        If ($ExportMailBoxList.status -eq "Failed") {
                            throw ("PST Export failed: " + $error[0].Exception)
                            Break
                        }
                    }
  
                  }
            }       
            #endregion Thread Functions         
            
            #Connect to PNP Online
            Connect-PnPOnline -Url $Settings.SPSite -UseWebLogin
            #Add users to array
            If ($Settings.Disable_User.text -contains ",") {
                $Settings.Users += $Settings.Disable_User.text -split ","
            }else{
                $Settings.Users += $Settings.Disable_User.text
            }
            ForEach ($User in $Settings.Users) {
                #Get AD User Info
                $ADUser = (Get-ADUser $User -Properties * -Credential  $Settings.AdminCred)

                #Check to make sure user is found
                If (-Not $ADUser) {
                   write-host ("AD User " + $User + " not found. Moving to next.")
                    continue 
                }

                #regon Find Home Drive
                 ForEach ($FS in $Settings.HomeDriveServers) {
                    Remove-PSDrive -Name "PSHome" -Force -ErrorAction SilentlyContinue | Out-Null
                    New-PSDrive -Name "PSHome" -PSProvider FileSystem -Root  ("\\" + $FS.Server + "\" + $FS.AdminShare + "\" + $User) -Credential $Settings.AdminCred -ErrorAction SilentlyContinue| Out-Null
                    #Validate Folders
                    If (Test-Path "PSHome:\") {
                        $Settings.HomeDrive = ("\\" + $FS.Server + "\" + $FS.AdminShare + "\" + $User)
                        break
                    }Else{ 
                        Remove-PSDrive -Name "PSHome" -Force -ErrorAction SilentlyContinue| Out-Null
                        New-PSDrive -Name "PSHome" -PSProvider FileSystem -Root  ("\\" + $FS.Server + "\" + $FS.AdminShare + "\" + $User + "$") -Credential $Settings.AdminCred  -ErrorAction SilentlyContinue | Out-Null
                        if (Test-Path -Path "PSHome:\") {
                            $Settings.HomeDrive = ("\\" + $FS.Server + "\" + $FS.AdminShare + "\" + $User + "$")
                            break
                        }
                    }
                }
            #Map Home drive to Archive if none exists. 
            If (-Not (Test-Path "PSHome:\")) {
                    If (-Not (Test-Path "PSHomeRoot:\")) {
                        New-PSDrive -Name "PSHomeRoot" -PSProvider FileSystem -Root  $Settings.HomeDriveArchive  -Credential $Settings.AdminCred  -ErrorAction SilentlyContinue | Out-Null
                    }
                    If (-Not (Test-Path -Path ("PSHomeRoot:\" + $User))){ 
                        New-Item -ItemType Directory -Path ( "PSHomeRoot:\" + $User)
                    }
                    $Settings.HomeDrive = ($Settings.HomeDriveArchive + "\" + $User)
                    New-PSDrive -Name "PSHome" -PSProvider FileSystem -Root  ($Settings.HomeDriveArchive + "\" + $User) -Credential $Settings.AdminCred  -ErrorAction SilentlyContinue | Out-Null
                }
                #endregon Find Home Drive
                #Set Description to the new value
                if ($ADUser.Description -notmatch $Settings.DisabledPrepend) {
                    Set-ADUser -Identity $ADUser -Description ($Settings.DisabledPrepend + " " + $ADUser.Description) -Credential $Settings.AdminCred
                }
                #region Save Group list to Disabled Users in SharePoint
                #Save to Desktop
                $groups = $ADUser.memberof 
                $groups = @()
                ForEach ($Mo in $ADUser.memberof) {
                    $groups += (Get-ADGroup -Identity $Mo).sAMAccountName
                }

                $groups | Out-File ($Settings.Desktop + "\" + $user + ".txt") -Force
                If ((Get-content -Path ($Settings.Desktop + "\" + $user + ".txt")).Length -gt 1) {
                    If(Test-Path "PSHome:\") {
                        Copy-Item -Path ($Settings.Desktop + "\" + $user + ".txt") -Destination ("PSHome:\" + $user + "_AD_Groups.txt")
                    }
                    #Save to SharePoint
                    Add-PnPFile -Path ($Settings.Desktop +"\" + $user + ".txt") -Folder $Settings.SPWebFolder
                    If (Find-PnPFile -Folder $Settings.SPWebFolder -Match ($user + ".txt")) {
                        remove-item -Force  ($Settings.Desktop +"\" + $user + ".txt")
                    }Else{
                        Write-Warning ("User " + $User + ".txt file failed to upload to SharePoint. Please Upload from: " + $Settings.Desktop) 
                        [System.Windows.Forms.MessageBox]::Show("User " + $User + ".txt file failed to upload to SharePoint. Please Upload from: " + $Settings.Desktop , "Error")
                    }
                    #endregion Save Group list to Disabled Users in SharePoint

                    #region Remove Groups
                    If ($groups) {
                       write-host ("Removing " + $user + " from groups.")
                        $groups |Where-Object {$_ -notin $Settings.ExcludeGroups} | Remove-ADGroupMember -Members $user -Confirm:$false -Credential  $Settings.AdminCred
                    }
                    #endregion Remove Groups
                    
                }Else{
                    [System.Windows.Forms.MessageBox]::Show("User " + $User + " please check groups are in " + $Settings.Desktop + "\" + $user + ".txt." , "Error")
                }
                
                #region Disable User
                If ($ADUser.enabled) {
                   write-host ("Disabling account: " + $User)
                    Disable-ADAccount -Identity $User -Credential  $Settings.AdminCred
                }
                #endregion Disable User

                #region Move User
                If (-Not $ADUser.distinguishedName -match $Settings.DisabledOU) {
                   write-host ("Moving " + $User + " to " +  $Settings.DisabledOU)
                    Move-ADObject -Identity $ADUser.DistinguishedName -TargetPath  $Settings.DisabledOU -Credential  $Settings.AdminCred
                }
                #endregion Move User

                #region Export On-Prem Mailbox
                If ($Settings.HomeDrive -and $Settings.ExportMail.Checked){
                    Export-Mail -User $User -Archive $Settings.HomeDrive -Disable -Credential  $Settings.AdminCred
                }
                #endregion Export On-Prem Mailbox

                #region Home Drive Archive
                If ($Settings.ArchiveHome.Checked) {
                    $ArchiveName = (($User  -replace " ","_") + ".7z")
                    $ArchiveName = Remove-InvalidFileNameChars -Name $ArchiveName
                   write-host ("Backing up Homedrive " + $HomeDrive)
                   write-host ("Creating archive " + ($Settings.HomeDriveArchive + "\" + $ArchiveName) + ". Please wait . . .")
                    Compress-7Zip -Path $Settings.HomeDrive -ArchiveFileName ($Settings.HomeDriveArchive + "\" + $ArchiveName) -Format SevenZip -CompressionLevel Ultra 
                }
                #endregion Home Drive Archive
                #region OoO and Forward
                If ($null -ne $ADUser.Manager) {
                    $UsersManager = (Get-ADUser $ADUser.Manager -Properties *)
                    If ($UsersManager) {
                        $OoOMesage = ( $OoO_Pre + $UsersManager.FirstName + " at " +  $UsersManager.Mail + $OoO_Post)
                    }Else{
                        $OoOMesage =( $OoO_Pre + "their manager" + $OoO_Post)

                    }
                    If ($Settings.OutOfOffice.Text) {
                        $OoOMesage = $Settings.OutOfOffice.Text
                        If ($Settings.OoOPrepend.Checked) {
                            $OoOMesage = ($ADUser.FirstName + " " + $OoOMesage)
                        }
                        if($Settings.OoOMan) {
                            $UsersManager = get-user $ADUser.Manager -ResultSize Unlimited
                            If (-Not $UsersManager) {
                                $OoOMesage = ( $OoOMesage + " For any business related needs please e-mail " + $UsersManager.FirstName + " at " +  $UsersManager.WindowsEmailAddress + ".")
                            }
                        }
                        If ($ADUser.targetAddress -match "onmicrosoft.com" -and $ADUser.Mail){
                            #"Online"
                            #Set OoO    
                            Set-MailboxAutoReplyConfiguration -Identity $User -AutoReplyState Enabled -InternalMessage $OoOMesage -ExternalMessage  $OoOMesage -ExternalAudience "all"
                                        
                            If ($UsersManager -and $Settings.ForwardMail.text) {
                                Set-Mailbox -Identity $User -DeliverToMailboxAndForward $false -ForwardingSMTPAddress $null
                                Set-Mailbox -Identity $User -DeliverToMailboxAndForward $true -ForwardingSMTPAddress  $UsersManager.Mail
                            }
                            
                        }Else{
                            #"On-Premise"
                            #Set OoO    
                            Set-EPMailboxAutoReplyConfiguration -Identity $User -AutoReplyState Enabled -InternalMessage $OoOMesage -ExternalMessage  $OoOMesage -ExternalAudience "all"
                                            
                            If ($UsersManager -and $Settings.ForwardMail.text) {
                                Set-EPMailbox -Identity $User -DeliverToMailboxAndForward $false -ForwardingSMTPAddress $null
                                Set-EPMailbox -Identity $User -DeliverToMailboxAndForward $true -ForwardingSMTPAddress  $UsersManager.Mail
                            }
                        }
                    }
                }
                #endregion OoO and Forward
            }

            #cleanup
            Remove-PSDrive -Name "PSHomeRoot" -Force -ErrorAction SilentlyContinue | Out-Null
            Remove-PSDrive -Name "PSHome" -Force -ErrorAction SilentlyContinue | Out-Null

        # })
        # $MainpsCmd.Powershell.Runspace = $MainRunspace
        # $MainpsCmd.Handle = $MainpsCmd.Powershell.BeginInvoke()
        
        # While ($MainpsCmd.Handle.IsCompleted -ne $true) {
        #     Start-Sleep -Milliseconds 100
        #     [gc]::collect()
        # }
        
        [gc]::collect()
        $Settings.sw.Stop()
        [gc]::collect()
        $Settings.Form.text = ( $Settings.WindowTitle + " Done. Time: " + (FormatElapsedTime($Settings.sw.Elapsed)) ) 
        #$MainpsCmd.Powershell.EndInvoke($MainpsCmd.Handle)
        [gc]::collect()
        $Settings.Stop.Text = "Exit"
        [gc]::collect()
        #$Settings.Form.Close()
        #[void]$Settings.Form.Close()
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
        $Settings.Form.text                = ( $Settings.WindowTitle + " Cleaning Up. Please Wait . . ." )
       
        
        [void]$Settings.Form.Close()

        #Exit
    } Else {
        [void]$Settings.Form.Close()
    }
}
#############################################################################
#endregion Functions
#############################################################################
#############################################################################
#region Setup Sessions
#############################################################################
# Hide-Console
#Load .Net Classes
#Popup
Add-Type -AssemblyName PresentationCore,PresentationFramework
#Form
Add-Type -AssemblyName System.Windows.Forms
#Password Generation
Add-Type -AssemblyName System.web

[System.Windows.Forms.Application]::EnableVisualStyles()

$Settings.Form                     = New-Object system.Windows.Forms.Form
$Settings.Form.ClientSize          = '450,330'
$Settings.Form.MinimumSize          = '450,330'
$Settings.Form.text                = $Settings.WindowTitle
$Settings.Form.TopMost             = $false
#Show Icon https://stackoverflow.com/questions/53376491/powershell-how-to-embed-icon-in-powershell-gui-exe
If ($iconBase64) {
    $iconBytes       = [Convert]::FromBase64String($iconBase64)
    $stream          = New-Object IO.MemoryStream($iconBytes, 0, $iconBytes.Length)
    $stream.Write($iconBytes, 0, $iconBytes.Length);
    $Settings.Form.icon       = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -Argument $stream).GetHIcon())
}

$Settings.Admin_User_Label                = New-Object system.Windows.Forms.Label
$Settings.Admin_User_Label.text           = "Administrator Username:"
$Settings.Admin_User_Label.AutoSize       = $true
$Settings.Admin_User_Label.width          = 25
$Settings.Admin_User_Label.height         = 10
$Settings.Admin_User_Label.location       = New-Object System.Drawing.Point(10,10)
$Settings.Admin_User_Label.Font           = 'Microsoft Sans Serif,10'
$Settings.Admin_User_Label.Anchor         = [System.Windows.Forms.AnchorStyles]::Top -bor  [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right


$Settings.Admin_User                      = New-Object system.Windows.Forms.TextBox
$Settings.Admin_User.multiline            = $false
$Settings.Admin_User.width                = 194
$Settings.Admin_User.height               = 20
$Settings.Admin_User.location             = New-Object System.Drawing.Point(170,10)
$Settings.Admin_User.Font                 = 'Microsoft Sans Serif,10'
$Settings.Admin_User.Text                 = ($env:USERDNSDOMAIN + "\DA-" + $env:USERNAME)
$Settings.Admin_User.Anchor               = [System.Windows.Forms.AnchorStyles]::Top -bor  [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

$Settings.Admin_Pass_Label                = New-Object system.Windows.Forms.Label
$Settings.Admin_Pass_Label.text           = "Administrator Password:"
$Settings.Admin_Pass_Label.AutoSize       = $true
$Settings.Admin_Pass_Label.width          = 25
$Settings.Admin_Pass_Label.height         = 10
$Settings.Admin_Pass_Label.location       = New-Object System.Drawing.Point(10,40)
$Settings.Admin_Pass_Label.Font           = 'Microsoft Sans Serif,10'
$Settings.Admin_Pass_Label.Anchor         = [System.Windows.Forms.AnchorStyles]::Top -bor  [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

$Settings.Admin_Pass                      = New-Object system.Windows.Forms.TextBox
$Settings.Admin_Pass.multiline            = $false
$Settings.Admin_Pass.width                = 194
$Settings.Admin_Pass.height               = 20
$Settings.Admin_Pass.location             = New-Object System.Drawing.Point(170,40)
$Settings.Admin_Pass.Font                 = 'Microsoft Sans Serif,10'
$Settings.Admin_Pass.Text                 = ""
$Settings.Admin_Pass.PasswordChar         = "*"
$Settings.Admin_Pass.Anchor               = [System.Windows.Forms.AnchorStyles]::Top -bor  [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

$Settings.Disable_User_Label              = New-Object system.Windows.Forms.Label
$Settings.Disable_User_Label.text         = "Disable User:"
$Settings.Disable_User_Label.AutoSize     = $true
$Settings.Disable_User_Label.width        = 25
$Settings.Disable_User_Label.height       = 10
$Settings.Disable_User_Label.location     = New-Object System.Drawing.Point(10,70)
$Settings.Disable_User_Label.Font         = 'Microsoft Sans Serif,10'
$Settings.Disable_User_Label.Anchor       = [System.Windows.Forms.AnchorStyles]::Top -bor  [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

$Settings.Disable_User                    = New-Object system.Windows.Forms.TextBox
$Settings.Disable_User.multiline          = $false
$Settings.Disable_User.width              = 194
$Settings.Disable_User.height             = 20
$Settings.Disable_User.location           = New-Object System.Drawing.Point(170,70)
$Settings.Disable_User.Font               = 'Microsoft Sans Serif,10'
$Settings.Disable_User.Enabled            = $true
$Settings.Disable_User.Anchor             = [System.Windows.Forms.AnchorStyles]::Top -bor  [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right 
# $Settings.Disable_User.text               = ""

$Settings.Disable_Browse                  = New-Object system.Windows.Forms.Button
$Settings.Disable_Browse.text             = "Browse..."
$Settings.Disable_Browse.width            = 70
$Settings.Disable_Browse.height           = 25
$Settings.Disable_Browse.location         = New-Object System.Drawing.Point(370,70)
$Settings.Disable_Browse.Font             = 'Microsoft Sans Serif,10'
$Settings.Disable_Browse.Anchor           = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right 
$Settings.Disable_Browse.Enabled          = $false
$Settings.Disable_Browse.Visible          = $False

$Settings.ForwardMail_Label               = New-Object system.Windows.Forms.Label
$Settings.ForwardMail_Label.text          = "Forward email to:"
$Settings.ForwardMail_Label.AutoSize      = $true
$Settings.ForwardMail_Label.width         = 25
$Settings.ForwardMail_Label.height        = 10
$Settings.ForwardMail_Label.location      = New-Object System.Drawing.Point(10,100)
$Settings.ForwardMail_Label.Font          = 'Microsoft Sans Serif,10'
$Settings.ForwardMail_Label.Anchor        = [System.Windows.Forms.AnchorStyles]::Top -bor  [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right 

$Settings.ForwardMail                     = New-Object system.Windows.Forms.TextBox
$Settings.ForwardMail.multiline           = $false
$Settings.ForwardMail.width               = 194
$Settings.ForwardMail.height              = 20
$Settings.ForwardMail.location            = New-Object System.Drawing.Point(170,100)
$Settings.ForwardMail.Font                = 'Microsoft Sans Serif,10'
$Settings.ForwardMail.Enabled             = $true
$Settings.ForwardMail.Anchor              = [System.Windows.Forms.AnchorStyles]::Top -bor  [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
# $Settings.ForwardMail.text                = ""

$Settings.OutOfOffice_Label               = New-Object system.Windows.Forms.Label
$Settings.OutOfOffice_Label.text          = "Out of Office Message:"
$Settings.OutOfOffice_Label.AutoSize      = $true
$Settings.OutOfOffice_Label.width         = 25
$Settings.OutOfOffice_Label.height        = 10
$Settings.OutOfOffice_Label.location      = New-Object System.Drawing.Point(10,130)
$Settings.OutOfOffice_Label.Font          = 'Microsoft Sans Serif,10'
$Settings.OutOfOffice_Label.Anchor        = [System.Windows.Forms.AnchorStyles]::Top -bor  [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right 

$Settings.OutOfOffice                     = New-Object system.Windows.Forms.TextBox
$Settings.OutOfOffice.multiline           = $false
$Settings.OutOfOffice.width               = 194
$Settings.OutOfOffice.height              = 90
$Settings.OutOfOffice.location            = New-Object System.Drawing.Point(170,130)
$Settings.OutOfOffice.Font                = 'Microsoft Sans Serif,10'
$Settings.OutOfOffice.Enabled             = $true
$Settings.OutOfOffice.multiline           = $true
$Settings.OutOfOffice.ScrollBars          = "Both" 
$Settings.OutOfOffice.Anchor              = [System.Windows.Forms.AnchorStyles]::Top -bor  [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
$Settings.OutOfOffice.text                = "The persion you are contacting is no longer with PLS Financial Services."

$Settings.ExportMail                      = New-Object System.Windows.Forms.Checkbox 
$Settings.ExportMail.Text                 = "Export Mail"
$Settings.ExportMail.width                = 120
$Settings.ExportMail.height               = 20
$Settings.ExportMail.Location             = New-Object System.Drawing.Size(10,230) 
$Settings.ExportMail.Font                 = 'Microsoft Sans Serif,10'
$Settings.ExportMail.Checked              = $False
$Settings.ExportMail.Anchor               = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left 

$Settings.ArchiveHome                     = New-Object System.Windows.Forms.Checkbox 
$Settings.ArchiveHome.Text                = "Archive Home"
$Settings.ArchiveHome.width               = 120
$Settings.ArchiveHome.height              = 20
$Settings.ArchiveHome.Location            = New-Object System.Drawing.Size(170,230) 
$Settings.ArchiveHome.Font                = 'Microsoft Sans Serif,10'
$Settings.ArchiveHome.Checked             = $False
$Settings.ArchiveHome.Anchor              = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left 

$Settings.OoOMan                          = New-Object System.Windows.Forms.Checkbox 
$Settings.OoOMan.Text                     = "Append manager e-mail to Out of Office."
$Settings.OoOMan.width                    = 320
$Settings.OoOMan.height                   = 20
$Settings.OoOMan.Location                 = New-Object System.Drawing.Size(10,250) 
$Settings.OoOMan.Font                     = 'Microsoft Sans Serif,10'
$Settings.OoOMan.Checked                  = $False
$Settings.OoOMan.Anchor                   = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left 

$Settings.OoOPrepend                       = New-Object System.Windows.Forms.Checkbox 
$Settings.OoOPrepend.Text                  = "Prepend User's First name to Out of Office."
$Settings.OoOPrepend.width                 = 320
$Settings.OoOPrepend.height                = 20
$Settings.OoOPrepend.Location              = New-Object System.Drawing.Size(10,250) 
$Settings.OoOPrepend.Font                  = 'Microsoft Sans Serif,10'
$Settings.OoOPrepend.Checked               = $False
$Settings.OoOPrepend.Anchor                = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left 

$Settings.Stop                             = New-Object system.Windows.Forms.Button
$Settings.Stop.text                        = "Exit"
$Settings.Stop.width                       = 70
$Settings.Stop.height                      = 25
$Settings.Stop.location                    = New-Object System.Drawing.Point(220,290)
$Settings.Stop.Font                        = 'Microsoft Sans Serif,10'
$Settings.Stop.Anchor                      = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right 

$Settings.Start                            = New-Object system.Windows.Forms.Button
$Settings.Start.text                       = "Disable"
$Settings.Start.width                      = 70
$Settings.Start.height                     = 25
$Settings.Start.location                   = New-Object System.Drawing.Point(290,290)
$Settings.Start.Font                       = 'Microsoft Sans Serif,10'
$Settings.Start.Anchor                     = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right 


$Settings.Form.controls.AddRange(@(
$Settings.Admin_User_Label,
$Settings.Admin_User,
$Settings.Admin_Pass_Label,
$Settings.Admin_Pass,
$Settings.Disable_User_Label,
$Settings.Disable_User,
$Settings.Disable_Browse,
$Settings.ForwardMail_Label,
$Settings.ForwardMail,
$Settings.OutOfOffice_Label,
$Settings.OutOfOffice,
$Settings.ExportMail,
$Settings.ArchiveHome,
$Settings.OoOMan,
$Settings.OoOPrepend,
$Settings.Stop,
$Settings.Start
))


#############################################################################
#endregion Setup Sessions
#############################################################################
#############################################################################
#region Main 
#############################################################################

$Settings.Disable_Browse.Add_Click({ Browse_File })
$Settings.Start.Add_Click({ Start_Work })
$Settings.Stop.Add_Click({ Stop_Work })


[void]$Settings.Form.ShowDialog()

#############################################################################
#endregion Main
#############################################################################
