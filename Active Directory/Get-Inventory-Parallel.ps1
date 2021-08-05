<#
.SYNOPSIS
  Name: Get-Inventory-Parallel.ps1
  The purpose of this script is to create a simple inventory.
  
.DESCRIPTION
  This is a simple script to retrieve all computer objects in Active Directory and then connect
  to each one and gather basic hardware information using Cim. The information includes Manufacturer,
  Model,Serial Number, CPU, RAM, Disks, Operating System, Sound Deivces and Graphics Card Controller.

.RELATED LINKS
  https://www.sconstantinou.com
  https://gallery.technet.microsoft.com/scriptcenter/PowerShell-Hardware-f99336f6

.NOTES
  Version 2.0.17

  Updated:      01-06-2018        - Replaced Get-CimInstance cmdlet with Get-CimInstance
                                  - Added Serial Number Information
                                  - Added Sound Device Information
                                  - Added Video Controller Information
                                  - Added option to send CSV file through email
                                  - Added parameters to enable email function option
                03/21/2019        - Added IP address
                                  - Added Software
                                  - Added Installed Roles
                                  - Added Installed Features
                09/24/2019        - Added DNS Server
                                  - Added DNS Search Suffix
                01/27/2020        - Added Certificates Experation info for Lisening connections.
                                  - Added Jobs to process computers in parallel
                11/09/2020        - Added Services and IIS info   
                11/10/2020        - Filtered disabled computer
                                  - Cleanup errors 
                                  - Reduce AD Calls
                                  - Fix video bug
                11/11/2020        - Cleaned up names
                                  - Get DNS Record IP
                                  - Ping by FQDN
                                  - Fixed Services Name
                                  - Fixed Drive info
                11/16/2020        - Fixed DNS Records null issue
                                  - Formatting fixes
                12/03/2020        - Added ablity to remove SMB1 protocal
                                  - Added Check for Reboot needed
                                  - Fixed Formatting
                12/05/2020        - Added PSVersion Fixes for SMB1
                12/08/2020        - Moved DN to end of list and more fixed for SMB1
                12/10/2020        - SMB1 detection fixed.  
                01/06/2021        - Added Hit Count to IIS sites.  
                01/07/2021        - Removed Hit Count as it caused computers to lock up
                03/24/2021        - IIS log memory work around
                04/07/2021        - Added SQL databases, FortiClient Version,LanDesk Agent detection.
                06/21/2021        - Remove AD module dependency in threads. Fix Errors.
                
  Release Date: 10-02-2018
   
  Author: Stephanos Constantinou

.EXAMPLES
  Get-Inventory-Parallel.ps1
  Find the output under Logs in the script directory

 Get-Inventory-Parallel.ps1 -Email -Recipients user1@domain.com
  Find the output under Logs in the script directory and an email will be sent
  also to user1@domain.com

  Get-Inventory-Parallel.ps1 -Email -Recipients user1@domain.com,user2@domain.com
  Find the output under Logs in the script directory and an email will be sent
  also to user1@domain.com and user2@domain.com
#>

Param(
    [switch]$Email = $false,
    [string]$Recipients = $null,
    [String]$OutputFolder = ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\Logs"),
    [Int]$MaxJobs = 120,
    [Int]$SleepTime = 5,
    [Int]$TimeOut = 90,
    [Int]$LogHistory = -90,
    [Int]$MaxMemoryUseage = 80,
    [string]$WSO = "",
    [switch]$InstallWMF,
    [switch]$RemoveSMB1
)
$LogFile = ($OutputFolder + "\" + `
           ($MyInvocation.MyCommand.Name -replace ".ps1","") + "_" + `
		   $env:computername + "_" + `
           (Get-Date -format yyyyMMdd-hhmm) + ".log")
$ScriptVersion = "2.0.19"
$sw = [Diagnostics.Stopwatch]::StartNew()
$Inventory =  [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
$ScanCount = 1
$JobCount = 1

$RemoveServices=@(
    "App Readiness - Manual",
    "Application Experience - Manual",
    "Application Identity - Manual",
    "Application Information - Manual",
    "Application Layer Gateway Service - Manual",
    "Application Management - Manual",
    "AppX Deployment Service (AppXSVC) - Manual",
    "Background Intelligent Transfer Service - Manual",
    "Background Tasks Infrastructure Service - Auto",
    "Base Filtering Engine - Auto",
    "Certificate Propagation - Manual",
    "CNG Key Isolation - Manual",
    "COM+ Event System - Auto",
    "COM+ System Application - Manual",
    "Credential Manager - Manual",
    "Cryptographic Services - Auto",
    "DCOM Server Process Launcher - Auto",
    "Device Association Service - Manual",
    "Device Install Service - Manual",
    "Device Setup Manager - Manual",
    "DHCP Client - Auto",
    "Diagnostic Policy Service - Auto",
    "Diagnostic Service Host - Manual",
    "Diagnostic System Host - Manual",
    "Diagnostics Tracking Service - Auto",
    "Distributed Link Tracking Client - Manual",
    "Distributed Transaction Coordinator - Auto",
    "DNS Client - Auto",
    "Encrypting File System (EFS) - Manual",
    "Extensible Authentication Protocol - Manual",
    "FortiClient Service Scheduler - Auto",
    "Function Discovery Provider Host - Manual",
    "Function Discovery Resource Publication - Manual",
    "Group Policy Client - Auto",
    "Human Interface Device Service - Manual",
    "Hyper-V Data Exchange Service - Manual",
    "Hyper-V Guest Service Interface - Manual",
    "Hyper-V Guest Shutdown Service - Manual",
    "Hyper-V Heartbeat Service - Manual",
    "Hyper-V Remote Desktop Virtualization Service - Manual",
    "Hyper-V Time Synchronization Service - Manual",
    "Hyper-V Volume Shadow Copy Requestor - Manual",
    "IKE and AuthIP IPsec Keying Modules - Auto",
    "Intel Local Scheduler Service - Auto",
    "Intel PDS - Auto",
    "Interactive Services Detection - Manual",
    "Internet Connection Sharing (ICS) - Disabled",
    "Internet Explorer ETW Collector Service - Manual",
    "IP Helper - Auto",
    "IPsec Policy Agent - Manual",
    "LANDESK Remote Control Service - Auto",
    "LANDesk Targeted Multicast - Auto",
    "LANDesk(R) Extended device discovery service - Manual",
    "LANDesk(R) Management Agent - Auto",
    "LANDesk(R) Software Monitoring Service - Auto",
    "Link-Layer Topology Discovery Mapper - Manual",
    "Local Session Manager - Auto",
    "Microsoft iSCSI Initiator Service - Manual",
    "Microsoft Software Shadow Copy Provider - Manual",
    "Microsoft Storage Spaces SMP - Manual",
    "Network Connections - Manual",
    "Network Connectivity Assistant - Manual",
    "Network List Service - Manual",
    "Network Location Awareness - Auto",
    "Network Store Interface Service - Auto",
    "Optimize drives - Manual",
    "Performance Counter DLL Host - Manual",
    "Performance Logs & Alerts - Manual",
    "Plug and Play - Manual",
    "Portable Device Enumerator Service - Manual",
    "Print Spooler - Disabled",
    "Printer Extensions and Notifications - Manual",
    "Problem Reports and Solutions Control Panel Support - Manual",
    "Quest Rapid Recovery Agent Service - Auto",
    "Remote Access Auto Connection Manager - Manual",
    "Remote Access Connection Manager - Manual",
    "Remote Desktop Configuration - Manual",
    "Remote Desktop Services - Manual",
    "Remote Desktop Services UserMode Port Redirector - Manual",
    "Remote Procedure Call (RPC) - Auto",
    "Remote Procedure Call (RPC) Locator - Manual",
    "Resultant Set of Policy Provider - Disabled",
    "Routing and Remote Access - Disabled",
    "RPC Endpoint Mapper - Auto",
    "Secondary Logon - Manual",
    "Secure Socket Tunneling Protocol Service - Manual",
    "Security Accounts Manager - Auto",
    "Server - Auto",
    "Shell Hardware Detection - Auto",
    "Smart Card - Auto",
    "Smart Card Device Enumeration Service - Manual",
    "Smart Card Removal Policy - Manual",
    "Software Protection - Auto",
    "Special Administration Console Helper - Manual",
    "Spot Verifier - Manual",
    "SSDP Discovery - Disabled",
    "Storage Tiers Management - Manual",
    "Superfetch - Manual",
    "System Event Notification Service - Auto",
    "System Events Broker - Auto",
    "Task Scheduler - Auto",
    "TCP/IP NetBIOS Helper - Auto",
    "Telephony - Manual",
    "Themes - Disabled",
    "Thread Ordering Server - Manual",
    "UPnP Device Host - Disabled",
    "User Access Logging Service - Auto",
    "User Profile Service - Auto",
    "Virtual Disk - Manual",
    "VMware Alias Manager and Ticket Service - Auto",
    "VMware Snapshot Provider - Manual",
    "VMware Tools - Auto",
    "Volume Shadow Copy - Manual",
    "Windows Audio - Manual",
    "Windows Audio Endpoint Builder - Manual",
    "Windows Color System - Manual",
    "Windows Connection Manager - Auto",
    "Windows Driver Foundation - User-mode Driver Framework - Manual",
    "Windows Encryption Provider Host Service - Manual",
    "Windows Error Reporting Service - Manual",
    "Windows Event Collector - Manual",
    "Windows Event Log - Auto",
    "Windows Firewall - Auto",
    "Windows Font Cache Service - Auto",
    "Windows Installer - Manual",
    "Windows Management Instrumentation - Auto",
    "Windows Modules Installer - Manual",
    "Windows Remote Management (WS-Management) - Auto",
    "Windows Store Service (WSService) - Manual",
    "Windows Time - Manual",
    "Windows Update - Manual",
    "WinHTTP Web Proxy Auto-Discovery Service - Manual",
    "Wired AutoConfig - Manual",
    "WMI Performance Adapter - Manual",
    "Workstation - Auto"
)
[string[]]$KillProcesses=@(
    "TrustedInstaller",
    "fcappdb",
    "FCDBLog",
    "fmon",
    "FortiESNAC",
    "FortiProxy"
)
function FormatElapsedTime($ts) {
  #https://stackoverflow.com/questions/3513650/timing-a-commands-execution-in-powershell
  $elapsedTime = ""
  if ( $ts.Hours -gt 0 ){
      $elapsedTime = [string]::Format( "{0:00} hours {1:00} min. {2:00}.{3:00} sec.", $ts.Hours, $ts.Minutes, $ts.Seconds, $ts.Milliseconds / 10 );
  }else{
      if ( $ts.Minutes -gt 0 ){
          $elapsedTime = [string]::Format( "{0:00} min. {1:00}.{2:00} sec.", $ts.Minutes, $ts.Seconds, $ts.Milliseconds / 10 );
      }else{
          $elapsedTime = [string]::Format( "{0:00}.{1:00} sec.", $ts.Seconds, $ts.Milliseconds / 10 );
      }

      if ($ts.Hours -eq 0 -and $ts.Minutes -eq 0 -and $ts.Seconds -eq 0){
          $elapsedTime = [string]::Format("{0:00} ms.", $ts.Milliseconds);
      }
      if ($ts.Milliseconds -eq 0){
          $elapsedTime = [string]::Format("{0} ms", $ts.TotalMilliseconds);
      }
  }
  return $elapsedTime
}


If (-Not [string]::IsNullOrEmpty($LogFile)) {
    try {
        Stop-transcript -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    } catch {

    }
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

if (Get-Job -name *)
{
    Write-Verbose "Removing old jobs."
    Get-Job -name * -ErrorAction SilentlyContinue| Remove-Job -Force -ErrorAction SilentlyContinue | Out-Null
}

$AllComputers = Get-ADComputer -Filter * -Properties Name,Enabled,OperatingSystem,OperatingSystemServicePack,OperatingSystemVersion,LastLogonDate,description,dNSHostName
$AllComputersNames = $AllComputers.Name
$AllComputersNamesCount = $AllComputersNames.Count
Write-Progress  -Id 0 -Activity "Active Directory Computers Count: $AllComputersNamesCount" -Status "Starting Scan" -Percentcomplete (0)
Foreach ($ComputerName in $AllComputersNames) {
    #Write-Progress  -Id 0 -Activity ("Processing Computer") -Status ("( " + $count + "\" + $AllComputersNames.count + "): " + $CurrentComputer) -percentComplete ($count / $AllComputersNames.count*100)
    #Write-Host ("( " + $count + "\" + $Computers.count + ") Computer: " + $ComputerName)
    If ($null -ne $ComputerName) {
        $ComputerOS = $AllComputers | Where-object {$_.Name -eq $ComputerName}
        Write-Verbose "Starting job ($((Get-Job -name * | Measure-Object).Count+1)/$MaxJobs) for $ComputerName."
        Start-Job -Name  $ComputerName -ArgumentList $ComputerName,$ComputerOS,$RemoveServices,$RemoveSMB1,$InstallWMF,$LogHistory,$KillProcesses -ScriptBlock{
            param
            (
                $ComputerName=$ComputerName,
                $ComputerOS=$ComputerOS,
                $RemoveServices=$RemoveServices,
                $RemoveSMB1=$RemoveSMB1,
                $InstallWMF=$InstallWMF,
                $LogHistory=$LogHistory,
                $KillProcesses=$KillProcesses
            )

            function Get-CheckUrl ()
            {
                [CmdletBinding()]	
                param(
                    [parameter(Mandatory=$true)][string]$url,
                    [int]$timeoutMilliseconds = 10000,
                    [int]$MinimumCertAgeDays = 90
                )
                #source: https://stackoverflow.com/questions/39253055/powershell-script-to-get-certificate-expiry-for-a-website-remotely-for-multiple
                [string]$details = $null
                #Ignore Certs
                #https://blog.ukotic.net/2017/08/15/could-not-establish-trust-relationship-for-the-ssltls-invoke-webrequest/
                if (-not ([System.Management.Automation.PSTypeName]'ServerCertificateValidationCallback').Type)
                {
                $certCallback = @"
    using System;
    using System.Net;
    using System.Net.Security;
    using System.Security.Cryptography.X509Certificates;
    public class ServerCertificateValidationCallback
    {
        public static void Ignore()
        {
            if(ServicePointManager.ServerCertificateValidationCallback ==null)
            {
                ServicePointManager.ServerCertificateValidationCallback += 
                    delegate
                    (
                        Object obj, 
                        X509Certificate certificate, 
                        X509Chain chain, 
                        SslPolicyErrors errors
                    )
                    {
                        return true;
                    };
            }
        }
    }
"@
                Add-Type $certCallback
                }
                [ServerCertificateValidationCallback]::Ignore()
                #disabling the cert validation check. This is what makes this whole thing work with invalid certs...
                #[Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

                #Start Request
                $req = [Net.HttpWebRequest]::Create($url)
                # $req.ServerCertificateValidationCallback = [ServerCertificateValidationCallback]::Ignore()
                $req.Timeout = $timeoutMilliseconds
                $req.AllowAutoRedirect = $false
                try 
                {
                    # GET WEB RESPONSE
                    $req.GetResponse().Dispose() | Out-Null
                    if ($null -eq $req.ServicePoint.Certificate) {$details = "No certificate in use for connection"}
                } 
                catch 
                {
                    $details = "Exception while checking URL $url`: $_ "
                }
                if ( $null -eq $details -or $details -eq "")
                {
                    [datetime]$expiration = [System.DateTime]::Parse($req.ServicePoint.Certificate.GetExpirationDateString())
                    [int]$certExpiresIn = ($expiration - $(get-date)).Days
                    $certName = $req.ServicePoint.Certificate.GetName()
                    $certPublicKeyString = $req.ServicePoint.Certificate.GetPublicKeyString()
                    $certSerialNumber = $req.ServicePoint.Certificate.GetSerialNumberString()
                    $certThumbprint = $req.ServicePoint.Certificate.GetCertHashString()
                    $certEffectiveDate = $req.ServicePoint.Certificate.GetEffectiveDateString()
                    $certIssuer = $req.ServicePoint.Certificate.GetIssuerName()
                    if ($certExpiresIn -gt $minimumCertAgeDays) {
                        $returnData += new-object psobject -property  @{Url = $url; CheckResult = "OK"; CertExpiresInDays = [int]$certExpiresIn; ExpirationOn = [datetime]$expiration; CertName = $certname; Details = $details}
                    }else{
                        $details = ""
                        $details += "Cert for site $url expires in $certExpiresIn days [on $expiration]`n"
                        $details += "Threshold is $minimumCertAgeDays days. Check details:`n"
                        $details += "Cert name: $certName`n"
                        $details += "Cert public key: $certPublicKeyString`n"
                        $details += "Cert serial number: $certSerialNumber`n"
                        $details += "Cert thumbprint: $certThumbprint`n"
                        $details += "Cert effective date: $certEffectiveDate`n"
                        $details += "Cert issuer: $certIssuer"
                        $returnData += new-object psobject -property  @{Url = $url; CheckResult = "WARNING"; CertExpiresInDays = [int]$certExpiresIn; ExpirationOn = [datetime]$expiration; CertName = $certname; Details = $details}
                    }
                    Remove-Variable expiration
                    Remove-Variable certExpiresIn
                }else{
                    $returnData += new-object psobject -property  @{Url = $url; CheckResult = "ERROR"; CertExpiresInDays = $null; ExpirationOn = $null; CertName = $certname; Details = $details}
                }

                Remove-Variable req
                #Remove-Variable res
                return $returnData
            }
            Function Get-DomainComputer {
                <#
                .SYNOPSIS
                    The Get-DomainComputer function allows you to get information from an Active Directory Computer object using ADSI.
                
                .DESCRIPTION
                    The Get-DomainComputer function allows you to get information from an Active Directory Computer object using ADSI.
                    You can specify: how many result you want to see, which credentials to use and/or which domain to query.
                
                .PARAMETER ComputerName
                    Specifies the name(s) of the Computer(s) to query
                
                .PARAMETER SizeLimit
                    Specifies the number of objects to output. Default is 100.
                
                .PARAMETER DomainDN
                    Specifies the path of the Domain to query.
                    Examples:     "FX.LAB"
                                "DC=FX,DC=LAB"
                                "Ldap://FX.LAB"
                                "Ldap://DC=FX,DC=LAB"
                .PARAMETER Properties
                    Specifies extra fields to return. 
                .PARAMETER Credential
                    Specifies the alternate credentials to use.
                
                .EXAMPLE
                    Get-DomainComputer
                
                    This will show all the computers in the current domain
                
                .EXAMPLE
                    Get-DomainComputer -ComputerName "Workstation001"
                
                    This will query information for the computer Workstation001.
                
                .EXAMPLE
                    Get-DomainComputer -ComputerName "Workstation001","Workstation002"
                
                    This will query information for the computers Workstation001 and Workstation002.
                
                .EXAMPLE
                    Get-Content -Path c:\WorkstationsList.txt | Get-DomainComputer
                
                    This will query information for all the workstations listed inside the WorkstationsList.txt file.
                
                .EXAMPLE
                    Get-DomainComputer -ComputerName "Workstation0*" -SizeLimit 10 -Verbose
                
                    This will query information for computers starting with 'Workstation0', but only show 10 results max.
                    The Verbose parameter allow you to track the progression of the script.
                
                .EXAMPLE
                    Get-DomainComputer -ComputerName "Workstation0*" -SizeLimit 10 -Verbose -DomainDN "DC=FX,DC=LAB" -Credential (Get-Credential -Credential FX\Administrator)
                
                    This will query information for computers starting with 'Workstation0' from the domain FX.LAB with the account FX\Administrator.
                    Only show 10 results max and the Verbose parameter allows you to track the progression of the script.
                
                .NOTES
                    NAME:    FUNCT-AD-COMPUTER-Get-DomainComputer.ps1
                    AUTHOR:    Francois-Xavier CAT
                    DATE:    2013/10/26
                    WWW:    www.lazywinadmin.com
                    TWITTER: @lazywinadmin
                
                    VERSION HISTORY:
                    1.0 2013.10.26
                        Initial Version
                    1.1 2021.06.21
                        Added ablity to have extra properties. 
                #>
                
                    [CmdletBinding()]
                    PARAM(
                        [Parameter(
                            ValueFromPipelineByPropertyName=$true,
                            ValueFromPipeline=$true)]
                        [Alias("Computer")]
                        [String[]]$ComputerName,
                
                        [Alias("ResultLimit","Limit")]
                        [int]$SizeLimit='100',
                
                        [Parameter(ValueFromPipelineByPropertyName=$true)]
                        [Alias("Domain")]
                        [String]$DomainDN=$(([adsisearcher]"").Searchroot.path),

                        [Parameter(ValueFromPipelineByPropertyName=$true)]
                        [Alias("Property")]
                        [String[]]$Properties,

                        [Alias("RunAs")]
                        [System.Management.Automation.Credential()]
                        $Credential = [System.Management.Automation.PSCredential]::Empty
                
                    )#PARAM
                
                    PROCESS{
                        IF ($ComputerName){
                            Write-Verbose -Message "One or more ComputerName specified"
                            FOREACH ($item in $ComputerName){
                                TRY{
                                    # Building the basic search object with some parameters
                                    Write-Verbose -Message "COMPUTERNAME: $item"
                                    $Searcher = New-Object -TypeName System.DirectoryServices.DirectorySearcher -ErrorAction 'Stop' -ErrorVariable ErrProcessNewObjectSearcher
                                    $Searcher.Filter = "(&(objectCategory=Computer)(name=$item))"
                                    $Searcher.SizeLimit = $SizeLimit
                                    $Searcher.SearchRoot = $DomainDN
                
                                    #Specify Other Properties to load
                                    # If ($Properties) {
                                    #     foreach ($Property in $Properties) {
                                    #         [void]$Searcher.PropertiesToLoad.Add($Property)
                                    #     }
                                    # }
                                    # Specify a different domain to query
                                    IF ($PSBoundParameters['DomainDN']){
                                        IF ($DomainDN -notlike "LDAP://*") {$DomainDN = "LDAP://$DomainDN"}#IF
                                        Write-Verbose -Message "Different Domain specified: $DomainDN"
                                        $Searcher.SearchRoot = $DomainDN}#IF ($PSBoundParameters['DomainDN'])
                
                                    # Alternate Credentials
                                    IF ($PSBoundParameters['Credential']) {
                                        Write-Verbose -Message "Different Credential specified: $($Credential.UserName)"
                                        $Domain = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList $DomainDN,$($Credential.UserName),$($Credential.GetNetworkCredential().password) -ErrorAction 'Stop' -ErrorVariable ErrProcessNewObjectCred
                                        $Searcher.SearchRoot = $Domain}#IF ($PSBoundParameters['Credential'])
                
                                    # Querying the Active Directory
                                    Write-Verbose -Message "Starting the ADSI Search..."
                                    FOREACH ($Computer in $($Searcher.FindAll())){
                                        Write-Verbose -Message "$($Computer.properties.name)"
                                        If ($Properties) {
                                            $Output = New-Object -TypeName PSObject -ErrorAction 'Continue' -ErrorVariable ErrProcessNewObjectOutputALL -Property @{}
                                            foreach ($Property in $Properties) {
                                                Try {
                                                    Write-Verbose -Message $([string]$Computer.properties.$Property + " ")
                                                    switch ($Property.ToLower()) {
                                                        "lastlogondate" { 
                                                            Add-Member -InputObject $Output -NotePropertyName ($Property) -NotePropertyValue $([datetime]::FromFileTime([int64][string]$computer.Properties.lastlogon).ToString('MM/dd/yyy hh:mm:ss tt')) -Force
                                                         }
                                                         "enabled" {
                                                             If ([int64][string]$computer.Properties.useraccountcontrol -band 2 -eq 0) {
                                                                Add-Member -InputObject $Output -NotePropertyName ($Property) -NotePropertyValue $true -Force
                                                             }Else {
                                                                Add-Member -InputObject $Output -NotePropertyName ($Property) -NotePropertyValue $false -Force
                                                             }
                                                         }
                                                         "operatingsystemservicepack" {
                                                             #Need to find/figure out
                                                            Add-Member -InputObject $Output -NotePropertyName ($Property) -NotePropertyValue $("") -Force
                                                         }
                                                        Default {
                                                            Add-Member -InputObject $Output -NotePropertyName ($Property) -NotePropertyValue $([string]$Computer.properties.($Property.ToLower())) -Force
                                                        }
                                                    }
                                                   
                                                }
                                                Catch {
                                                    Write-Warning -Message ('{0}: {1}' -f $Property, $_.Exception.Message)
                                                }   
                                            }
                                        }Else {
                                            $Output = New-Object -TypeName PSObject -ErrorAction 'Continue' -ErrorVariable ErrProcessNewObjectOutputALL -Property @{
                                                "Name" = $($Computer.properties.name)
                                                "DNShostName"    = $($Computer.properties.dnshostname)
                                                "Description" = $($Computer.properties.description)
                                                "OperatingSystem"=$($Computer.Properties.operatingsystem)
                                                "WhenCreated" = $($Computer.properties.whencreated)
                                                "DistinguishedName" = $($Computer.properties.distinguishedname)}#New-Object
                                        } #If $Properties
                                        return $Output
                                    }#FOREACH $Computer
                
                                    Write-Verbose -Message "ADSI Search completed"
                                }#TRY
                                CATCH{
                                    Write-Warning -Message ('{0}: {1}' -f $item, $_.Exception.Message)
                                    IF ($ErrProcessNewObjectSearcher){Write-Warning -Message "PROCESS BLOCK - Error during the creation of the searcher object"}
                                    IF ($ErrProcessNewObjectCred){Write-Warning -Message "PROCESS BLOCK - Error during the creation of the alternate credential object"}
                                    IF ($ErrProcessNewObjectOutput){Write-Warning -Message "PROCESS BLOCK - Error during the creation of the output object"}
                                }#CATCH
                            }#FOREACH $item
                
                
                        }#IF $ComputerName
                        ELSE {
                            Write-Verbose -Message "No ComputerName specified"
                            TRY{
                                # Building the basic search object with some parameters
                                Write-Verbose -Message "List All object"
                                $Searcher = New-Object -TypeName System.DirectoryServices.DirectorySearcher -ErrorAction 'Stop' -ErrorVariable ErrProcessNewObjectSearcherALL
                                $Searcher.Filter = "(objectCategory=Computer)"
                                $Searcher.SizeLimit = $SizeLimit
                                
                                #Specify Other Properties to load
                                # If ($Properties) {
                                #     foreach ($Property in $Properties) {
                                #         [void]$Searcher.PropertiesToLoad.Add($Property)
                                #     }
                                # }                
                                # Specify a different domain to query
                                IF ($PSBoundParameters['DomainDN']){
                                    $DomainDN = "LDAP://$DomainDN"
                                    Write-Verbose -Message "Different Domain specified: $DomainDN"
                                    $Searcher.SearchRoot = $DomainDN}#IF ($PSBoundParameters['DomainDN'])
                
                                # Alternate Credentials
                                IF ($PSBoundParameters['Credential']) {
                                    Write-Verbose -Message "Different Credential specified: $($Credential.UserName)"
                                    $DomainDN = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList $DomainDN, $Credential.UserName,$Credential.GetNetworkCredential().password -ErrorAction 'Stop' -ErrorVariable ErrProcessNewObjectCredALL
                                    $Searcher.SearchRoot = $DomainDN}#IF ($PSBoundParameters['Credential'])
                
                                # Querying the Active Directory
                                Write-Verbose -Message "Starting the ADSI Search..."
                                FOREACH ($Computer in $($Searcher.FindAll())){
                                    TRY{
                                        Write-Verbose -Message "$($Computer.properties.name)"
                                        If ($Properties) {
                                            Try {
                                                Write-Verbose -Message $([string]$Computer.properties.$Property + " ")
                                                switch ($Property.ToLower()) {
                                                    "lastlogondate" { 
                                                        Add-Member -InputObject $Output -NotePropertyName ($Property) -NotePropertyValue $([datetime]::FromFileTime([int64][string]$computer.Properties.lastlogon).ToString('MM/dd/yyy hh:mm:ss tt')) -Force
                                                     }
                                                     "enabled" {
                                                         If ([int64][string]$computer.Properties.useraccountcontrol -band 2 -eq 0) {
                                                            Add-Member -InputObject $Output -NotePropertyName ($Property) -NotePropertyValue $true -Force
                                                         }Else {
                                                            Add-Member -InputObject $Output -NotePropertyName ($Property) -NotePropertyValue $false -Force
                                                         }
                                                     }
                                                     "operatingsystemservicepack" {
                                                         #Need to find/figure out
                                                        Add-Member -InputObject $Output -NotePropertyName ($Property) -NotePropertyValue $("") -Force
                                                     }
                                                    Default {
                                                        Add-Member -InputObject $Output -NotePropertyName ($Property) -NotePropertyValue $([string]$Computer.properties.($Property.ToLower())) -Force
                                                    }
                                                }
                                               
                                            }
                                            Catch {
                                                Write-Warning -Message ('{0}: {1}' -f $Property, $_.Exception.Message)
                                            }   
                                        }Else {
                                            $Output = New-Object -TypeName PSObject -ErrorAction 'Continue' -ErrorVariable ErrProcessNewObjectOutputALL -Property @{
                                                "Name" = $($Computer.properties.name)
                                                "DNShostName"    = $($Computer.properties.dnshostname)
                                                "Description" = $($Computer.properties.description)
                                                "OperatingSystem"=$($Computer.Properties.operatingsystem)
                                                "WhenCreated" = $($Computer.properties.whencreated)
                                                "DistinguishedName" = $($Computer.properties.distinguishedname)}#New-Object
                                        } #If $Properties
                                        return $Output
                                    }#TRY
                                    CATCH{
                                        Write-Warning -Message ('{0}: {1}' -f $Computer, $_.Exception.Message)
                                        IF ($ErrProcessNewObjectOutputALL){Write-Warning -Message "PROCESS BLOCK - Error during the creation of the output object"}
                                    }
                                }#FOREACH $Computer
                
                                Write-Verbose -Message "ADSI Search completed"
                
                            }#TRY
                
                            CATCH{
                                Write-Warning -Message "Something Wrong happened"
                                IF ($ErrProcessNewObjectSearcherALL){Write-Warning -Message "PROCESS BLOCK - Error during the creation of the searcher object"}
                                IF ($ErrProcessNewObjectCredALL){Write-Warning -Message "PROCESS BLOCK - Error during the creation of the alternate credential object"}
                
                            }#CATCH
                        }#ELSE
                    }#PROCESS
                    END{Write-Verbose -Message "Function Completed"}
                }#function
                
            If ($null -eq $ComputerOS) {
                #Write-Host ("Recording Computer Info: " + $ComputerName)
                #If (-Not (Get-Module | Where-Object {$_.Name -Match "ActiveDirectory"})) {
                    #Write-Host ("Loading Active Directory Plugins") -foregroundcolor "Green"
                   # Import-Module "ActiveDirectory"  -ErrorAction SilentlyContinue
               # } Else {
                    #Write-Host ("Active Directory Plug-ins Already Loaded") -foregroundcolor "Green"
               # }
                #$ComputerOS = Get-ADComputer "$ComputerName" -Properties Name,Enabled,OperatingSystem,OperatingSystemServicePack,OperatingSystemVersion,LastLogonDate,description,dNSHostName -ErrorAction SilentlyContinue
                
                $ComputerOS = Get-DomainComputer -ComputerName "$ComputerName" -Properties Name,Enabled,OperatingSystem,OperatingSystemServicePack,OperatingSystemVersion,LastLogonDate,description,dNSHostName -ErrorAction SilentlyContinue

            }
            $ComputerInfo = New-Object System.Object

            if ($null -ne $ComputerOS.Name) {
                $ComputerInfo | Add-Member  -Force -MemberType NoteProperty -Name "AD Name" -Value $ComputerOS.Name
            }else {
                $ComputerInfo | Add-Member  -Force -MemberType NoteProperty -Name "AD Name" -Value $ComputerName
            }
            If ($null -ne $ComputerOS.dNSHostName) {
                $ComputerDNSIP =  (Resolve-DnsName -Name $ComputerOS.dNSHostName -ErrorAction SilentlyContinue).IPAddress
            }Else {
                $ComputerDNSIP =(Resolve-DnsName -Name $ComputerName -ErrorAction SilentlyContinue).IPAddress
            }
            If ($null -ne $ComputerDNSIP) {
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Name "DNS IP" -Value $ComputerDNSIP
            }else {
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Name "DNS IP" -Value ""
            }
            If ($null -ne $ComputerOS.Enabled) {
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Name "Enabled" -Value $ComputerOS.Enabled
            }Else{
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Name "Enabled" -Value ""
            }
            If ($null -ne $ComputerOS.description) {
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Name "Description" -Value $ComputerOS.description
            }Else{
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Name "Description" -Value ""
            }
            If ($null -ne $ComputerOS.OperatingSystem) {
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Name "Operating System" -Value $ComputerOS.OperatingSystem
            }Else{
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Name "Operating System" -Value ""
            }
            If ($null -ne $ComputerOS.OperatingSystemVersion) {
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Name "Operating System Version" -Value $ComputerOS.OperatingSystemVersion
            }Else{
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Name "Operating System Version" -Value ""
            }
            If ($null -ne $ComputerOS.OperatingSystemServicePack) {
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Name "Operating System Service Pack" -Value $ComputerOS.OperatingSystemServicePack
            }Else{
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Name "Operating System Service Pack" -Value ""
            }
            If ($null -ne $ComputerOS.LastLogonDate) {
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Name "Last Logon Date" -Value $ComputerOS.LastLogonDate
            }Else{
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Name "Last Logon Date" -Value ""
            }

            If (($ComputerOS.Enabled -eq $true)) {
                $ArrComputerIP= (Get-CimInstance -Class Win32_NetworkAdapterConfiguration -ComputerName $ComputerName -Filter 'IPEnabled = True' -ErrorAction SilentlyContinue | Where-object {$_.IPaddress -notlike "169.254.*" -and $_.IPAddress -ne "127.0.0.1" -and $_.IPAddress -notmatch ":"} | Sort-Object -Unique -Property IPAddress| Select-Object IPAddress,Description,DNSServerSearchOrder)
                If ($ArrComputerIP) {
                    Write-verbose ("`tHost is up")
                    #region NIC
                    $ComputerIP=@()
                    $ComputerDNS=@()
                    $ComputerDNSSuffix=@()
                    
                    foreach ($NIC in $ArrComputerIP) {
                        $ComputerIP += (($NIC.Description -replace '[,()]','') + " (" + ($NIC.IPAddress -join " ") + "), ")
                        $ComputerDNS += $NIC.DNSServerSearchOrder
                        $ComputerDNSSuffix += $_.DNSDomainSuffixSearchOrder
                    }
                    $ComputerDNS =  $ComputerDNS | Sort-Object -Unique
                    $ComputerDNSSuffix =  $ComputerDNSSuffix | Sort-Object -Unique
                    #endregion NIC
                    $ComputerHW = Get-CimInstance -Class Win32_ComputerSystem -ComputerName $ComputerName |
                        Select-Object Manufacturer,Model,NumberOfProcessors,@{Expression={($_.TotalPhysicalMemory / 1GB).ToString("#,###.##")};Label="TotalPhysicalMemoryGB"}

                    $ComputerCPU = Get-CimInstance win32_processor -ComputerName $ComputerName |
                        Select-Object DeviceID,Name,Manufacturer,NumberOfCores,NumberOfLogicalProcessors
                    #region Disk
                    $ArrComputerDisks = Get-CimInstance -Class Win32_LogicalDisk -Filter "DriveType=3" -ComputerName $ComputerName |
                        Select-Object DeviceID,VolumeName,@{Expression={($_.Size / 1GB).ToString("#,###.##")};Label="SizeGB"}
                        If (($ArrComputerDisks | Measure-Object).count -gt 1) {
                            If (!([string]::IsNullOrEmpty($_.VolumeName))) {
                                $ComputerDisks = ($ArrComputerDisks | ForEach-Object {($_.DeviceID + " (" + $_.VolumeName + ") " + $_.SizeGB + " GB")})  -join ", "
                            }else{
                                $ComputerDisks = ($ArrComputerDisks | ForEach-Object {($_.DeviceID + " " + $_.SizeGB + " GB")})  -join ", "
                            }
                        } else {
                            If ($null -ne $ArrComputerDisks.VolumeName) {
                                $ComputerDisks = ( $ArrComputerDisks.DeviceID + "  (" + $ArrComputerDisks.VolumeName + ") " + $ArrComputerDisks.SizeGB + " GB" )
                            }else {
                                $ComputerDisks = ( $ArrComputerDisks.DeviceID + " " + $ArrComputerDisks.SizeGB + " GB" )
                            }
                        }   
                    #endregion Disk
                    $ComputerSerial = (Get-CimInstance Win32_Bios -ComputerName $ComputerName).SerialNumber

                    $ComputerGraphics = Get-CimInstance -Class Win32_VideoController -ComputerName $ComputerName |Select-Object Name,@{Expression={($_.AdapterRAM / 1MB).ToString("#,###.##")};Label="GraphicsRAM"}

                    $ComputerSoundDevices = ((Get-CimInstance -Class Win32_SoundDevice -ComputerName $ComputerName).Name  -join ", ") 
                            
                    $ComputerInfoManufacturer = $ComputerHW.Manufacturer
                    $ComputerInfoModel = $ComputerHW.Model
                    $ComputerInfoNumberOfProcessors = $ComputerHW.NumberOfProcessors
                    $ComputerInfoProcessorManufacturer = $ComputerCPU.Manufacturer | Select-Object -Unique
                    $ComputerInfoProcessorName = $ComputerCPU.Name | Select-Object -Unique
                    $sum = 0
                    $ComputerCPU.NumberOfCores  | ForEach-Object { $sum += $_}
                    $ComputerInfoNumberOfCores = $sum
                    $sum = 0
                    $ComputerCPU.NumberOfLogicalProcessors | ForEach-Object { $sum += $_}
                    $ComputerInfoNumberOfLogicalProcessors = $sum
                    $ComputerInfoRAM = $ComputerHW.TotalPhysicalMemoryGB
                    $ComputerInfoGraphicsName = $ComputerGraphics.Name
                    $ComputerInfoGraphicsRAM = $ComputerGraphics.GraphicsRAM
                    #region Windows Optional Features
                    $ComputerFeatures =  ((Get-CimInstance -ComputerName $ComputerName -query "select Name from win32_optionalfeature where installstate= 1").name  -join ", ") 
                    If (($ComputerOS.OperatingSystem).Contains("Server")) {
                        $ComputerRolesTemp = Invoke-Command -ComputerName $ComputerName -Verbose -ScriptBlock { Import-Module servermanager;get-windowsfeature | Where-Object { $_.installed -eq $true -and $_.featuretype -eq 'Role'} | Select-Object name}
                        $ComputerRoles =  (($ComputerRolesTemp.Name) -join ", ") 
                    }
                    #endregion Windows Optional Features
                    #region Software
                    $ArrComputerSoftware = Get-CimInstance -Class Win32_Product -ComputerName $ComputerName | Select-Object Name,Version,InstallDate | Sort-Object -Unique -Property Name,Version,InstallDate
                    If (($ArrComputerSoftware | Measure-Object).count -gt 1) {
                        $ComputerSoftware =  (($ArrComputerSoftware | ForEach-Object {($_.Name + " (" + $_.Version + ") - " + $_.InstallDate)})  -join ", ")
                    } else {
                        $ComputerSoftware =  ($ArrComputerSoftware.Name + " - " + $ArrComputerSoftware.Version) 
                    }
                    #endregion Software
                    #region Servies
                    $ArrComputerServcies = Get-CimInstance -Class Win32_Service -ComputerName $ComputerName | Select-Object Name,Caption,StartMode | Sort-Object -Unique -Property Caption
                    $ComputerServcies=@()
                    ForEach ($ItemServcie in $ArrComputerServcies) {
                        If (($ItemServcie.Caption + " - " + $ItemServcie.StartMode) -notin $RemoveServices ) {
                            $ComputerServcies += ($ItemServcie.Caption + " - " + $ItemServcie.StartMode + ",")
                        }
                    }
                    #endregion Servies
                    #region IIS
                    IF ($ArrComputerServcies | Where-Object {$_.Name -eq "W3SVC"}) {
                        Try {
                            $scriptBlock = {
                                If (Get-Module -ListAvailable -Name "WebAdministration") {
                                    If (-Not (Get-Module "WebAdministration" -ErrorAction SilentlyContinue)) {
                                        Import-Module "WebAdministration" -DisableNameChecking
                                    } 
                                }	
                                If (get-psdrive | Where-Object {$_.Name -eq "IIS"}) {
                                    Get-ChildItem "IIS:\Sites" | Where-Object -Property State -eq "Started" | ForEach-Object {
                                        Write-Output ($_.Name + " (" + ([string]$_.Bindings.Collection -join "; ") + "),")
                                        $apps = $_ | Get-ChildItem | Where-Object { $_.nodetype -eq "application" }
                                        Foreach ($app  in $apps) {
                                            Write-Output ("`t" + $app.Name )
                                        }
                                    }
                                }   
                                function ConvertFrom-IISW3CLog {
                                    [CmdletBinding()]
                                    param (
                                        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
                                        [Alias('PSPath')]
                                        [string[]]
                                        $Path,
                                        [int]$MaxMemoryUseage = 75,
                                        [string[]]$FilterField,
                                        [string[]]$KillProcesses
                                    )
                                    #Source: https://gist.github.com/jstangroome/6189660
                                    process {
                                        If ($KillProcesses) {
                                            Stop-Process -Force -Name $KillProcesses -ErrorAction SilentlyContinue
                                        }
                                        foreach ($SinglePath in $Path) {
                            
                                            $FieldNames = $null
                                            $Properties = @{}
                                            
                                            #Test Memory Usage
                                            If (Get-Command Get-CimInstance -errorAction SilentlyContinue) {
                                                $OperatingSystem = Get-CimInstance -Class win32_OperatingSystem
                                            }Else{
                                                $OperatingSystem = Get-WmiObject -Class win32_OperatingSystem
                                            }
                                            # Lets grab the free memory
                                            $FreeMemory = $OperatingSystem.FreePhysicalMemory
                                            # Lets grab the total memory
                                            $TotalMemory = $OperatingSystem.TotalVisibleMemorySize
                                            # Calculate used memory
                                            $MemoryUsed = ($TotalMemory-$FreeMemory)
                                            # Lets do some math for percent
                                            $PercentMemoryUsed = "{0:N2}" -f (($MemoryUsed / $TotalMemory) * 100)
                                            
                                            If ( $PercentMemoryUsed -le $MaxMemoryUseage ) {
                                                Get-Content -Path $SinglePath |
                                                    ForEach-Object {
                                                        if ($_ -match '^#') {
                                                            #metadata
                                                            if ($_ -match '^#(?<k>[^:]+):\s*(?<v>.*)$') {
                                                                #key value pair
                                                                if ($Matches.k -eq 'Fields') {
                                                                    $FieldNames  = @(-split $Matches.v)
                                                                }
                                                            }
                                                        } else {
                                                            $FieldValues = @(-split $_)
                                                            $Properties.Clear()
                                                            for ($Index = 0; $Index -lt $FieldValues.Length; $Index++) {
                                                                If( $null -eq $FieldValues[$Index]) {
                                                                    If ($FilterField) {
                                                                        If ($FieldNames -in $FilterField) {
                                                                            $Properties[$FieldNames[$Index]] = ""
                                                                        }
                                                                    }Else{
                                                                        $Properties[$FieldNames[$Index]] = ""
                                                                    }
                                                                }else{
                                                                    If ($null -ne $FieldNames[$Index] ) {
                                                                        If ($FilterField) {
                                                                            If ($FieldNames -in $FilterField) {
                                                                                $Properties[$FieldNames[$Index]] = $FieldValues[$Index]
                                                                            }
                                                                        }Else{
                                                                            $Properties[$FieldNames[$Index]] = $FieldValues[$Index]
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            [pscustomobject]$Properties
                                                        }
                                                    }
                                            }Else{
                                                Write-Error ("Skipped file " + (split-path -path $SinglePath -Leaf) + " on computer " + $env:computername + " Memory usage at " + $PercentMemoryUsed + "%.")
                                                If ($KillProcesses) {
                                                    Stop-Process -Force -Name $KillProcesses -ErrorAction SilentlyContinue
                                                }
                                            }
                                        }
                            
                                    }
                                }                        
                                
                                If (Get-Module -ListAvailable -Name "WebAdministration") {
                                    If (-Not (Get-Module "WebAdministration" -ErrorAction SilentlyContinue)) {
                                        Import-Module "WebAdministration" -DisableNameChecking
                                    } 
                                }	
                                If (get-psdrive | Where-Object {$_.Name -eq "IIS"}) {
                                    Get-ChildItem "IIS:\Sites" | Where-Object -Property State -eq "Started" | ForEach-Object {
                                        #Get number of hits in the last $LogHistor days
                                        If ( Test-Path -Path ([System.Environment]::ExpandEnvironmentVariables($_.logfile.directory) + "\W3SVC" + $_.id)) {
                                            $Logs = Get-ChildItem ([System.Environment]::ExpandEnvironmentVariables($_.logfile.directory) + "\W3SVC" + $_.id) -include *.log -rec | Where-Object {$_.LastWriteTime -gt (Get-Date).AddDays($LogHistory)} | Sort-Object -Property LastWriteTime -Descending
                                            $Count = 0
                                            If ($null -ne $logs) {
                                                foreach ($log in $Logs) {
                                                  $Count += ($log | ConvertFrom-IISW3CLog -FilterField "c-ip" -MaxMemoryUseage $MaxMemoryUseage -KillProcesses $KillProcesses | Where-Object {$null -ne $_."c-ip" }| Measure-Object).Count
                                                }
                                            }
                                            Write-Output ($_.Name + " (" + ([string]$_.Bindings.Collection -join ";") + ") Hits in " + $LogHistory + " : " +  $Count)
                                            $apps = $_ | Get-ChildItem | Where-Object { $_.nodetype -eq "application" }
                                            Foreach ($app  in $apps) {
                                                Write-Output ("`t" + $app.Name )
                                            }
                                        }
                                    }
                                }
                            }
                            $ComputerIIS = Invoke-Command –ComputerName $ComputerName –ScriptBlock $scriptBlock
                        }
                        Catch {
                            $ComputerIIS =""
                        }
                        
                    } Else {
                        $ComputerIIS =""
                    }
                    #endregion IIS
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Name "Computer IPs" -Value $ComputerIP
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Name "DNS" -Value $ComputerDNS
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Name "DNS Suffix" -Value $ComputerDNSSuffix
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Name "Manufacturer" -Value "$ComputerInfoManufacturer" 
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Name "Model" -Value "$ComputerInfoModel"
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Name "Serial" -Value "$ComputerSerial"
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Name "Number Of Processors" -Value "$ComputerInfoNumberOfProcessors"
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Name "Processor Manufacturer" -Value "$ComputerInfoProcessorManufacturer"
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Name "Processor Name" -Value "$ComputerInfoProcessorName"
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Name "Number Of Cores" -Value "$ComputerInfoNumberOfCores"
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Name "Number Of Logical Processors" -Value "$ComputerInfoNumberOfLogicalProcessors"
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Name "RAM" -Value "$ComputerInfoRAM"
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Name "Disk Drive" -Value "$ComputerDisks"
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Name "Graphics" -Value "$ComputerInfoGraphicsName"
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Name "Graphics RAM (MB)" -Value "$ComputerInfoGraphicsRAM"
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Name "Sound Devices" -Value "$ComputerSoundDevices"
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Name "Windows Features" -Value "$ComputerFeatures"
       

                    If (($ComputerOS.OperatingSystem).Contains("Server")) {
                        $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Name "Windows Roles" -Value "$ComputerRoles" 
                    } else {
                        $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Name "Windows Roles" -Value ""
                    }
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Name "Software" -Value "$ComputerSoftware"
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Name "Services" -Value "$ComputerServcies"
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Name "IIS" -Value "$ComputerIIS"
                    #region SSL Certs
                    $NetConnections = Get-NetTCPConnection -CimSession $ComputerName -State Listen -ErrorAction SilentlyContinue | Where-Object {$_.RemoteAddress -ne "::1" -and $_.RemoteAddress -ne "127.0.0.1" -and $_.LocalAddress -ne "127.0.0.1" -and $_.LocalAddress -ne "::1" }
                    $WebCerts = @()
                    Foreach ($NetConnection in $NetConnections) {
                        $WebCerts += Get-CheckUrl -Url ("Https://" + $ComputerName + ":"  + $NetConnection.LocalPort)
                    }
                    If (( $WebCerts | Measure-Object).count -gt 1) {
                        $StrWebCerts = (( $WebCerts | Where-Object {$null -ne $_.ExpirationOn} | ForEach-Object {($_.Url + " - " + $_.ExpirationOn)})  -join ", ") 
                    } else {
                        $StrWebCerts = ( $WebCerts.Url + " - " +  $WebCerts.ExpirationOn) 
                    }
                    $ComputerInfo | Add-Member -MemberType NoteProperty -Name "SSL Certs Expiration" -Value "$StrWebCerts" -Force
                    #endregion SSL Certs
                    #region Databases
                        $SQLVersion = ""
                        $SQL= ""
                        $SQLInstances = Invoke-Command -errorAction SilentlyContinue -HideComputerName -ComputerName $ComputerName -ScriptBlock {
                            (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\').InstalledInstances
                        }
                        If ($SQLInstances){
                            $SQLVersion = Invoke-Command -errorAction SilentlyContinue -HideComputerName -ComputerName $ComputerName -ScriptBlock {
                                If (Get-Command Invoke-SqlCmd -errorAction SilentlyContinue) {
                                    #Get SQL Server Version
                                    $Output = Invoke-SqlCmd -query "select @@version" -ServerInstance $env:computerName
                                    If ($null -eq $Output) {
                                        #Get SQL Server Version for Cluster
                                        [string]$Output = Invoke-SqlCmd -query "select @@version" -ServerInstance ($env:computerName).SubString(0,($env:ComputerName).Length-2)
                                    }
                                    $Output
                                }
                            }
                            $SQL = Invoke-Command -errorAction SilentlyContinue -HideComputerName -ComputerName $ComputerName -ScriptBlock {
                                If (Get-Command Invoke-SqlCmd -errorAction SilentlyContinue) {
                                    #Get SQL Server Databases
                                    $Output = Get-ChildItem SQLSERVER:\SQL\$env:computername\Default\Databases | Select-Object Name
                                    If ($null -eq $Output) {
                                        #Get SQL Server Databases for Cluster
                                        $Output = Get-ChildItem SQLSERVER:\SQL\($env:computerName).SubString(0,($env:ComputerName).Length-2)\Default\Databases | Select-Object Name
                                    }
                                    $Output
                                }
                            }
                        }
                        $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Name "SQL Version" -Value ($SQLVersion.Column1 -split "\n")[0]
                        $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Name "SQL Databases" -Value  ($SQL | Select-Object Name).Name
                    #endregion Databases
                    #region Forti Client
                        $FortiClient = ($ArrComputerSoftware | Where-Object {$_.Name -match "FortiClient"}).version
                        $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Name "FortiClient" -Value "$FortiClient" 
                    #endregion Forti Client
                    #region LanDesk Agent
                        If ($ArrComputerSoftware | Where-Object {$_.Caption -match "LanDesk"}) {
                            $LANDeskAgent = "True"
                        }Else{
                            $LANDeskAgent = "False"
                        }
                        $ComputerInfo | Add-Member -MemberType NoteProperty -Force -Name "LANDesk Agent Installed" -Value "$LANDeskAgent" 
                    #endregion LanDesk Agent
                    #region RemoveSMB1
                    If ($RemoveSMB1) {
                        $SMB1 = Invoke-Command -errorAction SilentlyContinue -HideComputerName -ComputerName $ComputerName -ScriptBlock {
                            $SMB1Enabled = $true
                            If (Get-Command Get-WindowsOptionalFeature -errorAction SilentlyContinue) {
                                $WOF = Get-WindowsOptionalFeature -Online -ErrorAction SilentlyContinue | Where-Object {($_.FeatureName -eq "SMB1Protocol" -or $_.FeatureName -eq "SMB1Protocol-Server") -and $_.state -eq "Enabled" }
                                If ($WOF.state -eq "Enabled") {
                                    $WOF | Disable-WindowsOptionalFeature -online -NoRestart -WarningAction SilentlyContinue | Out-Null
                                }
                                #Test
                                If (Get-WindowsOptionalFeature -Online -ErrorAction SilentlyContinue | Where-Object {($_.FeatureName -eq "SMB1Protocol" -or $_.FeatureName -eq "SMB1Protocol-Server") -and $_.state -eq "Disabled"}) {
                                    $SMB1Enabled = $false
                                }
                            }
                            If (Get-Command Get-WindowsFeature -errorAction SilentlyContinue) {
                                $WF = Get-WindowsFeature -Name "FS-SMB1" -ErrorAction SilentlyContinue
                                If ( $WF.Installed) {
                                    $WF | Uninstall-WindowsFeature -Restart:$false -WarningAction SilentlyContinue | Out-Null
                                }
                                If ((Get-WindowsFeature -Name "FS-SMB1" -ErrorAction SilentlyContinue).Installed -eq $false) {
                                    $SMB1Enabled = $false
                                }
                            }
                            If (Get-Command Get-SmbServerConfiguration -errorAction SilentlyContinue) {	
                                If ((Get-SmbServerConfiguration).EnableSMB1Protocol) {
                                    Set-SmbServerConfiguration -EnableSMB1Protocol $false -confirm:$false
                                    If ((Get-SmbServerConfiguration).EnableSMB1Protocol -eq $false) {
                                        $SMB1Enabled = $false
                                    }
                                }else {
                                    $SMB1Enabled = $false
                                }
                            }
                            If ([Environment]::OSVersion.Version -lt (new-object 'Version' 6,2)) {
                                If ((Get-Item HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters | ForEach-Object {Get-ItemProperty $_.pspath}).SMB1 -ne 0) {
                                    Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters" SMB1 -Type DWORD -Value 0 –Force
                                }
                                If ((Get-Item HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters | ForEach-Object {Get-ItemProperty $_.pspath}).SMB1 -eq 0) {
                                    $SMB1Enabled = $false
                                }
                            }
                            Restart-Service -Force -Name "LanmanServer"
                            #Results
                            If ( $SMB1Enabled) {
                                return ("SMB1 Enabled")
                            }else{
                                return ("SMB1 Disabled")
                            }
                        }
                    }Else{
                        $SMB1 = Invoke-Command -errorAction SilentlyContinue -HideComputerName -ComputerName $ComputerName -ScriptBlock {
                            If (Get-Command Get-SmbServerConfiguration -errorAction SilentlyContinue) {	
                                If ((Get-SmbServerConfiguration).EnableSMB1Protocol ) {
                                    $SMB1Enabled = $true
                                }else{
                                    $SMB1Enabled = $false
                                }
                            }Elseif([Environment]::OSVersion.Version -lt (new-object 'Version' 6,2) -and (Get-Item HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters | ForEach-Object {Get-ItemProperty $_.pspath}).SMB1 -eq 0) {
                                $SMB1Enabled = $false
                            }Else{
                                $SMB1Enabled = $true
                            }
                            If (Get-Command Get-WindowsOptionalFeature -errorAction SilentlyContinue) {
                                If (Get-WindowsOptionalFeature -Online -ErrorAction SilentlyContinue | Where-Object {($_.FeatureName -eq "SMB1Protocol" -or $_.FeatureName -eq "SMB1Protocol-Server") -and $_.state -eq "Disabled"}) {
                                    $SMB1Enabled = $false
                                }
                            }
                            If (Get-Command Get-WindowsFeature -errorAction SilentlyContinue) {
                                If ((Get-WindowsFeature -Name "FS-SMB1" -ErrorAction SilentlyContinue).Installed -eq $False) {
                                    $SMB1Enabled = $false
                                }
                            }
                            #Results
                            If ($SMB1Enabled) {
                                return ("SMB1 Enabled")
                            }else{
                                return ("SMB1 Disabled")
                            }
                        }
                    }
                    $ComputerInfo | Add-Member  -Force -MemberType NoteProperty -Name "SMB Status" -Value $SMB1
                    #endregion RemoveSMB1
                    #region Pending Reboot
                        $NeedReboot = Invoke-Command -errorAction SilentlyContinue -ComputerName $ComputerName -ScriptBlock {
                            If ((Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') `
                                -or (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired')`
                                -or (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name 'PendingFileRenameOperations' -ErrorAction Ignore)) {
                                $true
                            }else {
                                $false
                            }
                        }
                        $ComputerInfo | Add-Member  -Force -MemberType NoteProperty -Name "Reboot Needed" -Value $NeedReboot
                    #endregion Pending Reboot
                    #region WMF and PS Version
                    $PSV = Invoke-Command -errorAction SilentlyContinue -HideComputerName -ComputerName $ComputerName -ScriptBlock {
                        If ($using:InstallWMF -and $PSVersionTable.PSVersion.Major -le 5 -and $PSVersionTable.PSVersion.Minor -le 1) {
                            If ((Get-PSDrive -PSProvider FileSystem) | Where-Object -Property Root -eq $useing:WSO) {
                                $DL = [string]((Get-PSDrive -PSProvider FileSystem) | Where-Object -Property Root -eq $useing:WSO).Name
                            }Else{
                                $DL = (68..90 | ForEach-Object{$L=[char]$_; if ((Get-PSDrive -PSProvider FileSystem).Name -notContains $L) {$L}}) | Select-Object -last 1
                                $pinfo = New-Object System.Diagnostics.ProcessStartInfo
                                $pinfo.FileName = "net.exe"
                                $pinfo.RedirectStandardError = $true
                                $pinfo.RedirectStandardOutput = $true
                                $pinfo.UseShellExecute = $false
                                $pinfo.Arguments = ("use " + $DL + ": " + $useing:WSO) 
                                $p = New-Object System.Diagnostics.Process
                                $p.StartInfo = $pinfo
                                $p.Start() | Out-Null
                                $p.WaitForExit()
                                [System.GC]::Collect()
                            }
                            return ($PSVersionTable.PSVersion + "`tInstalling WMF 5.1")
                            Set-Location -Path ($DL + ":\")
                            Start-Process -FilePath $($DL + ":\Update.cmd") -ArgumentList "/instwmf"
                        }else{
                            return ($PSVersionTable.PSVersion)
                        }
                        return ($PSVersionTable.PSVersion)
                    }
                    $ComputerInfo | Add-Member  -Force -MemberType NoteProperty -Name "PowerShell Version" -Value $PSV
                    #endregion WMF and PS Version
                    $ComputerInfo | Add-Member  -Force -MemberType NoteProperty -Name "Distinguished Name" -Value $ComputerOS.DistinguishedName
                } else {
                    #Write-Host ("`t`tHost WMI Not Reachable for computer: " + $ComputerName) -ForegroundColor Red
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "Computer IPs"    
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "DNSs"
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "DNS Suffixs"
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "Manufacturer"
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "Model"
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "Serial"
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "Number Of Processors"
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "Processor Manufacturer"
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "Processor Name"
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "Number Of Cores" 
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "Number Of Logical Processors"
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "RAM"
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "Disk Drive" 
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "Graphics"  
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "Graphics RAM (MB)" 
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "Sound Devices" 
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "Windows Features" 
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "Windows Roles"
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "Software" 
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "Services"
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "IIS" 
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "SSL Certs Expiration" 
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "SQL Version" 
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "SQL Databases" 
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "FortiClient" 
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "LANDesk Agent Installed" 
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "SMB Status"
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "Reboot Needed" 
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "PowerShell Version"
                    $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value $ComputerOS.DistinguishedName -Name "Distinguished Name" 
                }
            }Else {
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "Computer IPs"    
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "DNSs"
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "DNS Suffixs"
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "Manufacturer"
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "Model"
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "Serial"
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "Number Of Processors"
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "Processor Manufacturer"
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "Processor Name"
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "Number Of Cores" 
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "Number Of Logical Processors"
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "RAM"
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "Disk Drive" 
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "Size" 
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "Graphics"  
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "Graphics RAM (MB)" 
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "Sound Devices" 
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "Windows Features" 
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "Windows Roles"
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "Software" 
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "Services"
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "IIS" 
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "SSL Certs Expiration" 
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "SQL Version" 
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "SQL Databases" 
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "FortiClient" 
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "LANDesk Agent Installed" 
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "SMB Status"
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "Reboot Needed" 
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value "" -Name "PowerShell Version"
                $ComputerInfo | Add-Member -Force -MemberType NoteProperty -Value $ComputerOS.DistinguishedName -Name "Distinguished Name" 
            }

            If ($ComputerInfo) {
                #$Inventory.Add($ComputerInfo) | Out-Null
                return $ComputerInfo
            }
        #Stop Script block
        } | Out-Null
    }
    Write-Progress  -Id 0 -Activity "Active Directory Computers Count: $AllComputersNamesCount" -Percentcomplete ([int](($ScanCount+($Inventory.Count))/$AllComputersNames.count))

    do {
        Write-Verbose "Trying get part of data."
        Get-Job -State Completed | ForEach-Object {
            Write-Verbose "Geting job $($_.Name) result."
            $JobResult = Receive-Job -Id ($_.Id)

            if($ShowAll) {
                if($ShowInstantly) {
                    if($JobResult.Active -eq $true) {
                        Write-Host "$($JobResult.Name) is active." -ForegroundColor Green
                    } else {
                        Write-Host "$($JobResult.Name) is inactive." -ForegroundColor Red
                    }
                }
                
                	
            } else {
                if($JobResult.Active -eq $true) {
                    if($ShowInstantly) {
                        Write-Host "$($JobResult.Name) is active." -ForegroundColor Green
                    }
                }
            }
            Write-Verbose "Removing job $($_.Name)."
            $Inventory += $JobResult
            $JobCount++
            Remove-Job -Id ($_.Id)

            Write-Progress  -Id 1  -Activity "Retrieving results, please wait..." -Status "$($_.Name)" -Percentcomplete ([int](($JobCount)/$AllComputersNames.count))
        }
        
        if((Get-Job -name *).Count -eq $MaxJobs) {
            Write-Verbose "Jobs are not completed ($((Get-Job -name * | Measure-Object).Count)/$MaxJobs), please wait..."
            Write-Progress  -Id 1  -Activity "Jobs are not completed, please wait..." -Status "$($_.Name + " ")" -Percentcomplete ([int](($JobCount)/$AllComputersNames.count))
            Start-Sleep $SleepTime
        }
    } while((Get-Job -name *).Count -eq $MaxJobs)
    $ScanCount++
}

do {
        Write-Verbose "Trying get last part of data."
        Get-Job -State Completed | ForEach-Object {
            Write-Verbose "Getting job $($_.Name) result."
            $JobResult = Receive-Job -Id ($_.Id)

            if($ShowAll) {
                if($ShowInstantly) {
                    if($JobResult.Active -eq $true) {
                        Write-Host "$($JobResult.Name) is active." -ForegroundColor Green
                    } else {
                        Write-Host "$($JobResult.Name) is inactive." -ForegroundColor Red
                    }
                }
               $Inventory += $JobResult	
               $JobCount++
            } else {
                if($JobResult.Active -eq $true) {
                    if($ShowInstantly) {
                        Write-Host "$($JobResult.Name) is active." -ForegroundColor Green
                    }
                   $Inventory += $JobResult
                   $JobCount++
                }
            }
            Write-Verbose "Removing job $($_.Name)."
            Remove-Job -Id ($_.Id)
            Write-Progress  -Id 1  -Activity "Retrieving results, please wait..." -Status "$($_.Name + " ")" -Percentcomplete ([int](($JobCount)/$AllComputersNames.count))
        }
        
        if(Get-Job -name *) {
            Write-Verbose "All jobs are not completed ($((Get-Job -name *| Measure-Object).Count)/$MaxJobs), please wait... ($timeOutCounter)"
            Write-Progress  -Id 1  -Activity "All jobs are not completed, please wait..." -Status "$($_.Name + " ")" -Percentcomplete ([int](($JobCount)/$AllComputersNames.count))
            Start-Sleep $SleepTime
            $timeOutCounter += $SleepTime				

            if($timeOutCounter -ge $TimeOut) {
                Write-Verbose "Time out... $TimeOut. Can't finish some jobs  ($((Get-Job -name * | Measure-Object).Count)/$MaxJobs) try remove it manualy."
                Break
            }
        }
    } while(Get-Job -name *)
    
Write-Verbose "Scan finished."
$Inventory | Sort-Object ADName | Export-Csv -NoTypeInformation -Path ($OutputFolder + "\Inventory_" + (Get-Date -format yyyyMMdd-hhmm) + ".csv")

if ($Email -eq $true){send-mailmessage @EmailParameters}

$sw.Stop()
Write-Host ("Script took: " + (FormatElapsedTime($sw.Elapsed)) + " to run.")
#============================================================================
#endregion Main Local Machine Cleanup
#============================================================================
If (-Not [string]::IsNullOrEmpty($LogFile)) {
	Stop-Transcript
}
