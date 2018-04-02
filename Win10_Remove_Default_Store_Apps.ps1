#From: http://blog.randait.com/2016/08/remove-windows-10-bloat/
#region Elevate Powershell if needed
$myWindowsID=[System.Security.Principal.WindowsIdentity]::GetCurrent()
$myWindowsPrincipal=new-object System.Security.Principal.WindowsPrincipal($myWindowsID)
$adminRole=[System.Security.Principal.WindowsBuiltInRole]::Administrator
 
if ($myWindowsPrincipal.IsInRole($adminRole))
   {
   clear-host
   }
else
   {
   $newProcess = new-object System.Diagnostics.ProcessStartInfo "PowerShell";
   $newProcess.Arguments = "& '" + $myInvocation.MyCommand.Definition + "'";
   $newProcess.Verb = "runas";
   [System.Diagnostics.Process]::Start($newProcess);
   exit
   }
#endregion

#APSS to Keep:
#"Microsoft.Appconnector", 
#"Microsoft.BingWeather" 
#"Microsoft.MicrosoftStickyNotes", 
#"Microsoft.WindowsAlarms",
#"Microsoft.WindowsCalculator",
#"Windows.MiracastView",
#"Microsoft.Windows.Cortana"

#region Appx to be removed
$Remove = "Microsoft.3DBuilder",
"Microsoft.Advertising.Xaml",
"Microsoft.Getstarted",
"Microsoft.Messaging",
"Microsoft.Microsoft.Microsoft3DViewer",
"Microsoft.MicrosoftOfficeHub",
"Microsoft.MicrosoftSolitaireCollection",
"Microsoft.Office.OneNote",
"Microsoft.Office.Sway",
"Microsoft.Office.Desktop.Access",
"Microsoft.Office.Desktop.Excel",
"Microsoft.Office.Desktop.Outlook",
"Microsoft.Office.Desktop.PowerPoint",
"Microsoft.Office.Desktop.Publisher",
"Microsoft.Office.Desktop.Word",
"Microsoft.Office.Desktop",
"Microsoft.OneConnect",
"Microsoft.People",
"Microsoft.SkypeApp",
"Microsoft.Wallet",
"Microsoft.WindowsCamera",
"Microsoft.WindowsFeedbackHub",
"Microsoft.WindowsMaps",
"Microsoft.WindowsSoundRecorder",
"Microsoft.XboxApp",
"Microsoft.XboxGameCallableUI",
"Microsoft.XboxIdentityProvider",
"Microsoft.XboxSpeechToTextOverlay",
"Microsoft.Xbox.TCUI",
"Microsoft.ZuneMusic",
"Microsoft.ZuneVideo",
"microsoft.windowscommunicationsapps"
#endregion
 
#region Get list of currently installed and provisioned Appx packages
$AllInstalled = Get-AppxPackage -AllUsers | Foreach {$_.Name}
$AllProvisioned = Get-ProvisionedAppxPackage -Online | Foreach {$_.DisplayName}
#endregion
 
#region Remove Appx Packages
Write-Host "`n"
Write-Host '#####################################' -ForegroundColor Green
Write-Host -NoNewline '#' -ForegroundColor Green
Write-Host -NoNewline '           '
Write-Host -NoNewline "Appx Packages" -ForegroundColor White
Write-Host -NoNewline '           '
Write-Host '#' -ForegroundColor Green
Write-Host '#####################################' -ForegroundColor Green
Write-Host "`n"
Foreach($Appx in $Remove){
    $error.Clear()
    If($AllInstalled -like "*$Appx*"){
        Try{
            #Turn off the progress bar
            $ProgressPreference = 'silentlyContinue'
            Get-AppxPackage -Name $Appx | Remove-AppxPackage
            #Turn on the progress bar
            $ProgressPreference = 'Continue'
        }
         
        Catch{
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Write-Host "There was an error removing Appx:"
            Write-Host $ErrorMessage
            Write-Host $FailedItem
        }
        If(!$error){
            Write-Host "Removed Appx: $Appx"
        }
    }
    Else{
        Write-Host "Appx Package not installed: $Appx"
    }
}
#endregion
 
#region Remove Provisioned Appx Packages
Write-Host "`n"
Write-Host '#####################################' -ForegroundColor Green
Write-Host -NoNewline '#' -ForegroundColor Green
Write-Host -NoNewline '     '
Write-Host -NoNewline "Provisioned Appx Packages" -ForegroundColor White
Write-Host -NoNewline '     '
Write-Host '#' -ForegroundColor Green
Write-Host '#####################################' -ForegroundColor Green
Write-Host "`n"
Foreach($Appx in $Remove){
    $error.Clear()
    If($AllProvisioned -like "*$Appx*"){
        Try{
            Get-ProvisionedAppxPackage -Online | where {$_.DisplayName -eq $Appx} | Remove-ProvisionedAppxPackage -Online | Out-Null
        }
         
        Catch{
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Write-Host "There was an error removing Provisioned Appx:"
            Write-Host $ErrorMessage
            Write-Host $FailedItem
        }
        If(!$error){
            Write-Host "Removed Provisioned Appx: $Appx"
        }
    }
    Else{
        Write-Host "Provisioned Appx Package not installed: $Appx"
    }
}
#endregion

Write-Host "`n"
