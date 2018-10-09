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

#Windows 10 Rev. 1803 WhiteList
#APSS to Keep:
$Keep =  "1527c705-839a-4832-9118-54d4Bd6a0c89",
"c5e2524a-ea46-4f67-841f-6a9465d9d515",
"E2A4F912-2574-4A75-9BB0-0D023378592B",
"F46D4000-FD22-4DB4-AC8E-4E1DDDE828FE",
"InputApp",
"Microsoft.AAD.BrokerPlugin",
"Microsoft.AccountsControl",
"Microsoft.Appconnector",
"Microsoft.AsyncTextService",
"Microsoft.BingWeather", 
"Microsoft.BioEnrollment",
"Microsoft.CredDialogHost",
"Microsoft.ECApp",
"Microsoft.LockApp",
"Microsoft.MSPaint",
"Microsoft.MicrosoftEdge",
"Microsoft.MicrosoftEdgeDevToolsClient",
"Microsoft.MicrosoftStickyNotes", 
"Microsoft.NET.Native.Framework.1.6",
"Microsoft.NET.Native.Framework.1.7",
"Microsoft.NET.Native.Framework.2.1",
"Microsoft.NET.Native.Runtime.1.6",
"Microsoft.NET.Native.Runtime.1.7",
"Microsoft.NET.Native.Runtime.2.1",
"Microsoft.Office.OneNote",
"Microsoft.PPIProjection",
"Microsoft.People",
"Microsoft.SkypeApp",
"Microsoft.StorePurchaseApp",
"Microsoft.VCLibs.140.00",
"Microsoft.VCLibs.140.00.UWPDesktop",
"Microsoft.Wallet",
"Microsoft.Win32WebViewHost",
"Microsoft.Windows.Apprep.ChxApp",
"Microsoft.Windows.AssignedAccessLockApp",
"Microsoft.Windows.CapturePicker",
"Microsoft.Windows.CloudExperienceHost",
"Microsoft.Windows.ContentDeliveryManager",
"Microsoft.Windows.Cortana",
"Microsoft.Windows.HolographicFirstRun",
"Microsoft.Windows.OOBENetworkCaptivePortal",
"Microsoft.Windows.OOBENetworkConnectionFlow",
"Microsoft.Windows.ParentalControls",
"Microsoft.Windows.PeopleExperienceHost",
"Microsoft.Windows.Photos",
"Microsoft.Windows.PinningConfirmationDialog",
"Microsoft.Windows.SecHealthUI",
"Microsoft.Windows.SecureAssessmentBrowser",
"Microsoft.Windows.ShellExperienceHost",
"Microsoft.WindowsAlarms",
"Microsoft.WindowsCalculator",
"Microsoft.WindowsCamera",
"Microsoft.WindowsFeedbackHub",
"Microsoft.WindowsMaps",
"Microsoft.WindowsStore",
"Microsoft.Xbox.TCUI",
"Microsoft.XboxApp",
"Microsoft.XboxGameCallableUI",
"Microsoft.XboxIdentityProvider",
"Windows.CBSPreview",
"Windows.MiracastView",
"Windows.PrintDialog",
"windows.immersivecontrolpanel"


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
Foreach($Appx in $AllInstalled){
    $error.Clear()
    If(-Not $Keep.Contains($Appx)){
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
            Write-Host "There was an error removing Appx: $Appx"
            Write-Host $ErrorMessage
            Write-Host $FailedItem
        }
        If(!$error){
            Write-Host "Removed Appx: $Appx"
        }
    }
    Else{
        Write-Host "Appx Package is whitelisted: $Appx" -ForegroundColor yellow
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
Foreach($Appx in $AllProvisioned){
    $error.Clear()
    If(-Not $Keep.Contains($Appx)){
        Try{
            Get-ProvisionedAppxPackage -Online | where {$_.DisplayName -eq $Appx} | Remove-ProvisionedAppxPackage -Online | Out-Null
        }
         
        Catch{
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Write-Host "There was an error removing Provisioned Appx: $Appx"
            Write-Host $ErrorMessage
            Write-Host $FailedItem
        }
        If(!$error){
            Write-Host "Removed Provisioned Appx: $Appx"
        }
    }
    Else{
        Write-Host "Appx Package is whitelisted: $Appx" -ForegroundColor yellow
    }
}
#endregion

Write-Host "`n"
