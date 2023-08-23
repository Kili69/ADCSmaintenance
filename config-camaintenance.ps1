<#
Script Info

Author: Andreas Lucas [MSFT]
Download: 

Disclaimer:
This sample script is not supported under any Microsoft standard support program or service. 
The sample script is provided AS IS without warranty of any kind. Microsoft further disclaims 
all implied warranties including, without limitation, any implied warranties of merchantability 
or of fitness for a particular purpose. The entire risk arising out of the use or performance of 
the sample scripts and documentation remains with you. In no event shall Microsoft, its authors, 
or anyone else involved in the creation, production, or delivery of the scripts be liable for any 
damages whatsoever (including, without limitation, damages for loss of business profits, business 
interruption, loss of business information, or other pecuniary loss) arising out of the use of or 
inability to use the sample scripts or documentation, even if Microsoft has been advised of the 
possibility of such damages
#>
<#
.Synopsis
    This script provides the configuration of the CA-Mainaintance script

.DESCRIPTION
    This script create the JSON configuration file for the CA-Mainainance script

.EXAMPLE
    .\config.camaintenance.ps1

.INPUTS
  

.OUTPUTS
   none
.NOTES
    Version Tracking
    20230301
    Version 0.1
        - First internal release
        20230302
    Version 0.1 <AL>
        - Update documentation
        - added RemoveFailedRequestsAfter
        - added Remote CA parameter
        - removed the E-Mail notification
    
#>

$_scriptVersion = "0.1.230322"              #Script Version <major>.<minor>.<date of change>
$configFileName = ".\camaintain.config"     #JSON config location and file name


#Setting Debugging option
if ($DebugOutput -eq $true) { $DebugPreference = "Continue" } else { $DebugPreference = "SilentlyContinue" }

#region Read existing config file
if (Test-Path ".\$configFileName") {
    #read existing config file
    $config = Get-Content ".\$configFileName" | ConvertFrom-Json
}
else {
    #creating the default object 
    $config = New-Object psobject
    $config | Add-Member -MemberType NoteProperty -Name "ConfigScriptVersion"                   -Value $_scriptVersion
    $config | Add-Member -MemberType NoteProperty -Name "RemoveCertificatesAfter"               -Value 0
    $config | Add-Member -MemberType NoteProperty -Name "RemoveCertificatesExportBeforeDelete"  -Value $false
    $config | Add-Member -MemberType NoteProperty -Name "RemoveCertificatesExportPath"          -Value "$(Get-Location)\Export"
    $config | Add-Member -MemberType NoteProperty -Name "RemoveCertificatesEmailReceiver"       -Value "" #E-Mail is deprecated
    $config | Add-Member -MemberType NoteProperty -Name "RemoveCertificatesEmailTemplate"       -Value "" #E-Mail is deprecated
    $config | Add-Member -MemberType NoteProperty -Name "RemovePendingRequestAfter"             -Value 0
    $config | Add-Member -MemberType NoteProperty -Name "RemovePendingRequestAfterGrace"        -Value 0
    $config | Add-Member -MemberType NoteProperty -Name "RemoveFailedRequestsAfter"             -Value 0
    $config | Add-Member -MemberType NoteProperty -Name "CertificationAuthority"                -Value $env:COMPUTERNAME
}
#endregion

#update the existing configuration object
Write-Host "This script created the configuration for the CA-Maintainance powershell script."
$rhValue = Read-Host "Remove expired certificates from the database. Expired certificates will be deleted after(current value $($config.RemoveCertificatesAfter) months)"
if ($rhValue -ne "") { #if the user press return nothing will change
    #input validation. only numbers between 0 and 99 allowed
    while ($rhValue -notin 0..99) {
        Write-Host "Invalid value, please take number between 1 and 99. 0 Will disable the function" -ForegroundColor Red
        $rhValue = Read-Host "Remove expired certificates from the database. Expired certificates will be deleted after(current value $($config.RemoveCertificatesAfter) months):"
    }
    $config.RemoveCertificatesAfter = $rhValue
}
# if the Remove expired certificates is enabled (value > 0) continue with more configuration
if ($config.RemoveCertificatesAfter -gt 0) {
    $rhValue = Read-Host "Export certificates before they will be removed from the database?(Y/[N])"
    if ($rhValue -like "y") {
        if ($config.RemoveCertificatesExportPath -eq "") {
            $rhValue = "$(Get-Location)\Export"
        }
        else {
            $rhValue = $config.RemoveCertificatesExportPath
        }
        $rhValue = Read-Host "Export path:($($rhValue)):"
        if ($rhValue -ne "") {
            #validate the path exists
            while (-not(test-path $rhValue)) { 
                Write-Host "$rhValue is a invalid path" -ForegroundColor Red
                $rhValue = Read-Host "Export path:($rhValue):"
            }
            $config.RemoveCertificatesExportPath = $rhValue
        }
        $config.RemoveCertificatesExportBeforeDelete = $true
    }
    else {
        $config.RemoveCertificatesExportBeforeDelete = $false
    }
    <# Not in use anymore will be removed
    $rhValue = Read-Host "Do you want to send a email notification on certificate deletion? (Y/[N])"
    if ($rhValue -like "n" -or $rhValue -eq "") {
        $config.RemoveCertificatesEmailReceiver = ""
    }
    else {
        $rhValue = Read-Host "E-Mail address ($($config.RemoveCertificatesEmailReceiver)):"
        if ($rhValue -ne "") {
            while (-not (ValidEmailAddress($rhValue)) ) {
                Write-Host "Invalid email address" -ForegroundColor Red
                $rhValue = Read-Host "E-Mail address:"
            }
            $config.RemoveCertificatesEmailReceiver = $rhValue
        }
    }
    #>
}
#continue with deleting pending certificate requests
$rhValue = Read-Host "Remove pending certificate requests after ($($config.RemovePendingRequestAfter)) days:"
if ($rhValue -ne "") {
    while ($rhValue -notin 0..356) {
        Write-Host "Invalid value, please define 0..356 days before pending certificate requests will be deleted" -ForegroundColor Red
        $rhValue = Read-Host "Remove pending certificate requests after ($($config.RemovePendingRequestAfter)) days:"
    }
    $config.RemovePendingRequestAfter = $rhValue
}
if ($config.RemovePendingRequestAfter -gt 0) {
    $rhValue = Read-Host "Grace period before deleting pending certificate requests (0..356):[$($config.RemovePensingRequestAfterGrace)]"
    if ($rhValue -ne "") {
        while ($rhValue -notin 0..356) {
            Write-Host "Invalid value, please define 0..356 days before pending certificate requests will be deleted:" -ForegroundColor Red
            $rhValue = Read-Host "Grace period before deleting pending certificate requests (0..356):"
            if ($rhValue -eq "") {
                $rhValue = $config.RemovePensingRequestAfterGrace
            }
        }    
    }
}
#continue with deleting failed certificate requests
$rhValue = Read-Host "Remove failed certificate requests after ($($config.RemoveFailedRequestsAfter)) days:"
if ($rhValue -ne ""){
    while ($rhValue -notin 0..365) {
        Write-Host "Invalid value, please define 0..365 days before deleting failed certificate requests will be deleted" -ForegroundColor Red
        $rhValue = Read-Host "Remove failed certificate requests after ($($config.RemoveFailedRequestsAfter)) days:"
    }
    $config.RemoveFailedRequestsAfter = $rhValue
}
$rhValue = Read-Host "CertificationAuthority($($config.CertificationAuthority)) "
if ($rhValue -ne "") {
    $config.CertificationAuthority = $rhValue
}

#region Writing config file
ConvertTo-Json $config | Out-File "$configFileName" -Force
#endregion