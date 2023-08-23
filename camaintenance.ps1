<#
Script Info

Author: Sebastian Kerssen [MSFT] /Andreas Lucas [MSFT]
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
    This script run maintenance tasks on Microsoft PKI

.DESCRIPTION
    This script runs serveral maintenance task on a Microsoft Certification Authority. The maintenance tasks are:
    - Deleting expired certificates from the ADCS database
    - Deleting pending certificate requests from the ADCS database
    - Deleting failed certificate requests from the ADCS database
    - Creating a status report and send the report via e-mail
    - writing a status information to a Azure LogAnalytics table

.EXAMPLE
    .\camaintenance.ps1

.INPUTS
    -config <camaintenance.config>
        A alternat camaintenance.ps1 configuration file. the default value is .\camaintenance.config

.OUTPUTS
   none
.NOTES
    Error Codes:
    1 Error 
    2 Error
    3 Error config file missing
    1000 info exporting certificate
    1001 info removing pending requests
    1003 info removed failed certificates
    1004 info amount of sucessful removed certificates
    
    Version Tracking
    20230302
    Version 0.1
        - First internal release
    Version 0.1.20230330
        Adding comments [AL]
        If the export of certificates failed, no certificates will be deleted from database
#>
param(
    # alternate configuration file name
    [Parameter(Mandatory = $false)]
    $configFileName = ".\camaintain.config"
)

#region initializing environment
if (!(Get-Module -ListAvailable -Name PSPKI)) {
    Write-Host "Missing PSPKI Module. Install Module from https://www.powershellgallery.com/packages/PSPKI"
    return
} 
Import-Module -Name PSPKI

$env:ScriptVersion = "0.1.230330" #Script Version <major>.<minor>.<date of change>
$_configfileVersion = "0.1.230322"

$eventSource = "CAMaintenance"

#Setting Debugging option
if ($DebugOutput -eq $true) { $DebugPreference = "Continue" } else { $DebugPreference = "SilentlyContinue" }

if (!([System.Diagnostics.EventLog]::SourceExists($eventSource))) {
    [System.Diagnostics.EventLog]::CreateEventSource($eventsource, "Application")
}


#region reading and validating the configuration file
#Validate the config file is available and has a version number equal the required configfile version or higher
if (Test-Path $configFileName) {
    $config = Get-Content $configFileName | ConvertFrom-Json
    #Validate the config file version is equal or higher then the configuration file version
    if (([int]($config.ConfigScriptVersion -replace "\.", "")) -lt ([INT]($_configfileVersion -replace "\.", "") - 1)) {
        Write-Host "invalid configfile version"
    }
}
else {
    Write-Host "configuration file missing $configFileName" -ForegroundColor Red
    Write-EventLog Application -Source $eventSource -EventId 3 -EntryType Error -Message "missing configuration file $configFileName "
    break
}

#Register Eventsource
if ([System.Diagnostics.EventLog]::SourceExists($eventSource) -eq $false) {
    New-EventLog -LogName Application -Source $eventSource
}
#endregion
#region Delete expired certificates
#Removing certificates from the database will only occurs if the configuration.RemoveCertificatesAfter is higher then 0.
#The parameter is the amount on months how long a expired certificate will be stored inthe ADCS database before it will be delted
#from the database
if ($config.RemoveCertificatesAfter -gt 0) {
    #collect all expired certificates based on the RemoveCertificatesAfter parameter
    try {
        $certs = Get-AdcsDatabaseRow -CertificationAuthority $config.CertificationAuthority -Table issued -Filter "notafter -lt $((Get-date).AddMonths($config.RemoveCertificatesAfter))"
        # it is possible to store the certificates on a file share  before it will be deleted
        if ($config.RemoveCertificatesExportBeforeDelete) {
            foreach ($cert in $certs ) {
                $cert.RawCertificate | Out-File -FilePath "$($config.RemoveCertificatesExportPath)\$($cert.SerialNumber).cer"
            }
        }
        $certs | Remove-AdcsDatabaseRow
        Write-Host "$($certs.Count) removed from the database"
        Write-EventLog Application -Source $eventSource -EventId 1004 -EntryType Information -Message "$($certs.Count) removed from the database"
    }
    catch {
        Write-Host "A error occurs: $($Error[0]) while removeing certificates from database"
        Write-EventLog Application -Source $eventSource -EventId 1 -EntryType Error -Message "A error occurs: $($Error[0]) while removeing certificates from database"
    }
}
#endregion
#region Removing pending expired certificate requests
#This section remove pending request, if they are older then the amount of months configured in the configuration file
if ($config.RemovePendingRequestAfter -gt 0) {
    $RemoveRowsBefore = (Get-Date).AddMonths(-$config.RemovePendingRequestAfter)
    # $RemoveRowsBefore = Get-Date $RemoveRowsBefore -Format "MM/dd/yyyy HH:MM:ss"
    if ($config.RemovePendingRequestAfterGrace -gt 0) {
        $RemoveRowsBefore = $RemoveRowsBefore.AddDays($config.RemovePendingRequestAfterGrace)
    }
    try {
        $PendingRequests = Get-CertificationAuthority $config.CertificationAuthority | Get-PendingRequest -Filter "Request.SubmittedWhen -gt $RemoveRowsBefore" 
        if ($PendingRequests.Count -gt 0) {
            $pendingRequests | Remove-AdcsDatabaseRow
            Write-Eventlog Application -Source $eventSource -EventId "1001" -Message "Removing $($pendingRequests.Count) Pending Requests" -EntryType Information
        }
    }
    catch {
        Write-Host "A error occurs $($Error[0]) while removing pending certificares from database"
        Write-EventLog Application -Source $eventSource -EventId 2 -EntryType Error -Message "A error occurs $($Error[0]) while removing pending certificares from database"
    }
}
#endregion
#region delting failed requests
if ($config.RemoveFailedRequestsAfter -gt 0) {
    $RemoveRowsBefore = (Get-Date).AddDays($config.RemoveFailedRequestsAfter)
    try {
        $failedRequests = Get-CertificationAuthority $config.CertificationAuthority | Get-FailedRequest -Filter "Request.SubmittedWhen -lt $RemoveRowsBefore"
        Write-Eventlog -Source $eventSource -EventId "1003" -Message "Removing $($failedRequests.Count) Failed Requests" -EntryType Information
        $failedRequests | Remove-AdcsDatabaseRow
    }
    catch {
        Write-host "A error occurs while $($Error[0]) while removing filed requests"
        Write-EventLog Application -Source $eventSource -EventId 3 -EntryType Error -Message^"A error occurs while $($Error[0]) while removing filed requests"
    }
}
#endregion