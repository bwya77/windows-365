<#
.SYNOPSIS
    Deprovisions a single Windows 365 Reserve Cloud PC by ID.

.DESCRIPTION
    Designed to run as an Azure Automation runbook using a system-assigned managed identity,
    called once per Cloud PC from a Logic App "For each" loop after an approval email has been
    accepted. Calls Microsoft Graph's dedicated deprovision action, which deprovisions the Cloud
    PC and returns the Reserve license/allotment slot to the pool for reuse.

.PARAMETER CloudPcId
    The id (GUID) of the Cloud PC to deprovision, as returned in the idleCloudPcs array from
    Get-IdleWindows365ReserveCloudPCs.ps1.

.PARAMETER ManagedDeviceName
    Optional, for logging/output only -- not required by Graph, but makes job output and the
    resulting confirmation email readable without a second lookup.

.NOTES
    Required Microsoft Graph Application permission (granted to the Automation Account's
    managed identity): CloudPC.ReadWrite.All.

    Required PowerShell module in the Automation Account: Microsoft.Graph.Authentication.
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$CloudPcId,

    [Parameter(Mandatory = $false)]
    [string]$ManagedDeviceName
)

$ErrorActionPreference = 'Stop'

Connect-MgGraph -Identity -NoWelcome

$result = [ordered]@{
    cloudPcId         = $CloudPcId
    managedDeviceName = $ManagedDeviceName
    deprovisionedUtc  = $null
    success           = $false
    errorMessage      = $null
}

try {
    Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/cloudPCs/$CloudPcId/deprovision"

    $result.success = $true
    $result.deprovisionedUtc = (Get-Date).ToUniversalTime().ToString('o')
}
catch {
    $result.errorMessage = $_.Exception.Message
}

Write-Output ($result | ConvertTo-Json -Depth 4)
