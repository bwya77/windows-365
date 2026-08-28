<#
.SYNOPSIS
    Finds Windows 365 Reserve Cloud PCs that have never been signed into, or have not been signed
    into recently, so they can be flagged (alert) or later deprovisioned via a Logic App.

.DESCRIPTION
    Designed to run as an Azure Automation runbook using a system-assigned managed identity.
    Connects to Microsoft Graph, finds the Windows 365 Reserve provisioning policy (or policies),
    pulls every Cloud PC associated with it, and checks each one's real-time sign-in status.

    A Cloud PC is considered "idle" and included in the output when BOTH are true:
      1. It has been provisioned for at least $MinProvisionedHours (default 24).
      2. It has never signed in, OR its last sign-in was at least $MinIdleHours ago (default 24).

    Condition 1 (provisionedDateTime) matters because a user can provision a Reserve Cloud PC and
    never sign in at all -- in that case there is no sign-in timestamp to compare against, so we
    still need a way to catch it. Requiring 24 hours since provisioning also avoids false positives
    on brand-new Cloud PCs that simply haven't been signed into yet because they were just created.

    Output is a single JSON object (via Write-Output) intended to be consumed by a Logic App
    through the "Get Job Output" action on an Automation Account job, then parsed with a
    Parse JSON action and used to compose/send an email (or to drive an approval step before
    calling Invoke-DeprovisionWindows365ReserveCloudPc.ps1 / a direct deprovision Graph call).

.PARAMETER ReservePolicyDisplayName
    Display name of the Windows 365 Reserve provisioning policy. Used only as a fallback/filter
    when multiple provisioning policies of type 'reserve' exist and you want to target one by name.
    If omitted, every provisioning policy with provisioningType 'reserve' is included.

.PARAMETER MinProvisionedHours
    Minimum hours since provisionedDateTime before a Cloud PC is eligible to be flagged. Default 24.

.PARAMETER MinIdleHours
    Minimum hours since last sign-in (or "never signed in") before a Cloud PC is flagged. Default 24.

.NOTES
    Required Microsoft Graph Application permission (granted to the Automation Account's
    managed identity): CloudPC.Read.All (CloudPC.ReadWrite.All if you extend this script to
    also deprovision).

    Required PowerShell module in the Automation Account: Microsoft.Graph.Authentication and Microsoft.Graph.Applications
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$ReservePolicyDisplayName = "W365-Reserve",

    [Parameter(Mandatory = $false)]
    [int]$MinProvisionedHours = 24,

    [Parameter(Mandatory = $false)]
    [int]$MinIdleHours = 24
)

$ErrorActionPreference = 'Stop'
Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Applications


#region Connect to Microsoft Graph using the Automation Account's managed identity
Connect-MgGraph -Identity -NoWelcome
#endregion

#region Helper: page through a Graph collection response, following @odata.nextLink
function Get-MgGraphAllPages {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,

        # Set when the query combines $count=true with an advanced filter operator (e.g. "ne")
        # against an enum-backed property -- Graph requires ConsistencyLevel: eventual in that case.
        [switch]$ConsistencyLevelEventual
    )

    $results = [System.Collections.Generic.List[object]]::new()
    $nextUri = $Uri
    $headers = if ($ConsistencyLevelEventual) { @{ ConsistencyLevel = 'eventual' } } else { @{} }

    while ($nextUri) {
        $response = Invoke-MgGraphRequest -Method GET -Uri $nextUri -Headers $headers
        if ($response.value) {
            $results.AddRange($response.value)
        }
        $nextUri = $response.'@odata.nextLink'
    }

    return $results
}
#endregion

function Get-MgGraphReportJson {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri
    )

    $tempFile = [System.IO.Path]::GetTempFileName()
    try {
        Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputFilePath $tempFile | Out-Null
        $raw = Get-Content -Path $tempFile -Raw
        if ([string]::IsNullOrWhiteSpace($raw)) {
            return $null
        }
        return $raw | ConvertFrom-Json
    }
    finally {
        Remove-Item -Path $tempFile -Force -ErrorAction SilentlyContinue
    }
}


#region 1. Find the Reserve provisioning policy/policies
$policyFilter = "provisioningType eq 'reserve'"
$policyUri = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies?`$filter=$policyFilter"
$reservePolicies = Get-MgGraphAllPages -Uri $policyUri

if ($ReservePolicyDisplayName) {
    $reservePolicies = $reservePolicies | Where-Object { $_.displayName -eq $ReservePolicyDisplayName }
}

if (-not $reservePolicies -or $reservePolicies.Count -eq 0) {
    $noPolicyResult = [ordered]@{
        generatedUtc     = (Get-Date).ToUniversalTime().ToString('o')
        thresholdHours   = [ordered]@{
            minProvisionedHours = $MinProvisionedHours
            minIdleHours        = $MinIdleHours
        }
        reservePolicies  = @()
        totalCloudPcs    = 0
        idleCount        = 0
        idleCloudPcs     = @()
        emailBodyHtml    = "<p>No Windows 365 Reserve provisioning policy was found.</p>"
    }
    Write-Output ($noPolicyResult | ConvertTo-Json -Depth 6)
    return
}
#endregion

#region 2. Get every Cloud PC tied to each Reserve policy
$select = 'id,managedDeviceId,managedDeviceName,status,provisioningType,userPrincipalName,provisionedDateTime,sharedDeviceDetail,connectivityResult'
$allReserveCloudPcs = [System.Collections.Generic.List[object]]::new()

foreach ($policy in $reservePolicies) {
    $cloudPcFilter = "(provisioningPolicyId eq '$($policy.id)') and (status ne 'notProvisioned')"
    $cloudPcUri = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/cloudPCs" +
                  "?`$filter=$cloudPcFilter&`$select=$select&`$orderBy=provisionedDateTime desc&`$count=true"

    # ConsistencyLevel: eventual is required when combining $count=true with an advanced query
    # operator like "ne" against an enum-backed property such as status.
    $cloudPcs = Get-MgGraphAllPages -Uri $cloudPcUri -ConsistencyLevelEventual
    foreach ($cloudPc in $cloudPcs) {
        # Tag each Cloud PC with the policy it came from so multi-policy output is traceable
        $cloudPc | Add-Member -NotePropertyName 'reservePolicyId' -NotePropertyValue $policy.id -Force
        $cloudPc | Add-Member -NotePropertyName 'reservePolicyName' -NotePropertyValue $policy.displayName -Force
        $allReserveCloudPcs.Add($cloudPc)
    }
}
#endregion

#region 3. For each Cloud PC, get real-time sign-in status and evaluate idle criteria
$nowUtc = (Get-Date).ToUniversalTime()
$idleCloudPcs = [System.Collections.Generic.List[object]]::new()

foreach ($cloudPc in $allReserveCloudPcs) {

    if (-not $cloudPc.provisionedDateTime) {
        # No provisionedDateTime at all -- can't evaluate, skip rather than guess
        continue
    }

    $provisionedUtc = [DateTime]::Parse($cloudPc.provisionedDateTime, $null, [System.Globalization.DateTimeStyles]::AdjustToUniversal -bor [System.Globalization.DateTimeStyles]::AssumeUniversal)
    $hoursSinceProvisioned = ($nowUtc - $provisionedUtc).TotalHours

    if ($hoursSinceProvisioned -lt $MinProvisionedHours) {
        # Too new -- give it a chance to be signed into before flagging
        continue
    }

    # Real-time sign-in status/report call, one per Cloud PC
    $signInUri = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/reports/getRealTimeRemoteConnectionStatus(cloudPcId='$($cloudPc.id)')"
    $signInReport = Get-MgGraphReportJson -Uri $signInUri

    $lastActiveTimeUtc = $null
    $daysSinceLastSignIn = $null
    $signInStatus = $null

    if ($signInReport.TotalRowCount -gt 0 -and $signInReport.Values.Count -gt 0) {
        # Columnar Schema/Values response -- map column name to its index, then read the first row
        $columns = @{}
        for ($i = 0; $i -lt $signInReport.Schema.Count; $i++) {
            $columns[$signInReport.Schema[$i].Column] = $i
        }

        $row = $signInReport.Values[0]
        $signInStatus = $row[$columns['SignInStatus']]
        $daysSinceLastSignIn = $row[$columns['DaysSinceLastSignIn']]
        $rawLastActiveTime = $row[$columns['LastActiveTime']]

        if ($rawLastActiveTime) {
            $lastActiveTimeUtc = [DateTime]::Parse($rawLastActiveTime, $null, [System.Globalization.DateTimeStyles]::AdjustToUniversal -bor [System.Globalization.DateTimeStyles]::AssumeUniversal)
        }
    }

    $neverSignedIn = ($null -eq $lastActiveTimeUtc)
    $hoursSinceLastSignIn = if ($neverSignedIn) { $null } else { ($nowUtc - $lastActiveTimeUtc).TotalHours }

    $isIdle = $neverSignedIn -or ($hoursSinceLastSignIn -ge $MinIdleHours)

    if ($isIdle) {
        $assignedUser = if ($cloudPc.userPrincipalName) {
            $cloudPc.userPrincipalName
        } elseif ($cloudPc.sharedDeviceDetail -and $cloudPc.sharedDeviceDetail.assignedToUserPrincipalName) {
            $cloudPc.sharedDeviceDetail.assignedToUserPrincipalName
        } else {
            $null
        }

        $idleCloudPcs.Add([ordered]@{
            cloudPcId              = $cloudPc.id
            managedDeviceName      = $cloudPc.managedDeviceName
            managedDeviceId        = $cloudPc.managedDeviceId
            status                 = $cloudPc.status
            reservePolicyId        = $cloudPc.reservePolicyId
            reservePolicyName      = $cloudPc.reservePolicyName
            assignedUserPrincipalName = $assignedUser
            provisionedDateTimeUtc = $provisionedUtc.ToString('o')
            hoursSinceProvisioned  = [math]::Round($hoursSinceProvisioned, 1)
            neverSignedIn          = $neverSignedIn
            signInStatus           = $signInStatus
            lastActiveTimeUtc      = if ($lastActiveTimeUtc) { $lastActiveTimeUtc.ToString('o') } else { $null }
            daysSinceLastSignIn    = $daysSinceLastSignIn
            hoursSinceLastSignIn   = if ($null -ne $hoursSinceLastSignIn) { [math]::Round($hoursSinceLastSignIn, 1) } else { $null }
        })
    }
}
#endregion

#region 4. Build a ready-to-email HTML table (Logic App "Send an email" can use this directly)
$tableRows = $idleCloudPcs | ForEach-Object {
    $lastSignInDisplay = if ($_.neverSignedIn) { 'Never signed in' } else { $_.lastActiveTimeUtc }
    "<tr><td>$($_.managedDeviceName)</td><td>$($_.assignedUserPrincipalName)</td><td>$($_.reservePolicyName)</td><td>$($_.provisionedDateTimeUtc)</td><td>$lastSignInDisplay</td></tr>"
}

$emailBodyHtml = @"
<p>The following Windows 365 Reserve Cloud PCs have been provisioned for at least $MinProvisionedHours hour(s) and have not signed in during the last $MinIdleHours hour(s) (or have never signed in):</p>
<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif;font-size:13px;">
<tr style="background-color:#0078D4;color:#ffffff;">
<th>Device Name</th><th>Assigned User</th><th>Reserve Policy</th><th>Provisioned (UTC)</th><th>Last Sign-In (UTC)</th>
</tr>
$($tableRows -join "`n")
</table>
"@
#endregion

#region 5. Emit a single JSON object for Logic Apps to parse
$result = [ordered]@{
    generatedUtc    = $nowUtc.ToString('o')
    thresholdHours  = [ordered]@{
        minProvisionedHours = $MinProvisionedHours
        minIdleHours        = $MinIdleHours
    }
    reservePolicies = @($reservePolicies | ForEach-Object { [ordered]@{ id = $_.id; displayName = $_.displayName } })
    totalCloudPcs   = $allReserveCloudPcs.Count
    idleCount       = $idleCloudPcs.Count
    idleCloudPcs    = $idleCloudPcs
    emailBodyHtml   = $emailBodyHtml
}

Write-Output ($result | ConvertTo-Json -Depth 6)
#endregion
