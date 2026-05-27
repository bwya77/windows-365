$AppID = Get-AutomationVariable -Name 'appID' 
$TenantID = Get-AutomationVariable -Name 'tenantID'
$AppSecret = Get-AutomationVariable -Name 'appSecret' 
$verbosepreference = 'continue'
<#
CloudPC.Read.All
Directory.Read.All
#>

function Connect-MSGraphAPI {
    param (
        [Parameter(Mandatory)]
        [string]$AppID,
        [Parameter(Mandatory)]
        [string]$TenantID,
        [Parameter(Mandatory)]
        [string]$AppSecret
    )
    begin {
        $URI = "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token"
        $ReqTokenBody = @{
            Grant_Type    = "client_credentials"
            Scope         = "https://graph.microsoft.com/.default"
            client_Id     = $AppID
            Client_Secret = $AppSecret
        } 
    }
    Process {
        Write-Verbose "Connecting to the Graph API"
        $Response = Invoke-RestMethod -Uri $URI -Method POST -Body $ReqTokenBody -ErrorAction Stop
    }
    End {
        if ($null -eq $Response) {
            Write-Error "Failed to connect to the Graph API. Please check your credentials and try again."
            return $null
        }
        $Response
    }
}
function Get-MSGraphRequest {
    param (
        [Parameter(Mandatory)]
        [string]$Uri,
        [Parameter(Mandatory)]
        [string]$AccessToken
    )
    begin {
        $allPages = @()
        $ReqTokenBody = @{
            Headers = @{
                "Content-Type"  = "application/json"
                "Authorization" = "Bearer $($AccessToken)"
            }
            Method  = "Get"
            Uri     = $Uri
        }
    }
    process {
        Write-Verbose "GET request at endpoint: $Uri"
        $data = Invoke-RestMethod @ReqTokenBody
        if ($null -eq $data) {
            return
        }

        # If response contains a 'value' property (typical Graph collection), accumulate it
        if ($data.PSObject.Properties.Name -contains 'value') {
            $allPages += $data.value
            while ($data.'@odata.nextLink') {
                $ReqTokenBody.Uri = $data.'@odata.nextLink'
                $data = Invoke-RestMethod @ReqTokenBody
                Start-Sleep -Seconds 3
                if ($null -eq $data) { break }
                if ($data.PSObject.Properties.Name -contains 'value') {
                    $allPages += $data.value
                }
                else {
                    $allPages += $data
                }
            }
        }
        else {
            # If the response is already a collection/array, add its items.
            if ($data -is [System.Collections.IEnumerable] -and -not ($data -is [string])) {
                $allPages += $data
            }
            else {
                # Single object returned
                $allPages += $data
            }
        }
    }
    end {
        Write-Verbose "Returning all results"
        $allPages
    }
}
function Send-MSGraphEmail {
    param (
        [Parameter(Mandatory)]
        [string]$UserEmail,
        [Parameter(Mandatory)]
        [string]$Subject,
        [Parameter(Mandatory)]
        [string]$Content,
        [Parameter(Mandatory)]
        [string]$AccessToken,
        [Parameter(Mandatory)]
        [string]$FromEmail
    )
    begin {
        $params = @{
            message = @{
                subject = $Subject
                body = @{
                    contentType = "Text"
                    content = $Content
                }
                toRecipients = @(
                    @{
                        emailAddress = @{
                            address = $UserEmail
                        }
                    }
                )
            }
            saveToSentItems = "false"
        } | ConvertTo-Json -Depth 10
        $sendMailUri = "https://graph.microsoft.com/v1.0/users/$FromEmail/sendMail"
    }
    process {
        Write-Verbose "Sending email to user: $UserEmail with subject: $Subject"
        Invoke-RestMethod -Uri $sendMailUri -Headers @{ "Authorization" = "Bearer $($AccessToken)"; "Content-Type" = "application/json" } -Method POST -Body $params -ErrorAction Stop  
    }
    end {
        Write-Verbose "Email sent successfully to $UserEmail"
    }
}

# Get token for MSGraph API
$tokenResponse = Connect-MSGraphAPI -AppID $AppID -TenantID $TenantID -AppSecret $AppSecret

# Get recently provisioned Cloud PCs, include id, displayName, status, userPrincipalName, provisioningType, provisionedDateTime, provisioningPolicyId and lastRemoteActionResult
$since = (Get-Date).ToUniversalTime().AddMinutes(-5).ToString("yyyy-MM-ddTHH:mm:ss.ffffffZ")
$GraphSplat = @{
    Uri         = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/cloudPCs?`$select=id,managedDeviceName,displayName,status,userPrincipalName,provisioningPolicyId,provisioningType,provisionedDateTime&`$filter=provisionedDateTime ge $since and status eq 'provisioned'"
    AccessToken = $tokenResponse.access_token
}
[array]$CloudPCs = Get-MSGraphRequest @GraphSplat

# Iterate through the list of Cloud PC's
foreach ($PC in $CloudPCs) {
    Write-Verbose "Processing Cloud PC: $($PC.displayName) ($($PC.managedDeviceName)) with status: $($PC.status)"
    
    # If CloudPC is not shared, then send email to the user with the provisioning status and details
    if ($PC.provisioningType -ne "sharedByEntraGroup") {
        # Send email to the user with the provisioning status and details
        Write-Verbose "Sending email to user: $($PC.userPrincipalName) about provisioning status: $($PC.status)"
        [string]$userEmail = $PC.userPrincipalName
        [string]$emailSubject = "Your Cloud PC, `"$($PC.displayName)`", is ready!"
        [string]$emailContent = "Your Cloud PC, `"$($PC.displayName)`", is now ready! Open the Windows App or go to https://windows.cloud.microsoft/ to get started.`n`nProvisioning status: $($PC.status).`nProvisioned on: $($PC.provisionedDateTime)"

        $GraphSplatEmail = @{
            UserEmail   = $userEmail
            Subject     = $emailSubject
            Content     = $emailContent
            AccessToken = $tokenResponse.access_token
            FromEmail   = "brad@windowsfromanywhere.com"
        }
        Write-Verbose "Sending email to user: $userEmail with subject: $emailSubject"
        Send-MSGraphEmail @GraphSplatEmail
    } 
    # If CloudPC is shared by Entra Group
    else {
        # If the target Cloud PC is associated with a shared provisioning policy,
        # check whether another Cloud PC is provisioned
        # If so, skip sending an additional email for this Cloud PC,
        # since the end user would have already received a provisioning status email for
        # the shared provisioning policy.
        #
        # This prevents multiple emails from being sent to the end user when multiple
        # Cloud PCs are provisioned under the same shared provisioning policy within a
        # short time frame.           
        
        [string]$uri = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/cloudPCs?`$filter=provisioningPolicyId eq '$($PC.provisioningPolicyId)' and status eq 'provisioned'&`$select=id,displayName,status,userPrincipalName,provisioningPolicyId,provisioningType,provisionedDateTime"
        $policyCloudPCs = Get-MSGraphRequest -Uri $uri -AccessToken $tokenResponse.access_token
        # If multiple Cloud PCs were provisioned under the same provisioning policy,
        # only send one email for the earliest provisioned Cloud PC and skip the others.
        if ($policyCloudPCs.count -gt 1) {
            $sortedPCs = $policyCloudPCs | Sort-Object -Property provisionedDateTime
            $firstPC = $sortedPCs[0]
            if ($PC.id -ne $firstPC.id) {
                Write-Verbose "Multiple Cloud PCs provisioned under the same provisioning policy: $($PC.provisioningPolicyId). Skipping sending email for Cloud PC: $($PC.displayName) because another Cloud PC was provisioned earlier."
                continue
            }
        }
        [string]$id = $PC.provisioningPolicyId
        [string]$uri = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies/$id`?`$expand=assignments"
        $policyResponse = Get-MSGraphRequest -Uri $uri -AccessToken $tokenResponse.access_token
        [string]$allotmentDisplayName = $policyResponse.assignments[0].target.allotmentDisplayName
        [string]$emailSubject = "Your Cloud PC, `"$allotmentDisplayName`", is ready!"
        [string]$emailContent = "Your Cloud PC, `"$allotmentDisplayName`", is now ready! Open the Windows App or go to https://windows.cloud.microsoft/ to get started.`n`nProvisioning status: $($PC.status).`nProvisioned on: $($PC.provisionedDateTime)"


        [string]$groupID = $policyResponse.assignments[0].target.groupId

        # Lookup the groupID to get the list of users in the group
        [string]$groupUri = "https://graph.microsoft.com/v1.0/groups/$groupId/members"
        $groupMembers = Get-MSGraphRequest -Uri $groupUri -AccessToken $tokenResponse.access_token

        foreach ($member in $groupMembers) {
            $GraphSplatEmail = @{
                UserEmail   = $member.mail
                Subject     = $emailSubject
                Content     = $emailContent
                AccessToken = $tokenResponse.access_token
                FromEmail   = "brad@windowsfromanywhere.com"
            }
            Write-Verbose "Sending email to group member: $($member.mail) about provisioning status: $($PC.status) for shared Cloud PC $allotmentDisplayName"
            Send-MSGraphEmail @GraphSplatEmail
        }
    }
}


