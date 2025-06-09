#-- This script retrieves Conditional Access Policies (CAP) and exports them to a TXT file.
#-- Requires app-permission Policy.Read.All.
Set-Location -Path C:\Scripts\GraphAPI\tio

#------------------------------------------------------------------------------
#-- Authenticate to Microsoft Graph API using OAuth 2.0 client credentials flow
#------------------------------------------------------------------------------
function AuthGraphAPI
{
    param(
        [string]$tenantId,
        [string]$clientId,
        [string]$clientSecret
    )

    #-- Construct URI
    $uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

    #-- Construct Body
    $body = @{
        client_id     = $clientId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $clientSecret
        grant_type    = "client_credentials"
    }

    #-- Get OAuth 2.0 Token
    $tokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing
    return $tokenRequest.Content | ConvertFrom-Json | Select-Object -ExpandProperty access_token
}

#-- Main script execution
#-- Application (client) ID, tenant ID and secret
$tenantId = "ttttt"
$clientId = "aaaaa"
$clientSecret = "sssss"
$token = AuthGraphAPI -tenantId $tenantId -clientId $clientId -clientSecret $clientSecret
if ($null -eq $token)
{
    Write-Host "Failed to authenticate to Microsoft Graph API."
    exit 1
}

#-- Get all CAP named locations
$method = "GET"
$uri = "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies"
$rFinal = @()

#-- Get the first batch
$oTmp = Invoke-WebRequest -Method $method -Uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -ErrorAction Stop
$oContent = ConvertFrom-Json $oTmp.Content

#-- Save the results to a CSV file
foreach ($policy in $oContent.value)
{
    $oNew = [PSCustomObject]@{
        DisplayName = $policy.displayName
        State = $policy.state
        Created = $policy.createdDateTime
        Modified = $policy.modifiedDateTime
        Sessions = ($policy.sessionControls | ConvertTo-Json -Depth 10).ToString()
        Conditions = ($policy.conditions | ConvertTo-Json -Depth 10).ToString()
        GrantControls = ($policy.grantControls | ConvertTo-Json -Depth 10).ToString()
    }
    $rFinal += $oNew
}
$rFinal | Out-File cap.txt -Encoding ascii