#-- Azure AD OAuth Application Token for Graph API
#-- Get OAuth token for a AAD Application (returned as $token)
Write-Host "Authenticating with Graph-API" -ForegroundColor Green

# Application (client) ID, tenant ID and secret
$clientId = "aaa"
$tenantId = "bbb"
$clientSecret = "ccc"

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

#-- Access Token
$token = ($tokenRequest.Content | ConvertFrom-Json).access_token
$token

<#
#-- Get some Inbox items from the mailbox
$sId = "some email address"
$method = "GET"
$uri = "https://graph.microsoft.com/v1.0/users/" + $sId + "/messages"

#-- Set the variables and init GraphAPI
$rFinal = @()
$nCount = 0

#-- Get the first batch
$oTmp = Invoke-WebRequest -Method $method -Uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -ErrorAction Stop
$oContent = ConvertFrom-Json $oTmp.Content
$oContent.Value[0..2] | ft Subject
#>