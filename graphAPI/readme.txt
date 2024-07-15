#-- Azure AD OAuth Application Token for Graph API
#-- Get OAuth token for a AAD Application (returned as $token)
Write-Host "Authenticating with Graph-API" -ForegroundColor Green

# Application (client) ID, tenant ID and secret
$sClientId = ""
$sTenantId = ""
$sClientSecret = "Kgz8Q~JZrw9036WRdHrTtC_5HRfS5IOf2D6YWblU"

#-- Construct URI
$uri = "https://login.microsoftonline.com/$sTenantId/oauth2/v2.0/token"

#-- Construct Body
$body = @{
    client_id     = $sClientId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $sClientSecret
    grant_type    = "client_credentials"
}

#-- Get OAuth 2.0 Token
$tokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing

#-- Access Token
$token = ($tokenRequest.Content | ConvertFrom-Json).access_token
