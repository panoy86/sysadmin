#-- This script retrieves all named locations from Microsoft Graph API's Conditional Access Policies (CAP) and exports them to a CSV file.
#-- Requires app-permission Policy.Read.All.

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
$uri = "https://graph.microsoft.com/beta/identity/conditionalAccess/namedLocations"
$rFinal = @()

#-- Get the first batch
$oTmp = Invoke-WebRequest -Method $method -Uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -ErrorAction Stop
$oContent = ConvertFrom-Json $oTmp.Content

#-- Save to a file
foreach($nl in $oContent.value)
{
    if ($nl.'@odata.type' -eq "#microsoft.graph.ipNamedLocation")
    {
        foreach($cidr in $nl.ipRanges)
        {
            $oNew = New-Object PSObject
            $oNew | Add-Member -Type NoteProperty -Name type -Value "ipNamedLocation"
            $oNew | Add-Member -Type NoteProperty -Name displayName -Value $nl.displayName
            $oNew | Add-Member -Type NoteProperty -Name cidrAddress -Value $cidr.cidrAddress
            $oNew | Add-Member -Type NoteProperty -Name countryCode -Value "n/a"
            $oNew | Add-Member -Type NoteProperty -Name isTrusted -Value $nl.isTrusted
            $oNew | Add-Member -Type NoteProperty -Name Created -Value $nl.createdDateTime.ToString()
            $oNew | Add-Member -Type NoteProperty -Name Modified -Value $nl.modifiedDateTime.ToString()
            $rFinal += $oNew
        }
    }
    if ($nl.'@odata.type' -eq "#microsoft.graph.countryNamedLocation")
    {
        foreach($country in $nl.countriesAndRegions)
        {
            $oNew = New-Object PSObject
            $oNew | Add-Member -Type NoteProperty -Name type -Value "countryNamedLocation"
            $oNew | Add-Member -Type NoteProperty -Name displayName -Value $nl.displayName
            $oNew | Add-Member -Type NoteProperty -Name cidrAddress -Value "n/a"
            $oNew | Add-Member -Type NoteProperty -Name countryCode -Value $country
            $oNew | Add-Member -Type NoteProperty -Name isTrusted -Value "n/a"
            $oNew | Add-Member -Type NoteProperty -Name Created -Value $nl.createdDateTime.ToString()
            $oNew | Add-Member -Type NoteProperty -Name Modified -Value $nl.modifiedDateTime.ToString()
            $rFinal += $oNew
        }
    }
}
$rFinal | Export-Csv .\cap-named-locations.csv
$rFinal | Format-Table -AutoSize