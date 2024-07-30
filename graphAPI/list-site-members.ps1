#-- Declare global variables
$script:rgFiles = @()
$script:rgFolders = @()
$script:nCount = 0


#--------------------------------------------------------------------------------------------------
#-- Authenticate with Graph-API
#--------------------------------------------------------------------------------------------------
function udf_AuthGraph
{
    Write-Host "Authenticating with Graph-API" -ForegroundColor Green

    # Application (client) ID, tenant ID and secret
    $sClientId = "e15f4b6f-593c-41c5-b6ef-d49388e4b0e8"
    $sTenantId = "fb007914-6020-4374-977e-21bac5f3f4c8"
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
    $script:token = ($tokenRequest.Content | ConvertFrom-Json).access_token
}


#--------------------------------------------------------------------------------------------------
#-- Get the owners of a group
#--------------------------------------------------------------------------------------------------
function udf_GetGroupOwners
{
    param (
        [string]$sGroupId,
        [string]$sUrl
    )

    $sUri = "https://graph.microsoft.com/v1.0/groups/$sGroupId/owners"
    $sUri
    $oResult = Invoke-RestMethod -Headers @{Authorization = "Bearer $script:token"} -Uri $sUri -Method Get

    #-- Compile the owners
    $rReturn = @()
    if ($oResult.value.Count -gt 0)
    {
        foreach ($oOwner in $oResult.value)
        {
            $rReturn += New-Object PSObject -Property @{
                "Url" = $sUrl
                "DisplayName" = $oOwner.displayName
                "Mail" = $oOwner.mail
            }
        }
    }
    else
    {
        Write-Host "No owners found in the group" -ForegroundColor Yellow
        $rReturn += New-Object PSObject -Property @{
            "Url" = $sUrl
            "DisplayName" = "No owners found in the group"
            "Mail" = ""
        }
    }
    return $rReturn
}


#--------------------------------------------------------------------------------------------------
#-- Get the members of a group
#--------------------------------------------------------------------------------------------------
function udf_GetGroupMembers
{
    param (
        [string]$sGroupId,
        [string]$sUrl
    )

    $sUri = "https://graph.microsoft.com/v1.0/groups/$sGroupId/members"
    $sUri
    $oResult = Invoke-RestMethod -Headers @{Authorization = "Bearer $script:token"} -Uri $sUri -Method Get

    #-- Compile the members
    $rReturn = @()
    if ($oResult.value.Count -gt 0)
    {
        foreach ($oMember in $oResult.value)
        {
            $rReturn += New-Object PSObject -Property @{
                "Url" = $sUrl
                "DisplayName" = $oMember.displayName
                "Mail" = $oMember.mail
            }
        }
    }
    else
    {
        Write-Host "No members found in the group" -ForegroundColor Yellow
        $rReturn += New-Object PSObject -Property @{
            "Url" = $sUrl
            "DisplayName" = "No members found in the group"
            "Mail" = ""
        }
    }
    return $rReturn
}


#--------------------------------------------------------------------------------------------------
#-- Main script
#--------------------------------------------------------------------------------------------------
udf_AuthGraph
$rSites = Import-Csv .\list-url-groupid.csv
$rFinalMembers = @()
$rFinalOwners = @()
foreach ($oSite in $rSites)
{
    Write-Host $oSite.Url -ForegroundColor Green
    #-- Get the members
    $rTmp = udf_GetGroupMembers -sGroupId $oSite.GroupId -sUrl $oSite.Url
    if ($rTmp.Count -gt 0)
    {
        $rFinalMembers += $rTmp
    }

    #-- Get the owners
    $rTmp = udf_GetGroupOwners -sGroupId $oSite.GroupId -sUrl $oSite.Url
    if ($rTmp.Count -gt 0)
    {
        $rFinalOwners += $rTmp
    }
}
$rFinalMembers | Select-Object Url,DisplayName,Mail | Export-Csv hpy-members.csv -NoTypeInformation
$rFinalOwners | Select-Object Url,DisplayName,Mail | Export-Csv hpy-owners.csv -NoTypeInformation