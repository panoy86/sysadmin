#-- Gets all the holds in a given OneDrive account

#---------------------------------------------------------------------------------------------------
#-- Function to authenticate to the Graph API
#---------------------------------------------------------------------------------------------------
function udf_AuthGraphApi
{
    param
    (
        [Parameter()]$bSilent = $true
    )

    $sTenantId = ""
    $sClientId = ""
    $sSecretB64 = ""
    $bytes = [System.Convert]::FromBase64String($sSecretB64)
    $sClientSecret = [System.Text.Encoding]::UTF8.GetString($bytes)

    #-- Get the token
    if (-not $bSilent) {Write-Host "Authenticating to the Graph API..." -ForegroundColor Green}
    $sUri = "https://login.microsoftonline.com/$sTenantId/oauth2/v2.0/token"
    $oBody = @{
        client_id = $sClientId
        scope = "https://graph.microsoft.com/.default"
        client_secret = $sClientSecret
        grant_type = "client_credentials"
    }
    #-- Get OAuth 2.0 Token
    $tokenRequest = Invoke-WebRequest -Method Post -Uri $sUri -ContentType "application/x-www-form-urlencoded" -Body $oBody -UseBasicParsing
    #-- Access Token
    Set-Variable -Name token -Scope Script -Value ($tokenRequest.Content | ConvertFrom-Json).access_token
}

#--------------------------------------------------------------------------------------------------
#-- Function to get the drive id
#--------------------------------------------------------------------------------------------------
function udf_GetDriveId
{
    Param
    (
        [Parameter(Mandatory=$true)]$sRelativeUrl
    )

    #-- Get the site id
    $sUri = 'https://graph.microsoft.com/v1.0/sites/domain-my.sharepoint.com:/' + $sRelativeUrl
    $oAuthHeader = @{'Authorization'="Bearer $token"}
    Try {$oResult = Invoke-RestMethod -Method Get -Headers $oAuthHeader -Uri $sUri -Verbose:$VerbosePreference -ErrorAction SilentlyContinue}
    Catch {return $null}
    $sSiteId = $oResult.id

    #-- Get the drive id
    $sUri = "https://graph.microsoft.com/v1.0/sites/$($sSiteId)/drive"
    $oAuthHeader = @{'Authorization'="Bearer $token"}
    Try {$oResult = Invoke-RestMethod -Method Get -Headers $oAuthHeader -Uri $sUri -Verbose:$VerbosePreference -ErrorAction Stop}
    Catch {return $null}
    $sDriveId = $oResult.id
    return $sDriveId
}

function udf_GetHolds
{
    param
    (
        [Parameter(Mandatory=$true)]$sDriveId
    )

    #-- Get the holds
    $oAuthHeader = @{'Authorization'="Bearer $token"}
    $sUri = "https://graph.microsoft.com/v1.0/drives/$($sDriveId)/holds"
    $oResult = Invoke-WebRequest -Headers $oAuthHeader -Uri $sUri -ErrorAction Stop
    #Try {$oResult = Invoke-WebRequest -Headers $oAuthHeader -Uri $sUri -ErrorAction Stop}
    #Catch {$oResult = $null}
    $oResult
    if ($null -ne $oResult)
    {
        $oHolds = $oResult.Content | ConvertFrom-Json
        return $oHolds
    }
    else {return $null}
}

#-- Get the site-id based on the URL
function udf_GetSiteId
{
    Param
    (
        [Parameter(Mandatory=$true)]$sRelativeUrl
    )

    #-- Get the site id
    $sUri = 'https://graph.microsoft.com/v1.0/sites/domain-my.sharepoint.com:/' + $sRelativeUrl
    $oAuthHeader = @{'Authorization'="Bearer $token"}
    Try {$oResult = Invoke-WebRequest -Method GET -Headers $oAuthHeader -Uri $sUri -ErrorAction Stop}
    Catch {return $null}
    $sSiteId = ($oResult.Content | ConvertFrom-Json).id
    return $sSiteId
}


function udf_GetLists
{
    param
    (
        [Parameter(Mandatory=$true)]$sSiteId
    )

    #-- Get the lists
    $oAuthHeader = @{'Authorization'="Bearer $token"}
    $sUri = "https://graph.microsoft.com/v1.0/sites/$($sSiteId)/lists"
    $oResult = Invoke-WebRequest -Method GET -Headers $oAuthHeader -Uri $sUri -ErrorAction Stop  #;$oResult
    if ($null -ne $oResult)
    {
        $oReturn = $oResult.Content | ConvertFrom-Json
        return $oReturn
    }
    else {return $null}
}

function udf_GetItems
{
    param
    (
        [Parameter(Mandatory=$true)]$sSiteId,
        [Parameter(Mandatory=$true)]$sListId
    )

    #-- Get the items
    $oAuthHeader = @{'Authorization'="Bearer $token"}
    $sUri = "https://graph.microsoft.com/v1.0/sites/$($sSiteId)/lists/$($sListId)/items"
    $oResult = Invoke-WebRequest -Method GET -Headers $oAuthHeader -Uri $sUri -ErrorAction Stop
    if ($null -ne $oResult)
    {
        $oReturn = $oResult.Content | ConvertFrom-Json
        #$oReturn = $oResult.Content
        return $oReturn
    }
    else {return $null}
}

#-- Main
udf_AuthGraphApi -bSilent $false
$r = Get-Content .\list-onedrive-urls.txt
foreach ($sUrl in $r[0])
{
    #-- Attempt to find the site-id and list the lists
    Write-Progress -Activity $sUrl
    Write-Host "Processing $sUrl" -ForegroundColor Yellow
    $sRelativeUrl = $sUrl.Substring(32)
    $sSiteId = udf_GetSiteId -sRelativeUrl $sRelativeUrl
    
    #-- Get the lists details
    if ($null -ne $sSiteId)
    {
        [array]$rLists = udf_GetLists -sSiteId $sSiteId
        foreach ($oItem in $rLists.value)
        {
            if ($oItem.Name -match "DiscoveryPreservationHolds" -or $oItem.Name -match "PreservationHoldLibrary" -or $oItem.Name -like "Collection*")
            {
                Write-Host $oItem.Name
                Write-Host $oItem.Description -ForegroundColor Cyan
                [array]$rItems = udf_GetItems -sSiteId $sSiteId -sListId $oItem.Id
                $rItems.value.Count
            }
        }
    }
}
Write-Progress -Activity "End" -Completed
