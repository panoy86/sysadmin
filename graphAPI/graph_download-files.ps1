#-- Script that uses the list of files from "graph_get-list-of-folders-files.ps1"
#-- Adds column "download" to designate files to download

Set-Location C:\Scripts\OneDrive\Temp
$sWorkingFile = ".\tmp_files.csv"

#---------------------------------------------------------------------------------------------------
#-- Function to authenticate to the Graph API
#---------------------------------------------------------------------------------------------------
function udf_AuthGraphApi
{
    param
    (
        [Parameter()]$bSilent = $true
    )

    $sTenantId = "fb007914-6020-4374-977e-21bac5f3f4c8"
    $sClientId = "2417f530-7fe2-48ef-85c7-d5f09b9564d7"
    $sSecretB64 = "OWx4OFF+MGlrSnNabVVUYS1wNXZXfjVTS0hqVE9lY3ozX29BaGNqTQ=="
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


function udf_DownloadFile
{
    param
    (
        [Parameter()]$sName,
        [Parameter()]$sDriveId,
        [Parameter()]$sItemId
    )

    #-- Get the file download URL
    $oAuthHeader = @{'Authorization'="Bearer $token"; "Content-Type"= "application/json"}
    $sUri = "https://graph.microsoft.com/v1.0/drives/$($sDriveId)/items/$($sItemId)"
    Try {$oResult = Invoke-WebRequest -Method GET -Headers $oAuthHeader -Uri $sUri -ErrorAction Stop}
    Catch {$oResult = $null}
    if ($null -ne $oResult)
    {
        #-- Download the file
        $oTmp = $oResult.Content | ConvertFrom-Json
        Invoke-WebRequest -Uri $oTmp."@microsoft.graph.downloadUrl" -OutFile ".\$sName"
        return $true
    }
    else {return $false}
}


#-- Main script
udf_AuthGraphApi -bSilent $false
$rWork = Import-Csv $sWorkingFile
foreach ($oItem in $rWork)
{
    if ($oItem.download -eq "true")
    {
        Write-Host "Downloading: $($oItem.name)" -ForegroundColor Yellow
        $nLength = $oItem.parentReference.Length
        $sTmp = $oItem.parentReference.Substring(2, $nLength - 3)
        $sTmp = $sTmp -replace "; ", "`r`n"
        $oTmp = ConvertFrom-StringData -StringData $sTmp 
        udf_DownloadFile -sName $oItem.name -sDriveId $oTmp.driveId -sItemId $oItem.id
    }
}
