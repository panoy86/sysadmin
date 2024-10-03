#-- Creted this script to go thru inactive OneDrive accounts and remove folder and file
#-- retention labels that prevent m365 from auto-deleting the content once it has been unlicensed.

Set-Location C:\Scripts\OneDrive\
$sWorkingFile = ".\list-onedrive-urls.txt"

$script:rListFiles = @()
$script:rListFolders = @()
$script:nCtr = 0
$script:nOneDriveProcessed = 0

#---------------------------------------------------------------------------------------------------
#-- Function to authenticate to the Graph API
#---------------------------------------------------------------------------------------------------
function udf_AuthGraphApi
{
    param
    (
        [Parameter()]$bSilent = $true
    )

    $sTenantId = "XXX"
    $sClientId = "YYY"
    $sSecretB64 = "ZZZ"
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


#---------------------------------------------------------------------------------------------------
#-- Function to get the retention label (if any) applied to a file or folder
#---------------------------------------------------------------------------------------------------
function udf_GetRetentionLabel
{
    Param
    (
        [Parameter(Mandatory=$true)]$sDriveId,
        [Parameter(Mandatory=$true)]$sItemId
    )
    
    #-- Get the retention label
    $sUri = "https://graph.microsoft.com/v1.0/drives/$($sDriveId)/items/$($sItemId)/retentionLabel"
    Try {$oResult = Invoke-WebRequest -Headers $oAuthHeader -Uri $sUri -ErrorAction Stop}
    Catch {$oResult = $null}
    if ($null -ne $oResult)
    {
        $oRetentionLabel = $oResult.Content | ConvertFrom-Json
        return $oRetentionLabel
    }
    else {return $null}
}


#---------------------------------------------------------------------------------------------------
#-- Recursive function to iterate through all child items from a top-level folder
#---------------------------------------------------------------------------------------------------
function udf_GetChildItems
{

    Param
    (
        [Parameter(Mandatory=$true)]$sDriveId,
        [Parameter(Mandatory=$true)]$sFolderId
    )

    #-- See if we need to auth again
    $script:nCtr++
    if (($script:nCtr % 200) -eq 0) {udf_AuthGraphApi}

    #-- Get the child items
    $sUri = "https://graph.microsoft.com/v1.0/drives/$($sDriveId)/items/$($sFolderId)/children"
    do
    {
        $oResult = Invoke-WebRequest -Headers $oAuthHeader -Uri $sUri -Verbose:$VerbosePreference -ErrorAction Stop
        $sUri = $oResult.'@odata.nextLink'
        #-- Add delay to avoid SPO throttling
        Start-Sleep -Milliseconds 500
        $rItems += $oResult.Content | ConvertFrom-Json
    } while ($sUri)
    
    #-- Process each child item, there are 3 types: folder, file, notebook
    $rFolders = @()
    foreach ($oItem in $rItems.value)
    {
        if ($null -ne $oItem.folder)
        {
            if ($null -eq $oItem.package)
            {
                $script:rListFolders += $oItem
                $rFolders += $oItem
            }
            #-- Notebooks - just put them under files
            else {$script:rListFiles += $oItem}
        }
        else {$script:rListFiles += $oItem}
    }

    #-- Recursively process the folders from this level
    foreach ($oFolder in $rFolders)
    {
        Write-Progress $oFolder.name
        udf_GetChildItems -sDriveId $sDriveId -sFolderId $oFolder.id
    }
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
    $sUri = 'https://graph.microsoft.com/v1.0/sites/paypal-my.sharepoint.com:/' + $sRelativeUrl
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


#---------------------------------------------------------------------------------------------------
#-- Function to remove the retention label from a file or folder
#---------------------------------------------------------------------------------------------------
function udf_RemoveRetentionLabel
{
    Param
    (
        [Parameter(Mandatory=$true)]$sDriveId,
        [Parameter(Mandatory=$true)]$sItemId,
        [Parameter(Mandatory=$true)]$sRetentionLabel
    )
    
    #-- Remove the retention label
    #Set-Variable -Name oAuthHeader -Scope Local -Value @{'Authorization'="Bearer $token"}
    $oAuthHeader = @{'Authorization'="Bearer $token"}
    $sUri = "https://graph.microsoft.com/v1.0/drives/$($sDriveId)/items/$($sItemId)/retentionLabel"
    #$oResult = Invoke-RestMethod -Method Delete -Headers $oAuthHeader -Uri $sUri -Verbose:$VerbosePreference -ErrorAction Stop
    #Try {$oResult = Invoke-RestMethod -Method Delete -Headers $oAuthHeader -Uri $sUri -ErrorAction SilentlyContinue} Catch {$oResult = $null}
    $oResult = $true
    Try {$null = Invoke-RestMethod -Method Delete -Headers $oAuthHeader -Uri $sUri -ErrorAction SilentlyContinue} Catch {$oResult = $false}
    return $oResult

    <#-- To set a retention label, use this:
    Set-Variable -Name oAuthHeader -Scope Local -Value @{'Authorization'="Bearer $token";'Content-Type'='application/json'}
    Set-Variable -Name oBody -Scope Local -Value @{'name'=$sRetentionLabel}
    $sUri = "https://graph.microsoft.com/v1.0/drives/$($sDriveId)/items/$($sItemId)/retentionLabel"
    $oResult = Invoke-RestMethod -Method Delete -Headers $oAuthHeader -Uri $sUri -Verbose:$VerbosePreference -ErrorAction Stop -Body (ConvertTo-Json $oBody)
    #>
}   


#---------------------------------------------------------------------------------------------------
#-- Main script
#---------------------------------------------------------------------------------------------------
udf_AuthGraphApi
$rWork = Get-Content $sWorkingFile

#-- Loop thru our working list
$script:nCtr = 0
foreach ($sUrl in $rWork[33..100])
{
    #-- Attempt to find the drive id
    $sRelativeUrl = $sUrl.Substring(32)
    Try {$sDriveId = udf_GetDriveId -sRelativeUrl $sRelativeUrl}
    Catch {$sDriveId = $null}
    if ($null -eq $sDriveId)
    {
        Write-Host "Error:" -ForegroundColor Red -NoNewline
        Write-Host " Drive ID not found for $($sUrl)"
        continue
    }

    #-- Reset our variables
    $script:rListFiles = @()
    $script:rListFolders = @()
    $script:nCtr = 0

    #-- Get the top-level folder
    $sUri = "https://graph.microsoft.com/v1.0/drives/$($sDriveId)/root"
    $oAuthHeader = @{'Authorization'="Bearer $token"}
    Try {$oResult = Invoke-RestMethod -Method Get -Headers $oAuthHeader -Uri $sUri -ErrorAction Stop}
    Catch {$oResult = $null}
    if ($null -ne $oResult)
    {
        $sRootId = $oResult.id
        Write-Host "Root: $($oResult.webUrl)" -ForegroundColor Yellow

        #-- Get the list of files and folders
        Write-Host "Getting all files and folders..." -ForegroundColor Green
        udf_GetChildItems -sDriveId $sDriveId -sFolderId $sRootId
    }
    #-- Save the data temporarily
    $script:rListFiles | ForEach-Object {$_ | Add-Member -MemberType NoteProperty -Name "RetentionLabel" -Value "" -Force}
    $script:rListFiles | ForEach-Object {$_ | Add-Member -MemberType NoteProperty -Name "isLabelAppliedExplicitly" -Value "" -Force}
    $script:rListFiles | Export-Csv .\tmp_files.csv -NoTypeInformation
    $script:rListFolders | ForEach-Object {$_ | Add-Member -MemberType NoteProperty -Name "RetentionLabel" -Value "" -Force}
    $script:rListFolders | ForEach-Object {$_ | Add-Member -MemberType NoteProperty -Name "isLabelAppliedExplicitly" -Value "" -Force}
    $script:rListFolders | Export-Csv .\tmp_folders.csv -NoTypeInformation

    #-- Find any retention labels applied to the files and folders
    $nCountFiles = 0
    Write-Host "Finding file-applied retention labels..." -ForegroundColor Green
    foreach ($oFile in $script:rListFiles)
    {
        $oRetentionLabel = udf_GetRetentionLabel -sDriveId $sDriveId -sItemId $oFile.id
        if ($null -ne $oRetentionLabel.name)
        {
            Write-Progress "$($oFile.name) $($oRetentionLabel.name)"
            $oFile.RetentionLabel = $oRetentionLabel.name
            $oFile.isLabelAppliedExplicitly = $oRetentionLabel.isLabelAppliedExplicitly
            $nCountFiles++
        }
        #-- Re-authenticate every 200 items
        $script:nCtr++
        if (($script:nCtr % 200) -eq 0) {udf_AuthGraphApi}
    }
    Write-Host "Files with retention labels: $($nCountFiles)" -ForegroundColor Yellow
    $script:rListFiles | Export-Csv .\tmp_files.csv -NoTypeInformation

    #-- Do the folders
    $nCountFolders = 0
    Write-Host "Finding folder-applied retention labels..." -ForegroundColor Green
    foreach ($oFolder in $script:rListFolders)
    {
        $oRetentionLabel = udf_GetRetentionLabel -sDriveId $sDriveId -sItemId $oFolder.id
        if ($null -ne $oRetentionLabel.name)
        {
            Write-Progress "   $($oFolder.name) $($oRetentionLabel.name)"
            $oFolder.RetentionLabel = $oRetentionLabel.name
            $oFolder.isLabelAppliedExplicitly = $oRetentionLabel.isLabelAppliedExplicitly
            $nCountFolders++
        }
        #-- Re-authenticate every 200 items
        $script:nCtr++
        if (($script:nCtr % 200) -eq 0) {udf_AuthGraphApi}
    }
    Write-Host "Folders with retention labels: $($nCountFolders)" -ForegroundColor Yellow
    $script:rListFolders | Export-Csv .\tmp_folders.csv -NoTypeInformation
    
    #-- Increment our total OneDrive accounts processed
    if ($nCountFolders -gt 0 -or $nCountFiles -gt 0) {$script:nOneDriveProcessed++}

    #-- Remove the retention labels
    Write-Host "Removing retention labels..." -ForegroundColor Green
    foreach ($oFile in $script:rListFiles)
    {
        if (([string]$oFile.RetentionLabel).Trim().Length -gt 0)
        {
            #if ([string]$oFile.isLabelAppliedExplicitly -eq "True")
            if ($true)
            {
                Write-Progress "Removing file label: $($oFile.name) $($oFile.RetentionLabel)"
                if (-not (udf_RemoveRetentionLabel -sDriveId $sDriveId -sItemId $oFile.id -sRetentionLabel $oFile.RetentionLabel))
                {
                    Write-Host "Error:" -ForegroundColor Red -NoNewline
                    Write-Host " Unable to remove retention label from $($oFile.name)"
                }

                #-- Re-authenticate every 200 items
                $script:nCtr++
                if (($script:nCtr % 200) -eq 0) {udf_AuthGraphApi}
            }
        }
    }
    foreach ($oFolder in $script:rListFolders)
    {
        if (([string]$oFolder.RetentionLabel).Trim().Length -gt 0)
        {
            #if ([string]$oFolder.isLabelAppliedExplicitly -eq "True")
            if ($true)
            {
                Write-Progress "Removing folder label: $($oFolder.name) $($oFolder.RetentionLabel)"
                if (-not (udf_RemoveRetentionLabel -sDriveId $sDriveId -sItemId $oFolder.id -sRetentionLabel $oFolder.RetentionLabel))
                {
                    Write-Host "Error:" -ForegroundColor Red -NoNewline
                    Write-Host " Unable to remove retention label from $($oFolder.name)"
                }

                #-- Re-authenticate every 200 items
                $script:nCtr++
                if (($script:nCtr % 200) -eq 0) {udf_AuthGraphApi}
            }
        }
    }

    #-- Only process 3 OneDrive accounts per run
    if ($script:nOneDriveProcessed -ge 3) {break}
    Write-Progress -Completed
}
Write-Progress -Completed