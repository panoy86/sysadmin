
#-- Get the site id
$sRelativeSiteUrl = "/sites/nnn-publichrfolder"
#-- To get site-id manually: https://<url>/_api/site/id
$sUri = "https://graph.microsoft.com/v1.0/sites/<tenant>.sharepoint.com:" + $sRelativeSiteUrl
$oResult = Invoke-RestMethod -Headers @{Authorization = "Bearer $token"} -Uri $sUri -Method Get
Set-Variable -Name sSiteId -Value $oResult.id -Scope Script

#-- Declare global variables
$script:rgFiles = @()
$script:rgFolders = @()
$script:rgPerms = @()

#--------------------------------------------------------------------------------------------------
#-- Get permissions for a file/folder
#--------------------------------------------------------------------------------------------------
function udf_GetItemPermissions
{
    param (
        [string]$sSiteId,
        [string]$sFileId
    )
    #-- Get the permissions
    $sUri = "https://graph.microsoft.com/v1.0/sites/$sSiteId/drive/items/$sFileId/permissions"
    $oResult = Invoke-RestMethod -Headers @{Authorization = "Bearer $token"} -Uri $sUri -Method Get
    $rPermissions = $oResult.value

    #-- Create our return list
    $rReturn = @()
    foreach ($oPerm in $rPermissions)
    {
        $oNew = New-Object PSObject @{
            roles = ($oPerm.roles -join ',').ToString()
            displayName = $oPerm.grantedTo.user.displayName
            email = $oPerm.grantedTo.user.email
        }
        $rReturn += $oNew
    }
    return $rReturn
}

#--------------------------------------------------------------------------------------------------
#-- Recursively get all the child items
#--------------------------------------------------------------------------------------------------
function udf_GetDriveItems
{
    param (
        [string]$sSiteId,
        [string]$sFolderId
    )

    #-- Get all folders/files from this drive
    $sUri = "https://graph.microsoft.com/v1.0/sites/$sSiteId/drive/items/$sFolderId/children"
    $sUri = "https://graph.microsoft.com/v1.0/drives/$sFolderId/list/items"
    #$sUri
    $oResult = Invoke-RestMethod -Headers @{Authorization = "Bearer $token"} -Uri $sUri -Method Get
    $rItems = $oResult.value
    while ($null -ne $oResult."@odata.nextLink")
    {
        $sUri = $oResult."@odata.nextLink"
        $oResult = Invoke-RestMethod -Headers @{Authorization = "Bearer $token"} -Uri $sUri -Method Get
        $rItems += $oResult.value
    }
    Write-Host "Total: $($rItems.Count)"

    #-- Proces each child item
    $rFiles = @()
    $rFolders = @()
    foreach ($oItem in $rItems)
    {
        if ($oItem.contentType.name -eq "Folder") {$rFolders += $oItem}
        elseif ($oItem.contentType.name -eq "Document") {$rFiles += $oItem}
    }

    #-- Display it
    Write-Host "Files: $($rFiles.Count)"
    Write-Host "Folders: $($rFolders.Count)"
    #foreach ($oFile in $rFiles) {Write-Host $oFile.webUrl.substring($sSiteUrl.Length) $oFile.contentType.id -ForegroundColor Green}
    #foreach ($oFolder in $rFolders[0]) {$oFolder | fl} #Write-Host "Folder: $($oFolder.name)"}

    #-- Get the permissions for folders
    foreach ($oFolder in $rFolders[0..1])
    {
        #-- Troubleshooting
        #Write-Host "folder id:" $oFolder.contentType.id
        #$oFolder | fl

        #-- Add to our final list of folders
        $oNewItem = [PSCustomObject]@{
            id = $oFolder.eTag
            webUrl = $oFolder.webUrl
        }
        $script:rgFolders += $oNewItem

        #-- Get the permissions
        Write-Host "Folder: $($oFolder.webUrl)" -ForegroundColor Yellow
        $rPerms = udf_GetItemPermissions -sSiteId $sSiteId -sFileId $oFolder.contentType.id
        foreach ($oPerm in $rPerms)
        {
            #-- Skip certain entries
            $sRoles = $oPerm.roles -join ','
            if ($sRoles.Length -eq 0) {continue}

            #-- Add to our final list of permissions
            $oNewPerms = [PSCustomObject]@{
                id = $oFolder.eTag
                roles = $sRoles
                displayName = $oPerm.displayName
                email = $oPerm.email
            }
            $script:rgPerms += $oNewPerms

            #-- Display it
            if ($sRoles -match "owner") {continue}
            Write-Host "  $($sRoles) $($oPerm.displayName) $($oPerm.email)" -ForegroundColor Cyan
        }
    }

    #-- Get the permissions for files
}

#--------------------------------------------------------------------------------------------------
#-- Get all the document libraries for a site
#--------------------------------------------------------------------------------------------------
function udf_GetDocumentLibraries
{
    param (
        [string]$sSiteId
    )
    #-- Get the lists for this site
    #$sUri = "https://graph.microsoft.com/v1.0/sites/$sSiteId/lists"
    #$oResult = Invoke-RestMethod -Headers @{Authorization = "Bearer $token"} -Uri $sUri -Method Get
    #$rLists = $oResult.value

    #-- Get the drives for this site
    $sUri = "https://graph.microsoft.com/v1.0/sites/$sSiteId/drives"
    $oResult = Invoke-RestMethod -Headers @{Authorization = "Bearer $token"} -Uri $sUri -Method Get
    $oDrives = $oResult.value

    #-- Iterate through the drives
    $rReturn = @()
    foreach ($oDrive in $oDrives)
    {
        $oNew = [PSCustomObject]@{
            id = $oDrive.id
            name = $oDrive.name
        }
        $rReturn += $oNew
    }
    return $rReturn
}

#--------------------------------------------------------------------------------------------------
#-- Main script
#--------------------------------------------------------------------------------------------------

$rDocLibs = @()
$rDocLibs = udf_GetDocumentLibraries -sSiteId $sSiteId
if ($rDocLibs.Count -gt 0)
{
    foreach ($oDocLib in $rDocLibs)
    {
        if ($oDocLib.name -eq "Documents")
        {
            Write-Host "Document Library: $($oDocLib.name) $($oDocLib.id)" -ForegroundColor Green
            #udf_GetChildItems -sSiteId $sSiteId -sFolderId $oDocLib.id
            udf_GetDriveItems -sSiteId $sSiteId -sFolderId $oDocLib.id
        }
    }
}
#$script:rgFolders
#$script:rgPerms
