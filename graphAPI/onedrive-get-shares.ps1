<#
. c:\scripts\GraphApi\start-graph-api-ftan_script.ps1
$sUserUpn = ""
Set-Variable -Name oAuthHeader -Scope Global -Value @{'Authorization'="Bearer $token";'Content-Type'='application\json'}

#-- Get User ID
$sUri = "https://graph.microsoft.com/v1.0/users/?`$select=displayName,mail,userPrincipalName,id,userType&`$top=999&`$filter=userPrincipalName eq '$sUserUpn'"
$oResult = Invoke-WebRequest -Headers $oAuthHeader -Uri $sUri -Verbose:$VerbosePreference -ErrorAction Stop
$oUser = $oResult.Content | ConvertFrom-Json
$sUserId = $oUser.value.id
#>

#---------------------------------------------------------------------------------------------------
#-- Get the sharing permissions for a particular item
#---------------------------------------------------------------------------------------------------
function udf_GetPermissions
{

    Param(
        [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$sUserId,
        [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$sItemId
    )

    #-- Fetch permissions for the given item
    $sUri = "https://graph.microsoft.com/v1.0/users/$($sUserId)/drive/items/$($sItemId)/permissions"
    $oResult = Invoke-WebRequest -Headers $oAuthHeader -Uri $sUri -Verbose:$VerbosePreference -ErrorAction Stop
    $rPermissions = $oResult.Content | ConvertFrom-Json

    #-- Setup the return list
    $rReturnList = @()

    #-- There are 4 types of permissions: sharing link, invitations, direct, and inherited
    foreach ($oPermission in $rPermissions.value)
    {
        $oNew = New-Object PSObject
        $oNew | Add-Member -MemberType NoteProperty -Name Type -Value '' -Force
        $oNew | Add-Member -MemberType NoteProperty -Name Role -Value '' -Force
        $oNew | Add-Member -MemberType NoteProperty -Name Email -Value '' -Force

        #-- Sharing link
        if ($oPermission.link)
        {
            #Write-Host "   Link: $($oPermission.roles) $($oPermission.link.webUrl)"
            #$oPermission.link | Format-List
            #$oPermission | Format-List; exit

            $oNew.Type = "Link"
            $oNew.Role = $oPermission.roles
            $oNew.Email = $oPermission.link.webUrl
        }
        #-- Invitations
        elseif ($oPermission.invitation)
        {
            #Write-Host "   Invitation: $($oPermission.roles) $($oPermission.invitation.email)"
            #$oPermission | Format-List; exit

            $oNew.Type = "Invitation"
            $oNew.Role = $oPermission.roles
            $oNew.Email = $oPermission.invitation.email
        }
        #-- Direct
        elseif ($oPermission.roles)
        {
            $oNew.Type = "Direct"
            $oNew.Role = $oPermission.roles
            if ($oPermission.grantedTo.user.email) {$oNew.Email = $oPermission.grantedTo.user.email}
            else {$oNew.Email = $oPermission.grantedTo.user.displayName}

            #Write-Host "   Direct: $($oPermission.roles) $($oPermission.invitation.email)"
            #$oPermission.grantedToV2.siteUser | Format-List
            #$oPermission | Format-List; exit
        }
        #-- Inherited
        elseif ($oPermission.inheritedFrom)
        {
            #Write-Host "   Inherited from: $($entry.inheritedFrom.path)"

            $oNew.Type = "Inherited"
            $oNew.Role = $oPermission.roles
            $oNew.Email = $oPermission.inheritedFrom.path
        }
        #-- Some other permissions that we don't know of
        else
        {
            #Write-Verbose "   Permission $oPermission not covered by the script!"

            $oNew.Type = "Unknown"
            $oNew.Role = "Unknown"
            $oNew.Email = "Unknown"
        }
        #-- Add this permission entry to the return list
        $rReturnList += $oNew

        #-- Debug
        #if ($sItemId -eq "01PZQUGYGVN3GTVJME2NEZVD4ZP2GHB6EB") {$oPermission.grantedTo | Format-List; $oPermission.grantedToV2 | Format-List}
    }
    return $rReturnList
}


#---------------------------------------------------------------------------------------------------
#-- Recursive function to iterate through all child items from a top-level folder
#---------------------------------------------------------------------------------------------------
function udf_GetChildItems
{

    Param
    (
        [Parameter(Mandatory=$true)]$sUserId,
        [Parameter(Mandatory=$true)]$sFolderId
    )

    #-- Get the child items
    $sUri = "https://graph.microsoft.com/v1.0/users/$($sUserId)/drive/items/$($sFolderId)/children"
    do
    {
        $oResult = Invoke-WebRequest -Headers $oAuthHeader -Uri $sUri -Verbose:$VerbosePreference -ErrorAction Stop
        $sUri = $oResult.'@odata.nextLink'
        #-- Add delay to avoid SPO throttling
        Start-Sleep -Milliseconds 500
        $rItems += $oResult.Content | ConvertFrom-Json
    } while ($sUri)
    
    #-- Proces each child item, there are 3 types: folder, file, notebook
    $rFiles = @()
    $rFolders = @()
    $rNotebooks = @()
    foreach ($oItem in $rItems.value)
    {
        if ($null -ne $oItem.folder)
        {
            if ($null -eq $oItem.package) {$rFolders += $oItem}
            else {$rNotebooks += $oItem}
        }
        else {$rFiles += $oItem}
    }

    <#-- Display it
    foreach ($oFile in $rFiles) {Write-Host "File: $($oFile.name) $($oFile.file)"}
    foreach ($oNotebook in $rNotebooks) {Write-Host "Notebook: $($oNotebook.name) $($oNotebook.package)"}
    foreach ($oFolder in $rFolders) {Write-Host "Folder: $($oFolder.name) $($oFolder.folder)"}
    #>

    #-- Get the permissions
    #
    foreach ($oFile in $rFiles)
    {
        if ($oFile.shared)
        {
            Write-Host "$(' ' * $nTab)File: $($oFile.name) $($oFile.id)" -ForegroundColor Green
            $rTmp = udf_GetPermissions -sUserId $sUserId -sItemId $oFile.id
            foreach ($oTmp in $rTmp)
            {
                Write-Host "$(' ' * $nTab)   $($oTmp.Type) $($oTmp.Role) $($oTmp.Email)"
            }
        }
    }
    #>
    foreach ($oFolder in $rFolders)
    {
        #if ($false)
        if ($oFolder.shared)
        #if ($oFolder.name -eq "share-internal")
        {
            Write-Host "$(' ' * $nTab)Folder: $($oFolder.name) $($oFolder.folder) $($oFolder.id)" -ForegroundColor Green
            $rTmp = udf_GetPermissions -sUserId $sUserId -sItemId $oFolder.id
            foreach ($oTmp in $rTmp)
            {
                Write-Host "$(' ' * $nTab)   $($oTmp.Type) $($oTmp.Role) $($oTmp.Email)"
            }
        }
        #$nTab += 3
        Write-Host $oFolder.name -ForegroundColor DarkYellow
        udf_GetChildItems -sUserId $sUserId -sFolderId $oFolder.id
        #$nTab -= 3
    }
}

#---------------------------------------------------------------------------------------------------
#-- Main script
#---------------------------------------------------------------------------------------------------

#-- Get the top-level folder
Set-Variable -Name oAuthHeader -Scope Global -Value @{'Authorization'="Bearer $token";'Content-Type'='application\json'}
$sUri = "https://graph.microsoft.com/v1.0/users/$($sUserId)/drive/root"
$oResult = Invoke-WebRequest -Headers $oAuthHeader -Uri $sUri -Verbose:$VerbosePreference -ErrorAction Stop
$oRoot = $oResult.Content | ConvertFrom-Json
Write-Host "Root: $($oRoot.webUrl)" -ForegroundColor Yellow

#-- Get the permissions for the root folder items and children
Set-Variable -Name nTab -Scope Global -Value 0
udf_GetChildItems -sUserId $sUserId -sFolderId $oRoot.id
