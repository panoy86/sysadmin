#-- Graps app-pemissions: Files.Read.All, Sites.Read.All, for rw/: Sites.ReadWrite.All, Files.ReadWrite.All
#-- Find-MgGraphCommand -Command get-mguserdrive 
#-- (Get-MgContext).Scopes -> to check if user/app has the right permissions

$script:listOverSharedItems = @()
$script:hShareIds = @{}
$script:nPerUser = 0
$script:nPerUserTotal = 0

#------------------------------------------------------------------------------
#-- Function to get the permissions of a OneDrive item
#-- > adds overshared items to script-wide list
#------------------------------------------------------------------------------
function CheckOverSharingPermissions
{
    param (
        [Parameter(Mandatory = $true)]
        [string]$DriveId,
        [Parameter(Mandatory = $true)]
        [string]$DriveItemId
    )
    
    #-- Get permissions for the item
    $itemPermissions = Get-MgDriveItemPermission -DriveId $DriveId -DriveItemId $DriveItemId -All

    #-- Check for over-sharing
    $listOverShares = @()
    foreach ($permission in $itemPermissions)
    {
        #-- Everyone except external users
        if ($permission.GrantedTo.User -eq "Everyone except external users" -or $permission.GrantedToV2.SiteUser.DisplayName -eq "Everyone except external users")
        {
            #-- I *think* a duplicate ShareId indicates an inherited permission
            #-- so let's skip over items that are simply inheriting from the parent
            $script:nPerUserTotal++
            $sKey = $permission.ShareId
            if ($script:hShareIds.ContainsKey($sKey)) {$script:hShareIds[$sKey]++}
            else
            {
                $script:hShareIds.Add($sKey, 1)
                $listOverShares += "Everyone except external users"
            }
        }
    }

    #-- Check if we need to add this
    if ($listOverShares.Count -gt 0)
    {
        $overSharedItem = [PSCustomObject]@{
            User       = $sUpn
            ItemName   = $item.Name
            ItemId     = $item.Id
            WebUrl     = $item.WebUrl
            OverShared = $listOverShares -join ", "
        }
        $script:listOverSharedItems += $overSharedItem
        $script:nPerUser++
    }
}

#------------------------------------------------------------------------------
#-- Function to recursively get OneDrive items
#------------------------------------------------------------------------------
function RecursivelyGetOneDriveItems
{
    param (
        [Parameter(Mandatory = $true)]
        [string]$DriveId,
        [Parameter(Mandatory = $true)]
        [string]$DriveItemId
    )
    #-- Get child items of the current drive item
    $childItems = Get-MgDriveItemChild -DriveId $DriveId -DriveItemId $DriveItemId -All
    
    #-- Process each child item
    foreach ($item in $childItems)
    {
        Write-Progress -Activity "Processing item" -Status $item.Name
        CheckOverSharingPermissions -DriveId $DriveId -DriveItemId $item.Id

        #-- Recursively call this function for folders
        if ($null -ne $item.Folder.ChildCount)
        {
            RecursivelyGetOneDriveItems -DriveId $DriveId -DriveItemId $item.Id 
        }
    }
    Write-Progress -Activity "Processing item" -Completed
}

#------------------------------------------------------------------------------
#-- Main
#------------------------------------------------------------------------------
$PSStyle.Progress.View = "Minimal"  #-- Other value: "Classic", only works in PowerShell 7.2+
$ProgressPreference = "Continue"

Write-Host "This script will search for over-shared items in OneDrive accounts"
Write-Host "(for now just: Everyone except external users)"

#-- Get userPrincipalName (UPN) from the user
if ($args.Count -gt 0) {$rUpns = $args}
else
{
    Write-Host "List down the UPNs, separated by commas:"
    $sInput = Read-Host
    $rUpns = $sInput.Split(",")
}

<#-- Test
$rUpns = @()
$rUpns += "user1@somedomain.onmicrosoft.com"
$rUpns += "user2@somedomain.com"
$rUpns += "user3@somedomain.com"
#>

#-- Loops thru our list of users
foreach ($sUpn in $rUpns)
{
    Write-Host "Processing user" $sUpn -NoNewline
    $script:nPerUser = 0
    $script:nPerUserTotal = 0
    $script:hShareIds = @{}

    #-- Verify OneDrive
    $userOneDrive = Get-MgUserDefaultDrive -UserId $sUpn -ea SilentlyContinue    
    if ($null -eq $userOneDrive)
    {
        Write-Host " No OneDrive content found for:" $sUpn -ForegroundColor Yellow
        continue
    }
    
    #-- Get the root-id of the user's OneDrive
    $rootDriveItem = Get-MgDriveItem -DriveId $userOneDrive.Id -DriveItemId "root"
    RecursivelyGetOneDriveItems -DriveId $userOneDrive.Id -DriveItemId $rootDriveItem.Id

    #-- Show stats per user
    if ($script:nPerUserTotal -gt 0) {Write-Host " -> found" $script:nPerUser.Count "with a cumulative folders/files of" $script:nPerUserTotal -ForegroundColor Cyan}
    else {Write-Host "."}
}

Write-Host "Total found:" $script:listOverSharedItems.Count
if ($script:listOverSharedItems.Count -gt 0)
{
    Write-Host "Saving results to overshared-items.csv" -ForegroundColor Green
    $script:listOverSharedItems | Export-Csv overshared-items.csv -NoTypeInformation
}
