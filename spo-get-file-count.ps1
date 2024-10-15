<#-- Notes
https://martinday.co/determining-the-last-user-activity-of-a-sharepoint-or-onedrive-site/
Requires an existing PowerShell session for SPO and PnP
-> Connect-SPOService -Url https://paypal-admin.sharepoint.com
#>

$sFilename = ".\hpy-site-list.csv"

<#-- Manually generate our CSV file from the site list
$r = Get-Content .\site-urls-list.txt | where {$_.Trim() -notlike "#*"}
$rSites = @()
foreach ($t in $r)
{
    $oNew = New-Object PSObject -Property @{
        Url                      = $t
        FolderCount              = 0
        FileCount                = 0
        LastItemUserModifiedDate = ''
    }
    $rSites += $oNew 
}
$rSites | select Url,FileCount,FolderCount,LastItemUserModifiedDate | Export-Csv $sFilename
#>

$global:FolderCount = 0
$global:FileCount = 0
$global:LastItemUserModifiedDate = $null


#-- function to start on a top-level folder, and recursively go down and count the numbers of folders/files
function udf_RecursiveCount
{
    param ([string]$sRelativeFolder)
    #$sRelativeFolder
    $rItems = Get-PnPFolderItem -FolderSiteRelativeUrl $sRelativeFolder
    foreach ($oItem in $rItems)
    {
        if ($oItem.TypedObject.GetType().Name -eq "Folder")
        {
            $global:FolderCount++
            udf_RecursiveCount($sRelativeFolder + "/" + $oItem.Name)
        }
        else
        {
            if ($oItem.TypedObject.GetType().Name -eq "File") {$global:FileCount++}
        }
    }
}

#-- function to get the top-level folders and start a recursive iteration thru sub-folders
function udf_GetTopLevelFolders
{
    #-- Skip over system defined lists
    $rSkip = @()
    $rSkip += "appdata"
    $rSkip += "appfiles"
    $rSkip += "Composed Looks"
    $rSkip += "Converted Forms"
    $rSkip += "Form Templates"
    $rSkip += "List Template Gallery"
    $rSkip += "Maintenance Log Library"
    $rSkip += "Master Page Gallery"
    $rSkip += "Site Assets"
    $rSkip += "Site Pages"
    $rSkip += "Solution Gallery"
    $rSkip += "Style Library"
    $rSkip += "TaxonomyHiddenList"
    $rSkip += "Theme Gallery"
    $rSkip += "User Information List"
    $rSkip += "Web Part Gallery"

    #-- Get our top-level lists and iterate
    $rLists = Get-PnPList
    if ($rLists.Count -gt 0)
    {
        #-- Just process the lists meant for user documents
        foreach ($oList in $rLists)
        {
            if ($oList.Title -notin $rSkip)
            {
                udf_RecursiveCount($oList.RootFolder.ServerRelativeUrl.Substring($oList.ParentWebUrl.Length))
                #-- Initiate our last-user-modified date
                if ($null -eq $global:LastItemUserModifiedDate)
                {$global:LastItemUserModifiedDate = $oList.LastItemUserModifiedDate}
                #-- Update our last-user-modified date
                if ([datetime]$global:LastItemUserModifiedDate -gt [datetime]$oList.LastItemUserModifiedDate)
                {$global:LastItemUserModifiedDate = $oList.LastItemUserModifiedDate}
            }
        }
    }
    else {Write-Host "No top level folders found for:" $global:sUrl -ForegroundColor Red}
}

#-- Main
#-- Get the list of SPO sites to count the number of folders/files
$rSites = Import-Csv $sFilename

#-- Iterate thru each one
$nCtr = 0
foreach ($oSite in $rSites)
{
    #-- Show progress
    $nCtr++
    Write-Progress -Activity $oSite.Url -Status '.' -PercentComplete (($nCtr * 100)/$rSites.Count)
    Write-Host $oSite.Url -NoNewline

    #-- Add myself as an admin
    $sUrl = $oSite.Url
    $sAdmin = "ftan-a@paypal.com"
    $null = Set-SPOUser -Site $sUrl -LoginName $sAdmin -IsSiteCollectionAdmin:$true
    Start-Sleep 5
    
    #-- Connect, reset our counters, and get the folder/file count
    Connect-PnPOnline -Url $sUrl -UseWebLogin
    $global:FolderCount = 0
    $global:FileCount = 0
    $global:LastItemUserModifiedDate = $null
    udf_GetTopLevelFolders

    #-- Save the results and remove myself as an admin
    $null = Set-SPOUser -Site $sUrl -LoginName $sAdmin -IsSiteCollectionAdmin:$false
    Write-Host (" " + $global:FolderCount + "/" + $global:FileCount + " ") -NoNewline
    $oSite.FolderCount = $global:FolderCount
    $oSite.FileCount = $global:FileCount
    if ($null -ne $global:LastItemUserModifiedDate)
    {
        $oSite.LastItemUserModifiedDate = $global:LastItemUserModifiedDate
        Write-Host $global:LastItemUserModifiedDate
    }
    else {Write-Host "No modified date found"}
}
$rSites | Export-Csv $sFilename -NoTypeInformation