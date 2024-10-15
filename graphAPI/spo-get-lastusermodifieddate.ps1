#$rSites | foreach {$_ | Add-Member -Type NoteProperty -Name "LastItemUserModifiedDate" -Value '' -Force}
#$rSites | foreach {$_ | Add-Member -Type NoteProperty -Name "LastItemUserModifiedList" -Value '' -Force}

$global:LastItemUserModifiedDate = $null
$global:LastItemUserModifiedList = $null


#-- function to get the top-level folders and start a recursive iteration thru sub-folders
function udf_GetTopLevelFolders
{
    #-- Get our top-level lists and iterate
    $rLists = Get-PnPList
    if ($rLists.Count -gt 0)
    {
        #-- Just process the lists meant for user documents
        foreach ($oList in $rLists)
        {
            if ($oList.Title -notin $rSkip)
            {
                #-- Initiate our last-user-modified date/list
                if ($null -eq $global:LastItemUserModifiedDate)
                {$global:LastItemUserModifiedDate = $oList.LastItemUserModifiedDate}
                if ($null -eq $global:LastItemUserModifiedList)
                {$global:LastItemUserModifiedList = $oList.Title}
                
                #-- Update our last-user-modified date
                if ([datetime]$global:LastItemUserModifiedDate -lt [datetime]$oList.LastItemUserModifiedDate)
                {
                    $global:LastItemUserModifiedDate = $oList.LastItemUserModifiedDate
                    $global:LastItemUserModifiedList = $oList.Title
                }
            }
        }
    }
    else {Write-Host "No top level folders found for:" $global:sUrl -ForegroundColor Red}
}

#-- Main
$rSites = Import-Csv .\ownerless_inactive_modified.csv

#-- Iterate thru each one
$nCtr = 0
foreach ($oSite in $rSites[0..9])
{
    #-- Show progress
    $nCtr++
    #if ($nCtr -lt 6000) {continue}
    if ($oSite.SharePointURL.Trim().Length -eq 0) {continue}
    #if ($oSite.LastItemUserModifiedList.Trim().Length -eq 0)
    if ($true)
    {
        Write-Progress -Activity ($nCtr.ToString() + " " + $oSite.SharePointURL) -Status '.' -PercentComplete (($nCtr * 100)/$rSites.Count)
        Write-Host ($oSite.SharePointURL + " ") -NoNewline

        #-- Add myself as an admin
        $sUrl = $oSite.SharePointURL
        $sAdmin = "ftan-a@paypal.com"
        $null = Set-SPOUser -Site $sUrl -LoginName $sAdmin -IsSiteCollectionAdmin:$true
        Start-Sleep 5
        
        #-- Connect, reset our counters, and get the folder/file count
        Connect-PnPOnline -Url $sUrl -UseWebLogin
        $global:LastItemUserModifiedDate = $null
        $global:LastItemUserModifiedList = $null

        udf_GetTopLevelFolders

        #-- Save the results and remove myself as an admin
        $null = Set-SPOUser -Site $sUrl -LoginName $sAdmin -IsSiteCollectionAdmin:$false
        if ($null -ne $global:LastItemUserModifiedDate)
        {
            $oSite.LastItemUserModifiedDate = $global:LastItemUserModifiedDate
            $oSite.LastItemUserModifiedList = $global:LastItemUserModifiedList
            Write-Host $global:LastItemUserModifiedDate "-" $global:LastItemUserModifiedList
        }
        else {Write-Host "No modified date found"}
    }
}
#$rSites | Export-Csv .\ownerless_inactive_modified.csv -NoTypeInformation
$nCtr
