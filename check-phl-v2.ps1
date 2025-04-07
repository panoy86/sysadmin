#-- Gets all the holds in a given OneDrive account


function udf_GetHolds
{
    param
    (
        [Parameter(Mandatory=$true)]$sUrl
    )

    #-- Get the holds
    Write-Host $sUrl
    $rLists = Get-PnPList
    $nItemCount = 0
    $nFolderCount = 0
    foreach ($oList in $rLists)
    {
        if ($oList.Title -match "Preservation" -or $oList.Title -like "(Collection*")
        {
            Write-Host "  " $oList.Title -NoNewline -ForegroundColor Green
            Try {$oMeasure = Measure-PnPList $oList.Title -ea SilentlyContinue} Catch {$oMeasure = $null}
            if ($null -ne $oMeasure)
            {
                Write-Host "" $oMeasure.ItemCount $oMeasure.FolderCount $oMeasure.TotalFileSize
                $nItemCount += $oMeasure.ItemCount
                $nFolderCount += $oMeasure.FolderCount
            }
            else
            {
                if ($Error[0] -match "The attempted operation is prohibited because it exceeds the list view threshold") {$nItemCount += 5000}
                Write-Host ""
            }
            #break
        }
    }
    Write-Host "Total:" $nItemCount $nFolderCount
}


#-- Main
$r = Get-Content .\list-onedrive-urls.txt
foreach ($sUrl in $r)
{
    #-- Get the holds
    Write-Progress -Activity $sUrl
    $null = Set-SPOUser -Site $sUrl -LoginName ftan-a@something.com -IsSiteCollectionAdmin:$TRUE
    Connect-PnPOnline -Url $sUrl -SPOManagementShell
    udf_GetHolds -sUrl $sUrl
    Disconnect-PnPOnline
}
Write-Progress -Activity "End" -Completed
