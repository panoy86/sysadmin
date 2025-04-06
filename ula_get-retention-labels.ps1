$global:oTmp = $null


function Get-PnPOneDriveRetentionLabelReport()
{
    Param
    (
        [Parameter(Mandatory=$true)]$sUrl
    )
    
    $results = @()
    $SystemLibraries = @("Form Templates", "Pages", "Preservation Hold Library", "Site Assets", "Site Pages", "Images", "Site Collection Documents", "Site Collection Images", "Style Library", "Teams Wiki Data")
    [array]$lists = Get-PnPList | Where { $_.BaseType -eq "DocumentLibrary" -and $_.Hidden -eq $False -and $_.Title -notin $SystemLibraries }
                
    foreach ($list in $lists)
    {
        $ListItems = Get-PnPListItem -List $List.Title -PageSize 5000 
        #write-host "Count of the items: $($ListItems.count)"
        ForEach($Item in $ListItems)
        {   
            #Collect Data                                                                   
            $obj = [pscustomobject][ordered]@{
                Url         = $sUrl
                ListTitle   = $list.Title
                ItemID      = $Item.id
                ItemName    = $Item.FieldValues.FileLeafRef
                LabelTag    = $Item.FieldValues._ComplianceTag
                ItemPath    = $Item.fieldvalues.FileRef  
                Type        = $Item.FieldValues.FSObjType
                ParentUid   = $Item.FieldValues.ParentUniqueId.ToString()
                Uid         = $Item.FieldValues.UniqueId.ToString()
                Size_MB     = [Math]::Round($Item.FieldValues.SMTotalSize.LookupId/1MB,2)
            }
            $results += $obj
            $global:oTmp = $Item
        }
    }
    $rReturn = $results | where {$_.LabelTag -ne ""}
    Write-Host ", found:" $rReturn.Count "of" $results.Count
    return $rReturn

}

#-- Main script
$sAdmin = "ftan-a@nnn.com"
$rWork = Import-Csv .\zz.csv
$nCountToWorkOn = 4000

$nCtr = 0
foreach ($oItem in $rWork)
{
    #-- Skip certain items
    #if ($oItem.RetentionPolicy.Trim().Length -gt 0) {continue}
    #if ($oItem."Retention Labels".Trim().Length -gt 0) {continue}
    if ($oItem.IsDeleted -eq "true") {continue}
    
    #-- Check if we need to stop
    $nCtr++
    if ($nCtr -gt $nCountToWorkOn) {break}
    
    #-- Get the OneDrive URL
    Write-Host "  " $oItem.Url -NoNewline
    $sUrl = $oItem.Url
    $t = $null
    $t = Get-SPOSite -Identity $sUrl
    if ($t -eq $null)
    {
        $oItem.IsDeleted = "true"
        $oItem."Retention Labels" = 0
        continue
    }
    
    #-- Get the retention labels
    $null = Set-SPOUser -Site $sUrl -LoginName $sAdmin -IsSiteCollectionAdmin $true
    #Connect-PnPOnline -Url $sUrl -CurrentCredentials
    #Connect-PnPOnline -Url $sUrl -UseWebLogin
    Connect-PnPOnline -Url $sUrl -SPOManagementShell
    $rList = @()
    $rList = Get-PnPOneDriveRetentionLabelReport -sUrl $sUrl
    
    #-- Record the results
    if ($rList.Count -gt 0)
    {
        $sUser = ($oItem."Owner email" -split "@")[0]
        $rList | Export-Csv (".\results\results." + $sUser + ".csv") -NoTypeInformation
        $oItem."Retention Labels" = $rList.Count
    }
    else {$oItem."Retention Labels" = 0}
    
    #-- Disconnect
    Disconnect-PnPOnline
    $null = Set-SPOUser -Site $sUrl -LoginName $sAdmin -IsSiteCollectionAdmin $false
    #Sleep 3
}
#-- Save the work
$rWork | Export-Csv .\zz.csv -NoTypeInformation

#-- To Check: $rWork | where {$_."Retention Labels".Length -gt 0}
