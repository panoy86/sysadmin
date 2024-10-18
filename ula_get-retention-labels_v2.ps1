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
$sAdmin = "ftan-a@paypal.com"
$rWork = Get-Content .\list-onedrive-urls.txt
$rFinal = @()
$nCountToWorkOn = 4000

$nCtr = 0
foreach ($sUrl in $rWork)
{
    #-- Check if we need to stop
    $nCtr++
    if ($nCtr -gt $nCountToWorkOn) {break}
    
    $oNew = New-Object PSObject -Property @{url=$sUrl; Count=0}
    
    #-- Get the OneDrive URL
    Write-Host "  " $sUrl -NoNewline
    $t = $null
    $t = Get-SPOSite -Identity $sUrl
    if ($t -eq $null)
    {
        $rFinal += $oNew
        continue
    }
    
    #-- Get the retention labels
    $null = Set-SPOUser -Site $sUrl -LoginName $sAdmin -IsSiteCollectionAdmin $true
    Connect-PnPOnline -Url $sUrl -SPOManagementShell
    $rList = @()
    $rList = Get-PnPOneDriveRetentionLabelReport -sUrl $sUrl
    
    #-- Record the results
    if ($rList -ne $null -and $rList.Count -gt 0) {$oNew.Count = $rList.Count}
    $rFinal += $oNew
    
    #-- Disconnect
    Disconnect-PnPOnline
    $null = Set-SPOUser -Site $sUrl -LoginName $sAdmin -IsSiteCollectionAdmin $false

}
$rFinal