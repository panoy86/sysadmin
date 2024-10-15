#-- Get the shared links for a given site
$sUrl = "https://paypal.sharepoint.com/teams/TurnerDirects"
Set-SPOUser -Site $sUrl -LoginName ftan-a@paypal.com -IsSiteCollectionAdmin:$true

#-- Get the lists and choose one (manually)
Connect-PnPOnline -Url $sUrl -UseWebLogin
$oCtx = Get-PnPContext
$rLists = Get-PnPList
$rLists | ft Title,BaseType
$sListName = "Documents"

#-- Get the items from the chosen list
$rListItems = Get-PnPListItem -List $sListName -PageSize 2000
$rListItems.Count

#-- Iterate thru our items
$nCtr = 0
foreach ($oListItem in $rListItems)
{
    #-- Show Progress
    $nCtr++
    Write-Progress -Activity "Getting shared links..." -Status '.' -PercentComplete ($nCtr/$rListItems.Count * 100)
    
    #-- Check the item
    $bHasUniquePermissions = Get-PnPProperty -ClientObject $oListItem -Property "HasUniqueRoleAssignments"
    if ($bHasUniquePermissions)
    {       
        #-- Get shared links
        $oSharingInfo = [Microsoft.SharePoint.Client.ObjectSharingInformation]::GetObjectSharingInformation($oCtx, $oListItem, $false, $false, $false, $true, $true, $true, $true)
        $ctx.Load($oSharingInfo)
        $ctx.ExecuteQuery()
         
        ForEach($ShareLink in $oSharingInfo.SharingLinks)
        {
            If($ShareLink.Url)
            {           
                If($ShareLink.IsEditLink)
                {
                    $AccessType="Edit"
                }
                ElseIf($shareLink.IsReviewLink)
                {
                    $AccessType="Review"
                }
                Else
                {
                    $AccessType="ViewOnly"
                }
                 
                #Collect the data
                $Results += New-Object PSObject -property $([ordered]@{
                Name  = $oListItem.FieldValues["FileLeafRef"]           
                RelativeURL = $oListItem.FieldValues["FileRef"]
                FileType = $oListItem.FieldValues["File_x0020_Type"]
                ShareLink  = $ShareLink.Url
                ShareLinkAccess  =  $AccessType
                ShareLinkType  = $ShareLink.LinkKind
                AllowsAnonymousAccess  = $ShareLink.AllowsAnonymousAccess
                IsActive  = $ShareLink.IsActive
                Expiration = $ShareLink.Expiration
                })
            }
        }
    }
    $global:counter++
}


$Ctx = Get-PnPContext
$Results = @()
$global:counter = 0
  
#Get all list items in batches
$rListItems = Get-PnPListItem -List $sListName -PageSize 2000
$ItemCount = $rListItems.Count
   
#Iterate through each list item
ForEach($Item in $rListItems)
{
    Write-Progress -PercentComplete ($global:Counter / ($ItemCount) * 100) -Activity "Getting Shared Links from '$($Item.FieldValues["FileRef"])'" -Status "Processing Items $global:Counter to $($ItemCount)";
 
    #Check if the Item has unique permissions
    $HasUniquePermissions = Get-PnPProperty -ClientObject $Item -Property "HasUniqueRoleAssignments"
    If($HasUniquePermissions)
    {       
        #Get Shared Links
        $oSharingInfo = [Microsoft.SharePoint.Client.ObjectoSharingInformation]::GetObjectoSharingInformation($Ctx, $Item, $false, $false, $false, $true, $true, $true, $true)
        $ctx.Load($oSharingInfo)
        $ctx.ExecuteQuery()
         
        ForEach($ShareLink in $oSharingInfo.SharingLinks)
        {
            If($ShareLink.Url)
            {           
                If($ShareLink.IsEditLink)
                {
                    $AccessType="Edit"
                }
                ElseIf($shareLink.IsReviewLink)
                {
                    $AccessType="Review"
                }
                Else
                {
                    $AccessType="ViewOnly"
                }
                 
                #Collect the data
                $Results += New-Object PSObject -property $([ordered]@{
                Name  = $Item.FieldValues["FileLeafRef"]           
                RelativeURL = $Item.FieldValues["FileRef"]
                FileType = $Item.FieldValues["File_x0020_Type"]
                ShareLink  = $ShareLink.Url
                ShareLinkAccess  =  $AccessType
                ShareLinkType  = $ShareLink.LinkKind
                AllowsAnonymousAccess  = $ShareLink.AllowsAnonymousAccess
                IsActive  = $ShareLink.IsActive
                Expiration = $ShareLink.Expiration
                })
            }
        }
    }
    $global:counter++
}
$Results | Export-CSV $ReportOutput -NoTypeInformation
Write-host -f Green "Sharing Links Report Generated Successfully!"



