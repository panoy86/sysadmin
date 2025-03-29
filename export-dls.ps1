#-- This works onprem only, Exchange Online will yield different results
<#- This script will yield 6 files, namely:
    exp-dls-accept.csv
    exp-dls-managedby.csv
    exp-dls-members.csv
    exp-dls-moderation.csv
    exp-dls-reject.csv
    exp-dls.csv
-#>

#-- Get all DLs, assuming it's < 20K
Write-Host "Getting all Distribution Groups, please wait..." -NoNewline
$rd = Get-DistributionGroup -ResultSize 20000; $rd.Count

#-- Normalize certain properties
$rd | foreach {$_ | Add-Member -MemberType NoteProperty -Name "EmailAddressesString" -Value '' -Force}
$rd | foreach {$_.EmailAddressesString = $_.EmailAddresses -join ";"}

$rd | select Alias,CustomAttribute1,CustomAttribute2,CustomAttribute3,CustomAttribute4,CustomAttribute5,CustomAttribute6,CustomAttribute7,CustomAttribute8,CustomAttribute9,CustomAttribute10,CustomAttribute11,CustomAttribute12,CustomAttribute13,CustomAttribute14,CustomAttribute15,DisplayName,EmailAddressesString,EmailAddressPolicyEnabled,GroupType,Guid,HiddenFromAddressListsEnabled,Identity,LegacyExchangeDN,MemberJoinRestriction,MemberDepartRestriction,ModerationEnabled,MailTip,Name,PrimarySmtpAddress,ReportToManagerEnabled,ReportToOriginatorEnabledRecipientType,RecipientTypeDetails,RequireSenderAuthenticationEnabled,SamAccountName,SimpleDisplayName,SendModerationNotifications,WindowsEmailAddress | Export-Csv exp-dls.csv -NoTypeInformation

#-- Get the accept list
$nCtr = 0
$rFinal = @()
foreach ($d in $rd)
{
    $nCtr++; Write-Progress -Activity "Accept List" -Status $nCtr -PercentComplete ($nCtr / $rd.Count * 100)
    if ($d.AcceptMessagesOnlyFromSendersOrMembers.Count -gt 0)
    {
        $rTmp = @()
        foreach ($oAccept in $d.AcceptMessagesOnlyFromSendersOrMembers)
        {
            $oTmp = $null
            $oTmp = Get-Recipient -Identity $oAccept.DistinguishedName -ea SilentlyContinue
            if ($null -ne $oTmp) {$rTmp += $oTmp}
        }
        if ($rTmp.Count -gt 0)
        {
            foreach ($oTmp in $rTmp)
            {
                $rFinal += [PSCustomObject]@{
                    DLGuid = $d.Guid
                    Guid = $oTmp.Guid
                    RecipientTypeDetails = $oTmp.RecipientTypeDetails
                }
            }
        }
    }
}
if ($rFinal.Count -gt 0) {$rFinal | Export-Csv exp-dls-accept.csv -NoTypeInformation}

#-- Get the reject list
$nCtr = 0
$rFinal = @()
foreach ($d in $rd)
{
    $nCtr++; Write-Progress -Activity "Reject List" -Status $nCtr -PercentComplete ($nCtr / $rd.Count * 100)
    if ($d.RejectMessagesFromSendersOrMembers.Count -gt 0)
    {
        $rTmp = @()
        foreach ($oReject in $d.RejectMessagesFromSendersOrMembers)
        {
            $oTmp = $null
            $oTmp = Get-Recipient -Identity $oReject.DistinguishedName -ea SilentlyContinue
            if ($null -ne $oTmp) {$rTmp += $oTmp}
        }
        if ($rTmp.Count -gt 0)
        {
            foreach ($oTmp in $rTmp)
            {
                $rFinal += [PSCustomObject]@{
                    DLGuid = $d.Guid
                    Guid = $oTmp.Guid
                    RecipientTypeDetails = $oTmp.RecipientTypeDetails
                }
            }
        }
    }
}
if ($rFinal.Count -gt 0) {$rFinal | Export-Csv exp-dls-reject.csv -NoTypeInformation}

#-- Get the moderation list
$nCtr = 0
$rFinal = @()
foreach ($d in $rd)
{
    $nCtr++; Write-Progress -Activity "Moderation List" -Status $nCtr -PercentComplete ($nCtr / $rd.Count * 100)
    if ($d.ModeratedBy.Count -gt 0)
    {
        $rTmp = @()
        foreach ($oModerator in $d.ModeratedBy)
        {
            $oTmp = $null
            $oTmp = Get-Recipient -Identity $oModerator.DistinguishedName -ea SilentlyContinue
            if ($null -ne $oTmp) {$rTmp += $oTmp}
        }
        if ($rTmp.Count -gt 0)
        {
            foreach ($oTmp in $rTmp)
            {
                $rFinal += [PSCustomObject]@{
                    DLGuid = $d.Guid
                    Guid = $oTmp.Guid
                    RecipientTypeDetails = $oTmp.RecipientTypeDetails
                }
            }
        }
    }
}
if ($rFinal.Count -gt 0) {$rFinal | Export-Csv exp-dls-moderation.csv -NoTypeInformation}

#-- Get the ManagedBy list
$nCtr = 0
$rFinal = @()
foreach ($d in $rd)
{
    $nCtr++; Write-Progress -Activity "ManagedBy List" -Status $d.DisplayName -PercentComplete ($nCtr / $rd.Count * 100)
    if ($d.ManagedBy.Count -gt 0)
    {
        $rTmp = @()
        foreach ($oManagedBy in $d.ManagedBy)
        {
            $oTmp = $null
            $oTmp = Get-Recipient -Identity $oManagedBy.DistinguishedName -ea SilentlyContinue
            if ($null -ne $oTmp) {$rTmp += $oTmp}
        }
        if ($rTmp.Count -gt 0)
        {
            foreach ($oTmp in $rTmp)
            {
                $rFinal += [PSCustomObject]@{
                    DLGuid = $d.Guid
                    Guid = $oTmp.Guid
                    RecipientTypeDetails = $oTmp.RecipientTypeDetails
                }
            }
        }
    }
}
if ($rFinal.Count -gt 0) {$rFinal | Export-Csv exp-dls-managedby.csv -NoTypeInformation}

#-- Get the member list
$nCtr = 0
$rFinal = @()
foreach ($d in $rd)
{
    $nCtr++; Write-Progress -Activity "Member List" -Status $d.DisplayName -PercentComplete ($nCtr / $rd.Count * 100)
    $rdm = @()
    $rdm = Get-DistributionGroupMember -Identity $d.Guid.ToString() -ResultSize 20000 -ea SilentlyContinue
    if ($rdm.Count -gt 0)
    {
        foreach ($oMember in $rdm)
        {
            $rFinal += [PSCustomObject]@{
                DLGuid = $d.Guid
                Guid = $oMember.Guid
                RecipientTypeDetails = $oMember.RecipientTypeDetails
            }
        }
    }
}
if ($rFinal.Count -gt 0) {$rFinal | Export-Csv exp-dls-members.csv -NoTypeInformation}