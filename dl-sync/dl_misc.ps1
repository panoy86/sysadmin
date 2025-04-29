$script:sMembersFile = ".\dl_members-source-target.csv"

#------------------------------------------------------------------------------
#-- Aux function to return the target GUID, based on the source GUID
#------------------------------------------------------------------------------
function GetTargetGuid
{
    param (
        [string]$sSrcGuid
    )
    #-- Create a lookup table for source and target GUIDs
    if ($null -eq (Get-Variable -Name hSrcTgtGuidLookup -ea SilentlyContinue))
    {
        $script:hSrcTgtGuidLookup = @{}  #-- Create a new hashtable
        Import-Csv $script:sMembersFile | ForEach-Object {
            if ($_.SrcGuid.Length -gt 0 -and $_.TgtGuid.Length -gt 0)
            {
                $script:hSrcTgtGuidLookup.Add($_.SrcGuid, $_.TgtGuid)
            }
        }
    }

    #-- Check if the source GUID is in the lookup table
    if ($script:hSrcTgtGuidLookup.ContainsKey($sSrcGuid))
    {
        return $script:hSrcTgtGuidLookup[$sSrcGuid]
    } 
    return $null
}

#-- Adds new entries to the mapping file (e.g., new hires)
function UpdateMappingFile_AddNew
{
    #-- Add new entries to our mapping file, check if it exist othwerwise exit
    $sAddFile = ".\add.csv"  #-- Expect 3 columns: hid, honemail, caesemail
    if (-not (Test-Path $sAddFile)) {Write-Host "File" $sAddFile "not found"; return}
    #-- Get the mapping file
    $sFile = ".\dl_users-mapping.csv"
    $r = Import-Csv $sFile
    $rAdd = Import-Csv $sAddFile
    $hHon = @{}; $r | ForEach-Object {Try {$hHon.Add($_.honemail, 1)} Catch {}}
    $hCaes = @{}; $r | ForEach-Object {Try {$hCaes.Add($_.caesemail, 1)} Catch {}}
    $hTmp = @{}
    $nAdded = 0
    foreach ($t in $rAdd)
    {
        #-- Check for existing entries
        $bAdd = $true
        if ($hHon.ContainsKey($t.honemail)) {$bAdd = $false}
        if ($hCaes.ContainsKey($t.caesemail)) {$bAdd = $false}
        if ($bAdd -eq $true)
        {
            #-- Add new entry to the mapping file, but also look for duplicated in the add file
            if (-not $hTmp.ContainsKey($t.hid))
            {
                $hTmp.Add($t.hid, $t.honemail)
                $r += [PSCustomObject]@{
                    hid = $t.hid
                    honemail = $t.honemail
                    caesemail = $t.caesemail
                    caesguid = ''
                }
                $nAdded++
            }
        }
    }
    Write-Host "Added:" $nAdded "of" $rAdd.Count
    $r | Export-Csv $sFile -NoTypeInformation
}

#-- Update the mapping file with caes-guid
function UpdateMappingFile_CaesGUID
{
    #-- Update mapping file with caes-guid
    $sFile = ".\dl_users-mapping.csv"
    $r = Import-Csv $sFile
    $h = @{}
    Import-Csv exp-mbx.csv | ForEach-Object {$h.Add($_.PrimarySmtpAddress, $_.Guid)}
    $nUpdated = 0
    foreach ($t in $r)
    {
        #-- Bypass if GUID is already populated
        if ($t.caesguid -ne '') {continue}
        $sKey = $t.caesemail
        if ($h.ContainsKey($sKey))
        {
            $t.caesguid = $h[$sKey]
            $nUpdated++
        }
    }
    Write-Host "GUIDs updated:" $nUpdated
    Write-Host "Entries without GUID:" ($r | Where-Object {$_.caesguid.Length -eq 0}).Count
    $r | Export-Csv $sFile -NoTypeInformation
}

#-- Update the CU properties for the mail contacts
function SetMailContactProperties
{
    $rm = Import-Csv .\dl_users-mapping.csv
    $rmc = @()
    $nCtr = 0
    $nCount = 0
    foreach ($m in $rm)
    {
        #-- Show progress
        $nCtr++
        Write-Progress -Activity "Processing" -Status $m.hid -PercentComplete ($nCtr / $rm.Count * 100)

        #-- Get the source mailbox or mailuser object
        $oSrc = $null
        $oSrc = Get-MailBox $m.honemail -ea SilentlyContinue
        if ($null -eq $oSrc) {$oSrc = Get-MailUser $m.honemail -ea SilentlyContinue}
        if ($null -eq $oSrc) {continue}

        #-- Get the target mailcontact object
        $oTgt = $null
        $oTgt = Get-MailContact $m.caesemail -ea SilentlyContinue
        if ($null -eq $oTgt) {continue}
        $rmc += $oTgt

        #-- Set the target's CustomAttribute(s) to be the same as the source
        $bChange = $false
        if ($oSrc.CustomAttribute1 -ne $oTgt.CustomAttribute1) {$bChange = $true}
        if ($oSrc.CustomAttribute2 -ne $oTgt.CustomAttribute2) {$bChange = $true}
        if ($oSrc.CustomAttribute3 -ne $oTgt.CustomAttribute3) {$bChange = $true}
        if ($oSrc.CustomAttribute4 -ne $oTgt.CustomAttribute4) {$bChange = $true}
        if ($oSrc.CustomAttribute5 -ne $oTgt.CustomAttribute5) {$bChange = $true}
        if ($oSrc.CustomAttribute6 -ne $oTgt.CustomAttribute6) {$bChange = $true}
        if ($oSrc.CustomAttribute7 -ne $oTgt.CustomAttribute7) {$bChange = $true}
        if ($oSrc.CustomAttribute8 -ne $oTgt.CustomAttribute8) {$bChange = $true}
        if ($oSrc.CustomAttribute9 -ne $oTgt.CustomAttribute9) {$bChange = $true}
        if ($oSrc.CustomAttribute10 -ne $oTgt.CustomAttribute10) {$bChange = $true}
        if ($oSrc.CustomAttribute11 -ne $oTgt.CustomAttribute11) {$bChange = $true}
        if ($oSrc.CustomAttribute12 -ne $oTgt.CustomAttribute12) {$bChange = $true}
        if ($oSrc.DisplayName -ne $oTgt.DisplayName) {$bChange = $true}
        if ($bChange -eq $true)
        {
            Write-Host "Updating $($oTgt.PrimarySmtpAddress)"
            Set-MailContact -Identity $oTgt.Guid.ToString() -CustomAttribute1 $oSrc.CustomAttribute1 -CustomAttribute2 $oSrc.CustomAttribute2 -CustomAttribute3 $oSrc.CustomAttribute3 -CustomAttribute4 $oSrc.CustomAttribute4 -CustomAttribute5 $oSrc.CustomAttribute5 -CustomAttribute6 $oSrc.CustomAttribute6 -CustomAttribute7 $oSrc.CustomAttribute7 -CustomAttribute8 $oSrc.CustomAttribute8 -CustomAttribute9 $oSrc.CustomAttribute9 -CustomAttribute10 $oSrc.CustomAttribute10 -CustomAttribute11 $oSrc.CustomAttribute11 -CustomAttribute12 $oSrc.CustomAttribute12 -DisplayName $oSrc.DisplayName -ea SilentlyContinue
            $nCount++
        }
    }
    $rmc | Export-Csv -Path .\zz_contacts.csv -NoTypeInformation
    Write-Progress -Activity "Processing" -Completed
    Write-Host "Total number of updated mail contacts:" $nCount
}

#-- Update the members of the static DLs representing the dynamic DLs from the source tenant
function UpdateStaticDLs
{
    $rd = Import-Csv .\ddl-manually-created.csv
    $rm = Import-Csv .\exp-mbx.csv
    foreach ($d in $rd)
    {
        #-- Bypass certain static DLs
        if ($d.Name -notmatch "Users") {continue}
        $sLid = ($d.Name -split '\.')[0]
        if ($sLid -in ("CAES", "Frontgrade")) {continue}
        Write-Host $d.Name
        #-- Get the existing members of the static DL
        $oDist = Get-DistributionGroup $d.Name -ea SilentlyContinue
        if ($null -eq $oDist) {continue}
        $rTmp = Get-DistributionGroupMember -Identity $oDist.Guid.ToString() -ResultSize 5000 -ea SilentlyContinue
        $hPresentMembers = @{}
        $rTmp | ForEach-Object {$hPresentMembers.Add($_.Guid.Tostring(), 1)}
        #-- Loop thru our mailboxes' OU
        $rMembers = @()
        foreach ($m in $rm)
        {
            if ($m.OrganizationalUnit -match $sLid) {$rMembers += $m}
        }
        #-- Brute-force add the members to the static DL
        $nCount = 0
        foreach ($oMember in $rMembers)
        {
            $sTgtGuid = GetTargetGuid $oMember.Guid.ToString()            
            if ($null -eq $sTgtGuid) {} #Write-Host "  " $oMember.Guid.ToString()}
            else
            {
                #-- Check if the member is already present in the static DL
                if (-not $hPresentMembers.ContainsKey($sTgtGuid))
                {
                    Write-Progress -Activity "Adding members" -Status ($oMember.Guid.ToString() + " -> " + $sTgtGuid)
                    Add-DistributionGroupMember -Identity $d.Name -Member $sTgtGuid -BypassSecurityGroupManagerCheck -ea SilentlyContinue
                    $nCount++
                }
            }
        }
        Write-Host "  " $nCount "found/added, total from source:" $rMembers.Count

        #-- Lazy way to remove the mail-contacts from the static DL, where the mailbox is already a member
        $rTmp = Get-DistributionGroupMember -Identity $oDist.Guid.ToString() -ResultSize 5000 -ea SilentlyContinue
        $hPresentMembers = @{}
        $rTmp | ForEach-Object {$hPresentMembers.Add($_.Guid.Tostring(), 1)}
        $nCount = 0
        foreach ($oMember in $rTmp)
        {
            if ($oMember.RecipientTypeDetails -eq "UserMailbox")
            {
                $oMailbox = $null
                $oMailbox = Get-Mailbox $oMember.Guid.ToString() -ea SilentlyContinue
                if (($null -ne $oMailbox) -and ($null -ne $oMailbox.ForwardingAddress))
                {
                    $oContact = $null
                    $oContact = Get-MailContact $oMailbox.ForwardingAddress -ea SilentlyContinue
                    if (($null -ne $oContact) -and $hPresentMembers.ContainsKey($oContact.Guid.ToString()))
                    {
                        Write-Progress -Activity "Removing forwarders as members" -Status $oContact.PrimarySmtpAddress.ToString()
                        Remove-DistributionGroupMember -Identity $oDist.Guid.ToString() -Member $oContact.Guid.ToString() -BypassSecurityGroupManagerCheck -Confirm:$false -ea SilentlyContinue
                        $nCount++
                    }
                }
            }
        }
        Write-Progress -Activity "Removing forwarders as members" -Completed
        Write-Host "  " $nCount "forwarders removed from static DL"
    }
}

#-- Main
$PSStyle.Progress.View = "Minimal"  #-- Other values: "Classic"
$ProgressPreference = "Continue"

if ((Read-Host "Update the static DLs representation of the CAES dynamic DLs? (Y/N)").ToUpper() -eq 'Y')
{
    UpdateStaticDLs
}

Write-Host "You can add new entries to the mapping file" -ForegroundColor Green
Write-Host "Just create a CSV file called add.csv with the following columns:" -ForegroundColor Green
Write-Host "hid,honemail,caesemail" -ForegroundColor Yellow
if ((Read-Host "Add new entries to the mapping file? (Y/N)").ToUpper() -eq 'Y')
{
    UpdateMappingFile_AddNew
}
if ((Read-Host "Update the mapping file with caes-guid? (Y/N)").ToUpper() -eq 'Y')
{
    UpdateMappingFile_CaesGUID
}
if ((Read-Host "Update the mail contact properties? (Y/N)").ToUpper() -eq 'Y')
{
    SetMailContactProperties
}
Remove-Variable hSrcTgtGuidLookup
#-- End of script
