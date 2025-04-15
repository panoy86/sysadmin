#-- Adds new entries to the mapping file (e.g., new hires)
function UpdateMappingFile_AddNew
{
    #-- Add new entries to our mapping file
    $sAddFile = "c:\temp\zz.csv"  #-- Expect 3 columns: hid, honemail, caesemail
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
function SetCU4MailContacts
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
        if ($bChange -eq $true)
        {
            Write-Host "Updating $($oTgt.PrimarySmtpAddress)"
            Set-MailContact -Identity $oTgt.Guid.ToString() -CustomAttribute1 $oSrc.CustomAttribute1 -CustomAttribute2 $oSrc.CustomAttribute2 -CustomAttribute3 $oSrc.CustomAttribute3 -CustomAttribute4 $oSrc.CustomAttribute4 -CustomAttribute5 $oSrc.CustomAttribute5 -CustomAttribute6 $oSrc.CustomAttribute6 -CustomAttribute7 $oSrc.CustomAttribute7 -CustomAttribute8 $oSrc.CustomAttribute8 -CustomAttribute9 $oSrc.CustomAttribute9 -CustomAttribute10 $oSrc.CustomAttribute10 -CustomAttribute11 $oSrc.CustomAttribute11 -CustomAttribute12 $oSrc.CustomAttribute12 -ea SilentlyContinue
            $nCount++
        }
    }
    $rmc | Export-Csv -Path .\zz_contacts.csv -NoTypeInformation
    Write-Progress -Activity "Processing" -Completed
    Write-Host "Total number of updated mail contacts:" $nCount
}


#-- Main
$PSStyle.Progress.View = "Minimal"  #-- Other values: "Classic"
$ProgressPreference = "Continue"

#UpdateMappingFile_AddNew
#UpdateMappingFile_CaesGUID
SetCU4MailContacts
#-- End of script