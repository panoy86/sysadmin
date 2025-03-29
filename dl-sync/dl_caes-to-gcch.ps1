$script:sWorkFile    = ".\dl_workfile.csv"
$script:sMembersFile = ".\dl_members-source-target.csv"
$script:sMappingFile = ".\dl_users-mapping.csv"

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

#------------------------------------------------------------------------------
#-- Update the working file
#------------------------------------------------------------------------------
function UpdateWorkFile
{
    Write-Host "Updating the working file" -ForegroundColor Green

    #-- Check if the working file exists, if not create it
    if (-not (Test-Path $sWorkFile)) {$rWork = @()}
    else {$rWork = Import-Csv $sWorkFile}

    #-- Add to the working file
    $nAdd = 0
    $rDLs = Import-Csv ".\exp-dls.csv"
    $hTmp = @{}
    if ($rWork.Count -gt 0) {$rWork | ForEach-Object {$hTmp.Add($_.SrcGuid, 1)}}
    foreach ($oDL in $rDLs)
    {
        if (-not $hTmp.ContainsKey($oDL.Guid))
        {
            $oNew = [PSCustomObject]@{
                SrcGuid = $oDL.Guid
                TgtGuid = $null
                State = "New"
                Error = $null
                LastUpdate = (Get-Date).ToString()
            }
            $rWork += $oNew
            $nAdd++
        }
    }
    Write-Host "  New DL working entries added:" $nAdd

    #-- Set to delete on target
    $nDelete = 0
    $hTmp = @{}
    $rDLs | ForEach-Object {$hTmp.Add($_.Guid, 1)}
    foreach ($oItem in $rWork)
    {
        $sKey = $oItem.SrcGuid
        if (-not $hTmp.Contains($sKey))
        {
            $oItem.State = "Delete"
            $oItem.LastUpdate = (Get-Date).ToString()
            $nDelete++
        }
    }
    Write-Host "  DLs to delete from target:" $nDelete

    #-- Save the working file
    $rWork | Export-Csv $sWorkFile -NoTypeInformation
}

#------------------------------------------------------------------------------
#-- Create a new DL
#------------------------------------------------------------------------------
function CreateDL
{
    Write-Host "Creating new DLs" -ForegroundColor Green

    #-- Get the working files
    $rWork = Import-Csv $sWorkFile
    $rDLs = Import-Csv ".\exp-dls.csv"
    $hDLs = @{}
    $rDLs | ForEach-Object {$hDLs.Add($_.Guid, $_)}

    #-- Loop through the working file
    $nCount = 0
    foreach ($oItem in $rWork)
    {
        if ($oItem.State -eq "New" -and $oItem.TgtGuid.Length -eq 0)
        {
            #-- Get the details
            $oDL = $hDLs[$oItem.SrcGuid]
            if ($null -eq $oDL)
            {
                $oItem.State = "Error"
                $oItem.Error = "Source DL not found"
                $oItem.LastUpdate = (Get-Date).ToString()
                continue
            }

            #-- Bypass if the DL is hidden, is roomlist
            if ($oDL.HiddenFromAddressListsEnabled -eq "true") {continue}
            if ($oDL.RecipientTypeDetails -eq "RoomList") {continue}

            #-- Check first if another object with the same name exists
            $bExists = $false
            $sNewEmail = ($oDL.PrimarySmtpAddress -split "@")[0] + "@honeywell.us"
            if ($null -ne (Get-Recipient $sNewEmail -ea SilentlyContinue)) {$bExists = $true}
            if ($null -ne (Get-Recipient $oDL.Alias -ea SilentlyContinue)) {$bExists = $true}
            if ($null -ne (Get-Recipient $oDL.Name -ea SilentlyContinue)) {$bExists = $true}
            if ($bExists)
            {
                $oItem.State = "Error"
                $oItem.Error = "Target DL already exists"
                $oItem.LastUpdate = (Get-Date).ToString()
                continue
            }

            #-- Create the DL
            $sNewDisplayName = "DLC-" + $oDL.DisplayName
            Write-Host "  Creating:" $sNewDisplayName -NoNewline
            Write-Host " $($sNewEmail)" -ForegroundColor Yellow
            if ($oDL.RecipientTypeDetails -eq "MailUniversalDistributionGroup") {$null = New-DistributionGroup -DisplayName $sNewDisplayName -Alias $oDL.Alias -Name $oDL.Name -PrimarySmtpAddress $sNewEmail -Type Distribution}
            else {$null = New-DistributionGroup -DisplayName $sNewDisplayName -Alias $oDL.Alias -Name $oDL.Name -PrimarySmtpAddress $sNewEmail -Type Security}

            #-- Verify the new DL
            $oNewDL = $null
            $oNewDL = Get-DistributionGroup $sNewEmail -ea SilentlyContinue
            if ($null -ne $oNewDL)
            {
                $oItem.TgtGuid = $oNewDL.Guid.ToString()
                $oItem.State = "Created"
                $oItem.Error = $null
                $oItem.LastUpdate = (Get-Date).ToString()
                $null = Set-DistributionGroup $sNewEmail -CustomAttribute15 "caes-dl" -HiddenFromAddressListsEnabled:$true
                $nCount++
            }
            else
            {
                $oItem.Error = "Failed to create new DL"
                $oItem.LastUpdate = (Get-Date).ToString()
            }
            if ($nCount -ge 50) {break}  #-- Limit the number of DLs created
        }
    }
    Write-Host "  New DLs: $nCount"
    $rWork | Export-Csv $sWorkFile -NoTypeInformation
}

#------------------------------------------------------------------------------
#-- Create a mapping of the members from source to target
#------------------------------------------------------------------------------
function CreateMemberMapping
{
    Write-Host "Creating a mapping list of DL members, from source to target" -ForegroundColor Green

    #-- Get the export for DLs, Mailboxes, and Mail Contacts into 1 list
    if (Test-Path $script:sMembersFile) {$rMembers = Import-Csv $script:sMembersFile}
    else {$rMembers = @()}
    $rTmp = @()
    Import-Csv .\exp-dls.csv | ForEach-Object {$rTmp += [PSCustomObject]@{SrcGuid = $_.Guid; SrcEmail = $_.PrimarySmtpAddress; SrcType = $_.RecipientTypeDetails; TgtGuid = $null; TgtEmail = $null; TgtType = $null}}
    Import-Csv .\exp-mbx.csv | ForEach-Object {$rTmp += [PSCustomObject]@{SrcGuid = $_.Guid; SrcEmail = $_.PrimarySmtpAddress; SrcType = $_.RecipientTypeDetails; TgtGuid = $null; TgtEmail = $null; TgtType = $null}}
    Import-Csv .\exp-contacts.csv | ForEach-Object {$rTmp += [PSCustomObject]@{SrcGuid = $_.Guid; SrcEmail = $_.PrimarySmtpAddress; SrcType = "MailContact"; TgtGuid = $null; TgtEmail = $null; TgtType = $null}}

    #-- Add to our existing list
    $hTmp = @{}
    $rMembers | ForEach-Object {$hTmp.Add($_.SrcGuid, $_)}
    $nNew = 0
    $rTmp | ForEach-Object {if (-not $hTmp.ContainsKey($_.SrcGuid)) {$rMembers += $_; $nNew++}}
    Write-Host "  Total number of potential members: $($rMembers.Count)  New: $nNew"

    #-- Update the target DLs
    $rWork = Import-Csv $sWorkFile
    $hTmp = @{}
    $rWork | ForEach-Object {$hTmp.Add($_.SrcGuid, $_)}
    $nCtr = 0
    $nCount = 0
    #$ProgressPreference = "Continue"  #-- Re-enable the progress bar in Write-Progress
    foreach ($oItem in $rMembers)
    {
        #-- Show progress and bypass certain items
        $nCtr++
        if ($oItem.TgtGuid.Length -gt 0) {continue}  #-- already found
        Write-Progress -Activity "Mapping DL Members" -Status "Processing" -PercentComplete ($nCtr * 100 / $rMembers.Count)

        #-- Check if the source is a DL
        if ($hTmp.ContainsKey($oItem.SrcGuid))
        {
            $oWork = $hTmp[$oItem.SrcGuid]
            if ($oWork.State -eq "Created" -and $oWork.TgtGuid.Length -gt 0)
            {
                $oItem.TgtGuid = $oWork.TgtGuid
                $oTmp = $null
                $oTmp = Get-Recipient $oWork.TgtGuid -ea SilentlyContinue
                if ($null -ne $oTmp)
                {
                    $oItem.TgtEmail = $oTmp.PrimarySmtpAddress
                    $oItem.TgtType = $oTmp.RecipientTypeDetails
                    $nCount++
                }
            }
        }
    }
    Write-Progress -Activity "Mapping DL Members" -Completed
    Write-Host "  New DL members found:" $nCount

    #-- Update the target mailboxes
    $rMapping = Import-Csv $script:sMappingFile
    $hTmp = @{}
    $rMapping | ForEach-Object {if ($_.caesguid.Length -gt 0) {$hTmp.Add($_.caesguid, $_.honemail)}}
    $nCountMapped = 0
    $nCountFound = 0
    $nCtr = 0
    foreach ($oItem in $rMembers)
    {
        #-- Show progress, bypass certain items
        $nCtr++
        if ($oItem.SrcType -notmatch "Mailbox") {continue}
        if ($oItem.TgtGuid.Length -gt 0 -and $oItem.TgtType -match "Mailbox") {continue}  #-- already found
        Write-Progress -Activity "Mapping Mailbox Members" -Status $oItem.SrcEmail -PercentComplete ($nCtr * 100 / $rMembers.Count)

        #-- Check if the source is a mailbox, mailuser, or a temporary mail-contact
        if ($hTmp.ContainsKey($oItem.SrcGuid))
        {
            $nCountMapped++
            $sTgtEmail = $hTmp[$oItem.SrcGuid]
            $oTmp = $null
            $oTmp = Get-Recipient $sTgtEmail -ea SilentlyContinue
            if ($null -ne $oTmp)
            {
                $oItem.TgtGuid = $oTmp.Guid.ToString()
                $oItem.TgtEmail = $oTmp.PrimarySmtpAddress
                $oItem.TgtType = $oTmp.RecipientTypeDetails
                $nCountFound++
            }
        }
    }
    Write-Progress -Activity "Mapping Mailbox Members" -Completed
    Write-Host "  New mailbox/mailuser members found:" $nCountFound "Mapped:" $nCountMapped

    #-- Update that target contacts
    $rWork = Import-Csv .\exp-contacts.csv
    $hTmp = @{}
    $rWork | ForEach-Object {$hTmp.Add($_.Guid, $_.PrimarySmtpAddress)}
    $nCountMapped = 0
    $nCountFound = 0
    $nCtr = 0
    foreach ($oItem in $rMembers)
    {
        #-- Show progress and bypass certain items
        $nCtr++
        if ($oItem.SrcType -notmatch "MailContact") {continue}
        if ($oItem.TgtGuid.Length -gt 0) {continue}  #-- already found
        Write-Progress -Activity "Mapping Contact Members" -Status $oItem.SrcEmail -PercentComplete ($nCtr * 100 / $rMembers.Count)

        #-- Check if the source is a contact
        if ($hTmp.ContainsKey($oItem.SrcGuid))
        {
            $nCountMapped++
            $sTgtEmail = $hTmp[$oItem.SrcGuid]
            $oTmp = $null
            $oTmp = Get-MailContact $sTgtEmail -ea SilentlyContinue
            if ($null -ne $oTmp)
            {
                $oItem.TgtGuid = $oTmp.Guid.ToString()
                $oItem.TgtEmail = $oTmp.PrimarySmtpAddress
                $oItem.TgtType = $oTmp.RecipientTypeDetails
                $nCountFound++
            }
        }
    }
    Write-Progress -Activity "Mapping Contact Members" -Completed
    Write-Host "  New contact members found:" $nCountFound "Mapped:" $nCountMapped

    #-- Save the mapping file
    $rMembers | Export-Csv $script:sMembersFile -NoTypeInformation
}

#------------------------------------------------------------------------------
#-- Create mail contacts for users not yet migrated to target
#------------------------------------------------------------------------------
function CreateMailContacts
{
    Write-Host "Creating mail contacts for users not yet migrated to target" -ForegroundColor Green

    #-- Get all the dl-member objects and the mapping file
    $rMembers = Import-Csv $script:sMembersFile
    $rMappings = Import-Csv $script:sMappingFile
    $hMembers = @{}
    $rMembers | ForEach-Object {$hMembers.Add($_.SrcGuid, $_)}

    #-- Loop thru our user-mapping file
    $nCountCreated = 0
    foreach ($oMap in $rMappings)
    {
        #-- Check if the email is valid
        $sSrcEmail = $oMap.caesemail.Trim()
        if ($sSrcEmail.Length -eq 0) {continue}

        #-- Check if the source guid is valid and find the member data
        if ($oMap.caesguid.Trim.Length -eq 0) {continue}
        if (-not $hMembers.ContainsKey($oMap.caesguid)) {continue}

        #-- Check if the user is already defined in the target
        $oMember = $hMembers[$oMap.caesguid]
        if ($oMember.TgtGuid.Length -gt 0) {continue}

        #-- Create a new mail contact or assign the existing one
        $oTmp = $null
        $oTmp = Get-Recipient $sSrcEmail -ea SilentlyContinue
        if ($null -eq $oTmp)
        {
            #-- Create the mail contact
            Write-Host "  Creating mail contact:" $sSrcEmail
            $sAlias = "caes_" + (-join ((48..57) + (97..122) | Get-Random -Count 10 | ForEach-Object {[char]$_}))
            $null = New-MailContact -Alias $sAlias -Name $sAlias -DisplayName $sAlias -ExternalEmailAddress $sSrcEmail
            $null = Set-MailContact $sAlias -CustomAttribute15 "caes-mail-contact" -HiddenFromAddressListsEnabled:$true
            $nCountCreated++
        }
        #-- Confirm if new contact was created or get the existing object
        $oTmp = Get-Recipient $sSrcEmail -ea SilentlyContinue
        if ($null -ne $oTmp)
        {
            $oMember.TgtGuid = $oTmp.Guid.ToString()
            $oMember.TgtEmail = $null
            $oMember.TgtType = $oTmp.RecipientTypeDetails
            #-- Flag a warning if the object is not a mail contact
            if ($oTmp.RecipientTypeDetails -ne "MailContact") {Write-Host "  WARNING: Object is not a mail contact:" $sSrcEmail -ForegroundColor Yellow}
        }
        if ($nCountCreated -gt 500) {break}  #-- Limit the number of contacts created
    }
    Write-Host "  Mail contacts created:" $nCountCreated

    #-- Save the member ojects
    $rMembers | Export-Csv $script:sMembersFile -NoTypeInformation
}

#------------------------------------------------------------------------------
#-- Remove mail contacts for mailboxes that are migrated
#------------------------------------------------------------------------------
function RemoveMailContacts
{
    Write-Host "Removing mail contacts for mailboxes that are migrated" -ForegroundColor Green

    #-- Get all the dl-member objects and the mapping file
    $rMembers = Import-Csv $script:sMembersFile
    $rMappings = Import-Csv $script:sMappingFile
    $hMembers = @{}
    $rMembers | ForEach-Object {$hMembers.Add($_.SrcGuid, $_)}

    #-- Loop thru our user-mapping file
    $nCtr = 0
    $nCountRemoved = 0
    foreach ($oMap in $rMappings)
    {
        #-- Show progress and bypass certain items
        $nCtr++
        Write-Progress -Activity "Removing mail contacts" -Status '.' -PercentComplete ($nCtr * 100 / $rMappings.Count)

        #-- Check if the email is valid
        $sSrcEmail = $oMap.caesemail.Trim()
        if ($sSrcEmail.Length -eq 0) {continue}

        #-- Check if the source guid is valid and find the member data
        if ($oMap.caesguid.Trim.Length -eq 0) {continue}
        if (-not $hMembers.ContainsKey($oMap.caesguid)) {continue}

        #-- Get the member data
        if ($hMembers.ContainsKey($oMap.caesguid))
        {
            $oMember = $hMembers[$oMap.caesguid]
            if ($oMember.TgtGuid.Length -gt 0 -and $oMember.TgtType -match "Mailbox")
            {
                #-- Check if the source email is a mail contact
                $oTmp = Get-MailContact $sSrcEmail -ea SilentlyContinue
                if ($null -ne $oTmp)
                {
                    #-- Remove the mail contact
                    Write-Host "  Removing mail contact:" $sSrcEmail
                    #Remove-MailContact -Identity $sSrcEmail -Confirm:$false
                    $nCountRemoved++
                }
            }
        }
    }
    Write-Progress -Activity "Removing mail contacts" -Completed
    Write-Host "  Mail contacts removed:" $nCountRemoved
}

#------------------------------------------------------------------------------
#-- Update the members of the DLs
#------------------------------------------------------------------------------
function UpdateMembers
{
    Write-Host "Updating the members of the DLs" -ForegroundColor Green

    #-- Get the working file and DL member mappings
    $rWork = Import-Csv $script:sWorkFile
    $rExportedMembers = Import-Csv .\exp-dls-members.csv

    #-- Loop thru our list of DLs
    $nCtr = 0
	$nCount = 0
    $rErrors = @()
    foreach ($oWork in $rWork)
    {
        #-- Show progress and bypass certain items
        $nCtr++
        Write-Progress -Activity "Updating DL Members" -Status $oWork.TgtGuid -PercentComplete ($nCtr * 100 / $rWork.Count)
        if ($oWork.State -ne "Created") {continue}  #-- Only process the created DLs
        
        #-- Get the target DL and its current members
        $oDL = Get-DistributionGroup $oWork.TgtGuid -ea SilentlyContinue
        if ($null -eq $oDL) {continue}  #-- DL not found
        $rTgtMembers = Get-DistributionGroupMember $oDL.Identity -ResultSize 10000 -ea SilentlyContinue
        
        #-- Get the list of members from the source
        $rSrcMembers = $rExportedMembers | Where-Object {$_.DLGuid -eq $oWork.SrcGuid}
        Write-Host " " $oWork.TgtGuid -NoNewline
        Write-Host " - $($rSrcMembers.Count) members"

        #-- Convert the source GUID to target GUID
        $rSrcGuids = @()
        foreach ($oSrcMember in $rSrcMembers)
        {
            $sTargetGuid = GetTargetGuid $oSrcMember.Guid
            if ($null -ne $sTargetGuid)
            {
                #Write-Host $oSrcMember.Guid "->" $sTargetGuid
                $rSrcGuids += $sTargetGuid
            }
            else
            {
                $rErrors += [PSCustomObject]@{
                    SrcGuid = $oSrcMember.Guid
                    SrcType = $oSrcMember.RecipientTypeDetails
                    State = "Error"
                    Error = "Member not found in source"
                }
            }
        }

        #-- Add new members to the DL
        $hTgtMembers = @{}
        $rTgtMembers | ForEach-Object {$hTgtMembers.Add($_.Guid.ToString(), 1)}
        foreach ($sSrcGuid in $rSrcGuids)
        {
            if (-not $hTgtMembers.ContainsKey($sSrcGuid))
            {
                #-- Add the member to the DL
                Write-Host "    Adding: " $sSrcGuid -ForegroundColor Green -NoNewline
				Write-Host $sSrcGuid
                $null = Add-DistributionGroupMember -Identity $oDL.Identity -Member $sSrcGuid -Confirm:$false
            }
        }

        #-- Remove members that are not in the source DL
        $hSrcGuids = @{}
        $rSrcGuids | ForEach-Object {$hSrcGuids.Add($_, 1)}
        foreach ($oTgtMember in $rTgtMembers)
        {
            if (-not $hSrcGuids.ContainsKey($oTgtMember.Guid.ToString()))
            {
                #-- Remove the member from the DL
                Write-Host "    Removing: " -ForegroundColor Green -NoNewline
				Write-Host $oTgtMember.DisplayName
                $null = Remove-DistributionGroupMember -Identity $oDL.Identity -Member $oTgtMember.Guid.ToString() -Confirm:$false
            }
        }

        #-- For testing, stop at a breakpoint
		$nCount++
		if ($nCount -ge 100) {break}
    }
    if ($rErrors.Count -gt 0)
    {
        Write-Host "  Errors found:" $rErrors.Count
        $rErrors | Export-Csv ".\dl_errors.csv" -NoTypeInformation
    }
    Write-Progress -Activity "Updating DL Members" -Completed
}

#------------------------------------------------------------------------------
#-- Main
#------------------------------------------------------------------------------
$PSStyle.Progress.View = "Minimal"  #-- Other values: "Classic"
$ProgressPreference = "Continue"

UpdateWorkFile  #-- Update the list of CAES DLs to create or delete from target
if ((Read-Host "  Creating a member mapping for 6k objects can take an hour, do you want to proceed? (yes/no)").Trim().ToLower() -match "y") {CreateMemberMapping}  #-- Update the mapping file with source and target members

#CreateDL
#CreateMailContacts  #-- Will create mail contacts for users not yet migrated to target
#RemoveMailContacts  #-- Will remove mail contacts for users that are migrated
UpdateMembers

#-- Show stats for our work file
#$rWork = Import-Csv $sWorkFile
#$h = @{}; $rWork | ForEach-Object {$sKey = $_.State; if (-not $h.Contains($sKey)) {$h.Add($sKey, 1)} else {$h[$sKey]++}}; $h
