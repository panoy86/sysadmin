$script:sWorkFile    = ".\dl_workfile.csv"
$script:sMembersFile = ".\dl_members-source-target.csv"
$script:sMappingFile = ".\dl_users-mapping.csv"
$script:rErrors = @()

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

            #-- Bypass if the DL is hidden, is roomlist, except in our "force create" list
            $bBypass = $false
            #if ($oDL.HiddenFromAddressListsEnabled -eq "true") {$bBypass = $true}
            if ($oDL.RecipientTypeDetails -eq "RoomList") {$bBypass = $true}
            [array]$rTmp = (Get-Content dl_force-create-these.txt | Where-Object {$_.Trim() -notlike "#*"})
            if ($oItem.SrcGuid -in $rTmp) {$bBypass = $false}
            if ($bBypass) {continue}

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
            if ($nCount -ge 150) {break}  #-- Limit the number of DLs created
        }
    }
    Write-Host "  New DLs: $nCount"
    $rWork | Export-Csv $sWorkFile -NoTypeInformation
}

#------------------------------------------------------------------------------
#-- Update the properties of the DLs
#------------------------------------------------------------------------------
function UpdateDLProperties
{
    Write-Host "Updating the properties of the DLs" -ForegroundColor Green

    #-- Get the working files
    $rWork = Import-Csv $script:sWorkFile
    $rOwners = Import-Csv ".\exp-dls-managedby.csv"
    $rAccepts = Import-Csv ".\exp-dls-accept.csv"
    $hDLs = @{}
    Import-Csv ".\exp-dls.csv" | ForEach-Object {$hDLs.Add($_.Guid, $_)}

    #-- Loop thru our list of DLs
    $nCtr = 0
    foreach ($oWork in $rWork)
    {
        #-- Show progress and bypass certain items
        $nCtr++
        if ($oWork.State -ne "Created") {continue}  #-- Only process the created DLs
        $oDL = $null
        $oDL = Get-DistributionGroup $oWork.TgtGuid -ea SilentlyContinue
        if ($null -eq $oDL)
        {
            LogError $oWork.SrcGuid $oWork.SrcType $oWork.TgtGuid $oWork.TgtType "Target DL not found"
            continue
        }  #-- DL not found
        Write-Progress -Activity "Updating DL Properties" -Status $oDL.DisplayName -PercentComplete ($nCtr * 100 / $rWork.Count)

        #-- Update the owners
        [array]$rTmp = $rOwners | Where-Object {$_.DLGuid -match $oWork.SrcGuid}
        if ($rTmp.Count -gt 0)
        {
            #-- Normalize the current owners as a string
            $sOwners = $oDL.ManagedBy -join ','
            foreach ($oTmp in $rTmp)
            {
                #-- Get the target GUID if exists
                $sTgtGuid = GetTargetGuid $oTmp.Guid
                if ($null -ne $sTgtGuid)
                {
                    $oRecipient = $null
                    $oRecipient = Get-Recipient $sTgtGuid -ea SilentlyContinue
                    if ($null -ne $oRecipient)
                    {
                        #-- Add the owner to the DL
                        if ($sOwners -notmatch $oRecipient.Name)
                        {
                            Write-Host "  Adding owner: " -NoNewline -ForegroundColor Green
                            Write-Host $oWork.TgtGuid $oRecipient.DisplayName
                            Set-DistributionGroup $oWork.TgtGuid -ManagedBy @{Add=$($sTgtGuid)} -BypassSecurityGroupManagerCheck 
                        }
                        #-- Remove my account as owner
                        if ($sOwners -match "h592867")
                        {
                            $null = Set-DistributionGroup $oWork.TgtGuid -ManagedBy @{Remove="f292f968-c1c5-476c-a8d9-baf596dc01c0"}
                        }
                    }
                }
                else
                {
                    LogError $oTmp.Guid $oTmp.RecipientTypeDetails $null $null "Target DL owner not found"
                }
            }
        }

        #-- Update RequireSenderAuthenticationEnabled
        if ($hDLs.ContainsKey($oWork.SrcGuid))
        {
            $oSourceDL = $hDLs[$oWork.SrcGuid]
            $sRequire = $oSourceDL.RequireSenderAuthenticationEnabled
            #-- Check if we need to change it
            if ($oDL.RequireSenderAuthenticationEnabled -and $sRequire -eq "False")
            {
                Write-Host "  Updating RequireSenderAuthenticationEnabled: " -NoNewline -ForegroundColor Green
                Write-Host "From" $oDL.RequireSenderAuthenticationEnabled.ToString() "To" $sRequire
                $null = Set-DistributionGroup $oWork.TgtGuid -RequireSenderAuthenticationEnabled:$False
            }
            if (-not $oDL.RequireSenderAuthenticationEnabled -and $sRequire -eq "True")
            {
                Write-Host "  Updating RequireSenderAuthenticationEnabled: " -NoNewline -ForegroundColor Green
                Write-Host "From" $oDL.RequireSenderAuthenticationEnabled.ToString() "To" $sRequire
                $null = Set-DistributionGroup $oWork.TgtGuid -RequireSenderAuthenticationEnabled:$True
            }
        }

        #-- Add the LegacyExchangeDN as an x500 address
        if ($hDLs.ContainsKey($oWork.SrcGuid))
        {
            $oSourceDL = $hDLs[$oWork.SrcGuid]
            $sLegacyExchangeDN = $oSourceDL.LegacyExchangeDN
            $sEmailAddressString = $oDL.EmailAddresses -join ';'
            #-- Check if the x500 address already exists
            if ($null -eq ($sEmailAddressString | Select-String -Pattern $sLegacyExchangeDN -SimpleMatch))
            {
                Write-Host $sEmailAddressString $sLegacyExchangeDN
                Write-Host "  Adding LegacyExchangeDN: " -NoNewline -ForegroundColor Green
                Write-Host $sLegacyExchangeDN "To" $oDL.Name
                $null = Set-DistributionGroup $oWork.TgtGuid -EmailAddresses @{Add="x500:$sLegacyExchangeDN"}
            }
        }

        #-- Update the DL's accept list
        [array]$rTmp = $rAccepts | Where-Object {$_.DLGuid -match $oWork.SrcGuid}
        if ($rTmp.Count -gt 0)
        {
            #-- Normalize the current accept-list as a string
            $sAcceptList = $oDL.AcceptMessagesOnlyFromSendersOrMembers -join ','
            foreach ($oTmp in $rTmp)
            {
                #-- Get the target GUID if exists
                $sTgtGuid = GetTargetGuid $oTmp.Guid
                if ($null -ne $sTgtGuid)
                {
                    $oRecipient = $null
                    $oRecipient = Get-Recipient $sTgtGuid -ea SilentlyContinue
                    if ($null -ne $oRecipient)
                    {
                        #-- Add the owner to the DL
                        if ($sAcceptList -notmatch $oRecipient.Name)
                        {
                            Write-Host "  Adding accept-list: " -NoNewline -ForegroundColor Green
                            Write-Host $oWork.TgtGuid $oRecipient.DisplayName
                            Set-DistributionGroup $oWork.TgtGuid -AcceptMessagesOnlyFromSendersOrMembers @{Add=$($sTgtGuid)} -BypassSecurityGroupManagerCheck 
                        }
                    }
                    else 
                    {
                        LogError $oTmp.Guid $oTmp.RecipientTypeDetails $sTgtGuid $null "Accept-list target not found"
                    }
                }
                else
                {
                    LogError $oTmp.Guid $oTmp.RecipientTypeDetails $null $null "Accept-list source not found"
                }
            }
        }

        #-- Update if this DL is hidden or not
        if ($hDLs.ContainsKey($oWork.SrcGuid))
        {
            $oSourceDL = $hDLs[$oWork.SrcGuid]
            $sHidden = $oSourceDL.HiddenFromAddressListsEnabled
            if ($oDL.HiddenFromAddressListsEnabled -and $sHidden -eq "False")
            {
                Write-Host "  Updating HiddenFromAddressListsEnabled: " -NoNewline -ForegroundColor Green
                Write-Host "From" $oDL.HiddenFromAddressListsEnabled.ToString() "To" $sHidden
                $null = Set-DistributionGroup $oWork.TgtGuid -HiddenFromAddressListsEnabled:$false
            }
            if (-not $oDL.HiddenFromAddressListsEnabled -and $sHidden -eq "True")
            {
                Write-Host "  Updating HiddenFromAddressListsEnabled: " -NoNewline -ForegroundColor Green
                Write-Host "From" $oDL.HiddenFromAddressListsEnabled.ToString() "To" $sHidden
                $null = Set-DistributionGroup $oWork.TgtGuid -HiddenFromAddressListsEnabled:$true
            }
        }
        #-- Troubleshooting, break after N items processed
        #if ($nCtr -ge 1) {break}
    }
    Write-Progress -Activity "Updating DL Properties" -Completed
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

    #-- Update the source/target mailboxes
    $rMapping = Import-Csv $script:sMappingFile
    $hTmp = @{}
    $rMapping | ForEach-Object {if ($_.caesguid.Length -gt 0) {$hTmp.Add($_.caesguid, $_.honemail)}}
    $nCountMapped = 0
    $nCountFound = 0
    $nCtr = 0
    foreach ($oItem in $rMembers)
    {
        #if (1) {continue}  #-- Bypass this section for now
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
            $oTmp = Get-ExoMailbox $sTgtEmail -ea SilentlyContinue
            #-- Do not change if the target is not a mailbox-type
            if ($null -ne $oTmp -and $oTmp.RecipientTypeDetails -in ("UserMailbox", "SharedMailbox"))
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

    #-- Update that source/target contacts
    $rWork = Import-Csv .\exp-contacts.csv
    $hTmp = @{}
    $rWork | ForEach-Object {$hTmp.Add($_.Guid, $_.PrimarySmtpAddress)}
    $nCountMapped = 0
    $nCountFound = 0
    $nCtr = 0
    foreach ($oItem in $rMembers)
    {
        if (1) {continue}  #-- Bypass this section for now
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
    $nCtr = 0
    $nCountCreated = 0
    foreach ($oMap in $rMappings)
    {
        #-- Show progress and bypass certain items
        $nCtr++
        Write-Progress -Activity "Creating mail contacts" -Status '.' -PercentComplete ($nCtr * 100 / $rMappings.Count)

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
            $null = New-MailContact -Alias $sAlias -Name $sSrcEmail -DisplayName $sSrcEmail -ExternalEmailAddress $sSrcEmail
            $null = Set-MailContact $sAlias -CustomAttribute15 "caes-mail-contact" -HiddenFromAddressListsEnabled:$false
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
    }
    Write-Progress -Activity "Creating mail contacts" -Completed
    Write-Host "  Mail contacts created:" $nCountCreated

    #-- Save the member ojects
    $rMembers | Export-Csv $script:sMembersFile -NoTypeInformation
}

#------------------------------------------------------------------------------
#-- Remove mail contacts for mailboxes that are migrated
#------------------------------------------------------------------------------
function RemoveMailContacts
{
    Write-Host "Removing mail contacts for mailboxes that are fully migrated" -ForegroundColor Green

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
                #-- Check if the mailbox forwarding address is NOT set
                $oTmp = Get-Mailbox $oMember.TgtGuid -ea SilentlyContinue
                if (($null -ne $oTmp) -and ($null -eq $oTmp.ForwardingAddress))
                {
                    #-- Check if the source email is a mail contact
                    $oTmp = Get-MailContact $sSrcEmail -ea SilentlyContinue
                    if ($null -ne $oTmp)
                    {
                        #-- Remove the mail contact
                        Write-Host "  Removed mail contact:" $sSrcEmail
                        $null = Remove-MailContact -Identity $sSrcEmail -Confirm:$false
                        $nCountRemoved++
                    }
                }
            }
        }
    }
    Write-Progress -Activity "Removing mail contacts" -Completed
    Write-Host "  Mail contacts removed:" $nCountRemoved
}

#------------------------------------------------------------------------------
#-- Ensure the mailuser is hidden, and the mailcontact is visible from the GAL
#------------------------------------------------------------------------------
function UpdateMailObjectVisibility
{
    Write-Host "Updating mailuser and mailcontact visibility" -ForegroundColor Green

    #-- Get the working file
    $rMapping = Import-Csv $script:sMappingFile
    $nCtr = 0
    foreach ($oMap in $rMapping[1999..2415])
    {
        #-- Show progress and bypass certain items
        $nCtr++
        Write-Progress -Activity "Updating mail object visibility" -Status $oMap.hid -PercentComplete ($nCtr * 100 / $rMapping.Count)
        $oMailUser = $null
        $oMailUser = Get-MailUser $oMap.honemail -ea SilentlyContinue
        $oMailContact = $null
        $oMailContact = Get-MailContact $oMap.caesemail -ea SilentlyContinue
        if ($null -eq $oMailUser -or $null -eq $oMailContact) {continue}  #-- We don't have both objects
        if ($oMailUser.HiddenFromAddressListsEnabled -eq $false)
        {
            Write-Host "  Updating mailuser visibility: " -NoNewline -ForegroundColor Green
            Write-Host $oMailUser.DisplayName
            $null = Set-MailUser $oMailUser.Identity -HiddenFromAddressListsEnabled:$true
        }
        if ($oMailContact.HiddenFromAddressListsEnabled -eq $true)
        {
            Write-Host "  Updating mailcontact visibility: " -NoNewline -ForegroundColor Green
            Write-Host $oMailContact.DisplayName
            $null = Set-MailContact $oMailContact.Identity -HiddenFromAddressListsEnabled:$false
        }
    }
    Write-Progress -Activity "Updating mail object visibility" -Completed
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
    $rDoNotRemoveAsMembers = Get-Content .\dl_bypass-removal-members.txt | Where-Object {$_ -notlike "#*"}

    #-- Loop thru our list of DLs
    $nCtr = 0
    foreach ($oWork in $rWork)
    {
        #-- Show progress and bypass certain items
        $nCtr++
        if ($oWork.State -ne "Created") {continue}  #-- Only process the created DLs
        Write-Progress -Activity "Updating DL Members" -Status $oWork.TgtGuid -PercentComplete ($nCtr * 100 / $rWork.Count)
        
        #-- Get the target DL and its current members
        $oDL = Get-DistributionGroup $oWork.TgtGuid -ea SilentlyContinue
        if ($null -eq $oDL) {continue}  #-- DL not found
        $rTgtMembers = Get-DistributionGroupMember $oDL.Identity -ResultSize 10000 -ea SilentlyContinue
        
        #-- Get the list of members from the source
        $rSrcMembers = $rExportedMembers | Where-Object {$_.DLGuid -eq $oWork.SrcGuid}

        #-- Convert the source GUID to target GUID
        $rTgtMemberGuids = @()
        foreach ($oSrcMember in $rSrcMembers)
        {
            $sTargetGuid = GetTargetGuid $oSrcMember.Guid
            if ($null -ne $sTargetGuid) {$rTgtMemberGuids += $sTargetGuid}
            else {LogError $oSrcMember.Guid $oSrcMember.RecipientTypeDetails $null $null "Target member not found"}
        }

        #-- Add new members to the DL
        $hTgtMembers = @{}
        $rTgtMembers | ForEach-Object {$hTgtMembers.Add($_.Guid.ToString(), 1)}
        foreach ($sTgtMemberGuid in $rTgtMemberGuids)
        {
            if (-not $hTgtMembers.ContainsKey($sTgtMemberGuid))
            {
                #-- Add the member to the DL
                $oTmp = $null
                $oTmp = Get-ExoRecipient $sTgtMemberGuid -ea SilentlyContinue
                if ($null -ne $oTmp)
                {
                    Write-Host "    Adding: " -ForegroundColor Green -NoNewline
                    Write-Host $oTmp.DisplayName $oTmp.RecipientTypeDetails
                    $null = Add-DistributionGroupMember -Identity $oDL.Identity -Member $sTgtMemberGuid -Confirm:$false -BypassSecurityGroupManagerCheck -ea SilentlyContinue
                }
            }
        }

        #-- Remove members that are not in the source DL
        $hTgtMemberGuids = @{}
        $rTgtMemberGuids | ForEach-Object {$hTgtMemberGuids.Add($_, 1)}
        foreach ($oTgtMember in $rTgtMembers)
        {
            if (-not $hTgtMemberGuids.ContainsKey($oTgtMember.Guid.ToString()))
            {
                #-- Bypass certain members from removal
                if ($oTgtMember.Guid.ToString() -in $rDoNotRemoveAsMembers) {continue}
                #-- Remove the member from the DL
                Write-Host "    Removing: " -ForegroundColor Green -NoNewline
                Write-Host $oTgtMember.DisplayName $oTgtMember.RecipientTypeDetails
                $null = Remove-DistributionGroupMember -Identity $oDL.Identity -Member $oTgtMember.Guid.ToString() -Confirm:$false -BypassSecurityGroupManagerCheck -ea SilentlyContinue
            }
        }
    }
    Write-Progress -Activity "Updating DL Members" -Completed
}

#------------------------------------------------------------------------------
#-- Log an error
#------------------------------------------------------------------------------
function LogError
{
    param (
        [string]$sSrcGuid,
        [string]$sSrcType,
        [string]$sTgtGuid,
        [string]$sTgtType,
        [string]$sError
    )
    $oError = [PSCustomObject]@{
        SrcGuid = $sSrcGuid
        SrcType = $sSrcType
        TgtGuid = $sTgtGuid
        TgtType = $sTgtType
        Error = $sError
    }
    $script:rErrors += $oError
}

#------------------------------------------------------------------------------
#-- Analyze the error logs
#------------------------------------------------------------------------------
function AnalyzeErrors
{
    Write-Host "Analyzing the error logs" -ForegroundColor Green

    #-- Get the error logs
    $rErrors = Import-Csv ".\dl_errors.csv"
    $rTmp = @(); $h = @{}; $rErrors | ForEach-Object {$sKey = $_.SrcGuid; if (-not $h.ContainsKey($sKey)) {$h.Add($sKey,1); $rTmp += $_}}
    $h = @{}; $rTmp | ForEach-Object {$sKey = $_.SrcType; if ($h.ContainsKey($sKey)) {$h[$sKey]++} else {$h.Add($sKey,1)}}; $h.GetEnumerator() | Sort-Object Value

    $hTmp = @{}; $rErrors | ForEach-Object {$sKey = $_.Error; if (-not $hTmp.Contains($sKey)) {$hTmp.Add($sKey, 1)} else {$hTmp[$sKey]++}}
    Write-Host "  Total errors found:" $rErrors.Count
    Write-Host "  Errors by type:"
    foreach ($sKey in $hTmp.Keys) {Write-Host "   " $sKey ":" $hTmp[$sKey]}
}

#------------------------------------------------------------------------------
#-- Main
#------------------------------------------------------------------------------
$PSStyle.Progress.View = "Minimal"  #-- Other values: "Classic"
$ProgressPreference = "Continue"

UpdateWorkFile      #-- Update the list of CAES DLs to create or delete from target
CreateDL            #-- Create the new DLs in the target

if ((Read-Host "  Manage the mail-contacts? (Y/N)").Trim().ToLower() -match "y")
{
    CreateMailContacts  #-- Will create mail contacts for users not yet migrated to target
    RemoveMailContacts  #-- Will remove mail contacts for users that are migrated
}

if ((Read-Host "  Update the mail object visibility? (Y/N)").Trim().ToLower() -match "y")
{
    UpdateMailObjectVisibility  #-- Ensure the mailuser is hidden, and the mailcontact is visible from the GAL
}

if ((Read-Host "  Creating a member mapping can take an hour, do you want to proceed? (Y/N)").Trim().ToLower() -match "y")
{
    CreateMemberMapping  #-- Update the mapping file with source and target members
}

if ((Read-Host "  Update the DL properties and members? (Y/N)").Trim().ToLower() -match "y")
{
    UpdateDLProperties  #-- Update the properties of the DLs
    UpdateMembers
}

#-- Cleanup
$script:rErrors | Export-Csv ".\dl_errors.csv" -NoTypeInformation
Remove-Variable hSrcTgtGuidLookup -ea SilentlyContinue
Remove-Variable rErrors -ea SilentlyContinue
