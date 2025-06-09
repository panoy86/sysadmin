#-- Script to check for DLs (and all members of sub-DLs) that have restrictions on who can send emails to them.
#-- Assumes that an existing PowerShell session to Exchange/Online is already set
param (
    [string] $DistributionLists,  #-- Comma-separated list of Distribution Lists to search
    [string] $Keyword             #-- Keyword to search in DL names, not used in this script)
)

$bMoreDetails = $true

#-- Main program, do not change
$script:hMembers = @{}
$script:hGroupsFound = @{}  #-- This is used to detect loops; group1 is a member of group2, which is a member of group1
$script:rDlsToFixAccept = @()
$script:rDlsToFixReject = @()

function RecursivelyCheckDL
{
    param (
        [string] $sDL
    )
    #-- Show progress
    Write-Progress -Activity "Searching" -Status $sDL
    
    #-- Find the DL
    $dl = $null
    $dl = Get-DistributionGroup -Identity $sDL -ea SilentlyContinue
    if ($null -eq $dl)
    {
        Write-Host "DL not found: " -NoNewline
        Write-Host $sDL -ForegroundColor Red
        return
    }

    #-- Check for loop
    if ($script:hGroupsFound.ContainsKey($dl.Guid.ToString()))
    {
        if ($bMoreDetails) {Write-Host "Loop found, skipping" $dl.Identity.ToString() -ForegroundColor Yellow}
    }
    else
    {
        #-- Get members
        $script:hGroupsFound.Add($dl.Guid.ToString(), 1)
        [array]$rMembers = Get-DistributionGroupMember $dl.PrimarySmtpAddress -ResultSize unlimited
        
        #-- Show DL count info, warning if more than 500 members
        if ($rMembers.Count -ge 500) {Write-Host $dl.DisplayName $rMembers.Count -ForegroundColor Yellow}
        else {Write-Host $dl.DisplayName $rMembers.Count}
        
        #-- Process the accept list
        if ($dl.AcceptMessagesOnlyFromSendersOrMembers.Count -gt 0)
        {
            #-- Get more details on the accept list
            $rAcceptList = @()
            foreach ($sAcceptEntry in $dl.AcceptMessagesOnlyFromSendersOrMembers)
            {
                $rAcceptList += Get-EXORecipient $sAcceptEntry -ea SilentlyContinue
            }
            $rAcceptList | ForEach-Object {$_ | Add-Member -MemberType NoteProperty -Name 'Found' -Value $false -Force}

            #-- Loop and see if it matches our sender/keyword
            $bFound = $false
            foreach ($oAccept in $rAcceptList)
            {
                if ($oAccept.PrimarySmtpAddress.ToString() -match $script:sKeyword) {$oAccept.Found = $true}
                if ($oAccept.DisplayName -match $script:sKeyword) {$oAccept.Found = $true}
                if ($oAccept.Found) {$bFound = $true}
            }
            
            #-- Show our results
            $nCtr = 0
            foreach ($oAccept in $rAcceptList)
            {
                if ($nCtr -eq 0) {Write-Host "   Accept --> " -NoNewline} else {Write-Host "              " -NoNewline}
                $nCtr++
                if ($oAccept.Found) {Write-Host $oAccept.PrimarySmtpAddress.ToString() -ForegroundColor Green}
                else {Write-Host $oAccept.PrimarySmtpAddress.ToString()}
            }
            if (-not $bFound)
            {
                Write-Host "   Accept --> $($script:sKeyword) not found"  -ForegroundColor Red
                $script:rDlsToFixAccept += $dl
            }
        }
        else
        {
            if ($rMembers.Count -ge 500) {Write-Host "   Accept -->" -ForegroundColor Red}
            else {if ($bMoreDetails) {Write-Host "   Accept --> <none>"}}
        }
        
        #-- Process the reject list
        if ($dl.RejectMessagesFromSendersOrMembers.Count -gt 0)
        {
            #-- Get more details on the reject list
            $rRejectList = @()
            foreach ($sRejectEntry in $dl.RejectMessagesFromSendersOrMembers)
            {
                $rRejectList += Get-ExoRecipient $sRejectEntry -ea SilentlyContinue
            }
            $rRejectList | ForEach-Object {$_ | Add-Member -MemberType NoteProperty -Name 'Found' -Value $false -Force}

            #-- Loop and see if it matches our sender
            $bFound = $false
            foreach ($oReject in $rRejectList)
            {
                if ($oReject.PrimarySmtpAddress.ToString() -match $script:sKeyword) {$oReject.Found = $true}
                if ($oReject.DisplayName -match $script:sKeyword) {$oReject.Found = $true}
                if ($oReject.Found) {$bFound = $true}
            }
            
            #-- Show our results
            if ($bFound)
            {
                $nCtr = 0
                foreach ($oReject in $rRejectList)
                {
                    if ($nCtr -eq 0) {Write-Host "   Reject --> " -NoNewline} else {Write-Host "              " -NoNewline}
                    $nCtr++
                    if ($oReject.Found)
                    {
                        Write-Host $oReject.PrimarySmtpAddress.ToString() -ForegroundColor Red
                        $script:rDlsToFixReject += $dl
                    }
                    else {Write-Host $oReject.PrimarySmtpAddress.ToString()}
                }
            }
        }
        else {if ($bMoreDetails) {Write-Host "   Reject --> <none>"}}
        
        #-- Loop
        foreach($member in $rMembers)
        {
            #-- Recursively call if member is another group
            if ($member.RecipientType -like "*Group")
            {
                RecursivelyCheckDL($member.PrimarySmtpAddress.ToString())
            }
            else
            {
                $sKey = $member.Guid.ToString()
                if (-not $script:hMembers.ContainsKey($sKey)) {$script:hMembers.Add($sKey, 1)}
            }
        }
    }
    Write-Progress -Activity "Searching" -Completed
}

#------------------------------------------------------------------------------
#-- Main program
#------------------------------------------------------------------------------
$PSStyle.Progress.View = "Classic"  #-- Other value: "Minimal", only works in PowerShell 7.2+
$ProgressPreference = "Continue"
Write-Host ' '
Write-Host "This script will search for members of the specified Distribution Lists (DLs) and their sub-DLs."
Write-Host "and use the keyword (full or partial match of sender's email or display-name) to look for all"
Write-Host "the found DLs and show if the sender is in the accept or reject list of the DL."

#-- Normalize the parameters
$rDLs = @()
$DistributionLists -split ' ' | ForEach-Object {$rDLs += $_.Trim()}
$script:sKeyword = $Keyword.Trim().ToLower()

#-- Show command line usage
if ($DistributionLists.Trim().Length -eq 0 -or $rDLs.Count -eq 0 -or $script:sKeyword.Length -eq 0)
{
    Write-Host "Usage: " -NoNewline
    Write-Host ".\get-nested-dls-with-restrictions.ps1 -DistributionLists " -NoNewline -ForegroundColor Yellow
    Write-Host "dl1,dl2 " -NoNewline
    Write-Host "-Keyword " -NoNewline -ForegroundColor Yellow
    Write-Host "somekeyword"
    Write-Host ' '
    return
}

#-- Loop thru list of DLs
foreach ($sDL in $rDLs)
{
    $script:hMembers = @{}
    $script:hGroupsFound = @{}
    RecursivelyCheckDL($sDL)
    Write-Host "Total users:" $script:hMembers.Count
    Write-Host "Total groups:" $script:hGroupsFound.Count
    Write-Host ' '
}
#-- Show final results
if ($script:rDlsToFixAccept.Count -gt 0)
{
    Write-Host "DLs to add user to accept list:"
    foreach ($dl in $script:rDlsToFixAccept) {Write-Host "  " $dl.DisplayName}
}
if ($script:rDlsToFixReject.Count -gt 0)
{
    Write-Host "DLs to remove user from reject list:"
    foreach ($dl in $script:rDlsToFixReject) {Write-Host "  " $dl.DisplayName}
}
