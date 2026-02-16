<#
.SYNOPSIS
    Get nested Distribution Lists with restrictions
.DESCRIPTION
    This script will search for members of the specified Distribution Lists (DLs) and their sub-DLs.
    and use the keyword (full or partial match of sender's email or display-name) to look for all
    the found DLs and show if the sender is in the accept or reject list of the DL.
.PARAMETER DistributionLists
    Comma-separated list of Distribution Lists to search.
.PARAMETER Keyword
    Keyword to search in DL names (not used in this script).
.EXAMPLE
    .\get-nested-dls-with-restrictions.ps1 -DistributionLists "dl1,dl2" -Keyword "user1"
    This will search for members of dl1 and dl2, and all their sub-DLs, and check if "user1" is in the accept or reject list of any of the found DLs.
.NOTES
    Assumes that an existing PowerShell session to Exchange/Online is already set.
#>

param (
    [string] $DistributionLists,  #-- Comma-separated list of Distribution Lists to search
    [string] $Keyword             #-- Keyword to search in DL names (not used in this script)
)

$Script:MoreDetails = $true

# Main program, do not change
$Script:HashMembers = @{}
$Script:HashGroupsFound = @{}  # This is used to detect loops; group1 is a member of group2, which is a member of group1
$Script:DlsToFixAccept = @()
$Script:DlsToFixReject = @()

function RecursivelyCheckDL
{
    param (
        [string] $DLName
    )
    # Show progress
    Write-Progress -Activity "Searching" -Status $DLName
    
    # Find the DL
    $isDynamicDL = $false
    $dl = $null
    $dl = Get-DistributionGroup -Identity $DLName -ea SilentlyContinue
    if ($null -eq $dl) {$dl = Get-DynamicDistributionGroup -Identity $DLName -ea SilentlyContinue; if ($null -ne $dl) {$isDynamicDL = $true}}
    if ($null -eq $dl)
    {
        Write-Host "DL not found: " -NoNewline
        Write-Host $DLName -ForegroundColor Red
        return
    }

    # Check for DL loops (e.g., group1 is a member of group2, which is a member of group1); if loop found, skip processing this DL since we already processed it when we hit it the first time
    if ($Script:HashGroupsFound.ContainsKey($dl.Guid.ToString()))
    {
        if ($Script:MoreDetails) {Write-Host "Loop found, skipping" $dl.Identity.ToString() -ForegroundColor Yellow}
    }
    else
    {
        # Get members, if dynamic - just set to zero members since we don't need to expand it; we just want to know if the sender is in the accept/reject list of the DL itself, not worry about the members of the dynamic DL
        $Script:HashGroupsFound.Add($dl.Guid.ToString(), 1)
        if (-not $isDynamicDL) {[array]$members = Get-DistributionGroupMember $dl.PrimarySmtpAddress -ResultSize unlimited}
        else {[array]$members = @()}
        
        # Show DL count info, warning if more than 500 members
        if ($members.Count -ge 500)
        {
            Write-Host ($dl.DisplayName + " (" + $dl.PrimarySmtpAddress.ToString() + ") ") -NoNewline
            Write-Host $members.Count -ForegroundColor Yellow
        }
        else
        {
            if ($isDynamicDL) {Write-Host ($dl.DisplayName + " (" + $dl.PrimarySmtpAddress.ToString() + ") <Dynamic DL>")} 
            else {Write-Host ($dl.DisplayName + " (" + $dl.PrimarySmtpAddress.ToString() + ") " + $members.Count)}
        }
        
        # Process the accept list
        if ($dl.AcceptMessagesOnlyFromSendersOrMembers.Count -gt 0)
        {
            # Get more details on the accept list
            $acceptList = @()
            foreach ($sAcceptEntry in $dl.AcceptMessagesOnlyFromSendersOrMembers)
            {
                $acceptList += Get-EXORecipient $sAcceptEntry -ea SilentlyContinue
            }
            $acceptList | ForEach-Object {$_ | Add-Member -MemberType NoteProperty -Name 'Found' -Value $false -Force}

            # Loop and see if it matches our sender/keyword
            $isFound = $false
            foreach ($acceptEntry in $acceptList)
            {
                if ($acceptEntry.PrimarySmtpAddress.ToString() -match $Script:KeywordToMatch) {$acceptEntry.Found = $true}
                if ($acceptEntry.DisplayName -match $Script:KeywordToMatch) {$acceptEntry.Found = $true}
                if ($acceptEntry.Found) {$isFound = $true}
            }
            
            # Show our results
            $counter = 0
            foreach ($acceptEntry in $acceptList)
            {
                if ($counter -eq 0) {Write-Host "   Accept --> " -NoNewline} else {Write-Host "              " -NoNewline}
                $counter++
                if ($acceptEntry.Found) {Write-Host $acceptEntry.PrimarySmtpAddress.ToString() -ForegroundColor Green}
                else {Write-Host $acceptEntry.PrimarySmtpAddress.ToString()}
            }
            if (-not $isFound)
            {
                Write-Host "   Accept --> $($Script:KeywordToMatch) not found"  -ForegroundColor Red
                $Script:DlsToFixAccept += $dl
            }
        }
        else
        {
            if ($members.Count -ge 500) {Write-Host "   Accept -->" -ForegroundColor Red}
            else {if ($Script:MoreDetails) {Write-Host "   Accept --> <none>"}}
        }
        
        # Process the reject list
        if ($dl.RejectMessagesFromSendersOrMembers.Count -gt 0)
        {
            # Get more details on the reject list
            $rejectList = @()
            foreach ($sRejectEntry in $dl.RejectMessagesFromSendersOrMembers)
            {
                $rejectList += Get-ExoRecipient $sRejectEntry -ea SilentlyContinue
            }
            $rejectList | ForEach-Object {$_ | Add-Member -MemberType NoteProperty -Name 'Found' -Value $false -Force}

            # Loop and see if it matches our sender
            $isFound = $false
            foreach ($rejectEntry in $rejectList)
            {
                if ($rejectEntry.PrimarySmtpAddress.ToString() -match $Script:KeywordToMatch) {$rejectEntry.Found = $true}
                if ($rejectEntry.DisplayName -match $Script:KeywordToMatch) {$rejectEntry.Found = $true}
                if ($rejectEntry.Found) {$isFound = $true}
            }
            
            # Show our results
            if ($isFound)
            {
                $counter = 0
                foreach ($rejectEntry in $rejectList)
                {
                    if ($counter -eq 0) {Write-Host "   Reject --> " -NoNewline} else {Write-Host "              " -NoNewline}
                    $counter++
                    if ($rejectEntry.Found)
                    {
                        Write-Host $rejectEntry.PrimarySmtpAddress.ToString() -ForegroundColor Red
                        $Script:DlsToFixReject += $dl
                    }
                    else {Write-Host $rejectEntry.PrimarySmtpAddress.ToString()}
                }
            }
        }
        else {if ($Script:MoreDetails) {Write-Host "   Reject --> <none>"}}
        
        # Loop thru all members and recursively call if it's a DL; if it's a user, add to our hash of members
        foreach($member in $members)
        {
            # Recursively call if member is another group
            if ($member.RecipientType -like "*Group")
            {
                RecursivelyCheckDL -DLName ($member.PrimarySmtpAddress.ToString())
            }
            else
            {
                $sKey = $member.Guid.ToString()
                if (-not $Script:HashMembers.ContainsKey($sKey)) {$Script:HashMembers.Add($sKey, 1)}
            }
        }
    }
    Write-Progress -Activity "Searching" -Completed
}

#------------------------------------------------------------------------------
#-- Main program
#------------------------------------------------------------------------------
if ([int]$PSVersionTable.PSVersion.Major -ge 7)
{
    Import-Module PSReadLine -Force     # Fixes progress bar issues in PowerShell 7+
    $PSStyle.Progress.View = "Minimal"  # Other value: "Minimal", only works in PowerShell 7.2+
}
$ProgressPreference = "Continue"

# Normalize the parameters
$listOfDLs = @()
$DistributionLists -split ' ' | ForEach-Object {$listOfDLs += $_.Trim()}
$Script:KeywordToMatch = $Keyword.Trim().ToLower()

# If the parameters are empty, show command line usage and exit
if ($DistributionLists.Trim().Length -eq 0 -or $listOfDLs.Count -eq 0 -or $Script:KeywordToMatch.Length -eq 0)
{
    Write-Host "Usage: " -NoNewline
    Write-Host ".\get-nested-dls-with-restrictions.ps1 -DistributionLists " -NoNewline -ForegroundColor Yellow
    Write-Host "dl1,dl2 " -NoNewline
    Write-Host "-Keyword " -NoNewline -ForegroundColor Yellow
    Write-Host "somekeyword"
    Write-Host ' '
    return
}

# Loop thru list of DLs
foreach ($itemDL in $listOfDLs)
{
    $Script:HashMembers = @{}
    $Script:HashGroupsFound = @{}
    RecursivelyCheckDL -DLName $itemDL
    Write-Host "Total users:" $Script:HashMembers.Count
    Write-Host "Total groups:" $Script:HashGroupsFound.Count
    Write-Host ' '
}
# Show final results
if ($Script:DlsToFixAccept.Count -gt 0)
{
    Write-Host "DLs to add user to accept list:"
    foreach ($dl in $Script:DlsToFixAccept) {Write-Host "  " + $dl.DisplayName -NoNewline; Write-Host (" (" + $dl.PrimarySmtpAddress.ToString() + ")") -ForegroundColor Green}
}
if ($Script:DlsToFixReject.Count -gt 0)
{
    Write-Host "DLs to remove user from reject list:"
    foreach ($dl in $Script:DlsToFixReject) {Write-Host "  " - $dl.DisplayName -NoNewline; Write-Host (" (" + $dl.PrimarySmtpAddress.ToString() + ")") -ForegroundColor Red}
}