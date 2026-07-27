<#
.SYNOPSIS
    Get nested Distribution Lists with restrictions
.DESCRIPTION
    This script will search for members of the specified Distribution Lists (DLs) and their sub-DLs,
        and use the sender-keywords (full or partial match of sender's email or display-name) to look for all
        the found DLs and show if the sender(s) is in the accept (or reject list) of the DL.
.PARAMETER DistributionLists
    Comma-separated list of Distribution Lists to search.
.PARAMETER Keywords
    Comma-separated list of sender-keywords to search in DLs' accept/reject lists.
.EXAMPLE
    .\get-nested-dls-with-restrictions.ps1 -DistributionLists "dl1,dl2" -Keywords "user1,user2"
    This will search for members of dl1 and dl2, and all their sub-DLs, and check if "user1" or "user2" is
        in the accept list of any of the found DLs.
.NOTES
    Assumes that an existing PowerShell session to Exchange/Online is already set. Future updates to check for reject list as well.
#>
param (
    [string] $DistributionLists, # Comma-separated list of Distribution Lists to search
    [string] $Keywords           # Comma-separated list of keywords to search in DLs' accept/reject lists
)

# Script-wide variables
$Script:ListMembers = @()
$Script:HashMembers = @{}
$Script:DlsToFixAccept = @()

#------------------------------------------------------------------------------
# Get all members of this group, can filter by member type (Non-Group, Group, All)
#------------------------------------------------------------------------------
function RecursivelyGetGroupMembers
{
    param (
        [string] $DLName,
        [string] $MemberType = 'All' # Parameter to filter members by type (e.g., Non-Group, Group, All), defaults to All
    )
    # Show progress
    Write-Progress -Activity "Expanding DL" -Status $DLName

    # Find the DL
    $isDynamicDL = $false
    $dl = $null
    $dl = Get-DistributionGroup -Identity $DLName -ea SilentlyContinue
    if ($null -eq $dl) {
        $dl = Get-DynamicDistributionGroup -Identity $DLName -ea SilentlyContinue
        if ($null -ne $dl) {
            $isDynamicDL = $true
        }
    }
    # Exit function if DL not found
    if ($null -eq $dl) {
        Write-ToDisplay ("DL not found: {}" + $DLName + "{Red}")
        return
    }

    # Get the members of this DL
    if ($isDynamicDL) {
        $members = Get-Recipient -RecipientPreviewFilter $dl.RecipientFilter -ResultSize Unlimited
    }
    else {
        $members = Get-DistributionGroupMember -Identity $dl.PrimarySmtpAddress -ResultSize Unlimited
    }

    # Only get the group members
    if ($MemberType -eq 'Group') {
        foreach ($member in $members) {
            # If this member is a group, add to our list and recursively call this function to get its members
            if ($member.RecipientType -match 'Group') {
                if (-not $Script:HashMembers.ContainsKey($member.Guid.ToString())) {
                    $Script:HashMembers.Add($member.Guid.ToString(), 1)
                    $Script:ListMembers += $member
                }
                RecursivelyGetGroupMembers -DLName $member.PrimarySmtpAddress -MemberType $MemberType
            }
        }
    }
    else {
        foreach ($member in $members) {
            # If this member is a group, recursively call this function to get its members
            if ($member.RecipientType -match 'Group') {
                if ($MemberType -eq 'All') {
                    # Add this group to our script-wide list/hash
                    if (-not $Script:HashMembers.ContainsKey($member.Guid.ToString())) {
                        $Script:HashMembers.Add($member.Guid.ToString(), 1)
                        $Script:ListMembers += $member
                    }
                }
                RecursivelyGetGroupMembers -DLName $member.PrimarySmtpAddress -MemberType $MemberType
            }
            else {
                # By default, this is the MemberType -eq 'Non-Group' case, so we add the member to our script-wide list/hash
                if (-not $Script:HashMembers.ContainsKey($member.Guid.ToString())) {
                    $Script:HashMembers.Add($member.Guid.ToString(), 1)
                    $Script:ListMembers += $member
                }
            }
        }
    }
    # End progress
    Write-Progress -Activity "Expanding DL" -Completed
}

#------------------------------------------------------------------------------
# Check the accept/reject lists of the DL for the keywords
#------------------------------------------------------------------------------
function CheckDLRestrictions
{
    param (
        [string] $DLName,
        [string[]] $Keywords
    )

    # Find the DL
    $isDynamicDL = $false
    $dl = $null
    $dl = Get-DistributionGroup -Identity $DLName -ea SilentlyContinue
    if ($null -eq $dl) {
        $dl = Get-DynamicDistributionGroup -Identity $DLName -ea SilentlyContinue
        if ($null -ne $dl) {
            $isDynamicDL = $true
        }
    }
    # Exit function if DL not found
    if ($null -eq $dl) {
        Write-Host "DL not found:" $DLName -ForegroundColor Red
        return
    }

    # Prepare the list of keywords to match
    $keywordList = @()
    foreach ($keyword in $Keywords) {
        $keywordList += New-Object PSObject -Property @{Keyword = $keyword.Trim().ToLower(); Found = $false}
    }

    # Process the accept-user list
    if ($dl.AcceptMessagesOnlyFromSendersOrMembers.Count -gt 0) {
        # Get more details on the accept list
        $acceptList = @()
        foreach ($acceptEntry in $dl.AcceptMessagesOnlyFromSendersOrMembers) {
            $acceptList += Get-EXORecipient $acceptEntry -ea SilentlyContinue
        }
        $acceptList | ForEach-Object {$_ | Add-Member -MemberType NoteProperty -Name 'Found' -Value $false -Force}
        # Loop and see if it matches our sender/keywords
        foreach ($acceptEntry in $acceptList) {
            foreach ($keyword in $keywordList) {
                $stringMatch = $keyword.Keyword
                # Check if sender-keyword matches PrimarySmtpAddress
                if ($acceptEntry.PrimarySmtpAddress.ToString() -match $stringMatch) {
                    $acceptEntry.Found = $true
                    $keyword.Found = $true
                }
                # Check if sender-keyword matches DisplayName
                if ($acceptEntry.DisplayName -match $stringMatch) {
                    $acceptEntry.Found = $true
                    $keyword.Found = $true
                }
            }
        }
        # Show our results
        Write-Host "   Accept"
        foreach ($acceptEntry in ($acceptList | Sort-Object PrimarySmtpAddress)) {
            if ($acceptEntry.Found) {
                Write-Host "     " $acceptEntry.PrimarySmtpAddress.ToString() $acceptEntry.RecipientType -ForegroundColor Green
            }
            else {
                Write-Host "     " $acceptEntry.PrimarySmtpAddress.ToString()
            }
        }
    }
    # If the accept list contains groups, expand those as well and re-check
    if ($dl.AcceptMessagesOnlyFromDLMembers.Count -gt 0) {
        Write-Progress -Activity "Checking Accept-Groups" -Status $dl.DisplayName
        foreach ($acceptEntry in $dl.AcceptMessagesOnlyFromDLMembers) {
            # Reset our script-wide variables
            $Script:ListMembers = @()
            $Script:HashMembers = @{}
            # Get all members of this DL, including nested groups
            RecursivelyGetGroupMembers -DLName $acceptEntry -MemberType 'All'
            $Script:ListMembers | ForEach-Object {$_ | Add-Member -MemberType NoteProperty -Name 'Found' -Value $false -Force}
            $Script:ListMembers | ForEach-Object {$_ | Add-Member -MemberType NoteProperty -Name 'ParentDL' -Value '' -Force}
            foreach ($member in $Script:ListMembers) {
                foreach ($keyword in $keywordList) {
                    $stringMatch = $keyword.Keyword
                    # Check if sender-keyword matches PrimarySmtpAddress
                    if ($member.PrimarySmtpAddress.ToString() -match $stringMatch) {
                        $member.Found = $true
                        $member.ParentDL = $acceptEntry
                        $keyword.Found = $true
                    }
                    # Check if sender-keyword matches DisplayName
                    if ($member.DisplayName -match $stringMatch) {
                        $member.Found = $true
                        $member.ParentDL = $acceptEntry
                        $keyword.Found = $true
                    }
                }
            }
            Write-Progress -Activity "Checking Accept-Groups" -Completed
        }
        # Show our results
        Write-Host "   Accept-Groups"
        foreach ($member in ($Script:ListMembers | Sort-Object PrimarySmtpAddress)) {
            if ($member.Found) {
                Write-Host "     " $member.PrimarySmtpAddress.ToString() $member.RecipientType -ForegroundColor Green -NoNewline
                Write-Host (" (from DL: " + $member.ParentDL + ")")
            }
        }
    }
    # If accept list is empty, show it as well
    if ($dl.AcceptMessagesOnlyFromSendersOrMembers.Count -eq 0) {
        Write-Host "   Accept list is empty"
    }
    else {
        # Mark any keywords that were not found in the accept list
        foreach ($keyword in $keywordList) {
            if (-not $keyword.Found) {
                Write-Host ("   Accept -> " + $keyword.Keyword + " not found")
                $Script:DlsToFixAccept += [PSCustomObject]@{
                    dl = $dl.PrimarySmtpAddress.ToString()
                    sender = $keyword.Keyword
                }
            }
        }
    }
}

#------------------------------------------------------------------------------
# Main program
#------------------------------------------------------------------------------
$startTime = Get-Date
if ([int]$PSVersionTable.PSVersion.Major -ge 7) {
    Import-Module PSReadLine -Force     # Fixes progress bar issues in PowerShell 7+
    $PSStyle.Progress.View = "Minimal"  # Other value: "Minimal", only works in PowerShell 7.2+
}
$ProgressPreference = "Continue"

# Normalize the parameters
$ListOfDLs = @()
$ListOfKeywords = @()
$DistributionLists -split ' ' | ForEach-Object {$ListOfDLs += $_.Trim()}
$Keywords -split ' ' | ForEach-Object {$ListOfKeywords += $_.Trim()}

# If the parameters are empty, show command line usage and exit
if ($DistributionLists.Trim().Length -eq 0 -or $ListOfDLs.Count -eq 0 -or $ListOfKeywords.Count -eq 0 -or $Keywords.Length -eq 0) {
    Write-Host "Usage: " -NoNewline
    Write-Host ".\get-nested-dls-with-restrictions.ps1 -DistributionLists " -NoNewline -ForegroundColor Yellow
    Write-Host "dl1,dl2 " -NoNewline
    Write-Host "-Keywords " -NoNewline -ForegroundColor Yellow
    Write-Host "sender1,sender2,sender3"
    Write-Host ' '
    return
}

# Start the actual work
foreach ($itemDL in $ListOfDLs) {
    # Check for the top DL
    Write-Host "Checking DL: $itemDL" -ForegroundColor Cyan
    CheckDLRestrictions -DLName $itemDL -Keywords $ListOfKeywords

    # Get all group-members of this DL, including nested groups, then check those as well
    # Reset our script-wide variables
    Write-Host "Looking for nested DLs..." -ForegroundColor Cyan
    $Script:ListMembers = @()
    $Script:HashMembers = @{}
    RecursivelyGetGroupMembers -DLName $itemDL -MemberType 'Group'
    $nestedGroups = $Script:ListMembers
    foreach ($nestedGroup in $nestedGroups) {
        Write-Host "Checking nested DL: $($nestedGroup.PrimarySmtpAddress)" -ForegroundColor Cyan
        CheckDLRestrictions -DLName $nestedGroup.PrimarySmtpAddress -Keywords $ListOfKeywords
    }
}

# Show summary of results
if ($Script:DlsToFixAccept.Count -gt 0) {
    Write-Host "DLs to add sender to accept list:" -ForegroundColor Cyan
    $Script:DlsToFixAccept | Format-Table -AutoSize
}
else {
    Write-Host "No DLs found that need to add sender to accept list." -ForegroundColor Green
}
Write-Host ("Run time: " + ((Get-Date) - $startTime).ToString())

<#
#>
