<#
.SYNOPSIS
    Get nested Distribution Lists with restrictions
.DESCRIPTION
    This script will search for members of the specified Distribution Lists (DLs) and their sub-DLs,
        and use the sender-keywords (full or partial match of sender's email or display-name) to look for all
        the found DLs and show if the sender(s) is in the accept or reject list of the DL.
.PARAMETER DistributionLists
    Comma-separated list of Distribution Lists to search.
.PARAMETER Keywords
    Comma-separated list of sender-keywords to search in DLs' accept/reject lists.
.EXAMPLE
    .\get-nested-dls-with-restrictions.ps1 -DistributionLists "dl1,dl2" -Keywords "user1,user2"
    This will search for members of dl1 and dl2, and all their sub-DLs, and check if "user1" or "user2" is
        in the accept or reject list of any of the found DLs.
.NOTES
    Assumes that an existing PowerShell session to Exchange/Online is already set.
    If Azure automation/runbook is detected, it will assume a managed identity is set and attempt to
        authenticate to Exchange Online.
    Because a runbook does not have the same display capabilities as an interactive normal PowerShell session,
        a special function is created to handle the output differently for runbook vs local sessions, so that
        the output is still readable in a runbook.
#>

param (
    [string] $DistributionLists, # Comma-separated list of Distribution Lists to search
    [string] $Keywords           # Comma-separated list of keywords to search in DLs' accept/reject lists
)

$Script:MoreDetails = $true # Set to $true to show the full accept/reject list of each DL; set to $false to only show the matching entries in the accept/reject list of each DL

# Main program, do not change
$Script:HashMembers = @{}      # To keep track of all the unique members we found in our search, key is the member's guid
$Script:HashGroupsFound = @{}  # For detecting groups we already processed and skip those.
$Script:DlsToFixAccept = @()   # To keep track of DLs that we need to add the sender to the accept list
$Script:DlsToFixReject = @()   # To keep track of DLs that we need to remove the sender from the reject list
$Script:IsRunbook = $false     # To indicate if we are running in an Azure Automation runbook

#------------------------------------------------------------------------------
# Function to handle output differently for runbook vs local sessions, so that
# the output is still readable in a runbook, requires the use of curly brackets,
# hence curly brackets are excluded in the output.
# It expects a single string input with color tags in curly brackets,
# e.g. "This is a {Red}red{Green} and green{Yellow} message{Blue}"
#------------------------------------------------------------------------------
function Write-ToDisplay
{
    param(
        [string]$Message
    )
    # Remove color tags for runbook sessions
    if ($Script:IsRunbook) {   
        $Message = $Message -replace '\{.*?\}', ''
        Write-Output $Message
        return
    }
    # Else, process color tags for local sessions
    $listPhrases = $Message -split '}'
    foreach ($phrase in $listPhrases) {
        if ($phrase -match '\{') {
            $phrase += '}'
        }
        # Split the phrase into text and color components
        $text = $phrase -replace '\{.*?\}', ''
        $colorMatch = [regex]::Match($phrase, '\{(.*?)\}')
        if ($colorMatch.Success) {
            $color = $colorMatch.Groups[1].Value
            if ($color.Length -eq 0) {
                $color = "White"
            }
            Write-Host $text -ForegroundColor $color -NoNewline
        } else {
            Write-Host $text -NoNewline
        }
    }
    Write-Host "" # New line after processing all phrases
    return
}

#------------------------------------------------------------------------------
# Recursively search DLs for members and restrictions
#------------------------------------------------------------------------------
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

    # Check for DL loops (e.g., group1 is a member of group2, which is a member of group1);
    # if loop found, skip processing this DL since we already processed it when we hit it the first time
    if ($Script:HashGroupsFound.ContainsKey($dl.Guid.ToString())) {
        if ($Script:MoreDetails) {
            Write-ToDisplay ("already processed, skipping: {}" + $dl.Identity.ToString() + "{Yellow}")
        }
    }
    else {
        # Get members, if dynamic - just set to zero members since we don't need to expand it;
        # we just want to know if the sender is in the accept/reject list of the DL itself
        $Script:HashGroupsFound.Add($dl.Guid.ToString(), 1)
        if (-not $isDynamicDL) {
            [array]$members = Get-DistributionGroupMember $dl.PrimarySmtpAddress -ResultSize unlimited
        }
        else {
            [array]$members = @()
        }
        
        # Show DL count info, warning if more than 500 members
        if ($members.Count -ge 500) {
            Write-ToDisplay ($dl.DisplayName + "{Cyan} (" + $dl.PrimarySmtpAddress.ToString() + ") " + $members.Count + "{Red}")
        }
        else {
            if ($isDynamicDL) {
                Write-ToDisplay ($dl.DisplayName + "{Cyan} (" + $dl.PrimarySmtpAddress.ToString() + ") <Dynamic DL>")
            }
            else {
                Write-ToDisplay ($dl.DisplayName + "{Cyan} (" + $dl.PrimarySmtpAddress.ToString() + ") " + $members.Count)
            }
        }
        
        # Process the accept list
        if ($dl.AcceptMessagesOnlyFromSendersOrMembers.Count -gt 0) {
            # Get more details on the accept list
            $acceptList = @()
            foreach ($acceptEntry in $dl.AcceptMessagesOnlyFromSendersOrMembers) {
                $acceptList += Get-EXORecipient $acceptEntry -ea SilentlyContinue
            }
            $acceptList | ForEach-Object {$_ | Add-Member -MemberType NoteProperty -Name 'Found' -Value $false -Force}

            # Prepare the list of keywords to match
            $ListKeywords = @()
            foreach ($keyword in $Script:Keywords.Split(',')) {
                $ListKeywords += New-Object PSObject -Property @{Keyword = $keyword.Trim().ToLower(); Found = $false}
            }

            # Loop and see if it matches our sender/keywords
            foreach ($acceptEntry in $acceptList) {
                foreach ($keyword in $ListKeywords) {
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
            Write-ToDisplay ("   Accept")
            foreach ($acceptEntry in ($acceptList | Sort-Object PrimarySmtpAddress)) {
                if ($acceptEntry.Found) {
                    Write-ToDisplay ("      " + $acceptEntry.PrimarySmtpAddress.ToString() + " {Green}[matches sender]")
                }
                else {
                    Write-ToDisplay ("      " + $acceptEntry.PrimarySmtpAddress.ToString())
                }
            }
            foreach ($keyword in $ListKeywords) {
                if (-not $keyword.Found) {
                    Write-ToDisplay ("   Accept -> " + $keyword.Keyword + " not found")
                    $Script:DlsToFixAccept += [PSCustomObject]@{
                        dl = $dl
                        sender = $keyword
                    }
                }
            }
        }
        # DL has no accept-restrictions
        else {
            if ($Script:MoreDetails) {
                Write-ToDisplay "   Accept -> <none>"
            }
        }
        
        # Process the reject list
        if ($dl.RejectMessagesFromSendersOrMembers.Count -gt 0) {
            # Get more details on the reject list
            $rejectList = @()
            foreach ($rejectEntry in $dl.RejectMessagesFromSendersOrMembers) {
                $rejectList += Get-ExoRecipient $rejectEntry -ea SilentlyContinue
            }
            $rejectList | ForEach-Object {$_ | Add-Member -MemberType NoteProperty -Name 'Found' -Value $false -Force}

            # Prepare the list of keywords to match
            $ListKeywords = @()
            foreach ($keyword in $Script:Keywords.Split(',')) {
                $ListKeywords += New-Object PSObject -Property @{Keyword = $keyword.Trim().ToLower(); Found = $false}
            }
            
            # Loop and see if it matches our sender/keywords
            foreach ($rejectEntry in $rejectList) {
                foreach ($keyword in $ListKeywords) {
                    $stringMatch = $keyword.Keyword
                    # Check if sender-keyword matches PrimarySmtpAddress
                    if ($rejectEntry.PrimarySmtpAddress.ToString() -match $stringMatch) {
                        $rejectEntry.Found = $true
                        $keyword.Found = $true
                    }
                    # Check if sender-keyword matches DisplayName
                    if ($rejectEntry.DisplayName -match $stringMatch) {
                        $rejectEntry.Found = $true
                        $keyword.Found = $true
                    }
                }
            }

            # Show our results
            Write-ToDisplay ("   Reject")
            foreach ($rejectEntry in ($rejectList | Sort-Object PrimarySmtpAddress)) {
                if ($rejectEntry.Found) {
                    Write-ToDisplay ("      " + $rejectEntry.PrimarySmtpAddress.ToString() + " {Red}[matches sender]")
                }
                else {
                    Write-ToDisplay ("      " + $rejectEntry.PrimarySmtpAddress.ToString())
                }
            }
            foreach ($keyword in $ListKeywords) {
                if ($keyword.Found) {
                    Write-ToDisplay ("   Reject -> " + $keyword.Keyword + " found")
                    $Script:DlsToFixReject += [PSCustomObject]@{
                        dl = $dl
                        sender = $keyword
                    }
                }
            }
        }
        # DL has no reject-restrictions
        else {
            if ($Script:MoreDetails) {
                Write-ToDisplay "   Reject -> <none>"
            }
        }
        
        # Loop thru all members and recursively call if it's a DL; if it's a user, add to our hash of members
        foreach($member in $members) {
            # Recursively call if member is another group
            if ($member.RecipientType -like "*Group") {
                RecursivelyCheckDL -DLName ($member.PrimarySmtpAddress.ToString())
            }
            else {
                $sKey = $member.Guid.ToString()
                if (-not $Script:HashMembers.ContainsKey($sKey)) {$Script:HashMembers.Add($sKey, 1)}
            }
        }
    }
    Write-Progress -Activity "Searching" -Completed
}

#------------------------------------------------------------------------------
# Test if we are running in an Azure Automation runbook
#------------------------------------------------------------------------------
function Test-SessionInRunbook {
    return $env:AZUREPS_HOST_ENVIRONMENT -eq "AzureAutomation"
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
$listOfDLs = @()
$DistributionLists -split ',' | ForEach-Object {$listOfDLs += $_.Trim()}

# If the parameters are empty, show command line usage and exit
if ($DistributionLists.Trim().Length -eq 0 -or $listOfDLs.Count -eq 0 -or $Script:Keywords.Length -eq 0) {
    Write-Host "Usage: " -NoNewline
    Write-Host ".\get-nested-dls-with-restrictions.ps1 -DistributionLists " -NoNewline -ForegroundColor Yellow
    Write-Host "dl1,dl2 " -NoNewline
    Write-Host "-Keywords " -NoNewline -ForegroundColor Yellow
    Write-Host "sender1,sender2,sender3" -ForegroundColor Yellow
    Write-Host ' '
    return
}

# If we are in an Azure Automation runbook, attempt to authenticate to Exchange Online using a managed identity; if not in Azure Automation, assume we are already authenticated to Exchange Online.
if (Test-SessionInRunbook) {
    $Script:IsRunbook = $true
    try {
        Connect-ExchangeOnline -ManagedIdentity -Organization tionetworks.onmicrosoft.com
        Write-ToDisplay "Successfully authenticated to Exchange Online using managed identity.{Green}"
    }
    catch {
        Write-ToDisplay "Failed to authenticate to Exchange Online using managed identity. Please ensure the runbook has a system-managed identity configured."
        exit
    }
}

# Loop thru list of DLs, and show total members/groups found in all DLs
$Script:HashMembers = @{}
$Script:HashGroupsFound = @{}
foreach ($itemDL in $listOfDLs) {
    RecursivelyCheckDL -DLName $itemDL
}
Write-ToDisplay ("Total users: " + $Script:HashMembers.Count)
Write-ToDisplay ("Total groups: " + $Script:HashGroupsFound.Count)

# Show final results
if ($Script:DlsToFixAccept.Count -gt 0) {
    Write-ToDisplay "DLs to add user to accept list:"
    foreach ($dlFix in $Script:DlsToFixAccept) {
        Write-ToDisplay ("   +" + $dlFix.dl.DisplayName + "{} (" + $dlFix.dl.PrimarySmtpAddress.ToString() + `
            "){Green} " + $dlFix.sender.Keyword + "{}")
    }
}
if ($Script:DlsToFixReject.Count -gt 0) {
    Write-ToDisplay "DLs to remove user from reject list:"
    foreach ($dlFix in $Script:DlsToFixReject) {
        Write-ToDisplay ("   -" + $dlFix.dl.DisplayName + "{} (" + $dlFix.dl.PrimarySmtpAddress.ToString() +    
        "){Red} " + $dlFix.sender.Keyword + "{}")
    }
}
Write-ToDisplay ("Run time: {}" + ((Get-Date) - $startTime).ToString() + "{Yellow}")
