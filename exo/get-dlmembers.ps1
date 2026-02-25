<#
.DESCRIPTION
    This script retrieves all members of a specified Distribution List (DL) in Microsoft 365, including members of any nested DLs. It can also search for a specific keyword within the members' email addresses and display the DLs they belong to.
.PARAMETER ListOfDLs
    A comma-separated list of Distribution Lists (DLs, or Dynamic DLs) to search for members.
.PARAMETER Keyword
    An optional keyword to match against members' email addresses. If provided, the script will display the email addresses that match the keyword along with the DLs they belong to.
.EXAMPLE
    .\get-dlmembers.ps1 -ListOfDLs "DL1,DL2" -Keyword "john"
    This command retrieves all members of DL1 and DL2, including nested DLs, and displays any members whose email addresses contain the keyword "john" along with the DLs they belong to.
.NOTES
    - The script uses recursion to find members of nested DLs.
    - The output includes the total number of unique members found and the total number of DLs processed.
    - The script generates a file named "emails-list.txt" containing the email addresses of all members found.
#>

param (
    [string] $ListOfDLs, #-- Comma-separated list of DL to get the members list
    [string] $Keyword    #-- Optional keyword to match members against
)

#-- Main program, do not change
$script:HashMembersEmail = @{} # Unique list of members with email as key and the member object as value
$script:TotalDLCount = 0
$script:HashGroupsFound = @{}  # This is used to detect loops
#$script:aSubDLs = @()
$script:MemberKeywordMatch = $null

#---------------------------------------------------------------------------------
#-- Recursive function for DL search
#---------------------------------------------------------------------------------
function Get-DLMembersRecursive
{
    param ([String] $DLIdentity)
    
    # Find the DL
    $dl = $null
    $dl = Get-DistributionGroup -Identity $DLIdentity -ea SilentlyContinue
    if ($null -ne $dl) {
        # Check for loop
        if ($script:HashGroupsFound.ContainsKey($dl.Guid.ToString())) {
            Write-Host ($dl.Name.ToString() + " (loop found)") -ForegroundColor Yellow
        }
        else {
            # Get the members
            $script:HashGroupsFound.Add($dl.Guid.ToString(), 1)
            [array]$dlMembers = Get-DistributionGroupMember $dl.Guid.ToString() -ResultSize unlimited
            
            # Show progress
            Write-Host ("-> " + $dl.Name + " " + $dlMembers.Count.ToString())

            # Loop thru members
            if ($dlMembers.Count -gt 0) {
                # Track the sub-DLs
                #$script:aSubDLs += ($dlMembers.Count.ToString() + ", " + $dl.Name)
                Write-Progress -Activity "Finding member DLs" -Status ($dlMembers.Count.ToString() + ", " + $dl.Name)
                foreach($member in $dlMembers) {
                    # Recursively call if member is another group
                    if ($member.RecipientType -like "*Group") {
                        $script:TotalDLCount++
                        Get-DLMembersRecursive -DLIdentity $member.PrimarySmtpAddress.ToString()
                    }
                    #-- Add user-mailbox/mail-contact to list
                    else {
                        $email = $member.PrimarySmtpAddress.ToString().ToLower().Trim()
                        if (-not $script:HashMembersEmail.ContainsKey($email)) {
                            $script:HashMembersEmail.Add($email, $dl.Name)
                        }
                        if (-not $script:HashMembersEmail.ContainsKey($email)) {
                            $script:HashMembersEmail.Add($email, $member)
                        }

						#-- Check for the member keyword match
						if ($null -ne $script:MemberKeywordMatch) {
							if ($email -match $script:MemberKeywordMatch) {
								Write-Host "   match found: " -ForegroundColor Green -NoNewline
                            	Write-Host $email "->" $DLIdentity
							}
						}
                    }
                }
            }
        }
    }
    # If the DL is not found, check if it's a Dynamic Distribution List (DDL)
    else {
        $ddl = $null
        $ddl = Get-DynamicDistributionGroup -Identity $DLIdentity -ea SilentlyContinue
        if ($null -ne $ddl) {
            # Show progress
            $dlMembers = Get-Recipient -RecipientPreviewFilter $ddl.RecipientFilter -ResultSize unlimited
            Write-Host ("-> " + $ddl.Name + " (dynamic) " + $dlMembers.Count.ToString())
            foreach($member in $dlMembers) {
                $email = $member.PrimarySmtpAddress.ToString().ToLower().Trim()
                if (-not $script:HashMembersEmail.ContainsKey($email)) {
                    $script:HashMembersEmail.Add($email, $ddl.Name)
                }
                if (-not $script:HashMembersEmail.ContainsKey($email)) {
                    $script:HashMembersEmail.Add($email, $member)
                }
            
                # Check for the member keyword match
                if ($null -ne $script:MemberKeywordMatch) {
                    if ($email -match $script:MemberKeywordMatch) {
                        Write-Host "   match found: " -ForegroundColor Green -NoNewline
                        Write-Host $email "->" $DLIdentity #"->" $script:MemberKeywordMatch
                    }
                }
            }
        }
        else {Write-Host ($DLIdentity + " not found") -ForegroundColor Red}
    }
    Write-Progress -Activity "Finding member DLs" -Completed
}

#------------------------------------------------------------------------------
#-- Main
#------------------------------------------------------------------------------
if ([int]$PSVersionTable.PSVersion.Major -ge 7) {
    Import-Module PSReadLine -Force     # Fixes progress bar issues in PowerShell 7+
    $PSStyle.Progress.View = "Minimal"  # Other value: "Minimal", only works in PowerShell 7.2+
}
$ProgressPreference = "Continue"

Write-Host "This script will search for members of the specified Distribution Lists (DLs) and their sub-DLs."
if ($null -ne $ListOfDLs) {
    $listDls = $ListOfDLs.Split(" ")
}
else {
    # Get input of DLs from the user
    Write-Host "Enter the name of the Distribution Lists (DLs) to search, separated by commas:"
    $inputHost = Read-Host
    $listDls = $inputHost.Split(",")
}

# Get the optional member/keyword to search. If supplied, this will look for members that match this keyword
# and display the DL that this user is a member of. Handy for figuring out where a particular user is a
# member of when it involves hundreds of DLs to search.
if ($Keyword -ne $null -and $Keyword.Trim().Length -gt 0) {
	$script:MemberKeywordMatch = $Keyword
	Write-Host "Member keyword match: " -NoNewline
	Write-Host $script:MemberKeywordMatch -ForegroundColor Green
}
else {
    $script:MemberKeywordMatch = $null
}

# Loop thru the DLs
foreach ($DLIdentity in $listDls) {
    $script:hMembers = @{}
    Get-DLMembersRecursive -DLIdentity $DLIdentity.Trim()
    Write-Host ($DLIdentity  + " (" + $script:hMembers.Count + ")")
}
Write-Host "Total users:" $script:HashMembersEmail.Count
Write-Host "Total DLs (duplicates allowed):" $script:TotalDLCount

# Optional, get the SMTP list only
$rFinal = @()
$rTmp = $script:HashMembersEmail.Get_Keys()
foreach($t in $rTmp) {
    if ($script:HashMembersEmail.ContainsKey($t)) {$rFinal += $script:HashMembersEmail[$t]}
}
$rTmp = @()
$rFinal | ForEach-Object {$rTmp += $_.PrimarySmtpAddress}
$rTmp | Out-File -Encoding ASCII .\emails-list.txt
Write-Host "File created: emails-list.txt"
Write-Host "Members:" $rFinal.Count
return $rFinal.Count
