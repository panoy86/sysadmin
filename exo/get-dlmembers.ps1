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
$script:hTotalMembers = @{}
$script:nTotalDLCount = 0
$script:hGroupsFound = @{}  #-- This is used to detect loops
$script:aSubDLs = @()
$script:sMemberMatch = $null

#---------------------------------------------------------------------------------
#-- Recursive function for DL search
#---------------------------------------------------------------------------------
function Get-DLMembersRecursive
{
    param([String] $sDL)
    
    #-- Find the DL
    $oDL = $null
    $oDL = Get-DistributionGroup -Identity $sDL -ea SilentlyContinue
    if ($null -ne $oDL)
    {
        #-- Check for loop
        if ($script:hGroupsFound.ContainsKey($oDL.Guid.ToString()))
        {
            Write-Host ($oDL.Name.ToString() + " (loop found)") -ForegroundColor Yellow
        }
        else
        {
            #-- Get the members
            $script:hGroupsFound.Add($oDL.Guid.ToString(), 1)
            [array]$aMembers = Get-DistributionGroupMember $oDL.Guid.ToString() -ResultSize unlimited
            
            #-- Show progress
            Write-Host ("-> " + $oDL.Name + " " + $aMembers.Count.ToString())

            #-- Loop thru members
            if ($aMembers.Count -gt 0)
            {
                #-- Track the sub-DLs
                $script:aSubDLs += ($aMembers.Count.ToString() + ", " + $oDL.Name)
                Write-Progress -Activity "Finding member DLs" -Status ($aMembers.Count.ToString() + ", " + $oDL.Name)
                foreach($oMember in $aMembers)
                {
                    #-- Recursively call if member is another group
                    if ($oMember.RecipientType -like "*Group")
                    {
                        $script:nTotalDLCount++
                        Get-DLMembersRecursive($oMember.PrimarySmtpAddress.ToString())
                    }
                    #-- Add user-mailbox/mail-contact to list
                    else
                    {
                        $sEmail = $oMember.PrimarySmtpAddress.ToString().ToLower().Trim()
                        if (-not $script:hMembers.ContainsKey($sEmail)) {$script:hMembers.Add($sEmail, $oDL.Name)}
                        else
                        {
                            #Write-Host "   duplicate: " -ForegroundColor Yellow -NoNewline
                            #Write-Host $sEmail $script:hMembers[$sEmail]
                        }
                        if (-not $script:hTotalMembers.ContainsKey($sEmail))
                        {$script:hTotalMembers.Add($sEmail, $oMember)}

						#-- Check for the member keyword match
						if ($null -ne $script:sMemberMatch)
						{
							#if ($script:sMemberMatch -match $sEmail)
							if ($sEmail -match $script:sMemberMatch)
							{
								Write-Host "   match found: " -ForegroundColor Green -NoNewline
                            	Write-Host $sEmail "->" $sDL #"->" $script:sMemberMatch
							}
						}
                    }
                }
            }
        }
    }
    #-- If the DL is not found, check if it's a Dynamic Distribution List (DDL)
    else
    {
        $oDDL = $null
        $oDDL = Get-DynamicDistributionGroup -Identity $sDL -ea SilentlyContinue
        if ($null -ne $oDDL)
        {
            #-- Show progress
            $aMembers = Get-Recipient -RecipientPreviewFilter $oDDL.RecipientFilter -ResultSize unlimited
            Write-Host ("-> " + $oDDL.Name + " (dynamic) " + $aMembers.Count.ToString())
            foreach($oMember in $aMembers)
            {
                $sEmail = $oMember.PrimarySmtpAddress.ToString().ToLower().Trim()
                if (-not $script:hMembers.ContainsKey($sEmail)) {$script:hMembers.Add($sEmail, $oDDL.Name)}
                else
                {
                    #Write-Host "   duplicate: " -ForegroundColor Yellow -NoNewline
                    #Write-Host $sEmail $script:hMembers[$sEmail]
                }
                if (-not $script:hTotalMembers.ContainsKey($sEmail))
                {$script:hTotalMembers.Add($sEmail, $oMember)}
            
                #-- Check for the member keyword match
                if ($null -ne $script:sMemberMatch)
                {
                    if ($sEmail -match $script:sMemberMatch)
                    {
                        Write-Host "   match found: " -ForegroundColor Green -NoNewline
                        Write-Host $sEmail "->" $sDL #"->" $script:sMemberMatch
                    }
                }
            }
        }
        else {Write-Host ($sDL + " not found") -ForegroundColor Red}
    }
    Write-Progress -Activity "Finding member DLs" -Completed
}

#------------------------------------------------------------------------------
#-- Main
#------------------------------------------------------------------------------
$PSStyle.Progress.View = "Minimal"  #-- Other value: "Classic", only works in PowerShell 7.2+
$ProgressPreference = "Continue"

Write-Host "This script will search for members of the specified Distribution Lists (DLs) and their sub-DLs."
#-- Get input of DLs from the user
if ($null -ne $ListOfDLs)
{
    #Write-Host $ListOfDLs
    $aDLs = $ListOfDLs.Split(" ")
}
else
{
    Write-Host "Enter the name of the Distribution Lists (DLs) to search, separated by commas:"
    $sInput = Read-Host
    $aDLs = $sInput.Split(",")
}

#-- Get the optional member/keyword to search. If supplied, this will look for members that match this keyword
#-- and display the DL that this user is a member of. Handy for figuring out where a particular user is a
#-- member of when it involves hundreds of DLs to search.
if ($Keyword -ne $null -and $Keyword.Trim().Length -gt 0)
{
	$script:sMemberMatch = $Keyword
	Write-Host "Member keyword match: " -NoNewline
	Write-Host $script:sMemberMatch -ForegroundColor Green
}
else {$script:sMemberMatch = $null}

#-- Loop thru the DLs
foreach ($sDL in $aDLs)
{
    $script:hMembers = @{}
    Get-DLMembersRecursive($sDL.Trim())
    Write-Host ($sDL  + " (" + $script:hMembers.Count + ")")
}
Write-Host "Total users:" $script:hTotalMembers.Count
Write-Host "Total DLs (duplicates allowed):" $script:nTotalDLCount

#-- Optional, get the SMTP list only
$rFinal = @()
$rTmp = $script:hTotalMembers.Get_Keys()
foreach($t in $rTmp)
{
    if ($script:hTotalMembers.ContainsKey($t)) {$rFinal += $script:hTotalMembers[$t]}
}
$rTmp = @()
$rFinal | ForEach-Object {$rTmp += $_.PrimarySmtpAddress}
$rTmp | Out-File -Encoding ASCII .\emails-list.txt
Write-Host "File created: emails-list.txt"
Write-Host "Members:" $rFinal.Count
return $rFinal.Count