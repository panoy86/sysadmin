#-- Script to extract all members of a DL, and members of sub-DLS, etc...
#-- Main program, do not change
$script:hTotalMembers = @{}
$script:nTotalDLCount = 0
$script:hGroupsFound = @{}  #-- This is used to detect loops
$script:rSubDLs = @()

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
            Write-Host ($oDL.Name.ToString() + " (loop found)") -ForegroundColor Green
        }
        else
        {
            #-- Get the members
            $script:hGroupsFound.Add($oDL.Guid.ToString(), 1)
            [array]$rMembers = Get-DistributionGroupMember $oDL.Guid.ToString() -ResultSize unlimited
            
            #-- Show progress
            Write-Host ($oDL.Name + ' ' + $rMembers.Count.ToString())

            #-- Loop thru members
            if ($rMembers.Count -gt 0)
            {
                #-- Track the sub-DLs
                $script:rSubDLs += ($rMembers.Count.ToString() + ", " + $oDL.Name)
                Write-Progress -Activity "Finding member DLs" -Status ($rMembers.Count.ToString() + ", " + $oDL.Name)
                foreach($oMember in $rMembers)
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
                            Write-Host "   duplicate: " -ForegroundColor Yellow -NoNewline
                            Write-Host $sEmail $script:hMembers[$sEmail]
                        }
                        if (-not $script:hTotalMembers.ContainsKey($sEmail))
                        {$script:hTotalMembers.Add($sEmail, $oMember)}
                    }
                }
            }
        }
    }
    else {Write-Host ($sDL + " not found") -ForegroundColor Red}
    Write-Progress -Activity "Finding member DLs" -Completed
}

#------------------------------------------------------------------------------
#-- Main
#------------------------------------------------------------------------------
$PSStyle.Progress.View = "Minimal"  #-- Other value: "Classic", only works in PowerShell 7.2+
$ProgressPreference = "Continue"

Write-Host "This script will search for members of the specified Distribution Lists (DLs) and their sub-DLs."
#-- Get input of DLs from the user
if ($args.Count -gt 0) {$rDLs = $args}
else
{
    Write-Host "Enter the name of the Distribution Lists (DLs) to search, separated by commas:"
    $sInput = Read-Host
    $rDLs = $sInput.Split(",")
}

#-- Loop thru the DLs
foreach ($sDL in $rDLs)
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
