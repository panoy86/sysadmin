#-- Script to check if DLs that are members of the top DL have sender restrictions

#-- Add/remove DLs from this list
$rDLs = @()
$rDLs += "dl-something1"
$rDLs += "dl-something2"
$sLookFor = "person-keyword-to-search"

#-- Main program, do not change
$global:nTotalDLCount = 0
$global:hMembers = @{}
$global:hGroupsFound = @{}  #-- This is used to detect loops; group1 is a member of group2, which is a member of group1

function udfGet-DLMembers([String] $sDL)
{
    #-- Show progress
    #Write-Host "   " $sDL $global:nTotalDLCount.ToString()
    
    #-- Find the DL
    $d = $null
    $d = Get-DistributionGroup -Identity $sDL -ea SilentlyContinue
    if ($d -ne $null)
    {
        #-- Check for loop
        if ($global:hGroupsFound.ContainsKey($d.Identity.ToString()))
        {
            Write-Host "   Loop found, skipping" $d.Identity.ToString() -ForegroundColor Yellow
        }
        else
        {
            #-- Get members
            $global:hGroupsFound.Add($d.Identity.ToString(), 1)
            [array]$r = Get-DistributionGroupMember $d.PrimarySmtpAddress -ResultSize unlimited
            
            #-- Show DL info
            if ($r.Count -ge 500) {Write-Host $d.DisplayName $r.Count -ForegroundColor Yellow}
            else {Write-Host $d.DisplayName $r.Count}
            
            #-- Check the accept
            if ($d.AcceptMessagesOnlyFromSendersOrMembers.Count -gt 0)
            {
                $rTmp = @()
                $d.AcceptMessagesOnlyFromSendersOrMembers | foreach {$rTmp += Get-Recipient $_ -ea SilentlyContinue}
                $sTmp = ($rTmp | foreach {$_.PrimarySmtpAddress}) -join ','
                #Write-Host "   Accept -->" $sTmp
                $rTmp2 = ($sTmp -split ',') | sort
                for($i = 0; $i -lt $rTmp2.Count; $i++)
                {
                    if ($i -eq 0) {Write-Host "   Accept -->" $rTmp2[$i]}
                    else {Write-Host "             " $rTmp2[$i]}
                }
                if ($sTmp -match $sLookFor) {Write-Host "   Accept -->" $sLookFor -ForegroundColor Green}
                else {Write-Host "   Accept -->" $sLookFor "NOT FOUND" -ForegroundColor Red}
            }
            else
            {
                if ($r.Count -ge 500) {Write-Host "   Accept -->" -ForegroundColor Red}
                else {Write-Host "   Accept -->"}
            }
            
            #-- Check the reject
            if ($d.RejectMessagesFromSendersOrMembers.Count -gt 0)
            {
                $rTmp = @()
                $d.RejectMessagesFromSendersOrMembers | foreach {$rTmp += Get-Recipient $_}
                Write-Host "   Reject -->" (($rTmp | foreach {$_.PrimarySmtpAddress}) -join ', ')
            }
            else {Write-Host "   Reject -->"}
            #Write-Host "  A-->" $d.AcceptMessagesOnlyFromSendersOrMembers
            #Write-Host "  R-->" $d.RejectMessagesFromSendersOrMembers
            Write-Host ' '
            #-- Loop
            foreach($t in $r)
            {
                #-- Recursively call if member is another group
                if ($t.RecipientType -like "*Group")
                {
                    $global:nTotalDLCount++
                    udfGet-DLMembers($t.PrimarySmtpAddress.ToString())
                }
                else
                {
                    $sKey = $t.PrimarySmtpAddress
                    if (-not $global:hMembers.ContainsKey($sKey)) {$global:hMembers.Add($sKey, 1)}
                }
            }
        }
    }
}


#-- Loop thru list of DLs
for($i=0; $i -lt $rDLs.Count; $i++)
{
    $sDL = $rDLs[ $i ]
    $global:hMembers = @{}
    $global:hGroupsFound = @{}
    udfGet-DLMembers($sDL)
    Write-Host "Total users:" $global:hMembers.Count
    Write-Host "Total groups:" $global:hGroupsFound.Count
    Write-Host ' '
}
