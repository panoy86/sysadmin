#-- Script to check if DLs that are members of the top DL have sender restrictions
#-- assumes that an existing PowerShell session to Exchange/Online is already set

#-- Add/remove DLs from this list
$rDLs = @()
$rDLs += "dl-something1"
$rDLs += "dl-something2"
$sSenderEmail = "user@someorg.com"
$bMoreDetails = $false


#-- Main program, do not change
$global:nTotalDLCount = 0
$global:hMembers = @{}
$global:hGroupsFound = @{}  #-- This is used to detect loops; group1 is a member of group2, which is a member of group1
$global:sSenderGuid = ''
$global:rDlsToFixAccept = @()
$global:rDlsToFixReject = @()


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
            if ($bMoreDetails)
            {Write-Host "   Loop found, skipping" $d.Identity.ToString() -ForegroundColor Yellow}
        }
        else
        {
            #-- Get members
            $global:hGroupsFound.Add($d.Identity.ToString(), 1)
            [array]$rMembers = Get-DistributionGroupMember $d.PrimarySmtpAddress -ResultSize unlimited
            
            #-- Show DL info
            if ($rMembers.Count -ge 500) {Write-Host $d.DisplayName $r.Count -ForegroundColor Yellow}
            else {Write-Host $d.DisplayName $rMembers.Count}
            
            #-- Process the accept list
            if ($d.AcceptMessagesOnlyFromSendersOrMembers.Count -gt 0)
            {
                #-- Get more details on the accept list
                $rAcceptList = @()
                foreach ($sAcceptEntry in $d.AcceptMessagesOnlyFromSendersOrMembers)
                {
                    $rAcceptList += Get-Recipient $sAcceptEntry -ea SilentlyContinue
                }
                
                #-- Loop and see if it matches our sender
                $bFound = $false
                foreach ($oAccept in $rAcceptList)
                {
                    if ($oAccept.Guid -eq $global:sSenderGuid) {$bFound = $true}
                }
                
                #-- Show our results
                $nCtr = 0
                foreach ($oAccept in $rAcceptList)
                {
                    if ($nCtr -eq 0) {Write-Host "   Accept --> " -NoNewline} else {Write-Host "              " -NoNewline}
                    $nCtr++
                    if ($oAccept.Guid -eq $global:sSenderGuid)
                    {Write-Host $oAccept.PrimarySmtpAddress.ToString() -ForegroundColor Green}
                    else {Write-Host $oAccept.PrimarySmtpAddress.ToString()}
                }
                if (-not $bFound)
                {
                    Write-Host "   Accept -->" $sSenderEmail -ForegroundColor Red
                    $global:rDlsToFixAccept += $d
                }
            }
            else
            {
                if ($rMembers.Count -ge 500) {Write-Host "   Accept -->" -ForegroundColor Red}
                else {if ($bMoreDetails) {Write-Host "   Accept --> <none>"}}
            }
            
            #-- Process the reject list
            if ($d.RejectMessagesFromSendersOrMembers.Count -gt 0)
            {
                #-- Get more details on the reject list
                $rRejectList = @()
                foreach ($sRejectEntry in $d.RejectMessagesFromSendersOrMembers)
                {
                    $rRejectList += Get-Recipient $sRejectEntry -ea SilentlyContinue
                }
                
                #-- Loop and see if it matches our sender
                $bFound = $false
                $rTmp = @()
                foreach ($oReject in $rRejectList)
                {
                    if ($oReject.Guid -eq $global:sSenderGuid) {$bFound = $true}
                    $rTmp += $oReject.PrimarySmtpAddress.ToString()
                }
                
                #-- Show our results
                if ($bFound)
                {
                    $nCtr = 0
                    foreach ($oReject in $rRejectList)
                    {
                        if ($nCtr -eq 0) {Write-Host "   Reject --> " -NoNewline} else {Write-Host "              " -NoNewline}
                        $nCtr++
                        if ($oReject.Guid -eq $global:sSenderGuid)
                        {
                            Write-Host $oReject.PrimarySmtpAddress.ToString() -ForegroundColor Red
                            $global:rDlsToFixReject += $d
                        }
                        else {Write-Host $oAccept.PrimarySmtpAddress.ToString()}
                    }
                }
                else
                {
                    if ($bMoreDetails) {Write-Host "   Reject -->" ($rTmp -join ',')}
                }
            }
            else {if ($bMoreDetails) {Write-Host "   Reject --> <none>"}}
            #Write-Host ' '
            
            #-- Loop
            foreach($t in $rMembers)
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


#-- Main
$oTmp = $null
$oTmp = Get-Recipient $sSenderEmail -ea SilentlyContinue
if ($oTmp -eq $null) {Write-Host "Sender not found:" $sSenderEmail -ForegroundColor Red}
else
{
    #-- Get the sender GUID
    $global:sSenderGuid = $oTmp.Guid
    
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
}
#-- Show final results
if ($global:rDlsToFixAccept.Count -gt 0)
{
    Write-Host "DLs to add user to accept list:"
    foreach ($oDl in $global:rDlsToFixAccept) {Write-Host "  " $oDl.DisplayName}
}
if ($global:rDlsToFixReject.Count -gt 0)
{
    Write-Host "DLs to add user to reject list:"
    foreach ($oDl in $global:rDlsToFixReject) {Write-Host "  " $oDl.DisplayName}
}
